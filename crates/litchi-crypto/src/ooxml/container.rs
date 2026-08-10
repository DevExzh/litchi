//! `StrongEncryptionDataSpace` compound-container adapter.

use std::io::{Cursor, Seek, SeekFrom, Write};

use litchi_cfb::writer::OleWriter;
use litchi_cfb::{OleError, OleFile};

use crate::spaces::{
    self, Definition, ENCRYPTION_ID, ENCRYPTION_NAME, EncryptionTransform, Header, Map, MapEntry,
    PRIMARY, Reference, ReferenceKind, STORAGE, Version, VersionInfo,
};

use super::{Error, Limits, Mode, Result, malformed, mode};

const DATA_SPACE: &str = "StrongEncryptionDataSpace";
const TRANSFORM: &str = "StrongEncryptionTransform";

/// A `Write + Seek` sink that refuses growth before allocation.
struct Bounded {
    inner: Cursor<Vec<u8>>,
    maximum: u64,
    failure: Option<SinkFailure>,
}

#[derive(Clone, Copy)]
enum SinkFailure {
    Limit { actual: u64, maximum: u64 },
    Allocation,
}

impl Bounded {
    fn new(maximum: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            failure: None,
        }
    }

    fn take_failure(&mut self) -> Option<SinkFailure> {
        self.failure.take()
    }

    fn finish(self) -> Result<Vec<u8>> {
        let bytes = self.inner.into_inner();
        if u64::try_from(bytes.len()).unwrap_or(u64::MAX) > self.maximum {
            return Err(Error::Limit {
                resource: "compound output",
                actual: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
                maximum: self.maximum,
            });
        }
        Ok(bytes)
    }

    fn checked_position(&mut self, position: i128) -> std::io::Result<u64> {
        if position < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "OOXML encrypted output seeked before byte zero",
            ));
        }
        let Ok(absolute) = u64::try_from(position) else {
            return Err(self.limit(u64::MAX));
        };
        if absolute > self.maximum {
            return Err(self.limit(absolute));
        }
        Ok(absolute)
    }

    fn limit(&mut self, actual: u64) -> std::io::Error {
        self.failure.get_or_insert(SinkFailure::Limit {
            actual,
            maximum: self.maximum,
        });
        limit_io()
    }

    fn allocation(&mut self) -> std::io::Error {
        self.failure.get_or_insert(SinkFailure::Allocation);
        allocation_io()
    }
}

impl Write for Bounded {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let Ok(length) = u64::try_from(bytes.len()) else {
            return Err(self.limit(u64::MAX));
        };
        let Some(end) = self.inner.position().checked_add(length) else {
            return Err(self.limit(u64::MAX));
        };
        if end > self.maximum {
            return Err(self.limit(end));
        }
        let Ok(end_index) = usize::try_from(end) else {
            return Err(self.limit(end));
        };
        let capacity = self.inner.get_ref().capacity();
        let current = self.inner.get_ref().len();
        let maximum = usize::try_from(self.maximum).unwrap_or(usize::MAX);
        let target = end_index.max(capacity.saturating_mul(2)).min(maximum);
        if end_index > capacity
            && self
                .inner
                .get_mut()
                .try_reserve_exact(target.saturating_sub(current))
                .is_err()
        {
            return Err(self.allocation());
        }
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for Bounded {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let next = match position {
            SeekFrom::Start(value) => i128::from(value),
            SeekFrom::Current(delta) => i128::from(self.inner.position()) + i128::from(delta),
            SeekFrom::End(delta) => {
                i128::try_from(self.inner.get_ref().len()).map_err(|_err| limit_io())?
                    + i128::from(delta)
            },
        };
        let checked = self.checked_position(next)?;
        self.inner.seek(SeekFrom::Start(checked))
    }
}

/// Read only the bounded `EncryptionInfo` stream for password-free classification.
pub(super) fn read_info(bytes: &[u8], limits: &Limits) -> Result<Vec<u8>> {
    Limits::bytes("compound input", bytes.len(), limits.max_input_bytes)?;
    let mut ole = OleFile::open(Cursor::new(bytes)).map_err(map_reader_error)?;
    let info_len = ole
        .stream_len(&["EncryptionInfo"])
        .map_err(map_reader_error)?;
    declared_bytes("EncryptionInfo", info_len, limits.max_info_bytes)?;
    let info = ole
        .open_stream(&["EncryptionInfo"])
        .map_err(map_reader_error)?;
    Ok(info)
}

/// Wrap one `EncryptionInfo` and `EncryptedPackage` pair in the normative CFB graph.
pub(super) fn write(info: &[u8], encrypted: Vec<u8>, limits: &Limits) -> Result<Vec<u8>> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    Limits::bytes(
        "EncryptedPackage",
        encrypted.len(),
        limits.max_encrypted_bytes,
    )?;

    let mode = mode(info)?;
    let map = spaces::write_map(&expected_map()).map_err(|error| map_spaces_error(&error))?;
    let definition = spaces::write_definition(&Definition {
        transforms: vec![TRANSFORM.to_string()],
    })
    .map_err(|error| map_spaces_error(&error))?;
    let primary = spaces::write_encryption_transform(&EncryptionTransform {
        header: Header {
            transform_id: ENCRYPTION_ID.to_string(),
            transform_name: ENCRYPTION_NAME.to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        },
        encryption_name: match mode {
            Mode::Standard => Some("AES 128".to_string()),
            Mode::Agile => None,
        },
        encryption_block_size: 16,
        cipher_mode: 0,
    })
    .map_err(|error| map_spaces_error(&error))?;
    let version = spaces::write_version_info(&VersionInfo::default())
        .map_err(|error| map_spaces_error(&error))?;

    let mut writer = OleWriter::new();
    stream(&mut writer, &["EncryptionInfo"], info)?;
    writer
        .create_stream_owned(&["EncryptedPackage"], encrypted)
        .map_err(map_writer_error)?;
    storage(&mut writer, &[STORAGE])?;
    storage(&mut writer, &[STORAGE, "DataSpaceInfo"])?;
    storage(&mut writer, &[STORAGE, "TransformInfo"])?;
    storage(&mut writer, &[STORAGE, "TransformInfo", TRANSFORM])?;
    stream(&mut writer, &[STORAGE, "DataSpaceMap"], &map)?;
    stream(
        &mut writer,
        &[STORAGE, "DataSpaceInfo", DATA_SPACE],
        &definition,
    )?;
    stream(
        &mut writer,
        &[STORAGE, "TransformInfo", TRANSFORM, PRIMARY],
        &primary,
    )?;
    stream(&mut writer, &[STORAGE, "Version"], &version)?;

    let mut output = Bounded::new(limits.max_output_bytes);
    if let Err(error) = writer.write_to(&mut output) {
        if let Some(failure) = output.take_failure() {
            return Err(match failure {
                SinkFailure::Limit { actual, maximum } => Error::Limit {
                    resource: "compound output",
                    actual,
                    maximum,
                },
                SinkFailure::Allocation => Error::Allocation("compound output"),
            });
        }
        return Err(map_writer_error(error));
    }
    output.finish()
}

/// Read and validate the two encryption streams from a bounded compound file.
pub(super) fn read(bytes: Vec<u8>, limits: &Limits) -> Result<(Vec<u8>, Vec<u8>)> {
    Limits::bytes("compound input", bytes.len(), limits.max_input_bytes)?;
    let mut ole = OleFile::open(Cursor::new(bytes)).map_err(map_reader_error)?;

    let info_len = ole
        .stream_len(&["EncryptionInfo"])
        .map_err(map_reader_error)?;
    declared_bytes("EncryptionInfo", info_len, limits.max_info_bytes)?;
    let info = ole
        .open_stream(&["EncryptionInfo"])
        .map_err(map_reader_error)?;
    let mode = mode(&info)?;

    // LibreOffice has emitted otherwise valid encrypted packages without the
    // DataSpaces storage. Retain that narrow read compatibility, but when the
    // graph exists require the complete StrongEncryption profile.
    match spaces::inspect(&mut ole).map_err(|error| map_spaces_error(&error))? {
        Some(graph) => validate_graph(&graph, mode)?,
        None if limits.allow_missing_data_spaces => {},
        None => {
            return Err(malformed(
                "encrypted OOXML container is missing the required DataSpaces graph",
            ));
        },
    }

    let encrypted_len = ole
        .stream_len(&["EncryptedPackage"])
        .map_err(map_reader_error)?;
    declared_bytes(
        "EncryptedPackage",
        encrypted_len,
        limits.max_encrypted_bytes,
    )?;
    let encrypted = ole
        .open_stream(&["EncryptedPackage"])
        .map_err(map_reader_error)?;
    drop(ole);
    Ok((info, encrypted))
}

fn declared_bytes(resource: &'static str, actual: u64, maximum: usize) -> Result<()> {
    let Ok(maximum_u64) = u64::try_from(maximum) else {
        // On a target with a wider usize, every possible declared u64 length
        // is necessarily within this limit.
        return Ok(());
    };
    if actual > maximum_u64 {
        return Err(Error::Limit {
            resource,
            actual,
            maximum: maximum_u64,
        });
    }
    Ok(())
}

fn validate_graph(graph: &spaces::Graph, mode: Mode) -> Result<()> {
    if graph.map != expected_map() {
        return Err(malformed(
            "StrongEncryptionDataSpace map is not the required single stream reference",
        ));
    }
    if graph.definitions.len() != 1
        || graph.definitions[0].name != DATA_SPACE
        || graph.definitions[0].definition.transforms.as_slice() != [TRANSFORM]
    {
        return Err(malformed(
            "StrongEncryptionDataSpace definition is not the required single transform",
        ));
    }
    if graph.transforms.len() != 1 || graph.transforms[0].name != TRANSFORM {
        return Err(malformed(
            "StrongEncryptionDataSpace does not contain exactly one encryption transform",
        ));
    }
    let transform = &graph.transforms[0];
    if transform.header.transform_id != ENCRYPTION_ID
        || transform.header.transform_name != ENCRYPTION_NAME
        || transform.header.reader != Version::V1_0
        || transform.header.updater != Version::V1_0
        || transform.header.writer != Version::V1_0
    {
        return Err(malformed("StrongEncryptionTransform header is invalid"));
    }
    let Some(encryption) = &transform.encryption else {
        return Err(malformed(
            "StrongEncryptionTransform is missing its encryption parameters",
        ));
    };
    // EncryptionInfo is authoritative when these advisory parameters disagree.
    // Agile is the one exception: its EncryptionName MUST be a null string.
    if mode == Mode::Agile && encryption.encryption_name.is_some() {
        return Err(malformed(
            "Agile StrongEncryptionTransform EncryptionName is not null",
        ));
    }
    if graph.irm.is_some()
        || graph.label_info.is_some()
        || graph.summary_information_integrity.is_some()
        || graph.document_summary_information_integrity.is_some()
        || graph.custom_xml_data_store.is_some()
    {
        return Err(malformed(
            "encrypted OOXML DataSpaces graph contains an unrelated protected-content profile",
        ));
    }
    Ok(())
}

fn expected_map() -> Map {
    Map {
        entries: vec![MapEntry {
            references: vec![Reference {
                kind: ReferenceKind::Stream,
                component: "EncryptedPackage".to_string(),
            }],
            data_space_name: DATA_SPACE.to_string(),
        }],
    }
}

fn storage(writer: &mut OleWriter, path: &[&str]) -> Result<()> {
    writer.create_storage(path).map_err(map_writer_error)
}

fn stream(writer: &mut OleWriter, path: &[&str], bytes: &[u8]) -> Result<()> {
    writer.create_stream(path, bytes).map_err(map_writer_error)
}

fn map_writer_error(error: OleError) -> Error {
    match error {
        OleError::Io(io_error) => Error::Io(io_error),
        OleError::Allocation { resource, .. } => Error::Allocation(resource),
        OleError::Committed { source } => Error::Io(source),
        OleError::InvalidFormat(message) => Error::Container(format!("Invalid format: {message}")),
        OleError::InvalidData(message) => Error::Container(format!("Invalid data: {message}")),
        OleError::NotOleFile => Error::Container("Not an OLE file".to_string()),
        OleError::CorruptedFile(message) => Error::Container(format!("Corrupted file: {message}")),
        OleError::StreamNotFound => Error::Container("Stream not found".to_string()),
        OleError::SourceChanged { expected, observed } => Error::Container(format!(
            "CFB positional source changed (expected {expected:?}, observed {observed:?})"
        )),
    }
}

fn map_reader_error(error: OleError) -> Error {
    // All readers in this adapter are in-memory cursors. An I/O-looking error
    // therefore denotes truncated container bytes, not an external I/O fault.
    match error {
        OleError::Io(io_error) => Error::Container(format!("IO error: {io_error}")),
        OleError::Allocation { resource, .. } => Error::Allocation(resource),
        OleError::Committed { source } => Error::Io(source),
        OleError::InvalidFormat(message) => Error::Container(format!("Invalid format: {message}")),
        OleError::InvalidData(message) => Error::Container(format!("Invalid data: {message}")),
        OleError::NotOleFile => Error::Container("Not an OLE file".to_string()),
        OleError::CorruptedFile(message) => Error::Container(format!("Corrupted file: {message}")),
        OleError::StreamNotFound => Error::Container("Stream not found".to_string()),
        OleError::SourceChanged { expected, observed } => Error::Container(format!(
            "CFB positional source changed (expected {expected:?}, observed {observed:?})"
        )),
    }
}

fn map_spaces_error(error: &spaces::Error) -> Error {
    Error::Container(error.to_string())
}

fn limit_io() -> std::io::Error {
    std::io::Error::other("OOXML encrypted output exceeds configured maximum")
}

fn allocation_io() -> std::io::Error {
    std::io::Error::other("OOXML encrypted output allocation failed")
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        reason = "test code panics on failure; expect keeps assertions concise"
    )]
    use super::*;
    use crate::spaces::{ENCRYPTION_ID, inspect_bytes};

    #[test]
    fn wrapper_has_the_normative_strong_encryption_graph() {
        let info = [3, 0, 2, 0, 0x24, 0, 0, 0];
        let encrypted = [4, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 4];
        let limits = Limits::default();
        let bytes = write(&info, encrypted.to_vec(), &limits).expect("valid wrapper");

        let graph = inspect_bytes(&bytes)
            .expect("valid DataSpaces")
            .expect("present DataSpaces");
        assert_eq!(graph.map, expected_map());
        assert_eq!(graph.transforms[0].header.transform_id, ENCRYPTION_ID);
        assert_eq!(
            graph.transforms[0]
                .encryption
                .as_ref()
                .and_then(|encryption| encryption.encryption_name.as_deref()),
            Some("AES 128")
        );
        let (parsed_info, parsed_package) = read(bytes, &limits).expect("read wrapper");
        assert_eq!(parsed_info, info);
        assert_eq!(parsed_package, encrypted);
    }

    #[test]
    fn standard_encryption_info_overrides_mismatched_transform_parameters() {
        let info = [3, 0, 2, 0, 0x24, 0, 0, 0];
        let bytes = write(&info, vec![0; 8], &Limits::default()).expect("valid wrapper");
        let mut graph = inspect_bytes(&bytes)
            .expect("valid DataSpaces")
            .expect("present DataSpaces");
        let encryption = graph.transforms[0]
            .encryption
            .as_mut()
            .expect("encryption parameters");
        encryption.encryption_name = Some("AES 256".to_string());
        encryption.encryption_block_size = 32;
        encryption.cipher_mode = 1;

        validate_graph(&graph, Mode::Standard)
            .expect("EncryptionInfo is authoritative for Standard encryption");
    }

    #[test]
    fn agile_encryption_name_must_remain_null() {
        let info = [4, 0, 4, 0, 0x40, 0, 0, 0];
        let bytes = write(&info, vec![0; 8], &Limits::default()).expect("valid wrapper");
        let mut graph = inspect_bytes(&bytes)
            .expect("valid DataSpaces")
            .expect("present DataSpaces");
        {
            let encryption = graph.transforms[0]
                .encryption
                .as_mut()
                .expect("encryption parameters");
            encryption.encryption_block_size = 32;
            encryption.cipher_mode = 1;
        }
        validate_graph(&graph, Mode::Agile)
            .expect("EncryptionInfo is authoritative for Agile parameters");

        graph.transforms[0]
            .encryption
            .as_mut()
            .expect("encryption parameters")
            .encryption_name = Some("AES 128".to_string());
        assert!(matches!(
            validate_graph(&graph, Mode::Agile),
            Err(Error::Malformed(_))
        ));
    }

    #[test]
    fn accepts_the_narrow_libreoffice_no_dataspaces_profile() {
        let info = [3, 0, 2, 0, 0x24, 0, 0, 0];
        let encrypted = [0; 8];
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptionInfo"], &info)
            .expect("EncryptionInfo");
        writer
            .create_stream(&["EncryptedPackage"], &encrypted)
            .expect("EncryptedPackage");
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).expect("write CFB");

        assert!(matches!(
            read(bytes.get_ref().clone(), &Limits::default()),
            Err(Error::Malformed(_))
        ));

        let limits = Limits {
            allow_missing_data_spaces: true,
            ..Limits::default()
        };
        let (parsed_info, parsed_package) =
            read(bytes.into_inner(), &limits).expect("read compatibility package");
        assert_eq!(parsed_info, info);
        assert_eq!(parsed_package, encrypted);
    }

    #[test]
    fn encryption_info_declared_size_is_limited_before_materialization() {
        let info = [3, 0, 2, 0, 0x24, 0, 0, 0, 1, 2, 3, 4, 5];
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptionInfo"], &info)
            .expect("EncryptionInfo");
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).expect("write CFB");

        let at_boundary = Limits {
            max_info_bytes: info.len(),
            ..Limits::default()
        };
        assert_eq!(
            read_info(bytes.get_ref(), &at_boundary).expect("declared size at limit"),
            info
        );

        let over_boundary = Limits {
            max_info_bytes: info.len() - 1,
            ..Limits::default()
        };
        assert!(matches!(
            read_info(bytes.get_ref(), &over_boundary),
            Err(Error::Limit {
                resource: "EncryptionInfo",
                actual: 13,
                maximum: 12,
            })
        ));
    }

    #[test]
    fn encrypted_package_declared_size_is_limited_before_materialization() {
        let info = [3, 0, 2, 0, 0x24, 0, 0, 0];
        let encrypted = [0; 25];
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptionInfo"], &info)
            .expect("EncryptionInfo");
        writer
            .create_stream(&["EncryptedPackage"], &encrypted)
            .expect("EncryptedPackage");
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).expect("write CFB");

        let at_boundary = Limits {
            max_encrypted_bytes: encrypted.len(),
            allow_missing_data_spaces: true,
            ..Limits::default()
        };
        let (_, parsed_package) =
            read(bytes.get_ref().clone(), &at_boundary).expect("declared size at limit");
        assert_eq!(parsed_package, encrypted);

        let over_boundary = Limits {
            max_encrypted_bytes: encrypted.len() - 1,
            allow_missing_data_spaces: true,
            ..Limits::default()
        };
        assert!(matches!(
            read(bytes.into_inner(), &over_boundary),
            Err(Error::Limit {
                resource: "EncryptedPackage",
                actual: 25,
                maximum: 24,
            })
        ));
    }

    #[test]
    fn output_limit_stops_before_growth() {
        let mut sink = Bounded::new(4);
        assert_eq!(sink.write(&[1, 2, 3, 4]).expect("within limit"), 4);
        assert!(sink.write(&[5]).is_err());
        assert_eq!(sink.inner.get_ref(), &[1, 2, 3, 4]);
        assert!(matches!(
            sink.take_failure(),
            Some(SinkFailure::Limit {
                actual: 5,
                maximum: 4,
            })
        ));
    }

    #[test]
    fn wrapper_reports_output_ceiling_as_typed_limit() {
        let limits = Limits {
            max_output_bytes: 512,
            ..Limits::default()
        };
        let error = write(&[3, 0, 2, 0, 0x24, 0, 0, 0], vec![0; 16], &limits)
            .expect_err("compound file exceeds one sector");
        assert!(matches!(
            error,
            Error::Limit {
                resource: "compound output",
                actual,
                maximum: 512,
            } if actual > 512
        ));
    }

    #[test]
    fn cfb_allocation_failure_keeps_its_typed_category() {
        let mut impossible = Vec::<u8>::new();
        let source = impossible
            .try_reserve(usize::MAX)
            .expect_err("capacity overflow must fail");
        let error = map_writer_error(OleError::Allocation {
            resource: "FAT plan",
            source,
        });

        assert!(matches!(error, Error::Allocation("FAT plan")));
    }
}

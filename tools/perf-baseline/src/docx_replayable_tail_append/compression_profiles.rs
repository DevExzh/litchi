//! Opt-in compression profiles for the replayable DOCX benchmark fixture.
//!
//! `Current` is the package writer's normal output and is deliberately an
//! ownership-preserving identity operation.  The two explicit profiles replace
//! only `word/document.xml` through the ZIP preservation API.  Every other
//! source member is copied from its validated raw local span and central
//! record.  This keeps the Store/Deflate pair useful for compression work
//! without silently turning a reassembled archive into a comparison with the
//! default package-writer lineage.

use std::{collections::TryReserveError, error::Error, fmt, io};

use serde::Serialize;
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{
    CompressionMethod, PreservationAction, PreservationIndex, PreservationPlan,
    RECOMMENDED_BUFFER_SIZE, RegeneratedEntry, ZipArchive,
};

/// The benchmark's source-fixture compression lineage.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum CompressionProfile {
    /// Preserve the exact bytes returned by the ordinary package writer.
    Current,
    /// Store the selected main-document member without compression.
    Store,
    /// Deflate the selected main-document member.
    Deflate,
}

impl CompressionProfile {
    /// Parse the stable command-line/report spelling.
    pub fn parse(value: &str) -> Result<Self, CompressionProfileError> {
        match value {
            "current" | "default" => Ok(Self::Current),
            "store" => Ok(Self::Store),
            "deflate" => Ok(Self::Deflate),
            _ => Err(CompressionProfileError::InvalidProfile(value.to_owned())),
        }
    }

    /// Return the stable command-line/report spelling.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Current => "current",
            Self::Store => "store",
            Self::Deflate => "deflate",
        }
    }

    const fn method(self) -> Option<CompressionMethod> {
        match self {
            Self::Current => None,
            Self::Store => Some(CompressionMethod::Store),
            Self::Deflate => Some(CompressionMethod::Deflate),
        }
    }
}

/// The member selected for the explicit compression profiles.
pub const MAIN_DOCUMENT_PATH: &str = "word/document.xml";

/// Errors returned while validating or applying an explicit fixture profile.
#[derive(Debug)]
pub enum CompressionProfileError {
    /// The profile spelling was not recognized.
    InvalidProfile(String),
    /// The base archive has no exact main-document member.
    MissingMainDocument,
    /// The base archive has more than one exact main-document member.
    DuplicateMainDocument,
    /// The ZIP archive or its preservation layout was invalid.
    Zip(soapberry_zip::Error),
    /// The preservation plan could not reserve its bounded action list.
    PlanCapacity(TryReserveError),
    /// The generated output could not be written to its in-memory sink.
    Output(io::Error),
}

impl fmt::Display for CompressionProfileError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidProfile(value) => {
                write!(formatter, "invalid compression profile {value:?}")
            },
            Self::MissingMainDocument => {
                formatter.write_str("fixture archive is missing word/document.xml")
            },
            Self::DuplicateMainDocument => {
                formatter.write_str("fixture archive contains duplicate word/document.xml entries")
            },
            Self::Zip(error) => write!(formatter, "ZIP profile rewrite failed: {error}"),
            Self::PlanCapacity(error) => {
                write!(
                    formatter,
                    "compression profile plan allocation failed: {error}"
                )
            },
            Self::Output(error) => write!(formatter, "compression profile output failed: {error}"),
        }
    }
}

impl Error for CompressionProfileError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Zip(error) => Some(error),
            Self::PlanCapacity(error) => Some(error),
            Self::Output(error) => Some(error),
            Self::InvalidProfile(_) | Self::MissingMainDocument | Self::DuplicateMainDocument => {
                None
            },
        }
    }
}

impl From<soapberry_zip::Error> for CompressionProfileError {
    fn from(error: soapberry_zip::Error) -> Self {
        Self::Zip(error)
    }
}

impl From<TryReserveError> for CompressionProfileError {
    fn from(error: TryReserveError) -> Self {
        Self::PlanCapacity(error)
    }
}

impl From<io::Error> for CompressionProfileError {
    fn from(error: io::Error) -> Self {
        Self::Output(error)
    }
}

/// Apply one fixture compression profile to an owned base archive.
///
/// `Current` returns `base` directly, preserving the package writer's exact
/// bytes and allocation ownership.  `Store` and `Deflate` require one exact
/// `word/document.xml` member and regenerate only that member.  The source
/// archive is indexed before output is created, and every other member is
/// copied through [`PreservationIndex`], including its raw local framing and
/// central metadata.  A central local-header offset may be patched by the
/// preservation writer when the regenerated target changes the preceding
/// member span; that offset is layout, while all remaining untouched bytes and
/// fields stay source-owned.
pub fn apply(
    base: Vec<u8>,
    profile: CompressionProfile,
) -> Result<Vec<u8>, CompressionProfileError> {
    let Some(method) = profile.method() else {
        return Ok(base);
    };

    let archive = ZipArchive::from_slice(base.as_slice())?.into_zip_archive();
    let mut index_buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut index_buffer)?;
    let main_ids = index
        .entries()
        .iter()
        .filter(|entry| entry.raw_name_bytes() == MAIN_DOCUMENT_PATH.as_bytes())
        .map(|entry| entry.id())
        .collect::<Vec<_>>();
    let main_id = match main_ids.as_slice() {
        [] => return Err(CompressionProfileError::MissingMainDocument),
        [id] => *id,
        _ => return Err(CompressionProfileError::DuplicateMainDocument),
    };

    let reader = ArchiveReader::new(base.as_slice())?;
    let main_xml = reader.read(MAIN_DOCUMENT_PATH)?;
    let replacement =
        RegeneratedEntry::new(MAIN_DOCUMENT_PATH, main_xml).compression_method(method);
    let mut plan = PreservationPlan::new();
    plan.try_reserve_exact(index.entries().len())?;
    for entry in index.entries() {
        if entry.id() == main_id {
            plan.push(PreservationAction::Regenerate {
                id: entry.id(),
                entry: replacement.clone(),
            });
        } else {
            plan.push(PreservationAction::Copy(entry.id()));
        }
    }

    index
        .write_to(&plan, Vec::new())
        .map_err(CompressionProfileError::from)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;
    use std::io::Write;

    use soapberry_zip::office::ArchiveReader;
    use soapberry_zip::{ZipArchive, ZipArchiveWriter};

    const CONTENT_TYPES: &str = "[Content_Types].xml";
    const RELATIONSHIPS: &str = "_rels/.rels";
    const OPAQUE: &str = "word/perf-opaque.bin";

    #[derive(Clone, Copy, Debug)]
    enum MainPosition {
        First,
        Middle,
    }

    fn fixture(position: MainPosition) -> Vec<u8> {
        // Put the selected member before at least one untouched member.  The
        // Deflate profile must then move that later member's local offset;
        // source-v-output checks below normalize only that central offset.
        let mut archive = ZipArchiveWriter::builder()
            .with_capacity(4)
            .build(Vec::new());
        match position {
            MainPosition::First => {
                archive = write_main(archive);
                archive = write_content_types(archive);
                archive = write_relationships(archive);
                archive = write_opaque(archive);
            },
            MainPosition::Middle => {
                archive = write_content_types(archive);
                archive = write_main(archive);
                archive = write_relationships(archive);
                archive = write_opaque(archive);
            },
        }
        archive.finish().expect("fixture archive must finish")
    }

    fn write_main(archive: ZipArchiveWriter<Vec<u8>>) -> ZipArchiveWriter<Vec<u8>> {
        write_entry(
            archive,
            MAIN_DOCUMENT_PATH,
            CompressionMethod::Store,
            br#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#,
        )
    }

    fn write_content_types(archive: ZipArchiveWriter<Vec<u8>>) -> ZipArchiveWriter<Vec<u8>> {
        write_entry(
            archive,
            CONTENT_TYPES,
            CompressionMethod::Store,
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml"/></Types>"#,
        )
    }

    fn write_relationships(archive: ZipArchiveWriter<Vec<u8>>) -> ZipArchiveWriter<Vec<u8>> {
        write_entry(
            archive,
            RELATIONSHIPS,
            CompressionMethod::Store,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Target="word/document.xml"/></Relationships>"#,
        )
    }

    fn write_opaque(archive: ZipArchiveWriter<Vec<u8>>) -> ZipArchiveWriter<Vec<u8>> {
        write_entry(
            archive,
            OPAQUE,
            CompressionMethod::Store,
            b"opaque payload\0\xff\x01",
        )
    }

    fn write_entry(
        archive: ZipArchiveWriter<Vec<u8>>,
        name: &str,
        method: CompressionMethod,
        payload: &[u8],
    ) -> ZipArchiveWriter<Vec<u8>> {
        let mut entry = archive
            .start_file_owned(name, method)
            .expect("fixture entry must start");
        entry.write_all(payload).expect("fixture entry must write");
        entry.finish().expect("fixture entry must finish")
    }

    fn entry_method(bytes: &[u8], name: &str) -> CompressionMethod {
        let archive = ZipArchive::from_slice(bytes).expect("archive must parse");
        archive
            .entries()
            .find_map(|entry| {
                let entry = entry.expect("entry must parse");
                (entry
                    .file_path()
                    .try_normalize()
                    .expect("path must normalize")
                    .as_ref()
                    == name)
                    .then_some(entry.compression_method())
            })
            .expect("entry must exist")
    }

    fn decoded_member(bytes: &[u8], name: &str) -> Vec<u8> {
        ArchiveReader::new(bytes)
            .expect("archive reader must parse")
            .read(name)
            .expect("member must decode")
    }

    #[derive(Debug, PartialEq, Eq)]
    struct RawMember {
        local_offset: usize,
        local: Vec<u8>,
        central: Vec<u8>,
    }

    fn le_u16(bytes: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(
            bytes
                .get(offset..offset + 2)
                .expect("ZIP u16 field must fit")
                .try_into()
                .expect("ZIP u16 field must have width two"),
        )
    }

    fn le_u32(bytes: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(
            bytes
                .get(offset..offset + 4)
                .expect("ZIP u32 field must fit")
                .try_into()
                .expect("ZIP u32 field must have width four"),
        )
    }

    /// Read the tiny ZIP32 fixture framing independently of the production
    /// preservation index.  This intentionally checks raw bytes instead of
    /// trusting the same index implementation used by `apply`.
    fn raw_members(bytes: &[u8]) -> BTreeMap<String, RawMember> {
        const EOCD_SIZE: usize = 22;
        const CENTRAL_SIZE: usize = 46;
        let eocd = bytes
            .windows(4)
            .rposition(|signature| signature == b"PK\x05\x06")
            .expect("fixture must contain an EOCD");
        assert!(eocd + EOCD_SIZE <= bytes.len());
        let comment_len = usize::from(le_u16(bytes, eocd + 20));
        assert_eq!(eocd + EOCD_SIZE + comment_len, bytes.len());
        let count = usize::from(le_u16(bytes, eocd + 10));
        let central_start =
            usize::try_from(le_u32(bytes, eocd + 16)).expect("central offset must fit usize");
        let central_size =
            usize::try_from(le_u32(bytes, eocd + 12)).expect("central size must fit usize");
        let central_end = central_start
            .checked_add(central_size)
            .expect("central range must not overflow");
        assert_eq!(central_end, eocd);

        let mut central_records = Vec::with_capacity(count);
        let mut cursor = central_start;
        for _ in 0..count {
            let fixed = bytes
                .get(cursor..cursor + CENTRAL_SIZE)
                .expect("central fixed record must fit");
            assert_eq!(&fixed[..4], b"PK\x01\x02");
            let name_len = usize::from(le_u16(fixed, 28));
            let extra_len = usize::from(le_u16(fixed, 30));
            let comment_len = usize::from(le_u16(fixed, 32));
            let record_len = CENTRAL_SIZE
                .checked_add(name_len)
                .and_then(|length| length.checked_add(extra_len))
                .and_then(|length| length.checked_add(comment_len))
                .expect("central record length must not overflow");
            let end = cursor
                .checked_add(record_len)
                .expect("central record range must not overflow");
            let record = bytes
                .get(cursor..end)
                .expect("complete central record must fit")
                .to_vec();
            let name = String::from_utf8(record[CENTRAL_SIZE..CENTRAL_SIZE + name_len].to_vec())
                .expect("fixture member names must be UTF-8");
            let local_offset =
                usize::try_from(le_u32(&record, 42)).expect("fixture local offset must fit usize");
            assert_ne!(
                local_offset,
                usize::try_from(u32::MAX).expect("u32 fits usize")
            );
            central_records.push((name, local_offset, record));
            cursor = end;
        }
        assert_eq!(cursor, central_end);

        let mut physical_order = (0..central_records.len()).collect::<Vec<_>>();
        physical_order.sort_unstable_by_key(|&index| central_records[index].1);
        let mut members = BTreeMap::new();
        for (position, &index) in physical_order.iter().enumerate() {
            let (name, local_offset, central) = &central_records[index];
            let local_end = physical_order
                .get(position + 1)
                .map(|&next| central_records[next].1)
                .unwrap_or(central_start);
            assert!(*local_offset <= local_end);
            let local = bytes
                .get(*local_offset..local_end)
                .expect("complete local member must fit")
                .to_vec();
            assert!(
                members
                    .insert(
                        name.clone(),
                        RawMember {
                            local_offset: *local_offset,
                            local,
                            central: central.clone(),
                        },
                    )
                    .is_none()
            );
        }
        members
    }

    fn normalize_central_local_offset(record: &[u8]) -> Vec<u8> {
        assert!(record.len() >= 46);
        assert_eq!(&record[..4], b"PK\x01\x02");
        let mut normalized = record.to_vec();
        normalized[42..46].fill(0);
        normalized
    }

    fn assert_untouched_source_records(source: &[u8], output: &[u8]) {
        let source_members = raw_members(source);
        let output_members = raw_members(output);
        assert_eq!(source_members.len(), output_members.len());
        assert_eq!(
            source_members.keys().collect::<Vec<_>>(),
            output_members.keys().collect::<Vec<_>>()
        );
        for (name, source_member) in source_members {
            if name == MAIN_DOCUMENT_PATH {
                continue;
            }
            let output_member = output_members
                .get(&name)
                .expect("output must retain every untouched member");
            assert_eq!(
                source_member.local, output_member.local,
                "local record {name}"
            );
            assert_eq!(
                normalize_central_local_offset(&source_member.central),
                normalize_central_local_offset(&output_member.central),
                "central metadata {name}"
            );
        }
    }

    #[test]
    fn current_is_an_owned_identity_operation() {
        let base = fixture(MainPosition::Middle);
        let pointer = base.as_ptr();
        let capacity = base.capacity();
        let output = apply(base, CompressionProfile::Current).expect("current profile");
        assert_eq!(output.as_ptr(), pointer);
        assert_eq!(output.capacity(), capacity);
    }

    #[test]
    fn explicit_profiles_select_only_main_compression_and_keep_source_xml() {
        for position in [MainPosition::First, MainPosition::Middle] {
            let base = fixture(position);
            let source_xml = decoded_member(&base, MAIN_DOCUMENT_PATH);
            for profile in [
                CompressionProfile::Current,
                CompressionProfile::Store,
                CompressionProfile::Deflate,
            ] {
                let output = apply(base.clone(), profile).expect("compression profile");
                assert_eq!(
                    decoded_member(&output, MAIN_DOCUMENT_PATH),
                    source_xml,
                    "profile {} must retain source XML",
                    profile.name()
                );
                let expected_method = match profile {
                    CompressionProfile::Current | CompressionProfile::Store => {
                        CompressionMethod::Store
                    },
                    CompressionProfile::Deflate => CompressionMethod::Deflate,
                };
                assert_eq!(
                    entry_method(&output, MAIN_DOCUMENT_PATH),
                    expected_method,
                    "profile {} must select its main method",
                    profile.name()
                );
            }
        }
    }

    #[test]
    fn explicit_profiles_preserve_untouched_source_records_with_only_offset_relocation() {
        for position in [MainPosition::First, MainPosition::Middle] {
            let base = fixture(position);
            let source_members = raw_members(&base);
            for profile in [CompressionProfile::Store, CompressionProfile::Deflate] {
                let output = apply(base.clone(), profile).expect("compression profile");
                assert_untouched_source_records(&base, &output);

                let source_opaque = source_members
                    .get(OPAQUE)
                    .expect("source opaque member must exist");
                let output_members = raw_members(&output);
                let output_opaque = output_members
                    .get(OPAQUE)
                    .expect("output opaque member must exist");
                if profile == CompressionProfile::Deflate {
                    assert_ne!(
                        source_opaque.local_offset, output_opaque.local_offset,
                        "Deflate must relocate a later untouched member"
                    );
                    assert_ne!(
                        source_opaque.central, output_opaque.central,
                        "relocation must change only the central offset field"
                    );
                    assert_eq!(
                        normalize_central_local_offset(&source_opaque.central),
                        normalize_central_local_offset(&output_opaque.central)
                    );
                }
            }
        }
    }

    #[test]
    fn profile_parse_and_missing_or_duplicate_main_are_typed() {
        assert_eq!(
            CompressionProfile::parse("default").expect("alias"),
            CompressionProfile::Current
        );
        assert_eq!(
            CompressionProfile::parse("store").expect("store"),
            CompressionProfile::Store
        );
        assert!(matches!(
            CompressionProfile::parse("bogus"),
            Err(CompressionProfileError::InvalidProfile(_))
        ));

        let missing = write_entry(
            ZipArchiveWriter::builder()
                .with_capacity(1)
                .build(Vec::new()),
            "word/other.xml",
            CompressionMethod::Store,
            b"payload",
        )
        .finish()
        .expect("archive");
        assert!(matches!(
            apply(missing, CompressionProfile::Store),
            Err(CompressionProfileError::MissingMainDocument)
        ));

        let mut duplicate = ZipArchiveWriter::builder()
            .with_capacity(2)
            .build(Vec::new());
        duplicate = write_entry(
            duplicate,
            MAIN_DOCUMENT_PATH,
            CompressionMethod::Store,
            b"first",
        );
        duplicate = write_entry(
            duplicate,
            MAIN_DOCUMENT_PATH,
            CompressionMethod::Store,
            b"second",
        );
        let duplicate = duplicate.finish().expect("duplicate archive");
        assert!(matches!(
            apply(duplicate, CompressionProfile::Store),
            Err(CompressionProfileError::DuplicateMainDocument)
        ));
    }
}

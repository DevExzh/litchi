//! Typed MS-OFFCRYPTO DataSpaces and IRM metadata.
//!
//! This module validates the structural graph only. XrML licenses and
//! protected content remain inert and are never activated, fetched, or
//! decrypted.
//!
//! ```no_run
//! use litchi_ole::office_crypto::data_spaces::inspect_data_spaces_bytes;
//!
//! let bytes = std::fs::read("protected.docx")?;
//! if let Some(graph) = inspect_data_spaces_bytes(&bytes)? {
//!     println!("IRM profile: {:?}", graph.irm);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::io::{Read, Seek};

use litchi_cfb::OleFile;

use super::property_integrity::{
    DOCUMENT_SUMMARY_INFORMATION_STREAM, ENCRYPTED_DOCUMENT_SUMMARY_INFORMATION_HASH_STREAM,
    ENCRYPTED_SUMMARY_INFORMATION_HASH_STREAM, EncryptedPropertyStreamInfo,
    SUMMARY_INFORMATION_STREAM, checksum_matches, parse_encrypted_property_stream_info,
};
use super::sensitivity_labels::{SensitivityLabelList, parse_label_info};
use litchi_ole_common::custom_xml_data::{
    DataStorePromotion, MsoDataStore, inspect_mso_data_store,
};

const HEADER_LENGTH: u32 = 8;
const TRANSFORM_TYPE: u32 = 1;
const EXTENSIBILITY_HEADER_LENGTH: u32 = 4;
const MAX_STREAM_BYTES: usize = 16 * 1024 * 1024;
const MAX_ENTRIES: usize = 65_536;
const MAX_STRING_BYTES: usize = 1_048_576;
const MAX_XML_DEPTH: usize = 256;

pub const DATA_SPACES_STORAGE: &str = "\u{0006}DataSpaces";
pub const PRIMARY_STREAM: &str = "\u{0006}Primary";
pub const DATA_SPACES_FEATURE: &str = "Microsoft.Container.DataSpaces";
pub const DRM_TRANSFORM_ID: &str = "{C73DFACD-061F-43B0-8B64-0C620D2A8B50}";
pub const DRM_TRANSFORM_NAME: &str = "Microsoft.Metadata.DRMTransform";
pub const LZX_TRANSFORM_ID: &str = "{86DE7F2B-DDCE-486d-B016-405BBE82B8BC}";
pub const LZX_TRANSFORM_NAME: &str = "Microsoft.Metadata.CompressionTransform";
pub const ENCRYPTION_TRANSFORM_ID: &str = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
pub const ENCRYPTION_TRANSFORM_NAME: &str = "Microsoft.Container.EncryptionTransform";

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DataSpaceError {
    Invalid(String),
    Ole(String),
}

impl fmt::Display for DataSpaceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(message) => write!(formatter, "invalid DataSpaces structure: {message}"),
            Self::Ole(message) => write!(formatter, "OLE DataSpaces error: {message}"),
        }
    }
}

impl std::error::Error for DataSpaceError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataSpaceVersion {
    pub major: u16,
    pub minor: u16,
}

impl DataSpaceVersion {
    pub const V1_0: Self = Self { major: 1, minor: 0 };
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceVersionInfo {
    pub feature_identifier: String,
    pub reader: DataSpaceVersion,
    pub updater: DataSpaceVersion,
    pub writer: DataSpaceVersion,
}

impl Default for DataSpaceVersionInfo {
    fn default() -> Self {
        Self {
            feature_identifier: DATA_SPACES_FEATURE.to_string(),
            reader: DataSpaceVersion::V1_0,
            updater: DataSpaceVersion::V1_0,
            writer: DataSpaceVersion::V1_0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataSpaceReferenceKind {
    Stream,
    Storage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceReference {
    pub kind: DataSpaceReferenceKind,
    pub component: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceMapEntry {
    pub references: Vec<DataSpaceReference>,
    pub data_space_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceMap {
    pub entries: Vec<DataSpaceMapEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceDefinition {
    pub transforms: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TransformInfoHeader {
    pub transform_id: String,
    pub transform_name: String,
    pub reader: DataSpaceVersion,
    pub updater: DataSpaceVersion,
    pub writer: DataSpaceVersion,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IrmTransformInfo {
    pub header: TransformInfoHeader,
    /// Signed issuance license XML retained verbatim and never interpreted.
    pub publishing_license: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncryptionTransformInfo {
    pub header: TransformInfoHeader,
    /// Null when EncryptionInfo is authoritative, as with Agile encryption.
    pub encryption_name: Option<String>,
    pub encryption_block_size: u32,
    pub cipher_mode: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IrmEndUserLicense {
    pub stream_name: String,
    /// Base64-encoded Unicode LicenseID retained verbatim.
    pub encoded_license_id: String,
    /// Certificate-chain XML retained verbatim and never interpreted.
    pub certificate_chain: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedDataSpaceDefinition {
    pub name: String,
    pub definition: DataSpaceDefinition,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceTransform {
    pub name: String,
    pub header: TransformInfoHeader,
    pub irm: Option<IrmTransformInfo>,
    pub encryption: Option<EncryptionTransformInfo>,
    pub end_user_licenses: Vec<IrmEndUserLicense>,
    /// Non-IRM bytes following the transform header.
    pub opaque_tail: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IrmDocumentKind {
    Ooxml,
    LegacyBinary,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IrmDataSpace {
    pub document_kind: IrmDocumentKind,
    pub protected_content_stream: String,
    pub viewer_content_stream: Option<String>,
    pub transform_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSpaceGraph {
    pub version: DataSpaceVersionInfo,
    pub map: DataSpaceMap,
    pub definitions: Vec<NamedDataSpaceDefinition>,
    pub transforms: Vec<DataSpaceTransform>,
    pub irm: Option<IrmDataSpace>,
    /// Exact sensitivity-label XML bytes, retained inert when present.
    pub label_info: Option<Vec<u8>>,
    /// Validated typed view of `label_info`.
    pub sensitivity_labels: Option<SensitivityLabelList>,
    /// Integrity metadata for the public SummaryInformation property stream.
    pub summary_information_integrity: Option<PropertyStreamIntegrity>,
    /// Integrity metadata for the public DocumentSummaryInformation property stream.
    pub document_summary_information_integrity: Option<PropertyStreamIntegrity>,
    /// Public legacy Custom XML mirror and its IRM promotion semantics.
    pub custom_xml_data_store: Option<MsoDataStore>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PropertyStreamIntegrity {
    pub info: EncryptedPropertyStreamInfo,
    /// `None` means a future info-stream version that readers must ignore.
    pub checksum_matches: Option<bool>,
}

pub fn parse_version_info(data: &[u8]) -> Result<DataSpaceVersionInfo, DataSpaceError> {
    let mut reader = SliceReader::new(data)?;
    let value = DataSpaceVersionInfo {
        feature_identifier: reader.unicode_lpp4()?,
        reader: reader.version()?,
        updater: reader.version()?,
        writer: reader.version()?,
    };
    reader.finish()?;
    validate_version_info(&value)?;
    Ok(value)
}

pub fn write_version_info(value: &DataSpaceVersionInfo) -> Result<Vec<u8>, DataSpaceError> {
    validate_version_info(value)?;
    let mut output = Vec::new();
    write_unicode_lpp4(&mut output, &value.feature_identifier)?;
    write_version(&mut output, value.reader);
    write_version(&mut output, value.updater);
    write_version(&mut output, value.writer);
    Ok(output)
}

pub fn parse_data_space_map(data: &[u8]) -> Result<DataSpaceMap, DataSpaceError> {
    let mut reader = SliceReader::new(data)?;
    require_u32(reader.u32()?, HEADER_LENGTH, "DataSpaceMap.HeaderLength")?;
    let count = bounded_count(reader.u32()?, "DataSpaceMap.EntryCount")?;
    if count == 0 {
        return Err(invalid("DataSpaceMap requires at least one entry"));
    }
    let mut entries = Vec::with_capacity(count);
    for _ in 0..count {
        let start = reader.position();
        let length = usize::try_from(reader.u32()?)
            .map_err(|_| invalid("DataSpaceMapEntry.Length overflows usize"))?;
        if length < 12 || start.checked_add(length).is_none_or(|end| end > data.len()) {
            return Err(invalid("DataSpaceMapEntry.Length exceeds its stream"));
        }
        let reference_count =
            bounded_count(reader.u32()?, "DataSpaceMapEntry.ReferenceComponentCount")?;
        if reference_count == 0 {
            return Err(invalid("DataSpaceMapEntry requires a reference component"));
        }
        let mut references = Vec::with_capacity(reference_count);
        for _ in 0..reference_count {
            let kind = match reader.u32()? {
                0 => DataSpaceReferenceKind::Stream,
                1 => DataSpaceReferenceKind::Storage,
                value => {
                    return Err(invalid(format!(
                        "unknown DataSpaceReferenceComponent type {value}"
                    )));
                },
            };
            references.push(DataSpaceReference {
                kind,
                component: reader.unicode_lpp4()?,
            });
        }
        let data_space_name = reader.unicode_lpp4()?;
        if reader.position() != start + length {
            return Err(invalid(
                "DataSpaceMapEntry.Length does not match its fields",
            ));
        }
        entries.push(DataSpaceMapEntry {
            references,
            data_space_name,
        });
    }
    reader.finish()?;
    let value = DataSpaceMap { entries };
    validate_map(&value)?;
    Ok(value)
}

pub fn write_data_space_map(value: &DataSpaceMap) -> Result<Vec<u8>, DataSpaceError> {
    validate_map(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&HEADER_LENGTH.to_le_bytes());
    write_count(&mut output, value.entries.len(), "DataSpaceMap.EntryCount")?;
    for entry in &value.entries {
        let start = output.len();
        output.extend_from_slice(&0u32.to_le_bytes());
        write_count(
            &mut output,
            entry.references.len(),
            "DataSpaceMapEntry.ReferenceComponentCount",
        )?;
        for reference in &entry.references {
            let kind = match reference.kind {
                DataSpaceReferenceKind::Stream => 0u32,
                DataSpaceReferenceKind::Storage => 1u32,
            };
            output.extend_from_slice(&kind.to_le_bytes());
            write_unicode_lpp4(&mut output, &reference.component)?;
        }
        write_unicode_lpp4(&mut output, &entry.data_space_name)?;
        let length = u32::try_from(output.len() - start)
            .map_err(|_| invalid("DataSpaceMapEntry.Length exceeds u32"))?;
        output[start..start + 4].copy_from_slice(&length.to_le_bytes());
    }
    Ok(output)
}

pub fn parse_data_space_definition(data: &[u8]) -> Result<DataSpaceDefinition, DataSpaceError> {
    let mut reader = SliceReader::new(data)?;
    require_u32(
        reader.u32()?,
        HEADER_LENGTH,
        "DataSpaceDefinition.HeaderLength",
    )?;
    let count = bounded_count(reader.u32()?, "DataSpaceDefinition.TransformReferenceCount")?;
    if count == 0 {
        return Err(invalid(
            "DataSpaceDefinition requires at least one transform",
        ));
    }
    let mut transforms = Vec::with_capacity(count);
    for _ in 0..count {
        transforms.push(reader.unicode_lpp4()?);
    }
    reader.finish()?;
    let value = DataSpaceDefinition { transforms };
    validate_definition(&value)?;
    Ok(value)
}

pub fn write_data_space_definition(value: &DataSpaceDefinition) -> Result<Vec<u8>, DataSpaceError> {
    validate_definition(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&HEADER_LENGTH.to_le_bytes());
    write_count(
        &mut output,
        value.transforms.len(),
        "DataSpaceDefinition.TransformReferenceCount",
    )?;
    for transform in &value.transforms {
        write_unicode_lpp4(&mut output, transform)?;
    }
    Ok(output)
}

pub fn parse_transform_header(data: &[u8]) -> Result<(TransformInfoHeader, usize), DataSpaceError> {
    let mut reader = SliceReader::new(data)?;
    let transform_length =
        usize::try_from(reader.u32()?).map_err(|_| invalid("TransformLength overflows usize"))?;
    require_u32(reader.u32()?, TRANSFORM_TYPE, "TransformType")?;
    let transform_id = reader.unicode_lpp4()?;
    if reader.position() != transform_length {
        return Err(invalid("TransformLength does not end before TransformName"));
    }
    let value = TransformInfoHeader {
        transform_id,
        transform_name: reader.unicode_lpp4()?,
        reader: reader.version()?,
        updater: reader.version()?,
        writer: reader.version()?,
    };
    validate_transform_header(&value)?;
    Ok((value, reader.position()))
}

pub fn write_transform_header(value: &TransformInfoHeader) -> Result<Vec<u8>, DataSpaceError> {
    validate_transform_header(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&TRANSFORM_TYPE.to_le_bytes());
    write_unicode_lpp4(&mut output, &value.transform_id)?;
    let transform_length =
        u32::try_from(output.len()).map_err(|_| invalid("TransformLength exceeds u32"))?;
    output[..4].copy_from_slice(&transform_length.to_le_bytes());
    write_unicode_lpp4(&mut output, &value.transform_name)?;
    write_version(&mut output, value.reader);
    write_version(&mut output, value.updater);
    write_version(&mut output, value.writer);
    Ok(output)
}

pub fn parse_irm_transform(data: &[u8]) -> Result<IrmTransformInfo, DataSpaceError> {
    let (header, consumed) = parse_transform_header(data)?;
    validate_drm_header(&header)?;
    let mut reader = SliceReader::at(data, consumed)?;
    require_u32(
        reader.u32()?,
        EXTENSIBILITY_HEADER_LENGTH,
        "ExtensibilityHeader.Length",
    )?;
    let publishing_license = reader.utf8_lpp4()?;
    if publishing_license.as_deref().is_none_or(str::is_empty) {
        return Err(invalid("IRM publishing license cannot be null or empty"));
    }
    validate_inert_xml(
        publishing_license.as_deref().expect("presence checked"),
        "IRM publishing license",
    )?;
    reader.finish()?;
    Ok(IrmTransformInfo {
        header,
        publishing_license,
    })
}

pub fn write_irm_transform(value: &IrmTransformInfo) -> Result<Vec<u8>, DataSpaceError> {
    validate_drm_header(&value.header)?;
    if value
        .publishing_license
        .as_deref()
        .is_none_or(str::is_empty)
    {
        return Err(invalid("IRM publishing license cannot be null or empty"));
    }
    validate_inert_xml(
        value
            .publishing_license
            .as_deref()
            .expect("presence checked"),
        "IRM publishing license",
    )?;
    let mut output = write_transform_header(&value.header)?;
    output.extend_from_slice(&EXTENSIBILITY_HEADER_LENGTH.to_le_bytes());
    write_utf8_lpp4(&mut output, value.publishing_license.as_deref())?;
    Ok(output)
}

pub fn parse_encryption_transform(data: &[u8]) -> Result<EncryptionTransformInfo, DataSpaceError> {
    let (header, consumed) = parse_transform_header(data)?;
    validate_encryption_header(&header)?;
    let mut reader = SliceReader::at(data, consumed)?;
    let value = EncryptionTransformInfo {
        header,
        encryption_name: reader.utf8_lpp4()?,
        encryption_block_size: reader.u32()?,
        cipher_mode: reader.u32()?,
    };
    require_u32(reader.u32()?, 4, "EncryptionTransformInfo.Reserved")?;
    reader.finish()?;
    validate_encryption_transform(&value)?;
    Ok(value)
}

pub fn write_encryption_transform(
    value: &EncryptionTransformInfo,
) -> Result<Vec<u8>, DataSpaceError> {
    validate_encryption_transform(value)?;
    let mut output = write_transform_header(&value.header)?;
    write_utf8_lpp4(&mut output, value.encryption_name.as_deref())?;
    output.extend_from_slice(&value.encryption_block_size.to_le_bytes());
    output.extend_from_slice(&value.cipher_mode.to_le_bytes());
    output.extend_from_slice(&4u32.to_le_bytes());
    Ok(output)
}

pub fn parse_end_user_license(
    stream_name: &str,
    data: &[u8],
) -> Result<IrmEndUserLicense, DataSpaceError> {
    validate_eul_stream_name(stream_name)?;
    let mut reader = SliceReader::new(data)?;
    let header_start = reader.position();
    let header_length = usize::try_from(reader.u32()?)
        .map_err(|_| invalid("EndUserLicenseHeader.Length overflows usize"))?;
    if header_length < 8 || header_length > data.len() {
        return Err(invalid("invalid EndUserLicenseHeader.Length"));
    }
    let encoded_license_id = reader
        .utf8_lpp4()?
        .ok_or_else(|| invalid("EndUserLicenseHeader.ID_String cannot be null"))?;
    if reader.position() != header_start + header_length {
        return Err(invalid(
            "EndUserLicenseHeader.Length does not match ID_String",
        ));
    }
    let certificate_chain = reader.utf8_lpp4()?;
    if certificate_chain.as_deref().is_none_or(str::is_empty) {
        return Err(invalid(
            "end-user license certificate chain cannot be null or empty",
        ));
    }
    validate_inert_xml(
        certificate_chain.as_deref().expect("presence checked"),
        "end-user license certificate chain",
    )?;
    reader.finish()?;
    Ok(IrmEndUserLicense {
        stream_name: stream_name.to_string(),
        encoded_license_id,
        certificate_chain,
    })
}

pub fn write_end_user_license(value: &IrmEndUserLicense) -> Result<Vec<u8>, DataSpaceError> {
    validate_eul_stream_name(&value.stream_name)?;
    if value.encoded_license_id.is_empty() {
        return Err(invalid("EndUserLicenseHeader.ID_String cannot be empty"));
    }
    if value.certificate_chain.as_deref().is_none_or(str::is_empty) {
        return Err(invalid(
            "end-user license certificate chain cannot be null or empty",
        ));
    }
    validate_inert_xml(
        value
            .certificate_chain
            .as_deref()
            .expect("presence checked"),
        "end-user license certificate chain",
    )?;
    let mut output = vec![0; 4];
    write_utf8_lpp4(&mut output, Some(&value.encoded_license_id))?;
    let header_length = u32::try_from(output.len())
        .map_err(|_| invalid("EndUserLicenseHeader.Length exceeds u32"))?;
    output[..4].copy_from_slice(&header_length.to_le_bytes());
    write_utf8_lpp4(&mut output, value.certificate_chain.as_deref())?;
    Ok(output)
}

/// Inspect and cross-validate a complete DataSpaces graph in an OLE file.
pub fn inspect_data_spaces<R: Read + Seek>(
    ole: &mut OleFile<R>,
) -> Result<Option<DataSpaceGraph>, DataSpaceError> {
    let custom_xml_data_store = inspect_mso_data_store(ole)
        .map_err(|error| invalid(format!("MsoDataStore validation failed: {error}")))?;
    if !ole.exists(&[DATA_SPACES_STORAGE]) {
        validate_custom_xml_promotion(custom_xml_data_store.as_ref(), None)?;
        return Ok(None);
    }
    let version = parse_version_info(&read_stream(ole, &[DATA_SPACES_STORAGE, "Version"])?)?;
    let map = parse_data_space_map(&read_stream(ole, &[DATA_SPACES_STORAGE, "DataSpaceMap"])?)?;

    let definition_entries = ole
        .list_directory_entries(&[DATA_SPACES_STORAGE, "DataSpaceInfo"])
        .map_err(ole_error)?;
    if definition_entries.len() > MAX_ENTRIES {
        return Err(invalid("too many DataSpaceInfo entries"));
    }
    let mut definition_names = Vec::with_capacity(definition_entries.len());
    for entry in definition_entries {
        if entry.entry_type != 2 {
            return Err(invalid(format!(
                "DataSpaceInfo child '{}' is not a stream",
                entry.name
            )));
        }
        definition_names.push(entry.name.clone());
    }
    definition_names.sort();
    let mut definitions = Vec::with_capacity(definition_names.len());
    for name in definition_names {
        definitions.push(NamedDataSpaceDefinition {
            definition: parse_data_space_definition(&read_stream(
                ole,
                &[DATA_SPACES_STORAGE, "DataSpaceInfo", &name],
            )?)?,
            name,
        });
    }

    let transform_entries = ole
        .list_directory_entries(&[DATA_SPACES_STORAGE, "TransformInfo"])
        .map_err(ole_error)?;
    if transform_entries.len() > MAX_ENTRIES {
        return Err(invalid("too many TransformInfo entries"));
    }
    let mut transform_names = Vec::with_capacity(transform_entries.len());
    for entry in transform_entries {
        if entry.entry_type != 1 {
            // LabelInfo is a permitted stream sibling, not a transform.
            if entry.entry_type == 2 && entry.name == "LabelInfo" {
                continue;
            }
            return Err(invalid(format!(
                "TransformInfo child '{}' is not a storage",
                entry.name
            )));
        }
        transform_names.push(entry.name.clone());
    }
    transform_names.sort();
    let mut transforms = Vec::with_capacity(transform_names.len());
    for name in transform_names {
        let child_entries = ole
            .list_directory_entries(&[DATA_SPACES_STORAGE, "TransformInfo", &name])
            .map_err(ole_error)?
            .iter()
            .map(|entry| (entry.name.clone(), entry.entry_type))
            .collect::<Vec<_>>();
        if child_entries.len() > MAX_ENTRIES {
            return Err(invalid("too many transform-storage entries"));
        }
        let bytes = read_stream(
            ole,
            &[DATA_SPACES_STORAGE, "TransformInfo", &name, PRIMARY_STREAM],
        )?;
        let (header, consumed) = parse_transform_header(&bytes)?;
        let irm = if header.transform_id == DRM_TRANSFORM_ID
            && header.transform_name == DRM_TRANSFORM_NAME
        {
            Some(parse_irm_transform(&bytes)?)
        } else {
            None
        };
        let encryption = if header.transform_id == ENCRYPTION_TRANSFORM_ID
            && header.transform_name == ENCRYPTION_TRANSFORM_NAME
        {
            Some(parse_encryption_transform(&bytes)?)
        } else {
            None
        };
        let parsed_known_transform = irm.is_some() || encryption.is_some();
        let mut end_user_licenses = Vec::new();
        for (child_name, entry_type) in child_entries {
            if child_name == PRIMARY_STREAM {
                if entry_type != 2 {
                    return Err(invalid("transform Primary entry is not a stream"));
                }
                continue;
            }
            if child_name.starts_with("EUL-") {
                if entry_type != 2 {
                    return Err(invalid("end-user license entry is not a stream"));
                }
                end_user_licenses.push(parse_end_user_license(
                    &child_name,
                    &read_stream(
                        ole,
                        &[DATA_SPACES_STORAGE, "TransformInfo", &name, &child_name],
                    )?,
                )?);
            } else {
                return Err(invalid(format!(
                    "unexpected transform-storage entry '{child_name}'"
                )));
            }
        }
        if irm.is_some() && end_user_licenses.is_empty() {
            return Err(invalid(format!(
                "IRM transform '{name}' has no end-user license stream"
            )));
        }
        transforms.push(DataSpaceTransform {
            name,
            header,
            irm,
            encryption,
            end_user_licenses,
            opaque_tail: if parsed_known_transform {
                Vec::new()
            } else {
                bytes[consumed..].to_vec()
            },
        });
    }

    validate_graph(ole, &map, &definitions, &transforms)?;
    let irm = classify_irm(&map, &definitions, &transforms)?;
    let (label_info, sensitivity_labels) =
        if ole.exists(&[DATA_SPACES_STORAGE, "TransformInfo", "LabelInfo"]) {
            let bytes = read_stream(ole, &[DATA_SPACES_STORAGE, "TransformInfo", "LabelInfo"])?;
            let labels = parse_label_info(&bytes)
                .map_err(|error| invalid(format!("LabelInfo validation failed: {error}")))?;
            (Some(bytes), Some(labels))
        } else {
            (None, None)
        };
    if sensitivity_labels.is_some()
        && !irm.as_ref().is_some_and(|profile| {
            transforms.iter().any(|transform| {
                transform.name == profile.transform_name
                    && transform
                        .irm
                        .as_ref()
                        .is_some_and(|metadata| metadata.publishing_license.is_some())
            })
        })
    {
        return Err(invalid(
            "LabelInfo requires an IRM transform with a publishing license",
        ));
    }
    let summary_information_integrity = inspect_property_integrity(
        ole,
        ENCRYPTED_SUMMARY_INFORMATION_HASH_STREAM,
        SUMMARY_INFORMATION_STREAM,
    )?;
    let document_summary_information_integrity = inspect_property_integrity(
        ole,
        ENCRYPTED_DOCUMENT_SUMMARY_INFORMATION_HASH_STREAM,
        DOCUMENT_SUMMARY_INFORMATION_STREAM,
    )?;
    if (summary_information_integrity.is_some() || document_summary_information_integrity.is_some())
        && irm.is_none()
        && !transforms
            .iter()
            .any(|transform| transform.encryption.is_some())
    {
        return Err(invalid(
            "encrypted property hash stream is present without an encryption or IRM transform",
        ));
    }
    validate_custom_xml_promotion(custom_xml_data_store.as_ref(), irm.as_ref())?;
    Ok(Some(DataSpaceGraph {
        version,
        map,
        definitions,
        transforms,
        irm,
        label_info,
        sensitivity_labels,
        summary_information_integrity,
        document_summary_information_integrity,
        custom_xml_data_store,
    }))
}

fn validate_custom_xml_promotion(
    store: Option<&MsoDataStore>,
    irm: Option<&IrmDataSpace>,
) -> Result<(), DataSpaceError> {
    if store.is_some_and(|store| store.promotion != DataStorePromotion::Unspecified)
        && irm.is_none()
    {
        return Err(invalid(
            "MsoDataStore promotion marker requires an IRM data space",
        ));
    }
    Ok(())
}

/// Open an OLE compound file and inspect its DataSpaces graph.
pub fn inspect_data_spaces_bytes(bytes: &[u8]) -> Result<Option<DataSpaceGraph>, DataSpaceError> {
    let mut ole = OleFile::open(std::io::Cursor::new(bytes)).map_err(ole_error)?;
    inspect_data_spaces(&mut ole)
}

fn inspect_property_integrity<R: Read + Seek>(
    ole: &mut OleFile<R>,
    info_stream: &str,
    property_stream: &str,
) -> Result<Option<PropertyStreamIntegrity>, DataSpaceError> {
    if !ole.exists(&[info_stream]) {
        return Ok(None);
    }
    if !ole.exists(&[property_stream]) {
        return Err(invalid(format!(
            "{info_stream} is present without {property_stream}"
        )));
    }
    let info = parse_encrypted_property_stream_info(&read_stream(ole, &[info_stream])?)
        .map_err(|error| invalid(format!("{info_stream} is malformed: {error}")))?;
    let checksum_matches = checksum_matches(&info, &read_stream(ole, &[property_stream])?);
    Ok(Some(PropertyStreamIntegrity {
        info,
        checksum_matches,
    }))
}

fn validate_graph<R: Read + Seek>(
    ole: &OleFile<R>,
    map: &DataSpaceMap,
    definitions: &[NamedDataSpaceDefinition],
    transforms: &[DataSpaceTransform],
) -> Result<(), DataSpaceError> {
    if definitions.len() != map.entries.len()
        || definitions.iter().any(|definition| {
            !map.entries
                .iter()
                .any(|entry| entry.data_space_name == definition.name)
        })
    {
        return Err(invalid(
            "DataSpaceInfo streams do not correspond one-to-one with map entries",
        ));
    }
    let root_entries = ole.list_directory_entries(&[]).map_err(ole_error)?;
    let root_types = root_entries
        .iter()
        .map(|entry| (entry.name.as_str(), entry.entry_type))
        .collect::<std::collections::HashMap<_, _>>();
    for entry in &map.entries {
        let definition = definitions
            .iter()
            .find(|definition| definition.name == entry.data_space_name)
            .ok_or_else(|| {
                invalid(format!(
                    "map references missing data space '{}'",
                    entry.data_space_name
                ))
            })?;
        for reference in &entry.references {
            let component_type = root_types
                .get(reference.component.as_str())
                .ok_or_else(|| {
                    invalid(format!(
                        "map references missing root component '{}'",
                        reference.component
                    ))
                })?;
            let expected_type = match reference.kind {
                DataSpaceReferenceKind::Stream => 2,
                DataSpaceReferenceKind::Storage => 1,
            };
            if *component_type != expected_type {
                return Err(invalid(format!(
                    "root component '{}' has the wrong reference kind",
                    reference.component
                )));
            }
        }
        for transform_name in &definition.definition.transforms {
            if !transforms
                .iter()
                .any(|transform| transform.name == *transform_name)
            {
                return Err(invalid(format!(
                    "data space '{}' references missing transform '{}'",
                    definition.name, transform_name
                )));
            }
        }
    }
    Ok(())
}

fn classify_irm(
    map: &DataSpaceMap,
    definitions: &[NamedDataSpaceDefinition],
    transforms: &[DataSpaceTransform],
) -> Result<Option<IrmDataSpace>, DataSpaceError> {
    if let Some(entry) = map
        .entries
        .iter()
        .find(|entry| entry.data_space_name == "DRMEncryptedDataSpace")
    {
        if map.entries.len() != 1 {
            return Err(invalid("OOXML IRM requires exactly one DataSpaceMap entry"));
        }
        require_single_stream(entry, "EncryptedPackage")?;
        require_definition(
            definitions,
            "DRMEncryptedDataSpace",
            &["DRMEncryptedTransform"],
        )?;
        require_drm_transform(transforms, "DRMEncryptedTransform")?;
        return Ok(Some(IrmDataSpace {
            document_kind: IrmDocumentKind::Ooxml,
            protected_content_stream: "EncryptedPackage".to_string(),
            viewer_content_stream: None,
            transform_name: "DRMEncryptedTransform".to_string(),
        }));
    }
    if let Some(entry) = map
        .entries
        .iter()
        .find(|entry| entry.data_space_name == "0x09DRMDataSpace")
    {
        require_single_stream(entry, "0x09DRMContent")?;
        require_definition(definitions, "0x09DRMDataSpace", &["0x09DRMTransform"])?;
        require_drm_transform(transforms, "0x09DRMTransform")?;
        let viewer_content_stream = if let Some(viewer) = map
            .entries
            .iter()
            .find(|candidate| candidate.data_space_name == "0x09LZXDRMDataSpace")
        {
            require_single_stream(viewer, "0x09DRMViewerContent")?;
            require_definition(
                definitions,
                "0x09LZXDRMDataSpace",
                &["0x09DRMTransform", "0x09LZXTransform"],
            )?;
            require_named_transform(
                transforms,
                "0x09LZXTransform",
                LZX_TRANSFORM_ID,
                LZX_TRANSFORM_NAME,
            )?;
            Some("0x09DRMViewerContent".to_string())
        } else {
            None
        };
        let expected_entry_count = if viewer_content_stream.is_some() {
            2
        } else {
            1
        };
        if map.entries.len() != expected_entry_count {
            return Err(invalid(
                "binary IRM contains an unexpected DataSpaceMap entry",
            ));
        }
        return Ok(Some(IrmDataSpace {
            document_kind: IrmDocumentKind::LegacyBinary,
            protected_content_stream: "0x09DRMContent".to_string(),
            viewer_content_stream,
            transform_name: "0x09DRMTransform".to_string(),
        }));
    }
    Ok(None)
}

fn require_single_stream(entry: &DataSpaceMapEntry, expected: &str) -> Result<(), DataSpaceError> {
    if entry.references.as_slice()
        != [DataSpaceReference {
            kind: DataSpaceReferenceKind::Stream,
            component: expected.to_string(),
        }]
    {
        return Err(invalid(format!(
            "IRM data space '{}' does not reference exactly '{expected}'",
            entry.data_space_name
        )));
    }
    Ok(())
}

fn require_definition(
    definitions: &[NamedDataSpaceDefinition],
    name: &str,
    expected: &[&str],
) -> Result<(), DataSpaceError> {
    let definition = definitions
        .iter()
        .find(|definition| definition.name == name)
        .ok_or_else(|| invalid(format!("missing IRM data space '{name}'")))?;
    if definition
        .definition
        .transforms
        .iter()
        .map(String::as_str)
        .ne(expected.iter().copied())
    {
        return Err(invalid(format!(
            "IRM data space '{name}' has the wrong transform chain"
        )));
    }
    Ok(())
}

fn require_drm_transform(
    transforms: &[DataSpaceTransform],
    name: &str,
) -> Result<(), DataSpaceError> {
    let transform = transforms
        .iter()
        .find(|transform| transform.name == name)
        .ok_or_else(|| invalid(format!("missing IRM transform '{name}'")))?;
    validate_drm_header(&transform.header)
}

fn require_named_transform(
    transforms: &[DataSpaceTransform],
    name: &str,
    transform_id: &str,
    transform_name: &str,
) -> Result<(), DataSpaceError> {
    let transform = transforms
        .iter()
        .find(|transform| transform.name == name)
        .ok_or_else(|| invalid(format!("missing transform '{name}'")))?;
    if transform.header.transform_id != transform_id
        || transform.header.transform_name != transform_name
        || transform.header.reader != DataSpaceVersion::V1_0
        || transform.header.updater != DataSpaceVersion::V1_0
        || transform.header.writer != DataSpaceVersion::V1_0
    {
        return Err(invalid(format!("transform '{name}' has an invalid header")));
    }
    Ok(())
}

fn validate_version_info(value: &DataSpaceVersionInfo) -> Result<(), DataSpaceError> {
    if value.feature_identifier != DATA_SPACES_FEATURE
        || value.reader != DataSpaceVersion::V1_0
        || value.updater != DataSpaceVersion::V1_0
        || value.writer != DataSpaceVersion::V1_0
    {
        return Err(invalid("unsupported DataSpaceVersionInfo"));
    }
    Ok(())
}

fn validate_map(value: &DataSpaceMap) -> Result<(), DataSpaceError> {
    if value.entries.is_empty() || value.entries.len() > MAX_ENTRIES {
        return Err(invalid("DataSpaceMap entry count is out of bounds"));
    }
    let mut names = std::collections::HashSet::with_capacity(value.entries.len());
    for entry in &value.entries {
        validate_name(&entry.data_space_name, "data space name")?;
        if !names.insert(entry.data_space_name.as_str()) {
            return Err(invalid("duplicate data space name"));
        }
        if entry.references.is_empty() || entry.references.len() > MAX_ENTRIES {
            return Err(invalid("reference component count is out of bounds"));
        }
        for reference in &entry.references {
            validate_name(&reference.component, "reference component")?;
        }
    }
    Ok(())
}

fn validate_definition(value: &DataSpaceDefinition) -> Result<(), DataSpaceError> {
    if value.transforms.is_empty() || value.transforms.len() > MAX_ENTRIES {
        return Err(invalid("transform reference count is out of bounds"));
    }
    let mut names = std::collections::HashSet::with_capacity(value.transforms.len());
    for transform in &value.transforms {
        validate_name(transform, "transform reference")?;
        if !names.insert(transform.as_str()) {
            return Err(invalid("duplicate transform reference"));
        }
    }
    Ok(())
}

fn validate_transform_header(value: &TransformInfoHeader) -> Result<(), DataSpaceError> {
    validate_name(&value.transform_id, "transform identifier")?;
    validate_name(&value.transform_name, "transform name")?;
    Ok(())
}

fn validate_drm_header(value: &TransformInfoHeader) -> Result<(), DataSpaceError> {
    if value.transform_id != DRM_TRANSFORM_ID
        || value.transform_name != DRM_TRANSFORM_NAME
        || value.reader != DataSpaceVersion::V1_0
        || value.updater != DataSpaceVersion::V1_0
        || value.writer != DataSpaceVersion::V1_0
    {
        return Err(invalid("invalid IRM transform header"));
    }
    Ok(())
}

fn validate_encryption_header(value: &TransformInfoHeader) -> Result<(), DataSpaceError> {
    if value.transform_id != ENCRYPTION_TRANSFORM_ID
        || value.transform_name != ENCRYPTION_TRANSFORM_NAME
        || value.reader != DataSpaceVersion::V1_0
        || value.updater != DataSpaceVersion::V1_0
        || value.writer != DataSpaceVersion::V1_0
    {
        return Err(invalid("invalid encryption transform header"));
    }
    Ok(())
}

fn validate_encryption_transform(value: &EncryptionTransformInfo) -> Result<(), DataSpaceError> {
    validate_encryption_header(&value.header)?;
    if value.encryption_block_size == 0 {
        return Err(invalid("encryption transform block size cannot be zero"));
    }
    if value
        .encryption_name
        .as_deref()
        .is_some_and(|name| name.is_empty() || name.len() > MAX_STRING_BYTES)
    {
        return Err(invalid("invalid encryption transform algorithm name"));
    }
    Ok(())
}

fn validate_name(value: &str, label: &str) -> Result<(), DataSpaceError> {
    if value.is_empty()
        || value.len() > MAX_STRING_BYTES
        || value.chars().any(|character| character == '\0')
    {
        return Err(invalid(format!(
            "{label} is empty, too long, or contains NUL"
        )));
    }
    Ok(())
}

fn validate_eul_stream_name(value: &str) -> Result<(), DataSpaceError> {
    let Some(encoded_guid) = value.strip_prefix("EUL-") else {
        return Err(invalid("end-user license stream name lacks EUL- prefix"));
    };
    if encoded_guid.len() != 26
        || !encoded_guid
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid(
            "end-user license stream name does not contain a 26-character base-32 GUID",
        ));
    }
    Ok(())
}

fn validate_inert_xml(value: &str, label: &str) -> Result<(), DataSpaceError> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_str(value);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                if depth == 0 {
                    roots += 1;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid(format!("{label} XML depth overflow")))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid(format!("{label} XML is too deeply nested")));
                }
            },
            Ok(Event::Empty(_)) => {
                if depth == 0 {
                    roots += 1;
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid(format!("{label} XML is unbalanced")))?;
            },
            Ok(Event::DocType(_)) => {
                return Err(invalid(format!("{label} XML contains a forbidden DTD")));
            },
            Ok(Event::Text(text)) if depth == 0 => {
                if !text
                    .decode()
                    .map_err(|error| invalid(format!("{label} XML text is invalid: {error}")))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid(format!("{label} XML has text outside its root")));
                }
            },
            Ok(Event::CData(text)) if depth == 0 => {
                if !text
                    .decode()
                    .map_err(|error| invalid(format!("{label} XML CDATA is invalid: {error}")))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid(format!("{label} XML has CDATA outside its root")));
                }
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(invalid(format!("{label} XML is malformed: {error}"))),
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid(format!(
            "{label} XML must contain exactly one complete root"
        )));
    }
    Ok(())
}

fn bounded_count(value: u32, label: &str) -> Result<usize, DataSpaceError> {
    let value = usize::try_from(value).map_err(|_| invalid(format!("{label} overflows usize")))?;
    if value > MAX_ENTRIES {
        return Err(invalid(format!("{label} exceeds {MAX_ENTRIES}")));
    }
    Ok(value)
}

fn write_count(output: &mut Vec<u8>, count: usize, label: &str) -> Result<(), DataSpaceError> {
    if count > MAX_ENTRIES {
        return Err(invalid(format!("{label} exceeds {MAX_ENTRIES}")));
    }
    output.extend_from_slice(
        &u32::try_from(count)
            .map_err(|_| invalid(format!("{label} exceeds u32")))?
            .to_le_bytes(),
    );
    Ok(())
}

fn require_u32(value: u32, expected: u32, label: &str) -> Result<(), DataSpaceError> {
    if value != expected {
        return Err(invalid(format!(
            "{label} is {value:#010X}, expected {expected:#010X}"
        )));
    }
    Ok(())
}

fn write_version(output: &mut Vec<u8>, version: DataSpaceVersion) {
    output.extend_from_slice(&version.major.to_le_bytes());
    output.extend_from_slice(&version.minor.to_le_bytes());
}

fn write_unicode_lpp4(output: &mut Vec<u8>, value: &str) -> Result<(), DataSpaceError> {
    validate_name(value, "UNICODE-LP-P4 string")?;
    let units = value.encode_utf16().collect::<Vec<_>>();
    let byte_len = units
        .len()
        .checked_mul(2)
        .ok_or_else(|| invalid("UNICODE-LP-P4 length overflow"))?;
    output.extend_from_slice(
        &u32::try_from(byte_len)
            .map_err(|_| invalid("UNICODE-LP-P4 length exceeds u32"))?
            .to_le_bytes(),
    );
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    if byte_len % 4 == 2 {
        output.extend_from_slice(&[0, 0]);
    }
    Ok(())
}

fn write_utf8_lpp4(output: &mut Vec<u8>, value: Option<&str>) -> Result<(), DataSpaceError> {
    let Some(value) = value else {
        output.extend_from_slice(&0u32.to_le_bytes());
        return Ok(());
    };
    if value.len() > MAX_STRING_BYTES || value.contains('\0') {
        return Err(invalid("UTF-8-LP-P4 string is too long or contains NUL"));
    }
    output.extend_from_slice(
        &u32::try_from(value.len())
            .map_err(|_| invalid("UTF-8-LP-P4 length exceeds u32"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(value.as_bytes());
    output.resize(output.len().next_multiple_of(4), 0);
    Ok(())
}

fn read_stream<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
) -> Result<Vec<u8>, DataSpaceError> {
    let bytes = ole.open_stream(path).map_err(ole_error)?;
    if bytes.len() > MAX_STREAM_BYTES {
        return Err(invalid(format!(
            "stream '{}' exceeds {MAX_STREAM_BYTES} bytes",
            path.join("/")
        )));
    }
    Ok(bytes)
}

fn ole_error(error: impl fmt::Display) -> DataSpaceError {
    DataSpaceError::Ole(error.to_string())
}

fn invalid(message: impl Into<String>) -> DataSpaceError {
    DataSpaceError::Invalid(message.into())
}

struct SliceReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> SliceReader<'a> {
    fn new(data: &'a [u8]) -> Result<Self, DataSpaceError> {
        if data.len() > MAX_STREAM_BYTES {
            return Err(invalid("DataSpaces stream exceeds parser limit"));
        }
        Ok(Self { data, offset: 0 })
    }

    fn at(data: &'a [u8], offset: usize) -> Result<Self, DataSpaceError> {
        let mut reader = Self::new(data)?;
        if offset > data.len() {
            return Err(invalid("parser offset exceeds stream"));
        }
        reader.offset = offset;
        Ok(reader)
    }

    fn position(&self) -> usize {
        self.offset
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8], DataSpaceError> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid("stream offset overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| invalid("truncated DataSpaces stream"))?;
        self.offset = end;
        Ok(bytes)
    }

    fn u16(&mut self) -> Result<u16, DataSpaceError> {
        Ok(u16::from_le_bytes(
            self.take(2)?.try_into().expect("two-byte slice returned"),
        ))
    }

    fn u32(&mut self) -> Result<u32, DataSpaceError> {
        Ok(u32::from_le_bytes(
            self.take(4)?.try_into().expect("four-byte slice returned"),
        ))
    }

    fn version(&mut self) -> Result<DataSpaceVersion, DataSpaceError> {
        Ok(DataSpaceVersion {
            major: self.u16()?,
            minor: self.u16()?,
        })
    }

    fn unicode_lpp4(&mut self) -> Result<String, DataSpaceError> {
        let byte_len = usize::try_from(self.u32()?)
            .map_err(|_| invalid("UNICODE-LP-P4 length overflows usize"))?;
        if byte_len == 0 || byte_len > MAX_STRING_BYTES || byte_len % 2 != 0 {
            return Err(invalid("invalid UNICODE-LP-P4 length"));
        }
        let bytes = self.take(byte_len)?;
        let mut units = Vec::with_capacity(byte_len / 2);
        for bytes in bytes.chunks_exact(2) {
            units.push(u16::from_le_bytes([bytes[0], bytes[1]]));
        }
        let value =
            String::from_utf16(&units).map_err(|_| invalid("invalid UNICODE-LP-P4 UTF-16"))?;
        if value.contains('\0') {
            return Err(invalid("UNICODE-LP-P4 string contains NUL"));
        }
        let padding = (4 - (byte_len % 4)) % 4;
        if self.take(padding)?.iter().any(|byte| *byte != 0) {
            return Err(invalid("UNICODE-LP-P4 padding is nonzero"));
        }
        Ok(value)
    }

    fn utf8_lpp4(&mut self) -> Result<Option<String>, DataSpaceError> {
        let byte_len = usize::try_from(self.u32()?)
            .map_err(|_| invalid("UTF-8-LP-P4 length overflows usize"))?;
        if byte_len == 0 {
            return Ok(None);
        }
        if byte_len > MAX_STRING_BYTES {
            return Err(invalid("UTF-8-LP-P4 length exceeds parser limit"));
        }
        let bytes = self.take(byte_len)?;
        let value = std::str::from_utf8(bytes).map_err(|_| invalid("invalid UTF-8-LP-P4 UTF-8"))?;
        if value.contains('\0') {
            return Err(invalid("UTF-8-LP-P4 string contains NUL"));
        }
        let padding = (4 - (byte_len % 4)) % 4;
        if self.take(padding)?.iter().any(|byte| *byte != 0) {
            return Err(invalid("UTF-8-LP-P4 padding is nonzero"));
        }
        Ok(Some(value.to_string()))
    }

    fn finish(self) -> Result<(), DataSpaceError> {
        if self.offset == self.data.len() {
            Ok(())
        } else {
            Err(invalid("trailing bytes in DataSpaces stream"))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::OleWriter;
    use std::io::Cursor;

    fn drm_header() -> TransformInfoHeader {
        TransformInfoHeader {
            transform_id: DRM_TRANSFORM_ID.to_string(),
            transform_name: DRM_TRANSFORM_NAME.to_string(),
            reader: DataSpaceVersion::V1_0,
            updater: DataSpaceVersion::V1_0,
            writer: DataSpaceVersion::V1_0,
        }
    }

    #[test]
    fn core_streams_round_trip() {
        let version = DataSpaceVersionInfo::default();
        assert_eq!(
            parse_version_info(&write_version_info(&version).unwrap()).unwrap(),
            version
        );
        let map = DataSpaceMap {
            entries: vec![DataSpaceMapEntry {
                references: vec![DataSpaceReference {
                    kind: DataSpaceReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "DRMEncryptedDataSpace".to_string(),
            }],
        };
        assert_eq!(
            parse_data_space_map(&write_data_space_map(&map).unwrap()).unwrap(),
            map
        );
        let definition = DataSpaceDefinition {
            transforms: vec!["DRMEncryptedTransform".to_string()],
        };
        assert_eq!(
            parse_data_space_definition(&write_data_space_definition(&definition).unwrap())
                .unwrap(),
            definition
        );
    }

    #[test]
    fn irm_transform_round_trip_preserves_license() {
        let transform = IrmTransformInfo {
            header: drm_header(),
            publishing_license: Some("<XrML>inert</XrML>".to_string()),
        };
        assert_eq!(
            parse_irm_transform(&write_irm_transform(&transform).unwrap()).unwrap(),
            transform
        );
    }

    #[test]
    fn encryption_transform_round_trip_preserves_typed_parameters() {
        let transform = EncryptionTransformInfo {
            header: TransformInfoHeader {
                transform_id: ENCRYPTION_TRANSFORM_ID.to_string(),
                transform_name: ENCRYPTION_TRANSFORM_NAME.to_string(),
                reader: DataSpaceVersion::V1_0,
                updater: DataSpaceVersion::V1_0,
                writer: DataSpaceVersion::V1_0,
            },
            encryption_name: None,
            encryption_block_size: 16,
            cipher_mode: 0,
        };
        assert_eq!(
            parse_encryption_transform(&write_encryption_transform(&transform).unwrap()).unwrap(),
            transform
        );
    }

    #[test]
    fn end_user_license_round_trip_preserves_inert_xml() {
        let license = IrmEndUserLicense {
            stream_name: "EUL-ETRHA1143ZLUDD412YTI3M5CTZ".to_string(),
            encoded_license_id: "VwBpAG4AZABvAHcAOgB1AHMAZQByAEA".to_string(),
            certificate_chain: Some("<?xml version=\"1.0\"?><certificatechain/>".to_string()),
        };
        assert_eq!(
            parse_end_user_license(
                &license.stream_name,
                &write_end_user_license(&license).unwrap()
            )
            .unwrap(),
            license
        );
    }

    #[test]
    fn rejects_malformed_lengths_padding_counts_and_drm_identity() {
        let mut version = write_version_info(&DataSpaceVersionInfo::default()).unwrap();
        version[0] = 3;
        assert!(parse_version_info(&version).is_err());

        let map = DataSpaceMap {
            entries: Vec::new(),
        };
        assert!(write_data_space_map(&map).is_err());

        let mut transform = write_irm_transform(&IrmTransformInfo {
            header: drm_header(),
            publishing_license: Some("<XrML/>".to_string()),
        })
        .unwrap();
        transform[4] = 2;
        assert!(parse_irm_transform(&transform).is_err());
        assert!(
            write_irm_transform(&IrmTransformInfo {
                header: drm_header(),
                publishing_license: Some("<!DOCTYPE x><x/>".to_string()),
            })
            .is_err()
        );
    }

    #[test]
    fn inspects_and_classifies_complete_ooxml_irm_graph() {
        let map = write_data_space_map(&DataSpaceMap {
            entries: vec![DataSpaceMapEntry {
                references: vec![DataSpaceReference {
                    kind: DataSpaceReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "DRMEncryptedDataSpace".to_string(),
            }],
        })
        .unwrap();
        let definition = write_data_space_definition(&DataSpaceDefinition {
            transforms: vec!["DRMEncryptedTransform".to_string()],
        })
        .unwrap();
        let primary = write_irm_transform(&IrmTransformInfo {
            header: drm_header(),
            publishing_license: Some("<XrML/>".to_string()),
        })
        .unwrap();
        let end_user_license = IrmEndUserLicense {
            stream_name: "EUL-ETRHA1143ZLUDD412YTI3M5CTZ".to_string(),
            encoded_license_id: "VwBpAG4AZABvAHcAOgB1AHMAZQByAEA".to_string(),
            certificate_chain: Some("<certificatechain/>".to_string()),
        };
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptedPackage"], &[0; 16])
            .unwrap();
        writer.create_storage(&[DATA_SPACES_STORAGE]).unwrap();
        writer
            .create_storage(&[DATA_SPACES_STORAGE, "DataSpaceInfo"])
            .unwrap();
        writer
            .create_storage(&[DATA_SPACES_STORAGE, "TransformInfo"])
            .unwrap();
        writer
            .create_storage(&[
                DATA_SPACES_STORAGE,
                "TransformInfo",
                "DRMEncryptedTransform",
            ])
            .unwrap();
        writer
            .create_stream(&[DATA_SPACES_STORAGE, "DataSpaceMap"], &map)
            .unwrap();
        writer
            .create_stream(
                &[
                    DATA_SPACES_STORAGE,
                    "DataSpaceInfo",
                    "DRMEncryptedDataSpace",
                ],
                &definition,
            )
            .unwrap();
        writer
            .create_stream(
                &[
                    DATA_SPACES_STORAGE,
                    "TransformInfo",
                    "DRMEncryptedTransform",
                    PRIMARY_STREAM,
                ],
                &primary,
            )
            .unwrap();
        writer
            .create_stream(
                &[
                    DATA_SPACES_STORAGE,
                    "TransformInfo",
                    "DRMEncryptedTransform",
                    &end_user_license.stream_name,
                ],
                &write_end_user_license(&end_user_license).unwrap(),
            )
            .unwrap();
        writer
            .create_stream(
                &[DATA_SPACES_STORAGE, "Version"],
                &write_version_info(&DataSpaceVersionInfo::default()).unwrap(),
            )
            .unwrap();
        let label_info = format!(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><clbl:labelList xmlns:clbl=\"{}\"><clbl:label id=\"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee\" enabled=\"1\" method=\"Standard\" siteId=\"12345678-1234-5678-90ab-1234567890ab\" contentBits=\"8\" removed=\"0\"/></clbl:labelList>",
            crate::office_crypto::sensitivity_labels::LABEL_INFO_NAMESPACE
        );
        writer
            .create_stream(
                &[DATA_SPACES_STORAGE, "TransformInfo", "LabelInfo"],
                label_info.as_bytes(),
            )
            .unwrap();
        let summary_information = b"public summary property set";
        writer
            .create_stream(&[SUMMARY_INFORMATION_STREAM], summary_information)
            .unwrap();
        writer
            .create_stream(
                &[ENCRYPTED_SUMMARY_INFORMATION_HASH_STREAM],
                &crate::office_crypto::property_integrity::write_encrypted_property_stream_info(
                    crate::office_crypto::property_integrity::mso_crc32(summary_information),
                    &[],
                ),
            )
            .unwrap();
        let custom_properties = litchi_ole_common::custom_xml_data::CustomXmlDataProperties {
            item_id: "{11111111-2222-3333-4444-555555555555}".parse().unwrap(),
            schema_references: vec!["urn:test".to_string()],
        };
        let custom_item = litchi_ole_common::custom_xml_data::CustomXmlDataItem::new(
            custom_properties.item_id.storage_name(),
            br#"<test xmlns="urn:test"/>"#.to_vec(),
            custom_properties,
        )
        .unwrap();
        litchi_ole_common::custom_xml_data::write_mso_data_store(
            &mut writer,
            &litchi_ole_common::custom_xml_data::MsoDataStore::new(
                litchi_ole_common::custom_xml_data::DataStorePromotion::Modified,
                vec![custom_item],
            )
            .unwrap(),
        )
        .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();

        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        let graph = inspect_data_spaces(&mut ole).unwrap().unwrap();
        let irm = graph.irm.unwrap();
        assert_eq!(irm.document_kind, IrmDocumentKind::Ooxml);
        assert_eq!(irm.protected_content_stream, "EncryptedPackage");
        assert_eq!(graph.transforms[0].end_user_licenses, [end_user_license]);
        assert_eq!(graph.sensitivity_labels.unwrap().labels.len(), 1);
        assert_eq!(
            graph
                .summary_information_integrity
                .unwrap()
                .checksum_matches,
            Some(true)
        );
        assert!(graph.document_summary_information_integrity.is_none());
        let custom_xml = graph.custom_xml_data_store.unwrap();
        assert_eq!(
            custom_xml.promotion,
            litchi_ole_common::custom_xml_data::DataStorePromotion::Modified
        );
        assert_eq!(custom_xml.items().len(), 1);
    }

    #[test]
    fn classifies_binary_irm_with_optional_viewer_chain() {
        let map = DataSpaceMap {
            entries: vec![
                DataSpaceMapEntry {
                    references: vec![DataSpaceReference {
                        kind: DataSpaceReferenceKind::Stream,
                        component: "0x09DRMContent".to_string(),
                    }],
                    data_space_name: "0x09DRMDataSpace".to_string(),
                },
                DataSpaceMapEntry {
                    references: vec![DataSpaceReference {
                        kind: DataSpaceReferenceKind::Stream,
                        component: "0x09DRMViewerContent".to_string(),
                    }],
                    data_space_name: "0x09LZXDRMDataSpace".to_string(),
                },
            ],
        };
        let definitions = vec![
            NamedDataSpaceDefinition {
                name: "0x09DRMDataSpace".to_string(),
                definition: DataSpaceDefinition {
                    transforms: vec!["0x09DRMTransform".to_string()],
                },
            },
            NamedDataSpaceDefinition {
                name: "0x09LZXDRMDataSpace".to_string(),
                definition: DataSpaceDefinition {
                    transforms: vec![
                        "0x09DRMTransform".to_string(),
                        "0x09LZXTransform".to_string(),
                    ],
                },
            },
        ];
        let transforms = vec![
            DataSpaceTransform {
                name: "0x09DRMTransform".to_string(),
                header: drm_header(),
                irm: None,
                encryption: None,
                end_user_licenses: Vec::new(),
                opaque_tail: Vec::new(),
            },
            DataSpaceTransform {
                name: "0x09LZXTransform".to_string(),
                header: TransformInfoHeader {
                    transform_id: LZX_TRANSFORM_ID.to_string(),
                    transform_name: LZX_TRANSFORM_NAME.to_string(),
                    reader: DataSpaceVersion::V1_0,
                    updater: DataSpaceVersion::V1_0,
                    writer: DataSpaceVersion::V1_0,
                },
                irm: None,
                encryption: None,
                end_user_licenses: Vec::new(),
                opaque_tail: Vec::new(),
            },
        ];

        let irm = classify_irm(&map, &definitions, &transforms)
            .unwrap()
            .unwrap();
        assert_eq!(irm.document_kind, IrmDocumentKind::LegacyBinary);
        assert_eq!(
            irm.viewer_content_stream.as_deref(),
            Some("0x09DRMViewerContent")
        );
    }

    #[test]
    fn rejects_custom_xml_promotion_without_irm() {
        let store = litchi_ole_common::custom_xml_data::MsoDataStore::new(
            litchi_ole_common::custom_xml_data::DataStorePromotion::Redundant,
            Vec::new(),
        )
        .unwrap();
        assert!(validate_custom_xml_promotion(Some(&store), None).is_err());

        let mut writer = OleWriter::new();
        litchi_ole_common::custom_xml_data::write_mso_data_store(&mut writer, &store).unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        assert!(inspect_data_spaces(&mut ole).is_err());

        let unspecified = litchi_ole_common::custom_xml_data::MsoDataStore::default();
        assert!(validate_custom_xml_promotion(Some(&unspecified), None).is_ok());
    }
}

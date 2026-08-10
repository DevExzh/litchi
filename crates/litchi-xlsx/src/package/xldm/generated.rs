//! Typed, inert inspection of MS-XLDM section 2.4 system-generated files.
//!
//! Section 2.4 reuses the `.idf` and `.hidx` layouts. This module assigns
//! their generated roles and validates cross-file sets without decompressing
//! or interpreting any persisted integer array.

use std::collections::{HashMap, HashSet};
use std::error::Error;
use std::fmt;

use super::native::{HashIndexFile, IdfFile, parse_hash_index, parse_idf};
use super::{Storage, classify_generated_path};

const MAX_SYSTEM_GENERATED_FILES: usize = 65_536;

/// A section 2.4 structural validation error.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct GeneratedDataError(String);

impl GeneratedDataError {
    fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for GeneratedDataError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for GeneratedDataError {}

/// Result type for section 2.4 generated-file inspection.
pub type GeneratedDataResult<T> = Result<T, GeneratedDataError>;

/// The semantic role encoded by a section 2.4 generated filename.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum SystemGeneratedKind {
    PositionToIdentifier,
    IdentifierToPosition,
    RelationshipIndex,
    UserHierarchyChildCount,
    UserHierarchyFirstChildPosition,
    UserHierarchyMultilevelIdentifier,
    UserHierarchyParentPosition,
}

/// Compression state required by the selected section 2.4 representation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SystemGeneratedCompression {
    XmReNoSplit,
    Uncompressed,
}

/// The shared binary layout selected by the generated filename.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum SystemGeneratedData<'a> {
    Idf(IdfFile<'a>),
    HashIndex(HashIndexFile<'a>),
}

impl SystemGeneratedData<'_> {
    #[must_use]
    pub fn expected_compression(&self) -> SystemGeneratedCompression {
        match self {
            Self::Idf(_) => SystemGeneratedCompression::XmReNoSplit,
            Self::HashIndex(_) => SystemGeneratedCompression::Uncompressed,
        }
    }
}

/// A borrowed, typed system-generated storage member.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SystemGeneratedFile<'a> {
    pub storage_path: &'a str,
    pub kind: SystemGeneratedKind,
    pub object_key: String,
    pub version: u64,
    pub bytes: &'a [u8],
    pub data: SystemGeneratedData<'a>,
}

impl SystemGeneratedFile<'_> {
    #[must_use]
    pub fn expected_compression(&self) -> SystemGeneratedCompression {
        self.data.expected_compression()
    }
}

/// All recognized section 2.4 files in a validated MS-XLDM storage object.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SystemGeneratedModel<'a> {
    pub files: Vec<SystemGeneratedFile<'a>>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Layout {
    Idf,
    HashIndex,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct GeneratedName {
    kind: SystemGeneratedKind,
    object_key: String,
    version: u64,
    layout: Layout,
}

/// Parse one path/payload pair if its filename denotes a section 2.4 file.
/// Non-section-2.4 generated files return `Ok(None)`.
pub fn parse_system_generated_file<'a>(
    storage_path: &'a str,
    bytes: &'a [u8],
) -> GeneratedDataResult<Option<SystemGeneratedFile<'a>>> {
    classify_generated_path(storage_path).map_err(|error| {
        GeneratedDataError::new(format!("invalid generated path {storage_path}: {error}"))
    })?;
    let Some(name) = parse_generated_name(storage_path)? else {
        return Ok(None);
    };
    let data = match (name.kind, name.layout) {
        (SystemGeneratedKind::RelationshipIndex, Layout::Idf) => {
            let idf = parse_idf(bytes);
            let hash = parse_hash_index(bytes);
            match (idf, hash) {
                (Ok(idf), Err(_)) => SystemGeneratedData::Idf(idf),
                (Err(_), Ok(hash)) => SystemGeneratedData::HashIndex(hash),
                (Ok(_), Ok(_)) => {
                    return Err(GeneratedDataError::new(format!(
                        "ambiguous relationship index payload layout for {storage_path}"
                    )));
                },
                (Err(idf), Err(hash)) => {
                    return Err(GeneratedDataError::new(format!(
                        "invalid relationship index as .idf ({idf}) and .hidx ({hash})"
                    )));
                },
            }
        },
        (_, Layout::Idf) => SystemGeneratedData::Idf(parse_idf(bytes).map_err(|error| {
            GeneratedDataError::new(format!(
                "invalid generated .idf member {storage_path}: {error}"
            ))
        })?),
        (_, Layout::HashIndex) => {
            SystemGeneratedData::HashIndex(parse_hash_index(bytes).map_err(|error| {
                GeneratedDataError::new(format!(
                    "invalid generated .hidx member {storage_path}: {error}"
                ))
            })?)
        },
    };
    Ok(Some(SystemGeneratedFile {
        storage_path,
        kind: name.kind,
        object_key: name.object_key,
        version: name.version,
        bytes,
        data,
    }))
}

/// Discover section 2.4 members from the validated backup log, resolve each
/// path through the validated virtual directory, and enforce cross-file sets.
pub fn inspect_system_generated<'a>(
    storage: &'a Storage<'a>,
) -> GeneratedDataResult<SystemGeneratedModel<'a>> {
    let mut files = Vec::new();
    for group in &storage.backup_log.file_groups {
        for logged in &group.files {
            let path = logged.storage_path.as_str();
            if parse_generated_name(path)?.is_none() {
                continue;
            }
            if files.len() == MAX_SYSTEM_GENERATED_FILES {
                return Err(GeneratedDataError::new(
                    "too many system-generated data files",
                ));
            }
            let directory_index = storage
                .files
                .iter()
                .position(|entry| entry.path == path)
                .ok_or_else(|| {
                    GeneratedDataError::new(format!(
                        "logged generated member {path} is absent from the directory"
                    ))
                })?;
            let bytes = storage.file_payload(directory_index).ok_or_else(|| {
                GeneratedDataError::new(format!("cannot resolve generated member {path}"))
            })?;
            let file = parse_system_generated_file(path, bytes)?.ok_or_else(|| {
                GeneratedDataError::new("generated member classification changed")
            })?;
            files.push(file);
        }
    }
    validate_system_generated_files(&files)?;
    Ok(SystemGeneratedModel { files })
}

/// Validate section 2.4 relationships between already parsed generated files.
pub fn validate_system_generated_files(
    files: &[SystemGeneratedFile<'_>],
) -> GeneratedDataResult<()> {
    if files.len() > MAX_SYSTEM_GENERATED_FILES {
        return Err(GeneratedDataError::new(
            "too many system-generated data files",
        ));
    }
    let mut relationships = HashSet::new();
    let mut hierarchies: HashMap<&str, (u64, u8)> = HashMap::new();
    for file in files {
        match file.kind {
            SystemGeneratedKind::RelationshipIndex => {
                if !relationships.insert(file.object_key.as_str()) {
                    return Err(GeneratedDataError::new(format!(
                        "relationship {} has multiple index representations",
                        file.object_key
                    )));
                }
            },
            SystemGeneratedKind::UserHierarchyChildCount
            | SystemGeneratedKind::UserHierarchyFirstChildPosition
            | SystemGeneratedKind::UserHierarchyMultilevelIdentifier
            | SystemGeneratedKind::UserHierarchyParentPosition => {
                let bit = match file.kind {
                    SystemGeneratedKind::UserHierarchyChildCount => 0b0001,
                    SystemGeneratedKind::UserHierarchyFirstChildPosition => 0b0010,
                    SystemGeneratedKind::UserHierarchyMultilevelIdentifier => 0b0100,
                    SystemGeneratedKind::UserHierarchyParentPosition => 0b1000,
                    SystemGeneratedKind::PositionToIdentifier
                    | SystemGeneratedKind::IdentifierToPosition
                    | SystemGeneratedKind::RelationshipIndex => unreachable!(),
                };
                let entry = hierarchies
                    .entry(file.object_key.as_str())
                    .or_insert((file.version, 0));
                if entry.0 != file.version {
                    return Err(GeneratedDataError::new(format!(
                        "user hierarchy {} uses inconsistent file versions",
                        file.object_key
                    )));
                }
                if entry.1 & bit != 0 {
                    return Err(GeneratedDataError::new(format!(
                        "user hierarchy {} duplicates a generated role",
                        file.object_key
                    )));
                }
                entry.1 |= bit;
            },
            SystemGeneratedKind::PositionToIdentifier
            | SystemGeneratedKind::IdentifierToPosition => {},
        }
    }
    for (key, (_, roles)) in hierarchies {
        if roles != 0b1111 {
            return Err(GeneratedDataError::new(format!(
                "user hierarchy {key} does not contain exactly all four generated files"
            )));
        }
    }
    Ok(())
}

/// Reparse an inspected member and return its exact original bytes.
pub fn write_system_generated_file(file: &SystemGeneratedFile<'_>) -> GeneratedDataResult<Vec<u8>> {
    let reparsed = parse_system_generated_file(file.storage_path, file.bytes)?
        .ok_or_else(|| GeneratedDataError::new("file is not section 2.4 generated data"))?;
    if reparsed.kind != file.kind
        || reparsed.object_key != file.object_key
        || reparsed.version != file.version
    {
        return Err(GeneratedDataError::new(
            "generated filename metadata was mutated",
        ));
    }
    Ok(file.bytes.to_vec())
}

fn parse_generated_name(path: &str) -> GeneratedDataResult<Option<GeneratedName>> {
    let basename = path.rsplit('/').next().unwrap_or(path);
    let candidates = [
        (
            ".FIRST_CHILD_POS.",
            SystemGeneratedKind::UserHierarchyFirstChildPosition,
            "U$",
            false,
        ),
        (
            ".MULTI_LEVEL_ID.",
            SystemGeneratedKind::UserHierarchyMultilevelIdentifier,
            "U$",
            false,
        ),
        (
            ".CHILD_COUNT.",
            SystemGeneratedKind::UserHierarchyChildCount,
            "U$",
            false,
        ),
        (
            ".PARENT_POS.",
            SystemGeneratedKind::UserHierarchyParentPosition,
            "U$",
            false,
        ),
        (
            ".POS_TO_ID.",
            SystemGeneratedKind::PositionToIdentifier,
            "H$",
            false,
        ),
        (
            ".ID_TO_POS.",
            SystemGeneratedKind::IdentifierToPosition,
            "H$",
            false,
        ),
        (
            ".INDEX.",
            SystemGeneratedKind::RelationshipIndex,
            "R$",
            false,
        ),
    ];
    for (marker, kind, identity_marker, allow_hash) in candidates {
        let Some((prefix, suffix)) = basename.rsplit_once(marker) else {
            continue;
        };
        let (version_text, layout) = if let Some(version) = suffix.strip_suffix(".idf") {
            (version, Layout::Idf)
        } else if allow_hash {
            let Some(version) = suffix.strip_suffix(".hidx") else {
                return Err(GeneratedDataError::new(format!(
                    "invalid extension for section 2.4 role in {path}"
                )));
            };
            (version, Layout::HashIndex)
        } else {
            return Err(GeneratedDataError::new(format!(
                "section 2.4 role requires an .idf file in {path}"
            )));
        };
        if version_text.is_empty() || !version_text.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(GeneratedDataError::new(format!(
                "invalid generated file version in {path}"
            )));
        }
        let version = version_text.parse::<u64>().map_err(|_source| {
            GeneratedDataError::new(format!("generated file version overflows in {path}"))
        })?;
        let Some((ordinal, identity)) = prefix.split_once('.') else {
            return Err(GeneratedDataError::new(format!(
                "missing generated ordinal in {path}"
            )));
        };
        if ordinal.is_empty()
            || !ordinal.bytes().all(|byte| byte.is_ascii_digit())
            || !identity.starts_with(identity_marker)
        {
            return Err(GeneratedDataError::new(format!(
                "invalid section 2.4 object identity in {path}"
            )));
        }
        let parent = path.rsplit_once('/').map_or("", |(parent, _)| parent);
        let object_key = if parent.is_empty() {
            prefix.to_owned()
        } else {
            format!("{parent}/{prefix}")
        };
        return Ok(Some(GeneratedName {
            kind,
            object_key,
            version,
            layout,
        }));
    }
    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn one_segment_idf() -> Vec<u8> {
        0u64.to_le_bytes().to_vec()
    }

    fn valid_hash_index() -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&(-1i32).to_le_bytes());
        bytes.extend_from_slice(&8u32.to_le_bytes());
        bytes.extend_from_slice(&64u32.to_le_bytes());
        bytes.extend_from_slice(&6u32.to_le_bytes());
        bytes.extend_from_slice(&16i64.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&15u64.to_le_bytes());
        bytes.push(0);
        for bin in 0..16 {
            let mut raw = [0u8; 64];
            if bin == 0 {
                raw[8..12].copy_from_slice(&1u32.to_le_bytes());
            }
            bytes.extend_from_slice(&raw);
        }
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes
    }

    fn hierarchy_files<'a>(idf: &'a [u8]) -> Vec<SystemGeneratedFile<'a>> {
        [
            "Model.1.db/Table.0.dim/8.U$Table1$Geography.CHILD_COUNT.0.idf",
            "Model.1.db/Table.0.dim/8.U$Table1$Geography.FIRST_CHILD_POS.0.idf",
            "Model.1.db/Table.0.dim/8.U$Table1$Geography.MULTI_LEVEL_ID.0.idf",
            "Model.1.db/Table.0.dim/8.U$Table1$Geography.PARENT_POS.0.idf",
        ]
        .into_iter()
        .map(|path| parse_system_generated_file(path, idf).unwrap().unwrap())
        .collect()
    }

    #[test]
    fn parses_every_section_2_4_role_as_borrowed_data() {
        let idf = one_segment_idf();
        let cases = [
            (
                "Model.1.db/Table.0.dim/1.H$Table1$Label.POS_TO_ID.0.idf",
                SystemGeneratedKind::PositionToIdentifier,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table1$Label.ID_TO_POS.0.idf",
                SystemGeneratedKind::IdentifierToPosition,
            ),
            (
                "Model.1.db/Table.0.dim/73.R$Table1$c4047114-e5d3-4730-ab46-478baf7ae64f.INDEX.0.idf",
                SystemGeneratedKind::RelationshipIndex,
            ),
            (
                "Model.1.db/Table.0.dim/8.U$Table1$Geography.CHILD_COUNT.0.idf",
                SystemGeneratedKind::UserHierarchyChildCount,
            ),
            (
                "Model.1.db/Table.0.dim/8.U$Table1$Geography.FIRST_CHILD_POS.0.idf",
                SystemGeneratedKind::UserHierarchyFirstChildPosition,
            ),
            (
                "Model.1.db/Table.0.dim/8.U$Table1$Geography.MULTI_LEVEL_ID.0.idf",
                SystemGeneratedKind::UserHierarchyMultilevelIdentifier,
            ),
            (
                "Model.1.db/Table.0.dim/8.U$Table1$Geography.PARENT_POS.0.idf",
                SystemGeneratedKind::UserHierarchyParentPosition,
            ),
        ];
        for (path, expected) in cases {
            let file = parse_system_generated_file(path, &idf).unwrap().unwrap();
            assert_eq!(file.kind, expected);
            assert_eq!(
                file.expected_compression(),
                SystemGeneratedCompression::XmReNoSplit
            );
            assert!(std::ptr::eq(file.bytes.as_ptr(), idf.as_ptr()));
            assert_eq!(write_system_generated_file(&file).unwrap(), idf);
        }
        let hidx = valid_hash_index();
        let file = parse_system_generated_file(
            "Model.1.db/Table.0.dim/73.R$Table1$c4047114-e5d3-4730-ab46-478baf7ae64f.INDEX.0.idf",
            &hidx,
        )
        .unwrap()
        .unwrap();
        assert!(matches!(file.data, SystemGeneratedData::HashIndex(_)));
        assert_eq!(
            file.expected_compression(),
            SystemGeneratedCompression::Uncompressed
        );
    }

    #[test]
    fn validates_complete_hierarchy_sets_and_relationship_alternatives() {
        let idf = one_segment_idf();
        let complete = hierarchy_files(&idf);
        validate_system_generated_files(&complete).unwrap();

        for missing in 0..4 {
            let mut incomplete = complete.clone();
            incomplete.remove(missing);
            assert!(validate_system_generated_files(&incomplete).is_err());
        }

        let relationship_idf =
            parse_system_generated_file("Model.1.db/Table.0.dim/73.R$Table1$Rel.INDEX.0.idf", &idf)
                .unwrap()
                .unwrap();
        let hidx = valid_hash_index();
        let relationship_hidx = parse_system_generated_file(
            "Model.1.db/Table.0.dim/73.R$Table1$Rel.INDEX.0.idf",
            &hidx,
        )
        .unwrap()
        .unwrap();
        assert!(validate_system_generated_files(&[relationship_idf, relationship_hidx,]).is_err());
    }

    #[test]
    fn rejects_bad_names_versions_layouts_and_payloads() {
        let idf = one_segment_idf();
        for path in [
            "1.H$Table1$Label.POS_TO_ID.0.idf",
            "1.X$Table$Column.POS_TO_ID.0.idf",
            "1.H$Table$Column.POS_TO_ID.x.idf",
            "1.H$Table$Column.POS_TO_ID.0.hidx",
            "x.H$Table$Column.ID_TO_POS.0.idf",
            "8.U$Table$Hierarchy.CHILD_COUNT.18446744073709551616.idf",
        ] {
            assert!(parse_system_generated_file(path, &idf).is_err());
        }
        assert!(
            parse_system_generated_file(
                "Model.1.db/Table.0.dim/1.H$Table$Column.POS_TO_ID.0.idf",
                &[1, 2, 3],
            )
            .is_err()
        );
    }

    #[test]
    fn rejects_duplicate_or_version_skewed_hierarchy_roles() {
        let idf = one_segment_idf();
        let mut files = hierarchy_files(&idf);
        files.push(files[0].clone());
        assert!(validate_system_generated_files(&files).is_err());

        let mut files = hierarchy_files(&idf);
        files[3].version = 1;
        assert!(validate_system_generated_files(&files).is_err());
    }
}

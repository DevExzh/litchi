//! Public section 2.5 metadata model and borrowed snapshot records.

use std::error::Error;
use std::fmt;

use super::super::native::{NativeParseOptions, StringHashMode, StringHashOverride};

const NO_SPLIT_WIDTHS: &[u8] = &[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 21, 32];

/// A structural or cross-file section 2.5 validation error.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataError(pub(crate) String);

impl MetadataError {
    pub(super) fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for MetadataError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for MetadataError {}

/// Result type for section 2.5 metadata inspection.
pub type MetadataResult<T> = Result<T, MetadataError>;

/// A validated class token from `XMObjectClassNameEnum`.
#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub struct MetadataClass(pub(crate) String);

impl MetadataClass {
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// The persisted no-split bit width, if this is a no-split or hybrid class.
    #[must_use]
    pub fn no_split_width(&self) -> Option<u8> {
        no_split_width(&self.0).or_else(|| hybrid_width(&self.0))
    }

    #[must_use]
    pub fn is_hybrid_compression(&self) -> bool {
        self.0.starts_with("XMHybridRLECompressionInfo<class ")
    }
}

/// One scalar property. `value` is XML-unescaped but otherwise uninterpreted.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataProperty {
    pub name: String,
    pub value: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataMember {
    pub name: String,
    pub object: Box<MetadataObject>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataCollection {
    pub name: String,
    pub objects: Vec<MetadataObject>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataDataObject {
    pub object: Box<MetadataObject>,
}

/// The complete generic section 2.5.1 object shape.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataObject {
    pub class: MetadataClass,
    pub name: Option<String>,
    pub provider_version: Option<i32>,
    pub properties: Vec<MetadataProperty>,
    pub members: Vec<MetadataMember>,
    pub collections: Vec<MetadataCollection>,
    pub data_objects: Vec<MetadataDataObject>,
}

impl MetadataObject {
    #[must_use]
    pub fn property(&self, name: &str) -> Option<&str> {
        self.properties
            .iter()
            .find(|item| item.name == name)
            .map(|item| item.value.as_str())
    }

    #[must_use]
    pub fn member(&self, name: &str) -> Option<&MetadataObject> {
        self.members
            .iter()
            .find(|item| item.name == name)
            .map(|item| item.object.as_ref())
    }

    #[must_use]
    pub fn collection(&self, name: &str) -> Option<&[MetadataObject]> {
        self.collections
            .iter()
            .find(|item| item.name == name)
            .map(|item| item.objects.as_slice())
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MetadataFileKind {
    Table,
    ColumnHierarchy,
    UserHierarchy,
    TableRelationship,
}

/// One borrowed `.tbl.xml` file and its validated object graph.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataFile<'a> {
    pub storage_path: &'a str,
    pub bytes: &'a [u8],
    pub kind: MetadataFileKind,
    pub table: MetadataObject,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MetadataModel<'a> {
    pub files: Vec<MetadataFile<'a>>,
    pub columns: Vec<ColumnPolicy>,
    pub relationships: Vec<RelationshipPolicy>,
    pub hierarchies: Vec<HierarchyPolicy>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DictionaryPolicy {
    pub storage_name: String,
    pub class: MetadataClass,
    pub dictionary_flags: Option<u16>,
    pub operating_on_32: Option<bool>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ColumnPolicy {
    pub name: String,
    pub data_file: String,
    pub segment_count: usize,
    pub row_count: u64,
    pub compression_type: u8,
    pub settings: u16,
    pub dictionary: Option<DictionaryPolicy>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RelationshipIndexKind {
    Sparse,
    Dense,
    OneTwoThree,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct RelationshipPolicy {
    pub name: Option<String>,
    pub primary_table: String,
    pub primary_column: String,
    pub foreign_column: String,
    pub index_kind: RelationshipIndexKind,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HierarchyPolicy {
    pub table_store: String,
    pub processed: bool,
    pub position_to_id: bool,
    pub id_to_position: bool,
    pub id_to_position_hash: bool,
    /// Section 2.5 user-hierarchy ID components, matched to section 2.6 levels.
    pub level_ids: Vec<String>,
    /// Cumulative distinct-value offsets paired with `level_ids`.
    pub level_offsets: Vec<u64>,
}

impl MetadataModel<'_> {
    /// Section 2.5 `DictionaryFlags` bindings needed to parse string stores.
    #[must_use]
    pub fn native_parse_options(&self) -> NativeParseOptions {
        let string_hash_overrides = self
            .columns
            .iter()
            .filter_map(|column| {
                let dictionary = column.dictionary.as_ref()?;
                let flags = dictionary.dictionary_flags?;
                Some(StringHashOverride {
                    storage_path: dictionary.storage_name.clone(),
                    mode: if flags & 0x001 != 0 {
                        StringHashMode::Present
                    } else {
                        StringHashMode::Absent
                    },
                })
            })
            .collect();
        NativeParseOptions {
            string_hash_overrides,
        }
    }
}

pub(super) fn no_split_width(class: &str) -> Option<u8> {
    let value = class
        .strip_prefix("XMRENoSplitCompressionInfo<")?
        .strip_suffix('>')?
        .parse()
        .ok()?;
    NO_SPLIT_WIDTHS.contains(&value).then_some(value)
}

pub(super) fn hybrid_width(class: &str) -> Option<u8> {
    let value = class
        .strip_prefix("XMHybridRLECompressionInfo<class XMRENoSplitCompressionInfo<")?
        .strip_suffix(">>")?
        .parse()
        .ok()?;
    NO_SPLIT_WIDTHS.contains(&value).then_some(value)
}

//! Inert RTF 1.9.1 mail-merge metadata.

#![allow(
    clippy::shadow_reuse,
    reason = "builder-style helpers deliberately rebind a working value as it is refined"
)]
use crate::{RtfError, RtfResult};
use std::borrow::Cow;

/// Maximum decoded size of one mail-merge text destination.
pub const MAX_MAIL_MERGE_STRING_BYTES: usize = 32 * 1024;
/// Maximum aggregate decoded size of all mail-merge text destinations.
pub const MAX_MAIL_MERGE_TOTAL_BYTES: usize = 256 * 1024;
/// Maximum number of data-source field mappings.
pub const MAX_MAIL_MERGE_FIELD_MAPPINGS: usize = 1_024;
/// Maximum number of recipient-data destinations.
pub const MAX_MAIL_MERGE_RECIPIENT_DATA: usize = 1_024;
/// Maximum supported structural depth within the mail-merge destination.
pub const MAX_MAIL_MERGE_NESTING_DEPTH: usize = 4;

/// Forward-compatible typed value of the RTF `mmodsosrc` control.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MailMergeDataSourceType(i32);

impl MailMergeDataSourceType {
    #[must_use]
    pub const fn from_rtf(value: i32) -> Self {
        Self(value)
    }

    #[must_use]
    pub const fn rtf_value(self) -> i32 {
        self.0
    }
}

/// Zero-based data-source column index.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct MailMergeColumnIndex(u32);

impl MailMergeColumnIndex {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_rtf(value: i32) -> RtfResult<Self> {
        let value = u32::try_from(value).map_err(|_err| {
            RtfError::MalformedDocument(
                "RTF mail-merge column index cannot be negative".to_string(),
            )
        })?;
        Ok(Self(value))
    }

    #[must_use]
    pub const fn new(value: u32) -> Self {
        Self(value)
    }

    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn rtf_value(self) -> RtfResult<i32> {
        i32::try_from(self.0).map_err(|_err| {
            RtfError::MalformedDocument(
                "RTF mail-merge column index exceeds the signed control-word range".to_string(),
            )
        })
    }
}

/// One field-name mapping in an `mmodsofldmpdata` destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeFieldMapping<'a> {
    pub column: MailMergeColumnIndex,
    pub name: Cow<'a, str>,
    pub mapped_name: Option<Cow<'a, str>>,
}

impl<'a> MailMergeFieldMapping<'a> {
    pub fn new(column: MailMergeColumnIndex, name: impl Into<Cow<'a, str>>) -> Self {
        Self {
            column,
            name: name.into(),
            mapped_name: None,
        }
    }

    #[must_use]
    pub fn with_mapped_name(mut self, name: impl Into<Cow<'a, str>>) -> Self {
        self.mapped_name = Some(name.into());
        self
    }

    #[must_use]
    pub fn into_owned(self) -> MailMergeFieldMapping<'static> {
        MailMergeFieldMapping {
            column: self.column,
            name: Cow::Owned(self.name.into_owned()),
            mapped_name: self.mapped_name.map(|value| Cow::Owned(value.into_owned())),
        }
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        self.column.rtf_value()?;
        validate_required_text("field name", self.name.as_ref())?;
        validate_optional_text("mapped field name", self.mapped_name.as_deref())
    }
}

/// Office data-source-object metadata nested in a mail-merge group.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MailMergeDataSourceObject<'a> {
    pub active_record: Option<u32>,
    pub column_delimiter: Option<i32>,
    pub column_count: Option<u32>,
    pub dynamic_address: Option<bool>,
    pub first_row_header: Option<bool>,
    pub hash: Option<i32>,
    pub id: Option<i32>,
    pub source_type: Option<MailMergeDataSourceType>,
    pub filter: Option<Cow<'a, str>>,
    pub name: Option<Cow<'a, str>>,
    pub sort: Option<Cow<'a, str>>,
    pub table: Option<Cow<'a, str>>,
    pub udl: Option<Cow<'a, str>>,
    pub udl_data: Option<Cow<'a, str>>,
    pub unique_tag: Option<Cow<'a, str>>,
    pub field_mappings: Vec<MailMergeFieldMapping<'a>>,
    pub recipient_data: Vec<Cow<'a, str>>,
}

impl MailMergeDataSourceObject<'_> {
    pub fn into_owned(self) -> MailMergeDataSourceObject<'static> {
        MailMergeDataSourceObject {
            active_record: self.active_record,
            column_delimiter: self.column_delimiter,
            column_count: self.column_count,
            dynamic_address: self.dynamic_address,
            first_row_header: self.first_row_header,
            hash: self.hash,
            id: self.id,
            source_type: self.source_type,
            filter: owned(self.filter),
            name: owned(self.name),
            sort: owned(self.sort),
            table: owned(self.table),
            udl: owned(self.udl),
            udl_data: owned(self.udl_data),
            unique_tag: owned(self.unique_tag),
            field_mappings: self
                .field_mappings
                .into_iter()
                .map(MailMergeFieldMapping::into_owned)
                .collect(),
            recipient_data: self
                .recipient_data
                .into_iter()
                .map(|value| Cow::Owned(value.into_owned()))
                .collect(),
        }
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        for (name, value) in [
            ("active record", self.active_record),
            ("column count", self.column_count),
        ] {
            if let Some(value) = value {
                i32::try_from(value).map_err(|_err| {
                    RtfError::MalformedDocument(format!(
                        "RTF mail-merge {name} exceeds the signed control-word range"
                    ))
                })?;
            }
        }
        if self.field_mappings.len() > MAX_MAIL_MERGE_FIELD_MAPPINGS {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge field-mapping count exceeds the safety limit".to_string(),
            ));
        }
        if self.recipient_data.len() > MAX_MAIL_MERGE_RECIPIENT_DATA {
            return Err(RtfError::MalformedDocument(
                "RTF mail-merge recipient-data count exceeds the safety limit".to_string(),
            ));
        }
        for mapping in &self.field_mappings {
            mapping.validate()?;
        }
        for (name, value) in [
            ("filter", self.filter.as_deref()),
            ("data-source name", self.name.as_deref()),
            ("sort", self.sort.as_deref()),
            ("table", self.table.as_deref()),
            ("UDL", self.udl.as_deref()),
            ("UDL data", self.udl_data.as_deref()),
            ("unique tag", self.unique_tag.as_deref()),
        ] {
            validate_optional_text(name, value)?;
        }
        for value in &self.recipient_data {
            validate_text("recipient data", value.as_ref())?;
        }
        Ok(())
    }
}

/// Complete, inert RTF mail-merge information group.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MailMerge<'a> {
    pub connect_string: Option<Cow<'a, str>>,
    pub connect_string_data: Option<Cow<'a, str>>,
    pub query: Option<Cow<'a, str>>,
    pub data_source: Option<Cow<'a, str>>,
    pub header_source: Option<Cow<'a, str>>,
    pub link_to_query: bool,
    pub data_source_object: Option<MailMergeDataSourceObject<'a>>,
}

impl MailMerge<'_> {
    pub fn into_owned(self) -> MailMerge<'static> {
        MailMerge {
            connect_string: owned(self.connect_string),
            connect_string_data: owned(self.connect_string_data),
            query: owned(self.query),
            data_source: owned(self.data_source),
            header_source: owned(self.header_source),
            link_to_query: self.link_to_query,
            data_source_object: self
                .data_source_object
                .map(MailMergeDataSourceObject::into_owned),
        }
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        let mut total = 0usize;
        for (name, value) in [
            ("connection string", self.connect_string.as_deref()),
            (
                "connection-string data",
                self.connect_string_data.as_deref(),
            ),
            ("query", self.query.as_deref()),
            ("data source", self.data_source.as_deref()),
            ("header source", self.header_source.as_deref()),
        ] {
            if let Some(value) = value {
                validate_text(name, value)?;
                add_size(&mut total, value.len())?;
            }
        }
        if let Some(object) = &self.data_source_object {
            object.validate()?;
            for value in [
                object.filter.as_deref(),
                object.name.as_deref(),
                object.sort.as_deref(),
                object.table.as_deref(),
                object.udl.as_deref(),
                object.udl_data.as_deref(),
                object.unique_tag.as_deref(),
            ]
            .into_iter()
            .flatten()
            {
                add_size(&mut total, value.len())?;
            }
            for mapping in &object.field_mappings {
                add_size(&mut total, mapping.name.len())?;
                if let Some(value) = &mapping.mapped_name {
                    add_size(&mut total, value.len())?;
                }
            }
            for value in &object.recipient_data {
                add_size(&mut total, value.len())?;
            }
        }
        Ok(())
    }
}

fn owned(value: Option<Cow<'_, str>>) -> Option<Cow<'static, str>> {
    value.map(|value| Cow::Owned(value.into_owned()))
}

fn validate_required_text(name: &str, value: &str) -> RtfResult<()> {
    if value.is_empty() {
        return Err(RtfError::MalformedDocument(format!(
            "RTF mail-merge {name} cannot be empty"
        )));
    }
    validate_text(name, value)
}

fn validate_optional_text(name: &str, value: Option<&str>) -> RtfResult<()> {
    if let Some(value) = value {
        validate_text(name, value)?;
    }
    Ok(())
}

fn validate_text(name: &str, value: &str) -> RtfResult<()> {
    if value.len() > MAX_MAIL_MERGE_STRING_BYTES {
        return Err(RtfError::MalformedDocument(format!(
            "RTF mail-merge {name} exceeds the per-string safety limit"
        )));
    }
    Ok(())
}

fn add_size(total: &mut usize, size: usize) -> RtfResult<()> {
    *total = total.checked_add(size).ok_or_else(|| {
        RtfError::MalformedDocument("RTF mail-merge aggregate size overflow".to_string())
    })?;
    if *total > MAX_MAIL_MERGE_TOTAL_BYTES {
        return Err(RtfError::MalformedDocument(
            "RTF mail-merge aggregate text exceeds the safety limit".to_string(),
        ));
    }
    Ok(())
}

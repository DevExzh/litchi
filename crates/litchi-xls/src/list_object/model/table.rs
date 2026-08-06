//! Table aggregate semantics and ergonomic mutation facade.

use super::super::{LIST12_RECORD_TYPE, invalid};
use super::{
    ExternalTableMetadata, ListObjectColumn, ListObjectFeatureVersion, ListObjectId,
    ListObjectRange, ListObjectSourceMetadata, ListObjectStyleOptions, OpaqueListObjectFeature,
    OpaqueListObjectFutureRecord, WebTableMetadata, XmlTableMetadata,
};
use crate::Result;
use crate::autofilter12::TableAutoFilter12;
use crate::list_object::TableFlags;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListObject {
    pub(in crate::list_object) id: ListObjectId,
    pub(in crate::list_object) name: String,
    pub(in crate::list_object) range: ListObjectRange,
    pub(in crate::list_object) columns: Vec<ListObjectColumn>,
    pub(in crate::list_object) style: Option<ListObjectStyleOptions>,
    pub(in crate::list_object) has_header: bool,
    pub(in crate::list_object) has_totals: bool,
    pub(in crate::list_object) autofilter: bool,
    pub(in crate::list_object) table_flags: TableFlags,
    pub(in crate::list_object) comment: String,
    pub(in crate::list_object) feature_version: ListObjectFeatureVersion,
    pub(in crate::list_object) opaque_feature: Option<OpaqueListObjectFeature>,
    pub(in crate::list_object) opaque_future_records: Vec<OpaqueListObjectFutureRecord>,
    pub(in crate::list_object) autofilter12_criteria: Option<TableAutoFilter12>,
    pub(in crate::list_object) external_metadata: Option<ExternalTableMetadata>,
    pub(in crate::list_object) source_metadata: Option<ListObjectSourceMetadata>,
}
impl ListObject {
    pub fn try_new(
        id: ListObjectId,
        name: impl Into<String>,
        range: ListObjectRange,
        columns: Vec<ListObjectColumn>,
        style: ListObjectStyleOptions,
    ) -> Result<Self> {
        let feature_version = if columns
            .iter()
            .any(|c| c.total_formula.is_some() || c.total_string.is_some())
        {
            ListObjectFeatureVersion::Feature12
        } else {
            ListObjectFeatureVersion::Feature11
        };
        let value = Self {
            id,
            name: name.into(),
            range,
            columns,
            style: Some(style),
            has_header: true,
            has_totals: false,
            autofilter: true,
            table_flags: TableFlags::default_table(),
            comment: String::new(),
            feature_version,
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn with_header_row(mut self, v: bool) -> Result<Self> {
        self.has_header = v;
        if !v {
            self.autofilter = false;
            self.table_flags = self
                .table_flags
                .with_auto_filter(false)
                .with_persist_auto_filter(false)
                .with_apply_auto_filter(false);
            self.feature_version = ListObjectFeatureVersion::Feature12;
        } else if self.opaque_feature.is_none() {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_totals_row(mut self, v: bool) -> Result<Self> {
        self.has_totals = v;
        if v {
            self.table_flags = self.table_flags.with_shown_total_row(true);
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter(mut self, v: bool) -> Result<Self> {
        self.autofilter = v;
        self.table_flags = self.table_flags.with_auto_filter(v);
        if !v {
            self.table_flags = self
                .table_flags
                .with_persist_auto_filter(false)
                .with_apply_auto_filter(false);
        }
        self.validate()?;
        Ok(self)
    }
    /// Replace the fixed `TableFeatureType` metadata while retaining the
    /// table's row and column semantics.
    pub fn with_table_flags(mut self, flags: TableFlags) -> Result<Self> {
        self.table_flags = flags;
        self.autofilter = flags.auto_filter();
        self.validate()?;
        Ok(self)
    }
    pub fn with_autofilter12_criteria(mut self, value: TableAutoFilter12) -> Result<Self> {
        self.autofilter12_criteria = Some(value);
        self.validate()?;
        Ok(self)
    }
    pub fn with_comment(mut self, v: impl Into<String>) -> Result<Self> {
        self.comment = v.into();
        if self.comment.encode_utf16().count() > 255 {
            return Err(invalid(
                LIST12_RECORD_TYPE,
                "table comment exceeds 255 characters",
            ));
        }
        Ok(self)
    }
    pub fn with_external_data(mut self, metadata: ExternalTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.external_metadata = Some(metadata);
        self.feature_version = ListObjectFeatureVersion::Feature12;
        self.opaque_feature = None;
        self.validate()?;
        Ok(self)
    }
    pub fn with_web_source(mut self, metadata: WebTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.source_metadata = Some(ListObjectSourceMetadata::Web(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != ListObjectFeatureVersion::Feature12 {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub fn with_xml_source(mut self, metadata: XmlTableMetadata) -> Result<Self> {
        metadata.validate()?;
        self.source_metadata = Some(ListObjectSourceMetadata::Xml(metadata));
        self.external_metadata = None;
        self.opaque_feature = None;
        if self.feature_version != ListObjectFeatureVersion::Feature12 {
            self.feature_version = ListObjectFeatureVersion::Feature11;
        }
        self.validate()?;
        Ok(self)
    }
    pub const fn id(&self) -> ListObjectId {
        self.id
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub const fn range(&self) -> ListObjectRange {
        self.range
    }
    pub fn columns(&self) -> &[ListObjectColumn] {
        &self.columns
    }
    pub fn style(&self) -> Option<&ListObjectStyleOptions> {
        self.style.as_ref()
    }
    pub const fn has_header_row(&self) -> bool {
        self.has_header
    }
    pub const fn has_totals_row(&self) -> bool {
        self.has_totals
    }
    pub const fn shows_autofilter(&self) -> bool {
        self.autofilter
    }
    pub const fn table_flags(&self) -> TableFlags {
        self.table_flags
    }
    pub fn comment(&self) -> &str {
        &self.comment
    }
    pub const fn feature_version(&self) -> ListObjectFeatureVersion {
        self.feature_version
    }
    pub fn opaque_feature(&self) -> Option<&OpaqueListObjectFeature> {
        self.opaque_feature.as_ref()
    }
    pub fn opaque_future_records(&self) -> &[OpaqueListObjectFutureRecord] {
        &self.opaque_future_records
    }
    pub fn autofilter12_criteria(&self) -> Option<&TableAutoFilter12> {
        self.autofilter12_criteria.as_ref()
    }
    pub fn external_metadata(&self) -> Option<&ExternalTableMetadata> {
        self.external_metadata.as_ref()
    }
    pub fn source_metadata(&self) -> Option<&ListObjectSourceMetadata> {
        self.source_metadata.as_ref()
    }
}

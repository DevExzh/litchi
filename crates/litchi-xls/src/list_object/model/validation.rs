//! Cross-component invariants for list-object semantic values.

use super::super::{
    AUTO_FILTER12_RECORD_TYPE, FEATURE11_RECORD_TYPE, FEATURE12_RECORD_TYPE, LIST12_RECORD_TYPE,
    MAX_FEATURE_BYTES, invalid,
};
use super::{ListObject, ListObjectFeatureVersion, ListObjectSourceMetadata};
use crate::Result;
use std::collections::HashSet;

pub(in crate::list_object) fn validate_name(value: &str, field: &str) -> Result<()> {
    if !(1..=255).contains(&value.encode_utf16().count())
        || value
            .chars()
            .any(|c| c <= '\u{1f}' || matches!(c, '\u{fffe}' | '\u{ffff}'))
    {
        return Err(invalid(FEATURE11_RECORD_TYPE, format!("invalid {field}")));
    }
    Ok(())
}
pub(in crate::list_object) fn validate_table_name(value: &str) -> Result<()> {
    validate_name(value, "table name")?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first.is_alphabetic() || matches!(first, '_' | '\\'))
        || chars.any(|c| !(c.is_alphanumeric() || matches!(c, '_' | '.' | '\\')))
    {
        return Err(invalid(
            FEATURE11_RECORD_TYPE,
            "table name must use Excel identifier syntax",
        ));
    }
    Ok(())
}
pub(in crate::list_object) fn validate_column_name(value: &str) -> Result<()> {
    if !(1..=255).contains(&value.encode_utf16().count())
        || value.chars().any(|c| {
            (c < '\u{20}' && !matches!(c, '\t' | '\n' | '\r'))
                || matches!(c, '\u{fffe}' | '\u{ffff}')
        })
    {
        return Err(invalid(FEATURE11_RECORD_TYPE, "invalid column name"));
    }
    Ok(())
}

impl ListObject {
    pub(crate) fn validate(&self) -> Result<()> {
        validate_table_name(&self.name)?;
        if self.opaque_feature.is_none()
            && (self.columns.is_empty()
                || self.columns.len() > 256
                || self.columns.len() != self.range.column_count())
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "column count must match the table range",
            ));
        }
        if self.opaque_feature.is_some()
            && self.feature_version != ListObjectFeatureVersion::Feature12
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "opaque table feature must be Feature12",
            ));
        }
        if let Some(metadata) = &self.external_metadata {
            metadata.validate()?;
            if self.feature_version != ListObjectFeatureVersion::Feature12
                || metadata.fields.len() != self.columns.len()
                || metadata
                    .fields
                    .iter()
                    .zip(&self.columns)
                    .any(|(field, column)| field.column_id != column.id)
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "external metadata must be Feature12 and owned one-for-one by table columns",
                ));
            }
            for (field, column) in metadata.fields.iter().zip(&self.columns) {
                if field.total_array_formula != column.total_formula.is_some()
                    && field.total_array_formula
                {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "array formula metadata requires a total formula",
                    ));
                }
                if !field.total_array_formula && !field.formula_extra.is_empty() {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "scalar total formula cannot carry RgbExtra",
                    ));
                }
                if self.has_header
                    && (!field.header_cache.formatting_bytes().is_empty()
                        || field.header_cache.style_name().is_some())
                {
                    return Err(invalid(
                        FEATURE12_RECORD_TYPE,
                        "cached disk header requires a headerless external table",
                    ));
                }
            }
        }
        if let Some(source) = &self.source_metadata {
            if !matches!(
                self.feature_version,
                ListObjectFeatureVersion::Feature11 | ListObjectFeatureVersion::Feature12
            ) || self.external_metadata.is_some()
                || self.opaque_feature.is_some()
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Web/XML source metadata requires a typed Feature11 or Feature12",
                ));
            }
            let has_feature12_field = self
                .columns
                .iter()
                .any(|column| column.total_formula.is_some() || column.total_string.is_some());
            if self.feature_version == ListObjectFeatureVersion::Feature11 && has_feature12_field {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "Feature11 source fields cannot load total formulas or strings",
                ));
            }
            if self.feature_version == ListObjectFeatureVersion::Feature12
                && self.has_header
                && !has_feature12_field
            {
                return Err(invalid(
                    FEATURE12_RECORD_TYPE,
                    "Feature12 Web/XML source requires a Feature12-only property",
                ));
            }
            match source {
                ListObjectSourceMetadata::Web(metadata) => {
                    metadata.validate()?;
                    if !self.has_header
                        || metadata.fields.len() != self.columns.len()
                        || metadata
                            .fields
                            .iter()
                            .zip(&self.columns)
                            .any(|(field, column)| field.column_id != column.id)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "Web source fields must be owned one-for-one by a headered table",
                        ));
                    }
                    for (field, column) in metadata.fields.iter().zip(&self.columns) {
                        if column.total_formula.is_some()
                            && field.total_formula_extra.len() > MAX_FEATURE_BYTES
                        {
                            return Err(invalid(
                                FEATURE11_RECORD_TYPE,
                                "Web total formula extra data exceeds resource bound",
                            ));
                        }
                    }
                },
                ListObjectSourceMetadata::Xml(metadata) => {
                    metadata.validate()?;
                    if metadata
                        .entry_id
                        .as_deref()
                        .is_some_and(|entry| entry != self.id.value().to_string())
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "XML entry id must equal the decimal table id",
                        ));
                    }
                    if metadata.fields.len() != self.columns.len()
                        || metadata
                            .fields
                            .iter()
                            .zip(&self.columns)
                            .any(|(field, column)| field.column_id != column.id)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "XML source fields must be owned one-for-one by table columns",
                        ));
                    }
                    if metadata.single_cell
                        && (self.has_header
                            || self.has_totals
                            || self.columns.len() != 1
                            || self.range.first_row != self.range.last_row
                            || self.range.first_column != self.range.last_column)
                    {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "single-cell XML source requires one unheadered cell",
                        ));
                    }
                    if !metadata.single_cell && !self.has_header {
                        return Err(invalid(
                            FEATURE11_RECORD_TYPE,
                            "multi-cell XML source requires a header row",
                        ));
                    }
                },
            }
        }
        if self.autofilter && !self.has_header {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "AutoFilter requires a header row",
            ));
        }
        if self.table_flags.persists_auto_filter() && !self.table_flags.auto_filter()
            || self.table_flags.applies_auto_filter() && !self.table_flags.auto_filter()
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "TableFeatureType AutoFilter flags are inconsistent",
            ));
        }
        if self.table_flags.auto_filter() != self.autofilter {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "table AutoFilter metadata does not match the table semantic flag",
            ));
        }
        if let Some(filter) = &self.autofilter12_criteria {
            filter.validate()?;
            if !self.autofilter
                || !self.has_header
                || usize::from(filter.column_index()) >= self.range.column_count()
            {
                return Err(invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "typed AutoFilter12 criteria require an in-range column on a headered table AutoFilter",
                ));
            }
            if self
                .opaque_future_records
                .iter()
                .any(|future| future.record_type == AUTO_FILTER12_RECORD_TYPE)
            {
                return Err(invalid(
                    AUTO_FILTER12_RECORD_TYPE,
                    "typed and opaque AutoFilter12 records cannot coexist",
                ));
            }
        }
        if self.has_totals && self.range.first_row == self.range.last_row {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "totals row requires a range below the header",
            ));
        }
        let mut ids = HashSet::new();
        let mut names = HashSet::new();
        for column in &self.columns {
            column.validate_totals()?;
            if !ids.insert(column.id) || !names.insert(column.name.to_lowercase()) {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "duplicate column id or name",
                ));
            }
        }
        if self.style.is_none() {
            return Err(invalid(LIST12_RECORD_TYPE, "missing table style"));
        }
        Ok(())
    }
}

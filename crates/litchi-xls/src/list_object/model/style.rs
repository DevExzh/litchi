//! Table-column totals and visual style semantics.

use super::super::{FEATURE12_RECORD_TYPE, invalid};
use super::{ListColumnId, ListTotalAggregation, validate_column_name, validate_name};
use crate::Result;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListObjectColumn {
    pub(in crate::list_object) id: ListColumnId,
    pub(in crate::list_object) name: String,
    pub(in crate::list_object) aggregation: ListTotalAggregation,
    pub(in crate::list_object) total_formula: Option<Vec<u8>>,
    pub(in crate::list_object) total_string: Option<String>,
}
impl ListObjectColumn {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(id: ListColumnId, name: impl Into<String>) -> Result<Self> {
        let value = Self {
            id,
            name: name.into(),
            aggregation: ListTotalAggregation::None,
            total_formula: None,
            total_string: None,
        };
        validate_column_name(&value.name)?;
        Ok(value)
    }
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn with_total_aggregation(mut self, value: ListTotalAggregation) -> Result<Self> {
        self.aggregation = value;
        self.validate_totals()?;
        Ok(self)
    }
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn with_total_formula_tokens(mut self, tokens: Vec<u8>) -> Result<Self> {
        if tokens.is_empty() || tokens.len() > u16::MAX as usize {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total formula token length must be 1..=65535",
            ));
        }
        self.aggregation = ListTotalAggregation::Custom;
        self.total_formula = Some(tokens);
        self.validate_totals()?;
        Ok(self)
    }
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn with_total_string(mut self, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.encode_utf16().count() > 32767 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string exceeds 32767 UTF-16 units",
            ));
        }
        self.aggregation = ListTotalAggregation::None;
        self.total_string = Some(value);
        self.validate_totals()?;
        Ok(self)
    }
    pub(in crate::list_object) fn validate_totals(&self) -> Result<()> {
        if self.total_formula.is_some() != (self.aggregation == ListTotalAggregation::Custom) {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "custom aggregation and total formula must occur together",
            ));
        }
        if self.total_string.is_some() && self.aggregation != ListTotalAggregation::None {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "total string requires no aggregation",
            ));
        }
        Ok(())
    }
    #[must_use]
    pub const fn id(&self) -> ListColumnId {
        self.id
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub const fn total_aggregation(&self) -> ListTotalAggregation {
        self.aggregation
    }
    #[must_use]
    pub fn total_formula_tokens(&self) -> Option<&[u8]> {
        self.total_formula.as_deref()
    }
    #[must_use]
    pub fn total_string(&self) -> Option<&str> {
        self.total_string.as_deref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListObjectStyleOptions {
    pub(in crate::list_object) name: String,
    pub(in crate::list_object) first: bool,
    pub(in crate::list_object) last: bool,
    pub(in crate::list_object) row_stripes: bool,
    pub(in crate::list_object) column_stripes: bool,
    pub(in crate::list_object) default_style: bool,
}
impl ListObjectStyleOptions {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn try_new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_name(&name, "table style name")?;
        Ok(Self {
            name,
            first: false,
            last: false,
            row_stripes: true,
            column_stripes: false,
            default_style: false,
        })
    }
    #[must_use]
    pub fn with_first_column(mut self, v: bool) -> Self {
        self.first = v;
        self
    }
    #[must_use]
    pub fn with_last_column(mut self, v: bool) -> Self {
        self.last = v;
        self
    }
    #[must_use]
    pub fn with_row_stripes(mut self, v: bool) -> Self {
        self.row_stripes = v;
        self
    }
    #[must_use]
    pub fn with_column_stripes(mut self, v: bool) -> Self {
        self.column_stripes = v;
        self
    }
    #[must_use]
    pub fn with_default_style(mut self, v: bool) -> Self {
        self.default_style = v;
        self
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub const fn shows_first_column(&self) -> bool {
        self.first
    }
    #[must_use]
    pub const fn shows_last_column(&self) -> bool {
        self.last
    }
    #[must_use]
    pub const fn shows_row_stripes(&self) -> bool {
        self.row_stripes
    }
    #[must_use]
    pub const fn shows_column_stripes(&self) -> bool {
        self.column_stripes
    }
    #[must_use]
    pub const fn is_default_style(&self) -> bool {
        self.default_style
    }
}

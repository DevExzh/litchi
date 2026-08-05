//! Semantic cell assembly and typed value projection.

use super::{
    Annotation, Cell, CellDetective, CellMatrixSpan, CellMerge, CellRangeSource, CellTextContent,
    CellValue, Link,
};
use litchi_core::Result;

/// Builder for constructing a semantic [`Cell`] during content traversal.
pub(crate) struct CellBuilder {
    pub(crate) value_type: Option<String>,
    pub(crate) value_str: Option<String>,
    pub(crate) currency: Option<String>,
    pub(crate) formula: Option<String>,
    pub(crate) validation_name: Option<String>,
    pub(crate) style_name: Option<String>,
    pub(crate) matrix_span: Option<CellMatrixSpan>,
    pub(crate) protect: Option<bool>,
    pub(crate) protected: Option<bool>,
    pub(crate) repeated: usize,
    pub(crate) merge: CellMerge,
    pub(crate) annotation: Option<Annotation>,
    pub(crate) hyperlinks: Vec<Link>,
    pub(crate) range_source: Option<CellRangeSource>,
    pub(crate) detective: Option<CellDetective>,
}

impl CellBuilder {
    pub(crate) fn from_parts(
        value_type: Option<String>,
        value_str: Option<String>,
        currency: Option<String>,
        formula: Option<String>,
        validation_name: Option<String>,
        style_name: Option<String>,
        matrix_span: Option<CellMatrixSpan>,
        protect: Option<bool>,
        protected: Option<bool>,
        repeated: usize,
        merge: CellMerge,
    ) -> Self {
        Self {
            value_type,
            value_str,
            currency,
            formula,
            validation_name,
            style_name,
            matrix_span,
            protect,
            protected,
            repeated,
            merge,
            annotation: None,
            hyperlinks: Vec::new(),
            range_source: None,
            detective: None,
        }
    }

    pub(crate) fn mark_covered(&mut self) {
        self.merge = CellMerge::Covered;
    }

    pub(crate) fn set_range_source(&mut self, source: CellRangeSource) -> Result<()> {
        if self.range_source.replace(source).is_some() {
            return Err(litchi_core::Error::InvalidFormat(
                "table cell contains multiple table:cell-range-source elements".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn begin_detective(&mut self) -> Result<()> {
        if self.detective.is_some() {
            return Err(litchi_core::Error::InvalidFormat(
                "table cell contains multiple table:detective elements".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn set_detective(&mut self, detective: CellDetective) -> Result<()> {
        if self.detective.replace(detective).is_some() {
            return Err(litchi_core::Error::InvalidFormat(
                "table cell contains multiple table:detective elements".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn set_annotation(&mut self, annotation: Annotation) {
        self.annotation = Some(annotation);
    }

    pub(crate) fn push_hyperlink(&mut self, link: Link) {
        self.hyperlinks.push(link);
    }

    /// Whether this cell carries no user data whatsoever.
    ///
    /// A blank cell is exactly the attribute-free `<table:table-cell/>` filler
    /// producers emit to pad a row out to the full sheet width. Anything that a
    /// user could have authored — a value, formula, style, annotation,
    /// hyperlink, validation, protection flag, merge role, or text — makes the
    /// cell meaningful and therefore not blank.
    fn is_blank(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> bool {
        self.value_type.is_none()
            && self.value_str.is_none()
            && self.currency.is_none()
            && self.formula.is_none()
            && self.validation_name.is_none()
            && self.style_name.is_none()
            && self.matrix_span.is_none()
            && self.protect.is_none()
            && self.protected.is_none()
            && self.annotation.is_none()
            && self.hyperlinks.is_empty()
            && self.range_source.is_none()
            && self.detective.is_none()
            && self.merge == CellMerge::None
            && text_content.is_empty()
            && rich_text.is_none()
    }

    pub fn build(&self, text_content: &str, rich_text: Option<&CellTextContent>) -> Cell {
        let value = self.parse_value(text_content);

        Cell {
            value,
            text: text_content.to_string(),
            // Clone necessary: formula may be reused for repeated cells
            formula: self.formula.clone(),
            annotation: self.annotation.clone(),
            hyperlinks: self.hyperlinks.clone(),
            rich_text: rich_text.cloned(),
            range_source: self.range_source.clone(),
            detective: self.detective.clone(),
            validation_name: self.validation_name.clone(),
            style_name: self.style_name.clone(),
            matrix_span: self.matrix_span,
            merge: self.merge,
            protect: self.protect,
            protected: self.protected,
            row: 0, // Will be set by parent
            col: 0, // Will be set by parent
        }
    }

    fn parse_value(&self, text_content: &str) -> CellValue {
        match self.value_type.as_deref() {
            Some("float") | Some("double") | Some("decimal") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Number(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("currency") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        let currency_code = self.currency.as_deref().unwrap_or("USD").to_string();
                        CellValue::Currency(num, currency_code)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("percentage") => {
                if let Some(ref val_str) = self.value_str {
                    if let Ok(num) = val_str.parse::<f64>() {
                        CellValue::Percentage(num)
                    } else {
                        CellValue::Text(text_content.to_string())
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("boolean") => {
                if let Some(ref val_str) = self.value_str {
                    match val_str.as_str() {
                        "true" => CellValue::Boolean(true),
                        "false" => CellValue::Boolean(false),
                        _ => CellValue::Text(text_content.to_string()),
                    }
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("date") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Date(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            Some("time") => {
                if let Some(ref val_str) = self.value_str {
                    CellValue::Time(val_str.to_string())
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
            _ => {
                if text_content.trim().is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text_content.to_string())
                }
            },
        }
    }
}

//! Inert LibreOffice `calcext` conditional-format metadata.
//!
//! LibreOffice Calc persists conditional formatting in the experimental
//! `calcext` namespace (`urn:org:documentfoundation:names:experimental:calc:
//! xmlns:calcext:1.0`) as a `calcext:conditional-formats` container attached to
//! `table:table`. Each `calcext:conditional-format` names the cell ranges it
//! covers and carries `calcext:condition` rules that pair an inert condition
//! expression with the cell style applied when it holds.
//!
//! Litchi parses and stores this metadata as typed data only: conditions are
//! never evaluated and style references are never resolved or applied.
//! `calcext:color-scale`, `calcext:data-bar`, `calcext:icon-set`, and
//! `calcext:date-is` rules are not modeled yet and are skipped on read.

use super::structure::validate_cell_range_addresses;
use litchi_core::{Error, Result, xml::escape_xml};

/// Namespace URI of the LibreOffice `calcext` extension.
pub(crate) const CALCEXT_NAMESPACE_URI: &str =
    "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
/// Namespace declaration written when a document contains conditional formats.
pub(crate) const CALCEXT_NAMESPACE_DECLARATION: &str =
    " xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"";

/// Most conditional formats one sheet may declare.
pub(crate) const MAX_CONDITIONAL_FORMATS_PER_SHEET: usize = 16_384;
/// Most inert condition rules one conditional format may carry.
pub(crate) const MAX_CONDITIONS_PER_FORMAT: usize = 1_024;
/// Largest accepted length of any single lexical attribute value.
pub(crate) const MAX_CONDITIONAL_ATTRIBUTE_BYTES: usize = 64 * 1024;

/// One inert `calcext:conditional-format` element of a sheet.
///
/// Rules are retained in document order. Litchi does not evaluate their
/// conditions or compute an effective style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormat {
    /// ODF cell-range addresses the format applies to
    /// (`calcext:target-range-address`, split on unquoted whitespace).
    pub target_range_addresses: Vec<String>,
    /// Inert `calcext:condition` rules in document order.
    pub conditions: Vec<ConditionalFormatCondition>,
}

/// One inert `calcext:condition` rule of a conditional format.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalFormatCondition {
    /// The decoded condition expression (`calcext:value`). It is never
    /// evaluated by litchi.
    pub condition: String,
    /// Name of the table-cell style referenced by the rule
    /// (`calcext:apply-style-name`). It is never resolved by litchi.
    pub apply_style_name: String,
    /// Optional lexical base cell address for relative formula references
    /// (`calcext:base-cell-address`).
    pub base_cell_address: Option<String>,
}

impl ConditionalFormat {
    /// Create a validated inert conditional format.
    pub fn new(
        target_range_addresses: Vec<String>,
        conditions: Vec<ConditionalFormatCondition>,
    ) -> Result<Self> {
        let format = Self {
            target_range_addresses,
            conditions,
        };
        validate_conditional_format(&format)?;
        Ok(format)
    }
}

impl ConditionalFormatCondition {
    /// Create an inert condition rule without a base cell address.
    pub fn new(condition: impl Into<String>, apply_style_name: impl Into<String>) -> Self {
        Self {
            condition: condition.into(),
            apply_style_name: apply_style_name.into(),
            base_cell_address: None,
        }
    }

    /// Set the optional lexical base cell address.
    pub fn with_base_cell_address(mut self, address: impl Into<String>) -> Self {
        self.base_cell_address = Some(address.into());
        self
    }
}

pub(crate) fn validate_conditional_format(format: &ConditionalFormat) -> Result<()> {
    if format.target_range_addresses.is_empty() {
        return Err(Error::InvalidFormat(
            "conditional formats require at least one target range".to_string(),
        ));
    }
    validate_cell_range_addresses(&format.target_range_addresses)?;
    for range in &format.target_range_addresses {
        validate_attribute_length("calcext:target-range-address", range)?;
    }
    if format.conditions.is_empty() {
        return Err(Error::InvalidFormat(
            "conditional formats require at least one calcext:condition".to_string(),
        ));
    }
    if format.conditions.len() > MAX_CONDITIONS_PER_FORMAT {
        return Err(Error::InvalidFormat(format!(
            "conditional format exceeds the {MAX_CONDITIONS_PER_FORMAT} condition safety limit"
        )));
    }
    for condition in &format.conditions {
        validate_condition(condition)?;
    }
    Ok(())
}

pub(crate) fn validate_conditional_formats(formats: &[ConditionalFormat]) -> Result<()> {
    if formats.len() > MAX_CONDITIONAL_FORMATS_PER_SHEET {
        return Err(Error::InvalidFormat(format!(
            "sheet exceeds the {MAX_CONDITIONAL_FORMATS_PER_SHEET} conditional format safety limit"
        )));
    }
    for format in formats {
        validate_conditional_format(format)?;
    }
    Ok(())
}

pub(crate) fn validate_condition(condition: &ConditionalFormatCondition) -> Result<()> {
    if condition.condition.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:condition requires a non-empty calcext:value".to_string(),
        ));
    }
    validate_attribute_length("calcext:value", &condition.condition)?;
    if condition.apply_style_name.is_empty() {
        return Err(Error::InvalidFormat(
            "calcext:condition requires a non-empty calcext:apply-style-name".to_string(),
        ));
    }
    validate_attribute_length("calcext:apply-style-name", &condition.apply_style_name)?;
    if let Some(address) = &condition.base_cell_address {
        if address.trim() != address || address.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "invalid calcext:base-cell-address '{address}'"
            )));
        }
        validate_attribute_length("calcext:base-cell-address", address)?;
    }
    Ok(())
}

fn validate_attribute_length(name: &str, value: &str) -> Result<()> {
    if value.len() > MAX_CONDITIONAL_ATTRIBUTE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the {MAX_CONDITIONAL_ATTRIBUTE_BYTES} byte safety limit"
        )));
    }
    Ok(())
}

/// Write a sheet's `calcext:conditional-formats` container after its rows.
pub(crate) fn write_conditional_formats(
    out: &mut String,
    formats: &[ConditionalFormat],
) -> Result<()> {
    validate_conditional_formats(formats)?;
    if formats.is_empty() {
        return Ok(());
    }
    out.push_str("<calcext:conditional-formats>");
    for format in formats {
        out.push_str("<calcext:conditional-format calcext:target-range-address=\"");
        out.push_str(&escape_xml(&format.target_range_addresses.join(" ")));
        out.push_str("\">");
        for condition in &format.conditions {
            out.push_str("<calcext:condition calcext:apply-style-name=\"");
            out.push_str(&escape_xml(&condition.apply_style_name));
            out.push_str("\" calcext:value=\"");
            out.push_str(&escape_xml(&condition.condition));
            out.push('"');
            if let Some(address) = &condition.base_cell_address {
                out.push_str(" calcext:base-cell-address=\"");
                out.push_str(&escape_xml(address));
                out.push('"');
            }
            out.push_str("/>");
        }
        out.push_str("</calcext:conditional-format>");
    }
    out.push_str("</calcext:conditional-formats>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_format() -> ConditionalFormat {
        ConditionalFormat::new(
            vec!["Sheet1.A1:Sheet1.A5".to_string()],
            vec![
                ConditionalFormatCondition::new("cell-content()>5", "Good")
                    .with_base_cell_address("Sheet1.A1"),
                ConditionalFormatCondition::new("cell-content()<=5", "Bad"),
            ],
        )
        .unwrap()
    }

    #[test]
    fn writes_container_with_escaped_rules() {
        let mut xml = String::new();
        write_conditional_formats(&mut xml, &[sample_format()]).unwrap();
        assert!(xml.starts_with("<calcext:conditional-formats>"));
        assert!(xml.contains(
            r#"<calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A5">"#
        ));
        assert!(xml.contains(
            r#"<calcext:condition calcext:apply-style-name="Good" calcext:value="cell-content()&gt;5" calcext:base-cell-address="Sheet1.A1"/>"#
        ));
        assert!(xml.contains(
            r#"<calcext:condition calcext:apply-style-name="Bad" calcext:value="cell-content()&lt;=5"/>"#
        ));
        assert!(xml.ends_with("</calcext:conditional-formats>"));
    }

    #[test]
    fn writes_nothing_for_an_empty_collection() {
        let mut xml = String::new();
        write_conditional_formats(&mut xml, &[]).unwrap();
        assert!(xml.is_empty());
    }

    #[test]
    fn rejects_missing_ranges_conditions_and_blank_values() {
        assert!(
            ConditionalFormat::new(Vec::new(), vec![ConditionalFormatCondition::new("x", "S")])
                .is_err()
        );
        assert!(
            ConditionalFormat::new(vec![".A1".to_string()], Vec::new()).is_err()
        );
        assert!(
            ConditionalFormat::new(
                vec![".A1".to_string()],
                vec![ConditionalFormatCondition::new("", "S")],
            )
            .is_err()
        );
        assert!(
            ConditionalFormat::new(
                vec![".A1".to_string()],
                vec![ConditionalFormatCondition::new("x", "")],
            )
            .is_err()
        );
        assert!(
            ConditionalFormat::new(
                vec![".A1".to_string()],
                vec![ConditionalFormatCondition::new("x", "S")
                    .with_base_cell_address(" A1 ")],
            )
            .is_err()
        );
    }
}

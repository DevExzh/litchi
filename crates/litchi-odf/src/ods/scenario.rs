//! ODS what-if scenario metadata.

use super::{
    DdeSource, SheetTableSource, dde::write_dde_source, source::write_table_source,
    structure::validate_cell_range_addresses,
};
use litchi_core::{Error, Result, xml::escape_xml};

/// What-if scenario settings attached to an ODF spreadsheet sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetScenario {
    /// ODF cell-range addresses whose values belong to the scenario.
    pub ranges: Vec<String>,
    /// Whether this scenario currently supplies the active cell values.
    pub is_active: bool,
    /// Whether the scenario range border is displayed.
    pub display_border: Option<bool>,
    /// Optional border color in `#RRGGBB` form.
    pub border_color: Option<String>,
    /// Whether changed values are copied back into the scenario.
    pub copy_back: Option<bool>,
    /// Whether cell styles are copied into the scenario.
    pub copy_styles: Option<bool>,
    /// Whether formulas are copied into the scenario.
    pub copy_formulas: Option<bool>,
    /// Optional user-facing scenario comment.
    pub comment: Option<String>,
    /// Whether the scenario is protected against edits.
    pub protected: Option<bool>,
}

impl SheetScenario {
    /// Create a scenario with validated, non-empty cell-range addresses.
    pub fn new(ranges: Vec<String>, is_active: bool) -> Result<Self> {
        let scenario = Self {
            ranges,
            is_active,
            display_border: None,
            border_color: None,
            copy_back: None,
            copy_styles: None,
            copy_formulas: None,
            comment: None,
            protected: None,
        };
        validate_scenario(&scenario)?;
        Ok(scenario)
    }
}

pub(crate) fn validate_scenario(scenario: &SheetScenario) -> Result<()> {
    if scenario.ranges.is_empty() {
        return Err(Error::InvalidFormat(
            "table scenarios require at least one cell range".to_string(),
        ));
    }
    validate_cell_range_addresses(&scenario.ranges)?;
    if let Some(color) = &scenario.border_color
        && (color.len() != 7
            || !color.starts_with('#')
            || !color[1..].bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(Error::InvalidFormat(format!(
            "invalid table scenario border color '{color}'"
        )));
    }
    Ok(())
}

pub(crate) fn write_scenario(out: &mut String, scenario: &SheetScenario) -> Result<()> {
    validate_scenario(scenario)?;
    out.push_str("<table:scenario table:scenario-ranges=\"");
    out.push_str(&escape_xml(&scenario.ranges.join(" ")));
    out.push_str("\" table:is-active=\"");
    out.push_str(bool_str(scenario.is_active));
    out.push('"');
    write_optional_bool(out, "table:display-border", scenario.display_border);
    if let Some(color) = &scenario.border_color {
        out.push_str(" table:border-color=\"");
        out.push_str(color);
        out.push('"');
    }
    write_optional_bool(out, "table:copy-back", scenario.copy_back);
    write_optional_bool(out, "table:copy-styles", scenario.copy_styles);
    write_optional_bool(out, "table:copy-formulas", scenario.copy_formulas);
    if let Some(comment) = &scenario.comment {
        out.push_str(" table:comment=\"");
        out.push_str(&escape_xml(comment));
        out.push('"');
    }
    write_optional_bool(out, "table:protected", scenario.protected);
    out.push_str("/>");
    Ok(())
}

pub(crate) fn write_sheet_preamble(
    out: &mut String,
    title: Option<&str>,
    description: Option<&str>,
    table_source: Option<&SheetTableSource>,
    dde_source: Option<&DdeSource>,
    scenario: Option<&SheetScenario>,
) -> Result<()> {
    if let Some(title) = title {
        out.push_str("<table:title>");
        out.push_str(&escape_xml(title));
        out.push_str("</table:title>");
    }
    if let Some(description) = description {
        out.push_str("<table:desc>");
        out.push_str(&escape_xml(description));
        out.push_str("</table:desc>");
    }
    if let Some(table_source) = table_source {
        write_table_source(out, table_source)?;
    }
    if let Some(dde_source) = dde_source {
        write_dde_source(out, dde_source)?;
    }
    if let Some(scenario) = scenario {
        write_scenario(out, scenario)?;
    }
    Ok(())
}

fn write_optional_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(bool_str(value));
        out.push('"');
    }
}

fn bool_str(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn writes_complete_scenario_and_escapes_comment() {
        let mut scenario =
            SheetScenario::new(vec!["'Q1 Sales'.$A$1:$B$2".to_string()], true).unwrap();
        scenario.display_border = Some(false);
        scenario.border_color = Some("#12AbEF".to_string());
        scenario.copy_back = Some(true);
        scenario.copy_styles = Some(false);
        scenario.copy_formulas = Some(true);
        scenario.comment = Some("Best & worst".to_string());
        scenario.protected = Some(true);
        let mut xml = String::new();
        write_scenario(&mut xml, &scenario).unwrap();
        assert!(xml.contains(r#"table:is-active="true""#));
        assert!(xml.contains(r##"table:border-color="#12AbEF""##));
        assert!(xml.contains(r#"table:copy-styles="false""#));
        assert!(xml.contains(r#"table:comment="Best &amp; worst""#));
        assert!(xml.ends_with("/>"));
    }

    #[test]
    fn rejects_missing_ranges_and_invalid_colors() {
        assert!(SheetScenario::new(Vec::new(), false).is_err());
        let mut scenario = SheetScenario::new(vec![".A1:.B2".to_string()], false).unwrap();
        scenario.border_color = Some("red".to_string());
        assert!(validate_scenario(&scenario).is_err());
    }
}

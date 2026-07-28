//! Immutable XLSX worksheet sheet-format-properties read model.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::is_spreadsheetml_name;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const X14AC_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const DEFAULT_BASE_COLUMN_WIDTH: u32 = 8;
const MAX_BASE_COLUMN_WIDTH: u32 = 255;
const MAX_DEFAULT_COLUMN_WIDTH: f64 = 65_536.0;
const MAX_OUTLINE_LEVEL: u8 = 7;

/// Effective worksheet defaults and outline metadata from `sheetFormatPr`.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetSheetFormatProperties {
    base_column_width: u32,
    default_column_width: Option<f64>,
    default_row_height: f64,
    custom_height: bool,
    zero_height: bool,
    thick_top: bool,
    thick_bottom: bool,
    outline_level_row: u8,
    outline_level_column: u8,
    dy_descent: Option<f64>,
}

impl WorksheetSheetFormatProperties {
    pub fn base_column_width(&self) -> u32 {
        self.base_column_width
    }
    pub fn default_column_width(&self) -> Option<f64> {
        self.default_column_width
    }
    pub fn effective_default_column_width(&self) -> f64 {
        self.default_column_width
            .unwrap_or(self.base_column_width as f64)
    }
    pub fn default_row_height(&self) -> f64 {
        self.default_row_height
    }
    pub fn custom_height(&self) -> bool {
        self.custom_height
    }
    pub fn zero_height(&self) -> bool {
        self.zero_height
    }
    pub fn thick_top(&self) -> bool {
        self.thick_top
    }
    pub fn thick_bottom(&self) -> bool {
        self.thick_bottom
    }
    pub fn outline_level_row(&self) -> u8 {
        self.outline_level_row
    }
    pub fn outline_level_column(&self) -> u8 {
        self.outline_level_column
    }
    /// Excel 2010 typographical descent in pixels at 100% worksheet zoom.
    pub fn dy_descent(&self) -> Option<f64> {
        self.dy_descent
    }
}

#[derive(Default)]
struct Builder {
    base_column_width: Option<u32>,
    default_column_width: Option<f64>,
    default_row_height: Option<f64>,
    custom_height: Option<bool>,
    zero_height: Option<bool>,
    thick_top: Option<bool>,
    thick_bottom: Option<bool>,
    outline_level_row: Option<u8>,
    outline_level_column: Option<u8>,
    dy_descent: Option<f64>,
}

impl Builder {
    fn finish(self) -> Result<WorksheetSheetFormatProperties> {
        let default_row_height = self
            .default_row_height
            .ok_or_else(|| invalid("sheetFormatPr requires defaultRowHeight"))?;
        Ok(WorksheetSheetFormatProperties {
            base_column_width: self.base_column_width.unwrap_or(DEFAULT_BASE_COLUMN_WIDTH),
            default_column_width: self.default_column_width,
            default_row_height,
            custom_height: self.custom_height.unwrap_or(false) || self.dy_descent.is_some(),
            zero_height: self.zero_height.unwrap_or(false),
            thick_top: self.thick_top.unwrap_or(false),
            thick_bottom: self.thick_bottom.unwrap_or(false),
            outline_level_row: self.outline_level_row.unwrap_or(0),
            outline_level_column: self.outline_level_column.unwrap_or(0),
            dy_descent: self.dy_descent,
        })
    }
}

/// Parse the worksheet's direct `sheetFormatPr` child.
pub fn parse_worksheet_sheet_format_properties(
    xml: &[u8],
) -> Result<Option<WorksheetSheetFormatProperties>> {
    // MCE removes attributes in ignorable namespaces. Capture the one x14ac
    // attribute implemented here before preprocessing without claiming support
    // for the complete x14ac namespace.
    let raw_dy_descent = capture_raw_dy_descent(xml)?;
    let processed =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let mut parsed = parse_processed(processed.xml.as_ref())?;
    if let (Some(properties), Some(dy_descent)) = (&mut parsed, raw_dy_descent) {
        if properties
            .dy_descent
            .is_some_and(|value| value != dy_descent)
        {
            return Err(invalid("conflicting x14ac:dyDescent values"));
        }
        properties.dy_descent = Some(dy_descent);
        properties.custom_height = true;
    }
    Ok(parsed)
}

fn parse_processed(xml: &[u8]) -> Result<Option<WorksheetSheetFormatProperties>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut leaf_depth = None;
    let mut properties = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    if root_seen || !core || local.as_ref() != b"worksheet" {
                        return Err(invalid("sheetFormatPr parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("sheetFormatPr is a leaf element"));
                } else if depth == 1 && core && local.as_ref() == b"sheetFormatPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate worksheet sheetFormatPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish()?);
                    leaf_depth = Some(depth + 1);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    return Err(invalid("worksheet root cannot be empty"));
                }
                if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("sheetFormatPr is a leaf element"));
                }
                if depth == 1 && core && local.as_ref() == b"sheetFormatPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate worksheet sheetFormatPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish()?);
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet end element"));
                }
                if leaf_depth == Some(depth) {
                    leaf_depth = None;
                }
                depth -= 1;
            },
            Event::Text(text) if leaf_depth.is_some_and(|value| depth >= value)
                && !text.decode().map_err(xml_error)?.trim().is_empty() => {
                    return Err(invalid("sheetFormatPr cannot contain text"));
                },
            Event::CData(_) if leaf_depth.is_some_and(|value| depth >= value) => {
                return Err(invalid("sheetFormatPr cannot contain CDATA"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 || leaf_depth.is_some() {
        return Err(invalid("unterminated worksheet XML"));
    }
    Ok(properties)
}

fn capture_raw_dy_descent(xml: &[u8]) -> Result<Option<f64>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut value = None;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                capture_element_dy_descent(
                    depth, &namespace, &element, decoder, &resolver, &mut value,
                )?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => capture_element_dy_descent(
                depth, &namespace, &element, decoder, &resolver, &mut value,
            )?,
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet end element"));
                }
                depth -= 1;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(value)
}

fn capture_element_dy_descent(
    depth: usize,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    result: &mut Option<f64>,
) -> Result<()> {
    let local = element.local_name();
    if depth == 1
        && is_spreadsheetml_name(namespace, element.name(), local.as_ref())
        && local.as_ref() == b"sheetFormatPr"
    {
        if let Some(value) = parse_dy_descent_attribute(element, decoder, resolver)? {
            set_once(result, value, "x14ac:dyDescent")?;
        }
    }
    Ok(())
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Builder> {
    let mut builder = Builder::default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match namespace {
            ResolveResult::Unbound => match local.as_ref() {
                b"baseColWidth" => set_once(
                    &mut builder.base_column_width,
                    parse_bounded_u32(&value, "baseColWidth", MAX_BASE_COLUMN_WIDTH)?,
                    "baseColWidth",
                )?,
                b"defaultColWidth" => set_once(
                    &mut builder.default_column_width,
                    parse_default_column_width(&value)?,
                    "defaultColWidth",
                )?,
                b"defaultRowHeight" => set_once(
                    &mut builder.default_row_height,
                    parse_nonnegative_f64(&value, "defaultRowHeight")?,
                    "defaultRowHeight",
                )?,
                b"customHeight" => set_once(
                    &mut builder.custom_height,
                    parse_bool(&value, "customHeight")?,
                    "customHeight",
                )?,
                b"zeroHeight" => set_once(
                    &mut builder.zero_height,
                    parse_bool(&value, "zeroHeight")?,
                    "zeroHeight",
                )?,
                b"thickTop" => set_once(
                    &mut builder.thick_top,
                    parse_bool(&value, "thickTop")?,
                    "thickTop",
                )?,
                b"thickBottom" => set_once(
                    &mut builder.thick_bottom,
                    parse_bool(&value, "thickBottom")?,
                    "thickBottom",
                )?,
                b"outlineLevelRow" => set_once(
                    &mut builder.outline_level_row,
                    parse_outline_level(&value, "outlineLevelRow")?,
                    "outlineLevelRow",
                )?,
                b"outlineLevelCol" => set_once(
                    &mut builder.outline_level_column,
                    parse_outline_level(&value, "outlineLevelCol")?,
                    "outlineLevelCol",
                )?,
                name => {
                    return Err(invalid(format!(
                        "unknown sheetFormatPr attribute '{}'",
                        String::from_utf8_lossy(name)
                    )));
                },
            },
            ResolveResult::Bound(uri)
                if uri.as_ref() == X14AC_NAMESPACE && local.as_ref() == b"dyDescent" =>
            {
                set_once(
                    &mut builder.dy_descent,
                    parse_nonnegative_f64(&value, "x14ac:dyDescent")?,
                    "x14ac:dyDescent",
                )?;
            },
            _ => {
                return Err(invalid(format!(
                    "unknown namespaced sheetFormatPr attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                )));
            },
        }
    }
    Ok(builder)
}

fn parse_dy_descent_attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<f64>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(uri) if uri.as_ref() == X14AC_NAMESPACE)
            && local.as_ref() == b"dyDescent"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            set_once(
                &mut result,
                parse_nonnegative_f64(&value, "x14ac:dyDescent")?,
                "x14ac:dyDescent",
            )?;
        }
    }
    Ok(result)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(format!("duplicate {name} attribute")));
    }
    *slot = Some(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid {name} boolean '{value}'"))),
    }
}

fn parse_bounded_u32(value: &str, name: &str, maximum: u32) -> Result<u32> {
    let parsed = value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid {name} value '{value}'")))?;
    if parsed > maximum {
        return Err(invalid(format!("{name} exceeds Office maximum {maximum}")));
    }
    Ok(parsed)
}

fn parse_outline_level(value: &str, name: &str) -> Result<u8> {
    let parsed = value
        .parse::<u8>()
        .map_err(|_| invalid(format!("invalid {name} value '{value}'")))?;
    if parsed > MAX_OUTLINE_LEVEL {
        return Err(invalid(format!("{name} exceeds Office maximum 7")));
    }
    Ok(parsed)
}

fn parse_default_column_width(value: &str) -> Result<f64> {
    let parsed = parse_nonnegative_f64(value, "defaultColWidth")?;
    if parsed >= MAX_DEFAULT_COLUMN_WIDTH {
        return Err(invalid("defaultColWidth must be less than 65536"));
    }
    Ok(parsed)
}

fn parse_nonnegative_f64(value: &str, name: &str) -> Result<f64> {
    let parsed = value
        .parse::<f64>()
        .map_err(|_| invalid(format!("invalid {name} value '{value}'")))?;
    if !parsed.is_finite() || parsed < 0.0 {
        return Err(invalid(format!("{name} must be finite and non-negative")));
    }
    Ok(parsed)
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    invalid(format!("invalid worksheet sheetFormatPr XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetSheetFormatProperties>> {
        parse_worksheet_sheet_format_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_all_core_attributes_and_effective_defaults() {
        let value = parse(concat!(
            r#"<sheetFormatPr baseColWidth="9" defaultColWidth="11.5" "#,
            r#"defaultRowHeight="18.25" customHeight="false" zeroHeight="1" "#,
            r#"thickTop="true" thickBottom="1" outlineLevelRow="4" outlineLevelCol="3"/>"#,
        ))
        .unwrap()
        .unwrap();
        assert_eq!(value.base_column_width(), 9);
        assert_eq!(value.default_column_width(), Some(11.5));
        assert_eq!(value.effective_default_column_width(), 11.5);
        assert_eq!(value.default_row_height(), 18.25);
        assert!(!value.custom_height());
        assert!(value.zero_height() && value.thick_top() && value.thick_bottom());
        assert_eq!(value.outline_level_row(), 4);
        assert_eq!(value.outline_level_column(), 3);

        let defaults = parse(r#"<sheetFormatPr defaultRowHeight="15"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(defaults.base_column_width(), 8);
        assert_eq!(defaults.default_column_width(), None);
        assert_eq!(defaults.effective_default_column_width(), 8.0);
        assert!(!defaults.zero_height());
        assert_eq!(defaults.outline_level_row(), 0);
    }

    #[test]
    fn supports_strict_namespace_and_direct_child_only() {
        let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetFormatPr defaultRowHeight="16"/></worksheet>"#;
        assert_eq!(
            parse_worksheet_sheet_format_properties(strict)
                .unwrap()
                .unwrap()
                .default_row_height(),
            16.0
        );
        let nested = format!(
            r#"<worksheet xmlns="{NS}"><wrapper><sheetFormatPr defaultRowHeight="15"/></wrapper></worksheet>"#
        );
        assert!(
            parse_worksheet_sheet_format_properties(nested.as_bytes())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn dy_descent_survives_mce_and_forces_custom_height() {
        let xml = format!(
            concat!(
                r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
                r#"xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">"#,
                r#"<sheetFormatPr defaultRowHeight="15" customHeight="0" x14ac:dyDescent="0.25"/></worksheet>"#,
            ),
            NS
        );
        let value = parse_worksheet_sheet_format_properties(xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(value.dy_descent(), Some(0.25));
        assert!(value.custom_height());
    }

    #[test]
    fn rejects_invalid_bounds_attributes_and_leaf_content() {
        for child in [
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelRow="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelCol="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" baseColWidth="256"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" defaultColWidth="65536"/>"#,
            r#"<sheetFormatPr defaultRowHeight="NaN"/>"#,
            r#"<sheetFormatPr/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" mystery="1"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15"><child/></sheetFormatPr>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
        assert!(
            parse(concat!(
                r#"<sheetFormatPr defaultRowHeight="15"/>"#,
                r#"<sheetFormatPr defaultRowHeight="15"/>"#
            ))
            .is_err()
        );
    }

    fn fixture_sheet(bytes: &[u8]) -> WorksheetSheetFormatProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        parse_worksheet_sheet_format_properties(part.blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn reads_libreoffice_hidden_default_rows_fixture() {
        let value = fixture_sheet(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf105840_allRowsHidden.xlsx"
        )));
        assert_eq!(value.default_row_height(), 15.0);
        assert!(value.zero_height());
    }

    #[test]
    fn reads_libreoffice_custom_dimensions_fixture() {
        let value = fixture_sheet(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf120168.xlsx"
        )));
        assert_eq!(value.default_column_width(), Some(21.85546875));
        assert_eq!(value.default_row_height(), 39.0);
        assert_eq!(value.dy_descent(), Some(0.25));
        assert!(value.custom_height());
    }
}

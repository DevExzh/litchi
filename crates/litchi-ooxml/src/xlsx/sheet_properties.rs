//! Immutable XLSX worksheet sheet-properties read model.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::is_spreadsheetml_name;
use crate::xlsx::outline_properties::{
    WorksheetOutlineProperties, parse_worksheet_outline_properties,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// A bounded A1 anchor used to synchronize worksheet window positions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetSynchronizationReference {
    value: String,
    row: u32,
    column: u32,
}

impl WorksheetSynchronizationReference {
    pub fn value(&self) -> &str { &self.value }
    pub fn row(&self) -> u32 { self.row }
    pub fn column(&self) -> u32 { self.column }
}

/// Sheet-tab color metadata from `sheetPr/tabColor`.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetTabColor {
    automatic: bool,
    indexed: Option<u32>,
    argb: Option<[u8; 4]>,
    theme: Option<u32>,
    tint: f64,
}

impl WorksheetTabColor {
    pub fn automatic(&self) -> bool { self.automatic }
    pub fn indexed(&self) -> Option<u32> { self.indexed }
    pub fn argb(&self) -> Option<[u8; 4]> { self.argb }
    pub fn theme(&self) -> Option<u32> { self.theme }
    pub fn tint(&self) -> f64 { self.tint }
}

/// Effective page-setup flags from `sheetPr/pageSetUpPr`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorksheetPageSetupProperties {
    automatic_page_breaks: bool,
    fit_to_page: bool,
}

impl WorksheetPageSetupProperties {
    pub fn automatic_page_breaks(&self) -> bool { self.automatic_page_breaks }
    pub fn fit_to_page(&self) -> bool { self.fit_to_page }
}

/// Complete immutable metadata from a worksheet's direct `sheetPr` child.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetSheetProperties {
    code_name: Option<String>,
    synchronization_reference: Option<WorksheetSynchronizationReference>,
    synchronize_horizontally: bool,
    synchronize_vertically: bool,
    transition_evaluation: bool,
    transition_entry: bool,
    published: bool,
    filter_mode: bool,
    format_condition_calculation_enabled: bool,
    tab_color: Option<WorksheetTabColor>,
    outline_properties: Option<WorksheetOutlineProperties>,
    page_setup_properties: Option<WorksheetPageSetupProperties>,
}

impl WorksheetSheetProperties {
    /// Stable VBA-facing name, retained as inert metadata only.
    pub fn code_name(&self) -> Option<&str> { self.code_name.as_deref() }
    pub fn synchronization_reference(&self) -> Option<&WorksheetSynchronizationReference> {
        self.synchronization_reference.as_ref()
    }
    pub fn synchronize_horizontally(&self) -> bool { self.synchronize_horizontally }
    pub fn synchronize_vertically(&self) -> bool { self.synchronize_vertically }
    pub fn transition_evaluation_enabled(&self) -> bool { self.transition_evaluation }
    pub fn transition_entry_enabled(&self) -> bool { self.transition_entry }
    pub fn published(&self) -> bool { self.published }
    pub fn filter_mode(&self) -> bool { self.filter_mode }
    pub fn format_condition_calculation_enabled(&self) -> bool {
        self.format_condition_calculation_enabled
    }
    pub fn tab_color(&self) -> Option<&WorksheetTabColor> { self.tab_color.as_ref() }
    pub fn outline_properties(&self) -> Option<&WorksheetOutlineProperties> {
        self.outline_properties.as_ref()
    }
    pub fn page_setup_properties(&self) -> Option<&WorksheetPageSetupProperties> {
        self.page_setup_properties.as_ref()
    }
}

#[derive(Default)]
struct Builder {
    code_name: Option<String>,
    synchronization_reference: Option<WorksheetSynchronizationReference>,
    synchronize_horizontally: Option<bool>,
    synchronize_vertically: Option<bool>,
    transition_evaluation: Option<bool>,
    transition_entry: Option<bool>,
    published: Option<bool>,
    filter_mode: Option<bool>,
    format_condition_calculation_enabled: Option<bool>,
    tab_color: Option<WorksheetTabColor>,
    page_setup_properties: Option<WorksheetPageSetupProperties>,
}

impl Builder {
    fn finish(self, outline_properties: Option<WorksheetOutlineProperties>) -> WorksheetSheetProperties {
        WorksheetSheetProperties {
            code_name: self.code_name,
            synchronization_reference: self.synchronization_reference,
            synchronize_horizontally: self.synchronize_horizontally.unwrap_or(false),
            synchronize_vertically: self.synchronize_vertically.unwrap_or(false),
            transition_evaluation: self.transition_evaluation.unwrap_or(false),
            transition_entry: self.transition_entry.unwrap_or(false),
            published: self.published.unwrap_or(true),
            filter_mode: self.filter_mode.unwrap_or(false),
            format_condition_calculation_enabled:
                self.format_condition_calculation_enabled.unwrap_or(true),
            tab_color: self.tab_color,
            outline_properties,
            page_setup_properties: self.page_setup_properties,
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context { Outside, Worksheet, SheetProperties, TabColor, OutlineProperties, PageSetupProperties }

struct Parser {
    stack: Vec<Context>,
    builder: Option<Builder>,
    seen_sheet_properties: bool,
    seen_tab_color: bool,
    seen_outline_properties: bool,
    seen_page_setup_properties: bool,
}

/// Parse the worksheet's exact `worksheet/sheetPr` child path.
pub fn parse_worksheet_sheet_properties(xml: &[u8]) -> Result<Option<WorksheetSheetProperties>> {
    let outline_properties = parse_worksheet_outline_properties(xml)?;
    let processed = process_markup_compatibility(
        xml,
        &MceCapabilities::default(),
        &MceLimits::default(),
    )?;
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut parser = Parser {
        stack: Vec::new(), builder: None, seen_sheet_properties: false,
        seen_tab_color: false, seen_outline_properties: false,
        seen_page_setup_properties: false,
    };
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => parser.start(&namespace, &element, decoder, &resolver)?,
            Event::Empty(element) => parser.empty(&namespace, &element, decoder, &resolver)?,
            Event::End(element) => parser.end(element.local_name().as_ref())?,
            Event::Text(text) if matches!(parser.parent(),
                Context::SheetProperties | Context::TabColor | Context::OutlineProperties
                    | Context::PageSetupProperties) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("unexpected text in worksheet sheetPr"));
                }
            }
            Event::CData(_) if matches!(parser.parent(),
                Context::SheetProperties | Context::TabColor | Context::OutlineProperties
                    | Context::PageSetupProperties) => {
                return Err(invalid("unexpected CDATA in worksheet sheetPr"));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !parser.stack.is_empty() { return Err(invalid("unterminated worksheet sheetPr XML")); }
    Ok(parser.builder.map(|builder| builder.finish(outline_properties)))
}

impl Parser {
    fn parent(&self) -> Context {
        self.stack.last().copied().unwrap_or(Context::Outside)
    }

    fn start(
        &mut self, namespace: &ResolveResult<'_>, element: &BytesStart<'_>,
        decoder: Decoder, resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() && (!core || local.as_ref() != b"worksheet") {
            return Err(invalid("sheetPr parser requires a worksheet root"));
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Outside, true, b"worksheet") => self.stack.push(Context::Worksheet),
            (Context::Worksheet, true, b"sheetPr") => {
                self.begin_sheet_properties(element, decoder, resolver)?;
                self.stack.push(Context::SheetProperties);
            }
            (Context::SheetProperties, true, b"tabColor") => {
                self.add_tab_color(element, decoder, resolver)?;
                self.stack.push(Context::TabColor);
            }
            (Context::SheetProperties, true, b"outlinePr") => {
                self.add_outline_properties()?;
                self.stack.push(Context::OutlineProperties);
            }
            (Context::SheetProperties, true, b"pageSetUpPr") => {
                self.add_page_setup_properties(element, decoder, resolver)?;
                self.stack.push(Context::PageSetupProperties);
            }
            (Context::SheetProperties, _, name) => return Err(invalid(format!(
                "unknown sheetPr child '{}'", String::from_utf8_lossy(name)))),
            (Context::TabColor | Context::OutlineProperties | Context::PageSetupProperties, _, _) => {
                return Err(invalid("sheetPr property children must be leaf elements"));
            }
            _ => self.stack.push(Context::Outside),
        }
        Ok(())
    }

    fn empty(
        &mut self, namespace: &ResolveResult<'_>, element: &BytesStart<'_>,
        decoder: Decoder, resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() { return Err(invalid("worksheet root cannot be empty")); }
        match (self.parent(), core, local.as_ref()) {
            (Context::Worksheet, true, b"sheetPr") => {
                self.begin_sheet_properties(element, decoder, resolver)?;
            }
            (Context::SheetProperties, true, b"tabColor") => {
                self.add_tab_color(element, decoder, resolver)?;
            }
            (Context::SheetProperties, true, b"outlinePr") => self.add_outline_properties()?,
            (Context::SheetProperties, true, b"pageSetUpPr") => {
                self.add_page_setup_properties(element, decoder, resolver)?;
            }
            (Context::SheetProperties, _, name) => return Err(invalid(format!(
                "unknown sheetPr child '{}'", String::from_utf8_lossy(name)))),
            (Context::TabColor | Context::OutlineProperties | Context::PageSetupProperties, _, _) => {
                return Err(invalid("sheetPr property children must be leaf elements"));
            }
            _ => {}
        }
        Ok(())
    }

    fn end(&mut self, local: &[u8]) -> Result<()> {
        let context = self.stack.pop().ok_or_else(|| invalid("unexpected sheetPr end element"))?;
        match context {
            Context::Worksheet if local == b"worksheet" => Ok(()),
            Context::SheetProperties if local == b"sheetPr" => Ok(()),
            Context::TabColor if local == b"tabColor" => Ok(()),
            Context::OutlineProperties if local == b"outlinePr" => Ok(()),
            Context::PageSetupProperties if local == b"pageSetUpPr" => Ok(()),
            Context::Outside => Ok(()),
            _ => Err(invalid("mismatched worksheet sheetPr end element")),
        }
    }

    fn begin_sheet_properties(
        &mut self, element: &BytesStart<'_>, decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_sheet_properties { return Err(invalid("duplicate worksheet sheetPr element")); }
        self.seen_sheet_properties = true;
        self.builder = Some(parse_sheet_attributes(element, decoder, resolver)?);
        Ok(())
    }

    fn add_tab_color(
        &mut self, element: &BytesStart<'_>, decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_tab_color { return Err(invalid("duplicate worksheet tabColor element")); }
        self.seen_tab_color = true;
        self.builder.as_mut().expect("sheetPr builder")
            .tab_color = Some(parse_tab_color(element, decoder, resolver)?);
        Ok(())
    }

    fn add_outline_properties(&mut self) -> Result<()> {
        if self.seen_outline_properties { return Err(invalid("duplicate worksheet outlinePr element")); }
        self.seen_outline_properties = true;
        Ok(())
    }

    fn add_page_setup_properties(
        &mut self, element: &BytesStart<'_>, decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_page_setup_properties {
            return Err(invalid("duplicate worksheet pageSetUpPr element"));
        }
        self.seen_page_setup_properties = true;
        self.builder.as_mut().expect("sheetPr builder").page_setup_properties =
            Some(parse_page_setup_properties(element, decoder, resolver)?);
        Ok(())
    }
}

fn parse_sheet_attributes(
    element: &BytesStart<'_>, decoder: Decoder, resolver: &NamespaceResolver,
) -> Result<Builder> {
    let mut builder = Builder::default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!("unknown namespaced sheetPr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref()))));
        }
        let value = attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"codeName" => set_once(&mut builder.code_name, value.into_owned(), "sheetPr codeName")?,
            b"syncRef" => set_once(&mut builder.synchronization_reference,
                parse_reference(&value)?, "sheetPr syncRef")?,
            b"syncHorizontal" => set_once(&mut builder.synchronize_horizontally,
                parse_bool(&value, "sheetPr syncHorizontal")?, "sheetPr syncHorizontal")?,
            b"syncVertical" => set_once(&mut builder.synchronize_vertically,
                parse_bool(&value, "sheetPr syncVertical")?, "sheetPr syncVertical")?,
            b"transitionEvaluation" => set_once(&mut builder.transition_evaluation,
                parse_bool(&value, "sheetPr transitionEvaluation")?, "sheetPr transitionEvaluation")?,
            b"transitionEntry" => set_once(&mut builder.transition_entry,
                parse_bool(&value, "sheetPr transitionEntry")?, "sheetPr transitionEntry")?,
            b"published" => set_once(&mut builder.published,
                parse_bool(&value, "sheetPr published")?, "sheetPr published")?,
            b"filterMode" => set_once(&mut builder.filter_mode,
                parse_bool(&value, "sheetPr filterMode")?, "sheetPr filterMode")?,
            b"enableFormatConditionsCalculation" => set_once(
                &mut builder.format_condition_calculation_enabled,
                parse_bool(&value, "sheetPr enableFormatConditionsCalculation")?,
                "sheetPr enableFormatConditionsCalculation")?,
            name => return Err(invalid(format!("unknown sheetPr attribute '{}'",
                String::from_utf8_lossy(name)))),
        }
    }
    Ok(builder)
}

fn parse_tab_color(
    element: &BytesStart<'_>, decoder: Decoder, resolver: &NamespaceResolver,
) -> Result<WorksheetTabColor> {
    let mut automatic = None;
    let mut indexed = None;
    let mut argb = None;
    let mut theme = None;
    let mut tint = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!("unknown namespaced tabColor attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref()))));
        }
        let value = attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"auto" => set_once(&mut automatic, parse_bool(&value, "tabColor auto")?, "tabColor auto")?,
            b"indexed" => set_once(&mut indexed, parse_u32(&value, "tabColor indexed")?, "tabColor indexed")?,
            b"rgb" => set_once(&mut argb, parse_argb(&value)?, "tabColor rgb")?,
            b"theme" => set_once(&mut theme, parse_u32(&value, "tabColor theme")?, "tabColor theme")?,
            b"tint" => set_once(&mut tint, parse_tint(&value)?, "tabColor tint")?,
            name => return Err(invalid(format!("unknown tabColor attribute '{}'",
                String::from_utf8_lossy(name)))),
        }
    }
    Ok(WorksheetTabColor {
        automatic: automatic.unwrap_or(false), indexed, argb, theme,
        tint: tint.unwrap_or(0.0),
    })
}

fn parse_page_setup_properties(
    element: &BytesStart<'_>, decoder: Decoder, resolver: &NamespaceResolver,
) -> Result<WorksheetPageSetupProperties> {
    let mut automatic_page_breaks = None;
    let mut fit_to_page = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) { continue; }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!("unknown namespaced pageSetUpPr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref()))));
        }
        let value = attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"autoPageBreaks" => set_once(&mut automatic_page_breaks,
                parse_bool(&value, "pageSetUpPr autoPageBreaks")?, "pageSetUpPr autoPageBreaks")?,
            b"fitToPage" => set_once(&mut fit_to_page,
                parse_bool(&value, "pageSetUpPr fitToPage")?, "pageSetUpPr fitToPage")?,
            name => return Err(invalid(format!("unknown pageSetUpPr attribute '{}'",
                String::from_utf8_lossy(name)))),
        }
    }
    Ok(WorksheetPageSetupProperties {
        automatic_page_breaks: automatic_page_breaks.unwrap_or(true),
        fit_to_page: fit_to_page.unwrap_or(false),
    })
}

fn parse_reference(value: &str) -> Result<WorksheetSynchronizationReference> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_uppercase) { index += 1; }
    if index == column_start { return Err(invalid(format!("invalid sheetPr syncRef '{value}'"))); }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column.checked_mul(26)
            .and_then(|value| value.checked_add(u32::from(*byte - b'A' + 1)))
            .ok_or_else(|| invalid(format!("invalid sheetPr syncRef '{value}'")))?;
    }
    if bytes.get(index) == Some(&b'$') { index += 1; }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) { index += 1; }
    if index != bytes.len() || index == row_start {
        return Err(invalid(format!("invalid sheetPr syncRef '{value}'")));
    }
    let row = value[row_start..].parse::<u32>()
        .map_err(|_| invalid(format!("invalid sheetPr syncRef '{value}'")))?;
    if !(1..=MAX_ROW).contains(&row) || !(1..=MAX_COLUMN).contains(&column) {
        return Err(invalid(format!("sheetPr syncRef is outside worksheet bounds: '{value}'")));
    }
    Ok(WorksheetSynchronizationReference { value: value.to_owned(), row, column })
}

fn parse_argb(value: &str) -> Result<[u8; 4]> {
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(format!("invalid tabColor ARGB value '{value}'")));
    }
    let mut result = [0u8; 4];
    for (index, byte) in result.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16)
            .map_err(|_| invalid(format!("invalid tabColor ARGB value '{value}'")))?;
    }
    Ok(result)
}

fn parse_tint(value: &str) -> Result<f64> {
    let tint = value.parse::<f64>()
        .map_err(|_| invalid(format!("invalid tabColor tint '{value}'")))?;
    if !tint.is_finite() || !(-1.0..=1.0).contains(&tint) {
        return Err(invalid(format!("tabColor tint outside -1..1: '{value}'")));
    }
    Ok(tint)
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    value.parse::<u32>().map_err(|_| invalid(format!("invalid {name} unsigned integer '{value}'")))
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid {name} boolean '{value}'"))),
    }
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() { return Err(invalid(format!("duplicate {name} attribute"))); }
    *slot = Some(value);
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot")
            && !name.starts_with('#')
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
    invalid(format!("invalid worksheet sheetPr XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetSheetProperties>> {
        parse_worksheet_sheet_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_complete_metadata_and_effective_defaults() {
        let value = parse(concat!(
            r#"<sheetPr codeName="CodeSheet" syncRef="$XFD$1048576" syncHorizontal="1" "#,
            r#"syncVertical="true" transitionEvaluation="1" transitionEntry="true" "#,
            r#"published="0" filterMode="1" enableFormatConditionsCalculation="false">"#,
            r#"<tabColor auto="0" indexed="64" rgb="FF00B050" theme="11" tint="-0.25"/>"#,
            r#"<outlinePr summaryBelow="0"/><pageSetUpPr autoPageBreaks="0" fitToPage="1"/>"#,
            r#"</sheetPr>"#,
        )).unwrap().unwrap();
        assert_eq!(value.code_name(), Some("CodeSheet"));
        let reference = value.synchronization_reference().unwrap();
        assert_eq!((reference.value(), reference.row(), reference.column()),
            ("$XFD$1048576", 1_048_576, 16_384));
        assert!(value.synchronize_horizontally());
        assert!(value.synchronize_vertically());
        assert!(value.transition_evaluation_enabled());
        assert!(value.transition_entry_enabled());
        assert!(!value.published());
        assert!(value.filter_mode());
        assert!(!value.format_condition_calculation_enabled());
        let color = value.tab_color().unwrap();
        assert!(!color.automatic());
        assert_eq!(color.indexed(), Some(64));
        assert_eq!(color.argb(), Some([0xFF, 0x00, 0xB0, 0x50]));
        assert_eq!(color.theme(), Some(11));
        assert_eq!(color.tint(), -0.25);
        assert!(!value.outline_properties().unwrap().summary_below());
        let page = value.page_setup_properties().unwrap();
        assert!(!page.automatic_page_breaks());
        assert!(page.fit_to_page());

        let defaults = parse("<sheetPr><tabColor/><outlinePr/><pageSetUpPr/></sheetPr>")
            .unwrap().unwrap();
        assert_eq!(defaults.code_name(), None);
        assert_eq!(defaults.synchronization_reference(), None);
        assert!(!defaults.synchronize_horizontally());
        assert!(!defaults.synchronize_vertically());
        assert!(!defaults.transition_evaluation_enabled());
        assert!(!defaults.transition_entry_enabled());
        assert!(defaults.published());
        assert!(!defaults.filter_mode());
        assert!(defaults.format_condition_calculation_enabled());
        assert!(!defaults.tab_color().unwrap().automatic());
        assert_eq!(defaults.tab_color().unwrap().tint(), 0.0);
        assert!(defaults.page_setup_properties().unwrap().automatic_page_breaks());
        assert!(!defaults.page_setup_properties().unwrap().fit_to_page());
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn supports_strict_mce_and_exact_scoping() {
        let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetPr filterMode="1"><tabColor rgb="FF112233"/></sheetPr></worksheet>"#;
        assert!(parse_worksheet_sheet_properties(strict).unwrap().unwrap().filter_mode());
        let mce = format!(concat!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x">"#,
            r#"<sheetPr x:future="1"><mc:AlternateContent><mc:Choice Requires="x"><x:pageSetUpPr/>"#,
            r#"</mc:Choice><mc:Fallback><pageSetUpPr fitToPage="1"/></mc:Fallback>"#,
            r#"</mc:AlternateContent></sheetPr></worksheet>"#,
        ), NS);
        assert!(parse_worksheet_sheet_properties(mce.as_bytes()).unwrap().unwrap()
            .page_setup_properties().unwrap().fit_to_page());
        assert!(parse("<wrapper><sheetPr filterMode=\"1\"/></wrapper>").unwrap().is_none());
        assert!(parse("<sheetPr><wrapper><tabColor/></wrapper></sheetPr>").is_err());
    }

    #[test]
    fn rejects_invalid_values_duplicates_spoofing_and_leaf_content() {
        for child in [
            r#"<sheetPr published="yes"/>"#,
            r#"<sheetPr syncRef="XFE1"/>"#,
            r#"<sheetPr syncRef="A0"/>"#,
            r#"<sheetPr syncRef="a1"/>"#,
            r#"<sheetPr syncRef="A1:B2"/>"#,
            r#"<sheetPr mystery="1"/>"#,
            r#"<sheetPr xmlns:x="urn:x" x:filterMode="1"/>"#,
            r#"<sheetPr><tabColor rgb="00FF00"/></sheetPr>"#,
            r#"<sheetPr><tabColor tint="1.01"/></sheetPr>"#,
            r#"<sheetPr><tabColor tint="NaN"/></sheetPr>"#,
            r#"<sheetPr><tabColor theme="-1"/></sheetPr>"#,
            r#"<sheetPr><tabColor mystery="1"/></sheetPr>"#,
            r#"<sheetPr xmlns:x="urn:x"><tabColor x:rgb="FF00FF00"/></sheetPr>"#,
            r#"<sheetPr><pageSetUpPr fitToPage="maybe"/></sheetPr>"#,
            r#"<sheetPr><pageSetUpPr mystery="1"/></sheetPr>"#,
            r#"<sheetPr><tabColor><child/></tabColor></sheetPr>"#,
            r#"<sheetPr><pageSetUpPr>text</pageSetUpPr></sheetPr>"#,
            r#"<sheetPr><unknown/></sheetPr>"#,
            r#"<sheetPr><tabColor/><tabColor/></sheetPr>"#,
            r#"<sheetPr><pageSetUpPr/><pageSetUpPr/></sheetPr>"#,
            r#"<sheetPr/><sheetPr/>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
    }

    fn fixture(bytes: &[u8]) -> WorksheetSheetProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package.get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap()).unwrap();
        parse_worksheet_sheet_properties(part.blob()).unwrap().unwrap()
    }

    #[test]
    fn reads_libreoffice_and_poi_fixtures() {
        let color = fixture(include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/sheet-tab-color.xlsx")));
        assert_eq!(color.tab_color().unwrap().argb(), Some([0xFF, 0x00, 0xB0, 0x50]));

        let page = fixture(include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/page_scale.xlsx")));
        assert!(!page.filter_mode());
        assert!(!page.page_setup_properties().unwrap().fit_to_page());

        let outline = fixture(include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/poi/test-data/spreadsheet/66365.xlsx")));
        assert!(!outline.outline_properties().unwrap().summary_below());
        assert!(!outline.outline_properties().unwrap().summary_right());
    }
}

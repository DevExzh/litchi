//! Complete immutable worksheet page-setup metadata.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};

use super::namespace::relationship_attribute_value;

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;

/// Printed page orientation.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PageSetupOrientation {
    #[default]
    Default,
    Portrait,
    Landscape,
}

/// Order in which worksheet pages are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PageSetupOrder {
    #[default]
    DownThenOver,
    OverThenDown,
}

/// How cell comments are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PageSetupCellComments {
    #[default]
    None,
    AsDisplayed,
    AtEnd,
}

/// How cell errors are printed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PageSetupPrintErrors {
    #[default]
    Displayed,
    Blank,
    Dash,
    NotAvailable,
}

/// Unit identifier from `ST_PositiveUniversalMeasure`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UniversalMeasureUnit {
    Millimeter,
    Centimeter,
    Inch,
    Point,
    Pica,
    PicaAlternative,
}

/// Positive custom paper dimension with its original unit retained.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PositiveUniversalMeasure {
    value: f64,
    unit: UniversalMeasureUnit,
}

impl PositiveUniversalMeasure {
    pub fn value(self) -> f64 {
        self.value
    }
    pub fn unit(self) -> UniversalMeasureUnit {
        self.unit
    }

    pub fn inches(self) -> f64 {
        match self.unit {
            UniversalMeasureUnit::Millimeter => self.value / 25.4,
            UniversalMeasureUnit::Centimeter => self.value / 2.54,
            UniversalMeasureUnit::Inch => self.value,
            UniversalMeasureUnit::Point => self.value / 72.0,
            UniversalMeasureUnit::Pica | UniversalMeasureUnit::PicaAlternative => self.value / 6.0,
        }
    }

    pub fn millimeters(self) -> f64 {
        self.inches() * 25.4
    }
}

/// Complete effective settings from one worksheet `pageSetup` element.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetPageSetup {
    paper_size: u32,
    paper_width: Option<PositiveUniversalMeasure>,
    paper_height: Option<PositiveUniversalMeasure>,
    scale: u32,
    first_page_number: u32,
    fit_to_width: u32,
    fit_to_height: u32,
    page_order: PageSetupOrder,
    orientation: PageSetupOrientation,
    use_printer_defaults: bool,
    black_and_white: bool,
    draft: bool,
    cell_comments: PageSetupCellComments,
    use_first_page_number: bool,
    print_errors: PageSetupPrintErrors,
    horizontal_dpi: u32,
    vertical_dpi: u32,
    copies: u32,
    printer_settings_relationship_id: Option<String>,
}

impl WorksheetPageSetup {
    pub fn paper_size(&self) -> u32 {
        self.paper_size
    }
    pub fn paper_width(&self) -> Option<PositiveUniversalMeasure> {
        self.paper_width
    }
    pub fn paper_height(&self) -> Option<PositiveUniversalMeasure> {
        self.paper_height
    }
    pub fn scale(&self) -> u32 {
        self.scale
    }
    pub fn first_page_number(&self) -> u32 {
        self.first_page_number
    }
    pub fn fit_to_width(&self) -> u32 {
        self.fit_to_width
    }
    pub fn fit_to_height(&self) -> u32 {
        self.fit_to_height
    }
    pub fn page_order(&self) -> PageSetupOrder {
        self.page_order
    }
    pub fn orientation(&self) -> PageSetupOrientation {
        self.orientation
    }
    pub fn use_printer_defaults(&self) -> bool {
        self.use_printer_defaults
    }
    pub fn black_and_white(&self) -> bool {
        self.black_and_white
    }
    pub fn draft(&self) -> bool {
        self.draft
    }
    pub fn cell_comments(&self) -> PageSetupCellComments {
        self.cell_comments
    }
    pub fn use_first_page_number(&self) -> bool {
        self.use_first_page_number
    }
    pub fn print_errors(&self) -> PageSetupPrintErrors {
        self.print_errors
    }
    pub fn horizontal_dpi(&self) -> u32 {
        self.horizontal_dpi
    }
    pub fn vertical_dpi(&self) -> u32 {
        self.vertical_dpi
    }
    pub fn copies(&self) -> u32 {
        self.copies
    }
    pub fn printer_settings_relationship_id(&self) -> Option<&str> {
        self.printer_settings_relationship_id.as_deref()
    }
}

impl Default for WorksheetPageSetup {
    fn default() -> Self {
        Self {
            paper_size: 1,
            paper_width: None,
            paper_height: None,
            scale: 100,
            first_page_number: 1,
            fit_to_width: 1,
            fit_to_height: 1,
            page_order: PageSetupOrder::DownThenOver,
            orientation: PageSetupOrientation::Default,
            use_printer_defaults: false,
            black_and_white: false,
            draft: false,
            cell_comments: PageSetupCellComments::None,
            use_first_page_number: false,
            print_errors: PageSetupPrintErrors::Displayed,
            horizontal_dpi: 600,
            vertical_dpi: 600,
            copies: 1,
            printer_settings_relationship_id: None,
        }
    }
}

/// Parse a worksheet's optional complete core `pageSetup` element.
pub fn parse_complete_worksheet_page_setup(xml: &[u8]) -> Result<Option<WorksheetPageSetup>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let validated =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    parse_selected(selected)
}

fn parse_selected(xml: &[u8]) -> Result<Option<WorksheetPageSetup>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut result = None;
    let mut open: Option<(usize, WorksheetPageSetup)> = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if depth == 1 {
                    if root_seen
                        || !spreadsheet(&namespace)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("page-setup parser requires a worksheet root"));
                    }
                    root_seen = true;
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"pageSetup"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet pageSetup element"));
                    }
                    open = Some((depth, parse_setup(&element, decoder, &resolver)?));
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Empty(element) => {
                if depth == 1
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"pageSetup"
                {
                    if result.is_some() || open.is_some() {
                        return Err(invalid("duplicate worksheet pageSetup element"));
                    }
                    result = Some(parse_setup(&element, decoder, &resolver)?);
                } else if open.is_some() {
                    return Err(invalid("pageSetup must not contain child elements"));
                }
            },
            Event::Text(text) => {
                if open.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("pageSetup must not contain text"));
                }
            },
            Event::CData(_) if open.is_some() => {
                return Err(invalid("pageSetup must not contain CDATA"));
            },
            Event::End(element) => {
                if open
                    .as_ref()
                    .is_some_and(|(element_depth, _)| *element_depth == depth)
                {
                    let (_, setup) = open.take().expect("checked above");
                    result = Some(setup);
                }
                if depth == 1 {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("invalid worksheet closing element"));
                    }
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
            },
            Event::GeneralRef(reference) => {
                if reference.resolve_char_ref().map_err(xml_error)?.is_none()
                    && !matches!(
                        reference.decode().map_err(xml_error)?.as_ref(),
                        "amp" | "lt" | "gt" | "apos" | "quot"
                    )
                {
                    return Err(invalid("custom XML entities are rejected"));
                }
                if open.is_some() {
                    return Err(invalid("pageSetup must not contain entity text"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) | Event::CData(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || open.is_some() {
        return Err(invalid("incomplete worksheet page-setup XML"));
    }
    Ok(result)
}

fn parse_setup(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<WorksheetPageSetup> {
    let mut setup = WorksheetPageSetup::default();
    let mut seen = [false; 18];
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref().contains(&b':') {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let slot = match attribute.key.local_name().as_ref() {
            b"paperSize" => {
                setup.paper_size = parse_u32(&value, "paperSize")?;
                0
            },
            b"paperWidth" => {
                setup.paper_width = Some(parse_measure(&value, "paperWidth")?);
                1
            },
            b"paperHeight" => {
                setup.paper_height = Some(parse_measure(&value, "paperHeight")?);
                2
            },
            b"scale" => {
                setup.scale = parse_u32(&value, "scale")?;
                if !(10..=400).contains(&setup.scale) {
                    return Err(invalid("pageSetup scale must be between 10 and 400"));
                }
                3
            },
            b"firstPageNumber" => {
                setup.first_page_number = parse_u32(&value, "firstPageNumber")?;
                4
            },
            b"fitToWidth" => {
                setup.fit_to_width = parse_u32(&value, "fitToWidth")?;
                5
            },
            b"fitToHeight" => {
                setup.fit_to_height = parse_u32(&value, "fitToHeight")?;
                6
            },
            b"pageOrder" => {
                setup.page_order = parse_order(&value)?;
                7
            },
            b"orientation" => {
                setup.orientation = parse_orientation(&value)?;
                8
            },
            b"usePrinterDefaults" => {
                setup.use_printer_defaults = parse_bool(&value, "usePrinterDefaults")?;
                9
            },
            b"blackAndWhite" => {
                setup.black_and_white = parse_bool(&value, "blackAndWhite")?;
                10
            },
            b"draft" => {
                setup.draft = parse_bool(&value, "draft")?;
                11
            },
            b"cellComments" => {
                setup.cell_comments = parse_comments(&value)?;
                12
            },
            b"useFirstPageNumber" => {
                setup.use_first_page_number = parse_bool(&value, "useFirstPageNumber")?;
                13
            },
            b"errors" => {
                setup.print_errors = parse_errors(&value)?;
                14
            },
            b"horizontalDpi" => {
                setup.horizontal_dpi = parse_u32(&value, "horizontalDpi")?;
                15
            },
            b"verticalDpi" => {
                setup.vertical_dpi = parse_u32(&value, "verticalDpi")?;
                16
            },
            b"copies" => {
                setup.copies = parse_u32(&value, "copies")?;
                17
            },
            name => {
                return Err(invalid(format!(
                    "unknown pageSetup attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        };
        if seen[slot] {
            return Err(invalid("duplicate pageSetup attribute"));
        }
        seen[slot] = true;
    }
    setup.printer_settings_relationship_id =
        relationship_attribute_value(element, b"id", decoder, resolver)?;
    if let Some(id) = setup.printer_settings_relationship_id.as_ref() {
        if id.is_empty() || id.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(invalid(
                "invalid pageSetup printer-settings relationship id",
            ));
        }
    }
    Ok(setup)
}

fn parse_measure(raw: &str, field: &str) -> Result<PositiveUniversalMeasure> {
    let (number, unit) = [
        ("mm", UniversalMeasureUnit::Millimeter),
        ("cm", UniversalMeasureUnit::Centimeter),
        ("in", UniversalMeasureUnit::Inch),
        ("pt", UniversalMeasureUnit::Point),
        ("pc", UniversalMeasureUnit::Pica),
        ("pi", UniversalMeasureUnit::PicaAlternative),
    ]
    .into_iter()
    .find_map(|(suffix, unit)| raw.strip_suffix(suffix).map(|number| (number, unit)))
    .ok_or_else(|| invalid(format!("invalid {field} unit")))?;
    if number.is_empty() || number.starts_with('-') {
        return Err(invalid(format!("{field} must be positive")));
    }
    let value = number
        .parse::<f64>()
        .map_err(|_| invalid(format!("invalid {field} measure")))?;
    if !value.is_finite() || value <= 0.0 {
        return Err(invalid(format!("{field} must be finite and positive")));
    }
    Ok(PositiveUniversalMeasure { value, unit })
}

fn parse_u32(raw: &str, field: &str) -> Result<u32> {
    raw.parse()
        .map_err(|_| invalid(format!("invalid pageSetup {field}")))
}
fn parse_bool(raw: &str, field: &str) -> Result<bool> {
    match raw {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid pageSetup {field} boolean"))),
    }
}
fn parse_orientation(raw: &str) -> Result<PageSetupOrientation> {
    match raw {
        "default" => Ok(PageSetupOrientation::Default),
        "portrait" => Ok(PageSetupOrientation::Portrait),
        "landscape" => Ok(PageSetupOrientation::Landscape),
        _ => Err(invalid("invalid pageSetup orientation")),
    }
}
fn parse_order(raw: &str) -> Result<PageSetupOrder> {
    match raw {
        "downThenOver" => Ok(PageSetupOrder::DownThenOver),
        "overThenDown" => Ok(PageSetupOrder::OverThenDown),
        _ => Err(invalid("invalid pageSetup pageOrder")),
    }
}
fn parse_comments(raw: &str) -> Result<PageSetupCellComments> {
    match raw {
        "none" => Ok(PageSetupCellComments::None),
        "asDisplayed" => Ok(PageSetupCellComments::AsDisplayed),
        "atEnd" => Ok(PageSetupCellComments::AtEnd),
        _ => Err(invalid("invalid pageSetup cellComments")),
    }
}
fn parse_errors(raw: &str) -> Result<PageSetupPrintErrors> {
    match raw {
        "displayed" => Ok(PageSetupPrintErrors::Displayed),
        "blank" => Ok(PageSetupPrintErrors::Blank),
        "dash" => Ok(PageSetupPrintErrors::Dash),
        "NA" => Ok(PageSetupPrintErrors::NotAvailable),
        _ => Err(invalid("invalid pageSetup errors")),
    }
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    exact(namespace, CORE) || exact(namespace, STRICT)
}
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const START: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#;

    fn parse(body: &str) -> Result<Option<WorksheetPageSetup>> {
        parse_complete_worksheet_page_setup(format!("{START}{body}</worksheet>").as_bytes())
    }
    fn parse_fixture(path: &str) -> WorksheetPageSetup {
        let package = OpcPackage::open(path).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        parse_complete_worksheet_page_setup(package.get_part(&uri).unwrap().blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn parses_complete_settings_and_custom_measures() {
        let setup = parse(r#"<pageSetup paperSize="9" paperWidth="21cm" paperHeight="297mm" scale="125" firstPageNumber="3" fitToWidth="2" fitToHeight="4" pageOrder="overThenDown" orientation="landscape" usePrinterDefaults="1" blackAndWhite="true" draft="1" cellComments="atEnd" useFirstPageNumber="true" errors="NA" horizontalDpi="1200" verticalDpi="600" copies="2" r:id="rId7"/>"#).unwrap().unwrap();
        assert_eq!(setup.paper_size(), 9);
        assert!((setup.paper_width().unwrap().millimeters() - 210.0).abs() < 1e-10);
        assert_eq!(setup.page_order(), PageSetupOrder::OverThenDown);
        assert_eq!(setup.orientation(), PageSetupOrientation::Landscape);
        assert_eq!(setup.cell_comments(), PageSetupCellComments::AtEnd);
        assert_eq!(setup.print_errors(), PageSetupPrintErrors::NotAvailable);
        assert_eq!(setup.printer_settings_relationship_id(), Some("rId7"));
    }

    #[test]
    fn applies_schema_defaults_and_preserves_absence() {
        let setup = parse("<pageSetup/>").unwrap().unwrap();
        assert_eq!(setup.paper_size(), 1);
        assert_eq!(setup.scale(), 100);
        assert_eq!(setup.orientation(), PageSetupOrientation::Default);
        assert_eq!(setup.horizontal_dpi(), 600);
        assert_eq!(setup.copies(), 1);
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn rejects_bad_scale_enums_measures_and_content() {
        assert!(parse(r#"<pageSetup scale="9"/>"#).is_err());
        assert!(parse(r#"<pageSetup orientation="sideways"/>"#).is_err());
        assert!(parse(r#"<pageSetup paperWidth="0mm"/>"#).is_err());
        assert!(parse(r#"<pageSetup errors="na"/>"#).is_err());
        assert!(parse(r#"<pageSetup><x/></pageSetup>"#).is_err());
    }

    #[test]
    fn loads_poi_dpi_sentinel_and_relationship_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/45540_classic_Header.xlsx"
        );
        let setup = parse_fixture(path);
        assert_eq!(setup.orientation(), PageSetupOrientation::Portrait);
        assert_eq!(setup.horizontal_dpi(), 1200);
        assert_eq!(setup.vertical_dpi(), 1200);
        assert_eq!(setup.printer_settings_relationship_id(), Some("rId1"));
    }

    #[test]
    fn loads_libreoffice_paper_and_orientation_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf136721_letter_sized_paper.xlsx"
        );
        let setup = parse_fixture(path);
        assert_eq!(setup.paper_size(), 70);
        assert_eq!(setup.orientation(), PageSetupOrientation::Landscape);
        assert_eq!(setup.horizontal_dpi(), 0);
        assert_eq!(setup.printer_settings_relationship_id(), Some("rId1"));
    }
}

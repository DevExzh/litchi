//! Conservative closure checks for value-only cell publication.

use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, allocation, invalid};
use crate::raw;

const TRANSITIONAL_SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

#[cfg(test)]
#[path = "validation_borrow_tests.rs"]
mod borrow_tests;
#[cfg(test)]
#[path = "shared_traversal_tests.rs"]
mod shared_traversal_tests;

pub(super) fn workbook_xml(content: &[u8]) -> Result<()> {
    validate_xml(content, XmlOwner::Workbook)
}

pub(super) fn worksheet_xml(content: &[u8]) -> Result<()> {
    validate_xml(content, XmlOwner::Worksheet)
}

#[derive(Clone, Copy)]
enum XmlOwner {
    Workbook,
    Worksheet,
}

/// Validation state for the shared source traversal.
///
/// The state owns only the local element names required by the existing
/// closure checks. Event payloads remain borrowed for the duration of the
/// callback. On an observer failure the driver stops immediately and the
/// retained error is discarded before authoritative fallback validation.
struct Validator {
    owner: XmlOwner,
    depth: usize,
    elements: Vec<Box<[u8]>>,
    dialect: Option<Box<[u8]>>,
    saw_root: bool,
    first_error: Option<Error>,
}

impl Validator {
    fn new(owner: XmlOwner) -> Self {
        Self {
            owner,
            depth: 0,
            elements: Vec::new(),
            dialect: None,
            saw_root: false,
            first_error: None,
        }
    }

    fn observe(&mut self, namespace: &ResolveResult<'_>, event: &Event<'_>) -> bool {
        if self.first_error.is_some() {
            return false;
        }
        if let Err(error) = self.observe_inner(namespace, event) {
            self.first_error = Some(error);
            return false;
        }
        if matches!(event, Event::Eof) && (!self.saw_root || self.depth != 0) {
            self.first_error = Some(invalid("value-only XML has no complete root element"));
            return false;
        }
        true
    }

    fn observe_inner(&mut self, namespace: &ResolveResult<'_>, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(element) => {
                bind_dialect(namespace, &mut self.dialect)?;
                let local = element
                    .name()
                    .local_name()
                    .as_ref()
                    .to_vec()
                    .into_boxed_slice();
                validate_element(
                    self.owner,
                    namespace,
                    element,
                    &local,
                    self.elements.last().map(AsRef::as_ref),
                    self.depth,
                )?;
                self.saw_root = true;
                self.depth = self
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("value-only XML depth overflow"))?;
                self.elements
                    .try_reserve(1)
                    .map_err(|source| allocation("value-only XML element stack", source))?;
                self.elements.push(local);
            },
            Event::Empty(element) => {
                bind_dialect(namespace, &mut self.dialect)?;
                let local = element.name().local_name().as_ref().to_vec();
                validate_element(
                    self.owner,
                    namespace,
                    element,
                    &local,
                    self.elements.last().map(AsRef::as_ref),
                    self.depth,
                )?;
                self.saw_root = true;
            },
            Event::End(element) => {
                self.depth = self
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("value-only XML has an unmatched closing element"))?;
                let expected = self
                    .elements
                    .pop()
                    .ok_or_else(|| invalid("value-only XML has no open element to close"))?;
                if element.local_name().as_ref() != expected.as_ref()
                    || !matches!((namespace, self.dialect.as_deref()), (ResolveResult::Bound(Namespace(value)), Some(expected)) if *value == expected)
                {
                    return Err(invalid(
                        "value-only XML has a mismatched or foreign closing element",
                    ));
                }
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "value-only edits refuse XML document type declarations",
                ));
            },
            Event::Eof => {},
            Event::Text(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(
                    self.owner,
                    self.elements.last().map(AsRef::as_ref),
                    &decoded,
                ) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::CData(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(
                    self.owner,
                    self.elements.last().map(AsRef::as_ref),
                    &decoded,
                ) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::GeneralRef(_) => {
                if !text_context_allowed(self.owner, self.elements.last().map(AsRef::as_ref)) {
                    return Err(invalid(
                        "value-only XML has a reference outside a scalar value element",
                    ));
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if let Some(error) = self.first_error {
            return Err(error);
        }
        if !self.saw_root || self.depth != 0 {
            return Err(invalid("value-only XML has no complete root element"));
        }
        Ok(())
    }
}

/// Validate and parse an eligible source using one borrowed `NsReader`.
///
/// Early reader, parser-transition, or provisional-cap uncertainty is
/// converted into a provisional failure. After the observer accepts the
/// complete source, parser materialization is carried as a `Complete` result;
/// the caller finishes validation before forwarding that result through the
/// raw facade. No speculative diagnostic escapes the established error order.
pub(super) fn worksheet_xml_and_parse_source(content: &[u8]) -> Result<crate::cell::Store> {
    let mut validator = Validator::new(XmlOwner::Worksheet);
    let attempt = raw::worksheet::parse_source_with_observer(
        content,
        || Ok(None),
        |namespace, event| validator.observe(namespace, event),
    );
    match attempt {
        raw::worksheet::SourceParseAttempt::Complete(parsed) => {
            validator.finish()?;
            raw::worksheet::complete_source_parse(content, parsed)
        },
        raw::worksheet::SourceParseAttempt::ProvisionalFailed => {
            drop(validator);
            worksheet_xml(content)?;
            raw::worksheet::parse(content, || Ok(None))
        },
        raw::worksheet::SourceParseAttempt::ReaderFailed => {
            drop(validator);
            worksheet_xml(content)?;
            raw::worksheet::parse(content, || Ok(None))
        },
    }
}

fn validate_xml(content: &[u8], owner: XmlOwner) -> Result<()> {
    let mut reader = NsReader::from_reader(content);
    let mut depth = 0usize;
    let mut elements = Vec::<Box<[u8]>>::new();
    let mut dialect = None::<Box<[u8]>>;
    let mut saw_root = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("value-only XML scan failed: {error}")))?;
        // Events borrow the input slice; namespaces are inspected before the
        // next read advances the resolver's scope.
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                bind_dialect(&namespace, &mut dialect)?;
                let local = element
                    .name()
                    .local_name()
                    .as_ref()
                    .to_vec()
                    .into_boxed_slice();
                validate_element(
                    owner,
                    &namespace,
                    &element,
                    &local,
                    elements.last().map(AsRef::as_ref),
                    depth,
                )?;
                saw_root = true;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("value-only XML depth overflow"))?;
                elements.push(local);
            },
            Event::Empty(element) => {
                bind_dialect(&namespace, &mut dialect)?;
                let local = element.name().local_name().as_ref().to_vec();
                validate_element(
                    owner,
                    &namespace,
                    &element,
                    &local,
                    elements.last().map(AsRef::as_ref),
                    depth,
                )?;
                saw_root = true;
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("value-only XML has an unmatched closing element"))?;
                let expected = elements
                    .pop()
                    .ok_or_else(|| invalid("value-only XML has no open element to close"))?;
                if element.local_name().as_ref() != expected.as_ref()
                    || !matches!((&namespace, dialect.as_deref()), (ResolveResult::Bound(Namespace(value)), Some(expected)) if *value == expected)
                {
                    return Err(invalid(
                        "value-only XML has a mismatched or foreign closing element",
                    ));
                }
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "value-only edits refuse XML document type declarations",
                ));
            },
            Event::Eof => break,
            Event::Text(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(owner, elements.last().map(AsRef::as_ref), &decoded) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::CData(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(owner, elements.last().map(AsRef::as_ref), &decoded) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::GeneralRef(_) => {
                if !text_context_allowed(owner, elements.last().map(AsRef::as_ref)) {
                    return Err(invalid(
                        "value-only XML has a reference outside a scalar value element",
                    ));
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
    if !saw_root || depth != 0 {
        return Err(invalid("value-only XML has no complete root element"));
    }
    Ok(())
}

fn bind_dialect(namespace: &ResolveResult<'_>, dialect: &mut Option<Box<[u8]>>) -> Result<()> {
    let ResolveResult::Bound(Namespace(value)) = namespace else {
        return Err(invalid("value-only XML has an unbound element namespace"));
    };
    if *value != TRANSITIONAL_SML && *value != STRICT_SML {
        return Err(invalid("value-only XML has a foreign element namespace"));
    }
    match dialect {
        Some(expected) if expected.as_ref() != *value => {
            Err(invalid("value-only XML mixes SpreadsheetML dialects"))
        },
        Some(_) => Ok(()),
        None => {
            *dialect = Some(value.to_vec().into_boxed_slice());
            Ok(())
        },
    }
}

fn text_allowed(owner: XmlOwner, context: Option<&[u8]>, value: &str) -> bool {
    value.trim().is_empty() || text_context_allowed(owner, context)
}

fn text_context_allowed(owner: XmlOwner, context: Option<&[u8]>) -> bool {
    matches!(owner, XmlOwner::Worksheet) && matches!(context, Some(b"f" | b"v" | b"t"))
}

fn validate_element(
    owner: XmlOwner,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
    parent: Option<&[u8]>,
    depth: usize,
) -> Result<()> {
    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_SML || *value == STRICT_SML)
    {
        return Err(invalid(format!(
            "value-only edits refuse foreign or markup-compatible element '{}'",
            String::from_utf8_lossy(local)
        )));
    }
    let allowed = match owner {
        XmlOwner::Workbook => matches!(
            local,
            b"workbook"
                | b"fileVersion"
                | b"workbookPr"
                | b"bookViews"
                | b"workbookView"
                | b"sheets"
                | b"sheet"
                | b"calcPr"
        ),
        XmlOwner::Worksheet => matches!(
            local,
            b"worksheet"
                | b"dimension"
                | b"sheetViews"
                | b"sheetView"
                | b"pane"
                | b"selection"
                | b"sheetFormatPr"
                | b"cols"
                | b"col"
                | b"sheetData"
                | b"row"
                | b"c"
                | b"f"
                | b"v"
                | b"is"
                | b"t"
        ),
    };
    if !allowed {
        return Err(invalid(format!(
            "value-only edits refuse dependency-bearing or unknown element '{}'",
            String::from_utf8_lossy(local)
        )));
    }
    validate_parent(owner, local, parent)?;
    validate_attributes(owner, element, local)?;
    let expected_root = match owner {
        XmlOwner::Workbook => b"workbook".as_slice(),
        XmlOwner::Worksheet => b"worksheet".as_slice(),
    };
    if depth == 0 && local != expected_root {
        return Err(invalid("value-only XML has the wrong root element"));
    }
    Ok(())
}

fn validate_parent(owner: XmlOwner, local: &[u8], parent: Option<&[u8]>) -> Result<()> {
    let valid = match owner {
        XmlOwner::Workbook => matches!(
            (parent, local),
            (None, b"workbook")
                | (
                    Some(b"workbook"),
                    b"fileVersion" | b"workbookPr" | b"bookViews" | b"sheets" | b"calcPr"
                )
                | (Some(b"bookViews"), b"workbookView")
                | (Some(b"sheets"), b"sheet")
        ),
        XmlOwner::Worksheet => matches!(
            (parent, local),
            (None, b"worksheet")
                | (
                    Some(b"worksheet"),
                    b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols" | b"sheetData"
                )
                | (Some(b"sheetViews"), b"sheetView")
                | (Some(b"sheetView"), b"pane" | b"selection")
                | (Some(b"cols"), b"col")
                | (Some(b"sheetData"), b"row")
                | (Some(b"row"), b"c")
                | (Some(b"c"), b"f" | b"v" | b"is")
                | (Some(b"is"), b"t")
        ),
    };
    if !valid {
        return Err(invalid(format!(
            "value-only edits refuse element '{}' in this XML context",
            String::from_utf8_lossy(local)
        )));
    }
    Ok(())
}

fn validate_attributes(owner: XmlOwner, element: &BytesStart<'_>, local: &[u8]) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid value-only XML attribute: {error}")))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if name == b"r:id" && matches!(owner, XmlOwner::Workbook) && local == b"sheet" {
            continue;
        }
        if name == b"xml:space" && matches!(owner, XmlOwner::Worksheet) && local == b"t" {
            continue;
        }
        if matches!(owner, XmlOwner::Workbook) && local == b"calcPr" {
            continue;
        }
        if name.contains(&b':') || !allowed_unqualified_attribute(owner, local, name) {
            return Err(invalid(format!(
                "value-only edits refuse attribute '{}' on '{}'",
                String::from_utf8_lossy(name),
                String::from_utf8_lossy(local)
            )));
        }
    }
    Ok(())
}

fn allowed_unqualified_attribute(owner: XmlOwner, local: &[u8], name: &[u8]) -> bool {
    match owner {
        XmlOwner::Workbook => match local {
            b"fileVersion" => matches!(
                name,
                b"appName" | b"lastEdited" | b"lowestEdited" | b"rupBuild" | b"codeName"
            ),
            b"workbookPr" => matches!(
                name,
                b"date1904"
                    | b"showObjects"
                    | b"showBorderUnselectedTables"
                    | b"filterPrivacy"
                    | b"promptedSolutions"
                    | b"showInkAnnotation"
                    | b"backupFile"
                    | b"saveExternalLinkValues"
                    | b"updateLinks"
                    | b"codeName"
                    | b"hidePivotFieldList"
                    | b"showPivotChartFilter"
                    | b"allowRefreshQuery"
                    | b"publishItems"
                    | b"checkCompatibility"
                    | b"autoCompressPictures"
                    | b"refreshAllConnections"
                    | b"defaultThemeVersion"
            ),
            b"workbookView" => matches!(
                name,
                b"visibility"
                    | b"minimized"
                    | b"showHorizontalScroll"
                    | b"showVerticalScroll"
                    | b"showSheetTabs"
                    | b"xWindow"
                    | b"yWindow"
                    | b"windowWidth"
                    | b"windowHeight"
                    | b"tabRatio"
                    | b"firstSheet"
                    | b"activeTab"
                    | b"autoFilterDateGrouping"
            ),
            b"sheet" => matches!(name, b"name" | b"sheetId" | b"state"),
            b"calcPr" => true,
            _ => false,
        },
        XmlOwner::Worksheet => match local {
            b"dimension" => name == b"ref",
            b"sheetView" => matches!(
                name,
                b"windowProtection"
                    | b"showFormulas"
                    | b"showGridLines"
                    | b"showRowColHeaders"
                    | b"showZeros"
                    | b"rightToLeft"
                    | b"tabSelected"
                    | b"showRuler"
                    | b"showOutlineSymbols"
                    | b"defaultGridColor"
                    | b"showWhiteSpace"
                    | b"view"
                    | b"topLeftCell"
                    | b"colorId"
                    | b"zoomScale"
                    | b"zoomScaleNormal"
                    | b"zoomScaleSheetLayoutView"
                    | b"zoomScalePageLayoutView"
                    | b"workbookViewId"
            ),
            b"pane" => matches!(
                name,
                b"xSplit" | b"ySplit" | b"topLeftCell" | b"activePane" | b"state"
            ),
            b"selection" => matches!(name, b"pane" | b"activeCell" | b"activeCellId" | b"sqref"),
            b"sheetFormatPr" => matches!(
                name,
                b"baseColWidth"
                    | b"defaultColWidth"
                    | b"defaultRowHeight"
                    | b"customHeight"
                    | b"zeroHeight"
                    | b"thickTop"
                    | b"thickBottom"
                    | b"outlineLevelRow"
                    | b"outlineLevelCol"
            ),
            b"col" => matches!(
                name,
                b"min"
                    | b"max"
                    | b"width"
                    | b"style"
                    | b"hidden"
                    | b"bestFit"
                    | b"customWidth"
                    | b"phonetic"
                    | b"outlineLevel"
                    | b"collapsed"
            ),
            b"row" => matches!(
                name,
                b"r" | b"spans"
                    | b"s"
                    | b"customFormat"
                    | b"ht"
                    | b"hidden"
                    | b"customHeight"
                    | b"outlineLevel"
                    | b"collapsed"
                    | b"thickTop"
                    | b"thickBot"
                    | b"ph"
            ),
            b"c" => matches!(name, b"r" | b"s" | b"t"),
            b"f" => matches!(
                name,
                b"t" | b"ref"
                    | b"si"
                    | b"dt2D"
                    | b"dtr"
                    | b"del1"
                    | b"del2"
                    | b"r1"
                    | b"r2"
                    | b"ca"
                    | b"bx"
            ),
            b"t" => false,
            _ => false,
        },
    }
}

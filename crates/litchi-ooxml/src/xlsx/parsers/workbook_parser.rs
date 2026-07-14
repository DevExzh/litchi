//! Namespace-aware streaming parser for `xl/workbook.xml`.

use std::collections::HashSet;

use litchi_core::sheet::Result as SheetResult;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::{
    SPREADSHEETML_NAMESPACE, is_spreadsheetml_name, relationship_attribute_value,
};
use crate::xlsx::worksheet::WorksheetInfo;

const INITIAL_SHEETS_CAPACITY: usize = 16;
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Clone, Copy, PartialEq, Eq)]
enum WorkbookContext {
    Workbook,
    Sheets,
    BookViews,
    DefinedNames,
    DefinedName,
    PivotCaches,
    Other,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DefinedName {
    pub(crate) name: String,
    pub(crate) local_sheet_id: Option<u32>,
    pub(crate) value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PivotCacheInfo {
    pub(crate) cache_id: u32,
    pub(crate) relationship_id: String,
}

pub(crate) struct WorkbookParseResult {
    pub(crate) sheets: Vec<WorksheetInfo>,
    pub(crate) active_sheet_index: usize,
    pub(crate) uses_1904_date_system: bool,
    pub(crate) defined_names: Vec<DefinedName>,
    pub(crate) pivot_caches: Vec<PivotCacheInfo>,
}

struct WorkbookInfo {
    sheets: Vec<WorksheetInfo>,
    active_tab: Option<usize>,
    uses_1904_date_system: bool,
    seen_workbook_properties: bool,
    seen_sheets: bool,
    seen_book_views: bool,
    seen_defined_names: bool,
    seen_pivot_caches: bool,
    seen_workbook_view: bool,
    sheet_ids: HashSet<u32>,
    relationship_ids: HashSet<String>,
    defined_name_keys: HashSet<(Option<u32>, String)>,
    defined_names: Vec<DefinedName>,
    pending_defined_name: Option<DefinedName>,
    pivot_cache_ids: HashSet<u32>,
    pivot_cache_relationship_ids: HashSet<String>,
    pivot_caches: Vec<PivotCacheInfo>,
}

impl WorkbookInfo {
    fn new() -> Self {
        Self {
            sheets: Vec::with_capacity(INITIAL_SHEETS_CAPACITY),
            active_tab: None,
            uses_1904_date_system: false,
            seen_workbook_properties: false,
            seen_sheets: false,
            seen_book_views: false,
            seen_defined_names: false,
            seen_pivot_caches: false,
            seen_workbook_view: false,
            sheet_ids: HashSet::new(),
            relationship_ids: HashSet::new(),
            defined_name_keys: HashSet::new(),
            defined_names: Vec::new(),
            pending_defined_name: None,
            pivot_cache_ids: HashSet::new(),
            pivot_cache_relationship_ids: HashSet::new(),
            pivot_caches: Vec::new(),
        }
    }

    fn parse(content: &str) -> Result<Self> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut info = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    if stack.is_empty() {
                        if closed_root
                            || !is_spreadsheetml_name(&namespace, element.name(), b"workbook")
                        {
                            return Err(invalid(
                                "workbook XML must have one SpreadsheetML workbook root",
                            ));
                        }
                        stack.push(WorkbookContext::Workbook);
                        continue;
                    }
                    let parent = current_context(&stack)?;
                    info.process_element(parent, &namespace, &element, decoder, &resolver)?;
                    stack.push(info.child_context(parent, &namespace, &element)?);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"workbook")
                    {
                        return Err(invalid(
                            "workbook XML must have one SpreadsheetML workbook root",
                        ));
                    }
                    closed_root = true;
                },
                Event::Empty(element) => {
                    let parent = current_context(&stack)?;
                    info.process_element(parent, &namespace, &element, decoder, &resolver)?;
                    info.observe_empty_container(parent, &namespace, &element)?;
                },
                Event::Text(text) if stack.last() == Some(&WorkbookContext::DefinedName) => {
                    info.push_defined_name_text(
                        &text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                    )?;
                },
                Event::CData(text) if stack.last() == Some(&WorkbookContext::DefinedName) => {
                    info.push_defined_name_text(
                        &text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                    )?;
                },
                Event::GeneralRef(reference)
                    if stack.last() == Some(&WorkbookContext::DefinedName) =>
                {
                    info.push_defined_name_text(&decode_xml_reference(&reference)?)?;
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        invalid("workbook XML has a closing element outside its root")
                    })?;
                    if context == WorkbookContext::DefinedName {
                        info.finish_defined_name()?;
                    }
                    if context == WorkbookContext::Workbook {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"workbook") {
                            return Err(invalid(
                                "workbook XML has an invalid root closing element",
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid(
                        "workbook XML has a missing or unterminated SpreadsheetML workbook root",
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if let Some(active_tab) = info.active_tab
            && active_tab >= info.sheets.len()
            && !info.sheets.is_empty()
        {
            return Err(invalid(format!(
                "workbook activeTab {active_tab} exceeds the {} available sheets",
                info.sheets.len()
            )));
        }
        for defined_name in &info.defined_names {
            if let Some(local_sheet_id) = defined_name.local_sheet_id
                && usize::try_from(local_sheet_id).map_or(true, |index| index >= info.sheets.len())
            {
                return Err(invalid(format!(
                    "defined name '{}' has out-of-range localSheetId {local_sheet_id}",
                    defined_name.name
                )));
            }
        }
        Ok(info)
    }

    fn process_element(
        &mut self,
        parent: WorkbookContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"workbookPr")
        {
            mark_once(&mut self.seen_workbook_properties, "workbookPr element")?;
            self.uses_1904_date_system =
                optional_bool(element, b"date1904", decoder, "workbook date1904")?.unwrap_or(false);
        } else if parent == WorkbookContext::Sheets
            && is_spreadsheetml_name(namespace, element.name(), b"sheet")
        {
            let sheet = parse_sheet_element(element, decoder, resolver)?;
            if !self.sheet_ids.insert(sheet.sheet_id) {
                return Err(invalid(format!(
                    "duplicate workbook sheet ID {}",
                    sheet.sheet_id
                )));
            }
            if !self.relationship_ids.insert(sheet.relationship_id.clone()) {
                return Err(invalid(format!(
                    "duplicate workbook sheet relationship ID '{}'",
                    sheet.relationship_id
                )));
            }
            self.sheets.push(sheet);
        } else if parent == WorkbookContext::BookViews
            && is_spreadsheetml_name(namespace, element.name(), b"workbookView")
        {
            let active_tab = optional_u32(element, b"activeTab", decoder, "workbook activeTab")?
                .map(|value| {
                    usize::try_from(value)
                        .map_err(|_| invalid("workbook activeTab does not fit usize"))
                })
                .transpose()?;
            if !self.seen_workbook_view {
                self.active_tab = active_tab;
                self.seen_workbook_view = true;
            }
        } else if parent == WorkbookContext::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            if self.pending_defined_name.is_some() {
                return Err(invalid("nested workbook definedName element"));
            }
            let name = required_string(element, b"name", decoder, "defined name")?;
            if name.is_empty() {
                return Err(invalid("workbook defined name cannot be empty"));
            }
            let local_sheet_id = optional_u32(
                element,
                b"localSheetId",
                decoder,
                "defined name localSheetId",
            )?;
            self.pending_defined_name = Some(DefinedName {
                name,
                local_sheet_id,
                value: String::new(),
            });
        } else if parent == WorkbookContext::PivotCaches
            && is_spreadsheetml_name(namespace, element.name(), b"pivotCache")
        {
            let cache_id = required_u32(element, b"cacheId", decoder, "pivot cache ID")?;
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| invalid("workbook pivot cache is missing relationship ID"))?;
            if relationship_id.is_empty() {
                return Err(invalid(
                    "workbook pivot cache relationship ID cannot be empty",
                ));
            }
            if !self.pivot_cache_ids.insert(cache_id) {
                return Err(invalid(format!(
                    "duplicate workbook pivot cache ID {cache_id}"
                )));
            }
            if !self
                .pivot_cache_relationship_ids
                .insert(relationship_id.clone())
            {
                return Err(invalid(format!(
                    "duplicate workbook pivot cache relationship ID '{relationship_id}'"
                )));
            }
            self.pivot_caches.push(PivotCacheInfo {
                cache_id,
                relationship_id,
            });
        }
        Ok(())
    }

    fn child_context(
        &mut self,
        parent: WorkbookContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<WorkbookContext> {
        if parent == WorkbookContext::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            return Ok(WorkbookContext::DefinedName);
        }
        if parent != WorkbookContext::Workbook {
            return Ok(WorkbookContext::Other);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"sheets") {
            mark_once(&mut self.seen_sheets, "sheets element")?;
            Ok(WorkbookContext::Sheets)
        } else if is_spreadsheetml_name(namespace, element.name(), b"bookViews") {
            mark_once(&mut self.seen_book_views, "bookViews element")?;
            Ok(WorkbookContext::BookViews)
        } else if is_spreadsheetml_name(namespace, element.name(), b"definedNames") {
            mark_once(&mut self.seen_defined_names, "definedNames element")?;
            Ok(WorkbookContext::DefinedNames)
        } else if is_spreadsheetml_name(namespace, element.name(), b"pivotCaches") {
            mark_once(&mut self.seen_pivot_caches, "pivotCaches element")?;
            Ok(WorkbookContext::PivotCaches)
        } else {
            Ok(WorkbookContext::Other)
        }
    }

    fn observe_empty_container(
        &mut self,
        parent: WorkbookContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"sheets")
        {
            mark_once(&mut self.seen_sheets, "sheets element")?;
        } else if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"bookViews")
        {
            mark_once(&mut self.seen_book_views, "bookViews element")?;
        } else if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            mark_once(&mut self.seen_defined_names, "definedNames element")?;
        } else if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"pivotCaches")
        {
            mark_once(&mut self.seen_pivot_caches, "pivotCaches element")?;
        } else if parent == WorkbookContext::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            self.finish_defined_name()?;
        }
        Ok(())
    }

    fn push_defined_name_text(&mut self, text: &str) -> Result<()> {
        self.pending_defined_name
            .as_mut()
            .ok_or_else(|| invalid("defined-name text outside a definedName element"))?
            .value
            .push_str(text);
        Ok(())
    }

    fn finish_defined_name(&mut self) -> Result<()> {
        let defined_name = self
            .pending_defined_name
            .take()
            .ok_or_else(|| invalid("missing pending workbook defined name"))?;
        let key = (
            defined_name.local_sheet_id,
            defined_name.name.to_ascii_lowercase(),
        );
        if !self.defined_name_keys.insert(key) {
            return Err(invalid(format!(
                "duplicate workbook defined name '{}' in the same scope",
                defined_name.name
            )));
        }
        self.defined_names.push(defined_name);
        Ok(())
    }
}

pub(crate) fn parse_workbook_details(content: &str) -> SheetResult<WorkbookParseResult> {
    WorkbookInfo::parse(content)
        .map(|info| WorkbookParseResult {
            sheets: info.sheets,
            active_sheet_index: info.active_tab.unwrap_or(0),
            uses_1904_date_system: info.uses_1904_date_system,
            defined_names: info.defined_names,
            pivot_caches: info.pivot_caches,
        })
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
}

/// Parse workbook metadata, returning sheets, the active sheet index, and the date system.
pub fn parse_workbook_xml(content: &str) -> SheetResult<(Vec<WorksheetInfo>, usize, bool)> {
    parse_workbook_details(content).map(|details| {
        (
            details.sheets,
            details.active_sheet_index,
            details.uses_1904_date_system,
        )
    })
}

/// Parse a standalone `sheet` element.
pub fn parse_sheet_xml(sheet_xml: &str) -> SheetResult<Option<WorksheetInfo>> {
    let wrapped = format!(
        r#"<workbook xmlns="{}" xmlns:r="{}"><sheets>{sheet_xml}</sheets></workbook>"#,
        String::from_utf8_lossy(SPREADSHEETML_NAMESPACE),
        RELATIONSHIPS_NAMESPACE
    );
    parse_workbook_xml(&wrapped).map(|(mut sheets, _, _)| sheets.pop())
}

fn parse_sheet_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<WorksheetInfo> {
    let name = required_string(element, b"name", decoder, "workbook sheet name")?;
    if name.is_empty() {
        return Err(invalid("workbook sheet name cannot be empty"));
    }
    let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| invalid("workbook sheet is missing relationship ID"))?;
    if relationship_id.is_empty() {
        return Err(invalid("workbook sheet relationship ID cannot be empty"));
    }
    let sheet_id = required_u32(element, b"sheetId", decoder, "workbook sheet ID")?;
    if sheet_id == 0 {
        return Err(invalid("workbook sheet ID must be positive"));
    }
    if let Some(state) = unqualified_attribute_value(element, b"state", decoder)?
        && !matches!(state.as_str(), "visible" | "hidden" | "veryHidden")
    {
        return Err(invalid(format!("invalid workbook sheet state '{state}'")));
    }

    Ok(WorksheetInfo {
        name,
        relationship_id,
        sheet_id,
        is_active: false,
        print_area: None,
        repeating_rows: None,
        repeating_columns: None,
    })
}

fn current_context(stack: &[WorkbookContext]) -> Result<WorkbookContext> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("workbook XML is missing its root context"))
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
        })
        .transpose()
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid {description} value '{value}'"))),
        })
        .transpose()
}

fn mark_once(seen: &mut bool, description: &str) -> Result<()> {
    if *seen {
        return Err(invalid(format!("duplicate workbook {description}")));
    }
    *seen = true;
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    #[test]
    fn parses_namespaced_workbook_metadata_and_decodes_attributes() {
        let xml = format!(
            r#"<x:workbook xmlns:x="{S}" xmlns:rel="{R}" xmlns:f="urn:foreign">
                <x:workbookPr date1904="true"/>
                <x:bookViews><x:workbookView activeTab="1"/><x:workbookView activeTab="0"/></x:bookViews>
                <f:sheets><x:sheet name="Ignored" sheetId="99" rel:id="ignored"/></f:sheets>
                <x:sheets>
                    <x:sheet name="A &amp; B" sheetId="1" rel:id="custom-one"/>
                    <x:sheet name="Hidden" sheetId="7" state="veryHidden" rel:id="custom-two"/>
                </x:sheets>
            </x:workbook>"#
        );
        let (sheets, active, date_1904) = parse_workbook_xml(&xml).unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "A & B");
        assert_eq!(sheets[0].relationship_id, "custom-one");
        assert_eq!(sheets[1].sheet_id, 7);
        assert_eq!(active, 1);
        assert!(date_1904);
    }

    #[test]
    fn supports_strict_namespaces_and_standalone_sheet_fragments() {
        let xml = format!(
            r#"<workbook xmlns="{STRICT_S}" xmlns:r="{STRICT_R}"><sheets>
                <sheet name="Strict" sheetId="4" r:id="strictRel"/>
            </sheets></workbook>"#
        );
        let (sheets, active, date_1904) = parse_workbook_xml(&xml).unwrap();
        assert_eq!(sheets[0].name, "Strict");
        assert_eq!(sheets[0].relationship_id, "strictRel");
        assert_eq!(active, 0);
        assert!(!date_1904);

        let sheet = parse_sheet_xml(r#"<sheet name="One &amp; Two" sheetId="2" r:id="r9"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(sheet.name, "One & Two");
        assert_eq!(sheet.relationship_id, "r9");
    }

    #[test]
    fn parses_namespaced_sheet_scoped_defined_names() {
        let xml = format!(
            r#"<s:workbook xmlns:s="{STRICT_S}" xmlns:r="{STRICT_R}" xmlns:f="urn:foreign">
                <s:sheets><s:sheet name="A &amp; B" sheetId="1" r:id="r1"/></s:sheets>
                <f:definedNames><s:definedName name="ignored" localSheetId="0">bad</s:definedName></f:definedNames>
                <s:definedNames>
                    <s:definedName name="_xlnm.Print_Area" localSheetId="0">'A &amp; B'!$A$1:$D$20</s:definedName>
                    <s:definedName name="GlobalName">42</s:definedName>
                </s:definedNames>
            </s:workbook>"#
        );
        let details = parse_workbook_details(&xml).unwrap();

        assert_eq!(details.defined_names.len(), 2);
        assert_eq!(details.defined_names[0].name, "_xlnm.Print_Area");
        assert_eq!(details.defined_names[0].local_sheet_id, Some(0));
        assert_eq!(details.defined_names[0].value, "'A & B'!$A$1:$D$20");
        assert_eq!(details.defined_names[1].value, "42");
    }

    #[test]
    fn parses_strict_workbook_pivot_cache_references() {
        let xml = format!(
            r#"<s:workbook xmlns:s="{STRICT_S}" xmlns:r="{STRICT_R}" xmlns:f="urn:foreign">
                <f:pivotCaches><s:pivotCache cacheId="99" r:id="ignored"/></f:pivotCaches>
                <s:pivotCaches>
                    <s:pivotCache cacheId="7" r:id="custom-cache"/>
                    <s:pivotCache cacheId="11" r:id="second-cache"/>
                </s:pivotCaches>
            </s:workbook>"#
        );
        let details = parse_workbook_details(&xml).unwrap();

        assert_eq!(
            details.pivot_caches,
            vec![
                PivotCacheInfo {
                    cache_id: 7,
                    relationship_id: "custom-cache".to_string(),
                },
                PivotCacheInfo {
                    cache_id: 11,
                    relationship_id: "second-cache".to_string(),
                },
            ]
        );
    }

    #[test]
    fn rejects_malformed_defined_names() {
        let sheets = r#"<sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets>"#;
        let invalid = [
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName localSheetId="0">x</definedName></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x" localSheetId="1">x</definedName></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="Same">1</definedName><definedName name="same">2</definedName></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames/><definedNames/></workbook>"#
            ),
        ];
        for xml in invalid {
            assert!(parse_workbook_details(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn rejects_malformed_workbook_pivot_cache_references() {
        let invalid = [
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><pivotCaches><pivotCache r:id="r1"/></pivotCaches></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><pivotCaches><pivotCache cacheId="1"/></pivotCaches></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><pivotCaches><pivotCache cacheId="1" r:id="r1"/><pivotCache cacheId="1" r:id="r2"/></pivotCaches></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><pivotCaches><pivotCache cacheId="1" r:id="r1"/><pivotCache cacheId="2" r:id="r1"/></pivotCaches></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><pivotCaches/><pivotCaches/></workbook>"#
            ),
        ];
        for xml in invalid {
            assert!(parse_workbook_details(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn rejects_spoofed_duplicate_and_malformed_workbook_metadata() {
        let foreign = format!(
            r#"<workbook xmlns="{S}" xmlns:f="urn:foreign" xmlns:r="{R}"><f:sheets>
                <sheet name="Nested" sheetId="1" r:id="r1"/>
            </f:sheets><sheets><sheet name="Real" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        let (sheets, _, _) = parse_workbook_xml(&foreign).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Real");

        let duplicate = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets>
                <sheet name="Same" sheetId="1" r:id="r1"/>
                <sheet name="Other" sheetId="1" r:id="r2"/>
            </sheets></workbook>"#
        );
        assert!(parse_workbook_xml(&duplicate).is_err());

        for xml in [
            format!(r#"<workbook xmlns="{S}"><workbookPr date1904="yes"/></workbook>"#),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Missing" sheetId="0" r:id="r1"/></sheets></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView activeTab="3"/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
            ),
            format!(r#"<workbook xmlns="{S}"><sheets>"#),
        ] {
            assert!(parse_workbook_xml(&xml).is_err(), "accepted {xml}");
        }
    }
}

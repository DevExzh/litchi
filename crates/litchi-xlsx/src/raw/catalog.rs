//! Namespace-aware streaming parser for `xl/workbook.xml`.

use std::collections::{HashMap, HashSet};

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::{Catalog, DefinedName, PivotCache, Sheet, Visibility};
use crate::error::{Result, invalid};
use crate::raw::namespace::{
    SPREADSHEETML_NAMESPACE, is_spreadsheetml_name, relationship_attribute_value,
};

const INITIAL_SHEETS_CAPACITY: usize = 16;
const MAX_SHEETS: usize = 32_767;
const MAX_SHEET_ID: u32 = 65_534;
const MAX_RELATIONSHIP_ID_CHARACTERS: usize = 255;
const MAX_DEFINED_NAMES: usize = 65_536;
const MAX_DEFINED_NAME_FORMULA_BYTES: usize = 1_048_576;
const MAX_DEFINED_NAME_CHARACTERS: usize = 255;
const MAX_DEFINED_NAME_COMMENT_CHARACTERS: usize = 255;
const MAX_LOCAL_SHEET_ID: u32 = 32_766;
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Clone, Copy, PartialEq, Eq)]
enum WorkbookContext {
    Workbook,
    Sheets,
    BookViews,
    DefinedNames,
    DefinedName,
    ExternalReferences,
    PivotCaches,
    Other,
}

struct WorkbookInfo {
    sheets: Vec<Sheet>,
    active_tab: Option<usize>,
    uses_1904_date_system: bool,
    seen_workbook_properties: bool,
    seen_sheets: bool,
    seen_book_views: bool,
    seen_defined_names: bool,
    seen_pivot_caches: bool,
    seen_external_references: bool,
    seen_workbook_view: bool,
    sheet_ids: HashSet<u32>,
    sheet_names: HashMap<Box<str>, usize>,
    relationship_ids: HashSet<String>,
    defined_name_keys: HashSet<(Option<u32>, String)>,
    defined_names: Vec<DefinedName>,
    pending_defined_name: Option<DefinedName>,
    pivot_cache_ids: HashSet<u32>,
    pivot_cache_relationship_ids: HashSet<String>,
    pivot_caches: Vec<PivotCache>,
    external_reference_ids: Vec<String>,
    external_reference_id_set: HashSet<String>,
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
            seen_external_references: false,
            seen_workbook_view: false,
            sheet_ids: HashSet::new(),
            sheet_names: HashMap::new(),
            relationship_ids: HashSet::new(),
            defined_name_keys: HashSet::new(),
            defined_names: Vec::new(),
            pending_defined_name: None,
            pivot_cache_ids: HashSet::new(),
            pivot_cache_relationship_ids: HashSet::new(),
            pivot_caches: Vec::new(),
            external_reference_ids: Vec::new(),
            external_reference_id_set: HashSet::new(),
        }
    }

    fn parse(content: &[u8]) -> Result<Self> {
        let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
        let content = std::str::from_utf8(processed.as_ref())
            .map_err(|error| invalid(format!("workbook XML is not UTF-8: {error}")))?;
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut info = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?
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
                        &text.decode().map_err(|error| invalid(error.to_string()))?,
                    )?;
                },
                Event::CData(text) if stack.last() == Some(&WorkbookContext::DefinedName) => {
                    info.push_defined_name_text(
                        &text.decode().map_err(|error| invalid(error.to_string()))?,
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
        if info.seen_external_references && info.external_reference_ids.is_empty() {
            return Err(invalid("workbook externalReferences cannot be empty"));
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
            if self.sheets.len() >= MAX_SHEETS {
                return Err(invalid(format!(
                    "workbook sheet count exceeds Office's {MAX_SHEETS}-sheet limit"
                )));
            }
            let sheet = parse_sheet_element(element, decoder, resolver)?;
            let position = self.sheets.len();
            let name_key = crate::sheet::key(&sheet.name);
            if let Some(first) = self.sheet_names.insert(name_key, position) {
                return Err(crate::Error::SheetNameConflict {
                    name: sheet.name,
                    first,
                    second: position,
                });
            }
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
            if self.defined_names.len() >= MAX_DEFINED_NAMES {
                return Err(invalid("workbook defined-name limit exceeded"));
            }
            let name = required_string(element, b"name", decoder, "defined name")?;
            let name_characters = name.chars().count();
            if name_characters == 0 || name_characters > MAX_DEFINED_NAME_CHARACTERS {
                return Err(invalid(format!(
                    "workbook defined name must contain 1 to {MAX_DEFINED_NAME_CHARACTERS} characters"
                )));
            }
            let local_sheet_id = optional_u32(
                element,
                b"localSheetId",
                decoder,
                "defined name localSheetId",
            )?;
            if local_sheet_id.is_some_and(|value| value > MAX_LOCAL_SHEET_ID) {
                return Err(invalid(format!(
                    "defined name localSheetId exceeds {MAX_LOCAL_SHEET_ID}"
                )));
            }
            let comment = optional_string(element, b"comment", decoder)?;
            if comment
                .as_ref()
                .is_some_and(|value| value.chars().count() > MAX_DEFINED_NAME_COMMENT_CHARACTERS)
            {
                return Err(invalid(format!(
                    "defined name comment exceeds {MAX_DEFINED_NAME_COMMENT_CHARACTERS} characters"
                )));
            }
            self.pending_defined_name = Some(DefinedName {
                name,
                local_sheet_id,
                reference: String::new(),
                comment,
                custom_menu: optional_string(element, b"customMenu", decoder)?,
                description: optional_string(element, b"description", decoder)?,
                help: optional_string(element, b"help", decoder)?,
                status_bar: optional_string(element, b"statusBar", decoder)?,
                shortcut_key: optional_string(element, b"shortcutKey", decoder)?,
                hidden: optional_bool(element, b"hidden", decoder, "defined name hidden")?
                    .unwrap_or(false),
                function: optional_bool(element, b"function", decoder, "defined name function")?
                    .unwrap_or(false),
                vb_procedure: optional_bool(
                    element,
                    b"vbProcedure",
                    decoder,
                    "defined name vbProcedure",
                )?
                .unwrap_or(false),
                xlm: optional_bool(element, b"xlm", decoder, "defined name xlm")?.unwrap_or(false),
                function_group_id: optional_u32(
                    element,
                    b"functionGroupId",
                    decoder,
                    "defined name functionGroupId",
                )?,
                publish_to_server: optional_bool(
                    element,
                    b"publishToServer",
                    decoder,
                    "defined name publishToServer",
                )?
                .unwrap_or(false),
                workbook_parameter: optional_bool(
                    element,
                    b"workbookParameter",
                    decoder,
                    "defined name workbookParameter",
                )?
                .unwrap_or(false),
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
            self.pivot_caches.push(PivotCache {
                cache_id,
                relationship_id,
            });
        } else if parent == WorkbookContext::ExternalReferences
            && is_spreadsheetml_name(namespace, element.name(), b"externalReference")
        {
            if self.external_reference_ids.len() >= 4096 {
                return Err(invalid("workbook external-reference limit exceeded"));
            }
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| invalid("workbook external reference is missing relationship ID"))?;
            if relationship_id.is_empty() {
                return Err(invalid(
                    "workbook external-reference relationship ID cannot be empty",
                ));
            }
            if !self
                .external_reference_id_set
                .insert(relationship_id.clone())
            {
                return Err(invalid(format!(
                    "duplicate workbook external-reference relationship ID '{relationship_id}'"
                )));
            }
            self.external_reference_ids.push(relationship_id);
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
        } else if is_spreadsheetml_name(namespace, element.name(), b"externalReferences") {
            mark_once(
                &mut self.seen_external_references,
                "externalReferences element",
            )?;
            Ok(WorkbookContext::ExternalReferences)
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
        } else if parent == WorkbookContext::Workbook
            && is_spreadsheetml_name(namespace, element.name(), b"externalReferences")
        {
            mark_once(
                &mut self.seen_external_references,
                "externalReferences element",
            )?;
            return Err(invalid("workbook externalReferences cannot be empty"));
        } else if parent == WorkbookContext::DefinedNames
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            self.finish_defined_name()?;
        }
        Ok(())
    }

    fn push_defined_name_text(&mut self, text: &str) -> Result<()> {
        let defined_name = self
            .pending_defined_name
            .as_mut()
            .ok_or_else(|| invalid("defined-name text outside a definedName element"))?;
        let new_len = defined_name
            .reference
            .len()
            .checked_add(text.len())
            .ok_or_else(|| invalid("defined-name formula length overflow"))?;
        if new_len > MAX_DEFINED_NAME_FORMULA_BYTES {
            return Err(invalid("workbook defined-name formula limit exceeded"));
        }
        defined_name.reference.push_str(text);
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
        // Producers other than Excel repeat a name within one scope, most
        // often the built-in `_xlnm.*` print and filter names. Excel resolves
        // such a workbook to a single definition rather than refusing to open
        // it, so the first definition wins and later repeats in the same scope
        // are dropped. Distinct scopes remain independent.
        if !self.defined_name_keys.insert(key) {
            return Ok(());
        }
        self.defined_names.push(defined_name);
        Ok(())
    }
}

/// Parse and validate the workbook-level catalog from `workbook.xml` bytes.
pub fn parse_catalog(content: &[u8]) -> Result<Catalog> {
    WorkbookInfo::parse(content).map(|info| Catalog {
        sheets: info.sheets,
        active_sheet_index: info.active_tab.unwrap_or(0),
        uses_1904_date_system: info.uses_1904_date_system,
        defined_names: info.defined_names,
        pivot_caches: info.pivot_caches,
        external_reference_ids: info.external_reference_ids,
    })
}

fn parse_workbook_details(content: &str) -> Result<Catalog> {
    parse_catalog(content.as_bytes())
}

/// Parse workbook metadata, returning sheets, the active sheet index, and the date system.
fn parse_workbook_xml(content: &str) -> Result<(Vec<Sheet>, usize, bool)> {
    parse_workbook_details(content).map(|details| {
        (
            details.sheets,
            details.active_sheet_index,
            details.uses_1904_date_system,
        )
    })
}

/// Parse a standalone `sheet` element.
pub fn parse_sheet(sheet_xml: &str) -> Result<Option<Sheet>> {
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
) -> Result<Sheet> {
    let name = required_string(element, b"name", decoder, "workbook sheet name")?;
    crate::sheet::validate_str(&name)?;
    let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| invalid("workbook sheet is missing relationship ID"))?;
    if relationship_id.is_empty() {
        return Err(invalid("workbook sheet relationship ID cannot be empty"));
    }
    if relationship_id.chars().count() > MAX_RELATIONSHIP_ID_CHARACTERS {
        return Err(invalid(format!(
            "workbook sheet relationship ID exceeds {MAX_RELATIONSHIP_ID_CHARACTERS} characters"
        )));
    }
    let sheet_id = required_u32(element, b"sheetId", decoder, "workbook sheet ID")?;
    if !(1..=MAX_SHEET_ID).contains(&sheet_id) {
        return Err(invalid(format!(
            "workbook sheet ID must be between 1 and {MAX_SHEET_ID}"
        )));
    }
    let visibility = match optional_string(element, b"state", decoder)?.as_deref() {
        None | Some("visible") => Visibility::Visible,
        Some("hidden") => Visibility::Hidden,
        Some("veryHidden") => Visibility::VeryHidden,
        Some(value) => Visibility::Unknown(value.into()),
    };

    Ok(Sheet {
        name,
        relationship_id,
        sheet_id,
        visibility,
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

fn optional_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    Ok(unqualified_attribute_value(element, name, decoder)?)
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

        let sheet = parse_sheet(r#"<sheet name="One &amp; Two" sheetId="2" r:id="r9"/>"#)
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
        assert_eq!(details.defined_names[0].reference, "'A & B'!$A$1:$D$20");
        assert_eq!(details.defined_names[1].reference, "42");
        assert_eq!(details.defined_names[0].name, "_xlnm.Print_Area");
    }

    #[test]
    fn round_trips_defined_name_scope_and_standard_attributes() {
        let xml = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}">
                <sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets>
                <definedNames><definedName name="LocalName" localSheetId="0"
                    comment="A &amp; B" customMenu="Menu" description="Description"
                    help="Help" statusBar="Status" shortcutKey="K" hidden="1"
                    function="true" vbProcedure="1" xlm="true" functionGroupId="7"
                    publishToServer="1" workbookParameter="true">SUM(One!$A$1:$A$2)</definedName></definedNames>
            </workbook>"#
        );
        let parsed = parse_workbook_details(&xml).unwrap();
        let name = &parsed.defined_names[0];
        assert_eq!(name.local_sheet_id, Some(0));
        assert_eq!(name.comment.as_deref(), Some("A & B"));
        assert!(name.hidden && name.function && name.vb_procedure && name.xlm);
        assert!(name.publish_to_server && name.workbook_parameter);
        assert_eq!(name.function_group_id, Some(7));

        assert_eq!(name.custom_menu.as_deref(), Some("Menu"));
        assert_eq!(name.description.as_deref(), Some("Description"));
        assert_eq!(name.help.as_deref(), Some("Help"));
        assert_eq!(name.status_bar.as_deref(), Some("Status"));
        assert_eq!(name.shortcut_key.as_deref(), Some("K"));
        assert_eq!(name.reference, "SUM(One!$A$1:$A$2)");
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
                PivotCache {
                    cache_id: 7,
                    relationship_id: "custom-cache".to_string(),
                },
                PivotCache {
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
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames/><definedNames/></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x" hidden="yes">1</definedName></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x" localSheetId="32767">1</definedName></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="{}">1</definedName></definedNames></workbook>"#,
                "n".repeat(MAX_DEFINED_NAME_CHARACTERS + 1)
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x" comment="{}">1</definedName></definedNames></workbook>"#,
                "c".repeat(MAX_DEFINED_NAME_COMMENT_CHARACTERS + 1)
            ),
        ];
        for xml in invalid {
            assert!(parse_workbook_details(&xml).is_err(), "accepted {xml}");
        }

        let oversized_formula = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x">{}</definedName></definedNames></workbook>"#,
            "1".repeat(MAX_DEFINED_NAME_FORMULA_BYTES + 1)
        );
        assert!(parse_workbook_details(&oversized_formula).is_err());
    }

    /// Producers other than Excel repeat a defined name within one scope.
    /// The first definition wins instead of the workbook being rejected.
    #[test]
    fn keeps_the_first_of_duplicate_defined_names_in_the_same_scope() {
        let sheets = r#"<sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets>"#;
        // Name matching is case-insensitive, so `Same` and `same` collide.
        let xml = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="Same">1</definedName><definedName name="same">2</definedName></definedNames></workbook>"#
        );
        let parsed = parse_workbook_details(&xml).unwrap();
        assert_eq!(parsed.defined_names.len(), 1);
        assert_eq!(parsed.defined_names[0].name, "Same");
        assert_eq!(parsed.defined_names[0].reference, "1");
    }

    /// The same name in two different sheet scopes is not a duplicate.
    #[test]
    fn keeps_same_defined_name_in_distinct_scopes() {
        let sheets = r#"<sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets>"#;
        let xml = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}">{sheets}<definedNames><definedName name="x">1</definedName><definedName name="x" localSheetId="0">2</definedName></definedNames></workbook>"#
        );
        let parsed = parse_workbook_details(&xml).unwrap();
        assert_eq!(parsed.defined_names.len(), 2);
    }

    /// `sheet/@state` values outside `ST_SheetState` must not fail the
    /// workbook; LibreOffice writes `state="show"` (tdf#118668).
    #[test]
    fn tolerates_unrecognised_sheet_state_values() {
        for state in ["show", "", "VISIBLE", "bogus"] {
            let xml = format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" state="{state}" r:id="r1"/></sheets></workbook>"#
            );
            let parsed = parse_workbook_details(&xml).unwrap_or_else(|e| panic!("{state}: {e}"));
            assert_eq!(parsed.sheets.len(), 1, "{state}");
            assert_eq!(parsed.sheets[0].name, "One", "{state}");
            assert_eq!(
                parsed.sheets[0].visibility,
                Visibility::Unknown(state.into()),
                "{state}"
            );
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

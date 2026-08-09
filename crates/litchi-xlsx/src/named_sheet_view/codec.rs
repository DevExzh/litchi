//! Bounded `SpreadsheetML` named-sheet-view XML codec.
//!
//! MCE preprocessing and retained extension fragments stay inert; this layer
//! owns only XML validation, conversion, and serialization.

use crate::auto_filter::{Column, Definition, parse_auto_filter, write_auto_filter_fragment};
use crate::error::Result;
use crate::sort::{SortBy, SortMethod};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_sheet::Cell as Address;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};
use std::borrow::Cow;
use std::collections::HashSet;

use super::model::{
    ColumnFilter, DifferentialFormat, Extension, Filter, Guid, IconSet, Markup, Range,
    SortCondition, SortConditionKind, SortRule, SortRules, View, Views,
};
use super::{
    CORE, FRAGMENT_FILTER_ID, FRAGMENT_VIEW_ID, MAX_COLUMNS, MAX_EXTENSIONS, MAX_FILTERS,
    MAX_FRAGMENT_DEPTH, MAX_FRAGMENT_NODES, MAX_MARKUP_BYTES, MAX_NAMESPACE_DECLARATIONS,
    MAX_PART_BYTES, MAX_RETAINED_BYTES, MAX_VIEWS, NSV, RICH, STRICT, X14, invalid, xml_error,
};

pub fn parse_named_sheet_views(xml: &[u8]) -> Result<Views> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("Named Sheet Views part exceeds size limit"));
    }
    let mut caps = Capabilities::default();
    for ns in [NSV, X14, RICH] {
        caps.understand_namespace(String::from_utf8_lossy(ns).into_owned());
    }
    let limits = Limits {
        max_input_bytes: MAX_PART_BYTES,
        max_output_bytes: MAX_PART_BYTES * 2,
        ..Limits::default()
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?;
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut parser = Parser::default();
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        if matches!(event, Event::DocType(_) | Event::PI(_)) {
            return Err(invalid("DTD and processing instructions are rejected"));
        }
        if let Event::GeneralRef(v) = &event {
            let n = v.decode().map_err(xml_error)?;
            if !matches!(n.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !n.starts_with('#') {
                return Err(invalid("custom XML entities are rejected"));
            }
        }
        if parser.capture.is_some() {
            parser.capture_event(event)?;
            continue;
        }
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) => parser.start(&namespace, &e, decoder)?,
            Event::Empty(e) => parser.empty(&namespace, &e, decoder)?,
            Event::End(e) => parser.end(&namespace, e.local_name().as_ref())?,
            Event::Text(e) => {
                let text = e.decode().map_err(xml_error)?;
                if !text.chars().all(|c| matches!(c, ' ' | '\t' | '\r' | '\n')) {
                    return Err(invalid(
                        "text is not allowed in Named Sheet Views containers",
                    ));
                }
            },
            Event::CData(_) => {
                return Err(invalid(
                    "CDATA is not allowed in Named Sheet Views containers",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if parser.capture.is_some() || !parser.stack.is_empty() {
        return Err(invalid("unterminated Named Sheet Views XML"));
    }
    parser.finish()
}

pub fn write_named_sheet_views(value: &Views) -> Result<Vec<u8>> {
    validate_view_collection(&value.views)?;

    let mut output = Vec::new();
    output.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    output.extend_from_slice(b"<namedSheetViews");
    write_xml_attribute(&mut output, "xmlns", std::str::from_utf8(NSV).unwrap());

    let needs_spreadsheet_prefix = has_filter_payload(value);
    let spreadsheet_prefix_needs_local_binding =
        needs_spreadsheet_prefix && needs_local_spreadsheet_prefix_binding(value);
    let mut declared = HashSet::new();
    declared.insert("xmlns".to_string());
    for (name, namespace) in &value.namespace_declarations {
        if name == "xmlns" || !name.starts_with("xmlns:") || !declared.insert(name.clone()) {
            continue;
        }
        write_xml_attribute(&mut output, name, namespace);
    }
    if needs_spreadsheet_prefix && !declared.contains("xmlns:x") {
        write_xml_attribute(&mut output, "xmlns:x", std::str::from_utf8(CORE).unwrap());
    }
    output.push(b'>');

    for view in &value.views {
        write_named_sheet_view(&mut output, view, spreadsheet_prefix_needs_local_binding)?;
    }
    write_extensions(&mut output, &value.extensions);
    output.extend_from_slice(b"</namedSheetViews>");

    if output.len() > MAX_PART_BYTES {
        return Err(invalid(
            "serialized Named Sheet Views part exceeds size limit",
        ));
    }
    parse_named_sheet_views(&output)?;
    Ok(output)
}

fn has_filter_payload(value: &Views) -> bool {
    value.views.iter().any(view_has_filter_payload)
}

pub(crate) fn view_has_filter_payload(view: &View) -> bool {
    view.filters.iter().any(|filter| {
        filter
            .column_filters
            .iter()
            .any(|column| column.filters.iter().any(|filter| filter.payload.is_some()))
    })
}

fn needs_local_spreadsheet_prefix_binding(value: &Views) -> bool {
    value
        .namespace_declarations
        .iter()
        .find(|(name, _)| name == "xmlns:x")
        .is_some_and(|(_, namespace)| {
            namespace.as_bytes() != CORE && namespace.as_bytes() != STRICT
        })
}

fn write_named_sheet_view(
    output: &mut Vec<u8>,
    view: &View,
    spreadsheet_prefix_needs_local_binding: bool,
) -> Result<()> {
    validate_name(&view.name)?;
    output.extend_from_slice(b"<namedSheetView");
    write_xml_attribute(output, "name", &view.name);
    write_xml_attribute(output, "id", view.id.as_str());
    if view.filters.is_empty() && view.extensions.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for filter in &view.filters {
        write_named_sheet_view_filter(output, filter, spreadsheet_prefix_needs_local_binding)?;
    }
    write_extensions(output, &view.extensions);
    output.extend_from_slice(b"</namedSheetView>");
    Ok(())
}

fn write_named_sheet_view_filter(
    output: &mut Vec<u8>,
    filter: &Filter,
    spreadsheet_prefix_needs_local_binding: bool,
) -> Result<()> {
    output.extend_from_slice(b"<nsvFilter");
    write_xml_attribute(output, "filterId", filter.filter_id.as_str());
    if let Some(reference) = &filter.reference {
        write_xml_attribute(output, "ref", reference.as_str());
    }
    if let Some(table_id) = filter.table_id {
        write_xml_attribute(output, "tableId", &table_id.to_string());
    }
    if filter.column_filters.is_empty()
        && filter.sort_rules.is_none()
        && filter.extensions.is_empty()
    {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for column in &filter.column_filters {
        write_column_filter(output, column, spreadsheet_prefix_needs_local_binding)?;
    }
    if let Some(sort_rules) = &filter.sort_rules {
        write_sort_rules(output, sort_rules)?;
    }
    write_extensions(output, &filter.extensions);
    output.extend_from_slice(b"</nsvFilter>");
    Ok(())
}

fn write_column_filter(
    output: &mut Vec<u8>,
    column: &ColumnFilter,
    spreadsheet_prefix_needs_local_binding: bool,
) -> Result<()> {
    if column.column_id >= MAX_COLUMNS as u32 {
        return Err(invalid(
            "named-sheet-view colId exceeds worksheet column limit",
        ));
    }
    output.extend_from_slice(b"<columnFilter");
    write_xml_attribute(output, "colId", &column.column_id.to_string());
    if let Some(id) = &column.id {
        write_xml_attribute(output, "id", id.as_str());
    }
    if column.differential_format.is_none()
        && column.filters.is_empty()
        && column.extensions.is_empty()
    {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    if let Some(differential_format) = &column.differential_format {
        output.extend_from_slice(differential_format.xml());
    }
    for filter in &column.filters {
        write_filter_payload(output, filter, spreadsheet_prefix_needs_local_binding)?;
    }
    write_extensions(output, &column.extensions);
    output.extend_from_slice(b"</columnFilter>");
    Ok(())
}

fn write_filter_payload(
    output: &mut Vec<u8>,
    filter: &Column,
    spreadsheet_prefix_needs_local_binding: bool,
) -> Result<()> {
    output.extend_from_slice(b"<filter");
    write_xml_attribute(output, "colId", &filter.column_id.to_string());
    if filter.hidden_button {
        write_xml_attribute(output, "hiddenButton", "1");
    }
    if !filter.show_button {
        write_xml_attribute(output, "showButton", "0");
    }
    let Some(_) = &filter.payload else {
        output.extend_from_slice(b"/>");
        return Ok(());
    };
    if spreadsheet_prefix_needs_local_binding {
        write_xml_attribute(output, "xmlns:x", std::str::from_utf8(CORE).unwrap());
    }
    output.push(b'>');
    output.extend_from_slice(&filter_payload_markup(filter)?);
    output.extend_from_slice(b"</filter>");
    Ok(())
}

pub(crate) fn filter_payload_markup(filter: &Column) -> Result<Vec<u8>> {
    let fragment = write_auto_filter_fragment(&Definition {
        reference: None,
        columns: vec![filter.clone()],
        sort_state: None,
        opaque: None,
    })?;
    let mut reader = NsReader::from_reader(fragment.as_slice());
    let mut writer = Writer::new(Vec::new());
    let mut in_column = false;
    let mut depth = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        match event {
            Event::Start(element) if element.local_name().as_ref() == b"filterColumn" => {
                if in_column {
                    return Err(invalid("nested generated filterColumn"));
                }
                in_column = true;
                depth = 1;
            },
            Event::Start(element) if in_column => {
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                depth += 1;
            },
            Event::Empty(element) if in_column => {
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
            },
            Event::End(element) if in_column => {
                if depth == 1 && element.local_name().as_ref() == b"filterColumn" {
                    in_column = false;
                    depth = 0;
                } else {
                    writer.write_event(Event::End(element)).map_err(xml_error)?;
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("generated filter depth underflow"))?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if in_column || depth != 0 {
        return Err(invalid("unterminated generated filterColumn"));
    }
    let markup = writer.into_inner();
    if markup.is_empty() {
        return Err(invalid("generated Named Sheet Views filter has no payload"));
    }
    Ok(markup)
}

fn write_sort_rules(output: &mut Vec<u8>, rules: &SortRules) -> Result<()> {
    output.extend_from_slice(b"<sortRules");
    if rules.sort_method != SortMethod::None {
        write_xml_attribute(output, "sortMethod", rules.sort_method.as_str());
    }
    if rules.case_sensitive {
        write_xml_attribute(output, "caseSensitive", "1");
    }
    if rules.rules.is_empty() && rules.extensions.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for rule in &rules.rules {
        write_sort_rule(output, rule)?;
    }
    write_extensions(output, &rules.extensions);
    output.extend_from_slice(b"</sortRules>");
    Ok(())
}

fn write_sort_rule(output: &mut Vec<u8>, rule: &SortRule) -> Result<()> {
    if rule.column_id >= MAX_COLUMNS as u32 {
        return Err(invalid(
            "named-sheet-view colId exceeds worksheet column limit",
        ));
    }
    output.extend_from_slice(b"<sortRule");
    write_xml_attribute(output, "colId", &rule.column_id.to_string());
    if let Some(id) = &rule.id {
        write_xml_attribute(output, "id", id.as_str());
    }
    if rule.differential_format.is_none() && rule.condition.is_none() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    if let Some(differential_format) = &rule.differential_format {
        output.extend_from_slice(differential_format.xml());
    }
    if let Some(condition) = &rule.condition {
        write_sort_condition(output, condition)?;
    }
    output.extend_from_slice(b"</sortRule>");
    Ok(())
}

fn write_sort_condition(output: &mut Vec<u8>, condition: &SortCondition) -> Result<()> {
    let name = match condition.kind {
        SortConditionKind::Standard => "sortCondition",
        SortConditionKind::RichValue => "richSortCondition",
    };
    output.push(b'<');
    output.extend_from_slice(name.as_bytes());
    write_xml_attribute(output, "ref", condition.reference.as_str());
    if condition.descending {
        write_xml_attribute(output, "descending", "1");
    }
    if condition.sort_by != SortBy::Value {
        write_xml_attribute(output, "sortBy", condition.sort_by.as_str());
    }
    if let Some(custom_list) = &condition.custom_list {
        write_xml_attribute(output, "customList", custom_list);
    }
    if let Some(differential_format_id) = condition.differential_format_id {
        write_xml_attribute(output, "dxfId", &differential_format_id.to_string());
    }
    if let Some(icon_set) = condition.icon_set {
        write_xml_attribute(output, "iconSet", icon_set.as_str());
    }
    if let Some(icon_id) = condition.icon_id {
        write_xml_attribute(output, "iconId", &icon_id.to_string());
    }
    if let Some(rich_sort_key) = &condition.rich_sort_key {
        write_xml_attribute(output, "richSortKey", rich_sort_key);
    }
    output.extend_from_slice(b"/>");
    Ok(())
}

fn write_extensions(output: &mut Vec<u8>, extensions: &[Extension]) {
    if extensions.is_empty() {
        return;
    }
    output.extend_from_slice(b"<extLst>");
    for extension in extensions {
        output.extend_from_slice(extension.markup.xml());
    }
    output.extend_from_slice(b"</extLst>");
}

fn write_xml_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#9;"),
            '\n' => output.extend_from_slice(b"&#10;"),
            '\r' => output.extend_from_slice(b"&#13;"),
            _ => {
                let mut buffer = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
    output.push(b'"');
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Ctx {
    Outside,
    Root,
    View,
    Filter,
    Column,
    SortRules,
    SortRule,
    Leaf,
    ExtList(ExtOwner),
    Captured,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum ExtOwner {
    Root,
    View,
    Filter,
    Column,
    SortRules,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum DxfOwner {
    Column,
    SortRule,
}
enum Payload {
    Filter,
    Extension(ExtOwner, Extension),
    Dxf(DxfOwner),
}
struct Capture {
    depth: usize,
    writer: Writer<Vec<u8>>,
    payload: Payload,
}
#[derive(Default)]
struct Parser {
    stack: Vec<Ctx>,
    root: Option<Views>,
    view: Option<ViewBuilder>,
    filter: Option<FilterBuilder>,
    column: Option<ColumnBuilder>,
    sort_rules: Option<SortRulesBuilder>,
    sort_rule: Option<SortRuleBuilder>,
    capture: Option<Capture>,
    root_phase: u8,
    retained: usize,
    seen_root: bool,
}
struct ViewBuilder {
    value: View,
    phase: u8,
}
struct FilterBuilder {
    value: Filter,
    phase: u8,
}
struct ColumnBuilder {
    value: ColumnFilter,
    phase: u8,
}
struct SortRulesBuilder {
    value: SortRules,
    phase: u8,
}
struct SortRuleBuilder {
    value: SortRule,
    phase: u8,
}

impl Parser {
    fn parent(&self) -> Ctx {
        self.stack.last().copied().unwrap_or(Ctx::Outside)
    }
    fn start(&mut self, ns: &ResolveResult<'_>, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let local = e.local_name();
        let nsv = exact(ns, NSV);
        match (self.parent(), nsv, local.as_ref()) {
            (Ctx::Outside, true, b"namedSheetViews") => {
                self.begin_root(e, d)?;
                self.stack.push(Ctx::Root);
            },
            (Ctx::Root, true, b"namedSheetView") => {
                self.begin_view(e, d)?;
                self.stack.push(Ctx::View);
            },
            (Ctx::View, true, b"nsvFilter") => {
                self.begin_filter(e, d)?;
                self.stack.push(Ctx::Filter);
            },
            (Ctx::Filter, true, b"columnFilter") => {
                self.begin_column(e, d)?;
                self.stack.push(Ctx::Column);
            },
            (Ctx::Filter, true, b"sortRules") => {
                self.begin_sort_rules(e, d)?;
                self.stack.push(Ctx::SortRules);
            },
            (Ctx::SortRules, true, b"sortRule") => {
                self.begin_sort_rule(e, d)?;
                self.stack.push(Ctx::SortRule);
            },
            (Ctx::SortRule, true, b"sortCondition") => {
                self.add_condition(e, d, SortConditionKind::Standard)?;
                self.stack.push(Ctx::Leaf);
            },
            (Ctx::SortRule, true, b"richSortCondition") => {
                self.add_condition(e, d, SortConditionKind::RichValue)?;
                self.stack.push(Ctx::Leaf);
            },
            (parent, true, b"extLst") => {
                let owner = Self::ext_owner(parent)?;
                self.begin_ext(owner)?;
                self.stack.push(Ctx::ExtList(owner));
            },
            (Ctx::Column, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::Column)?;

                self.begin_capture(Payload::Dxf(DxfOwner::Column), e)?;
                self.stack.push(Ctx::Captured);
            },
            (Ctx::SortRule, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::SortRule)?;
                self.begin_capture(Payload::Dxf(DxfOwner::SortRule), e)?;
                self.stack.push(Ctx::Captured);
            },
            (Ctx::Column, true, b"filter") => {
                self.begin_capture(Payload::Filter, e)?;
                self.stack.push(Ctx::Captured);
            },
            (Ctx::ExtList(owner), _, b"ext") if core(ns) => {
                let x = parse_extension(e, d)?;
                self.begin_capture(Payload::Extension(owner, x), e)?;
                self.stack.push(Ctx::Captured);
            },
            (Ctx::Leaf, _, _) => {
                return Err(invalid(
                    "Named Sheet Views leaf element contains child content",
                ));
            },
            (parent, _, _) if inside(parent) => {
                return Err(invalid(format!(
                    "unexpected Named Sheet Views element '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
            (Ctx::Outside, _, _) => {
                return Err(invalid(
                    "Named Sheet Views part has an invalid or additional root element",
                ));
            },
            _ => self.stack.push(Ctx::Outside),
        }
        Ok(())
    }
    fn empty(&mut self, ns: &ResolveResult<'_>, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let local = e.local_name();
        let nsv = exact(ns, NSV);
        match (self.parent(), nsv, local.as_ref()) {
            (Ctx::Outside, true, b"namedSheetViews") => {
                return Err(invalid(
                    "Named Sheet Views part must contain a namedSheetView",
                ));
            },
            (Ctx::Root, true, b"namedSheetView") => {
                self.begin_view(e, d)?;
                self.finish_view()?;
            },
            (Ctx::View, true, b"nsvFilter") => {
                self.begin_filter(e, d)?;
                self.finish_filter()?;
            },
            (Ctx::Filter, true, b"columnFilter") => {
                self.begin_column(e, d)?;
                self.finish_column()?;
            },
            (Ctx::Filter, true, b"sortRules") => {
                self.begin_sort_rules(e, d)?;
                self.finish_sort_rules()?;
            },
            (Ctx::SortRules, true, b"sortRule") => {
                self.begin_sort_rule(e, d)?;
                self.finish_sort_rule()?;
            },
            (Ctx::SortRule, true, b"sortCondition") => {
                self.add_condition(e, d, SortConditionKind::Standard)?;
            },
            (Ctx::SortRule, true, b"richSortCondition") => {
                self.add_condition(e, d, SortConditionKind::RichValue)?;
            },
            (parent, true, b"extLst") => {
                let owner = Self::ext_owner(parent)?;
                self.begin_ext(owner)?;
            },
            (Ctx::Column, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::Column)?;
                self.empty_markup(Payload::Dxf(DxfOwner::Column), e)?;
            },
            (Ctx::SortRule, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::SortRule)?;
                self.empty_markup(Payload::Dxf(DxfOwner::SortRule), e)?;
            },
            (Ctx::Column, true, b"filter") => self.empty_markup(Payload::Filter, e)?,
            (Ctx::ExtList(owner), _, b"ext") if core(ns) => {
                let x = parse_extension(e, d)?;
                self.empty_markup(Payload::Extension(owner, x), e)?;
            },
            (Ctx::Leaf, _, _) => {
                return Err(invalid(
                    "Named Sheet Views leaf element contains child content",
                ));
            },
            (parent, _, _) if inside(parent) => {
                return Err(invalid(format!(
                    "unexpected Named Sheet Views element '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
            (Ctx::Outside, _, _) => {
                return Err(invalid(
                    "Named Sheet Views part has an invalid or additional root element",
                ));
            },
            _ => {},
        }
        Ok(())
    }
    fn end(&mut self, ns: &ResolveResult<'_>, name: &[u8]) -> Result<()> {
        let ctx = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unexpected Named Sheet Views end element"))?;
        match ctx {
            Ctx::Root if exact(ns, NSV) && name == b"namedSheetViews" => {},
            Ctx::View if exact(ns, NSV) && name == b"namedSheetView" => self.finish_view()?,
            Ctx::Filter if exact(ns, NSV) && name == b"nsvFilter" => self.finish_filter()?,
            Ctx::Column if exact(ns, NSV) && name == b"columnFilter" => self.finish_column()?,
            Ctx::SortRules if exact(ns, NSV) && name == b"sortRules" => self.finish_sort_rules()?,
            Ctx::SortRule if exact(ns, NSV) && name == b"sortRule" => self.finish_sort_rule()?,
            Ctx::Outside | Ctx::Leaf | Ctx::ExtList(_) => {},
            Ctx::Captured => return Err(invalid("invalid retained markup state")),
            _ => return Err(invalid("mismatched Named Sheet Views end element")),
        }
        Ok(())
    }
    fn begin_root(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        if self.seen_root {
            return Err(invalid("duplicate namedSheetViews root"));
        }
        self.seen_root = true;
        self.root_phase = 0;
        self.root = Some(Views {
            views: Vec::new(),
            extensions: Vec::new(),
            namespace_declarations: parse_namespace_declarations(e, d)?,
        });
        Ok(())
    }
    fn begin_view(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        if self.root_phase > 0 {
            return Err(invalid("namedSheetView appears after root extLst"));
        }
        if self
            .root
            .as_ref()
            .ok_or_else(|| invalid("namedSheetView outside root"))?
            .views
            .len()
            >= MAX_VIEWS
        {
            return Err(invalid("too many named sheet views"));
        }
        let name = required_attr(e, b"name", d)?;
        validate_name(&name)?;
        let id = parse_guid(&required_attr(e, b"id", d)?)?;
        self.view = Some(ViewBuilder {
            value: View {
                name,
                id,
                filters: Vec::new(),
                extensions: Vec::new(),
            },
            phase: 0,
        });
        Ok(())
    }
    fn finish_view(&mut self) -> Result<()> {
        let value = self
            .view
            .take()
            .ok_or_else(|| invalid("namedSheetView end without start"))?
            .value;
        self.root
            .as_mut()
            .ok_or_else(|| invalid("namedSheetView outside root"))?
            .views
            .push(value);
        Ok(())
    }
    fn begin_filter(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let view = self
            .view
            .as_mut()
            .ok_or_else(|| invalid("nsvFilter outside namedSheetView"))?;
        if view.phase > 0 {
            return Err(invalid("nsvFilter appears after extLst"));
        }
        if view.value.filters.len() >= MAX_FILTERS {
            return Err(invalid("too many named-sheet-view filters"));
        }
        self.filter = Some(FilterBuilder {
            value: Filter {
                filter_id: parse_guid(&required_attr(e, b"filterId", d)?)?,
                reference: attr(e, b"ref", d)?.map(|v| parse_range(&v)).transpose()?,
                table_id: optional_u32(e, b"tableId", d)?,
                column_filters: Vec::new(),
                sort_rules: None,
                extensions: Vec::new(),
            },
            phase: 0,
        });
        Ok(())
    }
    fn finish_filter(&mut self) -> Result<()> {
        let value = self
            .filter
            .take()
            .ok_or_else(|| invalid("nsvFilter end without start"))?
            .value;
        self.view
            .as_mut()
            .ok_or_else(|| invalid("nsvFilter outside view"))?
            .value
            .filters
            .push(value);
        Ok(())
    }
    fn begin_column(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let filter = self
            .filter
            .as_mut()
            .ok_or_else(|| invalid("columnFilter outside nsvFilter"))?;
        if filter.phase > 0 {
            return Err(invalid("columnFilter is out of schema order"));
        }
        if filter.value.column_filters.len() >= MAX_COLUMNS {
            return Err(invalid("too many named-sheet-view column filters"));
        }
        self.column = Some(ColumnBuilder {
            value: ColumnFilter {
                column_id: column_id(e, d)?,
                id: optional_guid(e, b"id", d)?,
                differential_format: None,
                filters: Vec::new(),
                extensions: Vec::new(),
            },
            phase: 0,
        });
        Ok(())
    }
    fn finish_column(&mut self) -> Result<()> {
        let value = self
            .column
            .take()
            .ok_or_else(|| invalid("columnFilter end without start"))?
            .value;
        self.filter
            .as_mut()
            .ok_or_else(|| invalid("columnFilter outside filter"))?
            .value
            .column_filters
            .push(value);
        Ok(())
    }
    fn begin_sort_rules(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let filter = self
            .filter
            .as_mut()
            .ok_or_else(|| invalid("sortRules outside nsvFilter"))?;
        if filter.phase > 1 || filter.value.sort_rules.is_some() {
            return Err(invalid("duplicate or misplaced sortRules"));
        }
        filter.phase = 1;
        let method = attr(e, b"sortMethod", d)?
            .as_deref()
            .unwrap_or("none")
            .parse::<SortMethod>()
            .map_err(|error| invalid(error.to_string()))?;
        self.sort_rules = Some(SortRulesBuilder {
            value: SortRules {
                sort_method: method,
                case_sensitive: bool_attr(e, b"caseSensitive", d)?.unwrap_or(false),
                rules: Vec::new(),
                extensions: Vec::new(),
            },
            phase: 0,
        });
        Ok(())
    }
    fn finish_sort_rules(&mut self) -> Result<()> {
        let value = self
            .sort_rules
            .take()
            .ok_or_else(|| invalid("sortRules end without start"))?
            .value;
        self.filter
            .as_mut()
            .ok_or_else(|| invalid("sortRules outside filter"))?
            .value
            .sort_rules = Some(value);
        Ok(())
    }
    fn begin_sort_rule(&mut self, e: &BytesStart<'_>, d: Decoder) -> Result<()> {
        let rules = self
            .sort_rules
            .as_mut()
            .ok_or_else(|| invalid("sortRule outside sortRules"))?;
        if rules.phase > 0 {
            return Err(invalid("sortRule appears after extLst"));
        }
        if rules.value.rules.len() >= 64 {
            return Err(invalid("sortRules exceeds 64 rules"));
        }
        self.sort_rule = Some(SortRuleBuilder {
            value: SortRule {
                column_id: column_id(e, d)?,
                id: optional_guid(e, b"id", d)?,
                differential_format: None,
                condition: None,
            },
            phase: 0,
        });
        Ok(())
    }
    fn finish_sort_rule(&mut self) -> Result<()> {
        let value = self
            .sort_rule
            .take()
            .ok_or_else(|| invalid("sortRule end without start"))?
            .value;
        if value
            .condition
            .as_ref()
            .is_some_and(|c| c.differential_format_id.is_some())
            != value.differential_format.is_some()
        {
            return Err(invalid(
                "sortRule dxf presence does not match sortCondition dxfId",
            ));
        }
        self.sort_rules
            .as_mut()
            .ok_or_else(|| invalid("sortRule outside sortRules"))?
            .value
            .rules
            .push(value);
        Ok(())
    }
    fn add_condition(
        &mut self,
        e: &BytesStart<'_>,
        d: Decoder,
        kind: SortConditionKind,
    ) -> Result<()> {
        let rule = self
            .sort_rule
            .as_mut()
            .ok_or_else(|| invalid("sort condition outside sortRule"))?;
        if rule.phase > 1 || rule.value.condition.is_some() {
            return Err(invalid("sortRule has multiple conditions"));
        }
        rule.phase = 1;
        rule.value.condition = Some(parse_condition(e, d, kind)?);
        Ok(())
    }
    fn ext_owner(parent: Ctx) -> Result<ExtOwner> {
        match parent {
            Ctx::Root => Ok(ExtOwner::Root),
            Ctx::View => Ok(ExtOwner::View),
            Ctx::Filter => Ok(ExtOwner::Filter),
            Ctx::Column => Ok(ExtOwner::Column),
            Ctx::SortRules => Ok(ExtOwner::SortRules),
            _ => Err(invalid("extLst is not allowed here")),
        }
    }
    fn begin_ext(&mut self, owner: ExtOwner) -> Result<()> {
        match owner {
            ExtOwner::Root => {
                if self.root_phase > 0 {
                    return Err(invalid("duplicate root extLst"));
                }
                self.root_phase = 1;
            },
            ExtOwner::View => {
                let v = self.view.as_mut().unwrap();
                if v.phase > 0 {
                    return Err(invalid("duplicate view extLst"));
                }
                v.phase = 1;
            },
            ExtOwner::Filter => {
                let v = self.filter.as_mut().unwrap();
                if v.phase > 1 {
                    return Err(invalid("duplicate filter extLst"));
                }
                v.phase = 2;
            },
            ExtOwner::Column => {
                let v = self.column.as_mut().unwrap();
                if v.phase > 1 {
                    return Err(invalid("duplicate columnFilter extLst"));
                }
                v.phase = 2;
            },
            ExtOwner::SortRules => {
                let v = self.sort_rules.as_mut().unwrap();
                if v.phase > 0 {
                    return Err(invalid("duplicate sortRules extLst"));
                }
                v.phase = 1;
            },
        }
        Ok(())
    }
    fn prepare_dxf(&mut self, owner: DxfOwner) -> Result<()> {
        match owner {
            DxfOwner::Column => {
                let v = self.column.as_mut().unwrap();
                if v.phase > 0 || v.value.differential_format.is_some() {
                    return Err(invalid("duplicate or misplaced columnFilter dxf"));
                }
                v.value.differential_format = Some(Markup(Vec::new()));
            },
            DxfOwner::SortRule => {
                let v = self.sort_rule.as_mut().unwrap();
                if v.phase > 0 || v.value.differential_format.is_some() {
                    return Err(invalid("duplicate or misplaced sortRule dxf"));
                }
                v.value.differential_format = Some(Markup(Vec::new()));
            },
        }
        Ok(())
    }
    fn begin_capture(&mut self, payload: Payload, e: &BytesStart<'_>) -> Result<()> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Start(e.clone()))
            .map_err(xml_error)?;
        self.capture = Some(Capture {
            depth: 1,
            writer,
            payload,
        });
        Ok(())
    }
    fn empty_markup(&mut self, payload: Payload, e: &BytesStart<'_>) -> Result<()> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Empty(e.clone()))
            .map_err(xml_error)?;
        self.attach(payload, writer.into_inner())
    }
    fn capture_event(&mut self, event: Event<'static>) -> Result<()> {
        let capture = self.capture.as_mut().unwrap();
        capture
            .writer
            .write_event(event.clone())
            .map_err(xml_error)?;
        match event {
            Event::Start(_) => capture.depth += 1,
            Event::End(_) => capture.depth -= 1,
            Event::Eof => return Err(invalid("unterminated retained Named Sheet Views markup")),
            _ => {},
        }
        if capture.depth == 0 {
            let capture = self.capture.take().unwrap();
            self.stack.pop();
            self.attach(capture.payload, capture.writer.into_inner())?;
        }
        Ok(())
    }
    fn attach(&mut self, payload: Payload, markup: Vec<u8>) -> Result<()> {
        if markup.len() > MAX_MARKUP_BYTES {
            return Err(invalid(
                "retained Named Sheet Views element exceeds size limit",
            ));
        }
        self.retained = self
            .retained
            .checked_add(markup.len())
            .ok_or_else(|| invalid("retained-byte overflow"))?;
        if self.retained > MAX_RETAINED_BYTES {
            return Err(invalid(
                "retained Named Sheet Views markup exceeds resource limit",
            ));
        }
        match payload {
            Payload::Filter => {
                let parsed = parse_filter_payload(&markup)?;
                let column = self
                    .column
                    .as_mut()
                    .ok_or_else(|| invalid("filter outside columnFilter"))?;
                if column.phase > 1 {
                    return Err(invalid("filter appears after extLst"));
                }
                column.phase = 1;
                if column.value.filters.len() >= MAX_FILTERS {
                    return Err(invalid("too many filter payloads"));
                }
                column.value.filters.push(parsed);
            },
            Payload::Dxf(DxfOwner::Column) => {
                self.column.as_mut().unwrap().value.differential_format = Some(Markup(markup));
            },
            Payload::Dxf(DxfOwner::SortRule) => {
                self.sort_rule.as_mut().unwrap().value.differential_format = Some(Markup(markup));
            },
            Payload::Extension(owner, mut x) => {
                x.markup = Markup(markup);
                match owner {
                    ExtOwner::Root => self.root.as_mut().unwrap().extensions.push(x),
                    ExtOwner::View => self.view.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::Filter => self.filter.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::Column => self.column.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::SortRules => {
                        self.sort_rules.as_mut().unwrap().value.extensions.push(x);
                    },
                }
            },
        }
        Ok(())
    }
    fn finish(self) -> Result<Views> {
        let root = self
            .root
            .ok_or_else(|| invalid("missing namedSheetViews root"))?;
        validate_view_collection(&root.views)?;
        Ok(root)
    }
}

fn parse_filter_payload(markup: &[u8]) -> Result<Column> {
    let mut reader = NsReader::from_reader(markup);
    let mut writer = Writer::new(Vec::new());
    let mut worksheet = BytesStart::new("x:worksheet");
    worksheet.push_attribute((
        "xmlns:x",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ));
    worksheet.push_attribute(("xmlns:s", "http://purl.oclc.org/ooxml/spreadsheetml/main"));
    worksheet.push_attribute((
        "xmlns:x14",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    ));
    writer
        .write_event(Event::Start(worksheet))
        .map_err(xml_error)?;
    writer
        .write_event(Event::Start(BytesStart::new("x:autoFilter")))
        .map_err(xml_error)?;
    let mut depth = 0usize;
    let mut seen = false;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        match event {
            Event::Start(e) if !seen => {
                seen = true;
                depth = 1;
                let c = adapt_filter_root(&e, decoder)?;
                writer.write_event(Event::Start(c)).map_err(xml_error)?;
            },
            Event::Empty(e) if !seen => {
                seen = true;
                let c = adapt_filter_root(&e, decoder)?;
                writer.write_event(Event::Empty(c)).map_err(xml_error)?;
                break;
            },
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e)).map_err(xml_error)?;
            },
            Event::Empty(e) => writer.write_event(Event::Empty(e)).map_err(xml_error)?,
            Event::End(e) => {
                depth -= 1;
                if depth == 0 {
                    writer
                        .write_event(Event::End(BytesEnd::new("x:filterColumn")))
                        .map_err(xml_error)?;
                    break;
                }
                writer.write_event(Event::End(e)).map_err(xml_error)?;
            },
            Event::Eof => break,
            other => writer.write_event(other).map_err(xml_error)?,
        }
    }
    if !seen {
        return Err(invalid("empty named-sheet-view filter payload"));
    }
    writer
        .write_event(Event::End(BytesEnd::new("x:autoFilter")))
        .map_err(xml_error)?;
    writer
        .write_event(Event::End(BytesEnd::new("x:worksheet")))
        .map_err(xml_error)?;
    let parsed = parse_auto_filter(&writer.into_inner())?
        .ok_or_else(|| invalid("failed to parse named-sheet-view filter payload"))?;
    if parsed.columns().len() != 1 {
        return Err(invalid(
            "named-sheet-view filter payload did not produce one column",
        ));
    }
    Ok(parsed.columns()[0].clone())
}
fn adapt_filter_root(e: &BytesStart<'_>, d: Decoder) -> Result<BytesStart<'static>> {
    let mut c = BytesStart::new("x:filterColumn");
    for name in [b"colId".as_slice(), b"hiddenButton", b"showButton"] {
        if let Some(value) = attr(e, name, d)? {
            c.push_attribute((std::str::from_utf8(name).unwrap(), Cow::Owned(value)));
        }
    }
    Ok(c)
}
fn parse_condition(
    e: &BytesStart<'_>,
    d: Decoder,
    kind: SortConditionKind,
) -> Result<SortCondition> {
    let reference = parse_range(&required_attr(e, b"ref", d)?)?;
    let sort_by = attr(e, b"sortBy", d)?
        .as_deref()
        .unwrap_or("value")
        .parse::<SortBy>()
        .map_err(|error| invalid(error.to_string()))?;
    let dxf = optional_u32(e, b"dxfId", d)?;
    let icon_set = attr(e, b"iconSet", d)?
        .map(|v| IconSet::parse(&v))
        .transpose()?;
    let icon_id = optional_u32(e, b"iconId", d)?;
    if sort_by == SortBy::Icon {
        if dxf.is_some() {
            return Err(invalid("icon sort condition cannot have dxfId"));
        }
        if let (Some(set), Some(id)) = (icon_set, icon_id)
            && set.cardinality().is_some_and(|n| id >= n)
        {
            return Err(invalid("iconId is outside icon set"));
        }
    } else if icon_set.is_some() || icon_id.is_some() {
        return Err(invalid(
            "non-icon sort condition cannot have icon attributes",
        ));
    }
    if !matches!(sort_by, SortBy::CellColor | SortBy::FontColor) && dxf.is_some() {
        return Err(invalid("dxfId is only valid for color sorting"));
    }
    let rich_sort_key = if kind == SortConditionKind::RichValue {
        attr(e, b"richSortKey", d)?
    } else {
        if attr(e, b"richSortKey", d)?.is_some() {
            return Err(invalid("standard sort condition cannot have richSortKey"));
        }
        None
    };
    if rich_sort_key
        .as_ref()
        .is_some_and(|v| v.chars().count() > 255)
    {
        return Err(invalid("richSortKey exceeds 255 characters"));
    }
    Ok(SortCondition {
        kind,
        reference,
        descending: bool_attr(e, b"descending", d)?.unwrap_or(false),
        sort_by,
        custom_list: bounded_attr(e, b"customList", d, 32767)?,
        differential_format_id: dxf,
        icon_set,
        icon_id,
        rich_sort_key,
    })
}
fn parse_extension(e: &BytesStart<'_>, d: Decoder) -> Result<Extension> {
    let uri = required_attr(e, b"uri", d)?;
    if uri.is_empty() || uri.len() > 1024 {
        return Err(invalid("invalid Named Sheet Views extension URI"));
    }
    Ok(Extension {
        uri,
        markup: Markup(Vec::new()),
    })
}
pub(crate) fn parse_authored_differential_format(xml: &[u8]) -> Result<DifferentialFormat> {
    if xml.is_empty() || xml.len() > MAX_MARKUP_BYTES {
        return Err(invalid("invalid authored Named Sheet Views dxf size"));
    }
    if validate_fragment_content(xml)? != 1 {
        return Err(invalid(
            "authored Named Sheet Views dxf must be exactly one XML element",
        ));
    }
    let mut document = fragment_document_prefix();
    document.extend_from_slice(
        format!(r#"<nsvFilter filterId="{FRAGMENT_FILTER_ID}"><columnFilter colId="0">"#)
            .as_bytes(),
    );
    document.extend_from_slice(xml);
    document.extend_from_slice(b"</columnFilter></nsvFilter></namedSheetView></namedSheetViews>");
    let mut parsed = parse_named_sheet_views(&document)?;
    let markup = parsed
        .views
        .pop()
        .and_then(|mut view| view.filters.pop())
        .and_then(|mut filter| filter.column_filters.pop())
        .and_then(|column| column.differential_format)
        .ok_or_else(|| invalid("authored fragment must contain exactly one dxf element"))?;
    Ok(DifferentialFormat { markup })
}
pub(crate) fn parse_authored_extension(uri: String, content_xml: &[u8]) -> Result<Extension> {
    if uri.is_empty() || uri.len() > 1024 {
        return Err(invalid("invalid Named Sheet Views extension URI"));
    }
    if content_xml.len() > MAX_MARKUP_BYTES {
        return Err(invalid(
            "authored Named Sheet Views extension content exceeds size limit",
        ));
    }
    validate_fragment_content(content_xml)?;
    let mut extension = Vec::with_capacity(content_xml.len().saturating_add(256));
    extension.extend_from_slice(b"<x:ext");
    write_xml_attribute(
        &mut extension,
        "xmlns:x",
        std::str::from_utf8(CORE).expect("constant namespace is UTF-8"),
    );
    write_xml_attribute(&mut extension, "uri", &uri);
    if content_xml.is_empty() {
        extension.extend_from_slice(b"/>");
    } else {
        extension.push(b'>');
        extension.extend_from_slice(content_xml);
        extension.extend_from_slice(b"</x:ext>");
    }
    let mut document = fragment_document_prefix();
    document.extend_from_slice(b"<extLst>");
    document.extend_from_slice(&extension);
    document.extend_from_slice(b"</extLst></namedSheetView></namedSheetViews>");
    let mut parsed = parse_named_sheet_views(&document)?;
    parsed
        .views
        .pop()
        .and_then(|mut view| view.extensions.pop())
        .filter(|extension| extension.uri == uri)
        .ok_or_else(|| invalid("authored extension did not round-trip"))
}
fn validate_fragment_content(content_xml: &[u8]) -> Result<usize> {
    let mut wrapped = Vec::with_capacity(content_xml.len().saturating_add(13));
    wrapped.extend_from_slice(b"<root>");
    wrapped.extend_from_slice(content_xml);
    wrapped.extend_from_slice(b"</root>");
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut nodes = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(_) => {
                if depth == 1 {
                    roots += 1;
                    if roots > 1 {
                        return Err(invalid("authored XML fragment has multiple root elements"));
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension XML depth overflow"))?;
                nodes += 1;
            },
            Event::Empty(_) => {
                if depth == 1 {
                    roots += 1;
                    if roots > 1 {
                        return Err(invalid("authored XML fragment has multiple root elements"));
                    }
                }
                nodes += 1;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid extension XML nesting"))?;
            },
            Event::Text(text)
                if depth == 1
                    && !text
                        .decode()
                        .map_err(xml_error)?
                        .chars()
                        .all(char::is_whitespace) =>
            {
                return Err(invalid(
                    "authored XML fragment must contain XML markup, not text",
                ));
            },
            Event::CData(_) if depth == 1 => {
                return Err(invalid("authored XML fragment must not contain root CDATA"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,

            _ => {},
        }
        if depth > MAX_FRAGMENT_DEPTH || nodes > MAX_FRAGMENT_NODES {
            return Err(invalid(
                "authored Named Sheet Views extension exceeds structural limits",
            ));
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated authored XML fragment"));
    }
    Ok(roots)
}
fn fragment_document_prefix() -> Vec<u8> {
    format!(
        r#"<namedSheetViews xmlns="{}"><namedSheetView name="Fragment" id="{FRAGMENT_VIEW_ID}">"#,
        std::str::from_utf8(NSV).expect("constant namespace is UTF-8")
    )
    .into_bytes()
}
pub(crate) fn add_extension(extensions: &mut Vec<Extension>, value: Extension) -> Result<()> {
    if extensions.len() >= MAX_EXTENSIONS {
        return Err(invalid("too many Named Sheet Views extensions"));
    }
    extensions.push(value);
    Ok(())
}
pub(crate) fn remove_extension(extensions: &mut Vec<Extension>, uri: &str) -> Option<Extension> {
    extensions
        .iter()
        .position(|extension| extension.uri == uri)
        .map(|index| extensions.remove(index))
}
pub(crate) fn validate_name(v: &str) -> Result<()> {
    let n = v.chars().count();
    if n == 0
        || n > 127
        || v.starts_with("_xlnsv.")
        || v.contains(';')
        || v.chars().next().is_some_and(char::is_whitespace)
        || v.chars().last().is_some_and(char::is_whitespace)
        || v.chars().any(|c| c != ' ' && c.is_whitespace())
    {
        return Err(invalid("invalid named sheet view name"));
    }
    Ok(())
}
pub(crate) fn validate_column_id(value: u32) -> Result<()> {
    if value >= MAX_COLUMNS as u32 {
        Err(invalid(
            "named-sheet-view colId exceeds worksheet column limit",
        ))
    } else {
        Ok(())
    }
}
fn validate_view_collection(views: &[View]) -> Result<()> {
    if views.is_empty() {
        return Err(invalid(
            "Named Sheet Views part must contain a namedSheetView",
        ));
    }
    if views.len() > MAX_VIEWS {
        return Err(invalid("too many named sheet views"));
    }
    let mut names = HashSet::new();
    let mut ids = HashSet::new();
    for view in views {
        if !names.insert(view.name.as_str()) {
            return Err(invalid("duplicate named sheet view name"));
        }
        if !ids.insert(&view.id) {
            return Err(invalid("duplicate named sheet view GUID"));
        }
    }
    Ok(())
}
pub(crate) fn parse_guid(v: &str) -> Result<Guid> {
    let b = v.as_bytes();
    let hyphens = [9usize, 14, 19, 24];
    if b.len() != 38
        || b.first() != Some(&b'{')
        || b.last() != Some(&b'}')
        || !hyphens.into_iter().all(|i| b[i] == b'-')
        || b[1..37]
            .iter()
            .enumerate()
            .any(|(i, x)| ![8usize, 13, 18, 23].contains(&i) && !x.is_ascii_hexdigit())
    {
        return Err(invalid(format!("invalid GUID '{v}'")));
    }
    Ok(Guid(v.into()))
}
fn optional_guid(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<Guid>> {
    attr(e, name, d)?.map(|v| parse_guid(&v)).transpose()
}
fn column_id(e: &BytesStart<'_>, d: Decoder) -> Result<u32> {
    let v = required_u32(e, b"colId", d)?;
    if v >= MAX_COLUMNS as u32 {
        return Err(invalid(
            "named-sheet-view colId exceeds worksheet column limit",
        ));
    }
    Ok(v)
}
pub(crate) fn parse_range(v: &str) -> Result<Range> {
    let mut parts = v.split(':');
    let a = parts.next().unwrap_or("");
    let b = parts.next();
    if a.is_empty() || parts.next().is_some() {
        return Err(invalid(format!("invalid named-sheet-view range '{v}'")));
    }
    let start = Address::from_a1(a).map_err(|e| invalid(e.to_string()))?;
    let end = if let Some(x) = b {
        Address::from_a1(x).map_err(|e| invalid(e.to_string()))?
    } else {
        start
    };
    if start.row() > end.row() || start.column() > end.column() {
        return Err(invalid(format!("reversed named-sheet-view range '{v}'")));
    }
    Ok(Range(v.into()))
}
fn parse_namespace_declarations(e: &BytesStart<'_>, d: Decoder) -> Result<Vec<(String, String)>> {
    let mut declarations = Vec::new();
    let mut names = HashSet::new();
    for attribute in e.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if name != "xmlns" && !name.starts_with("xmlns:") {
            continue;
        }
        if declarations.len() >= MAX_NAMESPACE_DECLARATIONS || !names.insert(name.clone()) {
            return Err(invalid(
                "too many or duplicate Named Sheet Views namespace declarations",
            ));
        }
        let namespace = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        declarations.push((name, namespace));
    }
    Ok(declarations)
}
fn attr(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<String>> {
    let mut value = None;
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() != name {
            continue;
        }
        if value.is_some() {
            return Err(invalid("duplicate attribute"));
        }
        value = Some(
            a.decoded_and_normalized_value(XmlVersion::Explicit1_0, d)
                .map_err(xml_error)?
                .into_owned(),
        );
    }
    Ok(value)
}
fn required_attr(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<String> {
    attr(e, name, d)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(name)
        ))
    })
}
fn bounded_attr(e: &BytesStart<'_>, name: &[u8], d: Decoder, max: usize) -> Result<Option<String>> {
    let v = attr(e, name, d)?;
    if v.as_ref().is_some_and(|x| x.chars().count() > max) {
        return Err(invalid("Named Sheet Views string exceeds limit"));
    }
    Ok(v)
}
fn optional_u32(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<u32>> {
    attr(e, name, d)?
        .map(|v| {
            v.parse::<u32>()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
fn required_u32(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<u32> {
    optional_u32(e, name, d)?.ok_or_else(|| invalid("missing required unsigned integer"))
}
fn bool_attr(e: &BytesStart<'_>, name: &[u8], d: Decoder) -> Result<Option<bool>> {
    attr(e, name, d)?
        .map(|v| match v.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
fn exact(ns: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(ns,ResolveResult::Bound(v)if v.as_ref()==expected)
}
fn core(ns: &ResolveResult<'_>) -> bool {
    exact(ns, CORE) || exact(ns, STRICT)
}
fn inside(ctx: Ctx) -> bool {
    !matches!(ctx, Ctx::Outside)
}
#[cfg(test)]
mod tests {
    use super::super::RELATIONSHIP;
    use super::super::package::{
        discover_named_sheet_views, load_worksheet_named_sheet_views,
        remove_worksheet_named_sheet_views, store_worksheet_named_sheet_views,
    };
    use super::*;
    use crate::auto_filter::{Calendar, DateGroup, Grouping, Icon, Item, Payload, Top10, Values};
    use litchi_opc::{OpcPackage, PackURI};

    fn libreoffice_fixture() -> OpcPackage {
        OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/NamedSheetViews.xlsx"
        ))
        .unwrap()
    }

    fn fixture_worksheet() -> PackURI {
        PackURI::new("/xl/worksheets/sheet1.xml").unwrap()
    }

    #[test]
    fn discovers_libreoffice_fixture_and_parses_filters() {
        let package = libreoffice_fixture();
        let sheet = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        let v = discover_named_sheet_views(&package, sheet.rels())
            .unwrap()
            .unwrap();
        assert_eq!(v.views().len(), 2);
        assert_eq!(v.views()[0].name(), "View1");
        assert_eq!(
            v.views()[0].filters()[0].sort_rules().unwrap().rules()[0]
                .condition()
                .unwrap()
                .reference()
                .as_str(),
            "C1:C8"
        );
        let columns = v.views()[1].filters()[0].column_filters();
        assert_eq!(columns.len(), 2);
        match columns[0].filters()[0].payload().unwrap() {
            Payload::Values(x) => assert_eq!(x.items().len(), 4),
            _ => panic!("wrong payload"),
        }
        match columns[1].filters()[0].payload().unwrap() {
            Payload::Values(x) => {
                assert!(matches!(&x.items()[0], Item::DateGroup(_)))
            },
            _ => panic!("wrong payload"),
        }
    }

    #[test]
    fn serializes_libreoffice_fixture_without_losing_filter_or_sort_metadata() {
        let package = libreoffice_fixture();
        let worksheet = fixture_worksheet();
        let value = load_worksheet_named_sheet_views(&package, &worksheet)
            .unwrap()
            .unwrap();

        let xml = write_named_sheet_views(&value).unwrap();
        let reparsed = parse_named_sheet_views(&xml).unwrap();
        assert_eq!(reparsed, value);
        let text = std::str::from_utf8(&xml).unwrap();
        assert!(text.contains("<filter colId=\"1\"><x:filters>"));
        assert!(text.contains("<sortCondition ref=\"C1:C8\"/>"));
    }

    #[test]
    fn constructs_core_views_and_extends_libreoffice_metadata() {
        let explicit_id = Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap();
        let primary = View::with_id("Personal", explicit_id.clone()).unwrap();
        let mut authored = Views::new(primary);
        let shared = View::new("Shared").unwrap();
        assert!(Guid::new(shared.id().as_str()).is_ok());
        authored.add_view(shared).unwrap();

        assert!(authored.add_view(View::new("Personal").unwrap()).is_err());
        assert!(
            authored
                .add_view(View::with_id("Duplicate GUID", explicit_id).unwrap())
                .is_err()
        );
        assert!(View::new(" _xlnsv.reserved").is_err());
        assert!(Guid::new("not-a-guid").is_err());
        assert!(Range::new("B2:A1").is_err());

        let package = libreoffice_fixture();
        let worksheet = fixture_worksheet();
        let mut preserved = load_worksheet_named_sheet_views(&package, &worksheet)
            .unwrap()
            .unwrap();
        preserved.add_view(View::new("Authored").unwrap()).unwrap();
        assert_eq!(
            parse_named_sheet_views(&write_named_sheet_views(&preserved).unwrap()).unwrap(),
            preserved
        );

        let removed = authored.remove_view("Shared").unwrap().unwrap();
        assert_eq!(removed.name(), "Shared");
        assert!(authored.remove_view("Missing").unwrap().is_none());
        assert!(authored.remove_view("Personal").is_err());
    }

    #[test]
    fn authors_detailed_value_date_and_sort_metadata() {
        let filter_id = Guid::new("{11111111-2222-3333-4444-555555555555}").unwrap();
        let mut column = ColumnFilter::new(1).unwrap();
        let mut payload = Column::new(1).unwrap();
        payload.set_payload(Some(Payload::Values(
            Values::new(
                true,
                Calendar::Gregorian,
                vec![
                    Item::Value("North".into()),
                    Item::DateGroup(
                        DateGroup::new(2026, Some(7), Some(26), None, None, None, Grouping::Day)
                            .unwrap(),
                    ),
                ],
            )
            .unwrap(),
        )));
        column.add_filter(payload).unwrap();

        let mut condition =
            SortCondition::new(SortConditionKind::RichValue, Range::new("B2:B20").unwrap());
        condition
            .set_descending(true)
            .set_custom_list(Some("North,South".into()))
            .unwrap()
            .set_rich_sort_key(Some("Region".into()))
            .unwrap();
        let mut rule = SortRule::new(1).unwrap();
        rule.set_condition(Some(condition)).unwrap();
        let mut rules = SortRules::new();
        rules
            .set_sort_method(SortMethod::PinYin)
            .set_case_sensitive(true)
            .add_rule(rule)
            .unwrap();

        let mut filter = Filter::new(filter_id);
        filter
            .set_reference(Some(Range::new("A1:C20").unwrap()))
            .set_table_id(Some(7))
            .add_column_filter(column)
            .unwrap()
            .set_sort_rules(Some(rules));
        let mut view = View::with_id(
            "Regional",
            Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
        )
        .unwrap();
        view.add_filter(filter).unwrap();
        let authored = Views::new(view);

        let xml = authored.to_xml().unwrap();
        let text = std::str::from_utf8(&xml).unwrap();
        assert!(text.contains(r#"<x:filters blank="1" calendarType="gregorian">"#));
        assert!(text.contains(r#"<x:filter val="North"/>"#));
        assert!(text.contains(
            r#"<x:dateGroupItem year="2026" month="7" day="26" dateTimeGrouping="day"/>"#
        ));
        assert!(text.contains(
            r#"<richSortCondition ref="B2:B20" descending="1" customList="North,South" richSortKey="Region"/>"#
        ));
        assert_eq!(parse_named_sheet_views(&xml).unwrap(), authored);
    }

    #[test]
    fn detailed_authoring_rejects_invalid_relationships_and_filter_values() {
        let mut column = ColumnFilter::new(3).unwrap();
        assert!(column.add_filter(Column::new(2).unwrap()).is_err());
        assert!(DateGroup::new(2026, None, Some(26), None, None, None, Grouping::Day,).is_err());
        assert!(Top10::new(true, true, 101.0, None).is_err());
        assert!(Icon::new(crate::auto_filter::IconSet::ThreeArrows, 3).is_err());

        assert!(ColumnFilter::new(MAX_COLUMNS as u32).is_err());
    }

    #[test]
    fn rejects_unsafe_fragments_and_mismatched_color_sort_formats() {
        for xml in [
            br#"<dxf xmlns="urn:wrong"/>"#.as_slice(),
            br#"<dxf xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"/><dxf xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"/>"#,
            br#"<!DOCTYPE x><dxf xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"/>"#,
        ] {
            assert!(DifferentialFormat::from_xml(xml).is_err());
        }
        assert!(Extension::new("", b"").is_err());
        assert!(Extension::new("urn:test", br#"<x:payload xmlns:x="urn:test">"#,).is_err());
        assert!(
            Extension::new(
                "urn:test",
                br#"<x:one xmlns:x="urn:test"/><x:two xmlns:x="urn:test"/>"#,
            )
            .is_err()
        );
        assert!(Extension::new("urn:test", b"not markup").is_err());
        assert!(
            Extension::new(
                "urn:test",
                br#"<?unsafe value?><x:payload xmlns:x="urn:test"/>"#,
            )
            .is_err()
        );

        let mut condition =
            SortCondition::new(SortConditionKind::Standard, Range::new("A1:A2").unwrap());
        assert!(condition.set_color_sort(SortBy::Value, 0).is_err());
        condition.set_color_sort(SortBy::FontColor, 0).unwrap();
        let mut rule = SortRule::new(0).unwrap();
        assert!(rule.set_condition(Some(condition.clone())).is_err());
        rule.set_differential_format(Some(DifferentialFormat::empty()))
            .unwrap()
            .set_condition(Some(condition))
            .unwrap();
        assert!(rule.set_differential_format(None).is_err());
    }

    #[test]
    fn package_crud_round_trips_real_fixture_and_removes_unreferenced_part() {
        let mut package = libreoffice_fixture();
        let worksheet = fixture_worksheet();
        let value = load_worksheet_named_sheet_views(&package, &worksheet)
            .unwrap()
            .unwrap();

        store_worksheet_named_sheet_views(&mut package, &worksheet, &value).unwrap();
        assert_eq!(
            load_worksheet_named_sheet_views(&package, &worksheet).unwrap(),
            Some(value.clone())
        );

        assert!(remove_worksheet_named_sheet_views(&mut package, &worksheet).unwrap());
        assert_eq!(
            load_worksheet_named_sheet_views(&package, &worksheet).unwrap(),
            None
        );
        assert!(!remove_worksheet_named_sheet_views(&mut package, &worksheet).unwrap());

        store_worksheet_named_sheet_views(&mut package, &worksheet, &value).unwrap();
        let relationship = package
            .get_part(&worksheet)
            .unwrap()
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == RELATIONSHIP)
            .unwrap();
        assert_eq!(
            relationship.target_partname().unwrap(),
            PackURI::new("/xl/namedSheetViews/namedSheetView1.xml").unwrap()
        );

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("named-sheet-views.xlsx");
        package.save(&path).unwrap();
        let reopened = OpcPackage::open(&path).unwrap();
        assert_eq!(
            load_worksheet_named_sheet_views(&reopened, &worksheet).unwrap(),
            Some(value)
        );
    }

    #[test]
    fn package_crud_rejects_invalid_existing_graph_before_replacement() {
        let mut package = libreoffice_fixture();
        let worksheet = fixture_worksheet();
        let value = load_worksheet_named_sheet_views(&package, &worksheet)
            .unwrap()
            .unwrap();
        let part_name = PackURI::new("/xl/namedSheetViews/namedSheetView1.xml").unwrap();
        let original = package.get_part(&part_name).unwrap().blob().to_vec();
        package
            .get_part_mut(&worksheet)
            .unwrap()
            .rels_mut()
            .add_relationship(
                RELATIONSHIP.into(),
                "../namedSheetViews/namedSheetView1.xml".into(),
                "rIdDuplicateNamedSheetView".into(),
                false,
            );

        assert!(store_worksheet_named_sheet_views(&mut package, &worksheet, &value).is_err());
        assert_eq!(package.get_part(&part_name).unwrap().blob(), original);
        assert!(remove_worksheet_named_sheet_views(&mut package, &worksheet).is_err());
    }
    #[test]
    fn poi_fixture_has_no_named_sheet_views_relationship() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx"
        ))
        .unwrap();
        let sheet = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        assert!(
            discover_named_sheet_views(&package, sheet.rels())
                .unwrap()
                .is_none()
        )
    }
    #[test]
    fn parses_mce_rich_sort_and_extensions() {
        let xml=br#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:no"><namedSheetView name="Rich View" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><mc:AlternateContent><mc:Choice Requires="u"><u:no/></mc:Choice><mc:Fallback><nsvFilter filterId="{11111111-2222-3333-4444-555555555555}" ref="A1:B9" tableId="0"><sortRules sortMethod="pinYin" caseSensitive="1"><sortRule colId="1"><richSortCondition ref="B1:B9" descending="1" richSortKey="City"/></sortRule></sortRules></nsvFilter></mc:Fallback></mc:AlternateContent><extLst><x:ext uri="urn:test"><u:payload/></x:ext></extLst></namedSheetView></namedSheetViews>"#;
        let v = parse_named_sheet_views(xml).unwrap();
        let rules = v.views()[0].filters()[0].sort_rules().unwrap();
        assert_eq!(rules.sort_method(), SortMethod::PinYin);
        let c = rules.rules()[0].condition().unwrap();
        assert_eq!(c.kind(), SortConditionKind::RichValue);
        assert_eq!(c.rich_sort_key(), Some("City"));
        assert_eq!(v.views()[0].extensions()[0].uri(), "urn:test");
        assert_eq!(
            parse_named_sheet_views(&write_named_sheet_views(&v).unwrap()).unwrap(),
            v
        );
    }
    #[test]
    fn rejects_schema_guid_range_and_security_errors() {
        let cases = [
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"/>"#,
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name=" bad" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"/></namedSheetViews>"#,
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name="v" id="bad"/></namedSheetViews>"#,
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name="v" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><nsvFilter filterId="{11111111-2222-3333-4444-555555555555}" ref="B2:A1"/></namedSheetView></namedSheetViews>"#,
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name="v" id="{01234567-89AB-CDEF-0123-456789ABCDEF}">text</namedSheetView></namedSheetViews>"#,
            r#"<namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"><namedSheetView name="v" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"/></namedSheetViews><extra/>"#,
            r#"<!DOCTYPE x><namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"/>"#,
        ];
        for xml in cases {
            assert!(parse_named_sheet_views(xml.as_bytes()).is_err(), "{xml}")
        }
    }
}

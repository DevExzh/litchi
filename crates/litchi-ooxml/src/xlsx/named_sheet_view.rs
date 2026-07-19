//! Immutable static reader for the MS-XLSX Named Sheet Views part.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::xlsx::Cell;
use crate::xlsx::auto_filter::{FilterColumnDefinition, parse_auto_filter};
use crate::xlsx::sort::{SortBy, SortMethod};
use litchi_opc::{OpcPackage, Relationships};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};
use std::borrow::Cow;
use std::collections::HashSet;

const NSV: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews";
const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const RICH: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
const RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView";
const CONTENT_TYPE: &str = "application/vnd.ms-excel.namedsheetviews+xml";
const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_VIEWS: usize = 1024;
const MAX_FILTERS: usize = 65_536;
const MAX_COLUMNS: usize = 16_384;
const MAX_RETAINED_BYTES: usize = 8 * 1024 * 1024;
const MAX_MARKUP_BYTES: usize = 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NamedSheetViewGuid(String);
impl NamedSheetViewGuid {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewRange(String);
impl NamedSheetViewRange {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewMarkup(Vec<u8>);
impl NamedSheetViewMarkup {
    pub fn xml(&self) -> &[u8] {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewExtension {
    uri: String,
    markup: NamedSheetViewMarkup,
}
impl NamedSheetViewExtension {
    pub fn uri(&self) -> &str {
        &self.uri
    }
    pub fn markup(&self) -> &NamedSheetViewMarkup {
        &self.markup
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NamedSheetViewSortConditionKind {
    Standard,
    RichValue,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NamedSheetViewIconSet {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
    ThreeStars,
    ThreeTriangles,
    FiveBoxes,
    NoIcons,
}
impl NamedSheetViewIconSet {
    fn parse(v: &str) -> Result<Self> {
        use NamedSheetViewIconSet::*;
        match v {
            "3Arrows" => Ok(ThreeArrows),
            "3ArrowsGray" => Ok(ThreeArrowsGray),
            "3Flags" => Ok(ThreeFlags),
            "3TrafficLights1" => Ok(ThreeTrafficLights1),
            "3TrafficLights2" => Ok(ThreeTrafficLights2),
            "3Signs" => Ok(ThreeSigns),
            "3Symbols" => Ok(ThreeSymbols),
            "3Symbols2" => Ok(ThreeSymbols2),
            "4Arrows" => Ok(FourArrows),
            "4ArrowsGray" => Ok(FourArrowsGray),
            "4RedToBlack" => Ok(FourRedToBlack),
            "4Rating" => Ok(FourRating),
            "4TrafficLights" => Ok(FourTrafficLights),
            "5Arrows" => Ok(FiveArrows),
            "5ArrowsGray" => Ok(FiveArrowsGray),
            "5Rating" => Ok(FiveRating),
            "5Quarters" => Ok(FiveQuarters),
            "3Stars" => Ok(ThreeStars),
            "3Triangles" => Ok(ThreeTriangles),
            "5Boxes" => Ok(FiveBoxes),
            "NoIcons" => Ok(NoIcons),
            _ => Err(invalid(format!("invalid named-sheet-view icon set '{v}'"))),
        }
    }
    fn cardinality(self) -> Option<u32> {
        use NamedSheetViewIconSet::*;
        Some(match self {
            ThreeArrows | ThreeArrowsGray | ThreeFlags | ThreeTrafficLights1
            | ThreeTrafficLights2 | ThreeSigns | ThreeSymbols | ThreeSymbols2 | ThreeStars
            | ThreeTriangles => 3,
            FourArrows | FourArrowsGray | FourRedToBlack | FourRating | FourTrafficLights => 4,
            FiveArrows | FiveArrowsGray | FiveRating | FiveQuarters | FiveBoxes => 5,
            NoIcons => return None,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewSortCondition {
    kind: NamedSheetViewSortConditionKind,
    reference: NamedSheetViewRange,
    descending: bool,
    sort_by: SortBy,
    custom_list: Option<String>,
    differential_format_id: Option<u32>,
    icon_set: Option<NamedSheetViewIconSet>,
    icon_id: Option<u32>,
    rich_sort_key: Option<String>,
}
impl NamedSheetViewSortCondition {
    pub fn kind(&self) -> NamedSheetViewSortConditionKind {
        self.kind
    }
    pub fn reference(&self) -> &NamedSheetViewRange {
        &self.reference
    }
    pub fn descending(&self) -> bool {
        self.descending
    }
    pub fn sort_by(&self) -> SortBy {
        self.sort_by
    }
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    pub fn icon_set(&self) -> Option<NamedSheetViewIconSet> {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
    pub fn rich_sort_key(&self) -> Option<&str> {
        self.rich_sort_key.as_deref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewSortRule {
    column_id: u32,
    id: Option<NamedSheetViewGuid>,
    differential_format: Option<NamedSheetViewMarkup>,
    condition: Option<NamedSheetViewSortCondition>,
}
impl NamedSheetViewSortRule {
    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    pub fn id(&self) -> Option<&NamedSheetViewGuid> {
        self.id.as_ref()
    }
    pub fn differential_format(&self) -> Option<&NamedSheetViewMarkup> {
        self.differential_format.as_ref()
    }
    pub fn condition(&self) -> Option<&NamedSheetViewSortCondition> {
        self.condition.as_ref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedSheetViewSortRules {
    sort_method: SortMethod,
    case_sensitive: bool,
    rules: Vec<NamedSheetViewSortRule>,
    extensions: Vec<NamedSheetViewExtension>,
}
impl NamedSheetViewSortRules {
    pub fn sort_method(&self) -> SortMethod {
        self.sort_method
    }
    pub fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }
    pub fn rules(&self) -> &[NamedSheetViewSortRule] {
        &self.rules
    }
    pub fn extensions(&self) -> &[NamedSheetViewExtension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct NamedSheetViewColumnFilter {
    column_id: u32,
    id: Option<NamedSheetViewGuid>,
    differential_format: Option<NamedSheetViewMarkup>,
    filters: Vec<FilterColumnDefinition>,
    extensions: Vec<NamedSheetViewExtension>,
}
impl NamedSheetViewColumnFilter {
    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    pub fn id(&self) -> Option<&NamedSheetViewGuid> {
        self.id.as_ref()
    }
    pub fn differential_format(&self) -> Option<&NamedSheetViewMarkup> {
        self.differential_format.as_ref()
    }
    pub fn filters(&self) -> &[FilterColumnDefinition] {
        &self.filters
    }
    pub fn extensions(&self) -> &[NamedSheetViewExtension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct NamedSheetViewFilter {
    filter_id: NamedSheetViewGuid,
    reference: Option<NamedSheetViewRange>,
    table_id: Option<u32>,
    column_filters: Vec<NamedSheetViewColumnFilter>,
    sort_rules: Option<NamedSheetViewSortRules>,
    extensions: Vec<NamedSheetViewExtension>,
}
impl NamedSheetViewFilter {
    pub fn filter_id(&self) -> &NamedSheetViewGuid {
        &self.filter_id
    }
    pub fn reference(&self) -> Option<&NamedSheetViewRange> {
        self.reference.as_ref()
    }
    pub fn table_id(&self) -> Option<u32> {
        self.table_id
    }
    pub fn column_filters(&self) -> &[NamedSheetViewColumnFilter] {
        &self.column_filters
    }
    pub fn sort_rules(&self) -> Option<&NamedSheetViewSortRules> {
        self.sort_rules.as_ref()
    }
    pub fn extensions(&self) -> &[NamedSheetViewExtension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct NamedSheetView {
    name: String,
    id: NamedSheetViewGuid,
    filters: Vec<NamedSheetViewFilter>,
    extensions: Vec<NamedSheetViewExtension>,
}
impl NamedSheetView {
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn id(&self) -> &NamedSheetViewGuid {
        &self.id
    }
    pub fn filters(&self) -> &[NamedSheetViewFilter] {
        &self.filters
    }
    pub fn extensions(&self) -> &[NamedSheetViewExtension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct NamedSheetViews {
    views: Vec<NamedSheetView>,
    extensions: Vec<NamedSheetViewExtension>,
}
impl NamedSheetViews {
    pub fn views(&self) -> &[NamedSheetView] {
        &self.views
    }
    pub fn extensions(&self) -> &[NamedSheetViewExtension] {
        &self.extensions
    }
}

pub(crate) fn discover_named_sheet_views(
    package: &OpcPackage,
    relationships: &Relationships,
) -> Result<Option<NamedSheetViews>> {
    let mut found = relationships.iter().filter(|r| r.reltype() == RELATIONSHIP);
    let Some(relationship) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "worksheet has multiple Named Sheet Views relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("Named Sheet Views relationship cannot be external"));
    }
    let part = package.get_part(&relationship.target_partname()?)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: CONTENT_TYPE.into(),
            got: part.content_type().into(),
        });
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Named Sheet Views part must not have relationships",
        ));
    }
    parse_named_sheet_views(part.blob()).map(Some)
}

pub fn parse_named_sheet_views(xml: &[u8]) -> Result<NamedSheetViews> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("Named Sheet Views part exceeds size limit"));
    }
    let mut caps = MceCapabilities::default();
    for ns in [NSV, X14, RICH] {
        caps.understand_namespace(String::from_utf8_lossy(ns).into_owned());
    }
    let mut limits = MceLimits::default();
    limits.max_input_bytes = MAX_PART_BYTES;
    limits.max_output_bytes = MAX_PART_BYTES * 2;
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
    Extension(ExtOwner, NamedSheetViewExtension),
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
    root: Option<NamedSheetViews>,
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
    value: NamedSheetView,
    phase: u8,
}
struct FilterBuilder {
    value: NamedSheetViewFilter,
    phase: u8,
}
struct ColumnBuilder {
    value: NamedSheetViewColumnFilter,
    phase: u8,
}
struct SortRulesBuilder {
    value: NamedSheetViewSortRules,
    phase: u8,
}
struct SortRuleBuilder {
    value: NamedSheetViewSortRule,
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
                self.begin_root()?;
                self.stack.push(Ctx::Root)
            },
            (Ctx::Root, true, b"namedSheetView") => {
                self.begin_view(e, d)?;
                self.stack.push(Ctx::View)
            },
            (Ctx::View, true, b"nsvFilter") => {
                self.begin_filter(e, d)?;
                self.stack.push(Ctx::Filter)
            },
            (Ctx::Filter, true, b"columnFilter") => {
                self.begin_column(e, d)?;
                self.stack.push(Ctx::Column)
            },
            (Ctx::Filter, true, b"sortRules") => {
                self.begin_sort_rules(e, d)?;
                self.stack.push(Ctx::SortRules)
            },
            (Ctx::SortRules, true, b"sortRule") => {
                self.begin_sort_rule(e, d)?;
                self.stack.push(Ctx::SortRule)
            },
            (Ctx::SortRule, true, b"sortCondition") => {
                self.add_condition(e, d, NamedSheetViewSortConditionKind::Standard)?;
                self.stack.push(Ctx::Leaf)
            },
            (Ctx::SortRule, true, b"richSortCondition") => {
                self.add_condition(e, d, NamedSheetViewSortConditionKind::RichValue)?;
                self.stack.push(Ctx::Leaf)
            },
            (parent, true, b"extLst") => {
                let owner = self.ext_owner(parent)?;
                self.begin_ext(owner)?;
                self.stack.push(Ctx::ExtList(owner))
            },
            (Ctx::Column, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::Column)?;
                self.begin_capture(Payload::Dxf(DxfOwner::Column), e)?;
                self.stack.push(Ctx::Captured)
            },
            (Ctx::SortRule, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::SortRule)?;
                self.begin_capture(Payload::Dxf(DxfOwner::SortRule), e)?;
                self.stack.push(Ctx::Captured)
            },
            (Ctx::Column, true, b"filter") => {
                self.begin_capture(Payload::Filter, e)?;
                self.stack.push(Ctx::Captured)
            },
            (Ctx::ExtList(owner), _, b"ext") if core(ns) => {
                let x = parse_extension(e, d)?;
                self.begin_capture(Payload::Extension(owner, x), e)?;
                self.stack.push(Ctx::Captured)
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
                self.finish_view()?
            },
            (Ctx::View, true, b"nsvFilter") => {
                self.begin_filter(e, d)?;
                self.finish_filter()?
            },
            (Ctx::Filter, true, b"columnFilter") => {
                self.begin_column(e, d)?;
                self.finish_column()?
            },
            (Ctx::Filter, true, b"sortRules") => {
                self.begin_sort_rules(e, d)?;
                self.finish_sort_rules()?
            },
            (Ctx::SortRules, true, b"sortRule") => {
                self.begin_sort_rule(e, d)?;
                self.finish_sort_rule()?
            },
            (Ctx::SortRule, true, b"sortCondition") => {
                self.add_condition(e, d, NamedSheetViewSortConditionKind::Standard)?
            },
            (Ctx::SortRule, true, b"richSortCondition") => {
                self.add_condition(e, d, NamedSheetViewSortConditionKind::RichValue)?
            },
            (parent, true, b"extLst") => {
                let owner = self.ext_owner(parent)?;
                self.begin_ext(owner)?
            },
            (Ctx::Column, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::Column)?;
                self.empty_markup(Payload::Dxf(DxfOwner::Column), e)?
            },
            (Ctx::SortRule, true, b"dxf") => {
                self.prepare_dxf(DxfOwner::SortRule)?;
                self.empty_markup(Payload::Dxf(DxfOwner::SortRule), e)?
            },
            (Ctx::Column, true, b"filter") => self.empty_markup(Payload::Filter, e)?,
            (Ctx::ExtList(owner), _, b"ext") if core(ns) => {
                let x = parse_extension(e, d)?;
                self.empty_markup(Payload::Extension(owner, x), e)?
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
    fn begin_root(&mut self) -> Result<()> {
        if self.seen_root {
            return Err(invalid("duplicate namedSheetViews root"));
        }
        self.seen_root = true;
        self.root_phase = 0;
        self.root = Some(NamedSheetViews {
            views: Vec::new(),
            extensions: Vec::new(),
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
            value: NamedSheetView {
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
            value: NamedSheetViewFilter {
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
            value: NamedSheetViewColumnFilter {
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
        let method = SortMethod::parse(attr(e, b"sortMethod", d)?.as_deref().unwrap_or("none"))
            .ok_or_else(|| invalid("invalid named-sheet-view sortMethod"))?;
        self.sort_rules = Some(SortRulesBuilder {
            value: NamedSheetViewSortRules {
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
            value: NamedSheetViewSortRule {
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
        kind: NamedSheetViewSortConditionKind,
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
    fn ext_owner(&self, parent: Ctx) -> Result<ExtOwner> {
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
                self.root_phase = 1
            },
            ExtOwner::View => {
                let v = self.view.as_mut().unwrap();
                if v.phase > 0 {
                    return Err(invalid("duplicate view extLst"));
                }
                v.phase = 1
            },
            ExtOwner::Filter => {
                let v = self.filter.as_mut().unwrap();
                if v.phase > 1 {
                    return Err(invalid("duplicate filter extLst"));
                }
                v.phase = 2
            },
            ExtOwner::Column => {
                let v = self.column.as_mut().unwrap();
                if v.phase > 1 {
                    return Err(invalid("duplicate columnFilter extLst"));
                }
                v.phase = 2
            },
            ExtOwner::SortRules => {
                let v = self.sort_rules.as_mut().unwrap();
                if v.phase > 0 {
                    return Err(invalid("duplicate sortRules extLst"));
                }
                v.phase = 1
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
                v.value.differential_format = Some(NamedSheetViewMarkup(Vec::new()))
            },
            DxfOwner::SortRule => {
                let v = self.sort_rule.as_mut().unwrap();
                if v.phase > 0 || v.value.differential_format.is_some() {
                    return Err(invalid("duplicate or misplaced sortRule dxf"));
                }
                v.value.differential_format = Some(NamedSheetViewMarkup(Vec::new()))
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
            self.attach(capture.payload, capture.writer.into_inner())?
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
                column.value.filters.push(parsed)
            },
            Payload::Dxf(DxfOwner::Column) => {
                self.column.as_mut().unwrap().value.differential_format =
                    Some(NamedSheetViewMarkup(markup))
            },
            Payload::Dxf(DxfOwner::SortRule) => {
                self.sort_rule.as_mut().unwrap().value.differential_format =
                    Some(NamedSheetViewMarkup(markup))
            },
            Payload::Extension(owner, mut x) => {
                x.markup = NamedSheetViewMarkup(markup);
                match owner {
                    ExtOwner::Root => self.root.as_mut().unwrap().extensions.push(x),
                    ExtOwner::View => self.view.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::Filter => self.filter.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::Column => self.column.as_mut().unwrap().value.extensions.push(x),
                    ExtOwner::SortRules => {
                        self.sort_rules.as_mut().unwrap().value.extensions.push(x)
                    },
                }
            },
        }
        Ok(())
    }
    fn finish(self) -> Result<NamedSheetViews> {
        let root = self
            .root
            .ok_or_else(|| invalid("missing namedSheetViews root"))?;
        if root.views.is_empty() {
            return Err(invalid(
                "Named Sheet Views part must contain a namedSheetView",
            ));
        }
        let mut names = HashSet::new();
        let mut ids = HashSet::new();
        for view in &root.views {
            if !names.insert(view.name.clone()) {
                return Err(invalid("duplicate named sheet view name"));
            }
            if !ids.insert(view.id.clone()) {
                return Err(invalid("duplicate named sheet view GUID"));
            }
        }
        Ok(root)
    }
}

fn parse_filter_payload(markup: &[u8]) -> Result<FilterColumnDefinition> {
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
                writer.write_event(Event::Start(c)).map_err(xml_error)?
            },
            Event::Empty(e) if !seen => {
                seen = true;
                let c = adapt_filter_root(&e, decoder)?;
                writer.write_event(Event::Empty(c)).map_err(xml_error)?;
                break;
            },
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e)).map_err(xml_error)?
            },
            Event::Empty(e) => writer.write_event(Event::Empty(e)).map_err(xml_error)?,
            Event::End(e) => {
                depth -= 1;
                if depth == 0 {
                    writer
                        .write_event(Event::End(BytesEnd::new("x:filterColumn")))
                        .map_err(xml_error)?;
                    break;
                } else {
                    writer.write_event(Event::End(e)).map_err(xml_error)?
                }
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
            c.push_attribute((std::str::from_utf8(name).unwrap(), Cow::Owned(value)))
        }
    }
    Ok(c)
}
fn parse_condition(
    e: &BytesStart<'_>,
    d: Decoder,
    kind: NamedSheetViewSortConditionKind,
) -> Result<NamedSheetViewSortCondition> {
    let reference = parse_range(&required_attr(e, b"ref", d)?)?;
    let sort_by = SortBy::parse(attr(e, b"sortBy", d)?.as_deref().unwrap_or("value"))
        .ok_or_else(|| invalid("invalid named-sheet-view sortBy"))?;
    let dxf = optional_u32(e, b"dxfId", d)?;
    let icon_set = attr(e, b"iconSet", d)?
        .map(|v| NamedSheetViewIconSet::parse(&v))
        .transpose()?;
    let icon_id = optional_u32(e, b"iconId", d)?;
    if sort_by == SortBy::Icon {
        if dxf.is_some() {
            return Err(invalid("icon sort condition cannot have dxfId"));
        }
        if let (Some(set), Some(id)) = (icon_set, icon_id) {
            if set.cardinality().is_some_and(|n| id >= n) {
                return Err(invalid("iconId is outside icon set"));
            }
        }
    } else if icon_set.is_some() || icon_id.is_some() {
        return Err(invalid(
            "non-icon sort condition cannot have icon attributes",
        ));
    }
    if !matches!(sort_by, SortBy::CellColor | SortBy::FontColor) && dxf.is_some() {
        return Err(invalid("dxfId is only valid for color sorting"));
    }
    let rich_sort_key = if kind == NamedSheetViewSortConditionKind::RichValue {
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
    Ok(NamedSheetViewSortCondition {
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
fn parse_extension(e: &BytesStart<'_>, d: Decoder) -> Result<NamedSheetViewExtension> {
    let uri = required_attr(e, b"uri", d)?;
    if uri.is_empty() || uri.len() > 1024 {
        return Err(invalid("invalid Named Sheet Views extension URI"));
    }
    Ok(NamedSheetViewExtension {
        uri,
        markup: NamedSheetViewMarkup(Vec::new()),
    })
}
fn validate_name(v: &str) -> Result<()> {
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
fn parse_guid(v: &str) -> Result<NamedSheetViewGuid> {
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
    Ok(NamedSheetViewGuid(v.into()))
}
fn optional_guid(
    e: &BytesStart<'_>,
    name: &[u8],
    d: Decoder,
) -> Result<Option<NamedSheetViewGuid>> {
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
fn parse_range(v: &str) -> Result<NamedSheetViewRange> {
    let mut parts = v.split(':');
    let a = parts.next().unwrap_or("");
    let b = parts.next();
    if a.is_empty() || parts.next().is_some() {
        return Err(invalid(format!("invalid named-sheet-view range '{v}'")));
    }
    let start = Cell::reference_to_coords(a).map_err(|e| invalid(e.to_string()))?;
    let end = if let Some(x) = b {
        Cell::reference_to_coords(x).map_err(|e| invalid(e.to_string()))?
    } else {
        start
    };
    if start.0 > end.0 || start.1 > end.1 {
        return Err(invalid(format!("reversed named-sheet-view range '{v}'")));
    }
    Ok(NamedSheetViewRange(v.into()))
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
        )
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
fn invalid(v: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(v.into())
}
fn xml_error(v: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(v.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::auto_filter::{FilterColumnPayload, FilterItem};
    use litchi_opc::{OpcPackage, PackURI};
    #[test]
    fn discovers_libreoffice_fixture_and_parses_filters() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/NamedSheetViews.xlsx"
        ))
        .unwrap();
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
            FilterColumnPayload::Values(x) => assert_eq!(x.items().len(), 4),
            _ => panic!("wrong payload"),
        }
        match columns[1].filters()[0].payload().unwrap() {
            FilterColumnPayload::Values(x) => {
                assert!(matches!(&x.items()[0], FilterItem::DateGroup(_)))
            },
            _ => panic!("wrong payload"),
        }
    }
    #[test]
    fn poi_fixture_has_no_named_sheet_views_relationship() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../3rdparty/poi/test-data/spreadsheet/right-to-left.xlsx"
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
        assert_eq!(c.kind(), NamedSheetViewSortConditionKind::RichValue);
        assert_eq!(c.rich_sort_key(), Some("City"));
        assert_eq!(v.views()[0].extensions()[0].uri(), "urn:test")
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

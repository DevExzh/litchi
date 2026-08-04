//! Immutable static XLSX worksheet-view read model.

use crate::error::{Error, Result};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_VIEWS: usize = 1024;
const MAX_SELECTION_RANGES: usize = 32_767;
const MAX_RETAINED_MARKUP: usize = 4 * 1024 * 1024;
const MAX_PIVOT_AREA_MARKUP: usize = 1024 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ViewType {
    Normal,
    PageBreakPreview,
    PageLayout,
}
impl ViewType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "normal" => Ok(Self::Normal),
            "pageBreakPreview" => Ok(Self::PageBreakPreview),
            "pageLayout" => Ok(Self::PageLayout),
            _ => Err(invalid(format!("invalid worksheet-view type '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PanePosition {
    BottomRight,
    TopRight,
    BottomLeft,
    TopLeft,
}
impl PanePosition {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "bottomRight" => Ok(Self::BottomRight),
            "topRight" => Ok(Self::TopRight),
            "bottomLeft" => Ok(Self::BottomLeft),
            "topLeft" => Ok(Self::TopLeft),
            _ => Err(invalid(format!("invalid worksheet-view pane '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneState {
    Split,
    Frozen,
    FrozenSplit,
}
impl PaneState {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "split" => Ok(Self::Split),
            "frozen" => Ok(Self::Frozen),
            "frozenSplit" => Ok(Self::FrozenSplit),
            _ => Err(invalid(format!(
                "invalid worksheet-view pane state '{value}'"
            ))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotSelectionAxis {
    Row,
    Column,
    Page,
    Values,
}
impl PivotSelectionAxis {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "axisRow" => Ok(Self::Row),
            "axisCol" => Ok(Self::Column),
            "axisPage" => Ok(Self::Page),
            "axisValues" => Ok(Self::Values),
            _ => Err(invalid(format!("invalid pivot-selection axis '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotAreaType {
    None,
    Normal,
    Data,
    All,
    Origin,
    Button,
    TopRight,
    TopEnd,
}
impl PivotAreaType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "normal" => Ok(Self::Normal),
            "data" => Ok(Self::Data),
            "all" => Ok(Self::All),
            "origin" => Ok(Self::Origin),
            "button" => Ok(Self::Button),
            "topRight" => Ok(Self::TopRight),
            "topEnd" => Ok(Self::TopEnd),
            _ => Err(invalid(format!("invalid pivot-area type '{value}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellReference(String);
impl CellReference {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeReference(String);
impl RangeReference {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sqref(Vec<RangeReference>);
impl Sqref {
    pub fn ranges(&self) -> &[RangeReference] {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    uri: String,
    markup: Vec<u8>,
}
impl Extension {
    pub fn uri(&self) -> &str {
        &self.uri
    }
    /// MCE-processed extension markup. It is retained, not interpreted or executed.
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Pane {
    x_split: Option<f64>,
    y_split: Option<f64>,
    top_left_cell: Option<CellReference>,
    active_pane: PanePosition,
    state: PaneState,
}
impl Pane {
    pub fn x_split(&self) -> Option<f64> {
        self.x_split
    }
    pub fn y_split(&self) -> Option<f64> {
        self.y_split
    }
    pub fn top_left_cell(&self) -> Option<&CellReference> {
        self.top_left_cell.as_ref()
    }
    pub fn active_pane(&self) -> PanePosition {
        self.active_pane
    }
    pub fn state(&self) -> PaneState {
        self.state
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Selection {
    pane: PanePosition,
    active_cell: CellReference,
    active_cell_id: u32,
    sqref: Sqref,
}
impl Selection {
    pub fn pane(&self) -> PanePosition {
        self.pane
    }
    pub fn active_cell(&self) -> &CellReference {
        &self.active_cell
    }
    pub fn active_cell_id(&self) -> u32 {
        self.active_cell_id
    }
    pub fn sqref(&self) -> &Sqref {
        &self.sqref
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotArea {
    field: Option<i32>,
    area_type: PivotAreaType,
    data_only: bool,
    label_only: bool,
    grand_row: bool,
    grand_column: bool,
    cache_index: bool,
    outline: bool,
    offset: Option<RangeReference>,
    collapsed_levels_are_subtotals: bool,
    axis: Option<PivotSelectionAxis>,
    field_position: Option<u32>,
    markup: Vec<u8>,
}
impl PivotArea {
    pub fn field(&self) -> Option<i32> {
        self.field
    }
    pub fn area_type(&self) -> PivotAreaType {
        self.area_type
    }
    pub fn data_only(&self) -> bool {
        self.data_only
    }
    pub fn label_only(&self) -> bool {
        self.label_only
    }
    pub fn grand_row(&self) -> bool {
        self.grand_row
    }
    pub fn grand_column(&self) -> bool {
        self.grand_column
    }
    pub fn cache_index(&self) -> bool {
        self.cache_index
    }
    pub fn outline(&self) -> bool {
        self.outline
    }
    pub fn offset(&self) -> Option<&RangeReference> {
        self.offset.as_ref()
    }
    pub fn collapsed_levels_are_subtotals(&self) -> bool {
        self.collapsed_levels_are_subtotals
    }
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    pub fn field_position(&self) -> Option<u32> {
        self.field_position
    }
    /// Complete, bounded, MCE-processed `pivotArea` markup, including references and extensions.
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotSelection {
    pane: PanePosition,
    show_header: bool,
    label: bool,
    data: bool,
    extendable: bool,
    count: u32,
    axis: Option<PivotSelectionAxis>,
    dimension: u32,
    start: u32,
    min: u32,
    max: u32,
    active_row: u32,
    active_column: u32,
    previous_row: u32,
    previous_column: u32,
    click: u32,
    relationship_id: Option<String>,
    area: PivotArea,
}
impl PivotSelection {
    pub fn pane(&self) -> PanePosition {
        self.pane
    }
    pub fn show_header(&self) -> bool {
        self.show_header
    }
    pub fn label(&self) -> bool {
        self.label
    }
    pub fn data(&self) -> bool {
        self.data
    }
    pub fn extendable(&self) -> bool {
        self.extendable
    }
    pub fn count(&self) -> u32 {
        self.count
    }
    pub fn axis(&self) -> Option<PivotSelectionAxis> {
        self.axis
    }
    pub fn dimension(&self) -> u32 {
        self.dimension
    }
    pub fn start(&self) -> u32 {
        self.start
    }
    pub fn min(&self) -> u32 {
        self.min
    }
    pub fn max(&self) -> u32 {
        self.max
    }
    pub fn active_row(&self) -> u32 {
        self.active_row
    }
    pub fn active_column(&self) -> u32 {
        self.active_column
    }
    pub fn previous_row(&self) -> u32 {
        self.previous_row
    }
    pub fn previous_column(&self) -> u32 {
        self.previous_column
    }
    pub fn click(&self) -> u32 {
        self.click
    }
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
    pub fn area(&self) -> &PivotArea {
        &self.area
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct View {
    workbook_view_id: u32,
    window_protection: bool,
    show_formulas: bool,
    show_grid_lines: bool,
    show_row_col_headers: bool,
    show_zeros: bool,
    right_to_left: bool,
    tab_selected: bool,
    show_ruler: bool,
    show_outline_symbols: bool,
    default_grid_color: bool,
    show_white_space: bool,
    view_type: ViewType,
    top_left_cell: Option<CellReference>,
    color_id: u32,
    zoom_scale: u16,
    zoom_scale_normal: u16,
    zoom_scale_sheet_layout_view: u16,
    zoom_scale_page_layout_view: u16,
    pane: Option<Pane>,
    selections: Vec<Selection>,
    pivot_selections: Vec<PivotSelection>,
    extensions: Vec<Extension>,
}
impl View {
    pub fn workbook_view_id(&self) -> u32 {
        self.workbook_view_id
    }
    pub fn window_protection(&self) -> bool {
        self.window_protection
    }
    pub fn show_formulas(&self) -> bool {
        self.show_formulas
    }
    pub fn show_grid_lines(&self) -> bool {
        self.show_grid_lines
    }
    pub fn show_row_col_headers(&self) -> bool {
        self.show_row_col_headers
    }
    pub fn show_zeros(&self) -> bool {
        self.show_zeros
    }
    pub fn right_to_left(&self) -> bool {
        self.right_to_left
    }
    pub fn tab_selected(&self) -> bool {
        self.tab_selected
    }
    pub fn show_ruler(&self) -> bool {
        self.show_ruler
    }
    pub fn show_outline_symbols(&self) -> bool {
        self.show_outline_symbols
    }
    pub fn default_grid_color(&self) -> bool {
        self.default_grid_color
    }
    pub fn show_white_space(&self) -> bool {
        self.show_white_space
    }
    pub fn view_type(&self) -> ViewType {
        self.view_type
    }
    pub fn top_left_cell(&self) -> Option<&CellReference> {
        self.top_left_cell.as_ref()
    }
    pub fn color_id(&self) -> u32 {
        self.color_id
    }
    pub fn zoom_scale(&self) -> u16 {
        self.zoom_scale
    }
    pub fn zoom_scale_normal(&self) -> u16 {
        self.zoom_scale_normal
    }
    pub fn zoom_scale_sheet_layout_view(&self) -> u16 {
        self.zoom_scale_sheet_layout_view
    }
    pub fn zoom_scale_page_layout_view(&self) -> u16 {
        self.zoom_scale_page_layout_view
    }
    pub fn pane(&self) -> Option<&Pane> {
        self.pane.as_ref()
    }
    pub fn selections(&self) -> &[Selection] {
        &self.selections
    }
    pub fn pivot_selections(&self) -> &[PivotSelection] {
        &self.pivot_selections
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Views {
    views: Vec<View>,
    extensions: Vec<Extension>,
}
impl Views {
    pub fn views(&self) -> &[View] {
        &self.views
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Outside,
    Worksheet,
    SheetViews,
    SheetView,
    Leaf,
    PivotSelection,
    PivotArea,
    ExtList(Owner),
    Extension,
}
#[derive(Clone, Copy, PartialEq, Eq)]
enum Owner {
    Collection,
    View,
}
enum CapturePayload {
    PivotArea(PivotArea),
    Extension(Owner, Extension),
}
struct Capture {
    depth: usize,
    bytes: usize,
    writer: Writer<Vec<u8>>,
    payload: CapturePayload,
}

struct Parser {
    stack: Vec<Context>,
    collection: Option<Views>,
    current_view: Option<View>,
    current_pivot: Option<PivotBuilder>,
    capture: Option<Capture>,
    seen_collection: bool,
    sheet_views_phase: u8,
    view_phase: u8,
    retained: usize,
}
struct PivotBuilder {
    pane: PanePosition,
    show_header: bool,
    label: bool,
    data: bool,
    extendable: bool,
    count: u32,
    axis: Option<PivotSelectionAxis>,
    dimension: u32,
    start: u32,
    min: u32,
    max: u32,
    active_row: u32,
    active_column: u32,
    previous_row: u32,
    previous_column: u32,
    click: u32,
    relationship_id: Option<String>,
    area: Option<PivotArea>,
}

/// Parse the worksheet's single `sheetViews` collection without resolving UI state or relationships.
pub fn parse_worksheet_views(xml: &[u8]) -> Result<Option<Views>> {
    let processed =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut parser = Parser {
        stack: Vec::new(),
        collection: None,
        current_view: None,
        current_pivot: None,
        capture: None,
        seen_collection: false,
        sheet_views_phase: 0,
        view_phase: 0,
        retained: 0,
    };
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        if matches!(event, Event::DocType(_) | Event::PI(_)) {
            return Err(invalid("DTD and processing instructions are rejected"));
        }
        if let Event::GeneralRef(reference) = &event {
            let name = reference.decode().map_err(xml_error)?;
            if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot")
                && !name.starts_with('#')
            {
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
            Event::Start(element) => parser.start(&namespace, &element, decoder, &resolver)?,
            Event::Empty(element) => parser.empty(&namespace, &element, decoder, &resolver)?,
            Event::End(element) => parser.end(&namespace, element.local_name().as_ref())?,
            Event::Eof => break,
            _ => {},
        }
    }
    if parser.capture.is_some() || !parser.stack.is_empty() {
        return Err(invalid("unterminated worksheet-view XML"));
    }
    Ok(parser.collection)
}

impl Parser {
    fn parent(&self) -> Context {
        self.stack.last().copied().unwrap_or(Context::Outside)
    }

    fn start(
        &mut self,
        ns: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(ns, element.name(), local.as_ref());
        match (self.parent(), core, local.as_ref()) {
            (Context::Outside, true, b"worksheet") => self.stack.push(Context::Worksheet),
            (Context::Worksheet, true, b"sheetViews") => {
                self.begin_collection()?;
                self.stack.push(Context::SheetViews);
            },
            (Context::SheetViews, true, b"sheetView") => {
                self.begin_view(element, decoder)?;
                self.stack.push(Context::SheetView);
            },
            (Context::SheetViews, true, b"extLst") => {
                self.begin_ext_list(Owner::Collection)?;
                self.stack.push(Context::ExtList(Owner::Collection));
            },
            (Context::SheetView, true, b"pane") => {
                self.add_pane(element, decoder)?;
                self.stack.push(Context::Leaf);
            },
            (Context::SheetView, true, b"selection") => {
                self.add_selection(element, decoder)?;
                self.stack.push(Context::Leaf);
            },
            (Context::SheetView, true, b"pivotSelection") => {
                self.begin_pivot(element, decoder, resolver)?;
                self.stack.push(Context::PivotSelection);
            },
            (Context::SheetView, true, b"extLst") => {
                self.begin_ext_list(Owner::View)?;
                self.stack.push(Context::ExtList(Owner::View));
            },
            (Context::PivotSelection, true, b"pivotArea") => {
                if self
                    .current_pivot
                    .as_ref()
                    .is_some_and(|value| value.area.is_some())
                {
                    return Err(invalid("duplicate pivotArea in pivotSelection"));
                }
                let area = parse_pivot_area(element, decoder)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element.clone()))
                    .map_err(xml_error)?;
                self.capture = Some(Capture {
                    depth: 1,
                    bytes: element.len(),
                    writer,
                    payload: CapturePayload::PivotArea(area),
                });
                self.stack.push(Context::PivotArea);
            },
            (Context::ExtList(owner), true, b"ext") => {
                let extension = parse_extension(element, decoder)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element.clone()))
                    .map_err(xml_error)?;
                self.capture = Some(Capture {
                    depth: 1,
                    bytes: element.len(),
                    writer,
                    payload: CapturePayload::Extension(owner, extension),
                });
                self.stack.push(Context::Extension);
            },
            (Context::Leaf, _, _) => {
                return Err(invalid(
                    "worksheet-view leaf element contains child content",
                ));
            },
            (parent, _, _) if in_view_tree(parent) => {
                return Err(invalid(format!(
                    "unexpected worksheet-view element '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
            _ => self.stack.push(Context::Outside),
        }
        Ok(())
    }

    fn empty(
        &mut self,
        ns: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(ns, element.name(), local.as_ref());
        match (self.parent(), core, local.as_ref()) {
            (Context::Worksheet, true, b"sheetViews") => {
                return Err(invalid("sheetViews must contain at least one sheetView"));
            },
            (Context::SheetViews, true, b"sheetView") => {
                self.begin_view(element, decoder)?;
                self.finish_view()?;
            },
            (Context::SheetViews, true, b"extLst") => self.begin_ext_list(Owner::Collection)?,
            (Context::SheetView, true, b"pane") => self.add_pane(element, decoder)?,
            (Context::SheetView, true, b"selection") => self.add_selection(element, decoder)?,
            (Context::SheetView, true, b"pivotSelection") => {
                return Err(invalid("pivotSelection requires one pivotArea"));
            },
            (Context::SheetView, true, b"extLst") => self.begin_ext_list(Owner::View)?,
            (Context::PivotSelection, true, b"pivotArea") => {
                if self
                    .current_pivot
                    .as_ref()
                    .is_some_and(|value| value.area.is_some())
                {
                    return Err(invalid("duplicate pivotArea in pivotSelection"));
                }
                let mut area = parse_pivot_area(element, decoder)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element.clone()))
                    .map_err(xml_error)?;
                area.markup = writer.into_inner();
                self.retain(area.markup.len(), MAX_PIVOT_AREA_MARKUP)?;
                self.current_pivot
                    .as_mut()
                    .ok_or_else(|| invalid("pivotArea outside pivotSelection"))?
                    .area = Some(area);
            },
            (Context::ExtList(owner), true, b"ext") => {
                let mut extension = parse_extension(element, decoder)?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element.clone()))
                    .map_err(xml_error)?;
                extension.markup = writer.into_inner();
                self.retain(extension.markup.len(), MAX_RETAINED_MARKUP)?;
                self.add_extension(owner, extension)?;
            },
            (Context::Leaf, _, _) => {
                return Err(invalid(
                    "worksheet-view leaf element contains child content",
                ));
            },
            (parent, _, _) if in_view_tree(parent) => {
                return Err(invalid(format!(
                    "unexpected worksheet-view element '{}'",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
            _ => {},
        }
        let _ = resolver;
        Ok(())
    }

    fn end(&mut self, ns: &ResolveResult<'_>, local: &[u8]) -> Result<()> {
        let context = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unexpected worksheet-view end element"))?;
        match context {
            Context::SheetViews if spreadsheet_ns(ns) && local == b"sheetViews" => {
                self.finish_collection()?
            },
            Context::SheetView if spreadsheet_ns(ns) && local == b"sheetView" => {
                self.finish_view()?
            },
            Context::PivotSelection if spreadsheet_ns(ns) && local == b"pivotSelection" => {
                self.finish_pivot()?
            },
            Context::Worksheet | Context::Outside | Context::Leaf | Context::ExtList(_) => {},
            Context::PivotArea
            | Context::Extension
            | Context::SheetViews
            | Context::SheetView
            | Context::PivotSelection => {
                return Err(invalid("mismatched worksheet-view end element"));
            },
        }
        Ok(())
    }

    fn begin_collection(&mut self) -> Result<()> {
        if self.seen_collection {
            return Err(invalid("duplicate worksheet sheetViews element"));
        }
        self.seen_collection = true;
        self.sheet_views_phase = 0;
        self.collection = Some(Views {
            views: Vec::new(),
            extensions: Vec::new(),
        });
        Ok(())
    }
    fn finish_collection(&self) -> Result<()> {
        if self
            .collection
            .as_ref()
            .is_none_or(|value| value.views.is_empty())
        {
            return Err(invalid("sheetViews must contain at least one sheetView"));
        }
        Ok(())
    }
    fn begin_view(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.sheet_views_phase > 0 {
            return Err(invalid("sheetView appears after sheetViews extLst"));
        }
        if self.current_view.is_some() {
            return Err(invalid("nested sheetView"));
        }
        let collection = self
            .collection
            .as_ref()
            .ok_or_else(|| invalid("sheetView outside sheetViews"))?;
        if collection.views.len() >= MAX_VIEWS {
            return Err(invalid("too many worksheet views"));
        }
        self.view_phase = 0;
        self.current_view = Some(parse_view(element, decoder)?);
        Ok(())
    }
    fn finish_view(&mut self) -> Result<()> {
        if self.current_pivot.is_some() {
            return Err(invalid("unterminated pivotSelection"));
        }
        let view = self
            .current_view
            .take()
            .ok_or_else(|| invalid("sheetView end without start"))?;
        self.collection
            .as_mut()
            .ok_or_else(|| invalid("sheetView outside collection"))?
            .views
            .push(view);
        Ok(())
    }
    fn begin_ext_list(&mut self, owner: Owner) -> Result<()> {
        match owner {
            Owner::Collection => {
                if self.sheet_views_phase > 0 {
                    return Err(invalid("duplicate sheetViews extLst"));
                }
                self.sheet_views_phase = 1;
            },
            Owner::View => {
                if self.view_phase > 2 {
                    return Err(invalid("duplicate sheetView extLst"));
                }
                self.view_phase = 3;
            },
        }
        Ok(())
    }
    fn add_pane(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.view_phase != 0 {
            return Err(invalid("pane is out of schema order"));
        }
        let view = self
            .current_view
            .as_mut()
            .ok_or_else(|| invalid("pane outside sheetView"))?;
        if view.pane.is_some() {
            return Err(invalid("duplicate worksheet-view pane"));
        }
        view.pane = Some(parse_pane(element, decoder)?);
        Ok(())
    }
    fn add_selection(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.view_phase > 1 {
            return Err(invalid("selection is out of schema order"));
        }
        self.view_phase = 1;
        let view = self
            .current_view
            .as_mut()
            .ok_or_else(|| invalid("selection outside sheetView"))?;
        if view.selections.len() >= 4 {
            return Err(invalid("sheetView exceeds four selections"));
        }
        view.selections.push(parse_selection(element, decoder)?);
        Ok(())
    }
    fn begin_pivot(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.view_phase > 2 {
            return Err(invalid("pivotSelection is out of schema order"));
        }
        self.view_phase = 2;
        let view = self
            .current_view
            .as_ref()
            .ok_or_else(|| invalid("pivotSelection outside sheetView"))?;
        if view.pivot_selections.len() >= 4 {
            return Err(invalid("sheetView exceeds four pivot selections"));
        }
        if self.current_pivot.is_some() {
            return Err(invalid("nested pivotSelection"));
        }
        self.current_pivot = Some(parse_pivot(element, decoder, resolver)?);
        Ok(())
    }
    fn finish_pivot(&mut self) -> Result<()> {
        let value = self
            .current_pivot
            .take()
            .ok_or_else(|| invalid("pivotSelection end without start"))?;
        let area = value
            .area
            .ok_or_else(|| invalid("pivotSelection requires one pivotArea"))?;
        self.current_view
            .as_mut()
            .ok_or_else(|| invalid("pivotSelection outside sheetView"))?
            .pivot_selections
            .push(PivotSelection {
                pane: value.pane,
                show_header: value.show_header,
                label: value.label,
                data: value.data,
                extendable: value.extendable,
                count: value.count,
                axis: value.axis,
                dimension: value.dimension,
                start: value.start,
                min: value.min,
                max: value.max,
                active_row: value.active_row,
                active_column: value.active_column,
                previous_row: value.previous_row,
                previous_column: value.previous_column,
                click: value.click,
                relationship_id: value.relationship_id,
                area,
            });
        Ok(())
    }
    fn add_extension(&mut self, owner: Owner, extension: Extension) -> Result<()> {
        match owner {
            Owner::Collection => self
                .collection
                .as_mut()
                .ok_or_else(|| invalid("extension outside sheetViews"))?
                .extensions
                .push(extension),
            Owner::View => self
                .current_view
                .as_mut()
                .ok_or_else(|| invalid("extension outside sheetView"))?
                .extensions
                .push(extension),
        }
        Ok(())
    }
    fn retain(&mut self, size: usize, individual_limit: usize) -> Result<()> {
        if size > individual_limit {
            return Err(invalid(
                "worksheet-view retained markup exceeds element limit",
            ));
        }
        self.retained = self
            .retained
            .checked_add(size)
            .ok_or_else(|| invalid("worksheet-view retained-byte overflow"))?;
        if self.retained > MAX_RETAINED_MARKUP {
            return Err(invalid(
                "worksheet-view retained markup exceeds resource limit",
            ));
        }
        Ok(())
    }
    fn capture_event(&mut self, event: Event<'static>) -> Result<()> {
        let Some(capture) = self.capture.as_mut() else {
            return Err(invalid("captured worksheet-view markup is missing"));
        };
        capture.bytes = capture
            .bytes
            .checked_add(event.as_ref().len())
            .ok_or_else(|| invalid("worksheet-view capture size overflow"))?;
        capture
            .writer
            .write_event(event.clone())
            .map_err(xml_error)?;
        match event {
            Event::Start(_) => capture.depth += 1,
            Event::End(_) => capture.depth -= 1,
            Event::Eof => return Err(invalid("unterminated retained worksheet-view markup")),
            _ => {},
        }
        if capture.depth != 0 {
            return Ok(());
        }
        let Some(capture) = self.capture.take() else {
            return Err(invalid("captured worksheet-view markup is missing"));
        };
        let markup = capture.writer.into_inner();
        self.stack.pop();
        match capture.payload {
            CapturePayload::PivotArea(mut area) => {
                area.markup = markup;
                self.retain(area.markup.len(), MAX_PIVOT_AREA_MARKUP)?;
                self.current_pivot
                    .as_mut()
                    .ok_or_else(|| invalid("pivotArea outside pivotSelection"))?
                    .area = Some(area);
            },
            CapturePayload::Extension(owner, mut extension) => {
                extension.markup = markup;
                self.retain(extension.markup.len(), MAX_RETAINED_MARKUP)?;
                self.add_extension(owner, extension)?;
            },
        }
        Ok(())
    }
}

fn parse_view(element: &BytesStart<'_>, decoder: Decoder) -> Result<View> {
    Ok(View {
        workbook_view_id: required_u32(element, b"workbookViewId", decoder)?,
        window_protection: bool_attr(element, b"windowProtection", decoder)?.unwrap_or(false),
        show_formulas: bool_attr(element, b"showFormulas", decoder)?.unwrap_or(false),
        show_grid_lines: bool_attr(element, b"showGridLines", decoder)?.unwrap_or(true),
        show_row_col_headers: bool_attr(element, b"showRowColHeaders", decoder)?.unwrap_or(true),
        show_zeros: bool_attr(element, b"showZeros", decoder)?.unwrap_or(true),
        right_to_left: bool_attr(element, b"rightToLeft", decoder)?.unwrap_or(false),
        tab_selected: bool_attr(element, b"tabSelected", decoder)?.unwrap_or(false),
        show_ruler: bool_attr(element, b"showRuler", decoder)?.unwrap_or(true),
        show_outline_symbols: bool_attr(element, b"showOutlineSymbols", decoder)?.unwrap_or(true),
        default_grid_color: bool_attr(element, b"defaultGridColor", decoder)?.unwrap_or(true),
        show_white_space: bool_attr(element, b"showWhiteSpace", decoder)?.unwrap_or(true),
        view_type: ViewType::parse(
            attr(element, b"view", decoder)?
                .as_deref()
                .unwrap_or("normal"),
        )?,
        top_left_cell: optional_cell(element, b"topLeftCell", decoder)?,
        color_id: optional_u32(element, b"colorId", decoder)?.unwrap_or(64),
        zoom_scale: current_zoom(element, b"zoomScale", decoder)?.unwrap_or(100),
        zoom_scale_normal: remembered_zoom(element, b"zoomScaleNormal", decoder)?.unwrap_or(0),
        zoom_scale_sheet_layout_view: remembered_zoom(
            element,
            b"zoomScaleSheetLayoutView",
            decoder,
        )?
        .unwrap_or(0),
        zoom_scale_page_layout_view: remembered_zoom(element, b"zoomScalePageLayoutView", decoder)?
            .unwrap_or(0),
        pane: None,
        selections: Vec::new(),
        pivot_selections: Vec::new(),
        extensions: Vec::new(),
    })
    .and_then(|view| {
        if view.color_id > 64 {
            Err(invalid("worksheet-view colorId exceeds 64"))
        } else {
            Ok(view)
        }
    })
}
fn parse_pane(element: &BytesStart<'_>, decoder: Decoder) -> Result<Pane> {
    Ok(Pane {
        x_split: nonnegative_f64(element, b"xSplit", decoder)?,
        y_split: nonnegative_f64(element, b"ySplit", decoder)?,
        top_left_cell: optional_cell(element, b"topLeftCell", decoder)?,
        active_pane: PanePosition::parse(
            attr(element, b"activePane", decoder)?
                .as_deref()
                .unwrap_or("topLeft"),
        )?,
        state: PaneState::parse(
            attr(element, b"state", decoder)?
                .as_deref()
                .unwrap_or("split"),
        )?,
    })
}
fn parse_selection(element: &BytesStart<'_>, decoder: Decoder) -> Result<Selection> {
    let active_cell =
        optional_cell(element, b"activeCell", decoder)?.unwrap_or(CellReference("A1".into()));
    let sqref = match attr(element, b"sqref", decoder)? {
        Some(value) => parse_sqref(&value)?,
        None => Sqref(vec![RangeReference("A1".into())]),
    };
    let active_cell_id = optional_u32(element, b"activeCellId", decoder)?.unwrap_or(0);
    if active_cell_id as usize >= sqref.0.len() {
        return Err(invalid("worksheet-view activeCellId is outside sqref"));
    }
    Ok(Selection {
        pane: PanePosition::parse(
            attr(element, b"pane", decoder)?
                .as_deref()
                .unwrap_or("topLeft"),
        )?,
        active_cell,
        active_cell_id,
        sqref,
    })
}
fn parse_pivot(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<PivotBuilder> {
    let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?;
    if let Some(value) = relationship_id.as_deref() {
        validate_relationship_id(value)?;
    }
    Ok(PivotBuilder {
        pane: PanePosition::parse(
            attr(element, b"pane", decoder)?
                .as_deref()
                .unwrap_or("topLeft"),
        )?,
        show_header: bool_attr(element, b"showHeader", decoder)?.unwrap_or(false),
        label: bool_attr(element, b"label", decoder)?.unwrap_or(false),
        data: bool_attr(element, b"data", decoder)?.unwrap_or(false),
        extendable: bool_attr(element, b"extendable", decoder)?.unwrap_or(false),
        count: optional_u32(element, b"count", decoder)?.unwrap_or(0),
        axis: attr(element, b"axis", decoder)?
            .map(|v| PivotSelectionAxis::parse(&v))
            .transpose()?,
        dimension: optional_u32(element, b"dimension", decoder)?.unwrap_or(0),
        start: optional_u32(element, b"start", decoder)?.unwrap_or(0),
        min: optional_u32(element, b"min", decoder)?.unwrap_or(0),
        max: optional_u32(element, b"max", decoder)?.unwrap_or(0),
        active_row: optional_u32(element, b"activeRow", decoder)?.unwrap_or(0),
        active_column: optional_u32(element, b"activeCol", decoder)?.unwrap_or(0),
        previous_row: optional_u32(element, b"previousRow", decoder)?.unwrap_or(0),
        previous_column: optional_u32(element, b"previousCol", decoder)?.unwrap_or(0),
        click: optional_u32(element, b"click", decoder)?.unwrap_or(0),
        relationship_id,
        area: None,
    })
}
fn parse_pivot_area(element: &BytesStart<'_>, decoder: Decoder) -> Result<PivotArea> {
    Ok(PivotArea {
        field: optional_i32(element, b"field", decoder)?,
        area_type: PivotAreaType::parse(
            attr(element, b"type", decoder)?
                .as_deref()
                .unwrap_or("normal"),
        )?,
        data_only: bool_attr(element, b"dataOnly", decoder)?.unwrap_or(true),
        label_only: bool_attr(element, b"labelOnly", decoder)?.unwrap_or(false),
        grand_row: bool_attr(element, b"grandRow", decoder)?.unwrap_or(false),
        grand_column: bool_attr(element, b"grandCol", decoder)?.unwrap_or(false),
        cache_index: bool_attr(element, b"cacheIndex", decoder)?.unwrap_or(false),
        outline: bool_attr(element, b"outline", decoder)?.unwrap_or(true),
        offset: attr(element, b"offset", decoder)?
            .map(|v| parse_range(&v))
            .transpose()?,
        collapsed_levels_are_subtotals: bool_attr(
            element,
            b"collapsedLevelsAreSubtotals",
            decoder,
        )?
        .unwrap_or(false),
        axis: attr(element, b"axis", decoder)?
            .map(|v| PivotSelectionAxis::parse(&v))
            .transpose()?,
        field_position: optional_u32(element, b"fieldPosition", decoder)?,
        markup: Vec::new(),
    })
}
fn parse_extension(element: &BytesStart<'_>, decoder: Decoder) -> Result<Extension> {
    let uri = attr(element, b"uri", decoder)?
        .ok_or_else(|| invalid("worksheet-view ext requires uri"))?;
    if uri.is_empty() || uri.len() > 1024 {
        return Err(invalid("invalid worksheet-view extension URI"));
    }
    Ok(Extension {
        uri,
        markup: Vec::new(),
    })
}

fn attr(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() != name {
            continue;
        }
        if value.is_some() {
            return Err(invalid(format!(
                "duplicate attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned(),
        );
    }
    Ok(value)
}
fn bool_attr(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    attr(element, name, decoder)?
        .map(|v| match v.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
fn optional_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<u32>> {
    attr(element, name, decoder)?
        .map(|v| {
            v.parse::<u32>()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
fn required_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<u32> {
    optional_u32(element, name, decoder)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(name)
        ))
    })
}
fn optional_i32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<i32>> {
    attr(element, name, decoder)?
        .map(|v| {
            v.parse::<i32>()
                .map_err(|_| invalid(format!("invalid integer '{v}'")))
        })
        .transpose()
}
fn nonnegative_f64(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<f64>> {
    attr(element, name, decoder)?
        .map(|v| {
            let n = v
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid split value '{v}'")))?;
            if !n.is_finite() || n < 0.0 {
                Err(invalid(
                    "worksheet-view split must be finite and nonnegative",
                ))
            } else {
                Ok(n)
            }
        })
        .transpose()
}
fn current_zoom(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<u16>> {
    optional_u32(element, name, decoder)?
        .map(|v| {
            if (10..=400).contains(&v) {
                Ok(v as u16)
            } else {
                Err(invalid(
                    "current worksheet-view zoom must be 10 through 400",
                ))
            }
        })
        .transpose()
}
fn remembered_zoom(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<u16>> {
    optional_u32(element, name, decoder)?
        .map(|v| {
            if v == 0 || (10..=400).contains(&v) {
                Ok(v as u16)
            } else {
                Err(invalid(
                    "remembered worksheet-view zoom must be zero or 10 through 400",
                ))
            }
        })
        .transpose()
}
fn optional_cell(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<CellReference>> {
    attr(element, name, decoder)?
        .map(|v| {
            validate_cell(&v)?;
            Ok(CellReference(v))
        })
        .transpose()
}
fn validate_cell(value: &str) -> Result<()> {
    if value == "#REF!" {
        return Err(invalid("worksheet-view references cannot be #REF!"));
    }
    litchi_sheet::Cell::from_a1(value)
        .map(|_| ())
        .map_err(Error::from)
}
fn parse_range(value: &str) -> Result<RangeReference> {
    if value.is_empty() || value.split(':').count() > 2 {
        return Err(invalid(format!("invalid worksheet range '{value}'")));
    }
    let mut parts = value.split(':');
    let Some(start) = parts.next() else {
        return Err(invalid(format!("invalid worksheet range '{value}'")));
    };
    validate_cell(start)?;
    if let Some(end) = parts.next() {
        validate_cell(end)?;
    }
    Ok(RangeReference(value.into()))
}
fn parse_sqref(value: &str) -> Result<Sqref> {
    let mut ranges = Vec::new();
    for token in value.split_whitespace() {
        if ranges.len() >= MAX_SELECTION_RANGES {
            return Err(invalid("worksheet-view sqref exceeds reference limit"));
        }
        ranges.push(parse_range(token)?);
    }
    if ranges.is_empty() {
        return Err(invalid("worksheet-view sqref is empty"));
    }
    Ok(Sqref(ranges))
}
fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > MAX_RELATIONSHIP_ID_BYTES
        || value.trim() != value
        || value.chars().any(char::is_control)
    {
        Err(invalid("invalid pivot-selection relationship ID"))
    } else {
        Ok(())
    }
}
fn spreadsheet_ns(value: &ResolveResult<'_>) -> bool {
    matches!(value,ResolveResult::Bound(namespace) if namespace.as_ref()==b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"||namespace.as_ref()==b"http://purl.oclc.org/ooxml/spreadsheetml/main")
}
fn in_view_tree(value: Context) -> bool {
    matches!(
        value,
        Context::SheetViews
            | Context::SheetView
            | Context::Leaf
            | Context::PivotSelection
            | Context::PivotArea
            | Context::ExtList(_)
            | Context::Extension
    )
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet-view XML: {error}"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};
    const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    fn fixture(bytes: &[u8]) -> Views {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        parse_worksheet_views(part.blob()).unwrap().unwrap()
    }

    #[test]
    fn reads_poi_and_libreoffice_view_fixtures() {
        let poi = fixture(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx"
        ));
        assert_eq!(poi.views()[0].selections()[0].active_cell().as_str(), "A4");
        let lo = fixture(include_bytes!(
            "../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/freezePaneStartCell.xlsx"
        ));
        let view = &lo.views()[0];
        assert_eq!(view.selections().len(), 4);
        assert_eq!(view.pane().unwrap().x_split(), Some(5.0));
        assert_eq!(view.pane().unwrap().state(), PaneState::Frozen);
    }

    #[test]
    fn reads_strict_mce_pivot_selection_and_defaults() {
        let xml = format!(
            r#"<s:worksheet xmlns:s="{S}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><u:no/></mc:Choice><mc:Fallback><s:sheetViews><s:sheetView workbookViewId="7" colorId="64" zoomScale="125" zoomScaleNormal="0"><s:pane xSplit="2" topLeftCell="C1" state="frozen"/><s:selection activeCell="C1" sqref="C1:D2 F3"/><s:pivotSelection pane="bottomRight" showHeader="1" axis="axisRow" dimension="2" activeRow="11" activeCol="1" r:id="rId1"><s:pivotArea dataOnly="0" labelOnly="1" fieldPosition="0" offset="A1:B2"><s:references count="1"><s:reference field="9" count="0"/></s:references></s:pivotArea></s:pivotSelection><s:extLst><s:ext uri="urn:view"><u:ignored/></s:ext></s:extLst></s:sheetView></s:sheetViews></mc:Fallback></mc:AlternateContent></s:worksheet>"#
        );
        let collection = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
        let view = &collection.views()[0];
        assert!(view.show_grid_lines());
        assert_eq!(view.view_type(), ViewType::Normal);
        assert_eq!(view.zoom_scale_normal(), 0);
        assert_eq!(view.selections()[0].sqref().ranges().len(), 2);
        let pivot = &view.pivot_selections()[0];
        assert_eq!(pivot.axis(), Some(PivotSelectionAxis::Row));
        assert_eq!(pivot.relationship_id(), Some("rId1"));
        assert_eq!(pivot.area().offset().unwrap().as_str(), "A1:B2");
        assert!(!pivot.area().markup().is_empty());
        assert_eq!(view.extensions()[0].uri(), "urn:view");
    }

    #[test]
    fn accepts_automatic_remembered_zoom_defaults() {
        let xml = format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" zoomScale="100" zoomScaleNormal="0" zoomScaleSheetLayoutView="0" zoomScalePageLayoutView="0"/></sheetViews></worksheet>"#
        );
        let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
        let view = &parsed.views()[0];
        assert_eq!(view.zoom_scale(), 100);
        assert_eq!(view.zoom_scale_normal(), 0);
        assert_eq!(view.zoom_scale_sheet_layout_view(), 0);
        assert_eq!(view.zoom_scale_page_layout_view(), 0);
    }

    #[test]
    fn accepts_spec_absolute_a1_references() {
        let xml = format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" topLeftCell="$XFD$1048576"><selection activeCell="$A$1" sqref="$A$1:$XFD$1048576"/></sheetView></sheetViews></worksheet>"#
        );
        let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
        let view = &parsed.views()[0];
        assert_eq!(view.top_left_cell().unwrap().as_str(), "$XFD$1048576");
        assert_eq!(view.selections()[0].active_cell().as_str(), "$A$1");
    }

    #[test]
    fn rejects_invalid_view_grammar_values_and_security() {
        let cases = [
            format!(
                r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" view="bad"/></sheetViews></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><selection activeCell="XFE1"/></sheetView></sheetViews></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCellId="1"/></sheetView></sheetViews></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><pivotSelection axis="bad"><pivotArea/></pivotSelection></sheetView></sheetViews></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><pivotSelection/></sheetView></sheetViews></worksheet>"#
            ),
            format!(r#"<!DOCTYPE x><worksheet xmlns="{T}"/>"#),
        ];
        for xml in cases {
            assert!(parse_worksheet_views(xml.as_bytes()).is_err(), "{xml}");
        }
    }
}

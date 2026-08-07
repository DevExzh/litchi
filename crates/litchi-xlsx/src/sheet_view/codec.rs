//! Bounded SpreadsheetML worksheet-view codec.

use crate::error::{Error, Result};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_sheet::view::{
    Color, Display, Mode, Pane, Position, Scale, Selection, Split, State, View, Window,
};
use litchi_sheet::{Cell, Rect};
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::invalid;
use super::model::*;

const MAX_VIEWS: usize = 1024;
const MAX_SELECTION_RANGES: usize = 32_767;
const MAX_RETAINED_MARKUP: usize = 4 * 1024 * 1024;
const MAX_PIVOT_AREA_MARKUP: usize = 1024 * 1024;
const MAX_RETAINED_SHEET_VIEW_XML: usize = 4 * 1024 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;

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
struct SheetViewCapture {
    bytes: usize,
    writer: Writer<Vec<u8>>,
}

struct Parser {
    stack: Vec<Context>,
    collection: Option<Collection>,
    current_view: Option<Entry>,
    current_pivot: Option<PivotBuilder>,
    capture: Option<Capture>,
    sheet_view_capture: Option<SheetViewCapture>,
    seen_collection: bool,
    sheet_views_phase: u8,
    view_phase: u8,
    retained: usize,
}
struct PivotBuilder {
    pane: Position,
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
pub fn parse_worksheet_views(xml: &[u8]) -> Result<Option<Collection>> {
    let processed =
        process_markup_compatibility(xml, &Capabilities::default(), &Limits::default())?;
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut parser = Parser {
        stack: Vec::new(),
        collection: None,
        current_view: None,
        current_pivot: None,
        capture: None,
        sheet_view_capture: None,
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
        if parser.sheet_view_capture.is_some() && !matches!(&event, Event::Eof) {
            parser.capture_sheet_view_event(event.clone())?;
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
                self.begin_sheet_view_capture(element, false)?;
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
                self.begin_sheet_view_capture(element, true)?;
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
        self.collection = Some(Collection {
            entries: Vec::new(),
            extensions: Vec::new(),
        });
        Ok(())
    }
    fn finish_collection(&self) -> Result<()> {
        if self
            .collection
            .as_ref()
            .is_none_or(|value| value.entries.is_empty())
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
        if collection.entries.len() >= MAX_VIEWS {
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
        let retained_xml = self
            .sheet_view_capture
            .take()
            .ok_or_else(|| invalid("sheetView retained markup is missing"))?
            .writer
            .into_inner();
        let mut view = self
            .current_view
            .take()
            .ok_or_else(|| invalid("sheetView end without start"))?;
        view.retained_xml = retained_xml;
        self.collection
            .as_mut()
            .ok_or_else(|| invalid("sheetView outside collection"))?
            .entries
            .push(view);
        Ok(())
    }
    fn begin_sheet_view_capture(&mut self, element: &BytesStart<'_>, empty: bool) -> Result<()> {
        if self.sheet_view_capture.is_some() {
            return Err(invalid("nested sheetView retained markup"));
        }
        if element.len() > MAX_RETAINED_SHEET_VIEW_XML {
            return Err(invalid(
                "worksheet-view retained sheetView XML exceeds resource limit",
            ));
        }
        let mut writer = Writer::new(Vec::new());
        if empty {
            writer
                .write_event(Event::Empty(element.clone()))
                .map_err(xml_error)?;
        } else {
            writer
                .write_event(Event::Start(element.clone()))
                .map_err(xml_error)?;
        }
        self.sheet_view_capture = Some(SheetViewCapture {
            bytes: element.len(),
            writer,
        });
        Ok(())
    }
    fn capture_sheet_view_event(&mut self, event: Event<'static>) -> Result<()> {
        let capture = self
            .sheet_view_capture
            .as_mut()
            .ok_or_else(|| invalid("sheetView retained markup is missing"))?;
        capture.bytes = capture
            .bytes
            .checked_add(event.as_ref().len())
            .ok_or_else(|| invalid("worksheet-view retained sheetView XML size overflow"))?;
        if capture.bytes > MAX_RETAINED_SHEET_VIEW_XML {
            return Err(invalid(
                "worksheet-view retained sheetView XML exceeds resource limit",
            ));
        }
        capture.writer.write_event(event).map_err(xml_error)?;
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
        if view.view.pane.is_some() {
            return Err(invalid("duplicate worksheet-view pane"));
        }
        view.view.pane = Some(parse_pane(element, decoder)?);
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
        if view.view.selections.len() >= 4 {
            return Err(invalid("sheetView exceeds four selections"));
        }
        view.view
            .selections
            .push(parse_selection(element, decoder)?);
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

fn parse_view(element: &BytesStart<'_>, decoder: Decoder) -> Result<Entry> {
    let color_id = optional_u32(element, b"colorId", decoder)?.unwrap_or(64);
    if color_id > 64 {
        return Err(invalid("worksheet-view colorId exceeds 64"));
    }

    let mut view = View::default();
    view.window = Window::new(required_u32(element, b"workbookViewId", decoder)?);
    view.display = Display {
        window_protection: bool_attr(element, b"windowProtection", decoder)?.unwrap_or(false),
        show_formulas: bool_attr(element, b"showFormulas", decoder)?.unwrap_or(false),
        grid_lines: bool_attr(element, b"showGridLines", decoder)?.unwrap_or(true),
        row_column_headers: bool_attr(element, b"showRowColHeaders", decoder)?.unwrap_or(true),
        zero_values: bool_attr(element, b"showZeros", decoder)?.unwrap_or(true),
        right_to_left: bool_attr(element, b"rightToLeft", decoder)?.unwrap_or(false),
        ruler: bool_attr(element, b"showRuler", decoder)?.unwrap_or(true),
        outline_symbols: bool_attr(element, b"showOutlineSymbols", decoder)?.unwrap_or(true),
        default_grid_color: bool_attr(element, b"defaultGridColor", decoder)?.unwrap_or(true),
        white_space: bool_attr(element, b"showWhiteSpace", decoder)?.unwrap_or(true),
    };
    view.tab_selected = bool_attr(element, b"tabSelected", decoder)?.unwrap_or(false);
    view.mode = parse_mode(
        attr(element, b"view", decoder)?
            .as_deref()
            .unwrap_or("normal"),
    )?;
    if let Some(origin) = optional_cell(element, b"topLeftCell", decoder)? {
        view.origin = origin;
    }
    view.color =
        Color::new(color_id as u8).map_err(|_| invalid("worksheet-view colorId exceeds 64"))?;
    view.zoom.current = current_zoom(element, b"zoomScale", decoder)?.unwrap_or_default();
    view.zoom.normal = remembered_zoom(element, b"zoomScaleNormal", decoder)?;
    view.zoom.page_break_preview = remembered_zoom(element, b"zoomScaleSheetLayoutView", decoder)?;
    view.zoom.page_layout = remembered_zoom(element, b"zoomScalePageLayoutView", decoder)?;
    view.selections.clear();

    Ok(Entry {
        view,
        pivot_selections: Vec::new(),
        extensions: Vec::new(),
        retained_xml: Vec::new(),
    })
}
fn parse_pane(element: &BytesStart<'_>, decoder: Decoder) -> Result<Pane> {
    let mut pane = Pane::default();
    pane.horizontal = nonnegative_f64(element, b"xSplit", decoder)?;
    pane.vertical = nonnegative_f64(element, b"ySplit", decoder)?;
    if let Some(top_left) = optional_cell(element, b"topLeftCell", decoder)? {
        pane.top_left = top_left;
    }
    pane.position = parse_position(
        attr(element, b"activePane", decoder)?
            .as_deref()
            .unwrap_or("topLeft"),
    )?;
    pane.state = parse_state(
        attr(element, b"state", decoder)?
            .as_deref()
            .unwrap_or("split"),
    )?;
    Ok(pane)
}
fn parse_selection(element: &BytesStart<'_>, decoder: Decoder) -> Result<Selection> {
    let defaults = Selection::default();
    let active_cell =
        optional_cell(element, b"activeCell", decoder)?.unwrap_or(defaults.active_cell());
    let ranges = match attr(element, b"sqref", decoder)? {
        Some(value) => parse_sqref(&value)?,
        None => defaults.ranges().to_vec(),
    };
    let active_range = optional_u32(element, b"activeCellId", decoder)?.unwrap_or(0) as usize;
    if active_range >= ranges.len() {
        return Err(invalid("worksheet-view activeCellId is outside sqref"));
    }
    Selection::new(
        parse_position(
            attr(element, b"pane", decoder)?
                .as_deref()
                .unwrap_or("topLeft"),
        )?,
        active_cell,
        active_range,
        ranges,
    )
    .map_err(|_| invalid("invalid worksheet-view selection"))
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
        pane: parse_position(
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
fn nonnegative_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<Split>> {
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
                Split::new(n)
                    .map_err(|_| invalid("worksheet-view split must be finite and nonnegative"))
            }
        })
        .transpose()
}
fn current_zoom(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<Scale>> {
    optional_u32(element, name, decoder)?
        .map(|v| {
            if (10..=400).contains(&v) {
                Scale::new(v as u16)
                    .map_err(|_| invalid("current worksheet-view zoom must be 10 through 400"))
            } else {
                Err(invalid(
                    "current worksheet-view zoom must be 10 through 400",
                ))
            }
        })
        .transpose()
}
fn remembered_zoom(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<Scale>> {
    match optional_u32(element, name, decoder)? {
        None | Some(0) => Ok(None),
        Some(value) if (10..=400).contains(&value) => Scale::new(value as u16)
            .map(Some)
            .map_err(|_| invalid("remembered worksheet-view zoom must be zero or 10 through 400")),
        Some(_) => Err(invalid(
            "remembered worksheet-view zoom must be zero or 10 through 400",
        )),
    }
}
fn optional_cell(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<Cell>> {
    attr(element, name, decoder)?
        .map(|v| parse_cell(&v))
        .transpose()
}
fn parse_cell(value: &str) -> Result<Cell> {
    if value == "#REF!" {
        return Err(invalid("worksheet-view references cannot be #REF!"));
    }
    Cell::from_a1(value).map_err(Error::from)
}
fn parse_range(value: &str) -> Result<Rect> {
    if value.is_empty() || value.split(':').count() > 2 {
        return Err(invalid(format!("invalid worksheet range '{value}'")));
    }
    if value.split(':').any(|part| part == "#REF!") {
        return Err(invalid("worksheet-view references cannot be #REF!"));
    }
    Rect::from_a1(value).map_err(Error::from)
}
fn parse_sqref(value: &str) -> Result<Vec<Rect>> {
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
    Ok(ranges)
}
fn parse_mode(value: &str) -> Result<Mode> {
    match value {
        "normal" => Ok(Mode::Normal),
        "pageBreakPreview" => Ok(Mode::PageBreakPreview),
        "pageLayout" => Ok(Mode::PageLayout),
        _ => Err(invalid(format!("invalid worksheet-view type '{value}'"))),
    }
}
fn parse_position(value: &str) -> Result<Position> {
    match value {
        "bottomRight" => Ok(Position::BottomRight),
        "topRight" => Ok(Position::TopRight),
        "bottomLeft" => Ok(Position::BottomLeft),
        "topLeft" => Ok(Position::TopLeft),
        _ => Err(invalid(format!("invalid worksheet-view pane '{value}'"))),
    }
}
fn parse_state(value: &str) -> Result<State> {
    match value {
        "split" => Ok(State::Split),
        "frozen" => Ok(State::Frozen),
        "frozenSplit" => Ok(State::FrozenSplit),
        _ => Err(invalid(format!(
            "invalid worksheet-view pane state '{value}'"
        ))),
    }
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
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet-view XML: {error}"
    )))
}

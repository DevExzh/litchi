//! Namespace-aware `SpreadsheetDrawing` reader.
//!
//! The inventory keeps the established chart-style cell anchor projection and
//! also records the complete `twoCellAnchor`,
//! `oneCellAnchor`, or `absoluteAnchor` geometry. The latter is the source of
//! truth for callers that need to select or edit an ordinary worksheet
//! picture.

use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use litchi_spreadsheet_drawing::shape::{
    Anchor as DrawingAnchor, CellMarker, EditAs, Emu, EmuExtent, EmuOffset,
};

use super::Anchor;
use super::model::{Chart, Drawing, Object, Picture, Unknown, UnknownKind};
use super::source::{MAX_RELATIONSHIP_ID_BYTES, RelationshipDialect};
use crate::error::{Error, Result, allocation};
use crate::raw::namespace::relationship_attribute_value;
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_chart_name, is_drawingml_name, is_ncname,
    unqualified_attribute_value, xsd_token_atom,
};

const SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_DRAWING_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWING_ANCHORS: usize = 100_000;
const MAX_DRAWING_DEPTH: usize = 256;
const MAX_DRAWING_EVENTS: usize = 1_000_000;
const MAX_MARKER_TEXT_BYTES: usize = 64;
const MAX_STRING_BYTES: usize = 1024 * 1024;

// ECMA-376 DrawingML ST_Coordinate and ST_PositiveCoordinate bounds. The
// parser intentionally accepts only the numeric lexical form here; callers
// that need universal-measure lexical forms use the shape owner.
const MIN_COORDINATE: i64 = -27_273_042_329_600;
const MAX_COORDINATE: i64 = 27_273_042_316_900;

#[derive(Clone, Copy, PartialEq, Eq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    Anchor(AnchorKind),
    From(MarkerTarget),
    To(MarkerTarget),
    Marker(MarkerTarget, MarkerField),
    Position,
    Extent,
    Picture,
    PictureNvPr,
    PictureBlipFill,
    ChartFrame,
    ChartGraphic,
    ChartGraphicData,
    UnknownObject,
    UnknownNvPr,
    ContentPart,
    Ignored,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum MarkerTarget {
    From,
    To,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum MarkerField {
    Column,
    ColumnOffset,
    Row,
    RowOffset,
}

impl MarkerField {
    const fn ordinal(self) -> u8 {
        match self {
            Self::Column => 0,
            Self::ColumnOffset => 1,
            Self::Row => 2,
            Self::RowOffset => 3,
        }
    }
}

#[derive(Default)]
struct Marker {
    column: Option<u32>,
    column_offset: Option<i64>,
    row: Option<u32>,
    row_offset: Option<i64>,
    next_field: u8,
}

impl Marker {
    fn finish(&self, description: &str) -> Result<CellMarker> {
        let column_offset = self
            .column_offset
            .ok_or_else(|| invalid(format!("{description} is missing its column offset")))?;
        let row_offset = self
            .row_offset
            .ok_or_else(|| invalid(format!("{description} is missing its row offset")))?;
        check_coordinate(column_offset, "drawing column offset")?;
        check_coordinate(row_offset, "drawing row offset")?;
        Ok(CellMarker {
            column: self
                .column
                .ok_or_else(|| invalid(format!("{description} is missing its column")))?,
            column_offset: Emu(column_offset),
            row: self
                .row
                .ok_or_else(|| invalid(format!("{description} is missing its row")))?,
            row_offset: Emu(row_offset),
        })
    }

    fn set(&mut self, field: MarkerField, value: i64, description: &str) -> Result<()> {
        if field.ordinal() != self.next_field {
            return Err(invalid(format!("{description} fields are out of order")));
        }
        self.next_field = self
            .next_field
            .checked_add(1)
            .ok_or_else(|| limit("drawing marker fields"))?;
        match field {
            MarkerField::Column => {
                if value < 0 || value > u32::MAX as i64 {
                    return Err(invalid("drawing column is outside its numeric range"));
                }
                set_once(&mut self.column, value as u32, "drawing column")
            },
            MarkerField::ColumnOffset => {
                set_once(&mut self.column_offset, value, "drawing column offset")
            },
            MarkerField::Row => {
                if value < 0 || value > u32::MAX as i64 {
                    return Err(invalid("drawing row is outside its numeric range"));
                }
                set_once(&mut self.row, value as u32, "drawing row")
            },
            MarkerField::RowOffset => set_once(&mut self.row_offset, value, "drawing row offset"),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ObjectKind {
    Picture,
    ChartFrame,
    Unknown(UnknownKind),
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum AnchorChild {
    From,
    To,
    Position,
    Extent,
    Object,
    ClientData,
}

struct PendingAnchor {
    kind: AnchorKind,
    edit_as: EditAs,
    from: Option<Marker>,
    to: Option<Marker>,
    position: Option<EmuOffset>,
    extent: Option<EmuExtent>,
    object_kind: Option<ObjectKind>,
    picture_relationship_id: Option<String>,
    chart_relationship_id: Option<String>,
    content_part_relationship_id: Option<String>,
    description: Option<String>,
    phase: u8,
    client_data_seen: bool,
}

impl PendingAnchor {
    fn new(kind: AnchorKind, edit_as: EditAs) -> Self {
        Self {
            kind,
            edit_as,
            from: None,
            to: None,
            position: None,
            extent: None,
            object_kind: None,
            picture_relationship_id: None,
            chart_relationship_id: None,
            content_part_relationship_id: None,
            description: None,
            phase: 0,
            client_data_seen: false,
        }
    }

    fn take_child(&mut self, child: AnchorChild) -> Result<()> {
        let expected = match (self.kind, self.phase) {
            (AnchorKind::TwoCell, 0) => AnchorChild::From,
            (AnchorKind::TwoCell, 1) => AnchorChild::To,
            (AnchorKind::TwoCell, 2) => AnchorChild::Object,
            (AnchorKind::TwoCell, 3) => AnchorChild::ClientData,
            (AnchorKind::OneCell, 0) => AnchorChild::From,
            (AnchorKind::OneCell, 1) => AnchorChild::Extent,
            (AnchorKind::OneCell, 2) => AnchorChild::Object,
            (AnchorKind::OneCell, 3) => AnchorChild::ClientData,
            (AnchorKind::Absolute, 0) => AnchorChild::Position,
            (AnchorKind::Absolute, 1) => AnchorChild::Extent,
            (AnchorKind::Absolute, 2) => AnchorChild::Object,
            (AnchorKind::Absolute, 3) => AnchorChild::ClientData,
            _ => return Err(invalid("drawing anchor has duplicate or trailing children")),
        };
        if expected != child {
            return Err(invalid("drawing anchor children are out of order"));
        }
        self.phase = self
            .phase
            .checked_add(1)
            .ok_or_else(|| limit("drawing anchor children"))?;
        Ok(())
    }
}

struct Parser {
    drawing: Drawing,
    anchor: Option<PendingAnchor>,
    marker_text: String,
    relationship_dialect: Option<RelationshipDialect>,
    max_depth: usize,
    max_events: usize,
}

impl Parser {
    fn parse(xml: &str, max_depth: usize, max_events: usize) -> Result<Option<Drawing>> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut parser = Self {
            drawing: Drawing::default(),
            anchor: None,
            marker_text: String::new(),
            relationship_dialect: None,
            max_depth,
            max_events,
        };
        let mut stack = Vec::new();
        let mut closed_root = false;
        let mut events = 0usize;
        loop {
            events = events
                .checked_add(1)
                .ok_or_else(|| limit("drawing XML event count"))?;
            if events > parser.max_events {
                return Err(limit("drawing XML event count"));
            }
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Invalid(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    check_context_depth(stack.len(), parser.max_depth)?;
                    if closed_root {
                        return Err(invalid("drawing XML contains multiple root elements"));
                    }
                    if !is_spreadsheet_drawing_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    parser.relationship_dialect = Some(drawing_relationship_dialect(&namespace)?);
                    push_context(&mut stack, Context::Root, parser.max_depth)?;
                },
                Event::Empty(element) if stack.is_empty() => {
                    check_context_depth(stack.len(), parser.max_depth)?;
                    if closed_root {
                        return Err(invalid("drawing XML contains multiple root elements"));
                    }
                    if !is_spreadsheet_drawing_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    return Ok(Some(parser.drawing));
                },
                Event::Start(element) => {
                    check_context_depth(stack.len(), parser.max_depth)?;
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context = parser.start(parent, &namespace, &element, decoder, resolver)?;
                    push_context(&mut stack, context, parser.max_depth)?;
                },
                Event::Empty(element) => {
                    check_context_depth(stack.len(), parser.max_depth)?;
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context = parser.start(parent, &namespace, &element, decoder, resolver)?;
                    parser.finish(context)?;
                },
                Event::Text(text) if matches!(stack.last(), Some(Context::Marker(..))) => {
                    if text.as_ref().len() > MAX_MARKER_TEXT_BYTES {
                        return Err(limit("drawing marker text"));
                    }
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    parser.append_marker_text(&text)?;
                },
                Event::Text(text) if matches!(stack.last(), Some(Context::ContentPart)) => {
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    if !text.chars().all(char::is_whitespace) {
                        return Err(invalid(
                            "core SpreadsheetDrawing contentPart must be childless",
                        ));
                    }
                },
                Event::CData(text) if matches!(stack.last(), Some(Context::Marker(..))) => {
                    if text.as_ref().len() > MAX_MARKER_TEXT_BYTES {
                        return Err(limit("drawing marker text"));
                    }
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    parser.append_marker_text(&text)?;
                },
                Event::CData(text) if matches!(stack.last(), Some(Context::ContentPart)) => {
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    if !text.chars().all(char::is_whitespace) {
                        return Err(invalid(
                            "core SpreadsheetDrawing contentPart must be childless",
                        ));
                    }
                },
                Event::GeneralRef(reference)
                    if matches!(stack.last(), Some(Context::Marker(..))) =>
                {
                    let text = decode_xml_reference(&reference)?;
                    parser.append_marker_text(&text)?;
                },
                Event::GeneralRef(reference)
                    if matches!(stack.last(), Some(Context::ContentPart)) =>
                {
                    let text = decode_xml_reference(&reference)?;
                    if !text.chars().all(char::is_whitespace) {
                        return Err(invalid(
                            "core SpreadsheetDrawing contentPart must be childless",
                        ));
                    }
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or_else(|| invalid("drawing XML closes outside its root"))?;
                    parser.finish(context)?;
                    if context == Context::Root {
                        if !is_spreadsheet_drawing_name(&namespace, element.name(), b"wsDr") {
                            return Err(invalid("drawing XML has an invalid root closing element"));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid("drawing XML has an unterminated root"));
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
        }
        Ok(Some(parser.drawing))
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        check_attribute_lengths(element, decoder)?;
        if parent == Context::Root {
            if !is_spreadsheet_drawing_namespace(namespace) {
                return Ok(Context::Ignored);
            }
            let kind = match element.name().local_name().as_ref() {
                b"twoCellAnchor" => AnchorKind::TwoCell,
                b"oneCellAnchor" => AnchorKind::OneCell,
                b"absoluteAnchor" => AnchorKind::Absolute,
                _ => return Ok(Context::Ignored),
            };
            if self.anchor.is_some() {
                return Err(invalid("nested drawing anchor"));
            }
            if self.drawing.len() >= MAX_DRAWING_ANCHORS {
                return Err(limit("anchors"));
            }
            self.drawing.reserve_object()?;
            let edit_as = if kind == AnchorKind::TwoCell {
                unqualified_attribute_value(element, b"editAs", decoder)?
                    .as_deref()
                    .map_or(Ok(EditAs::TwoCell), |value| {
                        xsd_token_atom(value)
                            .ok_or_else(|| invalid("drawing editAs is not one token"))?
                            .parse()
                            .map_err(|_| invalid("invalid drawing editAs"))
                    })?
            } else {
                EditAs::TwoCell
            };
            self.anchor = Some(PendingAnchor::new(kind, edit_as));
            return Ok(Context::Anchor(kind));
        }

        match parent {
            Context::Anchor(kind) => {
                self.start_anchor_child(kind, namespace, element, decoder, resolver)
            },
            Context::From(target) | Context::To(target) => {
                self.start_marker(target, namespace, element)
            },
            Context::Picture => {
                if is_spreadsheet_drawing_name(namespace, element.name(), b"nvPicPr") {
                    return Ok(Context::PictureNvPr);
                }
                if is_spreadsheet_drawing_name(namespace, element.name(), b"blipFill") {
                    return Ok(Context::PictureBlipFill);
                }
                Ok(Context::Ignored)
            },
            Context::PictureNvPr => {
                if is_spreadsheet_drawing_name(namespace, element.name(), b"cNvPr") {
                    self.capture_description(element, decoder)?;
                    return Ok(Context::Ignored);
                }
                Ok(Context::Ignored)
            },
            Context::PictureBlipFill => {
                if is_drawingml_name(namespace, element.name(), b"blip") {
                    let relationship_id =
                        relationship_attribute_value(element, b"embed", decoder, resolver)?
                            .ok_or_else(|| {
                                invalid("drawing picture blip is missing an embed relationship")
                            })?;
                    set_relationship(
                        &mut self.anchor_mut()?.picture_relationship_id,
                        relationship_id,
                        "picture",
                    )?;
                }
                Ok(Context::Ignored)
            },
            Context::ChartFrame => {
                if is_drawingml_name(namespace, element.name(), b"graphic") {
                    return Ok(Context::ChartGraphic);
                }
                Ok(Context::Ignored)
            },
            Context::ChartGraphic => {
                if is_drawingml_name(namespace, element.name(), b"graphicData") {
                    return Ok(Context::ChartGraphicData);
                }
                Ok(Context::Ignored)
            },
            Context::ChartGraphicData => {
                if is_drawingml_chart_name(namespace, element.name(), b"chart") {
                    let relationship_id =
                        relationship_attribute_value(element, b"id", decoder, resolver)?
                            .ok_or_else(|| invalid("drawing chart is missing a relationship ID"))?;
                    set_relationship(
                        &mut self.anchor_mut()?.chart_relationship_id,
                        relationship_id,
                        "chart",
                    )?;
                }
                Ok(Context::Ignored)
            },
            Context::UnknownObject => {
                if is_spreadsheet_drawing_name(namespace, element.name(), b"nvSpPr")
                    || is_spreadsheet_drawing_name(namespace, element.name(), b"nvGrpSpPr")
                    || is_spreadsheet_drawing_name(namespace, element.name(), b"nvCxnSpPr")
                {
                    return Ok(Context::UnknownNvPr);
                }
                Ok(Context::Ignored)
            },
            Context::ContentPart => Err(invalid(
                "core SpreadsheetDrawing contentPart must be childless",
            )),
            Context::UnknownNvPr => {
                if is_spreadsheet_drawing_name(namespace, element.name(), b"cNvPr") {
                    self.capture_description(element, decoder)?;
                }
                Ok(Context::Ignored)
            },
            Context::Marker(..)
            | Context::Position
            | Context::Extent
            | Context::Ignored
            | Context::Root => Ok(Context::Ignored),
        }
    }

    fn start_anchor_child(
        &mut self,
        kind: AnchorKind,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        let local = element.name().local_name();
        if !is_spreadsheet_drawing_namespace(namespace) {
            return Ok(Context::Ignored);
        }
        match local.as_ref() {
            b"from" => {
                self.take_anchor_child(AnchorChild::From)?;
                if self.anchor_mut()?.from.replace(Marker::default()).is_some() {
                    return Err(invalid("drawing anchor has duplicate from markers"));
                }
                Ok(Context::From(MarkerTarget::From))
            },
            b"to" => {
                self.take_anchor_child(AnchorChild::To)?;
                if kind != AnchorKind::TwoCell {
                    return Err(invalid("one-cell or absolute anchor has a to marker"));
                }
                if self.anchor_mut()?.to.replace(Marker::default()).is_some() {
                    return Err(invalid("drawing anchor has duplicate to markers"));
                }
                Ok(Context::To(MarkerTarget::To))
            },
            b"pos" => {
                self.take_anchor_child(AnchorChild::Position)?;
                if kind != AnchorKind::Absolute {
                    return Err(invalid("only an absolute anchor may contain pos"));
                }
                let position = EmuOffset {
                    x: Emu(coordinate_attribute(
                        element,
                        b"x",
                        decoder,
                        "drawing position x",
                    )?),
                    y: Emu(coordinate_attribute(
                        element,
                        b"y",
                        decoder,
                        "drawing position y",
                    )?),
                };
                if self.anchor_mut()?.position.replace(position).is_some() {
                    return Err(invalid("drawing anchor has duplicate positions"));
                }
                Ok(Context::Position)
            },
            b"ext" => {
                self.take_anchor_child(AnchorChild::Extent)?;
                if kind == AnchorKind::TwoCell {
                    return Err(invalid("a two-cell anchor cannot contain ext"));
                }
                let extent = EmuExtent {
                    width: positive_coordinate_attribute(
                        element,
                        b"cx",
                        decoder,
                        "drawing extent width",
                    )?,
                    height: positive_coordinate_attribute(
                        element,
                        b"cy",
                        decoder,
                        "drawing extent height",
                    )?,
                };
                if self.anchor_mut()?.extent.replace(extent).is_some() {
                    return Err(invalid("drawing anchor has duplicate extents"));
                }
                Ok(Context::Extent)
            },
            b"clientData" => {
                self.take_anchor_child(AnchorChild::ClientData)?;
                let anchor = self.anchor_mut()?;
                if anchor.client_data_seen {
                    return Err(invalid("drawing anchor has duplicate clientData"));
                }
                anchor.client_data_seen = true;
                Ok(Context::Ignored)
            },
            b"pic" => {
                self.take_anchor_child(AnchorChild::Object)?;
                self.open_object(ObjectKind::Picture, Context::Picture)
            },
            b"graphicFrame" => {
                self.take_anchor_child(AnchorChild::Object)?;
                self.open_object(ObjectKind::ChartFrame, Context::ChartFrame)
            },
            b"sp" => {
                self.take_anchor_child(AnchorChild::Object)?;
                self.open_object(
                    ObjectKind::Unknown(UnknownKind::Shape),
                    Context::UnknownObject,
                )
            },
            b"grpSp" => {
                self.take_anchor_child(AnchorChild::Object)?;
                self.open_object(
                    ObjectKind::Unknown(UnknownKind::Group),
                    Context::UnknownObject,
                )
            },
            b"cxnSp" => {
                self.take_anchor_child(AnchorChild::Object)?;
                self.open_object(
                    ObjectKind::Unknown(UnknownKind::Connection),
                    Context::UnknownObject,
                )
            },
            b"contentPart" => {
                self.take_anchor_child(AnchorChild::Object)?;
                let relationship_id = required_core_content_part_relationship(
                    element,
                    decoder,
                    resolver,
                    self.relationship_dialect,
                )?;
                set_relationship(
                    &mut self.anchor_mut()?.content_part_relationship_id,
                    relationship_id,
                    "content part",
                )?;
                self.open_object(
                    ObjectKind::Unknown(UnknownKind::ContentPart),
                    Context::ContentPart,
                )
            },
            _ => Ok(Context::Ignored),
        }
    }

    fn start_marker(
        &mut self,
        target: MarkerTarget,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<Context> {
        if !is_spreadsheet_drawing_namespace(namespace) {
            return Ok(Context::Ignored);
        }
        let field = match element.name().local_name().as_ref() {
            b"col" => MarkerField::Column,
            b"colOff" => MarkerField::ColumnOffset,
            b"row" => MarkerField::Row,
            b"rowOff" => MarkerField::RowOffset,
            _ => return Ok(Context::Ignored),
        };
        let marker = match target {
            MarkerTarget::From => self.anchor_mut()?.from.as_mut(),
            MarkerTarget::To => self.anchor_mut()?.to.as_mut(),
        }
        .ok_or_else(|| invalid("drawing marker value outside from/to"))?;
        if marker.next_field != field.ordinal() {
            return Err(invalid("drawing marker fields are out of order"));
        }
        self.marker_text.clear();
        Ok(Context::Marker(target, field))
    }

    fn open_object(&mut self, object_kind: ObjectKind, context: Context) -> Result<Context> {
        let anchor = self.anchor_mut()?;
        if anchor.object_kind.replace(object_kind).is_some() {
            return Err(invalid("drawing anchor has duplicate objects"));
        }
        Ok(context)
    }

    fn capture_description(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if let Some(description) = unqualified_attribute_value(element, b"descr", decoder)? {
            if description.len() > MAX_STRING_BYTES {
                return Err(limit("drawing description"));
            }
            if self
                .anchor_mut()?
                .description
                .replace(description)
                .is_some()
            {
                return Err(invalid("drawing object has duplicate descriptions"));
            }
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        match context {
            Context::Marker(target, field) => self.finish_marker(target, field),
            Context::From(target) | Context::To(target) => self.finish_marker_container(target),
            Context::Anchor(_) => self.finish_anchor(),
            Context::Root
            | Context::Position
            | Context::Extent
            | Context::Picture
            | Context::PictureNvPr
            | Context::PictureBlipFill
            | Context::ChartFrame
            | Context::ChartGraphic
            | Context::ChartGraphicData
            | Context::UnknownObject
            | Context::UnknownNvPr
            | Context::ContentPart
            | Context::Ignored => Ok(()),
        }
    }

    fn finish_marker(&mut self, target: MarkerTarget, field: MarkerField) -> Result<()> {
        let value = trim_xml_schema_whitespace(&self.marker_text);
        let parsed = match field {
            MarkerField::Column | MarkerField::Row => parse_value(value, "drawing marker")?,
            MarkerField::ColumnOffset | MarkerField::RowOffset => {
                parse_value(value, "drawing marker offset")?
            },
        };
        let marker = match target {
            MarkerTarget::From => self.anchor_mut()?.from.as_mut(),
            MarkerTarget::To => self.anchor_mut()?.to.as_mut(),
        }
        .ok_or_else(|| invalid("drawing marker value outside from/to"))?;
        marker.set(field, parsed, "drawing marker")
    }

    fn finish_marker_container(&mut self, target: MarkerTarget) -> Result<()> {
        let marker = match target {
            MarkerTarget::From => self.anchor_mut()?.from.as_ref(),
            MarkerTarget::To => self.anchor_mut()?.to.as_ref(),
        }
        .ok_or_else(|| invalid("drawing marker is missing its value"))?;
        marker.finish(match target {
            MarkerTarget::From => "drawing from marker",
            MarkerTarget::To => "drawing to marker",
        })?;
        Ok(())
    }

    fn finish_anchor(&mut self) -> Result<()> {
        let pending = self
            .anchor
            .take()
            .ok_or_else(|| invalid("missing pending drawing anchor"))?;
        if pending.phase != 4 || !pending.client_data_seen {
            return Err(invalid("drawing anchor is missing required children"));
        }
        let drawing_anchor = match pending.kind {
            AnchorKind::TwoCell => {
                let from = pending
                    .from
                    .as_ref()
                    .ok_or_else(|| invalid("drawing anchor is missing from marker"))?
                    .finish("drawing from marker")?;
                let to = pending
                    .to
                    .as_ref()
                    .ok_or_else(|| invalid("drawing anchor is missing to marker"))?
                    .finish("drawing to marker")?;
                check_marker_bounds(from)?;
                check_marker_bounds(to)?;
                if to.row < from.row
                    || to.column < from.column
                    || (to.row == from.row && to.row_offset < from.row_offset)
                    || (to.column == from.column && to.column_offset < from.column_offset)
                {
                    return Err(invalid("drawing anchor has descending markers"));
                }
                DrawingAnchor::TwoCell {
                    from,
                    to,
                    edit_as: pending.edit_as,
                }
            },
            AnchorKind::OneCell => {
                let from = pending
                    .from
                    .as_ref()
                    .ok_or_else(|| invalid("one-cell anchor is missing from marker"))?
                    .finish("drawing from marker")?;
                check_marker_bounds(from)?;
                let extent = pending
                    .extent
                    .ok_or_else(|| invalid("one-cell anchor is missing its extent"))?;
                DrawingAnchor::OneCell { from, extent }
            },
            AnchorKind::Absolute => {
                let position = pending
                    .position
                    .ok_or_else(|| invalid("absolute anchor is missing its position"))?;
                let extent = pending
                    .extent
                    .ok_or_else(|| invalid("absolute anchor is missing its extent"))?;
                DrawingAnchor::Absolute { position, extent }
            },
        };
        let compatibility_anchor = compatibility_anchor(&drawing_anchor);
        match (
            pending.object_kind,
            pending.picture_relationship_id,
            pending.chart_relationship_id,
            pending.content_part_relationship_id,
        ) {
            (Some(ObjectKind::Picture), Some(relationship_id), None, None) => {
                self.drawing.push(Object::Picture(Picture {
                    anchor: compatibility_anchor,
                    drawing_anchor,
                    relationship_id,
                    description: pending.description,
                }));
            },
            (Some(ObjectKind::ChartFrame), None, Some(relationship_id), None) => {
                self.drawing.push(Object::Chart(Chart {
                    anchor: compatibility_anchor,
                    drawing_anchor,
                    relationship_id,
                }));
            },
            (Some(ObjectKind::ChartFrame), None, None, None) => {
                self.drawing.push(Object::Unknown(Unknown {
                    anchor: compatibility_anchor,
                    drawing_anchor,
                    description: pending.description,
                    kind: UnknownKind::Other,
                }));
            },
            (Some(ObjectKind::Unknown(UnknownKind::ContentPart)), None, None, Some(_)) => {
                // The source-backed owner retains the physical relationship
                // edge.  The typed inventory only needs the validated
                // structural object and deliberately does not expose r:id.
                self.drawing.push(Object::Unknown(Unknown {
                    anchor: compatibility_anchor,
                    drawing_anchor,
                    description: pending.description,
                    kind: UnknownKind::ContentPart,
                }));
            },
            (Some(ObjectKind::Unknown(kind)), None, None, None) => {
                self.drawing.push(Object::Unknown(Unknown {
                    anchor: compatibility_anchor,
                    drawing_anchor,
                    description: pending.description,
                    kind,
                }));
            },
            (None, _, _, _) => return Err(invalid("drawing anchor has no object")),
            (Some(ObjectKind::Picture), _, _, _) => {
                return Err(invalid("drawing picture has an invalid image relationship"));
            },
            (Some(ObjectKind::ChartFrame), _, _, _) => {
                return Err(invalid("drawing chart frame has invalid relationships"));
            },
            (Some(ObjectKind::Unknown(_)), _, _, _) => {
                return Err(invalid("drawing unknown object has relationships"));
            },
        }
        Ok(())
    }

    fn take_anchor_child(&mut self, child: AnchorChild) -> Result<()> {
        self.anchor_mut()?.take_child(child)
    }

    fn anchor_mut(&mut self) -> Result<&mut PendingAnchor> {
        self.anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing object outside an anchor"))
    }

    fn append_marker_text(&mut self, text: &str) -> Result<()> {
        let length = self
            .marker_text
            .len()
            .checked_add(text.len())
            .ok_or_else(|| limit("drawing marker text"))?;
        if length > MAX_MARKER_TEXT_BYTES {
            return Err(limit("drawing marker text"));
        }
        self.marker_text
            .try_reserve(text.len())
            .map_err(|source| allocation("SpreadsheetDrawing marker text", source))?;
        self.marker_text.push_str(text);
        Ok(())
    }
}

pub fn parse(xml: &str) -> Result<Option<Drawing>> {
    parse_with_limits(xml, &litchi_opc::ReadLimits::default())
}

/// Parse a worksheet drawing under the caller's retained package read policy.
///
/// The MCE processor receives the same per-part, XML-event, and XML-depth
/// ceilings that bound the typed drawing parser.  Its bounded output is then
/// passed directly to the parser, so a selected MCE branch cannot grow beyond
/// the caller's part budget before semantic allocations begin.
pub(crate) fn parse_with_limits(
    xml: &str,
    caller: &litchi_opc::ReadLimits,
) -> Result<Option<Drawing>> {
    let caller_part_bytes = usize::try_from(caller.max_part_bytes()).unwrap_or(usize::MAX);
    let max_xml_bytes = MAX_DRAWING_XML_BYTES.min(caller_part_bytes);
    if xml.len() > max_xml_bytes {
        return Err(limit("drawing XML"));
    }

    let defaults = litchi_ooxml_common::mce::Limits::default();
    let max_events = MAX_DRAWING_EVENTS.min(caller.max_xml_events());
    let max_depth = MAX_DRAWING_DEPTH.min(caller.max_xml_depth());
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: defaults.max_input_bytes.min(max_xml_bytes),
        max_output_bytes: defaults.max_output_bytes.min(max_xml_bytes),
        max_depth: defaults.max_depth.min(max_depth),
        max_namespace_bindings: defaults.max_namespace_bindings.min(max_events),
        max_directive_tokens: defaults.max_directive_tokens.min(max_events),
        max_choices_per_alternate: defaults.max_choices_per_alternate.min(max_events),
    };
    let processed = litchi_ooxml_common::mce::process_markup_compatibility(
        xml.as_bytes(),
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?;
    let xml = match processed.xml {
        Cow::Borrowed(_) => Cow::Borrowed(xml),
        Cow::Owned(bytes) => Cow::Owned(
            String::from_utf8(bytes)
                .map_err(|error| Error::Invalid(format!("MCE output is not UTF-8: {error}")))?,
        ),
    };
    Parser::parse(xml.as_ref(), max_depth, max_events)
}

fn compatibility_anchor(anchor: &DrawingAnchor) -> Anchor {
    match anchor {
        DrawingAnchor::TwoCell { from, to, .. } => Anchor::with_offsets(
            from.column,
            from.column_offset.emu(),
            from.row,
            from.row_offset.emu(),
            to.column,
            to.column_offset.emu(),
            to.row,
            to.row_offset.emu(),
        ),
        // The legacy chart anchor has no extent vocabulary. Retain the source
        // origin as a useful compatibility projection while exposing the real
        // extent through DrawingAnchor.
        DrawingAnchor::OneCell { from, .. } => Anchor::with_offsets(
            from.column,
            from.column_offset.emu(),
            from.row,
            from.row_offset.emu(),
            from.column,
            from.column_offset.emu(),
            from.row,
            from.row_offset.emu(),
        ),
        // Existing absolute consumers observed a zero placeholder. Keep it
        // stable; the accurate public projection is never zeroed.
        DrawingAnchor::Absolute { .. } => Anchor::new(0, 0, 0, 0),
    }
}

fn push_context(stack: &mut Vec<Context>, context: Context, max_depth: usize) -> Result<()> {
    check_context_depth(stack.len(), max_depth)?;
    stack
        .try_reserve(1)
        .map_err(|source| allocation("SpreadsheetDrawing XML context stack", source))?;
    stack.push(context);
    Ok(())
}

fn check_context_depth(depth: usize, max_depth: usize) -> Result<()> {
    if depth >= max_depth {
        return Err(limit("drawing XML depth"));
    }
    Ok(())
}

fn check_marker_bounds(marker: CellMarker) -> Result<()> {
    if marker.column >= 16_384 || marker.row >= 1_048_576 {
        return Err(invalid("drawing anchor exceeds worksheet bounds"));
    }
    Ok(())
}

fn check_coordinate(value: i64, description: &str) -> Result<()> {
    if !(MIN_COORDINATE..=MAX_COORDINATE).contains(&value) {
        return Err(invalid(format!("{description} exceeds DrawingML bounds")));
    }
    Ok(())
}

fn check_positive_coordinate(value: i64, description: &str) -> Result<()> {
    if !(0..=MAX_COORDINATE).contains(&value) {
        return Err(invalid(format!("{description} exceeds DrawingML bounds")));
    }
    Ok(())
}

fn check_attribute_lengths(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Invalid(error.to_string()))?;
        if attribute.value.as_ref().len() > MAX_STRING_BYTES {
            return Err(limit("drawing attribute"));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Invalid(error.to_string()))?;
        if decoded.len() > MAX_STRING_BYTES {
            return Err(limit("decoded drawing attribute"));
        }
    }
    Ok(())
}

fn coordinate_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{description} attribute is missing")))?;
    let value = value
        .trim_matches([' ', '\t', '\n', '\r'])
        .parse()
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))?;
    check_coordinate(value, description)?;
    Ok(value)
}

fn positive_coordinate_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Emu> {
    Ok(Emu(check_positive_coordinate_value(
        element,
        name,
        decoder,
        description,
    )?))
}

fn check_positive_coordinate_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{description} attribute is missing")))?;
    let value = value
        .trim_matches([' ', '\t', '\n', '\r'])
        .parse()
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))?;
    check_positive_coordinate(value, description)?;
    Ok(value)
}

fn is_spreadsheet_drawing_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name && is_spreadsheet_drawing_namespace(namespace)
}

fn trim_xml_schema_whitespace(value: &str) -> &str {
    value.trim_matches([' ', '\t', '\n', '\r'])
}

fn is_spreadsheet_drawing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEET_DRAWING_NAMESPACE
                || *value == STRICT_SPREADSHEET_DRAWING_NAMESPACE
    )
}

fn drawing_relationship_dialect(namespace: &ResolveResult<'_>) -> Result<RelationshipDialect> {
    match namespace {
        ResolveResult::Bound(Namespace(value))
            if *value == STRICT_SPREADSHEET_DRAWING_NAMESPACE =>
        {
            Ok(RelationshipDialect::Strict)
        },
        ResolveResult::Bound(Namespace(value)) if *value == SPREADSHEET_DRAWING_NAMESPACE => {
            Ok(RelationshipDialect::Transitional)
        },
        _ => Err(invalid(
            "drawing root has no recognized SpreadsheetDrawing dialect",
        )),
    }
}

fn required_core_content_part_relationship(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    drawing_dialect: Option<RelationshipDialect>,
) -> Result<String> {
    let expected =
        drawing_dialect.ok_or_else(|| invalid("contentPart appears before the drawing root"))?;
    let mut found = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Invalid(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let dialect = match namespace {
            ResolveResult::Bound(Namespace(value)) if *value == *STRICT_RELATIONSHIPS_NAMESPACE => {
                RelationshipDialect::Strict
            },
            ResolveResult::Bound(Namespace(value)) if *value == *RELATIONSHIPS_NAMESPACE => {
                RelationshipDialect::Transitional
            },
            _ => {
                return Err(invalid(
                    "core SpreadsheetDrawing contentPart allows only a matching-dialect r:id",
                ));
            },
        };
        if attribute.key.local_name().as_ref() != b"id" {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart allows only a matching-dialect r:id",
            ));
        }
        if dialect != expected {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart r:id uses the wrong relationship dialect",
            ));
        }
        if found.is_some() {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart has duplicate r:id",
            ));
        }
        if attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4) {
            return Err(limit("contentPart relationship ID"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Invalid(error.to_string()))?
            .into_owned();
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES || !is_ncname(&value) {
            return Err(invalid("contentPart relationship ID is not an XML NCName"));
        }
        found = Some(value);
    }
    found.ok_or_else(|| invalid("core SpreadsheetDrawing contentPart is missing r:id"))
}

fn set_relationship(target: &mut Option<String>, value: String, kind: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("drawing {kind} relationship ID is empty")));
    }
    if value.len() > MAX_STRING_BYTES {
        return Err(limit(format!("drawing {kind} relationship ID")));
    }
    if target.replace(value).is_some() {
        return Err(invalid(format!(
            "drawing anchor has duplicate {kind} relationships"
        )));
    }
    Ok(())
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(invalid(format!("duplicate {description}")));
    }
    Ok(())
}

fn parse_value<T: std::str::FromStr>(value: &str, description: &str) -> Result<T> {
    value
        .parse()
        .map_err(|_source| invalid(format!("invalid {description} '{value}'")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(message: impl Into<String>) -> Error {
    Error::Invalid(format!(
        "SpreadsheetDrawing {} limit exceeded",
        message.into()
    ))
}

#[cfg(test)]
mod tests {
    use super::parse_with_limits;
    use litchi_ooxml_common::mce::process_markup_compatibility;

    const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    #[test]
    fn caller_part_cap_rejects_mce_output_growth_before_typed_parse() {
        let xml = format!(
            r#"<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{DRAWINGML}" xmlns:r="{RELATIONSHIPS}" xmlns:mc="{MCE}"/>"#
        );
        let expanded = process_markup_compatibility(
            xml.as_bytes(),
            &litchi_ooxml_common::mce::Capabilities::default(),
            &litchi_ooxml_common::mce::Limits {
                max_input_bytes: xml.len(),
                max_output_bytes: 1024 * 1024,
                ..Default::default()
            },
        )
        .expect("test MCE input should preprocess")
        .xml;
        assert!(expanded.len() > xml.len());

        let caller_limit = litchi_opc::ReadLimits::builder()
            .max_part_bytes((expanded.len() - 1) as u64)
            .expect("positive part limit")
            .build()
            .expect("consistent read limits");
        assert!(xml.len() <= caller_limit.max_part_bytes() as usize);
        assert!(parse_with_limits(&xml, &caller_limit).is_err());
    }

    #[test]
    fn exact_part_cap_accepts_unmodified_drawing_without_mce() {
        let xml = format!(r#"<wsDr xmlns="{XDR}"/>"#);
        let caller_limit = litchi_opc::ReadLimits::builder()
            .max_part_bytes(xml.len() as u64)
            .expect("positive part limit")
            .build()
            .expect("consistent read limits");

        let drawing = parse_with_limits(&xml, &caller_limit)
            .expect("exact no-MCE part limit should be accepted")
            .expect("drawing root should be present");
        assert!(drawing.is_empty());
    }

    #[test]
    fn depth_is_admitted_before_empty_element_semantics() {
        let empty_root = format!(r#"<wsDr xmlns="{XDR}"/>"#);
        assert!(super::Parser::parse(&empty_root, 0, 8).is_err());

        let empty_child = format!(r#"<wsDr xmlns="{XDR}"><future/></wsDr>"#);
        let caller_limit = litchi_opc::ReadLimits::builder()
            .max_xml_depth(1)
            .expect("positive depth limit")
            .build()
            .expect("consistent read limits");
        assert!(parse_with_limits(&empty_child, &caller_limit).is_err());
    }
}

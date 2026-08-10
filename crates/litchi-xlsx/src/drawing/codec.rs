//! Namespace-aware `SpreadsheetDrawing` reader.
//!
//! Both `xdr:twoCellAnchor` (worksheets) and `xdr:absoluteAnchor`
//! (chartsheets) are understood; absolute anchors record a zero placeholder
//! anchor because they carry EMU positions rather than cell markers.

use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::Anchor;
use super::model::{Chart, Drawing, Object, Picture, Unknown, UnknownKind};
use crate::error::{Error, Result};
use crate::raw::namespace::relationship_attribute_value;
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_chart_name, is_drawingml_name, unqualified_attribute_value,
};

const SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const MAX_DRAWING_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWING_ANCHORS: usize = 100_000;
const MAX_MARKER_TEXT_BYTES: usize = 64;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    TwoCellAnchor,
    AbsoluteAnchor,
    From,
    To,
    Marker(MarkerTarget, MarkerField),
    Other,
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

#[derive(Default)]
struct Marker {
    column: Option<u32>,
    column_offset: Option<i64>,
    row: Option<u32>,
    row_offset: Option<i64>,
}

impl Marker {
    fn finish(self, description: &str) -> Result<(u32, i64, u32, i64)> {
        Ok((
            self.column
                .ok_or_else(|| invalid(format!("{description} is missing its column")))?,
            self.column_offset
                .ok_or_else(|| invalid(format!("{description} is missing its column offset")))?,
            self.row
                .ok_or_else(|| invalid(format!("{description} is missing its row")))?,
            self.row_offset
                .ok_or_else(|| invalid(format!("{description} is missing its row offset")))?,
        ))
    }
}

#[derive(Default)]
struct PendingAnchor {
    from: Option<Marker>,
    to: Option<Marker>,
    picture_relationship_id: Option<String>,
    chart_relationship_id: Option<String>,
    description: Option<String>,
    unknown_kind: Option<UnknownKind>,
}

struct Parser {
    drawing: Drawing,
    anchor: Option<PendingAnchor>,
    marker_text: String,
}

impl Parser {
    fn parse(xml: &str) -> Result<Option<Drawing>> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let mut parser = Self {
            drawing: Drawing::default(),
            anchor: None,
            marker_text: String::new(),
        };
        let mut stack = Vec::new();
        let mut closed_root = false;
        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Invalid(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root {
                        return Err(invalid("drawing XML contains multiple root elements"));
                    }
                    if !is_spreadsheet_drawing_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    stack.push(Context::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_spreadsheet_drawing_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    return Ok(Some(parser.drawing));
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    stack.push(parser.start(parent, &namespace, &element, decoder, &resolver)?);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    parser.finish(context)?;
                },
                Event::Text(text) if matches!(stack.last(), Some(Context::Marker(_, _))) => {
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    parser.append_marker_text(&text)?;
                },
                Event::CData(text) if matches!(stack.last(), Some(Context::Marker(_, _))) => {
                    let text = text
                        .decode()
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    parser.append_marker_text(&text)?;
                },
                Event::GeneralRef(reference)
                    if matches!(stack.last(), Some(Context::Marker(_, _))) =>
                {
                    let text = decode_xml_reference(&reference)?;
                    parser.append_marker_text(&text)?;
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
        if parent == Context::Root
            && is_spreadsheet_drawing_name(namespace, element.name(), b"twoCellAnchor")
        {
            if self.anchor.is_some() {
                return Err(invalid("nested drawing anchor"));
            }
            if self.drawing.len() >= MAX_DRAWING_ANCHORS {
                return Err(invalid("drawing exceeds the anchor limit"));
            }
            self.anchor = Some(PendingAnchor::default());
            return Ok(Context::TwoCellAnchor);
        }
        if parent == Context::Root
            && is_spreadsheet_drawing_name(namespace, element.name(), b"absoluteAnchor")
        {
            if self.anchor.is_some() {
                return Err(invalid("nested drawing anchor"));
            }
            if self.drawing.len() >= MAX_DRAWING_ANCHORS {
                return Err(invalid("drawing exceeds the anchor limit"));
            }
            self.anchor = Some(PendingAnchor::default());
            return Ok(Context::AbsoluteAnchor);
        }
        if parent == Context::TwoCellAnchor
            && is_spreadsheet_drawing_name(namespace, element.name(), b"from")
        {
            let anchor = self.anchor_mut()?;
            if anchor.from.replace(Marker::default()).is_some() {
                return Err(invalid("drawing anchor has duplicate from markers"));
            }
            return Ok(Context::From);
        }
        if parent == Context::TwoCellAnchor
            && is_spreadsheet_drawing_name(namespace, element.name(), b"to")
        {
            let anchor = self.anchor_mut()?;
            if anchor.to.replace(Marker::default()).is_some() {
                return Err(invalid("drawing anchor has duplicate to markers"));
            }
            return Ok(Context::To);
        }
        let target = match parent {
            Context::From => Some(MarkerTarget::From),
            Context::To => Some(MarkerTarget::To),
            Context::Root
            | Context::TwoCellAnchor
            | Context::AbsoluteAnchor
            | Context::Marker(..)
            | Context::Other => None,
        };
        if let Some(target) = target {
            for (name, field) in [
                (b"col".as_slice(), MarkerField::Column),
                (b"colOff".as_slice(), MarkerField::ColumnOffset),
                (b"row".as_slice(), MarkerField::Row),
                (b"rowOff".as_slice(), MarkerField::RowOffset),
            ] {
                if is_spreadsheet_drawing_name(namespace, element.name(), name) {
                    self.marker_text.clear();
                    return Ok(Context::Marker(target, field));
                }
            }
        }
        if matches!(parent, Context::TwoCellAnchor | Context::AbsoluteAnchor)
            && is_spreadsheet_drawing_namespace(namespace)
            && let Some(kind) = unknown_kind(element.name().local_name().as_ref())
        {
            self.anchor_mut()?.unknown_kind = Some(kind);
        }
        if self.anchor.is_some()
            && is_spreadsheet_drawing_name(namespace, element.name(), b"cNvPr")
            && let Some(description) = unqualified_attribute_value(element, b"descr", decoder)?
        {
            if description.len() > MAX_STRING_BYTES {
                return Err(limit("drawing description"));
            }
            self.anchor_mut()?.description = Some(description);
        }
        if self.anchor.is_some() && is_drawingml_name(namespace, element.name(), b"blip") {
            let relationship_id = relationship_attribute_value(
                element, b"embed", decoder, resolver,
            )?
            .ok_or_else(|| invalid("drawing picture blip is missing an embed relationship"))?;
            set_relationship(
                &mut self.anchor_mut()?.picture_relationship_id,
                relationship_id,
                "picture",
            )?;
        }
        if self.anchor.is_some() && is_drawingml_chart_name(namespace, element.name(), b"chart") {
            let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
                .ok_or_else(|| invalid("drawing chart is missing a relationship ID"))?;
            set_relationship(
                &mut self.anchor_mut()?.chart_relationship_id,
                relationship_id,
                "chart",
            )?;
        }
        Ok(Context::Other)
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        match context {
            Context::Marker(target, field) => self.finish_marker(target, field),
            Context::TwoCellAnchor => self.finish_anchor(),
            Context::AbsoluteAnchor => self.finish_absolute_anchor(),
            Context::Root | Context::From | Context::To | Context::Other => Ok(()),
        }
    }

    fn finish_marker(&mut self, target: MarkerTarget, field: MarkerField) -> Result<()> {
        let value = self.marker_text.trim();
        let anchor = self
            .anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing marker outside an anchor"))?;
        let marker = match target {
            MarkerTarget::From => anchor.from.as_mut(),
            MarkerTarget::To => anchor.to.as_mut(),
        }
        .ok_or_else(|| invalid("drawing marker value outside from/to"))?;
        match field {
            MarkerField::Column => set_once(
                &mut marker.column,
                parse_value(value, "drawing column")?,
                "drawing column",
            ),
            MarkerField::ColumnOffset => set_once(
                &mut marker.column_offset,
                parse_value(value, "drawing column offset")?,
                "drawing column offset",
            ),
            MarkerField::Row => set_once(
                &mut marker.row,
                parse_value(value, "drawing row")?,
                "drawing row",
            ),
            MarkerField::RowOffset => set_once(
                &mut marker.row_offset,
                parse_value(value, "drawing row offset")?,
                "drawing row offset",
            ),
        }
    }

    fn finish_anchor(&mut self) -> Result<()> {
        let PendingAnchor {
            from,
            to,
            picture_relationship_id,
            chart_relationship_id,
            description,
            unknown_kind,
        } = self
            .anchor
            .take()
            .ok_or_else(|| invalid("missing pending drawing anchor"))?;
        let (from_col, from_col_offset, from_row, from_row_offset) = from
            .ok_or_else(|| invalid("drawing anchor is missing from marker"))?
            .finish("drawing from marker")?;
        let (to_col, to_col_offset, to_row, to_row_offset) = to
            .ok_or_else(|| invalid("drawing anchor is missing to marker"))?
            .finish("drawing to marker")?;
        if from_col >= 16_384 || to_col >= 16_384 || from_row >= 1_048_576 || to_row >= 1_048_576 {
            return Err(invalid("drawing anchor exceeds worksheet bounds"));
        }
        if to_row < from_row
            || to_col < from_col
            || (to_row == from_row && to_row_offset < from_row_offset)
            || (to_col == from_col && to_col_offset < from_col_offset)
        {
            return Err(invalid("drawing anchor has descending markers"));
        }
        let anchor = Anchor::with_offsets(
            from_col,
            from_col_offset,
            from_row,
            from_row_offset,
            to_col,
            to_col_offset,
            to_row,
            to_row_offset,
        );
        Self::push_anchor_object(
            &mut self.drawing,
            picture_relationship_id,
            chart_relationship_id,
            description,
            unknown_kind,
            anchor,
        )
    }

    /// Finish an `xdr:absoluteAnchor` (the only anchor kind chartsheets
    /// use). It carries EMU positions instead of cell markers, so the
    /// recorded anchor is a zero placeholder; callers that only need the
    /// anchored chart or picture are unaffected.
    fn finish_absolute_anchor(&mut self) -> Result<()> {
        let PendingAnchor {
            picture_relationship_id,
            chart_relationship_id,
            description,
            unknown_kind,
            ..
        } = self
            .anchor
            .take()
            .ok_or_else(|| invalid("missing pending drawing anchor"))?;
        Self::push_anchor_object(
            &mut self.drawing,
            picture_relationship_id,
            chart_relationship_id,
            description,
            unknown_kind,
            Anchor::new(0, 0, 0, 0),
        )
    }

    fn push_anchor_object(
        drawing: &mut Drawing,
        picture_relationship_id: Option<String>,
        chart_relationship_id: Option<String>,
        description: Option<String>,
        unknown_kind: Option<UnknownKind>,
        anchor: Anchor,
    ) -> Result<()> {
        match (picture_relationship_id, chart_relationship_id) {
            (Some(relationship_id), None) => drawing.push(Object::Picture(Picture {
                anchor,
                relationship_id,
                description,
            })),
            (None, Some(relationship_id)) => drawing.push(Object::Chart(Chart {
                anchor,
                relationship_id,
            })),
            (None, None) if let Some(kind) = unknown_kind => {
                drawing.push(Object::Unknown(Unknown {
                    anchor,
                    description,
                    kind,
                }));
            },
            (None, None) => return Err(invalid("drawing anchor has no picture or chart")),
            (Some(_), Some(_)) => {
                return Err(invalid("drawing anchor contains both a picture and chart"));
            },
        }
        Ok(())
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
        self.marker_text.push_str(text);
        Ok(())
    }
}

pub fn parse(xml: &str) -> Result<Option<Drawing>> {
    if xml.len() > MAX_DRAWING_XML_BYTES {
        return Err(limit("drawing XML"));
    }
    let xml = litchi_ooxml_common::mce::process_str(xml)?;
    Parser::parse(xml.as_ref())
}

fn is_spreadsheet_drawing_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == SPREADSHEET_DRAWING_NAMESPACE
                    || *value == STRICT_SPREADSHEET_DRAWING_NAMESPACE
        )
}

fn is_spreadsheet_drawing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEET_DRAWING_NAMESPACE
                || *value == STRICT_SPREADSHEET_DRAWING_NAMESPACE
    )
}

fn unknown_kind(local_name: &[u8]) -> Option<UnknownKind> {
    Some(match local_name {
        b"from" | b"to" | b"pos" | b"ext" | b"clientData" | b"pic" | b"graphicFrame" => {
            return None;
        },
        b"sp" => UnknownKind::Shape,
        b"grpSp" => UnknownKind::Group,
        b"cxnSp" => UnknownKind::Connection,
        b"contentPart" => UnknownKind::ContentPart,
        _ => UnknownKind::Other,
    })
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

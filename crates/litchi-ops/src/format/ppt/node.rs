//! Typed read nodes for legacy PowerPoint presentations.
//!
//! The legacy reader owns a catalog and lazily parses each slide's native
//! shape array.  A node keeps only the handle, its positional address, and an
//! `Arc` around the native shape selected from that array.  This gives the
//! resolver and the node view one stable object without retaining a whole
//! presentation scene for an unrelated sibling.

use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_ppt::shapes::shape::{Shape, ShapeProperties};
use litchi_ppt::shapes::{PictureFrameKind, ShapeEnum};

use super::Ppt;
use super::children::Children;
use super::elements::{Element, Values};
use super::spec::Kind;
use crate::Code;
use crate::error::{Error, Result};
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::path::navigate::KindSet;
use crate::value::Length;

const UNTITLED: &str = "(untitled)";

/// A node in a legacy PowerPoint presentation.
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Ppt,
    item: Item,
    steps: Arc<[PathStep<'static>]>,
}

/// The native-backed element carried by a node.
#[derive(Clone)]
pub(crate) enum Item {
    Presentation,
    Slide(SlideItem),
    Shape(ShapeItem),
    /// A slide's speaker notes: the notes page's text.
    Notes(NotesItem),
    /// One paragraph of a text shape: the text between paragraph marks.
    Paragraph(ParagraphItem),
    /// An element of the record-0323 surface (comments, table rows, cells
    /// and columns, hyperlinks, runs, line breaks, the theme, charts).
    Element(Box<Element>),
}

/// A slide's speaker notes, read once when the node is resolved.
#[derive(Clone)]
pub(crate) struct NotesItem {
    pub(crate) slide: usize,
    pub(crate) text: Arc<str>,
}

/// A paragraph of a text shape: the shape it belongs to and its zero-based
/// index among the shape's paragraphs. Its text is cut from the shape's
/// aggregate text at the paragraph marks (`\r`) the binary format stores.
#[derive(Clone)]
pub(crate) struct ParagraphItem {
    pub(crate) shape: ShapeItem,
    pub(crate) index: usize,
}

/// A catalog position.  Slide positions are zero-based internally and are
/// rendered one-based by [`PathStep`].
#[derive(Clone, Copy)]
pub(crate) struct SlideItem {
    pub(crate) position: usize,
}

/// A shape and its position in the shared native shape array.
#[derive(Clone)]
pub(crate) struct ShapeItem {
    pub(crate) slide: usize,
    pub(crate) kind: Kind,
    pub(crate) source: Arc<[ShapeEnum<'static>]>,
    pub(crate) path: Arc<[usize]>,
}

impl<'f> Node<'f> {
    /// Construct the presentation root.
    pub(crate) fn root(file: &'f Ppt) -> Self {
        Self {
            file,
            item: Item::Presentation,
            steps: Arc::from(Vec::<PathStep<'static>>::new().into_boxed_slice()),
        }
    }

    /// Construct a slide at its zero-based catalog position.
    pub(crate) fn slide_node(file: &'f Ppt, position: usize) -> Self {
        Self {
            file,
            item: Item::Slide(SlideItem { position }),
            steps: Arc::from(
                vec![PathStep::at("slide", Position::new(position))].into_boxed_slice(),
            ),
        }
    }

    /// Construct a shape under `parent_steps`.
    pub(crate) fn shape_node(
        file: &'f Ppt,
        parent_steps: &[PathStep<'static>],
        slide: usize,
        kind: Kind,
        ordinal: usize,
        source: Arc<[ShapeEnum<'static>]>,
        path: Arc<[usize]>,
    ) -> Self {
        let mut steps = parent_steps.to_vec();
        steps.push(PathStep::at(kind.name(), Position::new(ordinal)));
        Self {
            file,
            item: Item::Shape(ShapeItem {
                slide,
                kind,
                source,
                path,
            }),
            steps: Arc::from(steps.into_boxed_slice()),
        }
    }

    /// Construct a slide's notes node.
    pub(crate) fn notes_node(file: &'f Ppt, slide: usize, text: Arc<str>) -> Self {
        Self {
            file,
            item: Item::Notes(NotesItem { slide, text }),
            steps: Arc::from(
                vec![
                    PathStep::at("slide", Position::new(slide)),
                    PathStep::new("notes"),
                ]
                .into_boxed_slice(),
            ),
        }
    }

    /// Construct the paragraph at `index` of a text shape node.
    pub(crate) fn paragraph_node(parent: &Self, shape: ShapeItem, index: usize) -> Self {
        Self::paragraph_at(parent.file, &parent.steps, shape, index)
    }

    /// Construct the paragraph at `index` of the text shape at `shape_steps`.
    pub(crate) fn paragraph_at(
        file: &'f Ppt,
        shape_steps: &[PathStep<'static>],
        shape: ShapeItem,
        index: usize,
    ) -> Self {
        let mut steps = shape_steps.to_vec();
        steps.push(PathStep::at("paragraph", Position::new(index)));
        Self {
            file,
            item: Item::Paragraph(ParagraphItem { shape, index }),
            steps: Arc::from(steps.into_boxed_slice()),
        }
    }

    /// Construct a record-0323 element node under `parent_steps`.
    pub(crate) fn element_node(
        file: &'f Ppt,
        parent_steps: &[PathStep<'static>],
        step: PathStep<'static>,
        element: Element,
    ) -> Self {
        let mut steps = parent_steps.to_vec();
        steps.push(step);
        Self {
            file,
            item: Item::Element(Box::new(element)),
            steps: Arc::from(steps.into_boxed_slice()),
        }
    }

    /// The record-0323 element carried by this node.
    pub(crate) fn element_item(&self) -> Option<&Element> {
        match &self.item {
            Item::Element(element) => Some(element),
            _ => None,
        }
    }

    pub(crate) fn file(&self) -> &'f Ppt {
        self.file
    }

    pub(crate) fn item(&self) -> &Item {
        &self.item
    }

    pub(crate) fn steps(&self) -> &[PathStep<'static>] {
        &self.steps
    }

    /// The zero-based slide position containing this node.
    #[must_use]
    pub fn slide(&self) -> Option<Position> {
        self.slide_position().map(Position::new)
    }

    pub(crate) fn slide_position(&self) -> Option<usize> {
        match &self.item {
            Item::Presentation => None,
            Item::Slide(slide) => Some(slide.position),
            Item::Shape(shape) => Some(shape.slide),
            Item::Notes(notes) => Some(notes.slide),
            Item::Paragraph(paragraph) => Some(paragraph.shape.slide),
            Item::Element(element) => element.slide(),
        }
    }

    /// Return the native source-order target for a direct slide shape.
    ///
    /// The native PPT edit owner addresses a shape by its index in the
    /// complete slide shape array.  A navigation path's final ordinal is
    /// kind-specific (`/shape[N]`, `/picture[N]`, ...), so it cannot be used
    /// as the native target.  Nested group paths are deliberately refused by
    /// returning `None` until the native transaction grows a group-aware
    /// target identity.
    pub(crate) fn native_shape_target(&self) -> Option<litchi_ppt::text_edit::Target> {
        let shape = self.shape_item()?;
        let [position] = shape.path.as_ref() else {
            return None;
        };
        Some(litchi_ppt::text_edit::Target::new(
            Position::new(shape.slide),
            Position::new(*position),
        ))
    }

    /// Return the native complete anchor bounds for a shape that exposes
    /// them through the read projection.
    pub(crate) fn anchor_bounds(&self) -> Option<(i32, i32, i32, i32)> {
        match self.shape()? {
            ShapeEnum::TextBox(shape) => {
                let properties = shape.properties();
                Some((
                    properties.x,
                    properties.y,
                    properties.width,
                    properties.height,
                ))
            },
            ShapeEnum::Placeholder(shape) => {
                let properties = shape.properties();
                Some((
                    properties.x,
                    properties.y,
                    properties.width,
                    properties.height,
                ))
            },
            ShapeEnum::AutoShape(shape) => {
                let properties = shape.properties();
                Some((
                    properties.x,
                    properties.y,
                    properties.width,
                    properties.height,
                ))
            },
            ShapeEnum::Picture(shape) => {
                let properties = shape.properties();
                Some((
                    properties.x,
                    properties.y,
                    properties.width,
                    properties.height,
                ))
            },
            ShapeEnum::Table(shape) => {
                Some((shape.left(), shape.top(), shape.width(), shape.height()))
            },
            ShapeEnum::Group(shape) => {
                Some((shape.left(), shape.top(), shape.width(), shape.height()))
            },
            ShapeEnum::Line(_) => None,
        }
    }

    /// The native shape carried by a shape node.
    #[must_use]
    pub fn shape(&self) -> Option<&ShapeEnum<'static>> {
        match &self.item {
            Item::Shape(shape) => resolve_shape(&shape.source, &shape.path),
            Item::Paragraph(paragraph) => {
                resolve_shape(&paragraph.shape.source, &paragraph.shape.path)
            },
            Item::Element(element) => element
                .shape_item()
                .and_then(|shape| resolve_shape(&shape.source, &shape.path)),
            Item::Presentation | Item::Slide(_) | Item::Notes(_) => None,
        }
    }

    /// The shape item behind a shape node, or the shape a paragraph belongs to.
    pub(crate) fn shape_item(&self) -> Option<&ShapeItem> {
        match &self.item {
            Item::Shape(shape) => Some(shape),
            Item::Paragraph(paragraph) => Some(&paragraph.shape),
            Item::Element(element) => element.shape_item(),
            Item::Presentation | Item::Slide(_) | Item::Notes(_) => None,
        }
    }

    /// The zero-based paragraph index of a paragraph node.
    pub(crate) fn paragraph_index(&self) -> Option<usize> {
        match &self.item {
            Item::Paragraph(paragraph) => Some(paragraph.index),
            _ => None,
        }
    }

    /// The shape name (`wzName`), or a picture's file name when it has none.
    #[must_use]
    pub fn shape_name(&self) -> Option<&str> {
        self.shape().and_then(shape_name)
    }

    /// The physical element kind used by navigation and query matching.
    pub(crate) fn element(&self) -> Kind {
        match &self.item {
            Item::Presentation => Kind::Presentation,
            Item::Slide(_) => Kind::Slide,
            Item::Shape(shape) => shape.kind,
            Item::Notes(_) => Kind::Notes,
            Item::Paragraph(_) => Kind::Paragraph,
            Item::Element(element) => element.kind(),
        }
    }

    /// The semantic kind used by query inspection.  Placeholders retain
    /// their physical shape node and path, but are a distinct query subject.
    pub(crate) fn semantic_kind(&self) -> Kind {
        match (&self.item, self.shape()) {
            (Item::Shape(_), Some(ShapeEnum::Placeholder(_))) => Kind::Placeholder,
            _ => self.element(),
        }
    }

    /// The property table kind.  A placeholder remains physically a shape in
    /// the native tree, but has its own semantic properties.
    pub(crate) fn property_kind(&self) -> Kind {
        self.semantic_kind()
    }

    /// The OfficeCLI semantic type name.
    pub(crate) fn type_name(&self) -> &'static str {
        match (&self.item, self.shape()) {
            (Item::Shape(_), Some(ShapeEnum::Placeholder(placeholder)))
                if placeholder.is_title() =>
            {
                "title"
            },
            (Item::Shape(_), Some(ShapeEnum::Placeholder(_))) => "placeholder",
            _ => self.element().name(),
        }
    }

    /// The node's text member.  Legacy PPT exposes aggregate text for text
    /// shapes, including the native aggregate table/group text where the
    /// leaf provides it.
    pub(crate) fn text_member(&self) -> Result<Option<Cow<'_, str>>> {
        match &self.item {
            Item::Notes(notes) => Ok(Some(Cow::Borrowed(&notes.text))),
            Item::Paragraph(paragraph) => {
                let Some(shape) = self.shape() else {
                    return Err(vanished());
                };
                let text = shape.text().map_err(Error::from)?;
                Ok(Some(Cow::Owned(
                    paragraphs(&text)
                        .nth(paragraph.index)
                        .ok_or_else(vanished)?
                        .to_owned(),
                )))
            },
            Item::Element(element) => element.text(self),
            _ => match self.shape() {
                Some(shape)
                    if matches!(self.element(), Kind::Shape | Kind::Table | Kind::Group) =>
                {
                    shape
                        .text()
                        .map(|text| Some(Cow::Owned(text)))
                        .map_err(Error::from)
                },
                _ => Ok(None),
            },
        }
    }

    fn preview_member(&self) -> Result<Option<Preview<'_>>> {
        match &self.item {
            Item::Presentation => Ok(None),
            Item::Slide(slide) => Ok(Some(Preview::Text(Cow::Owned(slide_title(
                self.file,
                slide.position,
            )?)))),
            Item::Shape(shape) => shape_preview(shape),
            Item::Notes(notes) => Ok(Some(Preview::Prefix(Cow::Borrowed(&notes.text)))),
            Item::Paragraph(_) => Ok(self.text_member()?.map(Preview::Prefix)),
            Item::Element(element) => element.preview(self),
        }
    }

    fn child_count_member(&self) -> Result<Option<u32>> {
        let count = match &self.item {
            Item::Presentation => self.file.reader().slide_count(),
            Item::Slide(slide) => self
                .file
                .reader()
                .slide_at(slide.position)?
                .ok_or_else(|| missing_slide(slide.position))?
                .shape_count()?,
            Item::Shape(shape) => match resolve_shape(&shape.source, &shape.path) {
                Some(ShapeEnum::Group(group)) => group.children().len(),
                Some(native) => paragraph_count(native)?,
                None => 0,
            },
            Item::Notes(_) | Item::Paragraph(_) => 0,
            Item::Element(element) => element.child_count(self)?,
        };
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    pub(crate) fn format_member(&self) -> Result<Format<'_>> {
        let props = Props::new(self)?;
        let table = self.property_kind().spec().props;
        Ok(Format::new(table, Source::Ppt(props)))
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.type_name()
    }

    fn file_format(&self) -> FileFormat {
        FileFormat::Ppt
    }

    fn path(&self) -> NodePath<'_> {
        NodePath::new(self.steps())
    }

    fn positional_path(&self) -> NodePath<'_> {
        NodePath::new(self.steps())
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.check_source()?;
        let result = self.text_member();
        self.file.check_source().and(result)
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.check_source()?;
        let result = self.preview_member();
        self.file.check_source().and(result)
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        self.file.check_source()?;
        let result = self.child_count_member();
        self.file.check_source().and(result)
    }

    fn format(&self) -> Result<Format<'_>> {
        self.file.check_source()?;
        let result = self.format_member();
        self.file.check_source().and(result)
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        ViewChildren::ppt(Children::of(self, KindSet::ALL))
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Node")
            .field("type", &self.type_name())
            .field("slide", &self.slide())
            .finish_non_exhaustive()
    }
}

/// The values behind one node's `format` projection.
pub(crate) struct Props<'a> {
    kind: PropKind<'a>,
}

enum PropKind<'a> {
    Root {
        slide_count: usize,
    },
    /// Notes, whose only property is their text.
    Text(Cow<'a, str>),
    /// A paragraph, whose text is its node text, not a property.
    Paragraph,
    Slide {
        number: usize,
        slide_id: u32,
        persist_id: u32,
        hidden: Option<bool>,
    },
    Shape {
        kind: Kind,
        shape: &'a ShapeEnum<'static>,
    },
    /// A record-0323 element's values, in its table's order.
    Element(Values<'a>),
}

impl<'a> Props<'a> {
    fn new(node: &'a Node<'_>) -> Result<Self> {
        let kind = match &node.item {
            Item::Presentation => PropKind::Root {
                slide_count: node.file.reader().slide_count(),
            },
            Item::Slide(slide) => {
                let selected = node
                    .file
                    .reader()
                    .slide_at(slide.position)?
                    .ok_or_else(|| missing_slide(slide.position))?;
                PropKind::Slide {
                    number: selected.slide_number(),
                    slide_id: selected.slide_id(),
                    persist_id: selected.persist_id(),
                    hidden: selected.timing().map(|timing| timing.hidden),
                }
            },
            Item::Shape(shape) => PropKind::Shape {
                kind: node.property_kind(),
                shape: resolve_shape(&shape.source, &shape.path).ok_or_else(vanished)?,
            },
            Item::Notes(_) => PropKind::Text(node.text_member()?.ok_or_else(vanished)?),
            Item::Paragraph(_) => PropKind::Paragraph,
            Item::Element(element) => PropKind::Element(element.values(node)?),
        };
        Ok(Self { kind })
    }

    /// Read one property by its static table index.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match &self.kind {
            PropKind::Root { slide_count } => {
                match Kind::Presentation.spec().props.get(index)?.name {
                    "slideCount" => {
                        Some(Value::Int(i64::try_from(*slide_count).unwrap_or(i64::MAX)))
                    },
                    _ => None,
                }
            },
            PropKind::Slide {
                number,
                slide_id,
                persist_id,
                hidden,
            } => match Kind::Slide.spec().props.get(index)?.name {
                "slideNumber" => Some(Value::Int(i64::try_from(*number).unwrap_or(i64::MAX))),
                "slideId" => Some(Value::Int(i64::from(*slide_id))),
                "persistId" => Some(Value::Int(i64::from(*persist_id))),
                "hidden" => hidden.as_ref().map(|value| Value::Bool(*value)),
                _ => None,
            },
            PropKind::Shape { kind, shape } => shape_value(*kind, shape, index),
            PropKind::Text(text) => (index == 0).then(|| Value::Text(text.clone())),
            PropKind::Paragraph => None,
            PropKind::Element(values) => values.get(index),
        }
    }
}

/// The paragraphs of a shape's aggregate text, split at the paragraph marks
/// (`\r`) of the binary text atom. A text shape always has at least one
/// (possibly empty) paragraph.
pub(crate) fn paragraphs(text: &str) -> std::str::Split<'_, char> {
    text.split('\r')
}

/// The paragraph count of a text-capable shape; zero for other shapes.
pub(crate) fn paragraph_count(shape: &ShapeEnum<'static>) -> Result<usize> {
    if !has_text_body(shape) {
        return Ok(0);
    }
    let text = shape.text().map_err(Error::from)?;
    Ok(paragraphs(&text).count())
}

/// Whether a native shape holds a text body with paragraphs.
pub(crate) fn has_text_body(shape: &ShapeEnum<'_>) -> bool {
    matches!(
        shape,
        ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_)
    )
}

fn vanished() -> Error {
    Error::new(Code::InternalError, "shape vanished from the native scene")
        .with_format(FileFormat::Ppt)
}

fn shape_value<'a>(kind: Kind, shape: &'a ShapeEnum<'static>, index: usize) -> Option<Value<'a>> {
    let key = kind.spec().props.get(index)?.name;
    if matches!(key, "name" | "alt" | "sourceIndex") {
        return identity_value(shape, key);
    }
    match kind {
        Kind::Placeholder => placeholder_value(shape, key),
        Kind::Picture => picture_value(shape, key),
        Kind::Table => table_value(shape, key),
        Kind::Group => group_value(shape, key),
        Kind::Connector => connector_value(shape, key),
        Kind::Shape => generic_shape_value(shape, key),
        _ => None,
    }
}

/// The identity properties every shape kind shares: the `wzName` name, the
/// `wzDescription` alternative text and the stable source index.
fn identity_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    match key {
        "name" => shape_name(shape).map(|name| Value::Text(Cow::Borrowed(name))),
        "alt" => shape
            .description()
            .map(|alt| Value::Text(Cow::Borrowed(alt))),
        "sourceIndex" => shape.source_index().map(count),
        _ => None,
    }
}

fn generic_shape_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    match key {
        "id" => shape_id(shape).map(|id| Value::Int(i64::from(id))),
        "shapeType" => Some(Value::Text(Cow::Owned(shape.shape_type().to_string()))),
        "x" => common_properties(shape).map(|props| length(props.x)),
        "y" => common_properties(shape).map(|props| length(props.y)),
        "width" => common_properties(shape).map(|props| length(props.width)),
        "height" => common_properties(shape).map(|props| length(props.height)),
        "rotation" => {
            common_properties(shape).map(|props| Value::Number(f64::from(props.rotation)))
        },
        "hidden" => common_properties(shape).map(|props| Value::Bool(props.hidden)),
        _ => None,
    }
}

fn placeholder_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    let ShapeEnum::Placeholder(placeholder) = shape else {
        return None;
    };
    match key {
        "id" => Some(Value::Int(i64::from(placeholder.properties().id))),
        "placeholderType" => Some(Value::Text(Cow::Owned(
            placeholder.placeholder_type().to_string(),
        ))),
        "placeholderIndex" => placeholder
            .index()
            .map(|index| Value::Int(i64::from(index))),
        "x" => Some(length(placeholder.properties().x)),
        "y" => Some(length(placeholder.properties().y)),
        "width" => Some(length(placeholder.properties().width)),
        "height" => Some(length(placeholder.properties().height)),
        "rotation" => Some(Value::Number(f64::from(placeholder.properties().rotation))),
        "hidden" => Some(Value::Bool(placeholder.properties().hidden)),
        _ => None,
    }
}

fn picture_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    let ShapeEnum::Picture(picture) = shape else {
        return None;
    };
    match key {
        "id" => Some(Value::Int(i64::from(picture.properties().id))),
        "frameKind" => Some(Value::Text(Cow::Borrowed(match picture.frame_kind() {
            PictureFrameKind::Picture => "Picture",
            PictureFrameKind::OleObject => "OleObject",
            PictureFrameKind::Media => "Media",
        }))),
        "x" => Some(length(picture.properties().x)),
        "y" => Some(length(picture.properties().y)),
        "width" => Some(length(picture.properties().width)),
        "height" => Some(length(picture.properties().height)),
        "rotation" => Some(Value::Number(f64::from(picture.properties().rotation))),
        "hidden" => Some(Value::Bool(picture.properties().hidden)),
        _ => None,
    }
}

fn table_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    let table = shape.as_table()?;
    match key {
        "id" => Some(Value::Int(i64::from(table.id()))),
        "rows" => Some(count(table.rows())),
        "columns" => Some(count(table.columns())),
        "x" => Some(length(table.left())),
        "y" => Some(length(table.top())),
        "width" => Some(length(table.width())),
        "height" => Some(length(table.height())),
        _ => None,
    }
}

fn group_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    let group = shape.as_group()?;
    match key {
        "id" => Some(Value::Int(i64::from(group.id()))),
        "x" => Some(length(group.left())),
        "y" => Some(length(group.top())),
        "width" => Some(length(group.width())),
        "height" => Some(length(group.height())),
        _ => None,
    }
}

fn connector_value<'a>(shape: &'a ShapeEnum<'static>, key: &str) -> Option<Value<'a>> {
    let line = shape.as_line()?;
    match key {
        "id" => Some(Value::Int(i64::from(line.id()))),
        "shapeType" => Some(Value::Text(Cow::Owned(line.shape_type().to_string()))),
        "length" => Some(Value::Number(line.length())),
        _ => None,
    }
}

fn common_properties<'a>(shape: &'a ShapeEnum<'static>) -> Option<&'a ShapeProperties> {
    match shape {
        ShapeEnum::TextBox(shape) => Some(shape.properties()),
        ShapeEnum::Placeholder(shape) => Some(shape.properties()),
        ShapeEnum::AutoShape(shape) => Some(shape.properties()),
        ShapeEnum::Picture(shape) => Some(shape.properties()),
        ShapeEnum::Table(_) | ShapeEnum::Group(_) | ShapeEnum::Line(_) => None,
    }
}

fn resolve_shape<'a>(
    source: &'a [ShapeEnum<'static>],
    path: &[usize],
) -> Option<&'a ShapeEnum<'static>> {
    let (&first, rest) = path.split_first()?;
    let mut shape = source.get(first)?;
    for &index in rest {
        shape = shape.as_group()?.children().get(index)?;
    }
    Some(shape)
}

fn shape_id(shape: &ShapeEnum<'static>) -> Option<u32> {
    match shape {
        ShapeEnum::TextBox(shape) => Some(shape.properties().id),
        ShapeEnum::Placeholder(shape) => Some(shape.properties().id),
        ShapeEnum::AutoShape(shape) => Some(shape.properties().id),
        ShapeEnum::Picture(shape) => Some(shape.properties().id),
        ShapeEnum::Table(shape) => Some(shape.id()),
        ShapeEnum::Group(shape) => Some(shape.id()),
        ShapeEnum::Line(shape) => Some(shape.id()),
    }
}

/// The shape name (`wzName`); a picture without one falls back to the
/// file name its frame was authored with.
pub(crate) fn shape_name<'a>(shape: &'a ShapeEnum<'static>) -> Option<&'a str> {
    match shape {
        ShapeEnum::Picture(picture) => shape.name().or_else(|| picture.name()),
        _ => shape.name(),
    }
}

pub(crate) fn length(master: i32) -> Value<'static> {
    // Native PPT anchors use 1/576 inch; one master unit is 3175/2 EMUs.
    Value::Length(Length::from_emu(i64::from(master) * 3175 / 2))
}

pub(crate) fn count(value: usize) -> Value<'static> {
    Value::Int(i64::try_from(value).unwrap_or(i64::MAX))
}

fn shape_preview<'a>(shape: &'a ShapeItem) -> Result<Option<Preview<'a>>> {
    let native = resolve_shape(&shape.source, &shape.path).ok_or_else(|| {
        Error::new(Code::InternalError, "shape vanished from the native scene")
            .with_format(FileFormat::Ppt)
    })?;
    match native {
        ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_) => {
            let text = native.text().map_err(Error::from)?;
            if text.is_empty() {
                Ok(Some(Preview::Text(Cow::Borrowed("Shape"))))
            } else {
                // The text is owned by the native shape, so this temporary
                // allocation is retained by the preview value.
                Ok(Some(Preview::Prefix(Cow::Owned(text))))
            }
        },
        ShapeEnum::Picture(picture) => Ok(Some(Preview::Text(Cow::Owned(
            picture
                .name()
                .unwrap_or(match picture.frame_kind() {
                    PictureFrameKind::Picture => "Picture",
                    PictureFrameKind::OleObject => "OLE",
                    PictureFrameKind::Media => "Media",
                })
                .to_owned(),
        )))),
        ShapeEnum::Table(table) => Ok(Some(Preview::Text(Cow::Owned(format!(
            "Table ({}x{})",
            table.rows(),
            table.columns()
        ))))),
        ShapeEnum::Group(_) => Ok(Some(Preview::Text(Cow::Borrowed("Group")))),
        ShapeEnum::Line(line) => Ok(Some(Preview::Text(Cow::Borrowed(
            match line.shape_type() {
                litchi_ppt::shapes::ShapeType::Connector => "Connector",
                _ => "Line",
            },
        )))),
    }
}

fn slide_title(file: &Ppt, position: usize) -> Result<String> {
    let slide = file
        .reader()
        .slide_at(position)?
        .ok_or_else(|| missing_slide(position))?;
    for shape in slide.shapes()? {
        if let ShapeEnum::Placeholder(placeholder) = shape
            && placeholder.is_title()
        {
            let text = shape.text().map_err(Error::from)?;
            if !text.is_empty() {
                return Ok(text);
            }
        }
    }
    Ok(UNTITLED.to_owned())
}

fn missing_slide(position: usize) -> Error {
    Error::new(
        Code::InternalError,
        format!("slide {} vanished from the catalog", position + 1),
    )
    .with_format(FileFormat::Ppt)
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Props { .. }")
    }
}

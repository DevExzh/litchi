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
        match &self.item {
            Item::Presentation => None,
            Item::Slide(slide) => Some(Position::new(slide.position)),
            Item::Shape(shape) => Some(Position::new(shape.slide)),
        }
    }

    pub(crate) fn slide_position(&self) -> Option<usize> {
        match &self.item {
            Item::Presentation => None,
            Item::Slide(slide) => Some(slide.position),
            Item::Shape(shape) => Some(shape.slide),
        }
    }

    /// The native shape carried by a shape node.
    #[must_use]
    pub fn shape(&self) -> Option<&ShapeEnum<'static>> {
        match &self.item {
            Item::Shape(shape) => resolve_shape(&shape.source, &shape.path),
            Item::Presentation | Item::Slide(_) => None,
        }
    }

    /// The native picture name, where the leaf exposes one.
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
        }
    }

    /// The semantic kind used by query inspection.  Placeholders retain
    /// their physical shape node and path, but are a distinct query subject.
    pub(crate) fn semantic_kind(&self) -> Kind {
        match self.shape() {
            Some(ShapeEnum::Placeholder(_)) => Kind::Placeholder,
            _ => self.element(),
        }
    }

    /// The property table kind.  A placeholder remains physically a shape in
    /// the native tree, but has its own semantic properties.
    pub(crate) fn property_kind(&self) -> Kind {
        match self.shape() {
            Some(ShapeEnum::Placeholder(_)) => Kind::Placeholder,
            _ => self.element(),
        }
    }

    /// The OfficeCLI semantic type name.
    pub(crate) fn type_name(&self) -> &'static str {
        match self.shape() {
            Some(ShapeEnum::Placeholder(placeholder)) if placeholder.is_title() => "title",
            Some(ShapeEnum::Placeholder(_)) => "placeholder",
            _ => self.element().name(),
        }
    }

    /// The node's text member.  Legacy PPT exposes aggregate text for text
    /// shapes, including the native aggregate table/group text where the
    /// leaf provides it.
    pub(crate) fn text_member(&self) -> Result<Option<Cow<'_, str>>> {
        match self.shape() {
            Some(shape) if matches!(self.element(), Kind::Shape | Kind::Table | Kind::Group) => {
                shape
                    .text()
                    .map(|text| Some(Cow::Owned(text)))
                    .map_err(Error::from)
            },
            _ => Ok(None),
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
                _ => 0,
            },
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
    Slide {
        number: usize,
        slide_id: u32,
        persist_id: u32,
    },
    Shape {
        kind: Kind,
        shape: &'a ShapeEnum<'static>,
    },
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
                }
            },
            Item::Shape(shape) => PropKind::Shape {
                kind: node.property_kind(),
                shape: resolve_shape(&shape.source, &shape.path).ok_or_else(|| {
                    Error::new(Code::InternalError, "shape vanished from the native scene")
                        .with_format(FileFormat::Ppt)
                })?,
            },
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
            } => match Kind::Slide.spec().props.get(index)?.name {
                "slideNumber" => Some(Value::Int(i64::try_from(*number).unwrap_or(i64::MAX))),
                "slideId" => Some(Value::Int(i64::from(*slide_id))),
                "persistId" => Some(Value::Int(i64::from(*persist_id))),
                _ => None,
            },
            PropKind::Shape { kind, shape } => shape_value(*kind, shape, index),
        }
    }
}

fn shape_value<'a>(kind: Kind, shape: &'a ShapeEnum<'static>, index: usize) -> Option<Value<'a>> {
    let key = kind.spec().props.get(index)?.name;
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
        "name" => picture.name().map(|name| Value::Text(Cow::Borrowed(name))),
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

fn shape_name<'a>(shape: &'a ShapeEnum<'static>) -> Option<&'a str> {
    match shape {
        ShapeEnum::Picture(picture) => picture.name(),
        _ => None,
    }
}

fn length(master: i32) -> Value<'static> {
    // Native PPT anchors use 1/576 inch; one master unit is 3175/2 EMUs.
    Value::Length(Length::from_emu(i64::from(master) * 3175 / 2))
}

fn count(value: usize) -> Value<'static> {
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

//! Typed read nodes for OpenDocument presentations.
//!
//! A node keeps the handle, the scene index it was resolved in, a compact
//! address into that index and its canonical path steps. Text, previews
//! and property values are read from the indexed producer markup on demand.

use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};
use litchi_odp::scene::{
    FrameContent, Object, ObjectKind, ObjectRef, Scene, Scope, Story, StoryRef,
};

use super::Odp;
use super::children::Children;
use super::spec::{CHARACTER_KEYS, Kind};
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::path::navigate::KindSet;
use crate::value::Length;

const UNTITLED: &str = "(untitled)";

/// A node of an OpenDocument presentation.
#[derive(Clone)]
pub struct Node<'f> {
    pub(super) file: &'f Odp,
    pub(super) scene: &'f Scene,
    pub(super) item: Item,
    pub(super) steps: Arc<[PathStep<'static>]>,
}

/// The scene address a node carries.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub(crate) enum Item {
    Presentation,
    Slide(usize),
    /// A drawing object, with its physical kind.
    Object(ObjectRef, Kind),
    /// A slide's speaker notes (the notes page's text frame).
    Notes(usize),
    /// A paragraph of a story.
    Paragraph(StoryRef, usize),
    /// A run of a paragraph.
    Run(StoryRef, usize, usize),
    /// A table row.
    Row(ObjectRef, usize),
    /// A table cell.
    Cell(ObjectRef, usize, usize),
    /// An annotation: the index in the presentation's annotation inventory.
    Comment(usize),
    /// A master page.
    Master(usize),
    /// A presentation page layout.
    Layout(usize),
    /// A shape's hyperlink: its click action, else its first link run's
    /// target (record 0328).
    ShapeLink(ObjectRef),
    /// A link run's hyperlink.
    RunLink(StoryRef, usize, usize),
    /// A paragraph's `text:line-break`.
    LineBreak(StoryRef, usize, usize),
    /// A declared table column.
    Column(ObjectRef, usize),
    /// A chart object's series.
    Series(ObjectRef, usize),
    /// A chart object's axis.
    Axis(ObjectRef, usize),
}

/// The physical kind of a drawing object.
pub(crate) fn object_kind(object: &Object) -> Kind {
    match object.kind() {
        ObjectKind::Frame => match object.content() {
            FrameContent::Image => Kind::Picture,
            FrameContent::Table => Kind::Table,
            FrameContent::Object if object.class() == Some("chart") => Kind::Chart,
            FrameContent::Object => Kind::Ole,
            FrameContent::Plugin => Kind::Media,
            _ => Kind::Shape,
        },
        ObjectKind::Line | ObjectKind::Connector => Kind::Connector,
        ObjectKind::Group => Kind::Group,
        _ => Kind::Shape,
    }
}

/// Whether a shape is a presentation placeholder.
pub(crate) fn is_placeholder(object: &Object) -> bool {
    object.class().is_some() && object_kind(object) == Kind::Shape
}

/// The OfficeCLI type name of an object.
fn object_type(object: &Object, kind: Kind) -> &'static str {
    match kind {
        Kind::Shape if object.class() == Some("title") => "title",
        Kind::Shape if object.class().is_some() => "placeholder",
        kind => kind.name(),
    }
}

/// The notes page's story address of slide `page`.
pub(crate) fn notes_story(scene: &Scene, page: usize) -> Option<StoryRef> {
    let notes = scene.page(page)?.notes()?;
    let object = notes.text_object()?;
    Some(StoryRef::object(ObjectRef {
        page,
        scope: Scope::Notes,
        path: vec![object],
    }))
}

impl<'f> Node<'f> {
    pub(super) fn new(
        file: &'f Odp,
        scene: &'f Scene,
        item: Item,
        steps: Arc<[PathStep<'static>]>,
    ) -> Self {
        Self {
            file,
            scene,
            item,
            steps,
        }
    }

    pub(super) fn root(file: &'f Odp, scene: &'f Scene) -> Self {
        Self::new(file, scene, Item::Presentation, Arc::from(Vec::new()))
    }

    /// A child of `self` addressed by `item`, at `name[ordinal]`.
    pub(super) fn child(&self, item: Item, name: &'static str, ordinal: Option<usize>) -> Self {
        let mut steps = Vec::with_capacity(self.steps.len() + 1);
        steps.extend_from_slice(&self.steps);
        steps.push(match ordinal {
            Some(ordinal) => PathStep::at(name, Position::new(ordinal)),
            None => PathStep::new(name),
        });
        Self::new(self.file, self.scene, item, Arc::from(steps))
    }

    pub(crate) fn file(&self) -> &'f Odp {
        self.file
    }

    pub(crate) fn scene(&self) -> &'f Scene {
        self.scene
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
            Item::Slide(page) | Item::Notes(page) => Some(*page),
            Item::Object(object, _) | Item::Row(object, _) | Item::Cell(object, _, _) => {
                Some(object.page)
            },
            Item::Paragraph(story, _)
            | Item::Run(story, _, _)
            | Item::RunLink(story, _, _)
            | Item::LineBreak(story, _, _) => Some(story.object.page),
            Item::ShapeLink(object)
            | Item::Column(object, _)
            | Item::Series(object, _)
            | Item::Axis(object, _) => Some(object.page),
            Item::Comment(_) | Item::Presentation | Item::Master(_) | Item::Layout(_) => None,
        }
    }

    /// The drawing object a node addresses or belongs to.
    pub(crate) fn object(&self) -> Option<&'f Object> {
        let reference = match &self.item {
            Item::Object(object, _)
            | Item::Row(object, _)
            | Item::Cell(object, _, _)
            | Item::ShapeLink(object)
            | Item::Column(object, _)
            | Item::Series(object, _)
            | Item::Axis(object, _) => object,
            Item::Paragraph(story, _)
            | Item::Run(story, _, _)
            | Item::RunLink(story, _, _)
            | Item::LineBreak(story, _, _) => &story.object,
            _ => return None,
        };
        self.scene.object(reference)
    }

    /// The story a story-bearing node (shape, placeholder, cell, notes)
    /// owns.
    pub(crate) fn own_story(&self) -> Option<StoryRef> {
        match &self.item {
            Item::Object(object, Kind::Shape) => {
                let story = StoryRef::object(object.clone());
                self.scene.story(&story).map(|_| story)
            },
            Item::Cell(object, row, cell) => Some(StoryRef {
                object: object.clone(),
                cell: Some((*row, *cell)),
            }),
            Item::Notes(page) => notes_story(self.scene, *page),
            _ => None,
        }
    }

    fn story(&self, story: &StoryRef) -> Option<&'f Story> {
        self.scene.story(story)
    }

    /// The physical kind used by navigation and ordinals.
    pub(crate) fn element(&self) -> Kind {
        match &self.item {
            Item::Presentation => Kind::Presentation,
            Item::Slide(_) => Kind::Slide,
            Item::Object(_, kind) => *kind,
            Item::Notes(_) => Kind::Notes,
            Item::Paragraph(..) => Kind::Paragraph,
            Item::Run(..) => Kind::Run,
            Item::Row(..) => Kind::Row,
            Item::Cell(..) => Kind::Cell,
            Item::Comment(_) => Kind::Comment,
            Item::Master(_) => Kind::SlideMaster,
            Item::Layout(_) => Kind::SlideLayout,
            Item::ShapeLink(_) | Item::RunLink(..) => Kind::Hyperlink,
            Item::LineBreak(..) => Kind::LineBreak,
            Item::Column(..) => Kind::Col,
            Item::Series(..) => Kind::Series,
            Item::Axis(..) => Kind::Axis,
        }
    }

    /// The kind queries match: a presentation placeholder is a placeholder.
    pub(crate) fn semantic_kind(&self) -> Kind {
        match (&self.item, self.object()) {
            (Item::Object(_, Kind::Shape), Some(object)) if is_placeholder(object) => {
                Kind::Placeholder
            },
            _ => self.element(),
        }
    }

    /// The OfficeCLI type name.
    pub(crate) fn type_name(&self) -> &'static str {
        match (&self.item, self.object()) {
            (Item::Object(_, kind), Some(object)) => object_type(object, *kind),
            _ => self.element().name(),
        }
    }

    /// The node's text member.
    pub(crate) fn text_member(&self) -> Result<Option<Cow<'f, str>>> {
        Ok(match &self.item {
            Item::Object(_, Kind::Shape) => self.object().map(|object| Cow::Owned(object.text())),
            Item::Object(_, Kind::Table) => self.object().and_then(Object::table).map(|table| {
                let mut text = String::new();
                for row in table.rows() {
                    for cell in row.cells() {
                        let cell = cell.text();
                        if !cell.is_empty() {
                            if !text.is_empty() {
                                text.push('\n');
                            }
                            text.push_str(&cell);
                        }
                    }
                }
                Cow::Owned(text)
            }),
            Item::Notes(page) => {
                let notes = self.scene.page(*page).and_then(|page| page.notes());
                notes
                    .and_then(litchi_odp::scene::Notes::text)
                    .map(Cow::Owned)
            },
            Item::Paragraph(story, index) => self
                .story(story)
                .and_then(|story| story.paragraphs().get(*index))
                .map(|paragraph| Cow::Borrowed(paragraph.text())),
            Item::Run(story, paragraph, index) => self
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .and_then(|paragraph| paragraph.runs().get(*index))
                .map(|run| Cow::Borrowed(run.text())),
            Item::Cell(object, row, cell) => self
                .scene
                .object(object)
                .and_then(Object::table)
                .and_then(|table| table.rows().get(*row))
                .and_then(|row| row.cells().get(*cell))
                .map(|cell| Cow::Owned(cell.text())),
            Item::Comment(index) => self
                .file
                .annotations()?
                .get(*index)
                .map(|info| Cow::Owned(annotation_text(info))),
            Item::RunLink(story, paragraph, run) => self
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .and_then(|paragraph| paragraph.runs().get(*run))
                .map(|run| Cow::Borrowed(run.text())),
            Item::Object(reference, Kind::Equation) => self
                .file
                .parts()?
                .formula(reference)
                .map(|formula| Cow::Owned(formula.to_owned())),
            Item::Axis(object, index) => self
                .file
                .parts()?
                .chart(object)
                .and_then(|chart| chart.axes.get(*index))
                .and_then(|axis| axis.title.clone())
                .map(Cow::Owned),
            _ => None,
        })
    }

    fn preview_member(&self) -> Result<Option<Preview<'f>>> {
        Ok(match &self.item {
            Item::Presentation | Item::Row(..) => None,
            Item::Slide(page) => Some(Preview::Text(Cow::Owned(
                self.scene
                    .page(*page)
                    .and_then(litchi_odp::scene::Page::title)
                    .filter(|title| !title.is_empty())
                    .unwrap_or_else(|| UNTITLED.to_owned()),
            ))),
            Item::Object(_, kind) => {
                let Some(object) = self.object() else {
                    return Ok(None);
                };
                let label = |fallback: &'static str| {
                    Preview::Text(
                        object
                            .name()
                            .map_or(Cow::Borrowed(fallback), |name| Cow::Owned(name.to_owned())),
                    )
                };
                Some(match kind {
                    Kind::Shape => {
                        let text = object.text();
                        if text.is_empty() {
                            label("Shape")
                        } else {
                            Preview::Prefix(Cow::Owned(text))
                        }
                    },
                    Kind::Picture => label("Picture"),
                    Kind::Table => {
                        let (rows, columns) = object
                            .table()
                            .map_or((0, 0), |table| (table.rows().len(), table.columns()));
                        Preview::Text(Cow::Owned(format!("Table ({rows}x{columns})")))
                    },
                    Kind::Chart => label("Chart"),
                    Kind::Ole => label("OLE"),
                    Kind::Media => label("Media"),
                    Kind::Group => label("Group"),
                    Kind::Connector => label("Connector"),
                    _ => label("Shape"),
                })
            },
            Item::Master(index) => self
                .file
                .masters()?
                .get(*index)
                .map(|master| Preview::Text(Cow::Owned(master.name.clone()))),
            Item::Layout(index) => self
                .file
                .layouts()?
                .get(*index)
                .map(|layout| Preview::Text(Cow::Owned(layout.name.clone()))),
            _ => self.text_member()?.map(Preview::Prefix),
        })
    }

    fn child_count_member(&self) -> Result<usize> {
        Ok(match &self.item {
            Item::Presentation => {
                self.scene.pages().len() + self.file.masters()?.len() + self.file.layouts()?.len()
            },
            Item::Slide(page) => {
                let Some(slide) = self.scene.page(*page) else {
                    return Ok(0);
                };
                let notes = usize::from(notes_story(self.scene, *page).is_some());
                let comments = page_comments(self.file, *page)?.count();
                slide.objects().len() + notes + comments
            },
            Item::Object(_, Kind::Group) => {
                self.object().map_or(0, |object| object.children().len())
            },
            Item::Object(_, Kind::Table) => self
                .object()
                .and_then(Object::table)
                .map_or(0, |table| table.rows().len()),
            Item::Object(..) | Item::Notes(_) | Item::Cell(..) => self
                .own_story()
                .and_then(|story| self.story(&story))
                .map_or(0, |story| story.paragraphs().len()),
            Item::Paragraph(story, index) => self
                .story(story)
                .and_then(|story| story.paragraphs().get(*index))
                .map_or(0, |paragraph| paragraph.runs().len()),
            Item::Row(object, row) => self
                .scene
                .object(object)
                .and_then(Object::table)
                .and_then(|table| table.rows().get(*row))
                .map_or(0, |row| row.cells().len()),
            Item::Run(..)
            | Item::Comment(_)
            | Item::Master(_)
            | Item::Layout(_)
            | Item::ShapeLink(_)
            | Item::RunLink(..)
            | Item::LineBreak(..)
            | Item::Column(..)
            | Item::Series(..)
            | Item::Axis(..) => 0,
        })
    }

    pub(crate) fn format_member(&self) -> Result<Format<'_>> {
        let props = Props::new(self)?;
        let table = self.semantic_kind().spec().props;
        Ok(Format::new(table, Source::Odp(props)))
    }
}

/// The text of an annotation: its paragraphs joined by `\n`.
pub(crate) fn annotation_text(info: &litchi_odp::annotation::Info) -> String {
    let mut text = String::new();
    for child in info.annotation.children() {
        if child.name().ends_with(":p") || child.name().ends_with(":h") {
            if !text.is_empty() {
                text.push('\n');
            }
            text.push_str(&child.plain_text());
        }
    }
    text
}

/// The annotations of slide `page`, as indices into the inventory.
pub(crate) fn page_comments(file: &Odp, page: usize) -> Result<impl Iterator<Item = usize> + '_> {
    Ok(file
        .annotations()?
        .iter()
        .enumerate()
        .filter(move |(_, info)| match info.anchor.position() {
            litchi_odp::annotation::Position::Page { index } => *index == page,
            litchi_odp::annotation::Position::Shape { page_index, .. } => *page_index == page,
            _ => false,
        })
        .map(|(index, _)| index))
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        self.type_name()
    }

    fn file_format(&self) -> FileFormat {
        FileFormat::Odp
    }

    fn path(&self) -> NodePath<'_> {
        NodePath::new(&self.steps)
    }

    fn positional_path(&self) -> NodePath<'_> {
        NodePath::new(&self.steps)
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.check_source()?;
        self.text_member()
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.check_source()?;
        self.preview_member()
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        self.file.check_source()?;
        let count = self.child_count_member()?;
        Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
    }

    fn format(&self) -> Result<Format<'_>> {
        self.file.check_source()?;
        self.format_member()
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        match Children::of(self, KindSet::ALL) {
            Ok(children) => ViewChildren::odp(children),
            Err(error) => ViewChildren::failed(error),
        }
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("odp::Node")
            .field("type", &self.type_name())
            .field("slide", &self.slide())
            .finish_non_exhaustive()
    }
}

/// The values behind one node's `format` projection.
pub(crate) struct Props<'a> {
    kind: Kind,
    values: PropValues<'a>,
    /// The character properties a story's runs share, resolved once on the
    /// first such read (record 0316), in `CHARACTER_KEYS` order.
    /// Boxed: most nodes never read it, and the projection stays small.
    shared: std::cell::OnceCell<Box<[Option<Value<'static>>; 6]>>,
}

enum PropValues<'a> {
    None,
    Root {
        slides: usize,
        metadata: Option<Box<litchi_core::Metadata>>,
    },
    Slide {
        page: &'a litchi_odp::scene::Page,
        number: usize,
    },
    /// An object, with its text story for the shared character
    /// formatting read on demand (record 0316).
    Object(&'a Object, Option<&'a Story>, &'a Odp),
    /// A paragraph or run with the character properties its style chain
    /// resolves to (record 0316).
    Paragraph(
        &'a litchi_odp::scene::Paragraph,
        litchi_odp::scene::TextProperties,
    ),
    Run(
        &'a litchi_odp::scene::Run,
        litchi_odp::scene::TextProperties,
    ),
    Cell(&'a litchi_odp::scene::Cell, &'a Odp),
    Notes(Option<String>, Option<&'a Story>, &'a Odp),
    Comment(&'a litchi_odp::annotation::Info),
    Master(&'a litchi_odp::master::Summary),
    Layout(&'a litchi_odp::layout::Layout),
    /// A video, audio or formula object (record 0328).
    Special {
        object: &'a Object,
        reference: &'a ObjectRef,
        kind: Kind,
        file: &'a Odp,
    },
    /// A hyperlink: its `link` and `tooltip` (boxed, so the common
    /// variants stay as small as they were).
    Link(Box<(Option<String>, Option<String>)>),
    /// A declared table column's width.
    Column(Option<Length>),
    /// A chart series.
    Series(&'a super::parts::Series),
    /// A chart axis.
    Axis(&'a super::parts::Axis),
}

impl<'a> Props<'a> {
    fn new(node: &'a Node<'_>) -> Result<Self> {
        let scene = node.scene;
        let values = match &node.item {
            Item::Presentation => PropValues::Root {
                slides: scene.pages().len(),
                metadata: node.file.metadata()?.map(Box::new),
            },
            Item::Slide(page) => match scene.page(*page) {
                Some(slide) => PropValues::Slide {
                    page: slide,
                    number: page + 1,
                },
                None => PropValues::None,
            },
            Item::Object(reference, kind @ (Kind::Video | Kind::Audio | Kind::Equation)) => node
                .object()
                .map_or(PropValues::None, |object| PropValues::Special {
                    object,
                    reference,
                    kind: *kind,
                    file: node.file,
                }),
            Item::Object(..) => node.object().map_or(PropValues::None, |object| {
                let story = node.own_story().and_then(|story| scene.story(&story));
                PropValues::Object(object, story, node.file)
            }),
            Item::ShapeLink(_) | Item::RunLink(..) => {
                let (link, tooltip) = link_values(node)?;
                PropValues::Link(Box::new((link, tooltip)))
            },
            Item::Column(object, index) => {
                let xml = node.file.content_xml()?;
                let styles = node.file.automatic_styles()?;
                let width = scene
                    .object(object)
                    .and_then(Object::table)
                    .and_then(|table| {
                        litchi_odp::scene::parts::table_columns(xml, table)
                            .into_iter()
                            .flat_map(|column| {
                                std::iter::repeat_n(column.clone(), column.repeat() as usize)
                            })
                            .nth(*index)
                    })
                    .and_then(|column| column.width(xml, styles))
                    .and_then(|width| Length::parse(width.trim()).ok());
                PropValues::Column(width)
            },
            Item::Series(object, index) => node
                .file
                .parts()?
                .chart(object)
                .and_then(|chart| chart.series.get(*index))
                .map_or(PropValues::None, PropValues::Series),
            Item::Axis(object, index) => node
                .file
                .parts()?
                .chart(object)
                .and_then(|chart| chart.axes.get(*index))
                .map_or(PropValues::None, PropValues::Axis),
            Item::LineBreak(..) => PropValues::None,
            Item::Paragraph(story, index) => scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*index))
                .map_or(Ok::<_, crate::Error>(PropValues::None), |paragraph| {
                    let resolved = node
                        .file
                        .automatic_styles()?
                        .resolve_paragraph(node.file.common_styles()?, paragraph);
                    Ok(PropValues::Paragraph(paragraph, resolved))
                })?,
            Item::Run(story, paragraph, index) => scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .and_then(|paragraph| Some((paragraph, paragraph.runs().get(*index)?)))
                .map_or(
                    Ok::<_, crate::Error>(PropValues::None),
                    |(paragraph, run)| {
                        // A run inherits what its own style leaves unset from
                        // its paragraph (record 0316).
                        let styles = node.file.automatic_styles()?;
                        let common = node.file.common_styles()?;
                        let mut resolved = styles.resolve_run(common, run);
                        resolved.inherit(&styles.resolve_paragraph(common, paragraph));
                        Ok(PropValues::Run(run, resolved))
                    },
                )?,
            Item::Cell(object, row, cell) => scene
                .object(object)
                .and_then(Object::table)
                .and_then(|table| table.rows().get(*row))
                .and_then(|row| row.cells().get(*cell))
                .map_or(PropValues::None, |cell| PropValues::Cell(cell, node.file)),
            Item::Notes(page) => PropValues::Notes(
                scene
                    .page(*page)
                    .and_then(|page| page.notes())
                    .and_then(litchi_odp::scene::Notes::text),
                notes_story(scene, *page).and_then(|story| scene.story(&story)),
                node.file,
            ),
            Item::Comment(index) => node
                .file
                .annotations()?
                .get(*index)
                .map_or(PropValues::None, PropValues::Comment),
            Item::Master(index) => node
                .file
                .masters()?
                .get(*index)
                .map_or(PropValues::None, PropValues::Master),
            Item::Layout(index) => node
                .file
                .layouts()?
                .get(*index)
                .map_or(PropValues::None, PropValues::Layout),
            Item::Row(..) => PropValues::None,
        };
        Ok(Self {
            kind: node.semantic_kind(),
            values,
            shared: std::cell::OnceCell::new(),
        })
    }

    /// Read one property by its static table index.
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        let key = self.kind.spec().props.get(index)?.name;
        let text = |value: Option<&'a str>| value.map(|value| Value::Text(Cow::Borrowed(value)));
        let owned = |value: Option<String>| value.map(|value| Value::Text(Cow::Owned(value)));
        match &self.values {
            PropValues::None => None,
            PropValues::Root { slides, metadata } => match key {
                "slideCount" => Some(count(*slides)),
                "title" => owned(metadata.as_ref()?.title.clone()),
                "author" => owned(metadata.as_ref()?.author.clone()),
                "subject" => owned(metadata.as_ref()?.subject.clone()),
                "keywords" => Value::owned_list(&metadata.as_ref()?.keywords),
                "description" => owned(metadata.as_ref()?.description.clone()),
                _ => None,
            },
            PropValues::Slide { page, number } => match key {
                "name" => text(page.name()),
                "slideNumber" => Some(count(*number)),
                "title" => owned(page.title()),
                "layout" => text(page.layout()),
                "master" => text(page.master_page()),
                _ => None,
            },
            PropValues::Object(object, story, file) => match object_value(object, key) {
                Some(value) => Some(value),
                None if CHARACTER_KEYS.contains(&key) => self.shared(file, (*story)?, key),
                None => None,
            },
            PropValues::Paragraph(paragraph, properties) => match key {
                "style" => text(paragraph.style_name()),
                "level" => Some(Value::Int(i64::from(paragraph.level()))),
                key => character(properties, key).map(Value::into_owned),
            },
            PropValues::Run(run, properties) => match key {
                "style" => text(run.style_name()),
                "link" => text(run.href()),
                key => character(properties, key).map(Value::into_owned),
            },
            PropValues::Cell(cell, file) => match key {
                "colSpan" => Some(Value::Int(i64::from(cell.column_span()))),
                "rowSpan" => Some(Value::Int(i64::from(cell.row_span()))),
                key => self.shared(file, cell.story(), key),
            },
            PropValues::Notes(notes, story, file) => match key {
                "text" => owned(notes.clone()),
                key => self.shared(file, (*story)?, key),
            },
            PropValues::Comment(info) => match key {
                "text" => Some(Value::Text(Cow::Owned(annotation_text(info)))),
                "author" => owned(info.annotation.creator()),
                "initials" => owned(info.annotation.initials()),
                "date" => owned(info.annotation.date()),
                _ => None,
            },
            PropValues::Master(master) => match key {
                "name" => Some(Value::Text(Cow::Borrowed(&master.name))),
                "displayName" => text(master.display_name.as_deref()),
                "pageLayout" => text(master.page_layout_name.as_deref()),
                _ => None,
            },
            PropValues::Layout(layout) => match key {
                "name" => Some(Value::Text(Cow::Borrowed(&layout.name))),
                "displayName" => text(layout.display_name.as_deref()),
                _ => None,
            },
            PropValues::Special {
                object,
                reference,
                kind,
                file,
            } => special_value(file, reference, object, *kind, key),
            PropValues::Link(link) => match key {
                "link" => owned(link.0.clone()),
                "tooltip" => owned(link.1.clone()),
                _ => None,
            },
            PropValues::Column(width) => match key {
                "width" => width.map(Value::Length),
                _ => None,
            },
            PropValues::Series(series) => match key {
                "values" => text(series.values.as_deref()),
                "nameRef" => text(series.label.as_deref()),
                "chartType" => text(series.class.as_deref()),
                "axis" => text(series.axis.as_deref()),
                "pointCount" => Some(count(series.points)),
                _ => None,
            },
            PropValues::Axis(axis) => match key {
                "role" => axis.role().map(|role| Value::Text(Cow::Borrowed(role))),
                "title" => text(axis.title.as_deref()),
                "dimension" => axis
                    .dimension
                    .map(|dimension| Value::Text(Cow::Borrowed(dimension))),
                "name" => text(axis.name.as_deref()),
                "categories" => text(axis.categories.as_deref()),
                "style" => text(axis.style.as_deref()),
                _ => None,
            },
        }
    }
}

/// A video, audio or formula object's value: the frame keys, then the
/// media or the formula.
fn special_value<'a>(
    file: &'a Odp,
    reference: &ObjectRef,
    object: &'a Object,
    kind: Kind,
    key: &str,
) -> Option<Value<'a>> {
    if let Some(value) = object_value(object, key) {
        return Some(value);
    }
    if kind == Kind::Equation {
        return match key {
            "formula" => file
                .parts()
                .ok()?
                .formula(reference)
                .map(|formula| Value::Text(Cow::Owned(formula.to_owned()))),
            _ => None,
        };
    }
    let plugin = litchi_odp::scene::parts::plugin(file.content_xml().ok()?, object)?;
    let flag = |name: &str| {
        plugin
            .parameter(name)
            .map(|value| Value::Bool(value.eq_ignore_ascii_case("true")))
    };
    match key {
        "mediaType" => Some(Value::Text(Cow::Borrowed(if kind == Kind::Video {
            "video"
        } else {
            "audio"
        }))),
        "mimeType" => plugin
            .mime_type()
            .map(|mime| Value::Text(Cow::Owned(mime.to_owned()))),
        "loop" => flag("Loop"),
        "mute" => flag("Mute"),
        _ => None,
    }
}

/// A hyperlink node's `link` and `tooltip`: a shape's click action (a URL,
/// `slide[N]` or a jump name), else its first link run's; a link run's
/// target and `office:title`.
fn link_values(node: &Node<'_>) -> Result<(Option<String>, Option<String>)> {
    let xml = node.file.content_xml()?;
    let scene = node.scene;
    let run_values = |run: &litchi_odp::scene::Run| {
        (
            run.href().map(str::to_owned),
            litchi_odp::scene::parts::link_title(xml, run),
        )
    };
    Ok(match &node.item {
        Item::RunLink(story, paragraph, run) => scene
            .story(story)
            .and_then(|story| story.paragraphs().get(*paragraph))
            .and_then(|paragraph| paragraph.runs().get(*run))
            .map_or((None, None), run_values),
        Item::ShapeLink(object) => {
            let Some(shape) = scene.object(object) else {
                return Ok((None, None));
            };
            if let Some(click) = litchi_odp::scene::parts::shape_click(xml, shape) {
                (super::parts::click_link(scene, &click), None)
            } else {
                shape
                    .story()
                    .and_then(first_link_run)
                    .map_or((None, None), run_values)
            }
        },
        _ => (None, None),
    })
}

/// The first link run of a story.
pub(crate) fn first_link_run(story: &Story) -> Option<&litchi_odp::scene::Run> {
    story
        .paragraphs()
        .iter()
        .flat_map(|paragraph| paragraph.runs())
        .find(|run| run.kind() == litchi_odp::scene::RunKind::Link)
}

#[inline]
fn object_value<'a>(object: &'a Object, key: &str) -> Option<Value<'a>> {
    let text = |value: Option<&'a str>| value.map(|value| Value::Text(Cow::Borrowed(value)));
    match key {
        "name" => text(object.name()),
        "id" => text(object.id()),
        "alt" => text(object.title().or(object.description())),
        "shapeType" => Some(Value::Text(Cow::Borrowed(shape_type(object)))),
        "placeholderType" => text(object.class()),
        "isTitle" => Some(Value::Bool(object.class() == Some("title"))),
        "source" => text(object.href()),
        "x" => length(object.x()),
        "y" => length(object.y()),
        "width" => length(object.width()),
        "height" => length(object.height()),
        "zorder" => object
            .z_index()
            .and_then(|value| value.trim().parse::<i64>().ok())
            .map(Value::Int),
        "style" => match object.table() {
            Some(table) => text(table.style_name()),
            None => text(object.style_name()),
        },
        "rows" => object.table().map(|table| count(table.rows().len())),
        "cols" => object.table().map(|table| count(table.columns())),
        _ => None,
    }
}

/// The ODF drawing element name of a shape.
fn shape_type(object: &Object) -> &'static str {
    match object.kind() {
        ObjectKind::Frame => match object.content() {
            FrameContent::TextBox => "textBox",
            _ => "frame",
        },
        ObjectKind::Rectangle => "rect",
        ObjectKind::Ellipse => "ellipse",
        ObjectKind::Line => "line",
        ObjectKind::Connector => "connector",
        ObjectKind::CustomShape => "customShape",
        ObjectKind::Polygon => "polygon",
        ObjectKind::Annotation => "caption",
        ObjectKind::Control => "control",
        ObjectKind::Group => "group",
        ObjectKind::PageThumbnail => "pageThumbnail",
        ObjectKind::Scene3d => "scene3d",
        _ => "shape",
    }
}

/// An ODF length attribute as a unified length.
fn length(value: Option<&str>) -> Option<Value<'static>> {
    Length::parse(value?.trim()).ok().map(Value::Length)
}

fn count(value: usize) -> Value<'static> {
    Value::Int(i64::try_from(value).unwrap_or(i64::MAX))
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Props { .. }")
    }
}

impl Props<'_> {
    /// The character property `key` every run of `story` shares.
    fn shared(&self, file: &Odp, story: &Story, key: &str) -> Option<Value<'static>> {
        let index = CHARACTER_KEYS
            .iter()
            .position(|candidate| *candidate == key)?;
        self.shared
            .get_or_init(|| Box::new(shared_characters(file, story)))[index]
            .clone()
    }
}

/// The character properties every run of `story` shares (record 0316), in
/// `CHARACTER_KEYS` order: each run's style chain over its paragraph's; a
/// paragraph without runs counts by its own. A property is `None` when the
/// runs differ or none sets it.
fn shared_characters(file: &Odp, story: &Story) -> [Option<Value<'static>>; 6] {
    let (Ok(styles), Ok(common)) = (file.automatic_styles(), file.common_styles()) else {
        return Default::default();
    };
    let mut shared: [Option<Value<'static>>; 6] = Default::default();
    let mut agreed = [true; 6];
    let mut seen = false;
    let mut each = |properties: &litchi_odp::scene::TextProperties| {
        for (index, key) in CHARACTER_KEYS.iter().enumerate() {
            if !agreed[index] {
                continue;
            }
            let value = character(properties, key).map(Value::into_owned);
            if seen && value != shared[index] {
                agreed[index] = false;
                shared[index] = None;
            } else {
                shared[index] = value;
            }
        }
        seen = true;
    };
    for paragraph in story.paragraphs() {
        let base = styles.resolve_paragraph(common, paragraph);
        if paragraph.runs().is_empty() {
            each(&base);
            continue;
        }
        for run in paragraph.runs() {
            let mut resolved = styles.resolve_run(common, run);
            resolved.inherit(&base);
            each(&resolved);
        }
    }
    shared
}

/// A character property (record 0308) a node's style chain resolves to.
fn character<'a>(
    properties: &'a litchi_odp::scene::TextProperties,
    key: &str,
) -> Option<Value<'a>> {
    match key {
        "bold" => properties.bold().map(Value::Bool),
        "italic" => properties.italic().map(Value::Bool),
        "underline" => properties.underline().map(Value::Bool),
        "size" => {
            let raw = properties.size()?;
            Length::parse_in(raw, crate::value::LengthContext::FontSize)
                .or_else(|_| Length::parse(raw))
                .ok()
                .map(Value::Length)
        },
        "color" => {
            let hex = properties.color()?.strip_prefix('#')?;
            if hex.len() != 6 {
                return None;
            }
            let channel = |offset: usize| u8::from_str_radix(hex.get(offset..offset + 2)?, 16).ok();
            Some(Value::Color(crate::value::Color::rgb(
                channel(0)?,
                channel(2)?,
                channel(4)?,
            )))
        },
        "font" => properties
            .font()
            .map(|font| Value::Text(Cow::Borrowed(font))),
        _ => None,
    }
}

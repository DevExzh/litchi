//! Path resolution and query inspection for OpenDocument presentations.

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::selector::Position;
use litchi_odp::scene::{Object, ObjectRef, Scope};

use super::Odp;
use super::children::Children;
use super::node::{Item, Node, is_placeholder, notes_story, object_kind, page_comments};
use super::spec::{Kind, SPEC};
use crate::Result;
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::vocabulary::{Case, Identity};
use crate::query::inspect::Inspect;

/// The ODP adapter over one opened handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Odp,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Odp) -> Self {
        Self { file }
    }

    fn select(count: usize, select: Select<'_>) -> Option<Option<usize>> {
        Some(match select {
            Select::None => (count > 0).then_some(0),
            Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
            Select::Last => count.checked_sub(1),
            _ => return None,
        })
    }

    fn indexed(
        parent: &Node<'f>,
        count: usize,
        select: Select<'_>,
        kind: Kind,
        item: impl FnOnce(usize) -> Item,
    ) -> Lookup<Node<'f>, GapSpec> {
        match Self::select(count, select) {
            None => Lookup::NotApplicable,
            Some(None) => Lookup::Missing,
            Some(Some(index)) => Lookup::Found(parent.child(item(index), kind.name(), Some(index))),
        }
    }

    fn child_unfenced(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let kind = step.kind;
        if let Some(gap) = kind.spec().gap {
            return Ok(Lookup::Gap(gap));
        }
        let scene = parent.scene();
        Ok(match (parent.item(), kind) {
            (Item::Presentation, Kind::Slide) => {
                Self::indexed(parent, scene.pages().len(), step.select, kind, Item::Slide)
            },
            (Item::Presentation, Kind::SlideMaster) => Self::indexed(
                parent,
                self.file.masters()?.len(),
                step.select,
                kind,
                Item::Master,
            ),
            (Item::Presentation, Kind::SlideLayout) => Self::indexed(
                parent,
                self.file.layouts()?.len(),
                step.select,
                kind,
                Item::Layout,
            ),
            (Item::Slide(page), Kind::Notes) => {
                match step.select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                match notes_story(scene, *page) {
                    Some(_) => Lookup::Found(parent.child(Item::Notes(*page), "notes", None)),
                    None => Lookup::Missing,
                }
            },
            (Item::Slide(page), Kind::Comment) => {
                let comments: Vec<usize> = page_comments(self.file, *page)?.collect();
                match Self::select(comments.len(), step.select) {
                    None => Lookup::NotApplicable,
                    Some(None) => Lookup::Missing,
                    Some(Some(ordinal)) => Lookup::Found(parent.child(
                        Item::Comment(comments[ordinal]),
                        kind.name(),
                        Some(ordinal),
                    )),
                }
            },
            (Item::Slide(_) | Item::Object(_, Kind::Group), tree)
                if matches!(
                    tree,
                    Kind::Shape
                        | Kind::Placeholder
                        | Kind::Picture
                        | Kind::Table
                        | Kind::Chart
                        | Kind::Group
                        | Kind::Connector
                        | Kind::Ole
                        | Kind::Media
                        | Kind::Video
                        | Kind::Audio
                        | Kind::Equation
                ) =>
            {
                self.tree(parent, tree, step.select)?
            },
            (Item::Object(object, Kind::Shape), Kind::Hyperlink) => {
                match step.select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                if super::parts::shape_has_link(self.file, scene, object)? {
                    Lookup::Found(parent.child(Item::ShapeLink(object.clone()), "hyperlink", None))
                } else {
                    Lookup::Missing
                }
            },
            (Item::Run(story, paragraph, run), Kind::Hyperlink) => {
                match step.select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                let link = scene
                    .story(story)
                    .and_then(|story| story.paragraphs().get(*paragraph))
                    .and_then(|paragraph| paragraph.runs().get(*run))
                    .is_some_and(|run| run.kind() == litchi_odp::scene::RunKind::Link);
                if link {
                    Lookup::Found(parent.child(
                        Item::RunLink(story.clone(), *paragraph, *run),
                        "hyperlink",
                        None,
                    ))
                } else {
                    Lookup::Missing
                }
            },
            (Item::Paragraph(story, paragraph), Kind::LineBreak) => {
                let count = super::parts::line_break_count(self.file, scene, story, *paragraph)?;
                let story = story.clone();
                let paragraph = *paragraph;
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::LineBreak(story, paragraph, index)
                })
            },
            (Item::Object(object, Kind::Table), Kind::Col) => {
                let count = super::parts::column_count(self.file, scene, object)?;
                let object = object.clone();
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Column(object, index)
                })
            },
            (Item::Object(object, Kind::Chart), Kind::Series) => {
                let count = self
                    .file
                    .parts()?
                    .chart(object)
                    .map_or(0, |chart| chart.series.len());
                let object = object.clone();
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Series(object, index)
                })
            },
            (Item::Object(object, Kind::Chart), Kind::Axis) => {
                let parts = self.file.parts()?;
                let axes = parts.chart(object).map_or(&[][..], |chart| &chart.axes[..]);
                if let Select::Identity(identities) = step.select {
                    let mut found = None;
                    let mut matches = 0;
                    for (index, axis) in axes.iter().enumerate() {
                        let all = identities.iter().all(|(identity, value)| {
                            identity.key.eq_ignore_ascii_case("role")
                                && (axis
                                    .role()
                                    .is_some_and(|role| identity.case.matches(role, value))
                                    || axis.dimension.is_some_and(|dimension| {
                                        identity.case.matches(dimension, value)
                                    }))
                        });
                        if all {
                            matches += 1;
                            found = Some(index);
                        }
                    }
                    if matches > 1 {
                        return Ok(Lookup::Ambiguous(matches));
                    }
                    return Ok(match found {
                        Some(index) => Lookup::Found(parent.child(
                            Item::Axis(object.clone(), index),
                            kind.name(),
                            Some(index),
                        )),
                        None => Lookup::Missing,
                    });
                }
                let object = object.clone();
                Self::indexed(parent, axes.len(), step.select, kind, |index| {
                    Item::Axis(object, index)
                })
            },
            (Item::Object(_, Kind::Shape) | Item::Notes(_) | Item::Cell(..), Kind::Paragraph) => {
                let Some(story) = parent.own_story() else {
                    return Ok(Lookup::NotApplicable);
                };
                let count = scene
                    .story(&story)
                    .map_or(0, |story| story.paragraphs().len());
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Paragraph(story, index)
                })
            },
            (Item::Paragraph(story, paragraph), Kind::Run) => {
                let count = scene
                    .story(story)
                    .and_then(|story| story.paragraphs().get(*paragraph))
                    .map_or(0, |paragraph| paragraph.runs().len());
                let story = story.clone();
                let paragraph = *paragraph;
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Run(story, paragraph, index)
                })
            },
            (Item::Object(object, Kind::Table), Kind::Row) => {
                let count = scene
                    .object(object)
                    .and_then(Object::table)
                    .map_or(0, |table| table.rows().len());
                let object = object.clone();
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Row(object, index)
                })
            },
            (Item::Row(object, row), Kind::Cell) => {
                let count = scene
                    .object(object)
                    .and_then(Object::table)
                    .and_then(|table| table.rows().get(*row))
                    .map_or(0, |row| row.cells().len());
                let object = object.clone();
                let row = *row;
                Self::indexed(parent, count, step.select, kind, |index| {
                    Item::Cell(object, row, index)
                })
            },
            _ => Lookup::NotApplicable,
        })
    }

    fn tree(
        &self,
        parent: &Node<'f>,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let scene = parent.scene();
        let list = match parent.item() {
            Item::Slide(page) => ObjectRef {
                page: *page,
                scope: Scope::Page,
                path: Vec::new(),
            },
            Item::Object(object, Kind::Group) => object.clone(),
            _ => return Ok(Lookup::NotApplicable),
        };
        let objects: &[Object] = if list.path.is_empty() {
            scene.page(list.page).map_or(&[][..], |page| page.objects())
        } else {
            scene.object(&list).map_or(&[][..], Object::children)
        };
        // Video, audio and formula objects are addressed among the media
        // and embedded objects they are (record 0328).
        let semantic: Vec<Option<Kind>> =
            if matches!(kind, Kind::Video | Kind::Audio | Kind::Equation) {
                let mut kinds = Vec::with_capacity(objects.len());
                for (index, object) in objects.iter().enumerate() {
                    let mut path = list.path.clone();
                    path.push(index);
                    let reference = ObjectRef {
                        page: list.page,
                        scope: list.scope,
                        path,
                    };
                    kinds.push(super::parts::semantic_kind(self.file, &reference, object)?);
                }
                kinds
            } else {
                Vec::new()
            };
        let mut counters = [0_usize; Kind::COUNT];
        let candidates = objects.iter().enumerate().filter_map(|(index, object)| {
            let physical = object_kind(object);
            let found = if semantic.is_empty() {
                physical
            } else {
                semantic[index].unwrap_or(physical)
            };
            let ordinal = counters[found as usize];
            counters[found as usize] += 1;
            let matches = if kind == Kind::Placeholder {
                is_placeholder(object)
            } else {
                found == kind
            };
            matches.then_some((index, found, ordinal))
        });
        let mut candidates = candidates;
        let selected = match select {
            Select::None => candidates.next(),
            Select::Ordinal(position) => candidates.nth(position.get()),
            Select::Last => candidates.last(),
            Select::Identity(identities) => {
                let mut selected = None;
                let mut count = 0_usize;
                for candidate in candidates {
                    if identities.iter().all(|(identity, value)| {
                        identity_matches(&objects[candidate.0], identity, value)
                    }) {
                        count += 1;
                        selected = Some(candidate);
                    }
                }
                if count > 1 {
                    return Ok(Lookup::Ambiguous(count));
                }
                selected
            },
            Select::Name { name, case } => {
                candidates.find(|candidate| name_matches(objects[candidate.0].name(), name, case))
            },
            Select::Token(token) => candidates.find(|candidate| {
                objects[candidate.0]
                    .class()
                    .is_some_and(|class| class_matches(class, token))
            }),
            Select::Number(_) | Select::Cell(_) | Select::Range(_) => {
                return Ok(Lookup::NotApplicable);
            },
        };
        Ok(match selected {
            Some((index, physical, ordinal)) => {
                let mut path = list.path.clone();
                path.push(index);
                let reference = ObjectRef {
                    page: list.page,
                    scope: list.scope,
                    path,
                };
                Lookup::Found(parent.child(
                    Item::Object(reference, physical),
                    physical.name(),
                    Some(ordinal),
                ))
            },
            None => Lookup::Missing,
        })
    }

    fn available_unfenced(&self, parent: &Node<'f>, out: &mut Available) -> Result<()> {
        let scene = parent.scene();
        match parent.item() {
            Item::Presentation => {
                out.kind(Kind::Slide.name(), scene.pages().len());
                out.kind(Kind::SlideMaster.name(), self.file.masters()?.len());
                out.kind(Kind::SlideLayout.name(), self.file.layouts()?.len());
            },
            Item::Slide(_) | Item::Object(_, Kind::Group) => {
                let objects: &[Object] = match parent.item() {
                    Item::Slide(page) => scene.page(*page).map_or(&[][..], |page| page.objects()),
                    _ => parent.object().map_or(&[][..], Object::children),
                };
                let mut counts = [0_usize; Kind::COUNT];
                let mut placeholders = 0;
                for object in objects {
                    counts[object_kind(object) as usize] += 1;
                    placeholders += usize::from(is_placeholder(object));
                }
                for kind in [
                    Kind::Shape,
                    Kind::Picture,
                    Kind::Table,
                    Kind::Chart,
                    Kind::Group,
                    Kind::Connector,
                    Kind::Ole,
                    Kind::Media,
                ] {
                    out.kind(kind.name(), counts[kind as usize]);
                }
                out.kind(Kind::Placeholder.name(), placeholders);
                if let Item::Slide(page) = parent.item() {
                    out.kind("notes", usize::from(notes_story(scene, *page).is_some()));
                    out.kind(
                        Kind::Comment.name(),
                        page_comments(self.file, *page)?.count(),
                    );
                }
            },
            Item::Object(..) | Item::Notes(_) | Item::Cell(..) => {
                let count = parent
                    .own_story()
                    .and_then(|story| scene.story(&story).map(|story| story.paragraphs().len()))
                    .unwrap_or(0);
                out.kind(Kind::Paragraph.name(), count);
            },
            Item::Paragraph(story, index) => {
                let count = scene
                    .story(story)
                    .and_then(|story| story.paragraphs().get(*index))
                    .map_or(0, |paragraph| paragraph.runs().len());
                out.kind(Kind::Run.name(), count);
            },
            Item::Row(object, row) => {
                let count = scene
                    .object(object)
                    .and_then(Object::table)
                    .and_then(|table| table.rows().get(*row))
                    .map_or(0, |row| row.cells().len());
                out.kind(Kind::Cell.name(), count);
            },
            Item::Run(..)
            | Item::Comment(_)
            | Item::Master(_)
            | Item::Layout(_)
            | Item::ShapeLink(_)
            | Item::RunLink(..)
            | Item::LineBreak(..)
            | Item::Column(..)
            | Item::Series(..)
            | Item::Axis(..) => {},
        }
        Ok(())
    }
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Odp;
    const ABSENT: &'static [&'static str] = &[
        "smartart",
        "zoom",
        "modernComment",
        "reply",
        "model3d",
        "theme",
    ];

    fn check_source(&self) -> Result<()> {
        self.file.check_source()
    }

    fn root(&self) -> Result<Node<'f>> {
        let scene = self.file.scene()?;
        Ok(Node::root(self.file, scene))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, Self::Gap>> {
        self.child_unfenced(parent, step)
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self::Children> {
        Children::of(parent, kinds)
    }

    fn available(
        &self,
        parent: &Node<'f>,
        _step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        self.available_unfenced(parent, out)
    }
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Slide];
    const PLACEHOLDER_SHAPES: bool = true;

    fn kind(&self, node: &Node<'f>) -> Kind {
        // As in PPTX, a placeholder is a shape to selectors: `shape` matches
        // it, and a bare `placeholder` selects the shapes with a
        // presentation class (record 0290 F1).
        node.element()
    }

    fn is_placeholder(&self, node: &Node<'f>) -> bool {
        node.semantic_kind() == Kind::Placeholder
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        match node.item() {
            Item::Slide(page) => {
                let Some(slide) = node.scene().page(*page) else {
                    return Ok(None);
                };
                let mut text = String::new();
                for object in slide.objects() {
                    let object_text = object.text();
                    if !object_text.is_empty() {
                        if !text.is_empty() {
                            text.push('\n');
                        }
                        text.push_str(&object_text);
                    }
                }
                Ok(Some(Cow::Owned(text)))
            },
            _ => node.text_member(),
        }
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        node.format_member()
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        _kind: Kind,
        _ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        match node.steps().last() {
            Some(step) => fmt::Display::fmt(step, out),
            None => fmt::Display::fmt(&PathStep::new("presentation"), out),
        }
    }
}

fn identity_matches(object: &Object, identity: Identity, wanted: &str) -> bool {
    match identity.key {
        "id" => object.id().is_some_and(|id| id == wanted),
        "name" => name_matches(object.name(), wanted, identity.case),
        "type" => object
            .class()
            .is_some_and(|class| class_matches(class, wanted)),
        _ => false,
    }
}

fn name_matches(held: Option<&str>, wanted: &str, case: Case) -> bool {
    held.is_some_and(|held| case.matches(held, wanted))
}

/// Whether a `presentation:class` matches a placeholder type token, in the
/// ODF spelling or the PresentationML one.
fn class_matches(class: &str, wanted: &str) -> bool {
    if class.eq_ignore_ascii_case(wanted) {
        return true;
    }
    let wanted = wanted.to_ascii_lowercase();
    match wanted.as_str() {
        "ctrtitle" | "centertitle" => class == "title",
        "body" | "obj" | "object" | "content" => matches!(class, "outline" | "object" | "text"),
        "sub" => class == "subtitle",
        "pic" | "picture" => class == "graphic",
        "tbl" => class == "table",
        "dt" | "date" => class == "date-time",
        "ftr" => class == "footer",
        "hdr" => class == "header",
        "sldnum" | "slidenum" => class == "page-number",
        _ => false,
    }
}

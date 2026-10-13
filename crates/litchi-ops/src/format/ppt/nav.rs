//! Path resolution and query inspection for legacy PowerPoint.
//!
//! Resolution reads the catalog for slides and one native shape array for a
//! selected slide or group.  Unsupported presentation concepts are admitted
//! by the shared vocabulary and return their typed [`GapSpec`] instead of
//! being mistaken for unknown path names.

use litchi_core::FileFormat;
use litchi_core::selector::Position;
use litchi_ppt::shapes::shape::Shape as _;
use litchi_ppt::shapes::{PlaceholderType, ShapeEnum};
use std::borrow::Cow;
use std::fmt;

use super::Ppt;
use super::children::{Children, physical_kind, source_for};
use super::node::{Item, Node, ShapeItem, has_text_body, paragraph_count, shape_name};
use super::spec::{Kind, SPEC};
use crate::Result;
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::path::vocabulary::{Case, Identity};
use crate::query::inspect::Inspect;

/// The legacy PPT adapter over one opened handle.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Nav<'f> {
    file: &'f Ppt,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(file: &'f Ppt) -> Self {
        Self { file }
    }

    fn finish<T>(&self, operation: impl FnOnce() -> Result<T>) -> Result<T> {
        self.file.check_source()?;
        let result = operation();
        self.file.check_source().and(result)
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

        match (parent.item(), kind) {
            (Item::Presentation, Kind::Slide) => self.slide(step.select),
            (Item::Slide(slide), Kind::Notes) => self.notes(slide.position, step.select),
            (Item::Shape(shape), Kind::Paragraph) => self.paragraph(parent, shape, step.select),
            (Item::Slide(_) | Item::Shape(_), Kind::Shape | Kind::Placeholder)
            | (
                Item::Slide(_) | Item::Shape(_),
                Kind::Picture | Kind::Table | Kind::Group | Kind::Connector,
            ) => self.tree(parent, kind, step.select),
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn slide(&self, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        let count = self.file.reader().slide_count();
        let position = match select {
            Select::None => (count > 0).then_some(0),
            Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
            Select::Last => count.checked_sub(1),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match position {
            Some(position) => Lookup::Found(Node::slide_node(self.file, position)),
            None => Lookup::Missing,
        })
    }

    fn notes(&self, slide: usize, select: Select<'_>) -> Result<Lookup<Node<'f>, GapSpec>> {
        match select {
            Select::None | Select::Last => {},
            Select::Ordinal(position) if position.get() == 0 => {},
            Select::Ordinal(_) => return Ok(Lookup::Missing),
            _ => return Ok(Lookup::NotApplicable),
        }
        Ok(match slide_notes(self.file, slide)? {
            Some(text) => Lookup::Found(Node::notes_node(self.file, slide, text)),
            None => Lookup::Missing,
        })
    }

    fn paragraph(
        &self,
        parent: &Node<'f>,
        shape: &ShapeItem,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(native) = parent.shape() else {
            return Ok(Lookup::Missing);
        };
        if !has_text_body(native) {
            return Ok(Lookup::NotApplicable);
        }
        let count = paragraph_count(native)?;
        let index = match select {
            Select::None => (count > 0).then_some(0),
            Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
            Select::Last => count.checked_sub(1),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(match index {
            Some(index) => Lookup::Found(Node::paragraph_node(parent, shape.clone(), index)),
            None => Lookup::Missing,
        })
    }

    fn tree(
        &self,
        parent: &Node<'f>,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let Some(source) = source_for(parent)? else {
            return Ok(Lookup::NotApplicable);
        };
        let mut counters = [0_usize; Kind::COUNT];
        let Some(direct) = source.direct() else {
            return Ok(Lookup::NotApplicable);
        };
        // Ordinal/first lookups stop at the selected native shape. Keep only
        // physical-kind counters rather than allocating a candidate catalog.
        let mut candidates = direct.iter().enumerate().filter_map(|(index, shape)| {
            let physical = physical_kind(shape)?;
            let ordinal = counters[physical as usize];
            counters[physical as usize] = ordinal.saturating_add(1);
            let matches_kind = if kind == Kind::Placeholder {
                matches!(shape, ShapeEnum::Placeholder(_))
            } else {
                physical == kind
            };
            matches_kind.then_some(Candidate {
                index,
                kind: physical,
                ordinal,
            })
        });

        let candidate = match select {
            Select::None => candidates.next(),
            Select::Ordinal(position) => candidates.nth(position.get()),
            Select::Last => candidates.last(),
            Select::Identity(identities) => {
                let mut selected = None;
                let mut count = 0_usize;
                for candidate in candidates {
                    if identities.iter().all(|(identity, value)| {
                        identity_matches(&direct[candidate.index], identity, value)
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
            Select::Name { name, case } => candidates
                .find(|candidate| name_matches(shape_name(&direct[candidate.index]), name, case)),
            Select::Token(token) => {
                candidates.find(|candidate| placeholder_matches(&direct[candidate.index], token))
            },
            Select::Number(_) | Select::Cell(_) | Select::Range(_) => {
                return Ok(Lookup::NotApplicable);
            },
        };

        Ok(match candidate {
            Some(candidate) => {
                let Some(roots) = source.roots() else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(path) = source.child_path(candidate.index) else {
                    return Ok(Lookup::NotApplicable);
                };
                Lookup::Found(Node::shape_node(
                    self.file,
                    parent.steps(),
                    parent.slide_position().unwrap_or(0),
                    candidate.kind,
                    candidate.ordinal,
                    roots,
                    path,
                ))
            },
            None => Lookup::Missing,
        })
    }

    fn available_unfenced(&self, parent: &Node<'f>, out: &mut Available) -> Result<()> {
        match parent.item() {
            Item::Presentation => {
                out.kind(Kind::Slide.name(), self.file.reader().slide_count());
            },
            Item::Notes(_) | Item::Paragraph(_) => {},
            Item::Shape(_) if parent.shape().is_some_and(has_text_body) => {
                let count = parent.shape().map_or(Ok(0), paragraph_count)?;
                out.kind(Kind::Paragraph.name(), count);
            },
            Item::Slide(_) | Item::Shape(_) => {
                if let Item::Slide(slide) = parent.item() {
                    let notes = usize::from(slide_notes(self.file, slide.position)?.is_some());
                    out.kind(Kind::Notes.name(), notes);
                }
                if let Some(source) = source_for(parent)? {
                    let Some(direct) = source.direct() else {
                        return Ok(());
                    };
                    let mut counts = [0_usize; Kind::COUNT];
                    let mut placeholders = 0_usize;
                    for shape in direct {
                        if let Some(kind) = physical_kind(shape) {
                            counts[kind as usize] += 1;
                            if matches!(shape, ShapeEnum::Placeholder(_)) {
                                placeholders += 1;
                            }
                        }
                    }
                    for kind in [
                        Kind::Shape,
                        Kind::Placeholder,
                        Kind::Picture,
                        Kind::Table,
                        Kind::Group,
                        Kind::Connector,
                    ] {
                        let count = if kind == Kind::Placeholder {
                            placeholders
                        } else {
                            counts[kind as usize]
                        };
                        out.kind(kind.name(), count);
                    }
                }
            },
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
struct Candidate {
    index: usize,
    kind: Kind,
    ordinal: usize,
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Ppt;

    fn root(&self) -> Result<Node<'f>> {
        self.finish(|| Ok(Node::root(self.file)))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, Kind>,
    ) -> Result<Lookup<Node<'f>, Self::Gap>> {
        self.finish(|| self.child_unfenced(parent, step))
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self::Children> {
        self.finish(|| Ok(Children::walk(parent, kinds)))
    }

    fn available(
        &self,
        parent: &Node<'f>,
        _step: &Resolved<'_, Kind>,
        out: &mut Available,
    ) -> Result<()> {
        self.finish(|| self.available_unfenced(parent, out))
    }
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Slide];

    fn kind(&self, node: &Node<'f>) -> Kind {
        // A legacy placeholder is stored in the native scene as a physical
        // shape.  Queries use the semantic placeholder kind, while the node
        // itself keeps its physical `/shape[N]` address and ordinal.
        node.semantic_kind()
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        self.finish(|| match node.item() {
            Item::Slide(slide) => {
                let text = self
                    .file
                    .reader()
                    .slide_at(slide.position)?
                    .ok_or_else(|| missing_slide(slide.position))?
                    .text()?
                    .to_owned();
                Ok(Some(Cow::Owned(text)))
            },
            _ => node.text_member(),
        })
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        self.finish(|| node.format_member())
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        _kind: Kind,
        _ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        match (node.item(), node.steps().last()) {
            (Item::Notes(_), _) => fmt::Display::fmt(&PathStep::new("notes"), out),
            (_, Some(step)) => fmt::Display::fmt(step, out),
            (_, None) => fmt::Display::fmt(&PathStep::new("presentation"), out),
        }
    }
}

fn identity_matches(shape: &ShapeEnum<'static>, identity: Identity, wanted: &str) -> bool {
    match identity.key {
        "id" => shape_id(shape).is_some_and(|id| id_is(id, wanted)),
        "name" => name_matches(shape_name(shape), wanted, identity.case),
        "type" => placeholder_matches(shape, wanted),
        "idx" => match shape {
            ShapeEnum::Placeholder(placeholder) => placeholder
                .index()
                .is_some_and(|index| id_is(u32::from(index), wanted)),
            _ => false,
        },
        _ => false,
    }
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

/// A slide's speaker-notes text, or `None` when the slide has no notes page.
pub(crate) fn slide_notes(file: &Ppt, slide: usize) -> Result<Option<std::sync::Arc<str>>> {
    let selected = file
        .reader()
        .slide_at(slide)?
        .ok_or_else(|| missing_slide(slide))?;
    Ok(selected
        .speaker_notes()?
        .map(|notes| notes.text().map(std::sync::Arc::from))
        .transpose()?)
}

fn id_is(id: u32, wanted: &str) -> bool {
    wanted.parse::<u32>().ok() == Some(id)
}

fn name_matches(held: Option<&str>, wanted: &str, case: Case) -> bool {
    let Some(held) = held else {
        return false;
    };
    case.matches(held, wanted)
        || held
            .strip_prefix("!!")
            .is_some_and(|held| case.matches(held, wanted))
        || wanted
            .strip_prefix("!!")
            .is_some_and(|wanted| case.matches(held, wanted))
}

fn placeholder_matches(shape: &ShapeEnum<'_>, wanted: &str) -> bool {
    let ShapeEnum::Placeholder(placeholder) = shape else {
        return false;
    };
    let kind = placeholder.placeholder_type();
    let display = kind.to_string();
    if Case::Ascii.matches(&display, wanted) {
        return true;
    }
    let wanted = wanted.to_ascii_lowercase();
    match wanted.as_str() {
        "title" => matches!(kind, PlaceholderType::Title),
        "ctrtitle" | "centertitle" => matches!(kind, PlaceholderType::CenterTitle),
        "subtitle" | "sub" => matches!(kind, PlaceholderType::SubTitle),
        "body" => matches!(kind, PlaceholderType::Body),
        "content" | "obj" | "object" => {
            matches!(kind, PlaceholderType::Content | PlaceholderType::Object)
        },
        "pic" | "picture" | "slideimage" | "sldimg" => {
            matches!(kind, PlaceholderType::Picture | PlaceholderType::SlideImage)
        },
        "tbl" | "table" => matches!(kind, PlaceholderType::Table),
        "chart" => matches!(kind, PlaceholderType::Chart),
        "dgm" | "diagram" => matches!(kind, PlaceholderType::Diagram),
        "media" | "mediaclip" => matches!(kind, PlaceholderType::MediaClip),
        _ => false,
    }
}

fn missing_slide(position: usize) -> crate::Error {
    crate::Error::new(
        crate::Code::InternalError,
        format!("slide {} vanished from the catalog", position + 1),
    )
    .with_format(FileFormat::Ppt)
}

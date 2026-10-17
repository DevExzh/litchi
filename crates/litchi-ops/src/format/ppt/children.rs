//! Lazy direct-child traversal for legacy PowerPoint nodes.
//!
//! Presentation children are catalog entries.  Slide and group children are
//! native shape records loaded only when the iterator is first pulled.  Each
//! subsequent pull checks the retained source, so an iterator cannot quietly
//! continue over a source that changed after a node was selected.

use std::sync::Arc;

use litchi_ppt::shapes::{ShapeEnum, ShapeType};

use super::Ppt;
use super::node::{Item, Node, ShapeItem, paragraph_count};
use super::spec::Kind;
use crate::error::{Error, Result};
use crate::path::navigate::KindSet;

/// The children of one legacy PPT node.
pub(crate) struct Children<'f> {
    file: &'f Ppt,
    kinds: KindSet<Kind>,
    state: State,
    /// A query walk also visits a slide's notes after its shape tree; a view
    /// listing does not, as in the PPTX adapter.
    notes: Option<u32>,
    /// Record-0323 elements a query walk lists after the main axis: a
    /// slide's comments, a shape's hyperlink.
    extra: Option<Box<Extra<'f>>>,
}

/// The record-0323 elements of one query-walk step, or the error building
/// them, reported on the first pull.
struct Extra<'f> {
    nodes: std::vec::IntoIter<Node<'f>>,
    failed: Option<Error>,
}

// Only physical kinds through Connector are emitted by this iterator.
// Gap-only vocabulary kinds need no counters in every traversal frame.
const COUNTER_COUNT: usize = Kind::Connector as usize + 1;

enum State {
    Done,
    Slides {
        next: usize,
        end: usize,
    },
    Shapes {
        slide: usize,
        parent: Arc<[crate::node::PathStep<'static>]>,
        source: ShapeSource,
        next: usize,
        counters: [usize; COUNTER_COUNT],
    },
    /// A query walk's rows of one table (`row: None`) or cells of one row.
    Bands(Box<Bands>),
    /// The paragraphs of one text shape.
    Paragraphs {
        parent: Arc<[crate::node::PathStep<'static>]>,
        shape: ShapeItem,
        next: usize,
        end: usize,
    },
}

/// The rows of one table or the cells of one row, listed by a query walk.
struct Bands {
    parent: Arc<[crate::node::PathStep<'static>]>,
    shape: ShapeItem,
    row: Option<usize>,
    next: usize,
    end: usize,
}

pub(crate) enum ShapeSource {
    /// The slide is parsed only when the first child is requested.
    Slide(usize),
    /// A parsed slide/group root array and the index trail to a group whose
    /// direct children are being listed.  The trail points into `roots`, so a
    /// nested group never copies its child payloads.
    Tree {
        roots: Arc<[ShapeEnum<'static>]>,
        path: Arc<[usize]>,
    },
}

impl ShapeSource {
    pub(crate) fn tree(roots: Arc<[ShapeEnum<'static>]>, path: Arc<[usize]>) -> Self {
        Self::Tree { roots, path }
    }

    pub(crate) fn direct(&self) -> Option<&[ShapeEnum<'static>]> {
        match self {
            Self::Slide(_) => None,
            Self::Tree { roots, path } if path.is_empty() => Some(roots),
            Self::Tree { roots, path } => resolve_shape(roots, path)
                .and_then(|shape| shape.as_group())
                .map(|group| group.children()),
        }
    }

    pub(crate) fn roots(&self) -> Option<Arc<[ShapeEnum<'static>]>> {
        match self {
            Self::Slide(_) => None,
            Self::Tree { roots, .. } => Some(Arc::clone(roots)),
        }
    }

    pub(crate) fn child_path(&self, index: usize) -> Option<Arc<[usize]>> {
        let path = match self {
            Self::Slide(_) => return None,
            Self::Tree { path, .. } => path,
        };
        let mut child = path.to_vec();
        child.push(index);
        Some(Arc::from(child.into_boxed_slice()))
    }
}

/// Build an eager source for navigation or availability probes.  View
/// traversal uses the `Slide` variant above to defer this read until `next`.
pub(crate) fn source_for(parent: &Node<'_>) -> Result<Option<ShapeSource>> {
    match parent.item() {
        Item::Slide(slide) => {
            let Some(slide) = parent.file().reader().slide_at(slide.position)? else {
                return Err(missing_slide(slide.position));
            };
            Ok(Some(ShapeSource::tree(
                slide.shared_shapes()?,
                Arc::from(Vec::<usize>::new().into_boxed_slice()),
            )))
        },
        Item::Shape(shape) if shape.kind == Kind::Group => {
            let source = ShapeSource::tree(Arc::clone(&shape.source), Arc::clone(&shape.path));
            Ok(source.direct().is_some().then_some(source))
        },
        _ => Ok(None),
    }
}

impl<'f> Children<'f> {
    /// Construct the lazy child iterator a query walk uses: a view listing's
    /// children, then a slide's notes.
    pub(crate) fn walk(parent: &Node<'f>, kinds: KindSet<Kind>) -> Self {
        let mut children = Self::of(parent, kinds);
        // Record-0323 elements (comments, rows and cells, hyperlinks, runs)
        // are walked by a query that names them, never by a view or a `*`
        // walk, so existing views and query results keep their shape, as
        // the PPTX walk keeps its record-0321 elements.
        if Kind::ALL.iter().all(|kind| kinds.contains(*kind)) {
            if let Item::Slide(slide) = parent.item()
                && kinds.contains(Kind::Notes)
            {
                children.notes = u32::try_from(slide.position).ok();
            }
            return children;
        }
        match parent.item() {
            Item::Shape(shape) if shape.kind == Kind::Table && kinds.contains(Kind::Row) => {
                if let Ok(table) = super::elements::table(parent) {
                    children.state = State::Bands(Box::new(Bands {
                        parent: Arc::from(parent.steps().to_vec().into_boxed_slice()),
                        shape: shape.clone(),
                        row: None,
                        next: 0,
                        end: table.rows(),
                    }));
                }
            },
            Item::Element(element) if kinds.contains(Kind::Cell) => {
                if let super::elements::Element::Row { shape, row } = element.as_ref()
                    && let Ok(table) = super::elements::table(parent)
                {
                    children.state = State::Bands(Box::new(Bands {
                        parent: Arc::from(parent.steps().to_vec().into_boxed_slice()),
                        shape: shape.clone(),
                        row: Some(*row),
                        next: 0,
                        end: table.columns(),
                    }));
                }
            },
            _ => {},
        }
        if let Item::Slide(slide) = parent.item() {
            if kinds.contains(Kind::Notes) {
                children.notes = u32::try_from(slide.position).ok();
            }
        }
        let mut extra = Vec::new();
        let mut failed = None;
        if let Item::Slide(slide) = parent.item()
            && kinds.contains(Kind::Comment)
        {
            match super::elements::slide_comments(parent.file(), slide.position) {
                Ok(comments) => {
                    for (index, comment) in comments.into_iter().enumerate() {
                        extra.push(super::nav::comment_node(
                            parent.file(),
                            parent.steps(),
                            slide.position,
                            index,
                            comment,
                        ));
                    }
                },
                Err(error) => failed = Some(error),
            }
        }
        for kind in [Kind::Ole, Kind::Video, Kind::Audio, Kind::Media] {
            if !kinds.contains(kind)
                || (kind != Kind::Media && kinds.contains(Kind::Media))
                || !matches!(parent.item(), Item::Slide(_) | Item::Shape(_))
            {
                continue;
            }
            let found = (|| -> Result<()> {
                let Some(source) = source_for(parent)? else {
                    return Ok(());
                };
                let Some(direct) = source.direct() else {
                    return Ok(());
                };
                for (index, (position, program, linked)) in
                    super::elements::frames(parent.file(), direct, kind)?
                        .into_iter()
                        .enumerate()
                {
                    let (Some(roots), Some(path)) = (source.roots(), source.child_path(position))
                    else {
                        continue;
                    };
                    extra.push(Node::element_node(
                        parent.file(),
                        parent.steps(),
                        crate::node::PathStep::at(kind.name(), litchi_core::Position::new(index)),
                        super::elements::Element::Frame {
                            shape: ShapeItem {
                                slide: parent.slide_position().unwrap_or(0),
                                kind: Kind::Picture,
                                source: roots,
                                path,
                            },
                            kind,
                            program,
                            linked,
                        },
                    ));
                }
                Ok(())
            })();
            if let Err(error) = found {
                failed = Some(error);
            }
        }
        if let Item::Slide(slide) = parent.item()
            && kinds.contains(Kind::Chart)
        {
            match super::elements::slide_charts(parent.file(), slide.position) {
                Ok(charts) => {
                    for (index, chart) in charts.into_iter().enumerate() {
                        extra.push(super::nav::chart_node(parent, slide.position, index, chart));
                    }
                },
                Err(error) => failed = Some(error),
            }
        }
        if let Item::Shape(shape) = parent.item()
            && kinds.contains(Kind::Hyperlink)
        {
            match super::elements::shape_click(parent) {
                Ok(Some(click)) => {
                    extra.push(super::nav::hyperlink_node(parent, shape, None, click))
                },
                Ok(None) => {},
                Err(error) => failed = Some(error),
            }
        }
        if let Item::Presentation = parent.item()
            && kinds.contains(Kind::Theme)
        {
            match parent.file().reader().theme() {
                Ok(Some(theme)) => extra.push(Node::element_node(
                    parent.file(),
                    parent.steps(),
                    crate::node::PathStep::new(Kind::Theme.name()),
                    super::elements::Element::Theme(Arc::new(theme)),
                )),
                Ok(None) => {},
                Err(error) => failed = Some(error.into()),
            }
        }
        if let Item::Shape(shape) = parent.item()
            && shape.kind == Kind::Table
            && kinds.contains(Kind::Col)
        {
            match super::elements::table(parent) {
                Ok(table) => {
                    for column in 0..table.columns() {
                        extra.push(super::nav::band_node(
                            parent.file(),
                            parent.steps(),
                            shape,
                            Kind::Col,
                            column,
                        ));
                    }
                },
                Err(error) => failed = Some(error),
            }
        }
        if let Item::Paragraph(paragraph) = parent.item()
            && kinds.contains(Kind::LineBreak)
        {
            match parent.text_member() {
                Ok(text) => {
                    let count = super::elements::line_breaks(&text.unwrap_or_default(), 0);
                    for index in 0..count {
                        extra.push(Node::element_node(
                            parent.file(),
                            parent.steps(),
                            crate::node::PathStep::at(
                                Kind::LineBreak.name(),
                                litchi_core::Position::new(index),
                            ),
                            super::elements::Element::LineBreak {
                                shape: paragraph.shape.clone(),
                            },
                        ));
                    }
                },
                Err(error) => failed = Some(error),
            }
        }
        if let Item::Paragraph(paragraph) = parent.item()
            && (kinds.contains(Kind::Run) || kinds.contains(Kind::LineBreak))
        {
            if kinds.contains(Kind::Run) {
                match super::elements::shape_runs(parent) {
                    Ok(Some(body)) => {
                        let body = Arc::new(body);
                        for (index, piece) in super::elements::pieces(&body, paragraph.index)
                            .into_iter()
                            .enumerate()
                        {
                            extra.push(super::nav::run_node(
                                parent,
                                &paragraph.shape,
                                paragraph.index,
                                index,
                                Arc::clone(&body),
                                piece,
                            ));
                        }
                    },
                    Ok(None) => {},
                    Err(error) => failed = Some(error),
                }
            }
        }
        if let Item::Element(element) = parent.item()
            && kinds.contains(Kind::Hyperlink)
            && let super::elements::Element::Run { shape, .. } = element.as_ref()
        {
            match super::elements::run_click(parent) {
                Ok(Some((range, click))) => {
                    extra.push(super::nav::hyperlink_node(
                        parent,
                        shape,
                        Some(range),
                        click,
                    ));
                },
                Ok(None) => {},
                Err(error) => failed = Some(error),
            }
        }
        if !extra.is_empty() || failed.is_some() {
            children.extra = Some(Box::new(Extra {
                nodes: extra.into_iter(),
                failed,
            }));
        }
        children
    }

    /// Construct a lazy child iterator for `parent`.
    pub(crate) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Self {
        let file = parent.file();
        let state = match parent.item() {
            Item::Presentation if kinds.contains(Kind::Slide) => State::Slides {
                next: 0,
                end: file.reader().slide_count(),
            },
            Item::Presentation => State::Done,
            Item::Slide(slide) if wants_shapes(kinds) => State::Shapes {
                slide: slide.position,
                parent: Arc::from(parent.steps().to_vec().into_boxed_slice()),
                source: ShapeSource::Slide(slide.position),
                next: 0,
                counters: [0; COUNTER_COUNT],
            },
            Item::Slide(_) => State::Done,
            Item::Shape(shape) if shape.kind == Kind::Group => {
                if !wants_shapes(kinds) {
                    return Self {
                        file,
                        kinds,
                        state: State::Done,
                        notes: None,
                        extra: None,
                    };
                }
                let source = ShapeSource::tree(Arc::clone(&shape.source), Arc::clone(&shape.path));
                if source.direct().is_some() {
                    State::Shapes {
                        slide: shape.slide,
                        parent: Arc::from(parent.steps().to_vec().into_boxed_slice()),
                        source,
                        next: 0,
                        counters: [0; COUNTER_COUNT],
                    }
                } else {
                    State::Done
                }
            },
            Item::Shape(shape) if kinds.contains(Kind::Paragraph) => {
                match parent.shape().map(paragraph_count) {
                    Some(Ok(end)) if end > 0 => State::Paragraphs {
                        parent: Arc::from(parent.steps().to_vec().into_boxed_slice()),
                        shape: shape.clone(),
                        next: 0,
                        end,
                    },
                    _ => State::Done,
                }
            },
            Item::Shape(_) | Item::Notes(_) | Item::Paragraph(_) | Item::Element(_) => State::Done,
        };
        Self {
            file,
            kinds,
            state,
            notes: None,
            extra: None,
        }
    }

    fn step(&mut self) -> Result<Option<Node<'f>>> {
        // Once the iterator reaches EOF or reports an error, it is terminal.
        // In particular, do not perform another source check on a later pull:
        // a source that changes after EOF/error must not make this iterator
        // yield a second result.
        if let Some(error) = self.extra.as_mut().and_then(|extra| extra.failed.take()) {
            self.extra = None;
            return Err(error);
        }
        if matches!(self.state, State::Done) && self.notes.is_none() && self.extra.is_none() {
            return Ok(None);
        }
        self.file.check_source()?;
        let result = self.step_unfenced();
        self.file.check_source().and(result)
    }

    fn step_unfenced(&mut self) -> Result<Option<Node<'f>>> {
        let state = std::mem::replace(&mut self.state, State::Done);
        let next = self.step_state(state)?;
        if next.is_some() {
            return Ok(next);
        }
        // The shape tree is done: a query walk then visits the slide's notes.
        if let Some(slide) = self.notes.take().map(|slide| slide as usize)
            && let Some(text) = super::nav::slide_notes(self.file, slide)?
        {
            return Ok(Some(Node::notes_node(self.file, slide, text)));
        }
        // Then its record-0323 elements, one per pull.
        let Some(extra) = self.extra.as_mut() else {
            return Ok(None);
        };
        let next = extra.nodes.next();
        if next.is_none() {
            self.extra = None;
        }
        Ok(next)
    }

    fn step_state(&mut self, state: State) -> Result<Option<Node<'f>>> {
        match state {
            State::Done => Ok(None),
            State::Bands(mut bands) => {
                if bands.next >= bands.end {
                    return Ok(None);
                }
                let next = bands.next;
                let node = match bands.row {
                    None => super::nav::band_node(
                        self.file,
                        &bands.parent,
                        &bands.shape,
                        Kind::Row,
                        next,
                    ),
                    Some(row) => {
                        super::nav::cell_node(self.file, &bands.parent, &bands.shape, row, next)
                    },
                };
                bands.next += 1;
                self.state = State::Bands(bands);
                Ok(Some(node))
            },
            State::Paragraphs {
                parent,
                shape,
                next,
                end,
            } => {
                if next >= end {
                    return Ok(None);
                }
                let node = Node::paragraph_at(self.file, &parent, shape.clone(), next);
                self.state = State::Paragraphs {
                    parent,
                    shape,
                    next: next + 1,
                    end,
                };
                Ok(Some(node))
            },
            State::Slides { next, end } => {
                if next >= end {
                    return Ok(None);
                }
                self.state = State::Slides {
                    next: next + 1,
                    end,
                };
                Ok(Some(Node::slide_node(self.file, next)))
            },
            State::Shapes {
                slide,
                parent,
                source,
                mut next,
                mut counters,
            } => {
                let source = match source {
                    ShapeSource::Slide(position) => {
                        let slide_ref = self
                            .file
                            .reader()
                            .slide_at(position)?
                            .ok_or_else(|| missing_slide(position))?;
                        ShapeSource::tree(
                            slide_ref.shared_shapes()?,
                            Arc::from(Vec::<usize>::new().into_boxed_slice()),
                        )
                    },
                    source => source,
                };

                let Some(direct_len) = source.direct().map(|direct| direct.len()) else {
                    return Ok(None);
                };
                while next < direct_len {
                    let index = next;
                    next += 1;
                    let Some((kind, is_placeholder)) = (|| {
                        let direct = source.direct()?;
                        let shape = direct.get(index)?;
                        Some((
                            physical_kind(shape)?,
                            matches!(shape, ShapeEnum::Placeholder(_)),
                        ))
                    })() else {
                        continue;
                    };
                    let ordinal = counters[kind as usize];
                    counters[kind as usize] = ordinal.saturating_add(1);
                    let wanted = self.kinds.contains(kind)
                        || (is_placeholder && self.kinds.contains(Kind::Placeholder));
                    if !wanted {
                        continue;
                    }

                    let Some(roots) = source.roots() else {
                        return Ok(None);
                    };
                    let Some(path) = source.child_path(index) else {
                        return Ok(None);
                    };
                    let node =
                        Node::shape_node(self.file, &parent, slide, kind, ordinal, roots, path);
                    self.state = State::Shapes {
                        slide,
                        parent,
                        source,
                        next,
                        counters,
                    };
                    return Ok(Some(node));
                }
                Ok(None)
            },
        }
    }
}

fn wants_shapes(kinds: KindSet<Kind>) -> bool {
    [
        Kind::Shape,
        Kind::Placeholder,
        Kind::Picture,
        Kind::Table,
        Kind::Group,
        Kind::Connector,
    ]
    .into_iter()
    .any(|kind| kinds.contains(kind))
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.step() {
            Ok(Some(node)) => Some(Ok(node)),
            Ok(None) => None,
            Err(error) => {
                self.state = State::Done;
                self.notes = None;
                self.extra = None;
                Some(Err(error))
            },
        }
    }
}

impl<'f> std::iter::FusedIterator for Children<'f> {}

/// Classify one native shape into the physical navigation kind.
pub(crate) fn physical_kind(shape: &ShapeEnum<'_>) -> Option<Kind> {
    Some(match shape {
        ShapeEnum::TextBox(_) | ShapeEnum::Placeholder(_) | ShapeEnum::AutoShape(_) => Kind::Shape,
        ShapeEnum::Picture(_) => Kind::Picture,
        ShapeEnum::Table(_) => Kind::Table,
        ShapeEnum::Group(_) => Kind::Group,
        ShapeEnum::Line(line) => match line.shape_type() {
            ShapeType::Connector => Kind::Connector,
            _ => Kind::Shape,
        },
    })
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

fn missing_slide(position: usize) -> Error {
    Error::new(
        crate::Code::InternalError,
        format!("slide {} vanished from the catalog", position + 1),
    )
    .with_format(litchi_core::FileFormat::Ppt)
}

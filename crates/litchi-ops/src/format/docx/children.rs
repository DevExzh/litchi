//! The children of a DOCX node, produced lazily in document order (record
//! 0010 §4.5.1, §5.3).
//!
//! | Parent | Children |
//! |---|---|
//! | `/` | `body` (its blocks are read when it is entered), the styles part, the headers, the footers, the numbering part, the footnotes and endnotes parts (record 0319), the comments part, then the SmartArt graphics (record 0295) |
//! | a styles, numbering or notes part, a definition | its styles; its definitions, then its instances; its notes; its levels (record 0319) |
//! | `body`, a cell, a content control | paragraphs, tables, content controls, display equations and (in the body) section breaks, in visible order; `w:customXml` wrappers flattened; `w:altChunk` anchors and unknown children skipped |
//! | a table | rows |
//! | a row | cells |
//! | a paragraph | its addressable runs, then its direct hyperlinks, then its inline pictures, then its charts |
//! | a chart | its series, then its axes (record 0335) |
//! | a SmartArt graphic or node | its child nodes (record 0335) |
//!
//! The iterator owns what it reads (shared leaf values, an index cursor), so
//! it borrows neither the adapter nor the parent node, and a walk that
//! stops early reads no later sibling.

use std::sync::Arc;
use std::vec;

use litchi_docx::{Blocks, Cell, InlineHyperlink, Paragraph, Row, Run};

use super::Docx;
use super::nav::{self, Nav, block_node, container_of};
use super::node::{Address, Node, Step, Value};
use super::spec::Kind;
use crate::error::Result;
use crate::path::navigate::KindSet;

/// The children of one DOCX node whose kinds are in a set.
pub(crate) struct Children<'f> {
    docx: &'f Docx,
    /// The address the children extend.
    address: Address,
    kinds: KindSet<Kind>,
    source: Source,
}

enum Source {
    Done,
    /// The root's children: the body, then the part elements (record
    /// 0319), then the comments part.
    Root,
    /// The root's styles, header, footer, numbering and notes parts, or the
    /// children of one of them (record 0319).
    Parts {
        children: Box<super::parts::Children>,
        /// Whether the comments part follows (the root's listing).
        root: bool,
    },
    /// Marks of the parent, then the parent's other children (record 0319).
    Marks {
        children: Box<super::marks::Children>,
        then: Box<Source>,
    },
    /// The root's comments part, when the document has one.
    CommentsRoot,
    /// The root's SmartArt graphics, read when the first is needed
    /// (record 0295).
    SmartArtRoot,
    SmartArts {
        graphics: Arc<[Arc<super::smartart::SmartArtValue>]>,
        next: usize,
    },
    /// The comments of the comments part.
    Comments {
        comments: vec::IntoIter<litchi_docx::comment::Comment>,
        next: u32,
    },
    /// The direct blocks of a block container, from `next` on.
    Blocks {
        container: Blocks,
        next: usize,
    },
    Rows {
        rows: vec::IntoIter<Row>,
        next: u32,
    },
    Cells {
        cells: vec::IntoIter<Cell>,
        next: u32,
    },
    /// A paragraph's runs; its hyperlinks and pictures follow when `rest`
    /// holds it.
    Runs {
        runs: vec::IntoIter<Run>,
        next: u32,
        rest: Option<Rest>,
    },
    /// A paragraph's hyperlinks, read when the first is needed.
    Links(Rest),
    Hyperlinks {
        links: vec::IntoIter<InlineHyperlink>,
        next: u32,
        pictures: Option<Paragraph>,
        charts: Option<Box<Charted>>,
    },
    /// A paragraph's inline pictures, read when the first is needed; its
    /// charts follow when `charts` holds the paragraph.
    PictureList(Paragraph, Option<Box<Charted>>),
    Pictures {
        pictures: vec::IntoIter<litchi_docx::image::InlineImage>,
        next: u32,
        charts: Option<Box<Charted>>,
    },
    /// Children read in one go: a chart's series and axes, a graphic's
    /// nodes (record 0335).
    Built(vec::IntoIter<(Kind, Value, Step)>),
    /// A paragraph's charts, read when the first is needed.
    ChartList(Box<Charted>),
    Charts {
        charts: vec::IntoIter<litchi_docx::chart::InlineChart>,
        next: u32,
        paragraph: Box<Charted>,
    },
}

/// The paragraph whose charts a listing reads: its value and address.
pub(super) struct Charted {
    pub(super) value: Value,
    pub(super) address: Address,
}

impl Source {
    /// What follows a paragraph's hyperlinks: its pictures, then its charts,
    /// each when wanted.
    fn after_links(pictures: Option<Paragraph>, charts: Option<Box<Charted>>) -> Self {
        match (pictures, charts) {
            (Some(paragraph), charts) => Self::PictureList(paragraph, charts),
            (None, Some(charts)) => Self::ChartList(charts),
            (None, None) => Self::Done,
        }
    }
}

impl<'f> Children<'f> {
    /// The children of `parent` whose kinds are in `kinds`.
    pub(super) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let source = match &parent.value {
            Value::Document if kinds.contains(Kind::Body) => Source::Root,
            Value::Document if super::parts::Children::wants_root(kinds) => Source::Parts {
                children: Box::new(super::parts::Children::root(kinds)),
                root: true,
            },
            Value::Document if kinds.contains(Kind::Comments) => Source::CommentsRoot,
            Value::Document if kinds.contains(Kind::SmartArt) => Source::SmartArtRoot,
            Value::Part(value) => Source::Parts {
                children: Box::new(super::parts::Children::of(parent, *value, kinds)),
                root: false,
            },
            Value::Comments if kinds.contains(Kind::Comment) => Source::Comments {
                comments: nav::comments(parent.docx)?.into_iter(),
                next: 0,
            },
            Value::Body(_) | Value::Cell(_) | Value::Sdt(_) if wants_blocks(kinds) => {
                match container_of(parent)? {
                    Some(container) => Source::Blocks { container, next: 0 },
                    None => Source::Done,
                }
            },
            Value::Table(_) if kinds.contains(Kind::Row) => Source::Rows {
                rows: Nav::rows(parent)?.into_iter(),
                next: 0,
            },
            Value::Row(row) if kinds.contains(Kind::Cell) => Source::Cells {
                cells: row.cells()?.into_vec().into_iter(),
                next: 0,
            },
            Value::Chart(chart) if kinds.contains(Kind::Series) || kinds.contains(Kind::Axis) => {
                Source::Built(super::chart::chart_children(parent.docx, chart, kinds)?.into_iter())
            },
            Value::SmartArt(_) | Value::DiagramNode(_) if kinds.contains(Kind::Node) => {
                Source::Built(
                    super::smartart::node_children(parent.docx, parent)?
                        .into_iter()
                        .map(|(value, step)| (Kind::Node, value, step))
                        .collect::<Vec<_>>()
                        .into_iter(),
                )
            },
            Value::Paragraph { paragraph, .. } => {
                let rest = Rest {
                    paragraph: paragraph.clone(),
                    links: kinds.contains(Kind::Hyperlink),
                    pictures: kinds.contains(Kind::Picture),
                    charts: kinds.contains(Kind::Chart).then(|| {
                        Box::new(Charted {
                            value: parent.value.clone(),
                            address: parent.child_address(),
                        })
                    }),
                };
                let rest = (rest.links || rest.pictures || rest.charts.is_some()).then_some(rest);
                if kinds.contains(Kind::Run) {
                    Source::Runs {
                        runs: paragraph.addressable_runs()?.into_vec().into_iter(),
                        next: 0,
                        rest,
                    }
                } else {
                    rest.map_or(Source::Done, Source::Links)
                }
            },
            _ => Source::Done,
        };
        // Tab stops, breaks and the document-wide marks come first when a
        // query asks for them (record 0319).
        let source = match super::marks::Children::of(parent, kinds) {
            Some(marks) => Source::Marks {
                children: Box::new(marks),
                then: Box::new(source),
            },
            None => source,
        };
        Ok(Self {
            docx: parent.docx,
            address: parent.child_address(),
            kinds,
            source,
        })
    }

    fn produce(&mut self) -> Option<Result<Node<'f>>> {
        loop {
            match &mut self.source {
                Source::Done => return None,
                Source::Root => {
                    self.source = Source::Parts {
                        children: Box::new(super::parts::Children::root(self.kinds)),
                        root: true,
                    };
                    if !self.kinds.contains(Kind::Body) {
                        continue;
                    }
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Body,
                        Value::Body(std::sync::OnceLock::new()),
                        self.address.with(Step::new(Kind::Body, 0)),
                    )));
                },
                Source::Marks { children, then } => match children.next_node(self.docx) {
                    Some(item) => return Some(item),
                    None => {
                        self.source = std::mem::replace(then.as_mut(), Source::Done);
                    },
                },
                Source::Parts { children, root } => match children.next_node(self.docx) {
                    Some(item) => return Some(item),
                    None => {
                        self.source = if *root {
                            Source::CommentsRoot
                        } else {
                            Source::Done
                        };
                    },
                },
                Source::CommentsRoot => {
                    self.source = Source::SmartArtRoot;
                    if !self.kinds.contains(Kind::Comments) {
                        continue;
                    }
                    match nav::has_comments(self.docx) {
                        Ok(true) => {
                            return Some(Ok(Node::new(
                                self.docx,
                                Kind::Comments,
                                Value::Comments,
                                self.address.with(Step::new(Kind::Comments, 0)),
                            )));
                        },
                        Ok(false) => continue,
                        Err(error) => return Some(Err(error)),
                    }
                },
                Source::SmartArtRoot => {
                    if !self.kinds.contains(Kind::SmartArt) {
                        self.source = Source::Done;
                        continue;
                    }
                    match super::smartart::document_smartarts(self.docx) {
                        Ok(graphics) => {
                            self.source = Source::SmartArts { graphics, next: 0 };
                        },
                        Err(error) => {
                            self.source = Source::Done;
                            return Some(Err(error));
                        },
                    }
                },
                Source::Built(items) => {
                    let (kind, value, step) = items.next()?;
                    return Some(Ok(Node::new(
                        self.docx,
                        kind,
                        value,
                        self.address.with(step),
                    )));
                },
                Source::SmartArts { graphics, next } => {
                    let graphic = Arc::clone(graphics.get(*next)?);
                    *next += 1;
                    return Some(Ok(super::smartart::smartart_node(self.docx, graphic)));
                },
                Source::Comments { comments, next } => {
                    let comment = comments.next()?;
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Comment,
                        Value::Comment(Arc::new(comment)),
                        self.address.with(Step::new(Kind::Comment, ordinal)),
                    )));
                },
                Source::Blocks { container, next } => {
                    let block = container.get(litchi_core::Position::new(*next))?;
                    *next += 1;
                    let Some(kind) = Kind::of_block(block.kind()) else {
                        continue;
                    };
                    if !self.kinds.contains(kind) {
                        continue;
                    }
                    return Some(block_node(
                        self.docx,
                        kind,
                        container,
                        &block,
                        &self.address,
                    ));
                },
                Source::Rows { rows, next } => {
                    let row = rows.next()?;
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Row,
                        Value::Row(Arc::new(row)),
                        self.address.with(Step::new(Kind::Row, ordinal)),
                    )));
                },
                Source::Cells { cells, next } => {
                    let cell = cells.next()?;
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Cell,
                        Value::Cell(cell),
                        self.address.with(Step::new(Kind::Cell, ordinal)),
                    )));
                },
                Source::Runs { runs, next, rest } => {
                    if let Some(run) = runs.next() {
                        let ordinal = *next;
                        *next += 1;
                        return Some(Ok(Node::new(
                            self.docx,
                            Kind::Run,
                            Value::Run(run),
                            self.address.with(Step::new(Kind::Run, ordinal)),
                        )));
                    }
                    self.source = rest.take().map_or(Source::Done, Source::Links);
                },
                Source::Links(rest) => {
                    let pictures = rest.pictures.then(|| rest.paragraph.clone());
                    let charts = rest.charts.take();
                    if !rest.links {
                        self.source = Source::after_links(pictures, charts);
                        continue;
                    }
                    match nav::hyperlinks(self.docx, &rest.paragraph) {
                        Ok(links) => {
                            self.source = Source::Hyperlinks {
                                links: links.into_iter(),
                                next: 0,
                                pictures,
                                charts,
                            };
                        },
                        Err(error) => return Some(Err(error)),
                    }
                },
                Source::PictureList(paragraph, charts) => match paragraph.images() {
                    Ok(pictures) => {
                        self.source = Source::Pictures {
                            pictures: pictures.into_vec().into_iter(),
                            next: 0,
                            charts: charts.take(),
                        };
                    },
                    Err(error) => return Some(Err(error.into())),
                },
                Source::ChartList(paragraph) => {
                    let Value::Paragraph {
                        paragraph: value, ..
                    } = &paragraph.value
                    else {
                        self.source = Source::Done;
                        continue;
                    };
                    match value.charts() {
                        Ok(charts) => {
                            let paragraph = std::mem::replace(
                                paragraph,
                                Box::new(Charted {
                                    value: Value::Document,
                                    address: Address::root(),
                                }),
                            );
                            self.source = Source::Charts {
                                charts: charts.into_vec().into_iter(),
                                next: 0,
                                paragraph,
                            };
                        },
                        Err(error) => return Some(Err(error.into())),
                    }
                },
                Source::Charts {
                    charts,
                    next,
                    paragraph,
                } => {
                    let chart = charts.next()?;
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(super::chart::chart_node(
                        self.docx,
                        &paragraph.value,
                        &paragraph.address,
                        ordinal,
                        chart,
                    )));
                },
                Source::Pictures {
                    pictures,
                    next,
                    charts,
                } => {
                    let Some(picture) = pictures.next() else {
                        self.source = charts.take().map_or(Source::Done, Source::ChartList);
                        continue;
                    };
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Picture,
                        Value::Picture(Arc::new(picture)),
                        self.address.with(Step::new(Kind::Picture, ordinal)),
                    )));
                },
                Source::Hyperlinks {
                    links,
                    next,
                    pictures,
                    charts,
                } => {
                    let Some(link) = links.next() else {
                        self.source = Source::after_links(pictures.take(), charts.take());
                        continue;
                    };
                    let ordinal = *next;
                    *next += 1;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Hyperlink,
                        Value::Hyperlink(Arc::new(link)),
                        self.address.with(Step::new(Kind::Hyperlink, ordinal)),
                    )));
                },
            }
        }
    }
}

/// What follows a paragraph's runs: its hyperlinks and its pictures, each
/// when wanted.
struct Rest {
    paragraph: Paragraph,
    links: bool,
    pictures: bool,
    /// The paragraph's charts follow its pictures when wanted.
    charts: Option<Box<Charted>>,
}

/// Whether `kinds` holds a kind a block container lists.
fn wants_blocks(kinds: KindSet<Kind>) -> bool {
    [
        Kind::Paragraph,
        Kind::Table,
        Kind::Sdt,
        Kind::Equation,
        Kind::Section,
    ]
    .into_iter()
    .any(|kind| kinds.contains(kind))
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.source, Source::Done) {
            return None;
        }
        let owner = self.docx.owner();
        // Inside a resolver or query step for this owner, the step's
        // closing check fences the child: it observes a source changed
        // before or during this read, and its error wins (record 0315; the
        // walker still checks once per match, owner decision 2026-10-14).
        if crate::path::fence::is_active(owner.fence_key()) {
            let item = self.produce();
            if matches!(item, Some(Err(_))) {
                self.source = Source::Done;
            }
            return item;
        }
        let item = match owner.check_source() {
            Ok(()) => self.produce(),
            Err(error) => Some(Err(error)),
        };
        let item = match owner.check_source() {
            Ok(()) => item,
            Err(error) => Some(Err(error)),
        };
        if matches!(item, Some(Err(_))) {
            self.source = Source::Done;
        }
        item
    }
}

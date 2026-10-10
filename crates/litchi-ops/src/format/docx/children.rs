//! The children of a DOCX node, produced lazily in document order (record
//! 0010 §4.5.1, §5.3).
//!
//! | Parent | Children |
//! |---|---|
//! | `/` | `body` (its blocks are read when it is entered) |
//! | `body`, a cell, a content control | paragraphs, tables, content controls, display equations and (in the body) section breaks, in visible order; `w:customXml` wrappers flattened; `w:altChunk` anchors and unknown children skipped |
//! | a table | rows |
//! | a row | cells |
//! | a paragraph | its addressable runs, then its direct hyperlinks |
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
    /// The root's single child, the body.
    Root,
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
    /// A paragraph's runs; its hyperlinks follow when `links` holds it.
    Runs {
        runs: vec::IntoIter<Run>,
        next: u32,
        links: Option<Paragraph>,
    },
    /// A paragraph's hyperlinks, read when the first is needed.
    Links(Paragraph),
    Hyperlinks {
        links: vec::IntoIter<InlineHyperlink>,
        next: u32,
    },
}

impl<'f> Children<'f> {
    /// The children of `parent` whose kinds are in `kinds`.
    pub(super) fn of(parent: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let source = match &parent.value {
            Value::Document if kinds.contains(Kind::Body) => Source::Root,
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
            Value::Paragraph { paragraph, .. } => {
                let links = kinds.contains(Kind::Hyperlink).then(|| paragraph.clone());
                if kinds.contains(Kind::Run) {
                    Source::Runs {
                        runs: paragraph.addressable_runs()?.into_vec().into_iter(),
                        next: 0,
                        links,
                    }
                } else {
                    links.map_or(Source::Done, Source::Links)
                }
            },
            _ => Source::Done,
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
                    self.source = Source::Done;
                    return Some(Ok(Node::new(
                        self.docx,
                        Kind::Body,
                        Value::Body(std::sync::OnceLock::new()),
                        self.address.with(Step::new(Kind::Body, 0)),
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
                Source::Runs { runs, next, links } => {
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
                    self.source = links.take().map_or(Source::Done, Source::Links);
                },
                Source::Links(paragraph) => match nav::hyperlinks(self.docx, paragraph) {
                    Ok(links) => {
                        self.source = Source::Hyperlinks {
                            links: links.into_iter(),
                            next: 0,
                        };
                    },
                    Err(error) => return Some(Err(error)),
                },
                Source::Hyperlinks { links, next } => {
                    let link = links.next()?;
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
        let item = match self.docx.owner().check_source() {
            Ok(()) => self.produce(),
            Err(error) => Some(Err(error)),
        };
        let item = match self.docx.owner().check_source() {
            Ok(()) => item,
            Err(error) => Some(Err(error)),
        };
        if matches!(item, Some(Err(_))) {
            self.source = Source::Done;
        }
        item
    }
}

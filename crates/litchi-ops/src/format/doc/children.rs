//! Lazy children for the DOC node adapter.
//!
//! The main story streams: a body without tables walks the native body
//! paragraph traversal (borrowed text, one scan); a body with tables walks
//! the memoized story index in document order and draws each paragraph's
//! text from the same traversal. Every other parent lists the few native
//! values the parent already holds (rows, cells, a paragraph's runs, a
//! story's paragraphs) or the document's stories.

use std::sync::Arc;

use litchi_core::Position;

use super::Doc;
use super::node::{Item, Node, Story};
use super::spec::Kind;
use crate::error::Result;
use crate::path::navigate::KindSet;

/// Children of a DOC node.
pub(crate) struct Children<'f> {
    doc: &'f Doc,
    state: State<'f>,
    /// Nodes listed after the streamed body: its text boxes and shapes.
    tail: Option<std::vec::IntoIter<Node<'f>>>,
}

enum State<'f> {
    Done,
    Error(crate::Error),
    /// The body of a document without top-level tables.
    Paragraphs {
        paragraphs: litchi_doc::BodyParagraphs<'f>,
        next: usize,
    },
    /// The body of a document with top-level tables, in block order.
    Blocks {
        blocks: litchi_doc::StoryBlocks<'f>,
        cursor: usize,
        paragraphs: litchi_doc::BodyParagraphs<'f>,
        include_paragraphs: bool,
        include_tables: bool,
    },
    /// Nodes listed up front.
    List(std::vec::IntoIter<Node<'f>>),
}

impl<'f> Children<'f> {
    /// Every child of `parent`.
    pub(super) fn of(parent: &Node<'f>) -> Self {
        Self::filtered(parent, None)
    }

    /// The children of `parent` of the kinds in `kinds` (and the containers
    /// that lead to them); `None` lists every child.
    /// Views (`kinds` is `None`) list the structural kinds only; the parts
    /// and marks of record 0324 are listed when a query asks for them.
    pub(super) fn filtered(parent: &Node<'f>, kinds: Option<KindSet<Kind>>) -> Self {
        let wants = |kind: Kind| kinds.map_or(kind.is_structural(), |kinds| kinds.contains(kind));
        let doc = parent.doc;
        let (state, tail) = match Self::state(parent, &wants) {
            Ok(state) => match super::nav::tail(parent, &wants) {
                Ok(tail) => (state, tail.map(Vec::into_iter)),
                Err(error) => (State::Error(error), None),
            },
            Err(error) => (State::Error(error), None),
        };
        Self { doc, state, tail }
    }

    pub(super) const fn empty(doc: &'f Doc) -> Self {
        Self {
            doc,
            state: State::Done,
            tail: None,
        }
    }

    fn state(parent: &Node<'f>, wants: &dyn Fn(Kind) -> bool) -> Result<State<'f>> {
        let doc = parent.doc;
        Ok(match (&parent.item, parent.kind) {
            (Item::Bare, Kind::Document) => State::List(root_children(doc, wants)?.into_iter()),
            (Item::Bare, Kind::Body) => {
                let include_paragraphs = wants(Kind::Paragraph);
                let include_tables = wants(Kind::Table);
                let paragraphs = doc.body_paragraphs()?;
                let blocks = doc.reader()?.body_blocks()?;
                if blocks.count(litchi_doc::StoryBlockKind::Table) == 0 {
                    if include_paragraphs {
                        State::Paragraphs {
                            paragraphs,
                            next: 0,
                        }
                    } else {
                        State::Done
                    }
                } else {
                    State::Blocks {
                        blocks,
                        cursor: 0,
                        paragraphs,
                        include_paragraphs,
                        include_tables,
                    }
                }
            },
            (Item::Bare, Kind::Comments) => {
                let reader = doc.reader()?;
                let mut nodes = Vec::with_capacity(reader.comment_count());
                for (index, comment) in reader.comments()?.into_iter().enumerate() {
                    nodes.push(parent.child(
                        Kind::Comment,
                        Item::Story {
                            story: Arc::new(Story::Comment(comment)),
                            number: 0,
                        },
                        Position::new(index),
                    ));
                }
                State::List(nodes.into_iter())
            },
            (Item::BodyParagraph { .. } | Item::Paragraph(_), _) => {
                let Some(paragraph) = parent.native_paragraph()? else {
                    return Ok(State::Done);
                };
                let mut nodes = if wants(Kind::Run) {
                    run_nodes(parent, paragraph)?
                } else {
                    Vec::new()
                };
                if wants(Kind::Hyperlink) {
                    nodes.extend(super::nav::link_nodes(parent, paragraph)?);
                }
                if wants(Kind::Tab) {
                    nodes.extend(super::nav::tab_nodes(parent, paragraph));
                }
                State::List(nodes.into_iter())
            },
            (Item::Bare, Kind::Styles) => State::List(super::nav::style_nodes(parent)?.into_iter()),
            (Item::Bare, Kind::Numbering) => {
                State::List(super::nav::numbering_nodes(parent, wants).into_iter())
            },
            (Item::AbstractNum(at), _) => {
                State::List(super::nav::level_nodes(parent, *at).into_iter())
            },
            (Item::Table(table), _) => State::List(
                table
                    .rows()?
                    .into_iter()
                    .enumerate()
                    .map(|(index, row)| {
                        parent.child(Kind::Row, Item::Row(Arc::new(row)), Position::new(index))
                    })
                    .collect::<Vec<_>>()
                    .into_iter(),
            ),
            (Item::Row(row), _) => State::List(
                row.cells()?
                    .into_iter()
                    .enumerate()
                    .map(|(index, cell)| {
                        parent.child(Kind::Cell, Item::Cell(Arc::new(cell)), Position::new(index))
                    })
                    .collect::<Vec<_>>()
                    .into_iter(),
            ),
            (Item::Cell(cell), _) => {
                State::List(paragraph_nodes(parent, cell.paragraphs()?.into_iter()).into_iter())
            },
            (Item::Story { story, .. }, _) => {
                State::List(paragraph_nodes(parent, story.paragraphs().iter().cloned()).into_iter())
            },
            _ => State::Done,
        })
    }
}

/// The root's children: the body, then each non-empty header and footer
/// story, each footnote and endnote, and the comment collection.
fn root_children<'f>(doc: &'f Doc, wants: &dyn Fn(Kind) -> bool) -> Result<Vec<Node<'f>>> {
    let below = wants;
    let reader = doc.reader()?;
    let mut nodes = Vec::new();
    if wants(Kind::Body) {
        nodes.push(Node::body(doc));
    }
    if below(Kind::Header) {
        for (index, story) in reader.headers()?.into_iter().enumerate() {
            nodes.push(Node::story(
                doc,
                Kind::Header,
                Arc::new(Story::HeaderFooter(story)),
                Position::new(index),
            ));
        }
    }
    if below(Kind::Footer) {
        for (index, story) in reader.footers()?.into_iter().enumerate() {
            nodes.push(Node::story(
                doc,
                Kind::Footer,
                Arc::new(Story::HeaderFooter(story)),
                Position::new(index),
            ));
        }
    }
    if below(Kind::Footnote) {
        for (index, note) in reader.footnotes()?.into_iter().enumerate() {
            nodes.push(Node::story(
                doc,
                Kind::Footnote,
                Arc::new(Story::Note(note)),
                Position::new(index),
            ));
        }
    }
    if below(Kind::Endnote) {
        for (index, note) in reader.endnotes()?.into_iter().enumerate() {
            nodes.push(Node::story(
                doc,
                Kind::Endnote,
                Arc::new(Story::Note(note)),
                Position::new(index),
            ));
        }
    }
    if wants(Kind::Comments) && reader.comment_count() > 0 {
        nodes.push(Node::comments(doc));
    }
    super::nav::root_marks(doc, wants, &mut nodes)?;
    Ok(nodes)
}

/// The runs of a paragraph node. Runs of a main-story paragraph carry their
/// UTF-16 range in the paragraph's text, the leaf's run edit target.
pub(super) fn run_nodes<'f>(
    parent: &Node<'f>,
    paragraph: &litchi_doc::Paragraph,
) -> Result<Vec<Node<'f>>> {
    let start = paragraph.content_cp_range().map(|range| range.start());
    let body = parent.paragraph_position();
    Ok(paragraph
        .runs()?
        .into_iter()
        .enumerate()
        .map(|(index, run)| {
            let span = match (body, start, run.cp_range()) {
                (Some(position), Some(start), Some(range)) => range
                    .start()
                    .checked_sub(start)
                    .zip(range.end().checked_sub(start))
                    .map(|(from, to)| (position, from, to)),
                _ => None,
            };
            parent.child(
                Kind::Run,
                Item::Run {
                    run: Arc::new(run),
                    body: span,
                },
                Position::new(index),
            )
        })
        .collect())
}

fn paragraph_nodes<'f>(
    parent: &Node<'f>,
    paragraphs: impl Iterator<Item = litchi_doc::Paragraph>,
) -> Vec<Node<'f>> {
    paragraphs
        .enumerate()
        .map(|(index, paragraph)| {
            parent.child(
                Kind::Paragraph,
                Item::Paragraph(Arc::new(paragraph)),
                Position::new(index),
            )
        })
        .collect()
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            match &mut self.state {
                State::Done => match self.tail.take() {
                    Some(tail) => self.state = State::List(tail),
                    None => return None,
                },
                State::Error(_) => {
                    self.tail = None;
                    let State::Error(error) = std::mem::replace(&mut self.state, State::Done)
                    else {
                        return None;
                    };
                    return Some(Err(error));
                },
                State::List(nodes) => match nodes.next() {
                    Some(node) => return Some(Ok(node)),
                    None => self.state = State::Done,
                },
                State::Paragraphs { paragraphs, next } => {
                    return match paragraphs.next() {
                        Some(Ok(text)) => {
                            let position = Position::new(*next);
                            *next = next.saturating_add(1);
                            Some(Ok(Node::paragraph_with_text(self.doc, position, text)))
                        },
                        Some(Err(error)) => {
                            self.state = State::Done;
                            Some(Err(error.into()))
                        },
                        None => {
                            self.state = State::Done;
                            continue;
                        },
                    };
                },
                State::Blocks {
                    blocks,
                    cursor,
                    paragraphs,
                    include_paragraphs,
                    include_tables,
                } => {
                    let Some(block) = blocks.get(Position::new(*cursor)) else {
                        self.state = State::Done;
                        continue;
                    };
                    *cursor = cursor.saturating_add(1);
                    let ordinal = Position::new(block.ordinal());
                    match block.kind() {
                        litchi_doc::StoryBlockKind::Paragraph => {
                            // The body traversal yields exactly the paragraphs
                            // outside tables, in the order of these blocks.
                            let text = match paragraphs.next() {
                                Some(Ok(text)) => Some(text),
                                Some(Err(error)) => {
                                    self.state = State::Done;
                                    return Some(Err(error.into()));
                                },
                                None => None,
                            };
                            if !*include_paragraphs {
                                continue;
                            }
                            return Some(Ok(match text {
                                Some(text) => Node::paragraph_with_text(self.doc, ordinal, text),
                                None => Node::paragraph(self.doc, ordinal),
                            }));
                        },
                        litchi_doc::StoryBlockKind::Table => {
                            if !*include_tables {
                                continue;
                            }
                            return Some(match blocks.table(ordinal) {
                                Ok(Some(table)) => Ok(Node::body(self.doc).child(
                                    Kind::Table,
                                    Item::Table(Arc::new(table)),
                                    ordinal,
                                )),
                                Ok(None) => {
                                    self.state = State::Done;
                                    continue;
                                },
                                Err(error) => {
                                    self.state = State::Done;
                                    Err(error.into())
                                },
                            });
                        },
                        _ => continue,
                    }
                },
            }
        }
    }
}

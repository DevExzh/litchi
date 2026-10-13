//! Path and query navigation for the DOC vocabulary.

use std::borrow::Cow;
use std::fmt;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};

use super::Doc;
use super::children::{Children, run_nodes};
use super::node::{Item, Node, Story};
use super::spec::{Kind, SPEC};
use crate::error::Result;
use crate::node::spec::FormatSpec;
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::Inspect;

/// The DOC adapter used by path resolution, queries and generic views.
#[derive(Clone, Copy)]
pub(crate) struct Nav<'f> {
    doc: &'f Doc,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(doc: &'f Doc) -> Self {
        Self { doc }
    }
}

/// The zero-based index `select` names among `count` children, or `None`
/// when it names none. `Number` is a note's 1-based ordinal.
fn pick(select: &Select<'_>, count: usize) -> Option<Option<usize>> {
    Some(match select {
        Select::None => (count > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
        Select::Last => count.checked_sub(1),
        Select::Number(number) => number.checked_sub(1).filter(|index| *index < count),
        _ => return None,
    })
}

fn found<N, G>(node: Option<N>) -> Lookup<N, G> {
    node.map_or(Lookup::Missing, Lookup::Found)
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = crate::node::spec::GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Doc;
    const ABSENT: &'static [&'static str] = &["endnotes", "footnotes", "sdt"];

    fn check_source(&self) -> Result<()> {
        self.doc.check_source()
    }

    fn root(&self) -> Result<Self::Node> {
        Ok(Node::root(self.doc))
    }

    fn child(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
    ) -> Result<Lookup<Self::Node, Self::Gap>> {
        let doc = self.doc;
        let select = &step.select;
        match (parent.kind, &parent.item, step.kind) {
            (Kind::Document, _, Kind::Body) => match select {
                Select::None | Select::Last => Ok(Lookup::Found(Node::body(doc))),
                Select::Ordinal(position) if position.get() == 0 => {
                    Ok(Lookup::Found(Node::body(doc)))
                },
                Select::Ordinal(_) => Ok(Lookup::Missing),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Document, _, kind @ (Kind::Header | Kind::Footer)) => {
                let reader = doc.reader()?;
                let stories = if kind == Kind::Header {
                    reader.headers()?
                } else {
                    reader.footers()?
                };
                let index = match select {
                    // `header[@type=first]`: the first story of that kind.
                    Select::Identity(identities) => {
                        let mut wanted = None;
                        for (identity, value) in identities.iter() {
                            if !identity.key.eq_ignore_ascii_case("type") {
                                return Ok(Lookup::NotApplicable);
                            }
                            wanted = Some(value);
                        }
                        let Some(wanted) = wanted else {
                            return Ok(Lookup::NotApplicable);
                        };
                        Some(stories.iter().position(|story| {
                            super::node::story_type(story.header_footer_type)
                                .eq_ignore_ascii_case(wanted)
                        }))
                    },
                    _ => pick(select, stories.len()),
                };
                let Some(index) = index else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    stories.into_iter().nth(index).map(|story| {
                        Node::story(
                            doc,
                            kind,
                            Arc::new(Story::HeaderFooter(story)),
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Document, _, kind @ (Kind::Footnote | Kind::Endnote)) => {
                let reader = doc.reader()?;
                let count = if kind == Kind::Footnote {
                    reader.footnote_count()
                } else {
                    reader.endnote_count()
                };
                let Some(index) = pick(select, count) else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = index else {
                    return Ok(Lookup::Missing);
                };
                let note = if kind == Kind::Footnote {
                    reader.footnote(Position::new(index))?
                } else {
                    reader.endnote(Position::new(index))?
                };
                Ok(found(note.map(|note| {
                    Node::story(doc, kind, Arc::new(Story::Note(note)), Position::new(index))
                })))
            },
            (Kind::Document, _, Kind::Comments) => match select {
                Select::None | Select::Last => Ok(if doc.reader()?.comment_count() > 0 {
                    Lookup::Found(Node::comments(doc))
                } else {
                    Lookup::Missing
                }),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Comments, _, Kind::Comment) => {
                let reader = doc.reader()?;
                let Some(index) = pick(select, reader.comment_count()) else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = index else {
                    return Ok(Lookup::Missing);
                };
                Ok(found(reader.comment(Position::new(index))?.map(
                    |comment| {
                        parent.child(
                            Kind::Comment,
                            Item::Story {
                                story: Arc::new(Story::Comment(comment)),
                                number: 0,
                            },
                            Position::new(index),
                        )
                    },
                )))
            },
            (Kind::Body, _, Kind::Paragraph) => {
                let position = match select {
                    Select::None => Position::new(0),
                    Select::Ordinal(position) => *position,
                    Select::Last => match doc.body_paragraph_count()?.checked_sub(1) {
                        Some(last) => Position::new(last),
                        None => return Ok(Lookup::Missing),
                    },
                    _ => return Ok(Lookup::NotApplicable),
                };
                Ok(found(doc.body_paragraph_at(position)?.map(|text| {
                    Node::paragraph_with_text(doc, position, text)
                })))
            },
            (Kind::Body, _, Kind::Table) => {
                let blocks = doc.reader()?.body_blocks()?;
                let Some(index) = pick(select, blocks.count(litchi_doc::StoryBlockKind::Table))
                else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = index else {
                    return Ok(Lookup::Missing);
                };
                Ok(found(blocks.table(Position::new(index))?.map(|table| {
                    parent.child(
                        Kind::Table,
                        Item::Table(Arc::new(table)),
                        Position::new(index),
                    )
                })))
            },
            (_, Item::Table(table), Kind::Row) => {
                let rows = table.rows()?;
                let Some(index) = pick(select, rows.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    rows.into_iter().nth(index).map(|row| {
                        parent.child(Kind::Row, Item::Row(Arc::new(row)), Position::new(index))
                    })
                })))
            },
            (_, Item::Row(row), Kind::Cell) => {
                let cells = row.cells()?;
                let Some(index) = pick(select, cells.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    cells.into_iter().nth(index).map(|cell| {
                        parent.child(Kind::Cell, Item::Cell(Arc::new(cell)), Position::new(index))
                    })
                })))
            },
            (_, Item::Cell(cell), Kind::Paragraph) => {
                let paragraphs = cell.paragraphs()?;
                self.story_paragraph(parent, paragraphs, select)
            },
            (_, Item::Story { story, .. }, Kind::Paragraph) => {
                let paragraphs = story.paragraphs().to_vec();
                self.story_paragraph(parent, paragraphs, select)
            },
            (Kind::Paragraph, _, Kind::Run) => {
                let Some(paragraph) = parent.native_paragraph()? else {
                    return Ok(Lookup::Missing);
                };
                let runs = run_nodes(parent, paragraph)?;
                let Some(index) = pick(select, runs.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| runs.into_iter().nth(index))))
            },
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children> {
        let wanted = SPEC.elements[parent.kind as usize]
            .children
            .iter()
            .any(|kind| kinds.contains(*kind));
        Ok(if wanted {
            Children::filtered(parent, Some(kinds))
        } else {
            Children::empty(self.doc)
        })
    }

    fn available(
        &self,
        parent: &Self::Node,
        _: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        let reader = self.doc.reader()?;
        match (&parent.item, parent.kind) {
            (Item::Bare, Kind::Document) => {
                out.kind("body", 1);
                out.kind("header", reader.headers()?.len());
                out.kind("footer", reader.footers()?.len());
                out.kind("footnote", reader.footnote_count());
                out.kind("endnote", reader.endnote_count());
                out.kind("comments", usize::from(reader.comment_count() > 0));
            },
            (Item::Bare, Kind::Body) => {
                out.kind("p", self.doc.body_paragraph_count()?);
                out.kind(
                    "tbl",
                    reader
                        .body_blocks()?
                        .count(litchi_doc::StoryBlockKind::Table),
                );
            },
            (Item::Bare, Kind::Comments) => out.kind("comment", reader.comment_count()),
            (Item::Table(table), _) => out.kind("tr", table.row_count()?),
            (Item::Row(row), _) => out.kind("tc", row.cell_count()?),
            (Item::Cell(cell), _) => out.kind("p", cell.paragraphs()?.len()),
            (Item::Story { story, .. }, _) => out.kind("p", story.paragraphs().len()),
            (Item::BodyParagraph { .. } | Item::Paragraph(_), _) => out.kind(
                "r",
                parent
                    .native_paragraph()?
                    .map_or(Ok(0), |paragraph| paragraph.runs().map(|runs| runs.len()))?,
            ),
            _ => {},
        }
        Ok(())
    }
}

impl<'f> Nav<'f> {
    fn story_paragraph(
        &self,
        parent: &Node<'f>,
        paragraphs: Vec<litchi_doc::Paragraph>,
        select: &Select<'_>,
    ) -> Result<Lookup<Node<'f>, crate::node::spec::GapSpec>> {
        let Some(index) = pick(select, paragraphs.len()) else {
            return Ok(Lookup::NotApplicable);
        };
        Ok(found(index.and_then(|index| {
            paragraphs.into_iter().nth(index).map(|paragraph| {
                parent.child(
                    Kind::Paragraph,
                    Item::Paragraph(Arc::new(paragraph)),
                    Position::new(index),
                )
            })
        })))
    }
}

impl<'f> Inspect<'f> for Nav<'f> {
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Body];

    fn kind(&self, node: &Node<'f>) -> Kind {
        node.kind
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.kind.type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        node.text()
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        node.format_props()
    }

    fn number(&self, node: &Node<'f>) -> Option<usize> {
        node.number()
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        let step = match kind {
            Kind::Body => PathStep::new("body"),
            Kind::Comments => PathStep::new("comments"),
            Kind::Document => PathStep::new("document"),
            _ => node.path().steps().last().copied().unwrap_or_else(|| {
                PathStep::at(crate::path::navigate::ElementKind::name(kind), ordinal)
            }),
        };
        fmt::Display::fmt(&step, out)
    }
}

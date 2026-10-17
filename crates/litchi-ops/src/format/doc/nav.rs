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

/// The single identity value of `select` for `key`: `None` when `select`
/// names no identity, `Some(None)` when it names another key or the
/// stacked values disagree.
fn identity<'s>(select: &Select<'s>, key: &str) -> Option<Option<&'s str>> {
    let Select::Identity(identities) = select else {
        return None;
    };
    let mut wanted = None;
    for (identity, value) in identities.iter() {
        if !identity.key.eq_ignore_ascii_case(key) || wanted.is_some_and(|held| held != value) {
            return Some(None);
        }
        wanted = Some(value);
    }
    Some(wanted)
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
            (Kind::Paragraph, _, kind @ (Kind::Hyperlink | Kind::Tab)) => {
                let Some(paragraph) = parent.native_paragraph()? else {
                    return Ok(Lookup::Missing);
                };
                let nodes = if kind == Kind::Hyperlink {
                    link_nodes(parent, paragraph)?
                } else {
                    tab_nodes(parent, paragraph)
                };
                let Some(index) = pick(select, nodes.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| nodes.into_iter().nth(index))))
            },
            (Kind::Document, _, kind @ (Kind::Styles | Kind::Numbering)) => {
                let reader = doc.reader()?;
                let present = if kind == Kind::Styles {
                    reader.stylesheet().is_some()
                } else {
                    reader.list_tables().is_some()
                };
                match select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                Ok(if present {
                    Lookup::Found(Node::top(doc, kind, Item::Bare, Position::new(0)))
                } else {
                    Lookup::Missing
                })
            },
            (Kind::Styles, _, Kind::Style) => match select {
                Select::Name { name, .. } => Ok(found(
                    super::parts::find_style(doc, name)
                        .map(|istd| style_node(parent, usize::from(istd))),
                )),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Numbering, _, kind @ (Kind::AbstractNum | Kind::Num)) => {
                let count = if kind == Kind::AbstractNum {
                    super::parts::structures(doc).len()
                } else {
                    super::parts::num_count(doc)
                };
                let index = match identity(select, "id") {
                    Some(wanted) => {
                        let Some(wanted) =
                            wanted.and_then(|value| value.trim().parse::<usize>().ok())
                        else {
                            return Ok(Lookup::NotApplicable);
                        };
                        // A definition's id is its index; an instance's id is
                        // its one-based `PlfLfo` index.
                        let index = if kind == Kind::AbstractNum {
                            Some(wanted)
                        } else {
                            wanted.checked_sub(1)
                        };
                        index.filter(|index| *index < count)
                    },
                    None => {
                        let Some(index) = pick(select, count) else {
                            return Ok(Lookup::NotApplicable);
                        };
                        index
                    },
                };
                Ok(found(index.map(|index| {
                    let item = if kind == Kind::AbstractNum {
                        Item::AbstractNum(index)
                    } else {
                        Item::Num(index)
                    };
                    parent.child(kind, item, Position::new(index))
                })))
            },
            (Kind::AbstractNum, Item::AbstractNum(at), Kind::Level) => {
                let levels = super::parts::structures(doc)
                    .get(*at)
                    .map_or(&[][..], |structure| structure.levels.as_slice());
                // `level[N]` is the Nth level (1-based); `[@ilvl=L]` (raw
                // paths) names the level number.
                let index = match select {
                    Select::Number(number) => {
                        number.checked_sub(1).filter(|index| *index < levels.len())
                    },
                    _ => match identity(select, "ilvl") {
                        Some(wanted) => {
                            let Some(wanted) =
                                wanted.and_then(|value| value.trim().parse::<u8>().ok())
                            else {
                                return Ok(Lookup::NotApplicable);
                            };
                            levels.iter().position(|level| level.level == wanted)
                        },
                        None => {
                            let Some(index) = pick(select, levels.len()) else {
                                return Ok(Lookup::NotApplicable);
                            };
                            index
                        },
                    },
                };
                Ok(found(index.and_then(|index| {
                    levels
                        .get(index)
                        .map(|level| level_node(parent, *at, index, level))
                })))
            },
            (Kind::Document, _, Kind::Bookmark) => {
                let bookmarks = doc.reader()?.bookmarks()?;
                let index = match identity(select, "name") {
                    Some(Some(wanted)) => bookmarks
                        .iter()
                        .position(|bookmark| bookmark.name == wanted),
                    Some(None) => return Ok(Lookup::NotApplicable),
                    None => {
                        let Some(index) = pick(select, bookmarks.len()) else {
                            return Ok(Lookup::NotApplicable);
                        };
                        index
                    },
                };
                Ok(found(index.and_then(|index| {
                    bookmarks.into_iter().nth(index).map(|bookmark| {
                        Node::top(
                            doc,
                            Kind::Bookmark,
                            Item::Bookmark(Arc::new(bookmark)),
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Document, _, kind @ (Kind::Field | Kind::Toc)) => {
                let fields = if kind == Kind::Field {
                    super::marks::fields(doc)?
                } else {
                    super::marks::tocs(doc)?
                };
                let Some(index) = pick(select, fields.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    fields.into_iter().nth(index).map(|field| {
                        let field = Arc::new(field);
                        let item = if kind == Kind::Field {
                            Item::Field(field)
                        } else {
                            Item::Toc(field)
                        };
                        Node::top(doc, kind, item, Position::new(index))
                    })
                })))
            },
            (Kind::Document, _, Kind::Section) => {
                let sections = doc.reader()?.sections();
                let Some(index) = pick(select, sections.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    sections.get(index).map(|section| {
                        Node::top(
                            doc,
                            Kind::Section,
                            Item::Section(section),
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Body, _, Kind::TextBox) => {
                let text_boxes = super::marks::text_boxes(doc)?;
                let Some(index) = pick(select, text_boxes.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    text_boxes.into_iter().nth(index).map(|text_box| {
                        parent.child(
                            Kind::TextBox,
                            Item::TextBox(Arc::new(text_box)),
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Body, _, Kind::Shape) => {
                let anchors = doc.reader()?.shape_positions();
                let Some(index) = pick(select, anchors.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    anchors.get(index).map(|anchor| {
                        parent.child(Kind::Shape, Item::Shape(anchor), Position::new(index))
                    })
                })))
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
            (Item::Bare, Kind::Numbering) => {
                out.kind("abstractNum", super::parts::structures(self.doc).len());
                out.kind("num", super::parts::num_count(self.doc));
            },
            (Item::AbstractNum(at), _) => out.kind(
                "level",
                super::parts::structures(self.doc)
                    .get(*at)
                    .map_or(0, |structure| structure.levels.len()),
            ),
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
            Kind::Styles => PathStep::new("styles"),
            Kind::Numbering => PathStep::new("numbering"),
            Kind::Document => PathStep::new("document"),
            _ => node.last_step().unwrap_or_else(|| {
                PathStep::at(crate::path::navigate::ElementKind::name(kind), ordinal)
            }),
        };
        fmt::Display::fmt(&step, out)
    }
}

// ---------------------------------------------------------------------------
// Parts and marks (record 0324)
// ---------------------------------------------------------------------------

/// The text boxes and shapes listed after the body's blocks, when wanted.
pub(super) fn tail<'f>(
    parent: &Node<'f>,
    wants: &dyn Fn(Kind) -> bool,
) -> Result<Option<Vec<Node<'f>>>> {
    if parent.kind != Kind::Body || !matches!(parent.item, Item::Bare) {
        return Ok(None);
    }
    let doc = parent.doc;
    let mut nodes = Vec::new();
    if wants(Kind::TextBox) {
        for (index, text_box) in super::marks::text_boxes(doc)?.into_iter().enumerate() {
            nodes.push(parent.child(
                Kind::TextBox,
                Item::TextBox(Arc::new(text_box)),
                Position::new(index),
            ));
        }
    }
    if wants(Kind::Shape) {
        for (index, anchor) in doc.reader()?.shape_positions().iter().enumerate() {
            nodes.push(parent.child(Kind::Shape, Item::Shape(anchor), Position::new(index)));
        }
    }
    Ok((!nodes.is_empty()).then_some(nodes))
}

/// The document-level parts and marks, after the stories, when wanted.
pub(super) fn root_marks<'f>(
    doc: &'f Doc,
    wants: &dyn Fn(Kind) -> bool,
    nodes: &mut Vec<Node<'f>>,
) -> Result<()> {
    let reader = doc.reader()?;
    if wants(Kind::Styles) && reader.stylesheet().is_some() {
        nodes.push(Node::top(doc, Kind::Styles, Item::Bare, Position::new(0)));
    }
    if wants(Kind::Numbering) && reader.list_tables().is_some() {
        nodes.push(Node::top(
            doc,
            Kind::Numbering,
            Item::Bare,
            Position::new(0),
        ));
    }
    if wants(Kind::Bookmark) {
        for (index, bookmark) in reader.bookmarks()?.into_iter().enumerate() {
            nodes.push(Node::top(
                doc,
                Kind::Bookmark,
                Item::Bookmark(Arc::new(bookmark)),
                Position::new(index),
            ));
        }
    }
    if wants(Kind::Field) || wants(Kind::Toc) {
        let fields = super::marks::fields(doc)?;
        if wants(Kind::Field) {
            for (index, field) in fields.iter().enumerate() {
                nodes.push(Node::top(
                    doc,
                    Kind::Field,
                    Item::Field(Arc::new(field.clone())),
                    Position::new(index),
                ));
            }
        }
        if wants(Kind::Toc) {
            let tocs = fields.into_iter().filter(|field| {
                super::marks::keyword(&field.instruction).as_deref() == Some("TOC")
            });
            for (index, field) in tocs.enumerate() {
                nodes.push(Node::top(
                    doc,
                    Kind::Toc,
                    Item::Toc(Arc::new(field)),
                    Position::new(index),
                ));
            }
        }
    }
    if wants(Kind::Section) {
        for (index, section) in reader.sections().iter().enumerate() {
            nodes.push(Node::top(
                doc,
                Kind::Section,
                Item::Section(section),
                Position::new(index),
            ));
        }
    }
    Ok(())
}

/// Every style of the stylesheet, in `istd` order.
pub(super) fn style_nodes<'f>(parent: &Node<'f>) -> Result<Vec<Node<'f>>> {
    let Some(sheet) = parent.doc.reader()?.stylesheet() else {
        return Ok(Vec::new());
    };
    Ok(sheet
        .styles()
        .iter()
        .enumerate()
        .filter(|(_, slot)| slot.is_some())
        .map(|(istd, _)| style_node(parent, istd))
        .collect())
}

fn style_node<'f>(parent: &Node<'f>, istd: usize) -> Node<'f> {
    parent.child(
        Kind::Style,
        Item::Style(
            u16::try_from(istd).unwrap_or(u16::MAX),
            std::sync::OnceLock::new(),
        ),
        Position::new(istd),
    )
}

/// The list definitions then the list instances, when wanted.
pub(super) fn numbering_nodes<'f>(
    parent: &Node<'f>,
    wants: &dyn Fn(Kind) -> bool,
) -> Vec<Node<'f>> {
    let mut nodes = Vec::new();
    if wants(Kind::AbstractNum) {
        for index in 0..super::parts::structures(parent.doc).len() {
            nodes.push(parent.child(
                Kind::AbstractNum,
                Item::AbstractNum(index),
                Position::new(index),
            ));
        }
    }
    if wants(Kind::Num) {
        for index in 0..super::parts::num_count(parent.doc) {
            nodes.push(parent.child(Kind::Num, Item::Num(index), Position::new(index)));
        }
    }
    nodes
}

/// The levels of definition `at`.
pub(super) fn level_nodes<'f>(parent: &Node<'f>, at: usize) -> Vec<Node<'f>> {
    super::parts::structures(parent.doc)
        .get(at)
        .map_or(&[][..], |structure| structure.levels.as_slice())
        .iter()
        .enumerate()
        .map(|(index, level)| level_node(parent, at, index, level))
        .collect()
}

fn level_node<'f>(
    parent: &Node<'f>,
    definition: usize,
    index: usize,
    level: &'f litchi_doc::ListLevel,
) -> Node<'f> {
    parent.child(
        Kind::Level,
        Item::Level {
            definition,
            level,
            view: std::sync::OnceLock::new(),
        },
        Position::new(index),
    )
}

/// The `HYPERLINK` fields of a paragraph node.
pub(super) fn link_nodes<'f>(
    parent: &Node<'f>,
    paragraph: &litchi_doc::Paragraph,
) -> Result<Vec<Node<'f>>> {
    Ok(super::marks::links(parent.doc, paragraph)?
        .into_iter()
        .enumerate()
        .map(|(index, link)| {
            parent.child(
                Kind::Hyperlink,
                Item::Hyperlink(Arc::new(link)),
                Position::new(index),
            )
        })
        .collect())
}

/// The tab stops of a paragraph node.
pub(super) fn tab_nodes<'f>(parent: &Node<'f>, paragraph: &litchi_doc::Paragraph) -> Vec<Node<'f>> {
    paragraph
        .properties()
        .tab_stops
        .iter()
        .enumerate()
        .map(|(index, stop)| parent.child(Kind::Tab, Item::Tab(*stop), Position::new(index)))
        .collect()
}

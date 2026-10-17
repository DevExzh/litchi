//! Path and query navigation for the RTF vocabulary.

use std::borrow::Cow;
use std::fmt;

use litchi_core::{FileFormat, Position};

use super::Rtf;
use super::children::Children;
use super::node::{Item, Node};
use super::spec::{Kind, SPEC};
use super::stories;
use crate::error::Result;
use crate::node::spec::FormatSpec;
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::Inspect;

/// The RTF adapter used by path resolution, queries and generic views.
#[derive(Clone, Copy)]
pub(crate) struct Nav<'f> {
    doc: &'f Rtf,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(doc: &'f Rtf) -> Self {
        Self { doc }
    }
}

/// The zero-based index `select` names among `count` children: `Some(None)`
/// when it names none, `None` when the selector does not apply. `Number` is
/// a 1-based note number.
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

/// The single `[@key=value]` identity of `select` when its key is `key`.
fn identity<'s>(select: &Select<'s>, key: &str) -> Option<Option<&'s str>> {
    let Select::Identity(identities) = select else {
        return None;
    };
    let mut wanted = None;
    for (identity, value) in identities.iter() {
        if !identity.key.eq_ignore_ascii_case(key) {
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

    const FORMAT: FileFormat = FileFormat::Rtf;
    const ABSENT: &'static [&'static str] = &["sdt"];

    fn check_source(&self) -> Result<()> {
        Ok(())
    }

    fn root(&self) -> Result<Self::Node> {
        Ok(Node::root(self.doc))
    }

    #[allow(
        clippy::too_many_lines,
        reason = "one flat dispatch per native axis keeps the path grammar visible"
    )]
    fn child(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
    ) -> Result<Lookup<Self::Node, Self::Gap>> {
        let doc = self.doc;
        let package = &doc.0.package;
        let select = &step.select;
        match (parent.kind, step.kind) {
            (Kind::Document, Kind::Body) => match select {
                Select::None | Select::Last => Ok(Lookup::Found(Node::body(doc))),
                Select::Ordinal(position) if position.get() == 0 => {
                    Ok(Lookup::Found(Node::body(doc)))
                },
                Select::Ordinal(_) => Ok(Lookup::Missing),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Document, kind @ (Kind::Header | Kind::Footer)) => {
                let headers = kind == Kind::Header;
                let selected = if let Some(wanted) = identity(select, "type") {
                    let Some(wanted) = wanted else {
                        return Ok(Lookup::NotApplicable);
                    };
                    stories::header_footers(doc, headers)
                        .enumerate()
                        .find(|(_, story)| {
                            stories::story_type(story.value.header_type)
                                .eq_ignore_ascii_case(wanted)
                        })
                } else {
                    let count = stories::header_footers(doc, headers).count();
                    let Some(index) = pick(select, count) else {
                        return Ok(Lookup::NotApplicable);
                    };
                    index.and_then(|index| {
                        stories::header_footers(doc, headers)
                            .nth(index)
                            .map(|story| (index, story))
                    })
                };
                Ok(found(selected.map(|(index, story)| {
                    Node::top(
                        doc,
                        kind,
                        Item::Story {
                            section: story.section,
                            story: story.story,
                            value: story.value,
                        },
                        Position::new(index),
                    )
                })))
            },
            (Kind::Document, kind @ (Kind::Footnote | Kind::Endnote)) => {
                let footnotes = kind == Kind::Footnote;
                let count = stories::notes(doc, footnotes).count();
                // RTF notes carry no identifier; `[@footnoteId=N]` names the
                // Nth note of the kind, as `footnote[N]` does.
                let key = if footnotes { "footnoteId" } else { "endnoteId" };
                let index = if let Some(wanted) = identity(select, key) {
                    let Some(number) = wanted.and_then(|value| value.trim().parse::<usize>().ok())
                    else {
                        return Ok(Lookup::NotApplicable);
                    };
                    Some(number.checked_sub(1).filter(|index| *index < count))
                } else {
                    pick(select, count)
                };
                let Some(index) = index else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    stories::notes(doc, footnotes).nth(index).map(|note| {
                        Node::top(
                            doc,
                            kind,
                            Item::Note {
                                index: note.index,
                                number: note.number,
                                value: note.value,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Document, kind @ (Kind::Comments | Kind::Styles | Kind::Numbering)) => {
                let present = match kind {
                    Kind::Comments => !package.annotations().is_empty(),
                    Kind::Styles => !package.styles().is_empty(),
                    _ => !(package.lists().is_empty() && package.list_overrides().is_empty()),
                };
                match select {
                    Select::None | Select::Last => Ok(if present {
                        Lookup::Found(Node::collection(doc, kind))
                    } else {
                        Lookup::Missing
                    }),
                    _ => Ok(Lookup::NotApplicable),
                }
            },
            (Kind::Body, Kind::Paragraph) => {
                let position = match select {
                    Select::None => Some(0),
                    Select::Ordinal(position) => Some(position.get()),
                    Select::Last => package.paragraph_count().checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                Ok(found(position.and_then(|position| {
                    package
                        .body()
                        .paragraphs()
                        .nth(position)
                        .map(|paragraph| Node::paragraph(doc, Position::new(position), paragraph))
                })))
            },
            (Kind::Paragraph, Kind::Run) => {
                let (Some(paragraph), Some(paragraph_position)) =
                    (parent.body_paragraph(), parent.paragraph_position())
                else {
                    return Ok(Lookup::NotApplicable);
                };
                let position = match select {
                    Select::None => Some(0),
                    Select::Ordinal(position) => Some(position.get()),
                    Select::Last => paragraph.runs().count().checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                Ok(found(position.and_then(|position| {
                    paragraph
                        .runs()
                        .nth(position)
                        .map(|run| Node::run(doc, paragraph_position, Position::new(position), run))
                })))
            },
            (Kind::Paragraph, kind @ (Kind::Hyperlink | Kind::Picture)) => {
                let Some(paragraph) = parent.body_paragraph() else {
                    return Ok(Lookup::NotApplicable);
                };
                let range = stories::anchor_range(paragraph);
                let indices: Vec<usize> = if kind == Kind::Hyperlink {
                    stories::hyperlinks(doc, range).collect()
                } else {
                    stories::pictures(doc, range).collect()
                };
                let Some(index) = pick(select, indices.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|ordinal| {
                    let native = *indices.get(ordinal)?;
                    let item = if kind == Kind::Hyperlink {
                        Item::Hyperlink(package.fields().get(native)?)
                    } else {
                        Item::Picture {
                            index: native,
                            value: package.pictures().get(native)?,
                        }
                    };
                    Some(parent.child(kind, item, Position::new(ordinal)))
                })))
            },
            (Kind::Body, Kind::Table) => {
                let Some(index) = pick(select, package.tables().len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    package.tables().get(index).map(|table| {
                        parent.child(
                            Kind::Table,
                            Item::Table {
                                table: index,
                                value: table,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Table, Kind::Row) => {
                let Item::Table { table, value } = parent.item else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = pick(select, value.row_count()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    value.rows().get(index).map(|row| {
                        parent.child(
                            Kind::Row,
                            Item::Row {
                                table,
                                row: index,
                                value: row,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Row, Kind::Cell) => {
                let Item::Row { table, row, value } = parent.item else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = pick(select, value.cell_count()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    value.cells().get(index).map(|cell| {
                        parent.child(
                            Kind::Cell,
                            Item::Cell {
                                table,
                                row,
                                cell: index,
                                value: cell,
                                parent: value,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Header | Kind::Footer, Kind::Paragraph) => {
                let Item::Story {
                    section,
                    story,
                    value,
                } = parent.item
                else {
                    return Ok(Lookup::NotApplicable);
                };
                let Some(index) = pick(select, value.paragraphs.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    value.paragraphs.get(index).map(|paragraph| {
                        parent.child(
                            Kind::Paragraph,
                            Item::StoryParagraph {
                                section,
                                story,
                                index,
                                value: paragraph,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Comments, Kind::Comment) => {
                let index = if let Some(wanted) = identity(select, "commentId") {
                    let Some(wanted) = wanted.and_then(|value| value.trim().parse::<i32>().ok())
                    else {
                        return Ok(Lookup::NotApplicable);
                    };
                    Some(
                        package
                            .annotations()
                            .iter()
                            .position(|comment| comment.has_reference && comment.id == wanted),
                    )
                } else {
                    pick(select, package.annotations().len())
                };
                let Some(index) = index else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    package.annotations().get(index).map(|comment| {
                        parent.child(
                            Kind::Comment,
                            Item::Comment {
                                index,
                                value: comment,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Styles, Kind::Style) => match select {
                Select::Name { name, case } => Ok(found(
                    package
                        .styles()
                        .iter()
                        .enumerate()
                        .find(|(_, style)| case.matches(style.name.as_ref(), name))
                        .map(|(index, style)| {
                            parent.child(Kind::Style, Item::Style(style), Position::new(index))
                        }),
                )),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Numbering, kind @ (Kind::AbstractNum | Kind::Num)) => {
                let abstract_num = kind == Kind::AbstractNum;
                let count = if abstract_num {
                    package.lists().len()
                } else {
                    package.list_overrides().len()
                };
                let index = if let Some(wanted) = identity(select, "id") {
                    let Some(wanted) = wanted.and_then(|value| value.trim().parse::<i32>().ok())
                    else {
                        return Ok(Lookup::NotApplicable);
                    };
                    if abstract_num {
                        package.lists().iter().position(|list| list.id == wanted)
                    } else {
                        package
                            .list_overrides()
                            .iter()
                            .position(|list| list.index == wanted)
                    }
                } else {
                    let Some(index) = pick(select, count) else {
                        return Ok(Lookup::NotApplicable);
                    };
                    index
                };
                Ok(found(index.and_then(|index| {
                    let item = if abstract_num {
                        Item::AbstractNum(package.lists().get(index)?)
                    } else {
                        Item::Num(package.list_overrides().get(index)?)
                    };
                    Some(parent.child(kind, item, Position::new(index)))
                })))
            },
            (Kind::AbstractNum, Kind::Level) => {
                let Item::AbstractNum(list) = parent.item else {
                    return Ok(Lookup::NotApplicable);
                };
                // `level[N]` is the Nth level (1-based), whose ilvl is N-1 in
                // a well-formed list; `[@ilvl=L]` (raw paths) names the ilvl.
                let index = match select {
                    Select::Number(number) => number
                        .checked_sub(1)
                        .filter(|index| *index < list.levels.len()),
                    _ => {
                        if let Some(wanted) = identity(select, "ilvl") {
                            let Some(wanted) =
                                wanted.and_then(|value| value.trim().parse::<usize>().ok())
                            else {
                                return Ok(Lookup::NotApplicable);
                            };
                            list.levels
                                .iter()
                                .position(|level| usize::from(level.level) == wanted)
                        } else {
                            let Some(index) = pick(select, list.levels.len()) else {
                                return Ok(Lookup::NotApplicable);
                            };
                            index
                        }
                    },
                };
                Ok(found(index.and_then(|index| {
                    list.levels.get(index).map(|level| {
                        parent.child(Kind::Level, Item::Level(level), Position::new(index))
                    })
                })))
            },
            (Kind::Document, Kind::Section) => {
                let Some(index) = pick(select, package.sections().len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    package.sections().get(index).map(|value| {
                        Node::top(
                            doc,
                            Kind::Section,
                            Item::Section { index, value },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Document, Kind::Bookmark) => {
                let bookmarks = package.raw().bookmarks().bookmarks();
                let index = match identity(select, "name") {
                    Some(Some(wanted)) => bookmarks
                        .iter()
                        .position(|bookmark| bookmark.name.as_ref() == wanted),
                    Some(None) => return Ok(Lookup::NotApplicable),
                    None => {
                        let Some(index) = pick(select, bookmarks.len()) else {
                            return Ok(Lookup::NotApplicable);
                        };
                        index
                    },
                };
                Ok(found(index.and_then(|index| {
                    bookmarks.get(index).map(|value| {
                        Node::top(
                            doc,
                            Kind::Bookmark,
                            Item::Bookmark { index, value },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Document, kind @ (Kind::Field | Kind::Toc)) => {
                let indexes: Vec<usize> = super::parts::fields(doc, kind == Kind::Toc).collect();
                let Some(ordinal) = pick(select, indexes.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(ordinal.and_then(|ordinal| {
                    let index = *indexes.get(ordinal)?;
                    package.fields().get(index).map(|value| {
                        Node::top(
                            doc,
                            kind,
                            Item::Field { index, value },
                            Position::new(ordinal),
                        )
                    })
                })))
            },
            (Kind::Paragraph, Kind::Tab) => {
                let Some(paragraph) = parent.body_paragraph() else {
                    return Ok(Lookup::NotApplicable);
                };
                let stops = paragraph.format().tab_stops();
                let Some(index) = pick(select, stops.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|index| {
                    stops.get(index).map(|value| {
                        parent.child(
                            Kind::Tab,
                            Item::Tab {
                                index,
                                value: *value,
                            },
                            Position::new(index),
                        )
                    })
                })))
            },
            (Kind::Body, Kind::TextBox) => {
                let boxes = super::parts::text_boxes(doc);
                let Some(index) = pick(select, boxes.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(index.and_then(|ordinal| {
                    text_box_node(doc, parent, *boxes.get(ordinal)?, ordinal)
                })))
            },
            (Kind::Body, Kind::Shape) => {
                let indexes: Vec<usize> = super::parts::shapes(doc).collect();
                let Some(ordinal) = pick(select, indexes.len()) else {
                    return Ok(Lookup::NotApplicable);
                };
                Ok(found(ordinal.and_then(|ordinal| {
                    let index = *indexes.get(ordinal)?;
                    package.shapes().get(index).map(|value| {
                        parent.child(Kind::Shape, Item::Shape { value }, Position::new(ordinal))
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
            Children::empty()
        })
    }

    fn available(
        &self,
        parent: &Self::Node,
        _: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        let package = &self.doc.0.package;
        match parent.kind {
            Kind::Document => {
                out.kind("body", 1);
                out.kind("header", stories::header_footers(self.doc, true).count());
                out.kind("footer", stories::header_footers(self.doc, false).count());
                out.kind("footnote", stories::notes(self.doc, true).count());
                out.kind("endnote", stories::notes(self.doc, false).count());
                out.kind("comments", usize::from(!package.annotations().is_empty()));
            },
            Kind::Body => {
                out.kind("p", package.paragraph_count());
                out.kind("tbl", package.tables().len());
            },
            Kind::Paragraph => {
                if let Some(paragraph) = parent.body_paragraph() {
                    out.kind("r", paragraph.runs().count());
                    let range = stories::anchor_range(paragraph);
                    out.kind(
                        "hyperlink",
                        stories::hyperlinks(self.doc, range.clone()).count(),
                    );
                    out.kind("picture", stories::pictures(self.doc, range).count());
                }
            },
            _ => {
                let mut count = 0usize;
                let mut name = None;
                for child in Children::of(parent) {
                    let child = child?;
                    name = Some(super::node::kind_step(child.kind));
                    count += 1;
                }
                if let Some(name) = name {
                    out.kind(name, count);
                }
            },
        }
        Ok(())
    }
}

/// The text-box node at `ordinal` among the body's text boxes.
pub(super) fn text_box_node<'f>(
    doc: &'f Rtf,
    parent: &Node<'f>,
    reference: super::parts::TextBoxRef,
    ordinal: usize,
) -> Option<Node<'f>> {
    let package = &doc.0.package;
    let item = match reference {
        super::parts::TextBoxRef::Shape(index) => Item::ShapeTextBox {
            value: package.shapes().get(index)?,
        },
        super::parts::TextBoxRef::Legacy(index) => Item::LegacyTextBox {
            value: package.raw().legacy_text_boxes().get(index)?,
        },
    };
    Some(parent.child(Kind::TextBox, item, Position::new(ordinal)))
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
        Ok(node.format_props())
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

//! Lazy traversal of the native RTF stories behind unified nodes.
//!
//! The body's paragraph and run axes keep the native lazy cursors; every
//! other axis walks a native source-ordered collection by index. No second
//! tree is built.

use litchi_core::Position;

use super::node::{Item, Node};
use super::spec::{Kind, SPEC};
use super::stories;
use crate::Result;
use crate::path::navigate::KindSet;

/// The parent whose children are walked, with the native cursor of the axis
/// being walked. The hot main-story axes keep only what they read; any other
/// parent is kept by value (record 0287: an RTF node is small enough to copy,
/// which spares an allocation per walked container).
enum State<'f> {
    None,
    Body {
        doc: &'f super::Rtf,
        paragraphs: Option<litchi_rtf::text::Paragraphs<'f>>,
    },
    Paragraph {
        doc: &'f super::Rtf,
        position: Position,
        /// Where the paragraph's inline objects may be anchored: its text and
        /// the position of its terminating mark.
        anchors: std::ops::Range<usize>,
        runs: litchi_rtf::text::Runs<'f>,
        /// The paragraph's tab stops (record 0328).
        tabs: &'f [litchi_rtf::TabStop],
    },
    Other(Node<'f>),
}

/// The children of one node, in the order of its kind's child table.
pub(crate) struct Children<'f> {
    state: State<'f>,
    /// The parent's kind; its child table orders the phases.
    kind: Kind,
    /// Whether a view lists the children (header/footer stories list their
    /// paragraphs only to views).
    view: bool,
    filter: KindSet<Kind>,
    phase: u8,
    /// The zero-based position of the next child of the current kind.
    ordinal: usize,
    /// The next native index to inspect on a filtered collection axis.
    raw: usize,
}

impl<'f> Children<'f> {
    pub(super) const fn empty() -> Self {
        Self {
            state: State::None,
            kind: Kind::Run,
            view: false,
            filter: KindSet::EMPTY,
            phase: 0,
            ordinal: 0,
            raw: 0,
        }
    }

    /// Every child of `parent`.
    pub(super) fn of(parent: &Node<'f>) -> Self {
        Self::filtered(parent, None)
    }

    /// The children of `parent` whose kinds are in `filter`.
    pub(super) fn filtered(parent: &Node<'f>, filter: Option<KindSet<Kind>>) -> Self {
        let state = match (
            parent.kind,
            parent.paragraph_position(),
            parent.body_paragraph(),
        ) {
            (Kind::Body, _, _) => State::Body {
                doc: parent.doc,
                paragraphs: None,
            },
            (Kind::Paragraph, Some(position), Some(paragraph)) => {
                let range = paragraph.text_range();
                State::Paragraph {
                    doc: parent.doc,
                    position,
                    anchors: range.start..range.end.saturating_add(1),
                    runs: paragraph.runs(),
                    tabs: paragraph.format().tab_stops(),
                }
            },
            // Header/footer paragraphs expose text only; runs, links and
            // pictures are main-story axes.
            (Kind::Paragraph, _, _) => State::None,
            _ => State::Other(*parent),
        };
        Self {
            state,
            kind: parent.kind,
            view: filter.is_none(),
            filter: filter.unwrap_or(KindSet::ALL),
            phase: 0,
            ordinal: 0,
            raw: 0,
        }
    }

    /// The child kinds walked, in order.
    fn kinds(&self) -> &'static [Kind] {
        match self.kind {
            // Views list a story's paragraphs; the query walker does not
            // descend into header/footer stories.
            Kind::Header | Kind::Footer if self.view => &[Kind::Paragraph],
            // Views list the structural kinds only (record 0328).
            Kind::Document if self.view => &[
                Kind::Body,
                Kind::Header,
                Kind::Footer,
                Kind::Footnote,
                Kind::Endnote,
                Kind::Comments,
                Kind::Styles,
                Kind::Numbering,
            ],
            Kind::Body if self.view => &[Kind::Paragraph, Kind::Table],
            Kind::Paragraph if self.view => &[Kind::Run, Kind::Hyperlink, Kind::Picture],
            kind => SPEC.elements[kind as usize].children,
        }
    }

    /// The next child of `kind` under `parent`, advancing this axis.
    #[allow(
        clippy::too_many_lines,
        reason = "one flat dispatch per native axis keeps document order visible"
    )]
    fn produce(&mut self, kind: Kind) -> Option<Node<'f>> {
        let position = Position::new(self.ordinal);
        // The hot main-story axes read the compact state directly.
        match (&mut self.state, kind) {
            (State::None, _) => return None,
            (State::Body { doc, paragraphs }, Kind::Paragraph) => {
                let doc = *doc;
                let paragraph = paragraphs
                    .get_or_insert_with(|| doc.0.package.body().paragraphs())
                    .next()?;
                return Some(Node::paragraph(doc, position, paragraph));
            },
            (State::Body { doc, .. }, Kind::Table) => {
                let doc = *doc;
                return doc.0.package.tables().get(self.ordinal).map(|table| {
                    Node::body(doc).child(
                        Kind::Table,
                        Item::Table {
                            table: self.ordinal,
                            value: table,
                        },
                        position,
                    )
                });
            },
            (
                State::Paragraph {
                    doc,
                    position: paragraph,
                    runs,
                    ..
                },
                Kind::Run,
            ) => {
                let (doc, paragraph) = (*doc, *paragraph);
                return Some(Node::run(doc, paragraph, position, runs.next()?));
            },
            (
                State::Paragraph {
                    doc,
                    position: paragraph,
                    anchors,
                    ..
                },
                Kind::Hyperlink | Kind::Picture,
            ) => {
                let (doc, paragraph) = (*doc, *paragraph);
                let range = anchors.start..=anchors.end.saturating_sub(1);
                let package = &doc.0.package;
                let item = if kind == Kind::Hyperlink {
                    let index = stories::hyperlinks(doc, range).find(|index| *index >= self.raw)?;
                    self.raw = index.saturating_add(1);
                    Item::Hyperlink(package.fields().get(index)?)
                } else {
                    let index = stories::pictures(doc, range).find(|index| *index >= self.raw)?;
                    self.raw = index.saturating_add(1);
                    Item::Picture {
                        index,
                        value: package.pictures().get(index)?,
                    }
                };
                return Some(Node::paragraph_child(doc, paragraph, kind, item, position));
            },
            (
                State::Paragraph {
                    doc,
                    position: paragraph,
                    tabs,
                    ..
                },
                Kind::Tab,
            ) => {
                let (doc, paragraph) = (*doc, *paragraph);
                let stop = *tabs.get(self.ordinal)?;
                return Some(Node::paragraph_child(
                    doc,
                    paragraph,
                    Kind::Tab,
                    Item::Tab {
                        index: self.ordinal,
                        value: stop,
                    },
                    position,
                ));
            },
            (State::Body { doc, .. }, Kind::TextBox) => {
                let doc = *doc;
                let reference = *super::parts::text_boxes(doc).get(self.ordinal)?;
                return super::nav::text_box_node(doc, &Node::body(doc), reference, self.ordinal);
            },
            (State::Body { doc, .. }, Kind::Shape) => {
                let doc = *doc;
                let index = super::parts::shapes(doc).nth(self.ordinal)?;
                return doc.0.package.shapes().get(index).map(|value| {
                    Node::body(doc).child(Kind::Shape, Item::Shape { value }, position)
                });
            },
            (State::Body { .. } | State::Paragraph { .. }, _) => return None,
            (State::Other(_), _) => {},
        }
        let State::Other(parent) = &self.state else {
            return None;
        };
        let parent: &Node<'f> = parent;
        let doc = parent.doc;
        let package = &doc.0.package;
        match (parent.kind, kind) {
            (Kind::Document, Kind::Body) => (self.ordinal == 0).then(|| Node::body(doc)),
            (Kind::Document, Kind::Header | Kind::Footer) => {
                let story = stories::header_footers(doc, kind == Kind::Header).nth(self.ordinal)?;
                Some(Node::top(
                    doc,
                    kind,
                    Item::Story {
                        section: story.section,
                        story: story.story,
                        value: story.value,
                    },
                    position,
                ))
            },
            (Kind::Document, Kind::Footnote | Kind::Endnote) => {
                let notes = package.notes();
                let footnotes = kind == Kind::Footnote;
                while let Some(note) = notes.get(self.raw) {
                    let index = self.raw;
                    self.raw += 1;
                    if note.is_footnote == footnotes {
                        return Some(Node::top(
                            doc,
                            kind,
                            Item::Note {
                                index,
                                number: self.ordinal.saturating_add(1),
                                value: note,
                            },
                            position,
                        ));
                    }
                }
                None
            },
            (Kind::Document, Kind::Comments) => (self.ordinal == 0
                && !package.annotations().is_empty())
            .then(|| Node::collection(doc, Kind::Comments)),
            (Kind::Document, Kind::Styles) => (self.ordinal == 0 && !package.styles().is_empty())
                .then(|| Node::collection(doc, Kind::Styles)),
            (Kind::Document, Kind::Numbering) => (self.ordinal == 0
                && !(package.lists().is_empty() && package.list_overrides().is_empty()))
            .then(|| Node::collection(doc, Kind::Numbering)),
            (Kind::Table, Kind::Row) => {
                let Item::Table { table, value } = parent.item else {
                    return None;
                };
                value.rows().get(self.ordinal).map(|row| {
                    parent.child(
                        Kind::Row,
                        Item::Row {
                            table,
                            row: self.ordinal,
                            value: row,
                        },
                        position,
                    )
                })
            },
            (Kind::Row, Kind::Cell) => {
                let Item::Row { table, row, value } = parent.item else {
                    return None;
                };
                value.cells().get(self.ordinal).map(|cell| {
                    parent.child(
                        Kind::Cell,
                        Item::Cell {
                            table,
                            row,
                            cell: self.ordinal,
                            value: cell,
                            parent: value,
                        },
                        position,
                    )
                })
            },
            (Kind::Header | Kind::Footer, Kind::Paragraph) => {
                let Item::Story {
                    section,
                    story,
                    value,
                } = parent.item
                else {
                    return None;
                };
                value.paragraphs.get(self.ordinal).map(|paragraph| {
                    parent.child(
                        Kind::Paragraph,
                        Item::StoryParagraph {
                            section,
                            story,
                            index: self.ordinal,
                            value: paragraph,
                        },
                        position,
                    )
                })
            },
            (Kind::Comments, Kind::Comment) => {
                package.annotations().get(self.ordinal).map(|comment| {
                    parent.child(
                        Kind::Comment,
                        Item::Comment {
                            index: self.ordinal,
                            value: comment,
                        },
                        position,
                    )
                })
            },
            (Kind::Styles, Kind::Style) => package
                .styles()
                .get(self.ordinal)
                .map(|style| parent.child(Kind::Style, Item::Style(style), position)),
            (Kind::Numbering, Kind::AbstractNum) => package
                .lists()
                .get(self.ordinal)
                .map(|list| parent.child(Kind::AbstractNum, Item::AbstractNum(list), position)),
            (Kind::Numbering, Kind::Num) => package
                .list_overrides()
                .get(self.ordinal)
                .map(|list| parent.child(Kind::Num, Item::Num(list), position)),
            (Kind::AbstractNum, Kind::Level) => {
                let Item::AbstractNum(list) = parent.item else {
                    return None;
                };
                list.levels
                    .get(self.ordinal)
                    .map(|level| parent.child(Kind::Level, Item::Level(level), position))
            },
            (Kind::Document, Kind::Section) => package.sections().get(self.ordinal).map(|value| {
                Node::top(
                    doc,
                    Kind::Section,
                    Item::Section {
                        index: self.ordinal,
                        value,
                    },
                    position,
                )
            }),
            (Kind::Document, Kind::Bookmark) => package
                .raw()
                .bookmarks()
                .bookmarks()
                .get(self.ordinal)
                .map(|value| {
                    Node::top(
                        doc,
                        Kind::Bookmark,
                        Item::Bookmark {
                            index: self.ordinal,
                            value,
                        },
                        position,
                    )
                }),
            (Kind::Document, kind @ (Kind::Field | Kind::Toc)) => {
                let index = super::parts::fields(doc, kind == Kind::Toc).nth(self.ordinal)?;
                package
                    .fields()
                    .get(index)
                    .map(|value| Node::top(doc, kind, Item::Field { index, value }, position))
            },
            _ => None,
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        while let Some(kind) = self.kinds().get(usize::from(self.phase)).copied() {
            if self.filter.contains(kind)
                && let Some(node) = self.produce(kind)
            {
                self.ordinal = self.ordinal.saturating_add(1);
                return Some(Ok(node));
            }
            self.phase = self.phase.saturating_add(1);
            self.ordinal = 0;
            self.raw = 0;
        }
        None
    }
}

//! Test-only fake formats: one [`Navigate`] adapter over static trees,
//! instantiated with each family vocabulary, so the resolver's tests and the
//! counting-allocator test run every path, error code and suggestion without
//! a real document.
//!
//! A tree node carries its canonical element name and string attributes:
//! `n` is the number `[N]` selects for number-indexed elements (rows, note
//! ids, numbering levels), `name` is the name named children and bare
//! names match, `type` is the token `placeholder[ctrTitle]` matches, and any
//! other attribute is an identity value. Nothing here allocates on the
//! success path.

use std::fmt;
use std::marker::PhantomData;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRange, CellRef};

use super::navigate::{
    Available, ElementKind, GapInfo, KindSet, Lookup, Navigate, Resolved, Select,
};
use super::vocabulary::{self, Vocabulary};
use crate::Result;

/// A static document tree.
#[derive(Debug, PartialEq, Eq)]
pub(crate) struct Tree {
    kind: &'static str,
    attrs: &'static [(&'static str, &'static str)],
    children: &'static [Tree],
}

impl Tree {
    const fn new(
        kind: &'static str,
        attrs: &'static [(&'static str, &'static str)],
        children: &'static [Tree],
    ) -> Self {
        Self {
            kind,
            attrs,
            children,
        }
    }

    const fn leaf(kind: &'static str) -> Self {
        Self::new(kind, &[], &[])
    }

    fn attr(&self, key: &str) -> Option<&'static str> {
        self.attrs
            .iter()
            .find(|(held, _)| held.eq_ignore_ascii_case(key))
            .map(|(_, value)| *value)
    }
}

/// One fake format: its vocabulary, its tree, and the names it treats as
/// capability gaps, unmodeled, absent, or impossible under a parent.
pub(crate) trait Fixture: 'static {
    const VOCABULARY: &'static Vocabulary;
    const FORMAT: FileFormat;
    const TREE: &'static Tree;
    /// Elements whose lookup answers `Lookup::Gap`.
    const GAPS: &'static [&'static str];
    /// Elements `ElementKind::from_name` does not map.
    const UNMODELED: &'static [&'static str];
    /// Elements the format lacks (`Navigate::ABSENT`).
    const ABSENT: &'static [&'static str];
    /// `(parent, child)` pairs that answer `Lookup::NotApplicable`.
    const NOT_UNDER: &'static [(&'static str, &'static str)];
}

/// A fake element kind: an index into the fixture vocabulary's elements.
pub(crate) struct Kind<F>(u8, PhantomData<fn() -> F>);

impl<F> Clone for Kind<F> {
    fn clone(&self) -> Self {
        *self
    }
}

impl<F> Copy for Kind<F> {}

impl<F> PartialEq for Kind<F> {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}

impl<F> Eq for Kind<F> {}

impl<F: Fixture> fmt::Debug for Kind<F> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.name())
    }
}

/// The index of the vocabulary's root element.
const fn root_index(vocabulary: &Vocabulary) -> u8 {
    let mut index = 0;
    while index < vocabulary.elements.len() {
        let name = vocabulary.elements[index].name.as_bytes();
        let root = vocabulary.root.as_bytes();
        if name.len() == root.len() {
            let mut byte = 0;
            while byte < name.len() && name[byte] == root[byte] {
                byte += 1;
            }
            if byte == name.len() {
                return index as u8;
            }
        }
        index += 1;
    }
    panic!("the vocabulary root is not an element")
}

fn listed(list: &[&str], name: &str) -> bool {
    list.iter().any(|entry| entry.eq_ignore_ascii_case(name))
}

impl<F: Fixture> ElementKind for Kind<F> {
    const VOCABULARY: &'static Vocabulary = F::VOCABULARY;
    const ROOT: Self = Self(root_index(F::VOCABULARY), PhantomData);

    fn from_name(name: &str) -> Option<Self> {
        if listed(F::UNMODELED, name) || listed(F::ABSENT, name) {
            return None;
        }
        let index = F::VOCABULARY.index(name)?;
        Some(Self(u8::try_from(index).ok()?, PhantomData))
    }

    fn name(self) -> &'static str {
        F::VOCABULARY.elements[usize::from(self.0)].name
    }

    fn ordinal(self) -> u8 {
        self.0
    }
}

/// A node of a fake document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Node {
    /// An element of the tree.
    Tree(&'static Tree),
    /// A worksheet cell, stored or not.
    Cell(CellRef),
    /// A worksheet range.
    Range(CellRange),
    /// A worksheet row with no record.
    Row(usize),
}

impl Node {
    /// The canonical name of the node's element.
    pub(crate) fn kind(self) -> &'static str {
        match self {
            Self::Tree(tree) => tree.kind,
            Self::Cell(_) => "cell",
            Self::Range(_) => "range",
            Self::Row(_) => "row",
        }
    }

    /// An attribute of a tree node.
    pub(crate) fn attr(self, key: &str) -> Option<&'static str> {
        match self {
            Self::Tree(tree) => tree.attr(key),
            Self::Cell(_) | Self::Range(_) | Self::Row(_) => None,
        }
    }
}

/// The gap every fixture reports.
#[derive(Debug)]
pub(crate) struct Gap;

impl GapInfo for Gap {
    fn follow_up(&self) -> &'static str {
        "L-TEST-R1"
    }
}

static GAP: Gap = Gap;

/// The fake adapter.
pub(crate) struct Fake<F>(PhantomData<F>);

impl<F: Fixture> Fake<F> {
    pub(crate) const fn new() -> Self {
        Self(PhantomData)
    }
}

impl<'f, F: Fixture> Navigate<'f> for Fake<F> {
    type Kind = Kind<F>;
    type Node = Node;
    type Gap = Gap;
    type Children = Children<F>;

    const FORMAT: FileFormat = F::FORMAT;
    const ABSENT: &'static [&'static str] = F::ABSENT;

    fn root(&self) -> Result<Node> {
        Ok(Node::Tree(F::TREE))
    }

    fn child(&self, parent: &Node, step: &Resolved<'_, Kind<F>>) -> Result<Lookup<Node, Gap>> {
        let Node::Tree(parent) = *parent else {
            return Ok(Lookup::NotApplicable);
        };
        let kind = step.kind.name();
        if listed(F::GAPS, kind) {
            return Ok(Lookup::Gap(&GAP));
        }
        if F::NOT_UNDER
            .iter()
            .any(|(above, below)| *above == parent.kind && *below == kind)
        {
            return Ok(Lookup::NotApplicable);
        }
        let mut of_kind = parent.children.iter().filter(|child| child.kind == kind);
        let found = match step.select {
            Select::None => of_kind.next(),
            Select::Ordinal(position) => of_kind.nth(position.get()),
            Select::Last => of_kind.next_back(),
            Select::Number(number) => {
                let hit = of_kind.find(|child| {
                    child.attr("n").and_then(|held| held.parse::<usize>().ok()) == Some(number)
                });
                // A row with no record still resolves (record 0010 §4.7).
                if hit.is_none() && kind == "row" {
                    return Ok(Lookup::Found(Node::Row(number)));
                }
                hit
            },
            Select::Identity(identities) => {
                let mut matching = of_kind.filter(|child| {
                    identities.iter().all(|(identity, value)| {
                        child
                            .attr(identity.key)
                            .is_some_and(|held| identity.case.matches(held, value))
                    })
                });
                let first = matching.next();
                let more = matching.count();
                if more > 0 {
                    return Ok(Lookup::Ambiguous(more + 1));
                }
                first
            },
            Select::Name { name, case } => of_kind.find(|child| {
                child
                    .attr("name")
                    .is_some_and(|held| case.matches(held, name))
            }),
            Select::Token(token) => of_kind.find(|child| {
                child
                    .attr("type")
                    .is_some_and(|held| held.eq_ignore_ascii_case(token))
            }),
            Select::Cell(at) => return Ok(Lookup::Found(Node::Cell(at))),
            Select::Range(range) => return Ok(Lookup::Found(Node::Range(range))),
        };
        Ok(found.map_or(Lookup::Missing, |child| Lookup::Found(Node::Tree(child))))
    }

    fn children(&self, parent: &Node, kinds: KindSet<Kind<F>>) -> Result<Children<F>> {
        let children: &'static [Tree] = match parent {
            Node::Tree(tree) => tree.children,
            Node::Cell(_) | Node::Range(_) | Node::Row(_) => &[],
        };
        Ok(Children {
            children: children.iter(),
            kinds,
        })
    }

    fn available(
        &self,
        parent: &Node,
        step: &Resolved<'_, Kind<F>>,
        out: &mut Available,
    ) -> Result<()> {
        let Node::Tree(parent) = parent else {
            return Ok(());
        };
        if matches!(step.select, Select::Name { .. }) {
            let kind = step.kind.name();
            for child in parent.children.iter().filter(|child| child.kind == kind) {
                if let Some(name) = child.attr("name") {
                    out.name(name);
                }
            }
            return Ok(());
        }
        let mut counts: Vec<(&str, usize)> = Vec::new();
        for child in parent.children {
            match counts.iter_mut().find(|(kind, _)| *kind == child.kind) {
                Some((_, count)) => *count += 1,
                None => counts.push((child.kind, 1)),
            }
        }
        for (kind, count) in counts {
            out.kind(kind, count);
        }
        Ok(())
    }
}

/// The children of a fake tree node whose kinds are in a set.
pub(crate) struct Children<F> {
    children: std::slice::Iter<'static, Tree>,
    kinds: KindSet<Kind<F>>,
}

impl<F: Fixture> Iterator for Children<F> {
    type Item = Result<Node>;

    fn next(&mut self) -> Option<Self::Item> {
        let kinds = self.kinds;
        self.children
            .find(|child| Kind::<F>::from_name(child.kind).is_some_and(|kind| kinds.contains(kind)))
            .map(|child| Ok(Node::Tree(child)))
    }
}

const fn t(
    kind: &'static str,
    attrs: &'static [(&'static str, &'static str)],
    children: &'static [Tree],
) -> Tree {
    Tree::new(kind, attrs, children)
}

const fn l(kind: &'static str) -> Tree {
    Tree::leaf(kind)
}

const RUN: Tree = l("r");

/// A Word-processing document.
pub(crate) enum Word {}

impl Fixture for Word {
    const VOCABULARY: &'static Vocabulary = &vocabulary::WORD;
    const FORMAT: FileFormat = FileFormat::Docx;
    const TREE: &'static Tree = &t(
        "document",
        &[],
        &[
            t(
                "body",
                &[],
                &[
                    t(
                        "p",
                        &[("paraId", "1A2B3C4D"), ("textId", "77AA0001")],
                        &[RUN, RUN, l("hyperlink")],
                    ),
                    t("p", &[("paraId", "00000002")], &[RUN]),
                    t(
                        "tbl",
                        &[],
                        &[
                            t(
                                "tr",
                                &[],
                                &[
                                    t("tc", &[], &[t("p", &[], &[RUN])]),
                                    t("tc", &[], &[l("p")]),
                                ],
                            ),
                            t(
                                "tr",
                                &[],
                                &[
                                    t("tc", &[], &[t("p", &[], &[RUN, RUN])]),
                                    t("tc", &[], &[l("p")]),
                                ],
                            ),
                        ],
                    ),
                    t("p", &[], &[RUN, RUN, RUN]),
                    t("sdt", &[("sdtId", "7")], &[t("p", &[], &[RUN])]),
                    t("p", &[("paraId", "0000000A")], &[RUN]),
                ],
            ),
            t("header", &[("type", "default")], &[t("p", &[], &[RUN])]),
            t("header", &[("type", "first")], &[l("p")]),
            t("footer", &[("type", "default")], &[l("p")]),
            t(
                "footnotes",
                &[],
                &[
                    t(
                        "footnote",
                        &[("footnoteId", "1"), ("n", "1")],
                        &[t("p", &[], &[RUN])],
                    ),
                    t(
                        "footnote",
                        &[("footnoteId", "2"), ("n", "2")],
                        &[t("p", &[], &[RUN])],
                    ),
                ],
            ),
            // `/footnote[N]` reads note ids from the root, as OfficeCLI does.
            t(
                "footnote",
                &[("footnoteId", "1"), ("n", "1")],
                &[t("p", &[], &[RUN])],
            ),
            t(
                "footnote",
                &[("footnoteId", "2"), ("n", "2")],
                &[t("p", &[], &[RUN])],
            ),
            t(
                "comments",
                &[],
                &[
                    t("comment", &[("commentId", "0")], &[l("p")]),
                    t("comment", &[("commentId", "1")], &[l("p")]),
                ],
            ),
            t(
                "styles",
                &[],
                &[
                    t("style", &[("name", "Heading1")], &[l("tab")]),
                    t("style", &[("name", "Normal")], &[]),
                ],
            ),
            t(
                "numbering",
                &[],
                &[
                    t(
                        "abstractNum",
                        &[("id", "1")],
                        &[
                            t("level", &[("ilvl", "0"), ("n", "0")], &[]),
                            t("level", &[("ilvl", "1"), ("n", "1")], &[]),
                        ],
                    ),
                    t("num", &[("id", "1")], &[]),
                ],
            ),
            t("bookmark", &[("name", "Intro")], &[]),
            t("bookmark", &[("name", "a/b]")], &[]),
            t("formfield", &[("name", "agree")], &[]),
            l("section"),
        ],
    );
    const GAPS: &'static [&'static str] = &["chart"];
    const UNMODELED: &'static [&'static str] = &["toc"];
    const ABSENT: &'static [&'static str] = &["customXml"];
    const NOT_UNDER: &'static [(&'static str, &'static str)] = &[("body", "tr"), ("body", "tc")];
}

/// A presentation.
pub(crate) enum Presentation {}

impl Fixture for Presentation {
    const VOCABULARY: &'static Vocabulary = &vocabulary::PRESENTATION;
    const FORMAT: FileFormat = FileFormat::Pptx;
    const TREE: &'static Tree = &t(
        "presentation",
        &[],
        &[
            t(
                "slide",
                &[],
                &[
                    t("shape", &[("id", "2"), ("name", "Title 1")], &[]),
                    t("shape", &[("id", "3"), ("name", "Content")], &[]),
                    t(
                        "placeholder",
                        &[
                            ("type", "body"),
                            ("idx", "1"),
                            ("id", "4"),
                            ("name", "Body"),
                        ],
                        &[],
                    ),
                    t(
                        "placeholder",
                        &[
                            ("type", "pic"),
                            ("idx", "1"),
                            ("id", "5"),
                            ("name", "Picture"),
                        ],
                        &[],
                    ),
                    t(
                        "placeholder",
                        &[
                            ("type", "ctrTitle"),
                            ("idx", "0"),
                            ("id", "6"),
                            ("name", "Title"),
                        ],
                        &[],
                    ),
                    t(
                        "table",
                        &[("id", "7"), ("name", "Grid")],
                        &[t("tr", &[], &[l("tc"), l("tc")])],
                    ),
                    l("notes"),
                ],
            ),
            t(
                "slide",
                &[],
                &[t("shape", &[("id", "2"), ("name", "Ärger")], &[])],
            ),
            t(
                "slideMaster",
                &[],
                &[t(
                    "slideLayout",
                    &[],
                    &[t("shape", &[("id", "1"), ("name", "Layout Title")], &[])],
                )],
            ),
        ],
    );
    const GAPS: &'static [&'static str] = &["chart"];
    const UNMODELED: &'static [&'static str] = &["zoom"];
    const ABSENT: &'static [&'static str] = &[];
    const NOT_UNDER: &'static [(&'static str, &'static str)] = &[("presentation", "shape")];
}

/// A workbook.
pub(crate) enum Workbook {}

impl Fixture for Workbook {
    const VOCABULARY: &'static Vocabulary = &vocabulary::WORKBOOK;
    const FORMAT: FileFormat = FileFormat::Xlsx;
    const TREE: &'static Tree = &t(
        "workbook",
        &[],
        &[
            t(
                "sheet",
                &[("name", "Sheet1")],
                &[
                    t("row", &[("n", "1")], &[]),
                    t("row", &[("n", "3")], &[]),
                    t("col", &[("n", "2")], &[]),
                    t("table", &[("name", "Sales")], &[]),
                    l("comment"),
                ],
            ),
            t(
                "sheet",
                &[("name", "Data Q1")],
                &[t("row", &[("n", "2")], &[])],
            ),
            t("sheet", &[("name", "Chart")], &[]),
            t("sheet", &[("name", "Straße")], &[]),
            t("namedrange", &[("name", "Total")], &[]),
            t("namedrange", &[("name", "Rates")], &[]),
        ],
    );
    const GAPS: &'static [&'static str] = &["pivottable"];
    const UNMODELED: &'static [&'static str] = &["slicer"];
    const ABSENT: &'static [&'static str] = &[];
    const NOT_UNDER: &'static [(&'static str, &'static str)] = &[("workbook", "row")];
}

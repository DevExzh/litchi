//! Path and query navigation for the DOC root/body/paragraph vocabulary.

use std::borrow::Cow;
use std::fmt;

use litchi_core::{FileFormat, Position};

use super::Doc;
use super::children::Children;
use super::node::Node;
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

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = crate::node::spec::GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Doc;
    const ABSENT: &'static [&'static str] = &[
        "comments",
        "endnotes",
        "footnotes",
        "header",
        "footer",
        "styles",
        "numbering",
    ];

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
        match (parent.kind, step.kind) {
            (Kind::Document, Kind::Body) => match step.select {
                Select::None | Select::Last => Ok(Lookup::Found(Node::body(self.doc))),
                Select::Ordinal(position) if position.get() == 0 => {
                    Ok(Lookup::Found(Node::body(self.doc)))
                },
                Select::Ordinal(_) => Ok(Lookup::Missing),
                _ => Ok(Lookup::NotApplicable),
            },
            (Kind::Body, Kind::Paragraph) => {
                let position = match step.select {
                    Select::None => self.doc.body_paragraph_at(Position::new(0))?.map(|_| 0),
                    Select::Ordinal(position) => self
                        .doc
                        .body_paragraph_at(position)?
                        .map(|_| position.get()),
                    Select::Last => self.doc.body_paragraph_count()?.checked_sub(1),
                    _ => None,
                };
                Ok(position.map_or(Lookup::Missing, |position| {
                    Lookup::Found(Node::paragraph(self.doc, Position::new(position)))
                }))
            },
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children> {
        let wants = match parent.kind {
            Kind::Document => kinds.contains(Kind::Body),
            Kind::Body => kinds.contains(Kind::Paragraph),
            Kind::Paragraph => false,
        };
        Ok(if wants {
            Children::of(parent)
        } else {
            Children {
                doc: self.doc,
                next: 0,
                body: false,
                done: true,
                paragraphs: None,
                error: None,
            }
        })
    }

    fn available(
        &self,
        parent: &Self::Node,
        _: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        match parent.kind {
            Kind::Document => out.kind("body", 1),
            Kind::Body => out.kind("p", self.doc.body_paragraph_count()?),
            Kind::Paragraph => {},
        }
        Ok(())
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
        Ok(node.paragraph_text()?.map(Cow::Borrowed))
    }

    fn format<'n>(&'n self, _node: &'n Node<'f>) -> Result<Format<'n>> {
        Ok(Format::empty())
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
            Kind::Paragraph => node
                .path()
                .steps()
                .last()
                .copied()
                .unwrap_or_else(|| PathStep::at("p", ordinal)),
            Kind::Document => PathStep::new("document"),
        };
        fmt::Display::fmt(&step, out)
    }
}

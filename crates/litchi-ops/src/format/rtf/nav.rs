//! Path and query navigation for the RTF root/body/paragraph/run vocabulary.

use std::borrow::Cow;
use std::fmt;

use litchi_core::{FileFormat, Position};

use super::Rtf;
use super::children::Children;
use super::node::Node;
use super::spec::{Kind, SPEC};
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

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = crate::node::spec::GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Rtf;

    fn check_source(&self) -> Result<()> {
        Ok(())
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
                    Select::None => Some(0),
                    Select::Ordinal(position) => Some(position.get()),
                    Select::Last => self.doc.0.package.paragraph_count().checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                let paragraph = position.and_then(|position| {
                    self.doc
                        .0
                        .package
                        .body()
                        .paragraphs()
                        .nth(position)
                        .map(|paragraph| {
                            Node::paragraph(self.doc, Position::new(position), paragraph)
                        })
                });
                Ok(paragraph.map_or(Lookup::Missing, Lookup::Found))
            },
            (Kind::Paragraph, Kind::Run) => {
                let Some(paragraph) = parent.paragraph else {
                    return Ok(Lookup::Missing);
                };
                let Some(paragraph_position) = parent.paragraph_position else {
                    return Ok(Lookup::Missing);
                };
                let position = match step.select {
                    Select::None => Some(0),
                    Select::Ordinal(position) => Some(position.get()),
                    Select::Last => paragraph.runs().count().checked_sub(1),
                    _ => return Ok(Lookup::NotApplicable),
                };
                let run = position.and_then(|position| {
                    paragraph.runs().nth(position).map(|run| {
                        Node::run(
                            self.doc,
                            paragraph_position,
                            Position::new(position),
                            paragraph,
                            run,
                        )
                    })
                });
                Ok(run.map_or(Lookup::Missing, Lookup::Found))
            },
            _ => Ok(Lookup::NotApplicable),
        }
    }

    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children> {
        let wants = match parent.kind {
            Kind::Document => kinds.contains(Kind::Body),
            Kind::Body => kinds.contains(Kind::Paragraph),
            Kind::Paragraph => kinds.contains(Kind::Run),
            Kind::Run => false,
        };
        Ok(if wants {
            Children::of(parent)
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
        match parent.kind {
            Kind::Document => out.kind("body", 1),
            Kind::Body => out.kind("p", self.doc.0.package.paragraph_count()),
            Kind::Paragraph => {
                if let Some(paragraph) = parent.paragraph {
                    out.kind("r", paragraph.runs().count());
                }
            },
            Kind::Run => {},
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
        node.text()
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        crate::node::Project::format(node)
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
            Kind::Paragraph | Kind::Run => {
                node.path().steps().last().copied().unwrap_or_else(|| {
                    if kind == Kind::Run {
                        PathStep::at("r", ordinal)
                    } else {
                        PathStep::at("p", ordinal)
                    }
                })
            },
            Kind::Document => PathStep::new("document"),
        };
        fmt::Display::fmt(&step, out)
    }
}

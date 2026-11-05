//! Path and query navigation over the ODT body index.

use std::borrow::Cow;
use std::fmt;

use litchi_core::{FileFormat, Position};
use litchi_odt::body::{Attribute, NodeId};

use super::Odt;
use super::children::Children;
use super::node::{Address, Node, Step, Target};
use super::spec::{ABSENT, Kind, SPEC};
use crate::error::Result;
use crate::node::spec::FormatSpec;
use crate::node::{Format, PathStep};
use crate::path::navigate::{Available, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::Inspect;

/// The ODT adapter used by path resolution, queries and generic views.
#[derive(Clone, Copy)]
pub(crate) struct Nav<'f> {
    doc: &'f Odt,
}

impl<'f> Nav<'f> {
    pub(crate) const fn new(doc: &'f Odt) -> Self {
        Self { doc }
    }
}

/// Whether `kind` may sit under `parent` in the spec table.
fn admitted(parent: Kind, kind: Kind) -> bool {
    SPEC.elements[parent as usize].children.contains(&kind)
}

impl<'f> Navigate<'f> for Nav<'f> {
    type Kind = Kind;
    type Node = Node<'f>;
    type Gap = crate::node::spec::GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = FileFormat::Odt;
    const ABSENT: &'static [&'static str] = ABSENT;

    fn check_source(&self) -> Result<()> {
        self.doc.check_source()
    }

    fn root(&self) -> Result<Self::Node> {
        Ok(Node::new(
            self.doc,
            Kind::Document,
            Target::Synthetic,
            Address::root(),
        ))
    }

    fn child(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
    ) -> Result<Lookup<Self::Node, Self::Gap>> {
        if !admitted(parent.kind, step.kind) {
            return Ok(Lookup::NotApplicable);
        }
        match (parent.kind, step.kind) {
            (Kind::Document, kind @ (Kind::Body | Kind::Styles | Kind::Metadata)) => {
                match step.select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                let address = parent.address().with(Step::new(kind, 0));
                let target = if kind == Kind::Body {
                    Target::Body(self.doc.body()?, NodeId::BODY)
                } else {
                    Target::Synthetic
                };
                Ok(Lookup::Found(Node::new(self.doc, kind, target, address)))
            },
            (Kind::Styles, Kind::Style) => {
                let body = self.doc.body()?;
                let found = match step.select {
                    Select::Name { name, case } => {
                        body.common_styles().enumerate().find(|(_, style)| {
                            style.name().is_some_and(|held| case.matches(&held, name))
                        })
                    },
                    Select::Ordinal(position) => {
                        body.common_styles().enumerate().nth(position.get())
                    },
                    Select::None => body.common_styles().enumerate().next(),
                    Select::Last => body.common_styles().enumerate().last(),
                    _ => return Ok(Lookup::NotApplicable),
                };
                Ok(found.map_or(Lookup::Missing, |(ordinal, style)| {
                    Lookup::Found(Node::style(
                        self.doc,
                        style,
                        u32::try_from(ordinal).unwrap_or(u32::MAX),
                    ))
                }))
            },
            (_, kind) if kind.is_part() => self.part(parent, kind, step.select),
            (_, kind) => {
                if let (Target::Body(body, _), Select::None | Select::Ordinal(_) | Select::Last) =
                    (parent.target, step.select)
                {
                    let ordinal = match step.select {
                        Select::Ordinal(position) => Some(position.get()),
                        Select::None => Some(0),
                        _ => None,
                    };
                    let found = match ordinal {
                        Some(ordinal) => self.doc.nth_child(parent, kind, ordinal)?,
                        None => self.doc.last_child(parent, kind)?,
                    };
                    return Ok(found.map_or(Lookup::Missing, |(id, ordinal)| {
                        let address = parent
                            .address()
                            .with(Step::new(kind, u32::try_from(ordinal).unwrap_or(u32::MAX)));
                        Lookup::Found(Node::new(self.doc, kind, Target::Body(body, id), address))
                    }));
                }
                let children = Children::of(parent, KindSet::of(kind))?;
                let found = match step.select {
                    Select::None => children.take(1).last(),
                    Select::Ordinal(position) => children.skip(position.get()).take(1).last(),
                    Select::Last => children.last(),
                    Select::Identity(identities) => {
                        let mut matches = children.filter(|node| {
                            node.as_ref().is_ok_and(|node| {
                                identities.iter().all(|(identity, value)| {
                                    identity.key.eq_ignore_ascii_case("name")
                                        && node.leaf().is_some_and(|leaf| {
                                            leaf.attribute(Attribute::Name).is_some_and(|held| {
                                                identity.case.matches(&held, value)
                                            })
                                        })
                                })
                            })
                        });
                        let first = matches.next();
                        if first.is_some() && matches.next().is_some() {
                            return Ok(Lookup::Ambiguous(2));
                        }
                        first
                    },
                    _ => return Ok(Lookup::NotApplicable),
                };
                match found {
                    Some(node) => Ok(Lookup::Found(node?)),
                    None => Ok(Lookup::Missing),
                }
            },
        }
    }

    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children> {
        Children::of(parent, kinds)
    }

    fn available(
        &self,
        parent: &Self::Node,
        _: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        if parent.kind == Kind::Styles {
            for style in self.doc.body()?.common_styles() {
                if let Some(name) = style.name() {
                    out.name(&name);
                }
            }
            return Ok(());
        }
        let mut counts = [0usize; Kind::COUNT];
        for child in Children::of(parent, KindSet::ALL)? {
            counts[child?.kind as usize] += 1;
        }
        for element in SPEC.elements {
            let count = counts[element.kind as usize];
            if count > 0 {
                use crate::path::navigate::ElementKind;
                out.kind(element.kind.name(), count);
            }
        }
        Ok(())
    }
}

/// The index `select` picks among `count` children: `None` when the select
/// does not apply, `Some(None)` when it picks nothing.
fn pick(select: Select<'_>, count: usize) -> Option<Option<usize>> {
    Some(match select {
        Select::None => (count > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < count).then_some(position.get()),
        Select::Last => count.checked_sub(1),
        Select::Number(number) => number.checked_sub(1).filter(|index| *index < count),
        _ => return None,
    })
}

/// The single value `select` names for identity `key`: `None` when it
/// names no identity, `Some(None)` when it names another key or the
/// stacked values disagree.
fn identity<'s>(select: Select<'s>, key: &str) -> Option<Option<&'s str>> {
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

impl<'f> Nav<'f> {
    /// Resolve one step to a part of record 0328 under `parent`.
    fn part(
        &self,
        parent: &Node<'f>,
        kind: Kind,
        select: Select<'_>,
    ) -> Result<Lookup<Node<'f>, crate::node::spec::GapSpec>> {
        use super::node::Part;
        use super::parts;
        let odt = self.doc;
        let found = |index: Option<usize>, part: &dyn Fn(usize) -> Part| {
            index.map_or(Lookup::Missing, |index| {
                Lookup::Found(parts::node(parent, kind, part(index), index))
            })
        };
        let count = |value: usize| u32::try_from(value).unwrap_or(u32::MAX);
        Ok(match (parent.kind, parent.target, kind) {
            (Kind::Document, _, Kind::Header | Kind::Footer) => {
                let regions = parts::regions(odt, kind == Kind::Header)?;
                let index = match identity(select, "type") {
                    Some(Some(wanted)) => {
                        let mut index = None;
                        for (ordinal, region) in regions.iter().enumerate() {
                            if parts::region_type_of(odt, *region)?
                                .is_some_and(|held| held.eq_ignore_ascii_case(wanted))
                            {
                                index = Some(ordinal);
                                break;
                            }
                        }
                        index
                    },
                    Some(None) => return Ok(Lookup::NotApplicable),
                    None => match pick(select, regions.len()) {
                        Some(index) => index,
                        None => return Ok(Lookup::NotApplicable),
                    },
                };
                found(index, &|index| Part::Region(regions[index]))
            },
            (Kind::Document, _, Kind::Numbering) => {
                match select {
                    Select::None | Select::Last => {},
                    Select::Ordinal(position) if position.get() == 0 => {},
                    Select::Ordinal(_) => return Ok(Lookup::Missing),
                    _ => return Ok(Lookup::NotApplicable),
                }
                found((parts::list_count(odt)? > 0).then_some(0), &|_| {
                    Part::Numbering
                })
            },
            (Kind::Numbering, _, Kind::AbstractNum | Kind::Num) => {
                let lists = parts::list_count(odt)?;
                let index = match identity(select, "id") {
                    Some(Some(wanted)) => wanted
                        .trim()
                        .parse::<usize>()
                        .ok()
                        .filter(|index| *index < lists),
                    Some(None) => return Ok(Lookup::NotApplicable),
                    None => match pick(select, lists) {
                        Some(index) => index,
                        None => return Ok(Lookup::NotApplicable),
                    },
                };
                found(index, &|index| Part::List(count(index)))
            },
            (Kind::AbstractNum, Target::Part(Part::List(list)), Kind::Level) => {
                let levels = parts::level_count(odt, list)?;
                let index = match identity(select, "ilvl") {
                    Some(Some(wanted)) => match wanted.trim().parse::<u16>() {
                        Ok(ilvl) => parts::level_by_ilvl(odt, list, ilvl)?,
                        Err(_) => return Ok(Lookup::NotApplicable),
                    },
                    Some(None) => return Ok(Lookup::NotApplicable),
                    None => match pick(select, levels) {
                        Some(index) => index,
                        None => return Ok(Lookup::NotApplicable),
                    },
                };
                found(index, &|index| Part::Level(list, count(index)))
            },
            (Kind::Document, _, Kind::Field | Kind::Toc) => {
                let total = if kind == Kind::Field {
                    parts::field_count(odt)?
                } else {
                    parts::toc_count(odt)?
                };
                let Some(index) = pick(select, total) else {
                    return Ok(Lookup::NotApplicable);
                };
                if kind == Kind::Field {
                    found(index, &|index| Part::Field(count(index)))
                } else {
                    found(index, &|index| Part::Toc(count(index)))
                }
            },
            (Kind::Body, _, Kind::Shape | Kind::Equation) => {
                let total = if kind == Kind::Shape {
                    parts::shape_count(odt)?
                } else {
                    parts::equation_count(odt)?
                };
                let Some(index) = pick(select, total) else {
                    return Ok(Lookup::NotApplicable);
                };
                if kind == Kind::Shape {
                    found(index, &|index| Part::Shape(count(index)))
                } else {
                    found(index, &|index| Part::Equation(count(index)))
                }
            },
            (Kind::Paragraph | Kind::Heading, Target::Body(_, id), Kind::Tab) => {
                let Some(index) = pick(select, parts::tab_count(odt, id)?) else {
                    return Ok(Lookup::NotApplicable);
                };
                found(index, &|index| Part::Tab(id, count(index)))
            },
            (Kind::Table, Target::Body(_, id), Kind::Column) => {
                let Some(index) = pick(select, parts::column_count(odt, id)?) else {
                    return Ok(Lookup::NotApplicable);
                };
                found(index, &|index| Part::Column(id, count(index)))
            },
            _ => Lookup::NotApplicable,
        })
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

    fn style<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        crate::node::Project::style(node)
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        crate::node::Project::format(node)
    }

    fn childless(&self, node: &Node<'f>, kinds: KindSet<Kind>) -> bool {
        // A body element none of whose child elements projects to one of
        // `kinds` lists no such children (record 0280): a paragraph of runs
        // when the walk looks for paragraphs. Every other node is entered.
        if super::parts::has_part_children(node.kind, kinds) {
            return false;
        }
        match node.target {
            Target::Body(body, id) => body.node(id).is_some_and(|node| {
                node.children()
                    .all(|child| Kind::of(child.kind()).is_none_or(|kind| !kinds.contains(kind)))
            }),
            _ => false,
        }
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        use crate::path::navigate::ElementKind;
        let step = match kind {
            Kind::Body | Kind::Styles | Kind::Metadata | Kind::Numbering => {
                PathStep::new(kind.name())
            },
            Kind::Style => match node.style_name() {
                Some(name) => return fmt::Display::fmt(&PathStep::new(name), out),
                None => PathStep::at(kind.name(), ordinal),
            },
            Kind::Document => PathStep::new("document"),
            _ => node
                .address()
                .steps()
                .last()
                .map_or(PathStep::at(kind.name(), ordinal), |step| {
                    PathStep::at(kind.name(), Position::new(step.ordinal as usize))
                }),
        };
        fmt::Display::fmt(&step, out)
    }
}

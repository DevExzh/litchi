//! Path and query navigation over an iWork object tree, generic over the
//! format's element kinds.

use std::borrow::Cow;
use std::fmt;
use std::marker::PhantomData;

use litchi_core::{FileFormat, Position};

use super::Iwork;
use super::children::Children;
use super::node::Node;
use crate::error::Result;
use crate::node::spec::{FormatSpec, GapSpec};
use crate::node::{Format, PathStep, Source, Value};
use crate::path::navigate::{ElementKind, KindSet, Lookup, Navigate, Resolved, Select};
use crate::query::inspect::Inspect;

/// One iWork format's element kinds and spec table.
pub(crate) trait Flavor: 'static {
    /// The format's element kinds.
    type Kind: ElementKind;
    /// The format the flavor reads.
    const FORMAT: FileFormat;
    /// The spec table, each element at its kind's ordinal.
    const SPEC: &'static FormatSpec<Self::Kind>;
    /// The containers `*` lists the children of.
    const TOP: &'static [Self::Kind];
    /// Vocabulary names the format lacks as a concept.
    const ABSENT: &'static [&'static str] = &[];
    /// Kinds that occur once under their parent, whose path step has no
    /// index.
    const SINGLETONS: &'static [Self::Kind] = &[];
}

/// The kind of ordinal `ordinal` in `F`'s spec table.
pub(crate) fn kind<F: Flavor>(ordinal: u8) -> F::Kind {
    F::SPEC.elements[usize::from(ordinal)].kind
}

/// The generic adapter over one handle's tree.
pub(crate) struct Nav<'f, F> {
    file: &'f Iwork,
    flavor: PhantomData<fn() -> F>,
}

impl<F> Clone for Nav<'_, F> {
    fn clone(&self) -> Self {
        *self
    }
}

impl<F> Copy for Nav<'_, F> {}

impl<'f, F: Flavor> Nav<'f, F> {
    pub(crate) const fn new(file: &'f Iwork) -> Self {
        Self {
            file,
            flavor: PhantomData,
        }
    }
}

impl<'f, F: Flavor> Navigate<'f> for Nav<'f, F> {
    type Kind = F::Kind;
    type Node = Node<'f>;
    type Gap = GapSpec;
    type Children = Children<'f>;

    const FORMAT: FileFormat = F::FORMAT;
    const ABSENT: &'static [&'static str] = F::ABSENT;

    fn root(&self) -> Result<Node<'f>> {
        Ok(Node::new(self.file, super::At::ROOT))
    }

    fn child(
        &self,
        parent: &Node<'f>,
        step: &Resolved<'_, F::Kind>,
    ) -> Result<Lookup<Node<'f>, GapSpec>> {
        let kind = step.kind;
        if let Some(gap) = F::SPEC.elements[usize::from(kind.ordinal())].gap {
            return Ok(Lookup::Gap(gap));
        }
        let parent_kind = kind_of::<F>(parent);
        if !F::SPEC.elements[usize::from(parent_kind.ordinal())]
            .children
            .contains(&kind)
        {
            return Ok(Lookup::NotApplicable);
        }
        let file = self.file;
        let kids = file.kids(parent.at());
        let ordinal = kind.ordinal();
        let mut of_kind = kids.iter().filter(|child| file.kind_at(*child) == ordinal);
        let found = match step.select {
            Select::None => of_kind.next(),
            Select::Ordinal(position) => {
                // Children of one kind are usually contiguous: try the slot
                // `position` past the first of them before counting.
                let candidate = (0..kids.len())
                    .find(|slot| {
                        kids.get(*slot)
                            .is_some_and(|child| file.kind_at(child) == ordinal)
                    })
                    .and_then(|first| kids.get(first + position.get()))
                    .filter(|child| {
                        file.kind_at(*child) == ordinal
                            && usize::try_from(file.ordinal_at(*child))
                                .is_ok_and(|held| held == position.get())
                    });
                candidate.or_else(|| of_kind.nth(position.get()))
            },
            Select::Last => of_kind.next_back(),
            Select::Name { name, case } => of_kind.find(|child| {
                file.item(*child)
                    .label()
                    .is_some_and(|held| case.matches(held, name))
            }),
            Select::Token(token) => of_kind.find(|child| {
                file.item(*child)
                    .label()
                    .is_some_and(|held| held.eq_ignore_ascii_case(token))
            }),
            Select::Cell(cell) => file
                .cell_child(kids, ordinal, cell.row(), cell.column())
                .unwrap_or_else(|| {
                    of_kind.find(|child| {
                        let [_, _, row, column] = file.item(*child).leaf();
                        row == cell.row() && column == cell.column()
                    })
                }),
            Select::Identity(identities) => of_kind.find(|child| {
                identities.iter().all(|(identity, value)| {
                    file.item(*child)
                        .value(identity.key)
                        .is_some_and(|held| identity.case.matches(&held.to_string(), value))
                })
            }),
            _ => return Ok(Lookup::NotApplicable),
        };
        Ok(found.map_or(Lookup::Missing, |at| {
            Lookup::Found(Node::new(self.file, at))
        }))
    }

    fn children(&self, parent: &Node<'f>, kinds: KindSet<F::Kind>) -> Result<Children<'f>> {
        let mut mask = 0_u128;
        for element in F::SPEC.elements {
            if kinds.contains(element.kind) {
                mask |= 1_u128 << (element.kind.ordinal() & 127);
            }
        }
        Ok(Children::filtered(self.file, parent.at(), mask))
    }
}

/// The kind of `node` in `F`.
fn kind_of<F: Flavor>(node: &Node<'_>) -> F::Kind {
    kind::<F>(node.kind_ordinal())
}

impl<'f, F: Flavor> Inspect<'f> for Nav<'f, F> {
    const SPEC: &'static FormatSpec<F::Kind> = F::SPEC;
    const TOP: &'static [F::Kind] = F::TOP;
    const INDIVIDUAL_PROPERTIES: bool = true;

    fn kind(&self, node: &Node<'f>) -> F::Kind {
        kind_of::<F>(node)
    }

    fn type_name(&self, node: &Node<'f>) -> &'static str {
        node.item().type_name()
    }

    fn text<'n>(&'n self, node: &'n Node<'f>) -> Result<Option<Cow<'n, str>>> {
        Ok(node.item().text().map(Cow::Borrowed))
    }

    fn format<'n>(&'n self, node: &'n Node<'f>) -> Result<Format<'n>> {
        let item = node.item();
        Ok(Format::new(item.props(), Source::Tree(item.values())))
    }

    fn prop<'n>(&'n self, node: &'n Node<'f>, index: usize) -> Result<Option<Value<'n>>> {
        let Some(property) = node.item().props().get(index) else {
            return Ok(None);
        };
        node.value(property.name)
            .map(|value| value.map(Value::borrowed))
    }

    fn write_step(
        &self,
        node: &Node<'f>,
        kind: F::Kind,
        ordinal: Position,
        out: &mut fmt::Formatter<'_>,
    ) -> fmt::Result {
        // The tree recorded each item's canonical step (a cell's is its A1
        // address); fall back to the positional step for the root.
        if let Some(step) = node.item().path().last() {
            return fmt::Display::fmt(step, out);
        }
        if F::SINGLETONS.contains(&kind) {
            return fmt::Display::fmt(&PathStep::new(kind.name()), out);
        }
        fmt::Display::fmt(&PathStep::at(kind.name(), ordinal), out)
    }
}

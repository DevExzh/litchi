//! The seams between the resolver and the format adapters (record 0010
//! §4.1): the element kinds a format names, the [`Navigate`] trait each
//! adapter implements, the classified step it is asked for, and what it
//! answers.
//!
//! # Boundary with the node view (M1 step R1)
//!
//! The resolver is generic over the adapter's node type and places no bound
//! on it. R1's per-format `Node<'f>` variant is [`Navigate::Node`]; its
//! canonical `NodePath` is rendered from the node by R1, so the resolver needs
//! no address trait: its error messages quote the resolved prefix of the
//! path it was given. R1's `GapSpec` meets [`GapInfo`] by returning its
//! follow-up, and R1's `FormatSpec` element names are the canonical names of
//! the family [`Vocabulary`], which is how a format adapter (M1 step R4)
//! implements [`ElementKind`].

use std::fmt::{self, Write as _};
use std::marker::PhantomData;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRange, CellRef, Position, Predicate};

use super::vocabulary::{Case, Identity, Vocabulary};
use crate::Result;

/// The element kinds of one format.
///
/// A format adapter implements this on its own `Copy` kind enum. The kinds
/// are a subset of the family vocabulary's canonical names: [`Self::name`]
/// returns one, and [`Self::from_name`] maps one back, returning `None` for a
/// concept the adapter does not model.
pub(crate) trait ElementKind: Copy + Eq + fmt::Debug + 'static {
    /// The vocabulary of the format's family.
    const VOCABULARY: &'static Vocabulary;
    /// The kind of the document root (`Vocabulary::root`).
    const ROOT: Self;

    /// The kind with the canonical name `name`, if the format models it.
    fn from_name(name: &str) -> Option<Self>;

    /// The canonical name: an element of [`Self::VOCABULARY`].
    fn name(self) -> &'static str;

    /// A dense index below 128, for [`KindSet`].
    fn ordinal(self) -> u8;
}

/// Why an element the format has cannot be read yet: R1's `GapSpec`, seen
/// from the resolver.
pub(crate) trait GapInfo: 'static {
    /// The leaf request or milestone that closes the gap (`L-PPTX-R7`).
    fn follow_up(&self) -> &'static str;
}

/// A format adapter's read access, statically dispatched.
///
/// The resolver calls [`Navigate::root`] once and [`Navigate::child`] once per
/// step; only the leaf calls that step's axis needs run. On a miss, `get`
/// calls [`Navigate::available`] to describe what exists; `find` never does,
/// so a miss costs nothing beyond the lookup.
pub(crate) trait Navigate<'f> {
    /// The format's element kinds.
    type Kind: ElementKind;
    /// The format's typed node.
    type Node;
    /// The format's capability-gap descriptor (R1's `GapSpec`).
    type Gap: GapInfo;

    /// Children borrow the retained document, not this adapter value.
    /// Reference adapters may borrow their owner through that reference.
    type Children: Iterator<Item = Result<Self::Node>>;

    /// The format reported in errors.
    const FORMAT: FileFormat;
    /// Canonical names of the family vocabulary that this format lacks as a
    /// concept (`sdt` in DOC); a step naming one is `unsupported_path`. Any
    /// other name the format does not model is `unsupported_capability`.
    const ABSENT: &'static [&'static str] = &[];

    /// Validate retained source revisions around each public read.
    fn check_source(&self) -> Result<()> {
        Ok(())
    }

    /// The handle's key for [`super::fence::Scope`]: while a bracket of
    /// [`Self::check_source`] calls holds a scope with this key, the
    /// adapter's nested reads skip their own checks. 0 (the default) for an
    /// adapter whose reads do not check.
    fn fence_key(&self) -> usize {
        0
    }

    /// The document root (`/`).
    fn root(&self) -> Result<Self::Node>;

    /// The child of `parent` that `step` selects.
    fn child(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
    ) -> Result<Lookup<Self::Node, Self::Gap>>;

    /// The children of `parent` whose kinds are in `kinds`, in document
    /// order (for the query walker and the node view). The iterator is lazy:
    /// a walk that stops early decodes no later sibling.
    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children>;

    /// Describes what exists under `parent` for a `path_not_found` or
    /// ambiguity suggestion: counts per kind, or the names of named children.
    /// Called on the error path only; the default describes nothing.
    fn available(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        let _ = (parent, step, out);
        Ok(())
    }
}

/// A shared reference reads as the adapter it points to, so code that holds
/// an adapter by reference (`query::run(&nav, ..)`) and code that owns a
/// copy (`File::query`, whose walk owns its adapter) run the same resolver
/// and walker.
impl<'f, N: Navigate<'f>> Navigate<'f> for &N {
    type Kind = N::Kind;
    type Node = N::Node;
    type Gap = N::Gap;
    type Children = N::Children;

    const FORMAT: FileFormat = N::FORMAT;
    const ABSENT: &'static [&'static str] = N::ABSENT;

    fn check_source(&self) -> Result<()> {
        (**self).check_source()
    }

    fn fence_key(&self) -> usize {
        (**self).fence_key()
    }

    fn root(&self) -> Result<Self::Node> {
        (**self).root()
    }

    fn child(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
    ) -> Result<Lookup<Self::Node, Self::Gap>> {
        (**self).child(parent, step)
    }

    fn children(&self, parent: &Self::Node, kinds: KindSet<Self::Kind>) -> Result<Self::Children> {
        (**self).children(parent, kinds)
    }

    fn available(
        &self,
        parent: &Self::Node,
        step: &Resolved<'_, Self::Kind>,
        out: &mut Available,
    ) -> Result<()> {
        (**self).available(parent, step, out)
    }
}

/// What [`Navigate::child`] found.
#[derive(Debug)]
pub(crate) enum Lookup<N, G: 'static> {
    /// The selected child.
    Found(N),
    /// Nothing at that position, number, identity or name:
    /// `path_not_found` from `get`, `None` from `find`.
    Missing,
    /// That kind cannot occur under this parent: `unsupported_path`.
    NotApplicable,
    /// The format has the element but its leaf cannot answer yet:
    /// `unsupported_capability`, listed by `capabilities()`.
    Gap(&'static G),
    /// Stacked identities matched this many children (two or more):
    /// `invalid_path` naming the candidates.
    Ambiguous(usize),
}

/// One classified path step: the element kind and how to pick it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Resolved<'s, K> {
    /// The kind of the element the step selects.
    pub(crate) kind: K,
    /// How to pick it among the parent's children.
    pub(crate) select: Select<'s>,
}

/// How a step picks one child of its kind (record 0010 §4.3).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Select<'s> {
    /// No predicate: the only child of a singleton kind (`/body`, `notes`),
    /// or the first of a repeated one, as OfficeCLI does.
    None,
    /// `[N]` read as a zero-based ordinal among the parent's children of
    /// the kind.
    Ordinal(Position),
    /// `[N]` read as a number the element carries, as written (1-based): a
    /// row or column number, a note id, a numbering level. `col[B]` arrives
    /// here as 2.
    Number(usize),
    /// `[last()]`: the last child of the kind (for rows, the highest row
    /// with a record).
    Last,
    /// One or more `[@key=value]` identities, all of which must hold.
    Identity(Identities<'s>),
    /// A name: a worksheet or style named by the step itself, or a bare
    /// `namedrange[NAME]`.
    Name {
        /// The name as written.
        name: &'s str,
        /// How names compare.
        case: Case,
    },
    /// A bare type token: `placeholder[ctrTitle]`.
    Token(&'s str),
    /// An A1 cell inside the grid; a coordinate with no record still
    /// resolves (record 0010 §4.7).
    Cell(CellRef),
    /// An A1 range.
    Range(CellRange),
}

/// Validated `[@key=value]` predicates of one step.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct Identities<'s> {
    predicates: &'s [Predicate<'s>],
    keys: &'static [Identity],
}

impl<'s> Identities<'s> {
    /// Wraps predicates that are all identities of an element with `keys`.
    pub(crate) fn new(predicates: &'s [Predicate<'s>], keys: &'static [Identity]) -> Self {
        Self { predicates, keys }
    }

    /// Each identity with the value it must have, in path order.
    pub(crate) fn iter(&self) -> impl Iterator<Item = (Identity, &'s str)> + '_ {
        self.predicates.iter().filter_map(|predicate| {
            let (key, value) = predicate.identity()?;
            let identity = self
                .keys
                .iter()
                .find(|identity| identity.key.eq_ignore_ascii_case(key))?;
            Some((*identity, value))
        })
    }
}

impl fmt::Debug for Identities<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Identities { .. }")
    }
}

/// A set of a format's element kinds.
pub(crate) struct KindSet<K> {
    bits: u128,
    kind: PhantomData<fn() -> K>,
}

impl<K: ElementKind> KindSet<K> {
    /// No kinds.
    pub(crate) const EMPTY: Self = Self::from_bits(0);
    /// Every kind.
    pub(crate) const ALL: Self = Self::from_bits(u128::MAX);

    /// The set whose members are the kinds with ordinal `n` for every bit
    /// `n` set in `bits`.
    pub(crate) const fn from_bits(bits: u128) -> Self {
        Self {
            bits,
            kind: PhantomData,
        }
    }

    fn bit(kind: K) -> u128 {
        let ordinal = kind.ordinal();
        debug_assert!(ordinal < 128, "{kind:?} has ordinal {ordinal}");
        1_u128 << (ordinal & 127)
    }

    /// The set holding `kind` alone.
    pub(crate) fn of(kind: K) -> Self {
        Self::from_bits(Self::bit(kind))
    }

    /// The set with `kind` added.
    #[must_use]
    pub(crate) fn with(self, kind: K) -> Self {
        Self::from_bits(self.bits | Self::bit(kind))
    }

    /// The kinds in either set.
    #[must_use]
    pub(crate) fn union(self, other: Self) -> Self {
        Self::from_bits(self.bits | other.bits)
    }

    /// Whether `kind` is in the set.
    pub(crate) fn contains(self, kind: K) -> bool {
        self.bits & Self::bit(kind) != 0
    }

    /// Whether the set holds no kind.
    pub(crate) fn is_empty(self) -> bool {
        self.bits == 0
    }
}

impl<K> Clone for KindSet<K> {
    fn clone(&self) -> Self {
        *self
    }
}

impl<K> Copy for KindSet<K> {}

impl<K> PartialEq for KindSet<K> {
    fn eq(&self, other: &Self) -> bool {
        self.bits == other.bits
    }
}

impl<K> Eq for KindSet<K> {}

impl<K> fmt::Debug for KindSet<K> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "KindSet({:#x})", self.bits)
    }
}

/// The most entries a suggestion lists before summarizing the rest.
const MAX_AVAILABLE: usize = 32;

/// What exists under a parent, written by [`Navigate::available`] on the
/// error path and rendered as `p[1]..p[12], tbl[1]..tbl[3]` or as the names
/// of named children.
#[derive(Debug, Default)]
pub(crate) struct Available {
    text: String,
    listed: usize,
    more: usize,
}

impl Available {
    /// Lists `count` children of the element `name`: `p[1]..p[12]`.
    pub(crate) fn kind(&mut self, name: &'static str, count: usize) {
        match count {
            0 => {},
            1 => self.push(format_args!("{name}[1]")),
            _ => self.push(format_args!("{name}[1]..{name}[{count}]")),
        }
    }

    /// Records one named child without retaining its authored name.
    pub(crate) fn name(&mut self, _name: &str) {
        self.push(format_args!("named child"));
    }

    fn push(&mut self, entry: fmt::Arguments<'_>) {
        if self.listed == MAX_AVAILABLE {
            self.more += 1;
            return;
        }
        if self.listed > 0 {
            self.text.push_str(", ");
        }
        // Writing into a `String` cannot fail.
        let _ = self.text.write_fmt(entry);
        self.listed += 1;
    }

    /// Whether nothing was listed.
    pub(crate) fn is_empty(&self) -> bool {
        self.listed == 0
    }

    /// The rendered list, with a count of entries beyond the limit.
    pub(crate) fn render(&self) -> String {
        if self.more == 0 {
            self.text.clone()
        } else {
            format!("{} and {} more", self.text, self.more)
        }
    }
}

impl fmt::Debug for Select<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::None => "None",
            Self::Ordinal(_) => "Ordinal",
            Self::Number(_) => "Number",
            Self::Last => "Last",
            Self::Identity(_) => "Identity",
            Self::Name { .. } => "Name",
            Self::Token(_) => "Token",
            Self::Cell(_) => "Cell",
            Self::Range(_) => "Range",
        })
    }
}

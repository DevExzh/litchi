//! The compiled query plan (record 0010 §5.2, §8.1) and its diagnostics.
//!
//! A [`Plan`] is a selector bound to one format's element kinds and property
//! tables. It holds no reference into the parsed [`Query`]: literals, keys
//! and `:contains` texts borrow the query's text where it borrowed its input
//! and own the rest, so a plan compiled from text outlives the parsed AST.
//! Everything is stored in flat arenas indexed by `u32`, kinds as bits of a
//! `u128` (record 0010 §4.1, at most 128 kinds per format), and each
//! comparison as a lifetime-free compiled test, so evaluating a plan reads
//! no selector text and allocates nothing.

use std::borrow::Cow;
use std::fmt;

use litchi_core::FileFormat;
use litchi_core::selector::{Combinator, Path, Position, Query};

use crate::error::{Code, Error, Result};
use crate::value::compare::Compiled;

/// A compiled, reusable query: parse and compile once, run it on any
/// document of the same format (record 0010 §5.1, decision D9).
///
/// Compilation resolves element names through the format's alias tables,
/// binds each predicate key to the subject kinds' property tables and
/// coerces each literal to the key's type once; running the plan never looks
/// at selector text again. A plan is `Send + Sync`; using it on another
/// format is `invalid_argument`.
///
/// Compile-time findings that do not stop the query, such as an unknown
/// key, are [`Plan::diagnostics`].
///
/// ```
/// use litchi_core::FileFormat;
/// use litchi_ops::query::{Plan, Query};
///
/// let query = Query::parse("paragraph[style=Heading1]")?;
/// // A format whose adapter is not in this build cannot compile a plan.
/// if cfg!(not(feature = "pages")) {
///     let error = Plan::compile(&query, FileFormat::Pages).unwrap_err();
///     assert_eq!(error.code(), litchi_ops::Code::UnsupportedCapability);
/// }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone)]
pub struct Plan<'q> {
    pub(crate) raw: bool,
    /// The format whose kinds and tables the plan is bound to.
    pub(crate) format: FileFormat,
    /// Where the walk starts.
    pub(crate) scope: ScopePlan<'q>,
    /// The top-level alternatives (the union), as chain ids in query order.
    pub(crate) top: Vec<u32>,
    /// Every chain: the top-level alternatives and the selector lists of
    /// `:has` and `:not`.
    pub(crate) chains: Vec<Chain>,
    /// Chain links, subject first.
    pub(crate) links: Vec<Link>,
    /// Chain ids of the `:has`/`:not` lists, referenced by range.
    pub(crate) lists: Vec<u32>,
    /// Every compound.
    pub(crate) compounds: Vec<CompoundPlan<'q>>,
    /// Boolean filter expressions.
    pub(crate) exprs: Vec<ExprNode>,
    /// The operands of `and`/`or` expressions, referenced by range.
    pub(crate) expr_args: Vec<u32>,
    /// Leaf conditions.
    pub(crate) conds: Vec<CondPlan<'q>>,
    /// Header-row predicates (record 0010 §5.4).
    pub(crate) headers: Vec<HeaderPlan<'q>>,
    /// Bare row-property keys of row predicates, refused at run time when a
    /// table column has the same name.
    pub(crate) collisions: Vec<Cow<'q, str>>,
    /// For each kind ordinal, the kinds that can occur below it (the
    /// transitive closure of the spec's child lists): the pruning table,
    /// computed once per spec table (record 0280).
    pub(crate) reach: &'static [u128],
    /// The subject kinds of the top-level alternatives.
    pub(crate) wanted: u128,
    /// The kinds whose direct children `*` selects without a scope.
    pub(crate) top_parents: u128,
    /// The number of chain-link memo slots, for the walker's per-frame
    /// ancestor memo.
    pub(crate) memo_slots: u32,
    /// The header plan of a query whose only alternative is one header-row
    /// compound: the walker then scans the bound table's data rows by number,
    /// rows without a record included, as OfficeCLI does.
    pub(crate) rows_mode: Option<u32>,
    /// Whether every top-level alternative is a lone `*` compound, which
    /// matches only direct children of the scope or of a top-level
    /// container: the walk then enters nothing below them.
    pub(crate) shallow: bool,
    /// The single-subject fast path of the walk, when the plan has one
    /// (record 0315).
    pub(crate) fast: Option<FastPlan>,
    /// Findings that do not stop the query.
    pub(crate) diagnostics: Vec<Diagnostic>,
}

impl<'q> Plan<'q> {
    /// Compiles `query` for documents of `format`.
    ///
    /// # Errors
    ///
    /// `invalid_selector` for an element name no vocabulary knows (with a
    /// "did you mean"), a misplaced positional predicate or a malformed
    /// header-row predicate; `unsupported_path` for an element or scope the
    /// format's family does not have; `unsupported_capability` for an
    /// element or property Litchi cannot read yet, a regular-expression
    /// literal (record 0010 open question Q3), and every format whose
    /// adapter has not registered query support yet.
    pub fn compile(query: &Query<'q>, format: FileFormat) -> Result<Self> {
        super::bounds::validate(query)?;
        // Each format adapter (M1 steps R4-R7) adds an arm compiling for its
        // `Inspect` implementation.
        match format {
            #[cfg(feature = "doc")]
            FileFormat::Doc => super::compile::compile::<crate::format::doc::Nav<'static>>(query),
            #[cfg(feature = "docx")]
            FileFormat::Docx => super::compile::compile::<crate::format::docx::Nav<'static>>(query),
            #[cfg(feature = "xlsx")]
            FileFormat::Xlsx => super::compile::compile::<crate::format::xlsx::Nav<'static>>(query),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => super::compile::compile::<crate::format::xlsb::Nav<'static>>(query),
            #[cfg(feature = "xls")]
            FileFormat::Xls => super::compile::compile::<crate::format::xls::Nav<'static>>(query),
            #[cfg(feature = "pptx")]
            FileFormat::Pptx => super::compile::compile::<crate::format::pptx::Nav<'static>>(query),
            #[cfg(feature = "rtf")]
            FileFormat::Rtf => super::compile::compile::<crate::format::rtf::Nav<'static>>(query),
            #[cfg(feature = "odt")]
            FileFormat::Odt => super::compile::compile::<crate::format::odt::Nav<'static>>(query),
            // OTH and ODM run the ODT adapter's plans (record 0282).
            #[cfg(feature = "oth")]
            FileFormat::Oth => super::compile::compile::<crate::format::odt::Nav<'static>>(query),
            #[cfg(feature = "odm")]
            FileFormat::Odm => super::compile::compile::<crate::format::odt::Nav<'static>>(query),
            #[cfg(feature = "ppt")]
            FileFormat::Ppt => super::compile::compile::<crate::format::ppt::Nav<'static>>(query),
            #[cfg(feature = "ods")]
            FileFormat::Ods => super::compile::compile::<crate::format::ods::Nav<'static>>(query),
            #[cfg(feature = "odp")]
            FileFormat::Odp => super::compile::compile::<crate::format::odp::Nav<'static>>(query),
            #[cfg(odf_objects)]
            FileFormat::Odg
            | FileFormat::Odc
            | FileFormat::Odi
            | FileFormat::Odb
            | FileFormat::Odf
                if crate::read::Adapter::of(format).is_some() =>
            {
                crate::format::odf_tree::compile(query, format)
            },
            #[cfg(feature = "iwork-adapter")]
            FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers
                if crate::read::Adapter::of(format).is_some() =>
            {
                crate::format::iwork::compile(query, format)
            },
            _ => {
                let _ = query;
                Err(Error::new(
                    Code::UnsupportedCapability,
                    format!("querying {format:?} documents is not supported yet"),
                )
                .with_format(format)
                .with_suggestion(
                    "query support arrives with each format's adapter (record 0010, step R4)",
                ))
            },
        }
    }

    /// The format the plan is bound to.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        self.format
    }

    /// Whether this plan has a document-wide scope suitable for a mutation.
    pub(crate) fn is_scoped(&self) -> bool {
        match &self.scope {
            ScopePlan::Root => false,
            ScopePlan::Path(path) => path.is_absolute(),
            ScopePlan::Sheet(_) => true,
        }
    }

    /// The compile-time findings that do not stop the query, in selector
    /// order: unknown keys (which evaluate as absent), relational
    /// comparisons with a non-numeric literal, and tests that can never hold
    /// (record 0010 §5.1).
    #[must_use]
    pub fn diagnostics(&self) -> &[Diagnostic] {
        &self.diagnostics
    }

    /// The number of comma-separated alternatives.
    #[must_use]
    pub fn alternatives(&self) -> usize {
        self.top.len()
    }
}

/// Where a plan's walk starts.
#[derive(Debug, Clone)]
pub(crate) enum ScopePlan<'q> {
    /// The document root.
    Root,
    /// A path resolved through the format adapter (`/body/tbl[1]/...`),
    /// boxed because a path keeps its first steps inline.
    Path(Box<Path<'q>>),
    /// A worksheet named by an Excel-style prefix (`Sheet1!...`).
    Sheet(Cow<'q, str>),
}

/// One complex selector: compounds joined by combinators, stored subject
/// first.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Chain {
    /// The subject's link.
    pub(crate) first: u32,
    /// The number of links.
    pub(crate) len: u32,
    /// The subject compound's kinds.
    pub(crate) kinds: u128,
    /// A positional predicate on the subject of a top-level alternative:
    /// the Nth or last match overall (OfficeCLI's `SelectorPositionalIndex`).
    pub(crate) position: Option<Positional>,
    /// How a node of one of `kinds` is tested against the chain.
    pub(crate) shortcut: Shortcut,
}

/// The walk of a plan with one alternative, one compound, no position, no
/// header rows and no lone `*`, whose test is its kind and at most one
/// shortcut test (`cell`, `cell[value>1000]`, `shape:contains(Revenue)`):
/// its interest never changes, so which kinds match, which are entered and
/// which are asked for are fixed for the whole walk, and computed once with
/// the plan (record 0315). The walk's fast step makes the same reads, in
/// the same order, as the general step: the kind test, the shortcut test,
/// then the child test of a node it may enter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct FastPlan {
    /// The subject's kinds among the wanted ones.
    pub(crate) kinds: u128,
    /// The test beyond the kind: [`Shortcut::KindOnly`], [`Shortcut::Expr`]
    /// or [`Shortcut::Content`].
    pub(crate) test: Shortcut,
    /// The kinds below which a wanted kind can occur (the walk's
    /// `descend`).
    pub(crate) enter: u128,
    /// The kinds to ask the adapter for (the walk's `relevant`).
    pub(crate) ask: u128,
}

impl FastPlan {
    /// The fast path of `plan`, when it has one.
    pub(crate) fn of(plan: &Plan<'_>) -> Option<Self> {
        let [chain] = plan.top.as_slice() else {
            return None;
        };
        let chain = plan.chains[*chain as usize];
        let interest = plan.wanted;
        if plan.shallow || plan.rows_mode.is_some() || interest == 0 || chain.position.is_some() {
            return None;
        }
        if !matches!(
            chain.shortcut,
            Shortcut::KindOnly | Shortcut::Expr(_) | Shortcut::Content(_)
        ) {
            return None;
        }
        let enter = plan
            .reach
            .iter()
            .enumerate()
            .filter(|(_, reach)| *reach & interest != 0)
            .fold(0, |bits, (ordinal, _)| {
                bits | 1_u128
                    .checked_shl(u32::try_from(ordinal).unwrap_or(u32::MAX))
                    .unwrap_or(0)
            });
        Some(Self {
            kinds: chain.kinds & interest,
            test: chain.shortcut,
            enter,
            ask: interest | enter,
        })
    }
}

/// A cheaper test of a one-compound chain on a node of one of its kinds
/// (record 0262); each gives the result of the full test.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Shortcut {
    /// Test the chain in full.
    None,
    /// The compound tests the element kind only (`cell`, `shape`): every
    /// node of its kinds matches.
    KindOnly,
    /// The compound tests the kind and one filter expression
    /// (`cell[value>1000]`, `shape[width>=10cm]`): the expression decides.
    Expr(u32),
    /// The compound (by id) tests the kind and the node's content only
    /// (`shape:contains(Revenue)`): its content tests decide.
    Content(u32),
}

/// A subject-level positional selection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Positional {
    /// The Nth match (1-based).
    Nth(usize),
    /// The last match.
    Last,
}

/// One compound of a chain and how it relates to the link on its right.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Link {
    /// The compound.
    pub(crate) compound: u32,
    /// How this compound relates to the element its right-hand neighbour
    /// matched (`a > b`: `Child`); unused on the subject.
    pub(crate) combinator: Combinator,
    /// The ancestor-memo slot of this link; unused on the subject.
    pub(crate) memo: u32,
}

/// One compiled compound: the tests one element must pass, cheapest first.
#[derive(Debug, Clone)]
pub(crate) struct CompoundPlan<'q> {
    /// The element test.
    pub(crate) kinds: u128,
    /// `*`: the element is a direct child of the scope, or of a top-level
    /// container without one.
    pub(crate) any: bool,
    /// Structural tests: sibling position, element number, cell column or
    /// range.
    pub(crate) tests: Vec<Structural>,
    /// The header-row predicate of a worksheet row compound.
    pub(crate) header: Option<u32>,
    /// The bracket filters, ANDed.
    pub(crate) expr: Option<u32>,
    /// `:contains(...)` texts, all of which must occur (case-insensitively).
    pub(crate) contains: Vec<Cow<'q, str>>,
    /// `:empty`.
    pub(crate) empty: bool,
    /// `:no-alt`: for each kind ordinal, the index of its `alt` property.
    pub(crate) no_alt: Option<Vec<Option<u16>>>,
    /// `:not(...)` lists, as ranges of `Plan::lists`.
    pub(crate) not: Vec<(u32, u32)>,
    /// `:has(...)` lists, as ranges of `Plan::lists`.
    pub(crate) has: Vec<(u32, u32)>,
}

impl CompoundPlan<'_> {
    /// The cheaper test of the compound `id` as a one-compound chain.
    pub(crate) fn shortcut(&self, id: u32) -> Shortcut {
        // Kind tests aside, nothing about the node's place.
        let unplaced = !self.any && self.tests.is_empty() && self.header.is_none();
        let only_expr = self.contains.is_empty()
            && !self.empty
            && self.no_alt.is_none()
            && self.not.is_empty()
            && self.has.is_empty();
        match (unplaced, only_expr, self.expr) {
            (false, ..) => Shortcut::None,
            (true, true, None) => Shortcut::KindOnly,
            (true, true, Some(expr)) => Shortcut::Expr(expr),
            (true, false, _) => Shortcut::Content(id),
        }
    }
}

/// A structural test of a compound.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Structural {
    /// `[N]` on a compound that is not the subject: the Nth sibling of its
    /// kind (zero-based).
    Sibling(Position),
    /// `[last()]` on a compound that is not the subject: the last sibling of
    /// its kind.
    LastSibling,
    /// `[N]` on a number-indexed element (`row[3]`, `level[1]`) or a column
    /// letter (`col[B]`): the element's own number.
    Number(usize),
    /// A column element filter (`Sheet1!B[...]`): the cell's zero-based
    /// column.
    Column(u32),
    /// A bare `placeholder` element on a format that walks placeholders as
    /// shapes: the shape carries a placeholder binding.
    Placeholder,
    /// A cell or range element filter (`A1:C5`): the cell lies inside.
    Range {
        /// The top-left corner: zero-based column and row.
        first: (u32, u32),
        /// The bottom-right corner: zero-based column and row.
        last: (u32, u32),
    },
}

/// A boolean filter expression.
#[derive(Debug, Clone, Copy)]
pub(crate) enum ExprNode {
    /// A leaf condition.
    Cond(u32),
    /// All operands, as a range of `Plan::expr_args`.
    And(u32, u32),
    /// Any operand, as a range of `Plan::expr_args`.
    Or(u32, u32),
    /// The negated operand.
    Not(u32),
}

/// One leaf condition, compiled per subject kind.
#[derive(Debug, Clone)]
pub(crate) struct CondPlan<'q> {
    /// The literal (empty for `[key]`).
    pub(crate) literal: Cow<'q, str>,
    /// For each kind ordinal, its target in `targets`; kinds the compound
    /// does not select map to target 0.
    pub(crate) by_kind: Vec<u8>,
    /// The distinct targets; target 0 evaluates the key as absent.
    pub(crate) targets: Vec<Target>,
    /// The condition is the only one of its compound that reads properties
    /// (and the compound has no `:no-alt`): a property target reads its one
    /// value ([`Inspect::prop`](super::inspect::Inspect::prop)) instead of
    /// the node's whole property set.
    pub(crate) direct: bool,
}

impl CondPlan<'_> {
    /// Whether a target reads the node's property set.
    pub(crate) fn reads_properties(&self) -> bool {
        self.targets
            .iter()
            .any(|target| matches!(target, Target::Prop { .. } | Target::Style { .. }))
    }
}

/// What a condition reads for one kind, and its test typed by that read.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Target {
    /// The kind cannot have the key: the condition sees it absent.
    Absent(Compiled),
    /// A declared property, by index into the kind's property table.
    Prop {
        /// The property's index.
        index: u16,
        /// The test, typed by the property's value type.
        test: Compiled,
    },
    /// The node's text (`text`, and a cell's `value`).
    Text(Compiled),
    /// The node's OfficeCLI type (`type`).
    Type(Compiled),
    /// The `style` key: the declared `style` property, the node's style
    /// member, and for `=`/`!=` also the `styleName` property (OfficeCLI's
    /// dual style key).
    Style {
        /// The `style` property's index.
        prop: Option<u16>,
        /// The `styleName` property's index.
        name: Option<u16>,
        /// The test.
        test: Compiled,
    },
    /// A header-row column, by index into the header plan's keys.
    Column {
        /// The header plan.
        header: u16,
        /// The key's index.
        key: u16,
        /// The test, on the raw or displayed cell text.
        test: Compiled,
    },
}

/// A header-row predicate (record 0010 §5.4).
#[derive(Debug, Clone)]
pub(crate) struct HeaderPlan<'q> {
    /// The distinct column keys, without `col.`/`column.`, and whether the
    /// prefix forced them.
    pub(crate) keys: Vec<(Cow<'q, str>, bool)>,
    /// For each key, whether its conditions compare the raw stored value: a
    /// relational operator or a numeric literal (OfficeCLI's `wantsRaw`).
    pub(crate) raw: Vec<bool>,
    /// The worksheet kind's ordinal, to find a row's sheet among its
    /// ancestors.
    pub(crate) sheet_kind: u8,
    /// The worksheet row kind's ordinal.
    pub(crate) row_kind: u8,
}

/// A compile-time finding that does not stop the query (record 0010 §5.1).
///
/// OfficeCLI reports the same findings as warnings once a query has run
/// (`AttributeFilter.ApplyWithWarnings`); Litchi reports them from the plan,
/// before the first node is visited.
///
/// ```
/// use litchi_ops::query::DiagnosticKind;
///
/// assert_eq!(DiagnosticKind::UnknownKey.as_str(), "unknown_key");
/// ```
#[derive(Clone, PartialEq, Eq)]
pub struct Diagnostic {
    kind: DiagnosticKind,
    message: String,
    suggestion: Option<String>,
}

impl Diagnostic {
    pub(crate) fn new(
        kind: DiagnosticKind,
        _key: &str,
        message: String,
        suggestion: Option<String>,
    ) -> Self {
        Self {
            kind,
            message,
            suggestion,
        }
    }

    /// What was found.
    #[must_use]
    pub fn kind(&self) -> DiagnosticKind {
        self.kind
    }

    /// A human-readable description.
    #[must_use]
    pub fn message(&self) -> &str {
        &self.message
    }

    /// A correction, when one is known ("did you mean 'font.bold'?").
    #[must_use]
    pub fn suggestion(&self) -> Option<&str> {
        self.suggestion.as_deref()
    }
}

impl fmt::Display for Diagnostic {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}: {}", self.kind.as_str(), self.message)?;
        if let Some(suggestion) = &self.suggestion {
            write!(f, " ({suggestion})")?;
        }
        Ok(())
    }
}

/// The kind of a [`Diagnostic`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DiagnosticKind {
    /// No subject kind declares the key: the condition evaluates as if the
    /// key were absent (only `!=` holds).
    UnknownKey,
    /// A relational comparison (`>`, `<`, `>=`, `<=`) with a literal that is
    /// neither a number nor a length: it never holds.
    NonNumeric,
    /// A test no subject kind can pass, such as `:no-alt` on elements
    /// without alternative text.
    NeverMatches,
}

impl DiagnosticKind {
    /// The stable `snake_case` spelling, as OfficeCLI's `query --json`
    /// writes the warning kind.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::UnknownKey => "unknown_key",
            Self::NonNumeric => "non_numeric",
            Self::NeverMatches => "never_matches",
        }
    }
}

const fn assert_send_sync<T: Send + Sync>() {}
const _: () = assert_send_sync::<Plan<'static>>();

impl fmt::Debug for Plan<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Plan { .. }")
    }
}
impl fmt::Debug for Diagnostic {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Diagnostic")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

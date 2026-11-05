//! Query compilation (record 0010 §5.2): one pass over a parsed [`Query`]
//! that binds every element name, predicate key and literal to one format's
//! tables.
//!
//! - Element names go through the family vocabulary's alias table, ASCII
//!   case-insensitively, exactly as path steps do (record 0010 §4.2);
//!   `title` and `textbox` narrow presentation shapes; a workbook also
//!   accepts column letters (`B`) and A1 cells or ranges as element tests.
//! - Each predicate key is looked up once per subject kind in the kind's
//!   property table, through OfficeCLI's selector key aliases; its literal is
//!   coerced once into the key's [`ValueType`] by
//!   [`Compiled`](crate::value::compare::Compiled).
//! - `[N]` and `[last()]` on the subject of a top-level alternative select
//!   the Nth or last match; elsewhere they select among siblings of the
//!   element's kind; on a number-indexed element (`row[3]`, `col[B]`) they
//!   keep their path meaning.
//! - Header-row predicates (`row[Salary>5000]`, record 0010 §5.4) are
//!   classified into table columns and row properties here; the format
//!   adapter binds the columns when the query runs.
//!
//! Compilation allocates (the plan's arenas, owned copies of keys); running a
//! plan does not.

use std::borrow::Cow;
use std::marker::PhantomData;

use litchi_core::selector::{
    CellRange, Combinator, CompareOp, Complex, Compound, Condition, Element, Expr, Predicate,
    Pseudo, Query, Scope, Value as Literal, column_number,
};

use super::inspect::Inspect;
use super::plan::{
    Chain, CompoundPlan, CondPlan, Diagnostic, DiagnosticKind, ExprNode, HeaderPlan, Link, Plan,
    Positional, ScopePlan, Shortcut, Structural, Target,
};
use crate::Family;
use crate::error::{Code, Error, Result};
use crate::node::spec::{ElementSpec, PropSpec, Subject};
use crate::path::navigate::ElementKind;
use crate::path::report::{did_you_mean, distance};
use crate::path::vocabulary::{self, Bare, IndexRule};
use crate::value::ValueType;
use crate::value::compare::{self, Compiled};
use crate::value::parse_numeric_text;

/// Compiles `query` for the format adapter `N`.
pub(crate) fn compile<'q, 'f, N: Inspect<'f>>(query: &Query<'q>) -> Result<Plan<'q>> {
    compile_with_raw::<N>(query, false)
}

pub(super) fn compile_with_raw<'q, 'f, N: Inspect<'f>>(
    query: &Query<'q>,
    raw: bool,
) -> Result<Plan<'q>> {
    super::bounds::validate(query)?;
    let mut compiler = Compiler::<N>::new()?;
    compiler.plan.raw = raw;
    compiler.plan.scope = compiler.scope(query.scope())?;
    compiler.explicit_scope = !matches!(compiler.plan.scope, ScopePlan::Root);
    for complex in query.alternatives() {
        let chain = compiler.chain(complex, Role::Top)?;
        compiler.plan.top.push(chain);
    }
    compiler.finish()
}

/// Whether a chain is a top-level alternative or a `:has`/`:not` list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Role {
    Top,
    Nested,
}

/// Where a compound sits in its chain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Place {
    Subject(Role),
    Ancestor,
}

/// A selector-only element name that narrows its kind (OfficeCLI
/// `PowerPointHandler.Selector.cs`: `title` is a title or centred-title
/// placeholder, `textbox` any other shape).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Refine {
    Title,
    NotTitle,
}

/// `(family, element name)` → refinement. In paths both names stay plain
/// aliases of `shape`, counted with every other shape.
const REFINEMENTS: &[(Family, &str, Refine)] = &[
    (Family::Presentation, "textbox", Refine::NotTitle),
    (Family::Presentation, "title", Refine::Title),
];

/// Selector key aliases, applied when the element does not declare the key
/// itself: OfficeCLI's Excel cell aliases (`ExcelHandler.Selector.cs`
/// `_cellSelectorAliases`) and its PowerPoint `title` key
/// (`PowerPointHandler.Selector.cs`). `(family, canonical element, alias,
/// key)`; aliases are lowercase and matched ASCII case-insensitively.
const KEY_ALIASES: &[(Family, &str, &str, &str)] = &[
    (Family::Presentation, "shape", "title", "isTitle"),
    (Family::Workbook, "cell", "bold", "font.bold"),
    (Family::Workbook, "cell", "color", "font.color"),
    (Family::Workbook, "cell", "font", "font.name"),
    (Family::Workbook, "cell", "italic", "font.italic"),
    (Family::Workbook, "cell", "size", "font.size"),
];

/// The worksheet row properties a header-row predicate may filter
/// (OfficeCLI `ExcelHandler.Query.RowWhere.cs` `RowAttributeKeys`); any
/// other bare key names a table column.
const ROW_PROPERTIES: &[&str] = &[
    "collapsed",
    "customHeight",
    "height",
    "hidden",
    "outlineLevel",
];

/// Keys every element answers without declaring them (OfficeCLI
/// `AttributeFilter.ResolveValue`'s fallbacks).
const PSEUDO_KEYS: &[&str] = &["style", "text", "type"];

/// The element whose `value` key reads its text, as OfficeCLI's Excel
/// selector reads a cell's value.
const VALUE_ELEMENT: &str = "cell";

/// The keys a selector compares on an element of the canonical vocabulary
/// name `canonical` besides its declared properties: what `capabilities()`
/// reports as the element's predicate keys.
pub(crate) fn pseudo_keys(canonical: &str) -> impl Iterator<Item = &'static str> + Clone {
    PSEUDO_KEYS
        .iter()
        .copied()
        .chain((canonical == VALUE_ELEMENT).then_some("value"))
}

/// The selector spellings that read a declared key of an element of the
/// canonical name `canonical` in `family` (`bold` for a cell's
/// `font.bold`), with the key each one reads: what `capabilities()` reports
/// as the element's predicate aliases.
pub(crate) fn key_aliases(
    family: Family,
    canonical: &'static str,
) -> impl Iterator<Item = (&'static str, &'static str)> + Clone {
    KEY_ALIASES
        .iter()
        .filter(move |(owner, element, _, _)| *owner == family && *element == canonical)
        .map(|(_, _, alias, key)| (*alias, *key))
}

/// Native-key admission shared by query compilation and capability reporting.
pub(crate) fn native_key<K: ElementKind>(spec: &ElementSpec<K>, name: &str) -> bool {
    let semantic = ["name", "role", "type"]
        .iter()
        .any(|key| key.eq_ignore_ascii_case(name));
    !semantic
        && (spec
            .identities
            .iter()
            .any(|key| key.eq_ignore_ascii_case(name))
            || spec
                .props
                .iter()
                .any(|prop| prop.raw && prop.name.eq_ignore_ascii_case(name))
            || K::VOCABULARY
                .canonical(spec.name)
                .and_then(|canonical| K::VOCABULARY.rule(canonical))
                .and_then(|rule| rule.identity(name))
                .is_some_and(|identity| identity.is_native()))
}

struct Compiler<'q, 'f, N> {
    plan: Plan<'q>,
    /// The kinds a query can select: every kind without an element gap.
    all: u128,
    /// Whether the query names a scope, which `*` selects the children of.
    explicit_scope: bool,
    adapter: PhantomData<fn(&'f ()) -> N>,
}

/// The bit of the kind with ordinal `ordinal`.
/// The bit of the kind with `ordinal`; the spec table is checked to hold at
/// most 128 kinds before any is used.
const fn bit_of(ordinal: usize) -> u128 {
    if ordinal < 128 { 1 << ordinal } else { 0 }
}

/// The ordinals of the kinds in `bits`, ascending.
fn ordinals(mut bits: u128) -> impl Iterator<Item = usize> {
    std::iter::from_fn(move || {
        if bits == 0 {
            return None;
        }
        let ordinal = bits.trailing_zeros() as usize;
        bits &= bits - 1;
        Some(ordinal)
    })
}

impl<'q, 'f, N: Inspect<'f>> Compiler<'q, 'f, N> {
    fn new() -> Result<Self> {
        let elements = N::SPEC.elements;
        if elements.len() > 128 {
            return Err(internal(format!(
                "the {:?} spec table has {} kinds; a query kind set holds 128",
                N::FORMAT,
                elements.len()
            )));
        }
        for (index, element) in elements.iter().enumerate() {
            if usize::from(element.kind.ordinal()) != index {
                return Err(internal(format!(
                    "the {:?} spec table is not indexed by kind ordinal at {index}",
                    N::FORMAT
                )));
            }
        }
        let reach = reach_table(elements);
        let all = elements
            .iter()
            .enumerate()
            .filter(|(_, element)| element.gap.is_none())
            .fold(0, |bits, (index, _)| bits | bit_of(index));
        let top_parents = N::TOP
            .iter()
            .fold(0, |bits, kind| bits | bit_of(usize::from(kind.ordinal())));
        Ok(Self {
            plan: Plan {
                raw: false,
                format: N::FORMAT,
                scope: ScopePlan::Root,
                top: Vec::new(),
                chains: Vec::new(),
                links: Vec::new(),
                lists: Vec::new(),
                compounds: Vec::new(),
                exprs: Vec::new(),
                expr_args: Vec::new(),
                conds: Vec::new(),
                headers: Vec::new(),
                collisions: Vec::new(),
                reach,
                wanted: 0,
                top_parents,
                memo_slots: 0,
                rows_mode: None,
                shallow: false,
                fast: None,
                diagnostics: Vec::new(),
            },
            all,
            explicit_scope: false,
            adapter: PhantomData,
        })
    }

    fn kind_count(&self) -> usize {
        N::SPEC.elements.len()
    }

    fn spec(&self, ordinal: usize) -> &'static ElementSpec<N::Kind> {
        &N::SPEC.elements[ordinal]
    }

    /// The ordinal of a modeled kind, checked against the spec table.
    fn ordinal(&self, kind: N::Kind) -> Result<usize> {
        let ordinal = usize::from(kind.ordinal());
        if ordinal < self.kind_count() {
            Ok(ordinal)
        } else {
            Err(internal(format!(
                "element kind \"{}\" has no {:?} spec entry",
                kind.name(),
                N::FORMAT
            )))
        }
    }

    fn native_key(&self, ordinal: usize, name: &str) -> bool {
        native_key(self.spec(ordinal), name)
    }

    fn scope(&self, scope: Option<&Scope<'q>>) -> Result<ScopePlan<'q>> {
        Ok(match scope {
            None => ScopePlan::Root,
            Some(Scope::Path(path)) => {
                if !self.plan.raw {
                    for step in path.steps() {
                        if let Some(canonical) = N::Kind::VOCABULARY.canonical(step.name()) {
                            for predicate in step.predicates() {
                                if let Some((key, _)) = predicate.identity() {
                                    let native = N::Kind::VOCABULARY
                                        .rule(canonical)
                                        .and_then(|rule| rule.identity(key))
                                        .is_some_and(|identity| identity.is_native())
                                        || N::Kind::from_name(canonical).is_some_and(|kind| {
                                            self.native_key(usize::from(kind.ordinal()), key)
                                        });
                                    if native {
                                        return Err(Error::new(
                                            Code::InvalidSelector,
                                            "native identity scope requires an explicit raw query",
                                        )
                                        .with_suggestion(
                                            "use query::raw::query for native identity scope paths",
                                        ));
                                    }
                                }
                            }
                        }
                    }
                }
                ScopePlan::Path(path.clone())
            },
            Some(Scope::Sheet(name)) => {
                if N::Kind::VOCABULARY.family != Family::Workbook {
                    return Err(Error::new(
                        Code::UnsupportedPath,
                        "worksheet scope does not apply to this format",
                    )
                    .with_format(N::FORMAT)
                    .with_suggestion(format!(
                        "scope a query with a path such as {}",
                        N::Kind::VOCABULARY.roots
                    )));
                }
                ScopePlan::Sheet(name.clone())
            },
            Some(_) => return Err(unsupported("this kind of query scope")),
        })
    }

    fn chain(&mut self, complex: &Complex<'q>, role: Role) -> Result<u32> {
        let compounds: Vec<&Compound<'q>> = complex.compounds().collect();
        let subject_index = compounds.len() - 1;
        let (subject, position) = self.compound(compounds[subject_index], Place::Subject(role))?;
        let kinds = self.plan.compounds[subject as usize].kinds;
        let mut links = vec![Link {
            compound: subject,
            combinator: Combinator::Child,
            memo: 0,
        }];
        for index in (0..subject_index).rev() {
            let combinator = complex.rest()[index].0;
            let (compound, _) = self.compound(compounds[index], Place::Ancestor)?;
            let memo = self.plan.memo_slots;
            self.plan.memo_slots += 1;
            links.push(Link {
                compound,
                combinator,
                memo,
            });
        }
        let first = to_u32(self.plan.links.len())?;
        let len = to_u32(links.len())?;
        self.plan.links.extend(links);
        let id = to_u32(self.plan.chains.len())?;
        let shortcut = if len == 1 {
            self.plan.compounds[subject as usize].shortcut(subject)
        } else {
            Shortcut::None
        };
        self.plan.chains.push(Chain {
            first,
            len,
            kinds,
            position,
            shortcut,
        });
        Ok(id)
    }

    fn compound(
        &mut self,
        compound: &Compound<'q>,
        place: Place,
    ) -> Result<(u32, Option<Positional>)> {
        let mut tests = Vec::new();
        let mut refine = None;
        let mut any = false;
        let kinds = match compound.element() {
            Element::Implicit => self.all,
            Element::Any => {
                any = true;
                if self.explicit_scope {
                    self.all
                } else {
                    self.top_children()
                }
            },
            Element::Named(name) => {
                let (kinds, test, refinement) = self.element(name)?;
                tests.extend(test);
                refine = refinement;
                kinds
            },
            _ => return Err(unsupported("this kind of element test")),
        };
        let single = (kinds.count_ones() == 1).then(|| kinds.trailing_zeros() as usize);
        let rule =
            single.and_then(|ordinal| N::Kind::VOCABULARY.rule(self.spec(ordinal).kind.name()));
        let numbered = rule.is_some_and(|rule| {
            matches!(
                rule.index,
                IndexRule::Number | IndexRule::Row | IndexRule::Column
            )
        });
        let mut position = None;
        let mut conds = Vec::new();
        if let Some(refine) = refine {
            conds.push(self.refinement(kinds, refine)?);
        }
        let mut filters: Vec<&Expr<'q>> = Vec::new();
        for predicate in compound.predicates() {
            match predicate {
                Predicate::Index(at) => {
                    let number = at.one_based().ok_or_else(|| {
                        invalid_selector("the index is too large".to_owned(), None)
                    })?;
                    if numbered {
                        tests.push(Structural::Number(number));
                    } else if place == Place::Subject(Role::Top) {
                        set_position(&mut position, Positional::Nth(number))?;
                    } else {
                        tests.push(Structural::Sibling(*at));
                    }
                },
                Predicate::Last => {
                    if place == Place::Subject(Role::Top) {
                        set_position(&mut position, Positional::Last)?;
                    } else {
                        tests.push(Structural::LastSibling);
                    }
                },
                Predicate::Filter(expr) => {
                    match (predicate.bare_key(), rule.and_then(|rule| rule.bare)) {
                        (Some(token), Some(Bare::Column)) => {
                            let column = column_number(token).ok_or_else(|| {
                                invalid_selector(
                                    "invalid worksheet column",
                                    Some(Cow::Borrowed(
                                        "columns run from A to XFD, e.g. col[B] or col[2]",
                                    )),
                                )
                            })?;
                            tests.push(Structural::Number(column as usize));
                        },
                        _ => filters.push(expr),
                    }
                },
                _ => return Err(unsupported("this kind of predicate")),
            }
        }
        let mut plan = CompoundPlan {
            kinds,
            any,
            tests,
            header: None,
            expr: None,
            contains: Vec::new(),
            empty: false,
            no_alt: None,
            not: Vec::new(),
            has: Vec::new(),
        };
        if !filters.is_empty() {
            if self.is_row(single) {
                self.row_filters(&filters, kinds, &mut plan, &mut conds)?;
            } else {
                for expr in filters {
                    conds.push(self.expr(expr, kinds, None)?);
                }
            }
        }
        for pseudo in compound.pseudos() {
            match pseudo {
                Pseudo::Contains(text) => plan.contains.push(text.clone()),
                Pseudo::Empty => plan.empty = true,
                Pseudo::NoAlt => plan.no_alt = Some(self.no_alt(kinds)?),
                Pseudo::Has(list) => match self.literal_has(list, kinds) {
                    Some(key) => {
                        let cond = self.condition_parts(&key, None, kinds, None)?;
                        conds.push(self.push_expr(ExprNode::Cond(cond))?);
                    },
                    None => {
                        let range = self.list(list)?;
                        plan.has.push(range);
                    },
                },
                Pseudo::Not(list) => {
                    let range = self.list(list)?;
                    plan.not.push(range);
                },
                _ => return Err(unsupported("this pseudo-class")),
            }
        }
        plan.expr = self.and(conds)?;
        if plan.no_alt.is_none() {
            if let Some(expr) = plan.expr {
                self.mark_direct(expr);
            }
        }
        let id = to_u32(self.plan.compounds.len())?;
        self.plan.compounds.push(plan);
        Ok((id, position))
    }

    /// Marks the one condition of the expression `expr` that reads
    /// properties, when only one does, to read its value alone.
    fn mark_direct(&mut self, expr: u32) {
        let (mut readers, mut last) = (0_usize, None);
        self.property_readers(expr, &mut readers, &mut last);
        if let (1, Some(only)) = (readers, last) {
            self.plan.conds[only as usize].direct = true;
        }
    }

    /// Counts the conditions of `expr` that read properties, keeping the
    /// last one. Recurses as deep as the expression, which the grammar
    /// bounds.
    fn property_readers(&self, expr: u32, readers: &mut usize, last: &mut Option<u32>) {
        match self.plan.exprs[expr as usize] {
            ExprNode::Cond(cond) => {
                if self.plan.conds[cond as usize].reads_properties() {
                    *readers += 1;
                    *last = Some(cond);
                }
            },
            ExprNode::And(start, len) | ExprNode::Or(start, len) => {
                for &arg in &self.plan.expr_args[start as usize..(start + len) as usize] {
                    self.property_readers(arg, readers, last);
                }
            },
            ExprNode::Not(inner) => self.property_readers(inner, readers, last),
        }
    }

    /// The kind set, structural test and refinement of a named element.
    fn element(&mut self, name: &str) -> Result<(u128, Option<Structural>, Option<Refine>)> {
        let vocabulary = N::Kind::VOCABULARY;
        if let Some(&(_, _, refine)) = REFINEMENTS.iter().find(|(family, alias, _)| {
            *family == vocabulary.family && alias.eq_ignore_ascii_case(name)
        }) {
            let kind = self.model("shape")?;
            return Ok((bit_of(kind), None, Some(refine)));
        }
        if let Some(canonical) = vocabulary.canonical(name) {
            if N::PLACEHOLDER_SHAPES && canonical == "placeholder" {
                // Placeholders are walked as shapes (record 0290 F1).
                let shape = self.model("shape")?;
                return Ok((bit_of(shape), Some(Structural::Placeholder), None));
            }
            return Ok((bit_of(self.model(canonical)?), None, None));
        }
        if let Some(other) = vocabulary::elsewhere(vocabulary, name) {
            return Err(Error::new(
                Code::UnsupportedPath,
                format!(
                    "the selected element belongs to {}",
                    family_noun(other.family)
                ),
            )
            .with_format(N::FORMAT)
            .with_suggestion(format!(
                "{:?} paths start at {}",
                N::FORMAT,
                vocabulary.roots
            )));
        }
        if vocabulary.family == Family::Workbook {
            if let Some(range) = CellRange::parse(name) {
                let (start, end) = (range.start(), range.end());
                let test = Structural::Range {
                    first: (start.column().min(end.column()), start.row().min(end.row())),
                    last: (start.column().max(end.column()), start.row().max(end.row())),
                };
                return Ok((bit_of(self.model("cell")?), Some(test), None));
            }
            if name.len() <= 3 && name.bytes().all(|byte| byte.is_ascii_alphabetic()) {
                if let Some(column) = column_number(name) {
                    let cell = self.model("cell")?;
                    return Ok((bit_of(cell), Some(Structural::Column(column - 1)), None));
                }
            }
        }
        Err(Error::new(Code::InvalidSelector, "unknown query element")
            .with_format(N::FORMAT)
            .with_suggestion(did_you_mean(name, vocabulary)))
    }

    /// The ordinal of the kind with canonical name `canonical`:
    /// `unsupported_path` when the format lacks the concept,
    /// `unsupported_capability` when the adapter does not model it or lists
    /// it as a gap.
    fn model(&self, canonical: &str) -> Result<usize> {
        let Some(kind) = N::Kind::from_name(canonical) else {
            let absent = N::ABSENT
                .iter()
                .any(|absent| absent.eq_ignore_ascii_case(canonical));
            return Err(if absent {
                Error::new(
                    Code::UnsupportedPath,
                    format!("{:?} documents have no \"{canonical}\" elements", N::FORMAT),
                )
                .with_format(N::FORMAT)
                .with_suggestion(format!(
                    "{:?} paths start at {}",
                    N::FORMAT,
                    N::Kind::VOCABULARY.roots
                ))
            } else {
                Error::new(
                    Code::UnsupportedCapability,
                    format!(
                        "querying \"{canonical}\" elements of {:?} documents is not supported yet",
                        N::FORMAT
                    ),
                )
                .with_format(N::FORMAT)
                .with_suggestion("capabilities() lists the elements Litchi reads from this format")
            });
        };
        let ordinal = self.ordinal(kind)?;
        if let Some(gap) = self.spec(ordinal).gap {
            return Err(gap.error().with_format(N::FORMAT));
        }
        Ok(ordinal)
    }

    /// The children of the top-level containers: what `*` selects without a
    /// scope.
    fn top_children(&self) -> u128 {
        let mut bits = 0;
        for ordinal in ordinals(self.plan.top_parents) {
            for child in self.spec(ordinal).children {
                bits |= bit_of(usize::from(child.ordinal()));
            }
        }
        bits & self.all
    }

    /// Whether the compound selects exactly the workbook `row` kind.
    fn is_row(&self, single: Option<usize>) -> bool {
        N::Kind::VOCABULARY.family == Family::Workbook
            && single.is_some_and(|ordinal| self.spec(ordinal).kind.name() == "row")
    }

    /// The `isTitle` test a `title` or `textbox` element name adds.
    fn refinement(&mut self, kinds: u128, refine: Refine) -> Result<u32> {
        for ordinal in ordinals(kinds) {
            if find_prop(self.spec(ordinal), "isTitle").is_none() {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    format!(
                        "selecting title or text-box shapes of {:?} documents is not supported yet",
                        N::FORMAT
                    ),
                )
                .with_format(N::FORMAT)
                .with_suggestion(
                    "the format does not read the isTitle property yet; query shape instead",
                ));
            }
        }
        let op = match refine {
            Refine::Title => CompareOp::Eq,
            Refine::NotTitle => CompareOp::Ne,
        };
        let cond =
            self.condition_parts("isTitle", Some((op, Cow::Borrowed("true"))), kinds, None)?;
        self.push_expr(ExprNode::Cond(cond))
    }

    /// Splits a worksheet row compound's filters into table columns and row
    /// properties (record 0010 §5.4 rule 1, OfficeCLI
    /// `ExcelHandler.Query.cs` 1716-1791).
    fn row_filters(
        &mut self,
        filters: &[&Expr<'q>],
        kinds: u128,
        plan: &mut CompoundPlan<'q>,
        conds: &mut Vec<u32>,
    ) -> Result<()> {
        let mut leaves = Vec::new();
        for expr in filters {
            collect_leaves(expr, &mut leaves);
        }
        let row = kinds.trailing_zeros() as usize;
        let mut columns: Vec<(Cow<'q, str>, bool)> = Vec::new();
        let mut raw: Vec<bool> = Vec::new();
        let mut collisions: Vec<Cow<'q, str>> = Vec::new();
        let (mut column_count, mut property_count) = (0_usize, 0_usize);
        for leaf in leaves {
            let key = leaf.key();
            let name = key.name();
            let attribute = key.is_attribute();
            let stripped = strip_column_prefix(name);
            let forced = !attribute && stripped.is_some();
            if !forced && !name.is_empty() && name.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid_selector(
                    "a numeric header key is ambiguous with a row position",
                    Some(Cow::Borrowed("use col. to select a numeric column name")),
                ));
            }
            let declared = find_prop(self.spec(row), name).is_some();
            let property = is_row_property(name);
            if !forced && attribute && !property && !declared {
                return Err(invalid_selector(
                    "the attribute is not a row property",
                    Some(Cow::Borrowed(
                        "use a declared row property, or col. for a table column",
                    )),
                ));
            }
            if !forced && (attribute || property) {
                property_count += 1;
                if !attribute
                    && !collisions
                        .iter()
                        .any(|seen| seen.eq_ignore_ascii_case(name))
                {
                    collisions.push(Cow::Owned(name.to_owned()));
                }
                continue;
            }
            column_count += 1;
            let column = stripped.unwrap_or(name);
            let wants_raw = match leaf {
                Condition::Compare { op, value, .. } => {
                    matches!(
                        op,
                        CompareOp::Gt | CompareOp::Lt | CompareOp::Ge | CompareOp::Le
                    ) || parse_numeric_text(value.as_str()).is_some()
                },
                _ => false,
            };
            match columns
                .iter()
                .position(|(seen, _)| seen.eq_ignore_ascii_case(column))
            {
                Some(index) => {
                    raw[index] |= wants_raw;
                    columns[index].1 |= forced;
                },
                None => {
                    columns.push((Cow::Owned(column.to_owned()), forced));
                    raw.push(wants_raw);
                },
            }
        }
        if column_count > 0 && property_count > 0 {
            return Err(invalid_selector(
                "row[...] cannot mix table columns and row properties in one expression".to_owned(),
                Some(Cow::Borrowed("split it into separate queries")),
            ));
        }
        if column_count == 0 {
            self.plan.collisions.extend(collisions);
            for expr in filters {
                conds.push(self.expr(expr, kinds, None)?);
            }
            return Ok(());
        }
        let sheet = N::Kind::from_name("sheet")
            .map(|kind| self.ordinal(kind))
            .transpose()?
            .ok_or_else(|| unsupported("header-row predicates in a format without worksheets"))?;
        let header = to_u32(self.plan.headers.len())?;
        self.plan.headers.push(HeaderPlan {
            keys: columns,
            raw,
            sheet_kind: to_u8(sheet)?,
            row_kind: to_u8(row)?,
        });
        plan.header = Some(header);
        for expr in filters {
            conds.push(self.expr(expr, kinds, Some(header))?);
        }
        Ok(())
    }

    fn expr(&mut self, expr: &Expr<'q>, kinds: u128, header: Option<u32>) -> Result<u32> {
        let node = match expr {
            Expr::Condition(condition) => ExprNode::Cond(self.condition(condition, kinds, header)?),
            Expr::And(parts) | Expr::Or(parts) => {
                let mut ids = Vec::with_capacity(parts.len());
                for part in parts {
                    ids.push(self.expr(part, kinds, header)?);
                }
                let (start, len) = self.args(ids)?;
                if matches!(expr, Expr::And(_)) {
                    ExprNode::And(start, len)
                } else {
                    ExprNode::Or(start, len)
                }
            },
            Expr::Not(inner) => ExprNode::Not(self.expr(inner, kinds, header)?),
            _ => return Err(unsupported("this kind of filter expression")),
        };
        self.push_expr(node)
    }

    fn condition(
        &mut self,
        condition: &Condition<'q>,
        kinds: u128,
        header: Option<u32>,
    ) -> Result<u32> {
        match condition {
            Condition::Exists(key) => self.condition_parts(key.name(), None, kinds, header),
            Condition::Compare { key, op, value } => {
                compare::literal(*op, value)?;
                let Literal::Text(text) = value else {
                    return Err(unsupported("this kind of comparison value"));
                };
                self.condition_parts(key.name(), Some((*op, text.clone())), kinds, header)
            },
            _ => Err(unsupported("this kind of condition")),
        }
    }

    /// Compiles `[name]` (`comparison` is `None`) or `[name op literal]` for
    /// every kind in `kinds`, or for the header-row column `name`.
    fn condition_parts(
        &mut self,
        name: &str,
        comparison: Option<(CompareOp, Cow<'q, str>)>,
        kinds: u128,
        header: Option<u32>,
    ) -> Result<u32> {
        if !self.plan.raw
            && header.is_none()
            && ordinals(kinds).any(|ordinal| self.native_key(ordinal, name))
        {
            return Err(Error::new(
                Code::InvalidSelector,
                "native identity requires an explicit raw query",
            )
            .with_suggestion("use query::raw::query for native identity predicates"));
        }
        let (op, literal) = match comparison {
            Some((op, literal)) => (Some(op), literal),
            None => (None, Cow::Borrowed("")),
        };
        let test = |ty: Option<ValueType>| match op {
            Some(op) => Compiled::new(op, &literal, ty),
            None => Compiled::EXISTS,
        };
        let mut by_kind = vec![0_u8; self.kind_count()];
        let mut targets = vec![Target::Absent(test(None))];
        let mut keys: Vec<TargetKey> = vec![TargetKey::Absent];
        let mut known = false;
        match header {
            Some(header) => {
                let column = strip_column_prefix(name).unwrap_or(name);
                let key = self.plan.headers[header as usize]
                    .keys
                    .iter()
                    .position(|(seen, _)| seen.eq_ignore_ascii_case(column))
                    .ok_or_else(|| internal("a header-row column was not classified"))?;
                targets.push(Target::Column {
                    header: to_u16(header as usize)?,
                    key: to_u16(key)?,
                    test: test(None),
                });
                for ordinal in ordinals(kinds) {
                    by_kind[ordinal] = 1;
                }
                known = true;
            },
            None => {
                for ordinal in ordinals(kinds) {
                    let (key, target) = self.target(ordinal, name, &test)?;
                    known |= key != TargetKey::Absent;
                    let index = match keys.iter().position(|seen| *seen == key) {
                        Some(index) => index,
                        None => {
                            keys.push(key);
                            targets.push(target);
                            keys.len() - 1
                        },
                    };
                    by_kind[ordinal] = to_u8(index)?;
                }
            },
        }
        if !known && op != Some(CompareOp::Ne) {
            self.unknown_key(name, kinds);
        }
        if op.is_some() {
            if test(None).warning().is_some() {
                self.plan.diagnostics.push(Diagnostic::new(
                    DiagnosticKind::NonNumeric,
                    name,
                    "a non-numeric relational operand never matches".to_owned(),
                    Some("compare with a number or a length such as 12, 24pt or 2cm".to_owned()),
                ));
            }
        }
        let id = to_u32(self.plan.conds.len())?;
        self.plan.conds.push(CondPlan {
            literal,
            by_kind,
            targets,
            direct: false,
        });
        Ok(id)
    }

    /// What the key `name` reads for the kind with ordinal `ordinal`.
    fn target(
        &self,
        ordinal: usize,
        name: &str,
        test: &impl Fn(Option<ValueType>) -> Compiled,
    ) -> Result<(TargetKey, Target)> {
        let spec = self.spec(ordinal);
        if name.eq_ignore_ascii_case("style") {
            let prop = find_prop(spec, "style").map(|(index, _)| index);
            let styled = find_prop(spec, "styleName").map(|(index, _)| index);
            return Ok((
                TargetKey::Style(prop, styled),
                Target::Style {
                    prop,
                    name: styled,
                    test: test(None),
                },
            ));
        }
        let declared = find_prop(spec, name).or_else(|| {
            let canonical = spec.kind.name();
            KEY_ALIASES
                .iter()
                .find(|(family, element, alias, _)| {
                    *family == N::Kind::VOCABULARY.family
                        && *element == canonical
                        && alias.eq_ignore_ascii_case(name)
                })
                .and_then(|(_, _, _, key)| find_prop(spec, key))
        });
        if let Some((index, prop)) = declared {
            return Ok((
                TargetKey::Prop(index, prop.ty),
                Target::Prop {
                    index,
                    test: test(Some(prop.ty)),
                },
            ));
        }
        if name.eq_ignore_ascii_case("text")
            || (name.eq_ignore_ascii_case("value") && spec.kind.name() == VALUE_ELEMENT)
        {
            return Ok((TargetKey::Text, Target::Text(test(None))));
        }
        if name.eq_ignore_ascii_case("type") {
            return Ok((TargetKey::Type, Target::Type(test(None))));
        }
        for gap in N::SPEC.gaps {
            if let Subject::Property { element, key } = gap.subject {
                if element == spec.name && key.eq_ignore_ascii_case(name) {
                    return Err(gap.error().with_format(N::FORMAT));
                }
            }
        }
        Ok((TargetKey::Absent, Target::Absent(test(None))))
    }

    /// Records an `unknown_key` diagnostic, with a "did you mean" over the
    /// keys the selected kinds declare.
    fn unknown_key(&mut self, name: &str, kinds: u128) {
        let mut available: Vec<&'static str> = PSEUDO_KEYS.to_vec();
        for ordinal in ordinals(kinds) {
            for prop in self.spec(ordinal).props {
                if prop.access.get
                    && !available
                        .iter()
                        .any(|seen| seen.eq_ignore_ascii_case(prop.name))
                {
                    available.push(prop.name);
                }
            }
        }
        available.sort_by_key(|key| key.to_ascii_lowercase());
        let near = nearest_key(name, &available);
        let message = format!(
            "unknown key: no selected element declares it; only != can match; available: {}",
            available.join(", ")
        );
        self.plan.diagnostics.push(Diagnostic::new(
            DiagnosticKind::UnknownKey,
            name,
            message,
            near.map(|near| format!("did you mean '{near}'?")),
        ));
    }

    /// `:no-alt`: the `alt` property of each selected kind.
    fn no_alt(&mut self, kinds: u128) -> Result<Vec<Option<u16>>> {
        let mut table = vec![None; self.kind_count()];
        let mut any = false;
        for ordinal in ordinals(kinds) {
            if let Some((index, _)) = find_prop(self.spec(ordinal), "alt") {
                table[ordinal] = Some(index);
                any = true;
            }
        }
        if !any {
            for gap in N::SPEC.gaps {
                if let Subject::Property { element, key } = gap.subject {
                    if key == "alt"
                        && ordinals(kinds).any(|ordinal| self.spec(ordinal).name == element)
                    {
                        return Err(gap.error().with_format(N::FORMAT));
                    }
                }
            }
            self.plan.diagnostics.push(Diagnostic::new(
                DiagnosticKind::NeverMatches,
                ":no-alt",
                "':no-alt' selects elements without alternative text, and no selected element has an 'alt' property"
                    .to_owned(),
                Some("select pictures or shapes, e.g. picture:no-alt".to_owned()),
            ));
        }
        Ok(table)
    }

    /// OfficeCLI's `cell:has(formula)`: a `:has` whose list is one bare name
    /// that is no element but a key of the selected kinds is `[name]`.
    fn literal_has(&self, list: &[Complex<'q>], kinds: u128) -> Option<String> {
        let [complex] = list else {
            return None;
        };
        if !complex.rest().is_empty() {
            return None;
        }
        let compound = complex.first();
        if !compound.predicates().is_empty() || !compound.pseudos().is_empty() {
            return None;
        }
        let Element::Named(name) = compound.element() else {
            return None;
        };
        if N::Kind::VOCABULARY.canonical(name).is_some() {
            return None;
        }
        ordinals(kinds)
            .any(|ordinal| find_prop(self.spec(ordinal), name).is_some())
            .then(|| name.to_string())
    }

    /// Compiles a `:has`/`:not` selector list into chains.
    fn list(&mut self, list: &[Complex<'q>]) -> Result<(u32, u32)> {
        let mut ids = Vec::with_capacity(list.len());
        for complex in list {
            ids.push(self.chain(complex, Role::Nested)?);
        }
        let start = to_u32(self.plan.lists.len())?;
        let len = to_u32(ids.len())?;
        self.plan.lists.extend(ids);
        Ok((start, len))
    }

    fn and(&mut self, conds: Vec<u32>) -> Result<Option<u32>> {
        match conds.len() {
            0 => Ok(None),
            1 => Ok(Some(conds[0])),
            _ => {
                let (start, len) = self.args(conds)?;
                self.push_expr(ExprNode::And(start, len)).map(Some)
            },
        }
    }

    fn args(&mut self, ids: Vec<u32>) -> Result<(u32, u32)> {
        let start = to_u32(self.plan.expr_args.len())?;
        let len = to_u32(ids.len())?;
        self.plan.expr_args.extend(ids);
        Ok((start, len))
    }

    fn push_expr(&mut self, node: ExprNode) -> Result<u32> {
        let id = to_u32(self.plan.exprs.len())?;
        self.plan.exprs.push(node);
        Ok(id)
    }

    fn finish(mut self) -> Result<Plan<'q>> {
        let plan = &mut self.plan;
        let last = plan
            .top
            .iter()
            .any(|&chain| plan.chains[chain as usize].position == Some(Positional::Last));
        if last && plan.top.len() > 1 {
            return Err(invalid_selector(
                "[last()] on the selected element cannot be combined with ',' alternatives"
                    .to_owned(),
                Some(Cow::Borrowed("query each alternative separately")),
            ));
        }
        plan.wanted = plan
            .top
            .iter()
            .fold(0, |bits, &chain| bits | plan.chains[chain as usize].kinds);
        if let [only] = plan.top.as_slice() {
            let chain = plan.chains[*only as usize];
            if chain.len == 1 {
                let subject = plan.links[chain.first as usize].compound;
                plan.rows_mode = plan.compounds[subject as usize].header;
            }
        }
        plan.shallow = plan.top.iter().all(|&chain| {
            let chain = plan.chains[chain as usize];
            chain.len == 1 && plan.compounds[plan.links[chain.first as usize].compound as usize].any
        });
        plan.fast = super::plan::FastPlan::of(plan);
        Ok(self.plan)
    }
}

/// What a condition reads for one kind, for sharing targets between kinds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TargetKey {
    Absent,
    Prop(u16, ValueType),
    Text,
    Type,
    Style(Option<u16>, Option<u16>),
}

/// The index and spec of the property `name` (ASCII case-insensitive).
fn find_prop<K>(spec: &ElementSpec<K>, name: &str) -> Option<(u16, &'static PropSpec)> {
    spec.props
        .iter()
        .position(|prop| prop.access.get && prop.name.eq_ignore_ascii_case(name))
        .and_then(|index| Some((u16::try_from(index).ok()?, &spec.props[index])))
}

/// The key without a `col.` or `column.` prefix (ASCII case-insensitive),
/// when it has one.
fn strip_column_prefix(key: &str) -> Option<&str> {
    ["column.", "col."].into_iter().find_map(|prefix| {
        let head = key.get(..prefix.len())?;
        let rest = &key[prefix.len()..];
        (head.eq_ignore_ascii_case(prefix) && !rest.is_empty()).then_some(rest)
    })
}

fn is_row_property(key: &str) -> bool {
    ROW_PROPERTIES
        .iter()
        .any(|property| property.eq_ignore_ascii_case(key))
}

/// The leaf conditions of `expr`, in order.
fn collect_leaves<'e, 'q>(expr: &'e Expr<'q>, out: &mut Vec<&'e Condition<'q>>) {
    match expr {
        Expr::Condition(condition) => out.push(condition),
        Expr::And(parts) | Expr::Or(parts) => {
            for part in parts {
                collect_leaves(part, out);
            }
        },
        Expr::Not(inner) => collect_leaves(inner, out),
        _ => {},
    }
}

/// The key nearest to `input` by optimal-string-alignment distance within
/// `max(2, len / 3)`, also scoring a dotted key's last segment (OfficeCLI
/// `AttributeFilter.NearestMatch` with `matchKeySegment`); `None` on a tie.
fn nearest_key<'k>(input: &str, keys: &[&'k str]) -> Option<&'k str> {
    let wanted: Vec<char> = input.chars().map(|c| c.to_ascii_lowercase()).collect();
    let threshold = (wanted.len() / 3).max(2);
    let (mut best, mut best_distance, mut ties) = (None, usize::MAX, 0);
    for key in keys {
        let lower: Vec<char> = key.chars().map(|c| c.to_ascii_lowercase()).collect();
        if lower == wanted {
            continue;
        }
        let mut score = distance(&wanted, &lower);
        if let Some(dot) = lower.iter().rposition(|&c| c == '.') {
            if dot + 1 < lower.len() {
                score = score.min(distance(&wanted, &lower[dot + 1..]));
            }
        }
        if score > threshold {
            continue;
        }
        if score < best_distance {
            (best, best_distance, ties) = (Some(*key), score, 1);
        } else if score == best_distance {
            ties += 1;
        }
    }
    if ties == 1 { best } else { None }
}

fn set_position(slot: &mut Option<Positional>, position: Positional) -> Result<()> {
    if slot.is_some() {
        return Err(invalid_selector(
            "only one positional predicate ([N] or [last()]) may apply to the selected element"
                .to_owned(),
            Some(Cow::Borrowed(
                "keep one, or move the other onto an ancestor: tbl[2] > tr[1]",
            )),
        ));
    }
    *slot = Some(position);
    Ok(())
}

fn family_noun(family: Family) -> &'static str {
    match family {
        Family::Document => "word-processing documents",
        Family::Presentation => "presentations",
        Family::Workbook => "workbooks",
        _ => "other documents",
    }
}

fn invalid_selector(
    message: impl Into<Cow<'static, str>>,
    suggestion: Option<Cow<'static, str>>,
) -> Error {
    let error = Error::new(Code::InvalidSelector, message);
    match suggestion {
        Some(suggestion) => error.with_suggestion(suggestion),
        None => error,
    }
}

fn unsupported(what: &str) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        format!("{what} is not supported in queries"),
    )
}

fn internal(message: impl Into<Cow<'static, str>>) -> Error {
    Error::new(Code::InternalError, message)
}

fn to_u32(value: usize) -> Result<u32> {
    u32::try_from(value).map_err(|_| internal("a query plan table overflowed"))
}

fn to_u16(value: usize) -> Result<u16> {
    u16::try_from(value).map_err(|_| internal("a query plan index overflowed"))
}

fn to_u8(value: usize) -> Result<u8> {
    u8::try_from(value).map_err(|_| internal("a query plan kind index overflowed"))
}

/// The transitive closure of the child lists of `elements`: what can occur
/// below each kind, the walker's pruning table (record 0010 §5.3).
///
/// A spec table is static, so the closure is computed once per table and
/// kept for the process (record 0280: recomputing it was about a fifth of
/// compiling a short selector). Tables are told apart by address; there is
/// one per format adapter, so the cache stays a handful of entries.
fn reach_table<K: ElementKind>(elements: &'static [ElementSpec<K>]) -> &'static [u128] {
    type Tables = std::sync::RwLock<Vec<(usize, &'static [u128])>>;
    static TABLES: Tables = std::sync::RwLock::new(Vec::new());
    let key = elements.as_ptr().addr();
    if let Ok(tables) = TABLES.read()
        && let Some(&(_, reach)) = tables.iter().find(|(at, _)| *at == key)
    {
        return reach;
    }
    let mut reach = vec![0_u128; elements.len()];
    loop {
        let mut changed = false;
        for (index, element) in elements.iter().enumerate() {
            let mut below = reach[index];
            for child in element.children {
                let child = usize::from(child.ordinal());
                below |= bit_of(child) | reach.get(child).copied().unwrap_or(0);
            }
            if below != reach[index] {
                reach[index] = below;
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    let Ok(mut tables) = TABLES.write() else {
        // A poisoned cache only costs the recomputation.
        return Box::leak(reach.into_boxed_slice());
    };
    if let Some(&(_, existing)) = tables.iter().find(|(at, _)| *at == key) {
        return existing;
    }
    let reach: &'static [u128] = Box::leak(reach.into_boxed_slice());
    tables.push((key, reach));
    reach
}

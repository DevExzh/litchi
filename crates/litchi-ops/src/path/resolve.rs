//! The resolver: walks a path step by step through a [`Navigate`] adapter
//! (record 0010 §4.1).
//!
//! For each step it normalizes the name through the family vocabulary's
//! static alias table (§4.2), classifies the predicates (§4.3) into a
//! [`Resolved`] step, and asks the adapter for that one child. Nothing on the
//! success path allocates: names fold into a stack buffer, tables are
//! `static`, and the classified step borrows from the parsed path. A miss
//! from `get` becomes `path_not_found` with what exists, computed only then;
//! `find` returns `None` without computing it.

use litchi_core::selector::{CellRange, Condition, Expr, Position, Predicate, Step, column_number};

use super::arg::{Origin, PathArg};
use super::navigate::{ElementKind, Identities, Lookup, Navigate, Resolved, Select};
use super::report::{self, MAX_COLUMN, MAX_ROW, Spot};
use super::vocabulary::{self, Bare, ElementRule, IndexRule, Named};
use crate::{Family, Result};

/// Resolves an absolute path from the document root; a miss is
/// `path_not_found`.
pub(crate) fn get<'f, N: Navigate<'f>>(nav: &N, path: &PathArg<'_>) -> Result<N::Node> {
    let result = fenced(nav, || {
        walk(nav, path, None).and_then(|outcome| found(nav, path, outcome))
    });
    nav.check_source()?;
    result
}

/// Resolves an absolute path from the document root; a miss is `Ok(None)`.
pub(crate) fn find<'f, N: Navigate<'f>>(nav: &N, path: &PathArg<'_>) -> Result<Option<N::Node>> {
    let result = fenced(nav, || walk(nav, path, None).map(Outcome::into_found));
    nav.check_source()?;
    result
}

/// Resolves a relative path from `start`, a node of kind `kind`; an absolute
/// path starts at the document root instead. A miss is `path_not_found`.
pub(crate) fn get_from<'f, N: Navigate<'f>>(
    nav: &N,
    start: N::Node,
    kind: N::Kind,
    path: &PathArg<'_>,
) -> Result<N::Node> {
    let result = fenced(nav, || {
        walk(nav, path, Some((start, kind))).and_then(|outcome| found(nav, path, outcome))
    });
    nav.check_source()?;
    result
}

/// As [`get_from`], with a miss as `Ok(None)`.
pub(crate) fn find_from<'f, N: Navigate<'f>>(
    nav: &N,
    start: N::Node,
    kind: N::Kind,
    path: &PathArg<'_>,
) -> Result<Option<N::Node>> {
    let result = fenced(nav, || {
        walk(nav, path, Some((start, kind))).map(Outcome::into_found)
    });
    nav.check_source()?;
    result
}

/// Runs `read` with the adapter's nested source checks covered by the
/// caller's [`Navigate::check_source`] after it: a source that changed
/// before or during `read` fails that check, whose error wins over what
/// `read` returned, so one check after the whole resolution gives what a
/// check around each nested read gave (record 0262).
fn fenced<'f, N: Navigate<'f>, T>(nav: &N, read: impl FnOnce() -> T) -> T {
    let _scope = super::fence::Scope::enter(nav.fence_key());
    read()
}

/// Where a walk ended.
enum Outcome<'s, N, K> {
    /// Every step resolved.
    Found(N),
    /// Step `step` selected nothing under `parent`.
    Missed {
        step: usize,
        parent: N,
        wanted: Resolved<'s, K>,
    },
}

impl<N, K> Outcome<'_, N, K> {
    fn into_found(self) -> Option<N> {
        match self {
            Self::Found(node) => Some(node),
            Self::Missed { .. } => None,
        }
    }
}

fn found<'f, N: Navigate<'f>>(
    nav: &N,
    path: &PathArg<'_>,
    outcome: Outcome<'_, N::Node, N::Kind>,
) -> Result<N::Node> {
    match outcome {
        Outcome::Found(node) => Ok(node),
        Outcome::Missed {
            step,
            parent,
            wanted,
        } => {
            let spot = Spot::new(path, step, N::FORMAT);
            Err(report::not_found(nav, &spot, &parent, &wanted))
        },
    }
}

fn walk<'s, 'f, N: Navigate<'f>>(
    nav: &N,
    path: &'s PathArg<'_>,
    start: Option<(N::Node, N::Kind)>,
) -> Result<Outcome<'s, N::Node, N::Kind>> {
    let vocabulary = N::Kind::VOCABULARY;
    let steps = path.path().steps();
    if steps.len() > litchi_core::selector::limits::MAX_STEPS {
        return Err(crate::Error::new(
            crate::Code::MaxDepthExceeded,
            "path exceeds the step limit",
        ));
    }
    super::arg::validate(path)?;
    if matches!(path.origin(), Origin::Sheet { .. }) && vocabulary.family != Family::Workbook {
        return Err(report::no_sheets(
            &Spot::new(path, 0, N::FORMAT),
            vocabulary,
        ));
    }
    // OfficeCLI's Word and Excel handlers read `body/p[1]` as `/body/p[1]`;
    // Litchi refuses relative text at the document level, as the grammar's
    // `Path::parse` does, and suggests the rooted form.
    let (mut node, mut kind) = if path.is_absolute() {
        (nav.root()?, N::Kind::ROOT)
    } else {
        match start {
            Some(start) => start,
            None => return Err(report::relative(path, N::FORMAT)),
        }
    };
    if steps.is_empty() {
        return Ok(Outcome::Found(node));
    }
    // The rule of the current node's element; each step looks up its child's
    // rule once and carries it to the next step.
    let mut parent = rule(&Spot::new(path, 0, N::FORMAT), kind)?;
    for (index, step) in steps.iter().enumerate() {
        let spot = Spot::new(path, index, N::FORMAT);
        // Children named by free text: worksheets at a workbook's root are
        // tried first; styles under `/styles` are the only reading.
        let mut named_miss = None;
        match parent.named {
            Named::Only { child, case } => {
                // OfficeCLI accepts `/styles/Heading1[1]` (the first style
                // with that id); the name alone selects one style, so Litchi
                // refuses the predicate instead of reading it.
                if !step.predicates().is_empty() {
                    return Err(report::named_with_predicates(&spot, child));
                }
                let wanted = Resolved {
                    kind: model::<N>(&spot, child)?,
                    select: Select::Name {
                        name: step.name(),
                        case,
                    },
                };
                match nav.child(&node, &wanted)? {
                    Lookup::Found(next) => {
                        (node, kind, parent) = (next, wanted.kind, rule(&spot, wanted.kind)?);
                        continue;
                    },
                    Lookup::Missing => {
                        return Ok(Outcome::Missed {
                            step: index,
                            parent: node,
                            wanted,
                        });
                    },
                    answer => return Err(report::lookup(nav, &spot, &node, &wanted, answer)),
                }
            },
            Named::First { child, case } if step.predicates().is_empty() => {
                if let Some(child) = N::Kind::from_name(child) {
                    let wanted = Resolved {
                        kind: child,
                        select: Select::Name {
                            name: step.name(),
                            case,
                        },
                    };
                    match nav.child(&node, &wanted)? {
                        Lookup::Found(next) => {
                            (node, kind, parent) = (next, wanted.kind, rule(&spot, wanted.kind)?);
                            continue;
                        },
                        Lookup::Missing => named_miss = Some(wanted),
                        answer => return Err(report::lookup(nav, &spot, &node, &wanted, answer)),
                    }
                }
            },
            Named::First { .. } | Named::No => {},
        }
        let Some((wanted, child)) =
            classify::<N>(&spot, path, index, step, parent, kind == N::Kind::ROOT)?
        else {
            return match named_miss {
                Some(wanted) => Ok(Outcome::Missed {
                    step: index,
                    parent: node,
                    wanted,
                }),
                None => Err(report::unknown(&spot, step.name(), vocabulary)),
            };
        };
        match nav.child(&node, &wanted)? {
            Lookup::Found(next) => (node, kind, parent) = (next, wanted.kind, child),
            Lookup::Missing => {
                return Ok(Outcome::Missed {
                    step: index,
                    parent: node,
                    wanted,
                });
            },
            answer => return Err(report::lookup(nav, &spot, &node, &wanted, answer)),
        }
    }
    Ok(Outcome::Found(node))
}

/// The vocabulary rule of an adapter kind.
fn rule<K: ElementKind>(spot: &Spot<'_, '_>, kind: K) -> Result<&'static ElementRule> {
    K::VOCABULARY
        .rule(kind.name())
        .ok_or_else(|| report::kind_without_rule(spot, kind.name()))
}

/// The adapter's kind for a canonical name: `unsupported_path` when the
/// format lacks the concept, `unsupported_capability` when the adapter does
/// not model it yet.
fn model<'f, N: Navigate<'f>>(spot: &Spot<'_, '_>, canonical: &str) -> Result<N::Kind> {
    if let Some(kind) = N::Kind::from_name(canonical) {
        return Ok(kind);
    }
    if N::ABSENT
        .iter()
        .any(|absent| absent.eq_ignore_ascii_case(canonical))
    {
        Err(report::absent(spot, canonical, N::Kind::VOCABULARY))
    } else {
        Err(report::unmodeled(spot, canonical))
    }
}

/// Classifies a step as an element kind and a selection, with the rule of
/// that kind, or `None` when no vocabulary knows its name.
fn classify<'s, 'f, N: Navigate<'f>>(
    spot: &Spot<'_, '_>,
    path: &PathArg<'_>,
    index: usize,
    step: &'s Step<'s>,
    parent: &ElementRule,
    parent_is_root: bool,
) -> Result<Option<(Resolved<'s, N::Kind>, &'static ElementRule)>> {
    let vocabulary = N::Kind::VOCABULARY;
    let name = step.name();
    if parent.cells && step.predicates().is_empty() {
        if let Some(range) = CellRange::parse(name) {
            let single = !name.contains(':');
            let kind = model::<N>(spot, if single { "cell" } else { "range" })?;
            let select = if single {
                Select::Cell(range.start())
            } else {
                Select::Range(range)
            };
            return Ok(Some((Resolved { kind, select }, rule(spot, kind)?)));
        }
    }
    let Some((canonical, rule)) = vocabulary.canonical_rule(name) else {
        if let Some(other) = vocabulary::elsewhere(vocabulary, name) {
            return Err(report::foreign(spot, name, vocabulary, other));
        }
        if vocabulary.family != Family::Workbook
            && report::addresses_cells(path.path().steps(), index)
        {
            return Err(report::no_sheets(spot, vocabulary));
        }
        return Ok(None);
    };
    let kind = model::<N>(spot, canonical)?;
    // The alias table's rule is the kind's own when the adapter names its
    // kind canonically, as every adapter does; otherwise look it up.
    let rule = if kind.name() == canonical {
        rule
    } else {
        self::rule(spot, kind)?
    };
    let select = select(spot, step, rule, parent_is_root, path.raw)?;
    Ok(Some((Resolved { kind, select }, rule)))
}

/// Classifies a step's predicates for its element (record 0010 §4.3).
fn select<'s>(
    spot: &Spot<'_, '_>,
    step: &'s Step<'s>,
    rule: &ElementRule,
    parent_is_root: bool,
    raw: bool,
) -> Result<Select<'s>> {
    let predicates = step.predicates();
    match predicates {
        [] => return Ok(Select::None),
        [Predicate::Index(position)] => return index(spot, rule, *position, parent_is_root),
        [Predicate::Last] => return Ok(Select::Last),
        [single] => {
            if let Some(token) = single.bare_key() {
                return bare(spot, rule, token);
            }
        },
        _ => {},
    }
    // Everything else must be one or more identities, all of which hold.
    for predicate in predicates {
        if let Some((key, _)) = predicate.identity() {
            let Some(identity) = rule.identity(key) else {
                return Err(report::not_identity(spot, rule, key));
            };
            if identity.is_native() && !raw {
                return Err(crate::Error::new(
                    crate::Code::InvalidPath,
                    "native identity requires an explicit raw path",
                )
                .with_suggestion("use path::raw::path for native identity addressing"));
            }
            continue;
        }
        if let Predicate::Filter(Expr::Condition(Condition::Exists(key))) = predicate {
            if key.is_attribute() {
                return Err(report::identity_without_value(spot, rule, key.name()));
            }
        }
        let positional = matches!(predicate, Predicate::Index(_) | Predicate::Last)
            || predicate.bare_key().is_some();
        return Err(if positional {
            report::stacked(spot)
        } else {
            report::content_filter(spot)
        });
    }
    Ok(Select::Identity(Identities::new(
        predicates,
        rule.identities,
    )))
}

/// What `[N]` selects for this element.
fn index<'s>(
    spot: &Spot<'_, '_>,
    rule: &ElementRule,
    position: Position,
    parent_is_root: bool,
) -> Result<Select<'s>> {
    let number = || {
        position
            .one_based()
            .ok_or_else(|| report::index_too_large(spot))
    };
    match rule.index {
        IndexRule::Ordinal => Ok(Select::Ordinal(position)),
        IndexRule::NumberUnderRoot if !parent_is_root => Ok(Select::Ordinal(position)),
        IndexRule::Number | IndexRule::NumberUnderRoot => number().map(Select::Number),
        IndexRule::Row => match number()? {
            row if row > MAX_ROW => Err(report::out_of_grid(spot, "row", row, MAX_ROW)),
            row => Ok(Select::Number(row)),
        },
        IndexRule::Column => match number()? {
            column if column > MAX_COLUMN => {
                Err(report::out_of_grid(spot, "column", column, MAX_COLUMN))
            },
            column => Ok(Select::Number(column)),
        },
    }
}

/// What a bare token selects for this element.
fn bare<'s>(spot: &Spot<'_, '_>, rule: &ElementRule, token: &'s str) -> Result<Select<'s>> {
    match rule.bare {
        Some(Bare::Column) => match column_number(token) {
            Some(column) => Ok(Select::Number(column as usize)),
            None => Err(report::bad_column(spot, token)),
        },
        Some(Bare::Name(case)) => Ok(Select::Name { name: token, case }),
        Some(Bare::Token) => Ok(Select::Token(token)),
        None => Err(report::bare_token(spot, rule, token)),
    }
}

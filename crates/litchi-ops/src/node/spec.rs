//! Static spec tables: one per format, read by the node view, path
//! resolution, query compilation and `capabilities()` (record 0010 §7.2).
//!
//! Each format declares one `static FormatSpec<Kind>` whose `ElementSpec`s sit
//! at their kind's index. The node view's `format` member is produced from an
//! element's `PropSpec` table; resolution normalizes step names through the
//! alias lists; the query walker prunes with the child lists; and
//! `capabilities()` reports the same entries, so the report cannot claim
//! what the code does not do. M2 widens `PropSpec` into a registry entry
//! (constraints, coercion, typed setter) through its access flags, without
//! changing these shapes.
//!
//! Drift is caught at compile time by [`check_format_spec!`], which a format
//! invokes once on its table: each element must sit at its kind's index, alias
//! lists must be sorted lowercase ASCII and unique across the format, property
//! keys must be unique per element, every child must list its parent, and
//! every gap must carry a gap code and a follow-up.

// The format adapters declare the tables; a build without any adapter
// declares none, and only this crate's tests read every part of them.
#![cfg_attr(
    not(test),
    allow(
        dead_code,
        unused_imports,
        unused_macros,
        reason = "read through the format adapters' tables, which a build may not have"
    )
)]

use std::fmt;

use crate::error::{Code, Error};
use crate::value::ValueType;

/// One readable property of an element kind: its key, value type and the
/// verbs that read and write it.
///
/// M1's `PropSpec` is the read-only subset of an M2 registry entry (record
/// 0010 §7.3). The seam where the registry's write side attaches is
/// [`PropSpec::access`]: M2 builds entries that also carry constraints,
/// coercion and the typed setter, and turns the `set` and `add` flags on, so
/// `capabilities()` reports them without changing its shape or this table's
/// readers (resolution, the query compiler and the node view read `name`
/// and `ty` only).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct PropSpec {
    /// The OfficeCLI property name, which is also the key `set` accepts in
    /// M2 (`bold`, `font.latin`, `x`, `cachedValue`).
    pub(crate) name: &'static str,
    /// The value type: how a query literal or a `set` value for the key is
    /// coerced ([`ValueType::coerce`]) and how the key's value prints (a
    /// length in its [`LengthSpec`](crate::value::LengthSpec)'s style, so one
    /// [`Length`](crate::value::Length) prints the way OfficeCLI prints that
    /// key; record 0010 §2.3).
    pub(crate) ty: ValueType,
    /// Which verbs read and write the key: [`Access::READ`] in M1.
    pub(crate) access: Access,
    /// Native metadata available only through explicit raw operations.
    pub(crate) raw: bool,
    /// The registry's write declaration (constraint, clearing, pending
    /// setter), which [`crate::registry::Entry::readable`] copies; the
    /// capability report and the schema export read it from here.
    pub(crate) admission: crate::registry::Admission,
}

impl PropSpec {
    /// Mark a property as native metadata without making it a path identity.
    #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
    pub(crate) const fn raw(mut self) -> Self {
        self.raw = true;
        self
    }

    /// A readable property with the given key and value type.
    pub(crate) const fn new(name: &'static str, ty: ValueType) -> Self {
        Self {
            name,
            ty,
            access: Access::READ,
            raw: false,
            admission: crate::registry::Admission::NONE,
        }
    }

    /// Mark this property as accepting `null` to clear its value.
    pub(crate) const fn clearable(mut self) -> Self {
        self.admission = self.admission.clearable();
        self
    }
}

/// Which verbs read and write a property: the `get`, `set` and `add` flags
/// of OfficeCLI's help schema (record 0010 §7.1).
///
/// M1 reads only. M2's registry entries set `set` and `add` where the write
/// kernel has a typed setter (record 0010 §7.3); `capabilities()` reports
/// the flags as they are.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) struct Access {
    /// `get` and `query` read the key.
    pub(crate) get: bool,
    /// `set` writes it (M2).
    pub(crate) set: bool,
    /// `add` accepts it when it creates the element (M2).
    pub(crate) add: bool,
}

impl Access {
    /// Read only: every M1 property.
    pub(crate) const READ: Self = Self {
        get: true,
        set: false,
        add: false,
    };
}

/// One element kind of a format: its OfficeCLI type, names, structure,
/// identity keys, path templates, readable properties and capability gap.
#[derive(Debug)]
pub(crate) struct ElementSpec<K: 'static> {
    /// The kind; `FormatSpec::elements[kind as usize].kind == kind`.
    pub(crate) kind: K,
    /// The OfficeCLI `type` string (`paragraph`, `cell`, `slide`).
    pub(crate) name: &'static str,
    /// Every step or selector name for the kind, lowercase ASCII and sorted;
    /// input is matched ASCII case-insensitively (record 0010 §4.2).
    pub(crate) aliases: &'static [&'static str],
    /// The kinds that may contain this one.
    pub(crate) parents: &'static [K],
    /// The kinds this one may contain; also the query walker's pruning table.
    pub(crate) children: &'static [K],
    /// The `@key`s a path step may use to address the kind by identity.
    pub(crate) identities: &'static [&'static str],
    /// Canonical path templates (`/body/p[N]`, `/body/p[@paraId=X]`).
    pub(crate) paths: &'static [&'static str],
    /// The readable properties, in the order the view emits them.
    pub(crate) props: &'static [PropSpec],
    /// `Some` when the format has the kind but the leaf cannot answer it yet:
    /// resolving it is `unsupported_capability`, listed in `capabilities()`.
    pub(crate) gap: Option<&'static GapSpec>,
}

impl<K> ElementSpec<K> {
    /// The index and spec of the property named `key` (exact match; keys are
    /// case-sensitive, as OfficeCLI's are on output).
    pub(crate) fn prop(&self, key: &str) -> Option<(usize, &'static PropSpec)> {
        self.props
            .iter()
            .enumerate()
            .find(|(_, prop)| prop.name == key)
    }
}

/// The spec table of one format.
#[derive(Debug)]
pub(crate) struct FormatSpec<K: 'static> {
    /// The kind of the document root (`/`).
    pub(crate) root: K,
    /// Every element kind, each at its kind's index.
    pub(crate) elements: &'static [ElementSpec<K>],
    /// Axes and properties OfficeCLI has for the format that Litchi does not
    /// answer, beyond the per-element `gap`s.
    pub(crate) gaps: &'static [GapSpec],
}

/// What a capability gap is about: an element, a path axis, one property or
/// operation context of an element, or a view mode.
///
/// ```
/// use litchi_ops::capabilities::Subject;
///
/// let subject = Subject::Property { element: "cell", key: "computedValue" };
/// assert_eq!(subject.to_string(), "the `computedValue` property of `cell`");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Subject {
    /// An element kind (`comment`).
    Element(&'static str),
    /// A path axis (`/header`).
    Axis(&'static str),
    /// One property key of an element kind.
    Property {
        /// The element's OfficeCLI type.
        element: &'static str,
        /// The property key.
        key: &'static str,
    },
    /// One operation on an element in a context the handler cannot admit.
    ///
    /// `verb` is the canonical lower-case spelling from
    /// [`crate::capabilities::Verb`].  The strings are static because gap
    /// subjects are part of format specification tables and must remain
    /// allocation-free during capability discovery.
    Operation {
        /// The unified operation (`set`, `add`, `remove`, ...).
        verb: &'static str,
        /// The element kind whose operation is conditional.
        element: &'static str,
        /// The selection or leaf context that is refused.
        context: &'static str,
    },
    /// A `view` mode (`html`).
    View(&'static str),
}

impl fmt::Display for Subject {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Element(name) => write!(f, "the `{name}` element"),
            Self::Axis(axis) => write!(f, "the `{axis}` axis"),
            Self::Property { element, key } => {
                write!(f, "the `{key}` property of `{element}`")
            },
            Self::Operation {
                verb,
                element,
                context,
            } => write!(f, "the `{verb}` operation on `{element}` in `{context}`"),
            Self::View(mode) => write!(f, "the `{mode}` view"),
        }
    }
}

/// An element, axis, property, operation context or view mode the format lacks
/// or the leaf cannot answer yet, with the follow-up that closes it (record
/// 0010 §4.4, §7.1).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct GapSpec {
    /// What is missing.
    pub(crate) subject: Subject,
    /// [`Code::UnsupportedPath`] when the format lacks the concept,
    /// [`Code::UnsupportedCapability`] when the leaf has no model yet, or
    /// [`Code::DocumentProtected`] when the operation would violate file
    /// protection.
    pub(crate) code: Code,
    /// The leaf request or milestone that closes the gap (`L-DOCX-R4`, `M6`).
    pub(crate) follow_up: &'static str,
    /// One sentence on why the gap exists.
    pub(crate) reason: &'static str,
}

/// Path resolution reports a gap through its follow-up (M1 step R2).
impl crate::path::navigate::GapInfo for GapSpec {
    fn follow_up(&self) -> &'static str {
        self.follow_up
    }
}

impl GapSpec {
    /// The typed error a request that hits this gap returns. Built on the
    /// error path only.
    pub(crate) fn error(&self) -> Error {
        let message = match self.code {
            Code::UnsupportedPath => format!("{} does not exist in this format", self.subject),
            _ => format!("{} is not supported yet", self.subject),
        };
        Error::new(self.code, message)
            .with_suggestion(format!("{} (follow-up: {})", self.reason, self.follow_up))
    }
}

/// Byte-wise comparison usable in constant evaluation: negative, zero or
/// positive as `a` sorts before, equal to or after `b`.
const fn compare(a: &[u8], b: &[u8]) -> i32 {
    let mut i = 0;
    while i < a.len() && i < b.len() {
        if a[i] != b[i] {
            return if a[i] < b[i] { -1 } else { 1 };
        }
        i += 1;
    }
    if a.len() == b.len() {
        0
    } else if a.len() < b.len() {
        -1
    } else {
        1
    }
}

/// Whether `names` are non-empty lowercase ASCII and strictly ascending, so a
/// lowercased input can be binary-searched (the `Axis::TABLE` pattern).
pub(crate) const fn aliases_valid(names: &[&str]) -> bool {
    let mut i = 0;
    while i < names.len() {
        let bytes = names[i].as_bytes();
        if bytes.is_empty() {
            return false;
        }
        let mut j = 0;
        while j < bytes.len() {
            if !bytes[j].is_ascii() || bytes[j].is_ascii_uppercase() {
                return false;
            }
            j += 1;
        }
        if i > 0 && compare(names[i - 1].as_bytes(), bytes) >= 0 {
            return false;
        }
        i += 1;
    }
    true
}

/// Whether two alias lists share a name.
pub(crate) const fn aliases_overlap(a: &[&str], b: &[&str]) -> bool {
    let mut i = 0;
    while i < a.len() {
        let mut j = 0;
        while j < b.len() {
            if compare(a[i].as_bytes(), b[j].as_bytes()) == 0 {
                return true;
            }
            j += 1;
        }
        i += 1;
    }
    false
}

/// Whether every key in `props` is non-empty and unique.
pub(crate) const fn props_valid(props: &[PropSpec]) -> bool {
    let mut i = 0;
    while i < props.len() {
        if props[i].name.is_empty() {
            return false;
        }
        let mut j = 0;
        while j < i {
            if compare(props[i].name.as_bytes(), props[j].name.as_bytes()) == 0 {
                return false;
            }
            j += 1;
        }
        i += 1;
    }
    true
}

/// Whether every gap carries a gap code and a follow-up.
pub(crate) const fn gaps_valid(gaps: &[GapSpec]) -> bool {
    let mut i = 0;
    while i < gaps.len() {
        if !gap_valid(&gaps[i]) {
            return false;
        }
        i += 1;
    }
    true
}

/// Whether one gap carries a gap code and a follow-up.
pub(crate) const fn gap_valid(gap: &GapSpec) -> bool {
    let subject_valid = match gap.subject {
        Subject::Element(name) | Subject::Axis(name) | Subject::View(name) => !name.is_empty(),
        Subject::Property { element, key } => !element.is_empty() && !key.is_empty(),
        Subject::Operation {
            verb,
            element,
            context,
        } => !verb.is_empty() && !element.is_empty() && !context.is_empty(),
    };
    let code_valid = match gap.code {
        Code::UnsupportedPath | Code::UnsupportedCapability => true,
        Code::DocumentProtected => matches!(gap.subject, Subject::Operation { .. }),
        _ => false,
    };
    subject_valid && code_valid && !gap.follow_up.is_empty() && !gap.reason.is_empty()
}

/// Checks a format's spec table at compile time.
///
/// `$spec` is a `static` or `const` [`FormatSpec`] whose kind type is a
/// field-less enum (the check casts kinds to `usize`). The macro expands to an
/// anonymous `const` item, so a table that breaks a rule fails the build:
///
/// - each element sits at its kind's index, and the root kind is in range;
/// - each element has an OfficeCLI type name;
/// - each alias list is sorted lowercase ASCII, and no alias names two kinds;
/// - each element's property keys are non-empty and unique;
/// - each listed child names the element among its parents;
/// - each gap carries a gap code, a follow-up and a reason.
macro_rules! check_format_spec {
    ($spec:path) => {
        const _: () = {
            let spec = &$spec;
            let elements = spec.elements;
            assert!(!elements.is_empty(), "a format spec has no elements");
            assert!(
                (spec.root as usize) < elements.len(),
                "the root kind has no element spec"
            );
            let mut i = 0;
            while i < elements.len() {
                let element = &elements[i];
                assert!(
                    element.kind as usize == i,
                    "an element spec is not at its kind's index"
                );
                assert!(!element.name.is_empty(), "an element has no type name");
                assert!(
                    $crate::node::spec::aliases_valid(element.aliases),
                    "an alias list is not sorted lowercase ASCII"
                );
                assert!(
                    $crate::node::spec::props_valid(element.props),
                    "an element repeats a property key"
                );
                if let Some(gap) = element.gap {
                    assert!(
                        $crate::node::spec::gap_valid(gap),
                        "an element gap lacks a gap code, follow-up or reason"
                    );
                }
                let mut other = 0;
                while other < i {
                    assert!(
                        !$crate::node::spec::aliases_overlap(
                            element.aliases,
                            elements[other].aliases
                        ),
                        "an alias names two kinds"
                    );
                    other += 1;
                }
                let mut c = 0;
                while c < element.children.len() {
                    let child = &elements[element.children[c] as usize];
                    let mut listed = false;
                    let mut p = 0;
                    while p < child.parents.len() {
                        if child.parents[p] as usize == i {
                            listed = true;
                        }
                        p += 1;
                    }
                    assert!(listed, "a child does not list its parent");
                    c += 1;
                }
                i += 1;
            }
            assert!(
                $crate::node::spec::gaps_valid(spec.gaps),
                "a format gap lacks a gap code, follow-up or reason"
            );
        };
    };
}

pub(crate) use check_format_spec;

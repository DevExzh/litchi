//! One property declaration for read projection and strict write admission.
//!
//! Tables retain declaration order for node views, and build sorted lookup
//! permutations at compile time. Setter enums belong to the format handler;
//! lookup and coercion introduce no dynamic dispatch or document model.

// Write handlers attach in M2. Read tables already use the same declarations.
#![cfg_attr(
    not(test),
    allow(dead_code, reason = "write admission is consumed as M2 handlers land")
)]

use crate::error::{Code, Error, Result};
use crate::node::spec::{Access as Flags, GapSpec, PropSpec};
use crate::props::PropValue;
use crate::value::{Length, Scalar, ValueType};

#[derive(Clone, Copy)]
pub(crate) enum Access<S> {
    ReadOnly,
    Set(S),
    Add(S),
    SetAndAdd(S),
    Gap(&'static GapSpec),
}

/// The domain a property's value must lie in, beyond its [`ValueType`]:
/// the registry checks it after coercion, before a setter runs, and the
/// schema export ([`crate::schema()`]) reports it.
///
/// ```
/// use litchi_core::FileFormat;
/// use litchi_ops::capabilities::Constraint;
///
/// let schema = litchi_ops::schema(FileFormat::Xlsx);
/// # #[cfg(feature = "xlsx")] {
/// let row = schema.elements().find(|element| element.name() == "row").unwrap();
/// let level = row.property("outline").unwrap();
/// assert_eq!(level.constraint(), Constraint::Int { min: 0, max: 7 });
/// # }
/// ```
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum Constraint {
    /// No constraint beyond the value type.
    None,
    /// An integer in `min..=max`.
    Int {
        /// The smallest admitted value.
        min: i64,
        /// The largest admitted value.
        max: i64,
    },
    /// A finite number in `min..=max`.
    Number {
        /// The smallest admitted value.
        min: f64,
        /// The largest admitted value.
        max: f64,
    },
    /// A length in `min..=max`.
    Length {
        /// The shortest admitted length.
        min: Length,
        /// The longest admitted length.
        max: Length,
    },
    /// Text of at most `max` characters (Unicode scalar values).
    Chars {
        /// The longest admitted text, in characters.
        max: u32,
    },
    /// Text that is not empty.
    NotEmpty,
}

/// Bounds compare by bit pattern, so equality is total; constructors reject
/// non-finite number bounds.
impl PartialEq for Constraint {
    fn eq(&self, other: &Self) -> bool {
        match (*self, *other) {
            (Self::None, Self::None) | (Self::NotEmpty, Self::NotEmpty) => true,
            (Self::Int { min, max }, Self::Int { min: a, max: b }) => min == a && max == b,
            (Self::Number { min, max }, Self::Number { min: a, max: b }) => {
                min.to_bits() == a.to_bits() && max.to_bits() == b.to_bits()
            },
            (Self::Length { min, max }, Self::Length { min: a, max: b }) => min == a && max == b,
            (Self::Chars { max }, Self::Chars { max: b }) => max == b,
            _ => false,
        }
    }
}

impl Eq for Constraint {}

/// What the registry declares about a property's write admission beyond its
/// access flags: carried by every [`PropSpec`] a table projects, so the
/// capability report and the schema export read the same declaration the
/// write path checks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Admission {
    /// The value domain checked after coercion.
    pub(crate) constraint: Constraint,
    /// Whether `null` clears the property.
    pub(crate) clear: bool,
    /// The gap a write of the property reports, when the registry declares
    /// the property but its setter is pending.
    pub(crate) gap: Option<&'static GapSpec>,
}

impl Admission {
    /// No declaration: a read-only property or one outside any table.
    pub(crate) const NONE: Self = Self {
        constraint: Constraint::None,
        clear: false,
        gap: None,
    };

    /// Mark a writable property as accepting `null` to clear its value.
    #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
    pub(crate) const fn clearable(self) -> Self {
        Self {
            clear: true,
            ..self
        }
    }
}

/// A registry alias with the canonical property name it resolves to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct PropAlias {
    /// The accepted spelling (`font.strike`).
    pub(crate) alias: &'static str,
    /// The canonical property it names (`strike`).
    pub(crate) name: &'static str,
}

#[derive(Clone, Copy)]
pub(crate) enum Operation {
    Set,
    Add,
}

#[derive(Clone, Copy)]
pub(crate) struct Entry<S> {
    pub(crate) spec: PropSpec,
    pub(crate) access: Access<S>,
    pub(crate) constraint: Constraint,
    pub(crate) clear: bool,
}

impl<S: Copy> Entry<S> {
    pub(crate) const fn read(spec: PropSpec) -> Self {
        Self {
            spec,
            access: Access::ReadOnly,
            constraint: Constraint::None,
            clear: false,
        }
    }

    pub(crate) const fn readable(self) -> PropSpec {
        let (set, add) = match self.access {
            Access::Set(_) => (true, false),
            Access::Add(_) => (false, true),
            Access::SetAndAdd(_) => (true, true),
            Access::ReadOnly | Access::Gap(_) => (false, false),
        };
        PropSpec {
            access: Flags {
                get: self.spec.access.get,
                set,
                add,
            },
            admission: Admission {
                constraint: self.constraint,
                clear: self.clear,
                gap: match self.access {
                    Access::Gap(gap) => Some(gap),
                    _ => None,
                },
            },
            ..self.spec
        }
    }

    fn setter(&self, operation: Operation) -> Result<S> {
        match (self.access, operation) {
            (Access::Set(setter), Operation::Set)
            | (Access::Add(setter), Operation::Add)
            | (Access::SetAndAdd(setter), _) => Ok(setter),
            (Access::Gap(gap), _) => Err(Error::new(gap.code, gap.reason)
                .with_suggestion(gap.follow_up)
                .with_property(self.spec.name)),
            _ => Err(Error::new(
                Code::UnsupportedProperty,
                "the property is not writable by this operation",
            )
            .with_property(self.spec.name)),
        }
    }

    pub(crate) fn prepare<T>(
        &self,
        operation: Operation,
        value: &PropValue<'_>,
        stage: impl FnOnce(S, Option<Scalar<'_>>) -> Result<T>,
    ) -> Result<T> {
        let setter = self.setter(operation)?;
        if matches!(value, PropValue::Null) {
            return if self.clear {
                stage(setter, None)
            } else {
                Err(
                    Error::new(Code::InvalidValue, "the property cannot be cleared")
                        .with_property(self.spec.name),
                )
            };
        }
        value
            .with_coerced(self.spec.ty, |value| {
                let value = value.ok_or_else(|| {
                    Error::new(
                        Code::InternalError,
                        "non-null property coercion returned no value",
                    )
                })?;
                self.constraint
                    .check(value)
                    .map_err(|error| error.with_property(self.spec.name))?;
                stage(setter, Some(value))
            })
            .map_err(|error| error.with_property(self.spec.name))
    }
}

impl Constraint {
    const fn valid(self, ty: ValueType) -> bool {
        match (self, ty) {
            (Self::None, _) => true,
            (Self::Int { min, max }, ValueType::Int) => min <= max,
            (Self::Number { min, max }, ValueType::Number) => {
                min.is_finite() && max.is_finite() && min <= max
            },
            (Self::Length { min, max }, ValueType::Length(_)) => min.emu() <= max.emu(),
            (Self::Chars { .. } | Self::NotEmpty, ValueType::Text | ValueType::Token(_)) => true,
            _ => false,
        }
    }

    fn check(self, value: Scalar<'_>) -> Result<()> {
        let valid = match (self, value) {
            (Self::None, _) => true,
            (Self::Int { min, max }, Scalar::Int(value)) => min <= value && value <= max,
            (Self::Number { min, max }, Scalar::Number(value)) => {
                value.is_finite() && min <= value && value <= max
            },
            (Self::Length { min, max }, Scalar::Length(value)) => {
                min.emu() <= value.emu() && value.emu() <= max.emu()
            },
            (Self::Chars { max }, Scalar::Text(value) | Scalar::Token(value)) => {
                value.chars().count() as u64 <= u64::from(max)
            },
            (Self::NotEmpty, Scalar::Text(value) | Scalar::Token(value)) => !value.is_empty(),
            _ => false,
        };
        if valid {
            Ok(())
        } else {
            Err(Error::new(
                Code::InvalidValue,
                "the property value is outside its supported domain",
            ))
        }
    }
}

#[derive(Clone, Copy)]
pub(crate) struct Alias {
    pub(crate) name: &'static str,
    pub(crate) entry: usize,
}

pub(crate) struct Table<S: 'static, const N: usize, const A: usize> {
    entries: [Entry<S>; N],
    aliases: [Alias; A],
    order: [usize; N],
    alias_order: [usize; A],
}

impl<S: Copy + 'static, const N: usize, const A: usize> Table<S, N, A> {
    pub(crate) const fn new(entries: [Entry<S>; N], aliases: [Alias; A]) -> Self {
        let mut names = [""; N];
        let mut i = 0;
        while i < N {
            assert!(
                entries[i].constraint.valid(entries[i].spec.ty),
                "property constraint does not match type or bounds"
            );
            names[i] = entries[i].spec.name;
            i += 1;
        }
        let order = sorted(names);
        let mut alias_names = [""; A];
        i = 0;
        while i < A {
            assert!(aliases[i].entry < N, "property alias has no entry");
            alias_names[i] = aliases[i].name;
            let mut j = 0;
            while j < N {
                assert!(
                    compare(aliases[i].name, names[j]) != 0,
                    "alias shadows a canonical property"
                );
                j += 1;
            }
            i += 1;
        }
        let alias_order = sorted(alias_names);
        Self {
            entries,
            aliases,
            order,
            alias_order,
        }
    }

    pub(crate) const fn readable(&self) -> [PropSpec; N] {
        let mut props = [PropSpec::new("", ValueType::Text); N];
        let mut i = 0;
        while i < N {
            props[i] = self.entries[i].readable();
            i += 1;
        }
        props
    }

    /// Every alias with the canonical name it resolves to, in declaration
    /// order: the schema export's per-property `aliases`.
    #[cfg_attr(
        not(any(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "pptx")),
        allow(dead_code, reason = "only these formats' tables export aliases")
    )]
    pub(crate) const fn prop_aliases(&self) -> [PropAlias; A] {
        let mut aliases = [PropAlias {
            alias: "",
            name: "",
        }; A];
        let mut i = 0;
        while i < A {
            aliases[i] = PropAlias {
                alias: self.aliases[i].name,
                name: self.entries[self.aliases[i].entry].spec.name,
            };
            i += 1;
        }
        aliases
    }

    pub(crate) fn lookup(&self, key: &str) -> Option<(usize, &Entry<S>)> {
        if let Ok(position) = self
            .order
            .binary_search_by(|&index| compare(self.entries[index].spec.name, key).cmp(&0))
        {
            let index = self.order[position];
            return Some((index, &self.entries[index]));
        }
        let position = self
            .alias_order
            .binary_search_by(|&index| compare(self.aliases[index].name, key).cmp(&0))
            .ok()?;
        let index = self.aliases[self.alias_order[position]].entry;
        Some((index, &self.entries[index]))
    }

    pub(crate) fn resolve(&self, key: &str) -> Result<(usize, &Entry<S>)> {
        self.lookup(key).ok_or_else(|| {
            let error = Error::new(
                Code::UnsupportedProperty,
                "the element has no such property",
            );
            let candidates = self
                .entries
                .iter()
                .map(|entry| (entry.spec.name, entry.spec.name))
                .chain(
                    self.aliases
                        .iter()
                        .map(|alias| (alias.name, self.entries[alias.entry].spec.name)),
                );
            match crate::value::suggest::nearest_keyed(key, candidates) {
                Some(name) => error.with_suggestion(format!("did you mean '{name}'?")),
                None => error,
            }
        })
    }
}

const fn sorted<const N: usize>(names: [&str; N]) -> [usize; N] {
    let mut order = [0; N];
    let mut i = 0;
    while i < N {
        assert!(!names[i].is_empty(), "property names cannot be empty");
        let mut j = 0;
        while j < names[i].len() {
            assert!(
                names[i].as_bytes()[j].is_ascii(),
                "property names must be ASCII"
            );
            j += 1;
        }
        order[i] = i;
        j = i;
        while j > 0 && compare(names[order[j - 1]], names[order[j]]) > 0 {
            order.swap(j - 1, j);
            j -= 1;
        }
        i += 1;
    }
    i = 1;
    while i < N {
        assert!(
            compare(names[order[i - 1]], names[order[i]]) != 0,
            "duplicate property spelling"
        );
        i += 1;
    }
    order
}

const fn compare(a: &str, b: &str) -> i8 {
    let (a, b) = (a.as_bytes(), b.as_bytes());
    let mut i = 0;
    while i < a.len() && i < b.len() {
        let (left, right) = (a[i].to_ascii_lowercase(), b[i].to_ascii_lowercase());
        if left < right {
            return -1;
        }
        if left > right {
            return 1;
        }
        i += 1;
    }
    if a.len() < b.len() {
        -1
    } else if a.len() > b.len() {
        1
    } else {
        0
    }
}

/// Generate the readable table from its write declarations in the same order.
#[cfg(any(test, feature = "docx"))]
macro_rules! registry {
    ($vis:vis $name:ident, $table:ident : $setter:ty { $($entry:expr),* $(,)? }) => {
        $vis const $table: $crate::registry::Table<$setter, { [$(stringify!($entry)),*].len() }, 0> =
            $crate::registry::Table::new([$($entry),*], []);
        $vis const $name: &[$crate::node::spec::PropSpec] = &$table.readable();
    };
}
#[cfg(any(test, feature = "docx"))]
pub(crate) use registry;

#[cfg(test)]
mod tests;

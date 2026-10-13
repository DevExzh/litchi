//! Closed vocabularies (enumeration properties).

use super::error::{ErrorKind, ValueError, What};
use super::text::{find_folded, is_lower_ascii, str_lt};

/// A closed vocabulary such as an `underline` or `align` property's values:
/// canonical tokens plus alias spellings, matched ASCII case-insensitively
/// after trimming white space, as OfficeCLI's `ToLowerInvariant()` switches
/// match them.
///
/// The tables are `static` and checked when the vocabulary is built: in a
/// `static` or `const` initializer, an alias table that is not sorted,
/// not lower case, or points outside `values` fails to compile. Lookups are
/// allocation-free binary searches; "did you mean" suggestions run only on
/// the error path.
///
/// ```
/// use litchi_ops::value::Tokens;
///
/// static ALIGN: Tokens = Tokens::new(
///     "alignment",
///     &["left", "center", "right", "justify"],
///     // Aliases, sorted by their lower-case spelling, with the index of the
///     // canonical value each one means.
///     &[("centre", 1), ("justified", 3), ("middle", 1)],
/// );
///
/// assert_eq!(ALIGN.parse(" Center ")?, "center");
/// assert_eq!(ALIGN.parse("middle")?, "center");
/// let error = ALIGN.parse("rihgt").unwrap_err();
/// assert_eq!(error.suggestion().as_deref(), Some("did you mean 'right'?"));
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, PartialEq, Eq)]
pub struct Tokens {
    name: &'static str,
    values: &'static [&'static str],
    aliases: &'static [(&'static str, u16)],
}

impl Tokens {
    /// Builds a vocabulary named `name` (used in messages) from its canonical
    /// `values` and its `aliases`, given as `(lower-case spelling, index into
    /// values)` sorted by spelling. Every canonical value is also accepted in
    /// any ASCII case without being listed as an alias.
    ///
    /// # Panics
    ///
    /// Panics, at compile time in a `static` or `const` initializer, when
    /// `values` is empty, an alias spelling is not lower-case ASCII, the
    /// aliases are not strictly sorted, or an alias index is out of range.
    #[must_use]
    pub const fn new(
        name: &'static str,
        values: &'static [&'static str],
        aliases: &'static [(&'static str, u16)],
    ) -> Self {
        assert!(!values.is_empty(), "a vocabulary needs at least one value");
        let mut index = 0;
        while index < aliases.len() {
            let (spelling, target) = aliases[index];
            assert!(
                is_lower_ascii(spelling),
                "alias spellings must be lower-case ASCII"
            );
            assert!((target as usize) < values.len(), "alias index out of range");
            if index > 0 {
                assert!(
                    str_lt(aliases[index - 1].0, spelling),
                    "aliases must be sorted by spelling without duplicates"
                );
            }
            index += 1;
        }
        Self {
            name,
            values,
            aliases,
        }
    }

    /// The vocabulary's name, as messages print it.
    #[must_use]
    pub const fn name(&self) -> &'static str {
        self.name
    }

    /// The canonical values, in declaration order.
    #[must_use]
    pub const fn values(&self) -> &'static [&'static str] {
        self.values
    }

    /// The accepted input spellings besides the canonical values, each with
    /// the index of the value it means, sorted by spelling.
    pub(crate) const fn alias_pairs(&self) -> &'static [(&'static str, u16)] {
        self.aliases
    }

    /// The canonical value `text` means, if any.
    #[must_use]
    pub fn get(&self, text: &str) -> Option<&'static str> {
        let text = text.trim();
        if let Some(value) = self
            .values
            .iter()
            .find(|value| value.eq_ignore_ascii_case(text))
        {
            return Some(value);
        }
        find_folded(self.aliases, |(spelling, _)| spelling, text)
            .map(|&(_, target)| self.values[usize::from(target)])
    }

    /// Parses `text` to its canonical value.
    ///
    /// # Errors
    ///
    /// Returns an `invalid_value` [`ValueError`] listing the values, with the
    /// nearest spelling as its suggestion, when `text` is not in the
    /// vocabulary.
    pub fn parse<'a>(&'static self, text: &'a str) -> Result<&'static str, ValueError<'a>> {
        self.get(text).ok_or_else(|| {
            let kind = if text.trim().is_empty() {
                ErrorKind::Empty
            } else {
                ErrorKind::UnknownToken
            };
            ValueError::new(What::Token(self), kind, text)
        })
    }

    /// Every accepted spelling with the canonical value it means: the values
    /// themselves, then the aliases.
    pub(super) fn spellings(&self) -> impl Iterator<Item = (&'static str, &'static str)> {
        let values = self.values;
        values.iter().map(|&value| (value, value)).chain(
            self.aliases
                .iter()
                .map(move |&(spelling, target)| (spelling, values[usize::from(target)])),
        )
    }
}

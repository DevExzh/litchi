//! Formula values and stored calculation caches.

/// Shared-formula expansion and reference translation.
pub mod shared;

use std::fmt;

use crate::cell::{Text, Value};
use crate::error::{Result, invalid};

const MAX_FORMULA_CHARACTERS: usize = 8_192;

/// Formula expression together with its semantic storage kind and cached value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    text: Box<str>,
    kind: Kind,
    cached: Option<Cache>,
}

impl Formula {
    /// Create an ordinary formula with no cached result.
    ///
    /// The expression never includes a leading `=`.
    pub fn new(text: impl Into<Box<str>>) -> Result<Self> {
        let text = text.into();
        if text.trim().is_empty() || text.trim_start().starts_with('=') {
            return Err(invalid(
                "formula expression must be non-empty and omit the leading '='",
            ));
        }
        if let Some(character) = text
            .chars()
            .find(|&character| !is_xml_10_character(character))
        {
            return Err(invalid(format!(
                "formula contains XML 1.0-forbidden character U+{:04X}",
                character as u32
            )));
        }
        if text.chars().count() > MAX_FORMULA_CHARACTERS {
            return Err(invalid(format!(
                "formula exceeds {MAX_FORMULA_CHARACTERS} characters"
            )));
        }
        Ok(Self {
            text,
            kind: Kind::Scalar,
            cached: None,
        })
    }

    pub(crate) fn parsed(text: String, kind: Kind, cached: Option<Cache>) -> Self {
        Self {
            text: text.into_boxed_str(),
            kind,
            cached,
        }
    }

    /// Formula expression without the leading `=`.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Semantic formula kind.
    pub fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Stored or calculated result, when one is available.
    pub fn cached(&self) -> Option<&Cache> {
        self.cached.as_ref()
    }
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || (character >= '\u{20}' && character <= '\u{D7FF}')
        || (character >= '\u{E000}' && character <= '\u{FFFD}')
        || (character >= '\u{10000}' && character <= '\u{10FFFF}')
}

/// Semantic formula form.
///
/// Shared-formula records are expanded during parsing and therefore never
/// appear here.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// An ordinary scalar formula.
    Scalar,
    /// An array formula anchored at this cell.
    Array {
        /// Inclusive A1 range exactly as stored, when supplied.
        range: Option<Text>,
    },
    /// A what-if analysis data-table formula.
    DataTable {
        /// Inclusive A1 range exactly as stored, when supplied.
        range: Option<Text>,
    },
    /// A producer extension not yet modeled semantically.
    Unknown(Text),
}

/// One formula result cache with explicit provenance and freshness.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    value: Value,
    origin: Origin,
    freshness: Freshness,
}

impl Cache {
    pub(crate) fn stored(value: Value) -> Self {
        Self {
            value,
            origin: Origin::Stored,
            freshness: Freshness::Unknown,
        }
    }

    /// Cached result value.
    pub fn value(&self) -> &Value {
        &self.value
    }

    /// How the result entered this snapshot.
    pub fn origin(&self) -> Origin {
        self.origin
    }

    /// Whether the cache is known to match current dependencies.
    pub fn freshness(&self) -> Freshness {
        self.freshness
    }
}

/// Provenance of a formula result.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Origin {
    /// Read from the document package.
    Stored,
    /// Produced by an explicitly selected calculation engine.
    Calculated,
}

/// Knowledge about whether a formula result is current.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Freshness {
    /// No dependency proof is available.
    Unknown,
    /// The result is proven current for this snapshot.
    Current,
    /// A dependency changed after the result was produced.
    Stale,
}

impl fmt::Display for Formula {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.text)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn explicit_formula_construction_is_checked() {
        let formula = Formula::new("SUM(A1:A3)").expect("valid formula");
        assert_eq!(formula.text(), "SUM(A1:A3)");
        assert!(Formula::new("").is_err());
        assert!(Formula::new("  \t").is_err());
        assert!(Formula::new("=SUM(A1:A3)").is_err());
        assert!(Formula::new("x".repeat(MAX_FORMULA_CHARACTERS + 1)).is_err());
    }

    #[test]
    fn explicit_formula_construction_rejects_xml_forbidden_characters() {
        assert!(Formula::new("SUM(A1,\u{1})").is_err());
        assert!(Formula::new("IF(A1=\"\t\",1,0)").is_ok());
        assert!(Formula::new("SUM(😀,A1)").is_ok());
    }
}

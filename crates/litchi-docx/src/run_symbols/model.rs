//! Semantic values for the Word 2015 `symEx` run extension.

use crate::error::Result;

use super::validation;

/// Maximum number of Unicode scalar values retained for one symbol font name.
pub const MAX_FONT_CHARS: usize = 255;
/// Maximum number of `symEx` elements retained in one run snapshot.
pub const MAX_SYMBOLS: usize = 1024;

/// A Word 2015 extended symbol character.
///
/// Both attributes are optional in `CT_SymEx`.  Consequently, `None` means
/// that the corresponding attribute was absent, while `Some("")` or
/// `Some(0)` remains an explicit authored value rather than being collapsed
/// into absence.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[must_use]
pub struct Symbol {
    font: Option<String>,
    character: Option<u32>,
}

impl Symbol {
    /// Construct an extended symbol with both schema attributes present.
    pub fn new(font: impl Into<String>, character: u32) -> Result<Self> {
        Self::from_parts(Some(font.into()), Some(character))
    }

    /// Construct an extended symbol from its optional wire attributes.
    pub fn from_parts(font: Option<String>, character: Option<u32>) -> Result<Self> {
        let value = Self { font, character };
        value.validate()?;
        Ok(value)
    }

    /// Construct a symbol with only its Unicode character code.
    pub fn with_character(character: u32) -> Result<Self> {
        Self::from_parts(None, Some(character))
    }

    /// Construct an explicitly empty `symEx` element.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            font: None,
            character: None,
        }
    }

    /// Return the optional font attribute.
    #[inline]
    #[must_use]
    pub fn font(&self) -> Option<&str> {
        self.font.as_deref()
    }

    /// Return the optional Unicode scalar value encoded by `char`.
    #[inline]
    #[must_use]
    pub const fn character(&self) -> Option<u32> {
        self.character
    }

    /// Alias for [`Self::character`] that follows the wire attribute name.
    #[inline]
    #[must_use]
    pub const fn char_code(&self) -> Option<u32> {
        self.character()
    }

    /// Set or remove the optional font attribute.
    pub fn set_font(&mut self, font: Option<impl Into<String>>) -> Result<&mut Self> {
        let candidate = Self {
            font: font.map(Into::into),
            character: self.character,
        };
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    /// Set or remove the optional Unicode scalar value.
    pub fn set_character(&mut self, character: Option<u32>) -> Result<&mut Self> {
        let candidate = Self {
            font: self.font.clone(),
            character,
        };
        candidate.validate()?;
        *self = candidate;
        Ok(self)
    }

    /// Alias for [`Self::set_character`].
    pub fn set_char_code(&mut self, character: Option<u32>) -> Result<&mut Self> {
        self.set_character(character)
    }

    /// Validate the complete semantic symbol value.
    pub fn validate(&self) -> Result<()> {
        validation::validate_symbol(self)
    }

    pub(crate) fn font_value(&self) -> Option<&str> {
        self.font()
    }

    pub(crate) const fn character_value(&self) -> Option<u32> {
        self.character()
    }
}

/// Ordered `symEx` elements in one run.
///
/// The empty collection represents an absent extension.  An element with no
/// attributes is represented by `Some(Symbol::empty())` at the singular run
/// facade, preserving the distinction between an absent element and an
/// explicitly authored empty element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Symbols {
    values: Vec<Symbol>,
}

impl Symbols {
    /// Construct an empty symbol collection.
    #[must_use]
    pub const fn new() -> Self {
        Self { values: Vec::new() }
    }

    /// Construct and validate a symbol collection.
    pub fn from_values(values: Vec<Symbol>) -> Result<Self> {
        let value = Self { values };
        value.validate()?;
        Ok(value)
    }

    /// Number of symbols in source order.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Whether no `symEx` elements are present.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Borrow symbols in their authored order.
    #[inline]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Symbol> {
        self.values.iter()
    }

    /// Return a symbol by zero-based source position.
    #[inline]
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Symbol> {
        self.values.get(index)
    }

    /// Return the first symbol, if present.
    #[inline]
    #[must_use]
    pub fn first(&self) -> Option<&Symbol> {
        self.values.first()
    }

    /// Append a symbol after enforcing collection and value bounds.
    pub fn push(&mut self, symbol: Symbol) -> Result<&mut Self> {
        let next_len =
            self.values.len().checked_add(1).ok_or_else(|| {
                crate::Error::Invalid("Word run symbol count overflows usize".into())
            })?;
        if next_len > MAX_SYMBOLS {
            return Err(crate::Error::Invalid(format!(
                "Word run symbols exceed {MAX_SYMBOLS} elements"
            )));
        }
        symbol.validate()?;
        self.values.push(symbol);
        Ok(self)
    }

    /// Remove all symbols and return the previous values.
    pub fn clear(&mut self) {
        self.values.clear();
    }

    /// Validate the complete collection.
    pub fn validate(&self) -> Result<()> {
        validation::validate_symbols(self)
    }

    pub(crate) fn values(&self) -> &[Symbol] {
        &self.values
    }
}

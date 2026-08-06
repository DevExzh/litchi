//! Run and writer facades for the `symEx` semantic owner.

use crate::error::Result;
use crate::paragraph::Run;
use crate::writer::{MutableRun, RunContent};

use super::codec;
use super::{Symbol, Symbols};

impl Run {
    /// Read all direct Word 2015 `symEx` elements in source order.
    pub fn symbols(&self) -> Result<Symbols> {
        codec::read(self.xml_bytes())
    }

    /// Read the first direct symbol, preserving an absent element as `None`.
    pub fn symbol(&self) -> Result<Option<Symbol>> {
        Ok(self.symbols()?.first().cloned())
    }

    /// Alias for [`Self::symbol`] using the specification element name.
    pub fn sym_ex(&self) -> Result<Option<Symbol>> {
        self.symbol()
    }

    /// Replace the complete direct symbol collection without disturbing other
    /// run content or unknown XML.
    pub fn set_symbols(&mut self, value: Symbols) -> Result<&mut Self> {
        let original = self.xml_bytes();
        let rewritten = codec::rewrite(original, &value)?;
        if rewritten.as_slice() != original {
            self.replace_xml(rewritten);
        }
        Ok(self)
    }

    /// Replace the direct symbol with zero or one `symEx` element.
    pub fn set_symbol(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        let mut symbols = Symbols::new();
        if let Some(value) = value {
            symbols.push(value)?;
        }
        self.set_symbols(symbols)
    }

    /// Alias for [`Self::set_symbol`] using the specification element name.
    pub fn set_sym_ex(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        self.set_symbol(value)
    }

    /// Remove all direct `symEx` elements.
    pub fn clear_symbols(&mut self) -> Result<&mut Self> {
        self.set_symbols(Symbols::new())
    }

    /// Remove the direct symbol extension.
    pub fn clear_symbol(&mut self) -> Result<&mut Self> {
        self.clear_symbols()
    }
}

impl MutableRun {
    /// Borrow the symbol authored on this new run, if any.
    #[inline]
    #[must_use]
    pub fn symbol(&self) -> Option<&Symbol> {
        match &self.content {
            RunContent::Symbol(value) => Some(value),
            _ => None,
        }
    }

    /// Alias for [`Self::symbol`] using the specification element name.
    #[inline]
    #[must_use]
    pub fn sym_ex(&self) -> Option<&Symbol> {
        self.symbol()
    }

    /// Set or remove the single symbol authored by this mutable run.
    pub fn set_symbol(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        if let Some(value) = &value {
            value.validate()?;
        }
        self.content = match value {
            Some(value) => RunContent::Symbol(value),
            None => RunContent::Text(String::new()),
        };
        Ok(self)
    }

    /// Alias for [`Self::set_symbol`] using the specification element name.
    pub fn set_sym_ex(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        self.set_symbol(value)
    }

    /// Remove the authored symbol from this run.
    pub fn clear_symbol(&mut self) -> &mut Self {
        self.content = RunContent::Text(String::new());
        self
    }

    /// Set one symbol from a font name and Unicode scalar value.
    pub fn set_symbol_character(
        &mut self,
        font: impl Into<String>,
        character: u32,
    ) -> Result<&mut Self> {
        self.set_symbol(Some(Symbol::new(font, character)?))
    }
}

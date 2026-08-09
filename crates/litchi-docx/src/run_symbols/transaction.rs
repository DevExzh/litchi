#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Failure-atomic snapshots, edits, commits, and reversible patches for runs.

use std::sync::Arc;

use crate::error::{Error, Result};

use super::codec;
use super::model::{Symbol, Symbols};
use super::validation::validate_symbols;

/// An immutable, cheaply clonable run snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    symbols: Symbols,
}

impl Snapshot {
    /// Parse and retain a bounded source-preserving `w:r` snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let symbols = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            symbols,
        })
    }

    /// Borrow the original run XML without copying it.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow all direct symbols in source order.
    #[inline]
    pub const fn symbols(&self) -> &Symbols {
        &self.symbols
    }

    /// Return the first direct symbol, preserving absence as `None`.
    #[inline]
    #[must_use]
    pub fn symbol(&self) -> Option<&Symbol> {
        self.symbols.first()
    }

    /// Alias for [`Self::symbol`].
    #[inline]
    #[must_use]
    pub fn sym_ex(&self) -> Option<&Symbol> {
        self.symbol()
    }

    /// Start an isolated edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.symbols.clone(),
        }
    }
}

/// A run edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Symbols,
}

impl Transaction {
    /// Borrow all symbols projected by this transaction.
    #[inline]
    pub const fn symbols(&self) -> &Symbols {
        &self.next
    }

    /// Return the first symbol projected by this transaction.
    #[inline]
    #[must_use]
    pub fn symbol(&self) -> Option<&Symbol> {
        self.next.first()
    }

    /// Replace the complete ordered symbol collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_symbols(&mut self, value: Symbols) -> Result<&mut Self> {
        validate_symbols(&value)?;
        self.next = value;
        Ok(self)
    }

    /// Replace the run's direct symbol with zero or one element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_symbol(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        let mut symbols = Symbols::new();
        if let Some(value) = value {
            symbols.push(value)?;
        }
        self.set_symbols(symbols)
    }

    /// Alias for [`Self::set_symbol`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_sym_ex(&mut self, value: Option<Symbol>) -> Result<&mut Self> {
        self.set_symbol(value)
    }

    /// Remove all direct `symEx` elements.
    pub fn clear_symbols(&mut self) -> &mut Self {
        self.next.clear();
        self
    }

    /// Remove the direct symbol collection.
    pub fn clear_symbol(&mut self) -> &mut Self {
        self.clear_symbols()
    }

    /// Validate and publish the edit without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        let xml = codec::rewrite(self.base.xml_bytes(), &self.next)?;
        let snapshot = Snapshot::from_xml(xml)?;
        Ok(Commit {
            patch: Patch {
                before: self.base.symbols,
                after: self.next,
            },
            snapshot,
        })
    }
}

/// A successful publication containing a new snapshot and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A lineage-independent reversible symbol collection patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Symbols,
    after: Symbols,
}

impl Patch {
    /// Borrow the expected source symbols.
    #[inline]
    pub const fn before(&self) -> &Symbols {
        &self.before
    }

    /// Borrow the symbols produced by this patch.
    #[inline]
    pub const fn after(&self) -> &Symbols {
        &self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply the patch only when the target has the expected source state.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.symbols != self.before {
            return Err(Error::InvalidFormat(
                "symEx patch source state does not match its precondition".into(),
            ));
        }
        let xml = codec::rewrite(source.xml_bytes(), &self.after)?;
        Snapshot::from_xml(xml)
    }
}

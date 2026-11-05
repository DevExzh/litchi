//! Source-checked protection XML transactions.

use litchi_core::Result;

#[cfg(test)]
use super::Key;
use super::codec;
use super::model::Policy;
use super::{Kind, invalid};

/// A staged protection-policy edit against one immutable XML source.
#[derive(Debug, Clone)]
pub(crate) struct Transaction {
    source: Vec<u8>,
    kind: Kind,
    before: Policy,
    current: Policy,
}

impl Transaction {
    /// Start an edit against package `settings.xml` bytes.
    pub(crate) fn package(source: impl AsRef<[u8]>) -> Result<Self> {
        Self::new(source.as_ref(), Kind::Package)
    }

    /// Start an edit against a flat `OpenDocument` XML source.
    pub(crate) fn flat(source: impl AsRef<[u8]>) -> Result<Self> {
        Self::new(source.as_ref(), Kind::Flat)
    }

    fn new(source: &[u8], kind: Kind) -> Result<Self> {
        let before = codec::parse(source, kind)?;
        Ok(Self {
            source: source.to_vec(),
            kind,
            current: before.clone(),
            before,
        })
    }

    /// Replace the complete staged policy after validation.
    pub(crate) fn set(&mut self, policy: Policy) -> Result<()> {
        policy.validate()?;
        self.current = policy;
        Ok(())
    }

    #[cfg(test)]
    /// Stage a bookmark-protection toggle or clear it with `None`.
    pub(crate) fn set_bookmarks(&mut self, value: Option<bool>) {
        self.current.bookmarks = value;
    }

    #[cfg(test)]
    /// Stage a read-only loading hint or clear it with `None`.
    pub(crate) fn set_read_only(&mut self, value: Option<bool>) {
        self.current.read_only = value;
    }

    #[cfg(test)]
    /// Stage tracked-change protection digest material or clear it with `None`.
    pub(crate) fn set_redline_key(&mut self, value: Option<Key>) -> Result<()> {
        if let Some(key) = &value {
            Key::new(key.as_bytes().to_vec())?;
        }
        self.current.redline_key = value;
        Ok(())
    }

    /// Commit the staged edit without mutating the source bytes.
    pub(crate) fn commit(self) -> Result<Commit> {
        let xml = codec::rewrite(&self.source, self.kind, &self.before, &self.current)?;
        let after = codec::parse(&xml, self.kind)?;
        if after != self.current {
            return invalid("protection transaction did not round-trip its staged policy");
        }
        Ok(Commit {
            patch: Patch {
                source: self.source.clone(),
                target: xml.clone(),
                kind: self.kind,
                before: self.before.clone(),
                after: after.clone(),
            },
            source: self.source,
            xml,
            before: self.before,
            after,
        })
    }
}

/// The immutable result of a successful protection transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Commit {
    patch: Patch,
    source: Vec<u8>,
    xml: Vec<u8>,
    before: Policy,
    after: Policy,
}

impl Commit {
    #[cfg(test)]
    /// Borrow the committed XML bytes.
    pub(crate) fn as_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return committed XML bytes, consuming the commit.
    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.xml
    }

    #[cfg(test)]
    /// Whether the transaction left the XML source byte-for-byte unchanged.
    pub(crate) fn is_unchanged(&self) -> bool {
        self.source == self.xml
    }
}

/// An exact-source, reversible protection patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Patch {
    source: Vec<u8>,
    target: Vec<u8>,
    kind: Kind,
    before: Policy,
    after: Policy,
}

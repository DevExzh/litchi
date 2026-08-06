//! Source-checked protection XML transactions.

use litchi_core::Result;

use super::codec;
use super::model::Policy;
use super::{Kind, invalid};

/// A staged protection-policy edit against one immutable XML source.
#[derive(Debug, Clone)]
pub struct Transaction {
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

    /// Start an edit against a flat OpenDocument XML source.
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

    /// Return the policy currently staged in this transaction.
    pub fn policy(&self) -> &Policy {
        &self.current
    }

    /// Replace the complete staged policy after validation.
    pub fn set(&mut self, policy: Policy) -> Result<()> {
        policy.validate()?;
        self.current = policy;
        Ok(())
    }

    /// Stage a form-protection toggle or clear it with `None`.
    pub fn set_forms(&mut self, value: Option<bool>) {
        self.current.forms = value;
    }

    /// Stage a bookmark-protection toggle or clear it with `None`.
    pub fn set_bookmarks(&mut self, value: Option<bool>) {
        self.current.bookmarks = value;
    }

    /// Stage a read-only loading hint or clear it with `None`.
    pub fn set_read_only(&mut self, value: Option<bool>) {
        self.current.read_only = value;
    }

    /// Stage tracked-change protection digest material or clear it with `None`.
    pub fn set_redline_key(&mut self, value: Option<super::Key>) -> Result<()> {
        if let Some(key) = &value {
            super::Key::new(key.as_bytes().to_vec())?;
        }
        self.current.redline_key = value;
        Ok(())
    }

    /// Commit the staged edit without mutating the source bytes.
    pub fn commit(self) -> Result<Commit> {
        let xml = codec::rewrite(&self.source, self.kind, &self.before, &self.current)?;
        let after = codec::parse(&xml, self.kind)?;
        if after != self.current {
            return invalid("protection transaction did not round-trip its staged policy");
        }
        Ok(Commit {
            source: self.source,
            xml,
            before: self.before,
            after,
        })
    }
}

/// The immutable result of a successful protection transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    source: Vec<u8>,
    xml: Vec<u8>,
    before: Policy,
    after: Policy,
}

impl Commit {
    /// Return the source policy observed when the transaction began.
    pub fn before(&self) -> &Policy {
        &self.before
    }

    /// Return the policy represented by the committed XML.
    pub fn after(&self) -> &Policy {
        &self.after
    }

    /// Borrow the committed XML bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return committed XML bytes, consuming the commit.
    pub fn into_bytes(self) -> Vec<u8> {
        self.xml
    }

    /// Whether the transaction left the XML source byte-for-byte unchanged.
    pub fn is_unchanged(&self) -> bool {
        self.source == self.xml
    }
}

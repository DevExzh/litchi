//! Immutable transform snapshots and detached, failure-atomic edits.

use std::sync::Arc;

use crate::Error;
use crate::Result;

use super::{Angle, Limits, Point, Size, Transform, codec, validation};

pub use crate::LimitKind;
pub use litchi_core::edit::Diagnostics;

/// An immutable, cheaply clonable source snapshot of one `a:xfrm` element.
///
/// The original bytes are replayed for a no-op commit, retaining producer
/// prefixes, attribute order, whitespace, and lexical defaults exactly.
/// Changed commits are rebuilt from the fully typed transform model.
#[derive(Debug, Clone)]
pub struct Package {
    xml: Arc<[u8]>,
    value: Transform,
    limits: Limits,
}

impl Package {
    /// Parse and retain a bounded transform fragment.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn from_xml(xml: impl Into<Vec<u8>>, limits: &Limits) -> Result<Self> {
        let xml = xml.into();
        let limits = limits.validate()?;
        let value = codec::read(&xml, &limits)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
            limits,
        })
    }

    /// Create a canonical snapshot from a detached transform value.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn new(value: Transform, limits: &Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let xml = codec::write(&value, &limits)?;
        drop(value);
        Self::from_xml(xml, &limits)
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed transform value.
    #[inline]
    #[must_use]
    pub const fn value(&self) -> &Transform {
        &self.value
    }

    /// Start an isolated edit based on this immutable snapshot.
    #[inline]
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            working: self.value.clone(),
        }
    }
}

/// A transform edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Package,
    working: Transform,
}

impl Edit {
    /// Borrow the projected typed transform.
    #[inline]
    #[must_use]
    pub const fn transform(&self) -> &Transform {
        &self.working
    }

    /// Replace the complete typed transform after validating its output
    /// budget.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set(&mut self, value: Transform) -> Result<&mut Self> {
        validation::validate(&value)?;
        self.working = value;
        Ok(self)
    }

    /// Set or clear the object offset.
    #[inline]
    pub fn set_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.working.set_offset(value);
        self
    }

    /// Set or clear the object extent.
    #[inline]
    pub fn set_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.working.set_extent(value);
        self
    }

    /// Set or clear the group child-coordinate offset.
    #[inline]
    pub fn set_child_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.working.set_child_offset(value);
        self
    }

    /// Set or clear the group child-coordinate extent.
    #[inline]
    pub fn set_child_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.working.set_child_extent(value);
        self
    }

    /// Set or clear the authored rotation.
    #[inline]
    pub fn set_rotation(&mut self, value: Option<Angle>) -> &mut Self {
        self.working.set_rotation(value);
        self
    }

    /// Set or clear the authored horizontal flip.
    #[inline]
    pub fn set_flip_horizontal(&mut self, value: Option<bool>) -> &mut Self {
        self.working.set_flip_horizontal(value);
        self
    }

    /// Set or clear the authored vertical flip.
    #[inline]
    pub fn set_flip_vertical(&mut self, value: Option<bool>) -> &mut Self {
        self.working.set_flip_vertical(value);
        self
    }

    /// Validate and publish the edit without changing the source package.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn commit(self) -> Result<Commit> {
        validation::validate(&self.working)?;
        let xml = if self.base.value == self.working {
            self.base.xml.as_ref().to_vec()
        } else {
            codec::write(&self.working, &self.base.limits)?
        };
        let changed = self.base.value != self.working;
        let package = Package::from_xml(xml, &self.base.limits)?;
        let patch = Patch {
            before: self.base.value,
            after: self.working,
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
    }
}

/// A successful transform publication containing a new package and a
/// reversible semantic patch.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A lineage-independent, preconditioned reversible transform patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Transform,
    after: Transform,
}

impl Patch {
    /// Borrow the expected source state.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Transform {
        &self.before
    }

    /// Borrow the state produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Transform {
        &self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }

    /// Whether this patch leaves the semantic transform unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Apply the patch only when the target has the expected semantic state.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    fn apply(&self, source: &Package) -> Result<Package> {
        if source.value != self.before {
            return Err(Error::PatchConflict);
        }
        let xml = if self.before == self.after {
            source.xml.as_ref().to_vec()
        } else {
            codec::write(&self.after, &source.limits)?
        };
        Package::from_xml(xml, &source.limits)
    }
}

impl Package {
    /// Replay a patch against this exact package and publish a commit triad.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.apply(self)?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

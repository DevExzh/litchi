use std::fmt;

use litchi_iwa_graph::ObjectId;

/// One directed object dependency.
///
/// Both endpoints are validated [`ObjectId`] values. The optional constructor
/// is provided for native adapters where a zero/null reference is represented
/// as an absent field; ordinary CRUD APIs use the typed fields directly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Reference {
    source: ObjectId,
    target: ObjectId,
}

impl Reference {
    /// Construct a reference from two non-null typed identities.
    #[must_use]
    pub const fn new(source: ObjectId, target: ObjectId) -> Self {
        Self { source, target }
    }

    /// Construct a reference while reporting a null endpoint explicitly.
    ///
    /// # Errors
    ///
    /// Returns [`ReferenceError::NullSource`] or
    /// [`ReferenceError::NullTarget`] when the corresponding endpoint is
    /// absent.
    pub fn try_new(
        source: Option<ObjectId>,
        target: Option<ObjectId>,
    ) -> Result<Self, ReferenceError> {
        let source_id = source.ok_or(ReferenceError::NullSource)?;
        let target_id = target.ok_or(ReferenceError::NullTarget)?;
        Ok(Self::new(source_id, target_id))
    }

    /// Return the referencing object.
    #[must_use]
    pub const fn source(self) -> ObjectId {
        self.source
    }

    /// Return the referenced object.
    #[must_use]
    pub const fn target(self) -> ObjectId {
        self.target
    }
}

/// Failure while constructing a reference from nullable native fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceError {
    /// The source endpoint was absent/null.
    NullSource,
    /// The target endpoint was absent/null.
    NullTarget,
}

impl fmt::Display for ReferenceError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NullSource => formatter.write_str("reference source must be non-null"),
            Self::NullTarget => formatter.write_str("reference target must be non-null"),
        }
    }
}

impl std::error::Error for ReferenceError {}

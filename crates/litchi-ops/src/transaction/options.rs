//! [`EditOptions`]: the configuration of one explicit `File` transaction
//! (record 0351 §2.5 rule 7, record 0362).

use super::SemanticLimits;

/// Per-call configuration of one explicit transaction: the
/// [`SemanticLimits`] that bound its operations, retained snapshots, target
/// expansion and staged bytes.
///
/// [`File::edit_with_options`](crate::File::edit_with_options),
/// [`File::batch_with_options`](crate::File::batch_with_options),
/// [`File::import_with_options`](crate::File::import_with_options),
/// [`File::merge_with_options`](crate::File::merge_with_options) and
/// [`File::refresh_with_options`](crate::File::refresh_with_options) take it;
/// the forms without options use [`EditOptions::default`]. Source (ingress)
/// limits belong to [`OpenOptions`](crate::OpenOptions), not here.
///
/// ```
/// use litchi_ops::EditOptions;
/// use litchi_ops::transaction::SemanticLimits;
///
/// let limits = SemanticLimits::builder().max_operations(16)?.build()?;
/// let options = EditOptions::new().semantic_limits(limits);
/// assert_eq!(options.semantic_limits_value().max_operations(), 16);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    semantic_limits: SemanticLimits,
}

impl EditOptions {
    /// The default transaction bounds; the same as [`EditOptions::default`].
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the operation, retained-snapshot, target-expansion and
    /// staging bounds.
    pub const fn semantic_limits(mut self, limits: SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The selected semantic limits.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> SemanticLimits {
        self.semantic_limits
    }
}

//! Shared diagnostics of the scoped PPTX edit families (record 0252).

/// The outcome of a scoped edit or commit.
///
/// `changed()` is the canonical no-op introspection (ADR 0034 §3); patches
/// answer the same question through `is_noop()`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    pub(crate) const fn new(changed: bool) -> Self {
        Self { changed }
    }

    /// Whether the edit or commit changes its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

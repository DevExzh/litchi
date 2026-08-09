//! Configurable resource ceilings for spreadsheet tracked changes.

/// Resource ceilings applied while validating tracked changes.
///
/// A zero ceiling is meaningful: it rejects any value that consumes that
/// resource. Builders therefore do not silently replace or reject zero.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_integer_digits: usize,
    max_changes: usize,
    max_nodes: usize,
    max_value_bytes: usize,
    max_aggregate_bytes: usize,
}

impl Limits {
    /// Return the safe default validation budget.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_input_bytes: 32 * 1024 * 1024,
            max_output_bytes: 32 * 1024 * 1024,
            max_integer_digits: 4_096,
            max_changes: 1_000_000,
            max_nodes: 1_000_000,
            max_value_bytes: 65_536,
            max_aggregate_bytes: 16 * 1_048_576,
        }
    }

    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[must_use]
    pub const fn max_integer_digits(self) -> usize {
        self.max_integer_digits
    }

    #[must_use]
    pub const fn max_changes(self) -> usize {
        self.max_changes
    }

    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }

    #[must_use]
    pub const fn max_value_bytes(self) -> usize {
        self.max_value_bytes
    }

    #[must_use]
    pub const fn max_aggregate_bytes(self) -> usize {
        self.max_aggregate_bytes
    }

    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_integer_digits(mut self, maximum: usize) -> Self {
        self.max_integer_digits = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_changes(mut self, maximum: usize) -> Self {
        self.max_changes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_nodes(mut self, maximum: usize) -> Self {
        self.max_nodes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_value_bytes(mut self, maximum: usize) -> Self {
        self.max_value_bytes = maximum;
        self
    }

    #[must_use]
    pub const fn with_max_aggregate_bytes(mut self, maximum: usize) -> Self {
        self.max_aggregate_bytes = maximum;
        self
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

impl Limits {
    pub(crate) fn same_semantic_limits(&self, other: &Self) -> bool {
        self.max_integer_digits() == other.max_integer_digits()
            && self.max_changes() == other.max_changes()
            && self.max_nodes() == other.max_nodes()
            && self.max_value_bytes() == other.max_value_bytes()
            && self.max_aggregate_bytes() == other.max_aggregate_bytes()
    }
}

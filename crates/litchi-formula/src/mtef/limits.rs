//! Caller-selected resource policy for the MTEF codec.

/// A resource checked by the MTEF parser or writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes admitted from the source stream.
    InputBytes,
    /// Recursive MTEF/AST nesting.
    Depth,
    /// Parser and serializer work units.
    Work,
    /// Retained records or AST items.
    Items,
    /// One retained string or font name.
    StringBytes,
    /// Bytes emitted by one serialization.
    OutputBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Self::InputBytes => "input bytes",
            Self::Depth => "depth",
            Self::Work => "work",
            Self::Items => "items",
            Self::StringBytes => "string bytes",
            Self::OutputBytes => "output bytes",
        };
        f.write_str(name)
    }
}

/// Finite limits for one MTEF parse or serialization.
///
/// Raw MTEF entry points require a value selected by their caller. `DEFAULT`
/// is only a convenient finite profile for applications and tests; no codec
/// entry point silently selects it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub(crate) max_input_bytes: usize,
    pub(crate) max_depth: usize,
    pub(crate) max_work: usize,
    pub(crate) max_items: usize,
    pub(crate) max_string_bytes: usize,
    pub(crate) max_output_bytes: usize,
}

impl Limits {
    /// Largest source stream accepted by the codec.
    pub const MAX_INPUT_BYTES: usize = 64 * 1024 * 1024;
    /// Largest recursive MTEF/AST depth accepted by the codec.
    pub const MAX_DEPTH: usize = 256;
    /// Largest parser or serializer work budget.
    pub const MAX_WORK: usize = 8 * 1024 * 1024;
    /// Largest retained record/AST item count.
    pub const MAX_ITEMS: usize = 2 * 1024 * 1024;
    /// Largest single retained string or font name.
    pub const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
    /// Largest serialized MTEF stream.
    pub const MAX_OUTPUT_BYTES: usize = 64 * 1024 * 1024;

    /// A conservative finite profile suitable for one equation.
    pub const DEFAULT: Self = Self {
        max_input_bytes: Self::MAX_INPUT_BYTES,
        max_depth: Self::MAX_DEPTH,
        max_work: Self::MAX_WORK,
        max_items: Self::MAX_ITEMS,
        max_string_bytes: Self::MAX_STRING_BYTES,
        max_output_bytes: Self::MAX_OUTPUT_BYTES,
    };

    /// Construct a finite MTEF policy.
    ///
    /// `None` means that a value is zero or exceeds the codec hard ceiling.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_depth: usize,
        max_work: usize,
        max_items: usize,
        max_string_bytes: usize,
        max_output_bytes: usize,
    ) -> Option<Self> {
        if max_input_bytes == 0
            || max_input_bytes > Self::MAX_INPUT_BYTES
            || max_depth == 0
            || max_depth > Self::MAX_DEPTH
            || max_work == 0
            || max_work > Self::MAX_WORK
            || max_items == 0
            || max_items > Self::MAX_ITEMS
            || max_string_bytes == 0
            || max_string_bytes > Self::MAX_STRING_BYTES
            || max_output_bytes == 0
            || max_output_bytes > Self::MAX_OUTPUT_BYTES
        {
            None
        } else {
            Some(Self {
                max_input_bytes,
                max_depth,
                max_work,
                max_items,
                max_string_bytes,
                max_output_bytes,
            })
        }
    }

    /// Maximum source bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum recursive depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum parser/serializer work units.
    #[must_use]
    pub const fn max_work(self) -> usize {
        self.max_work
    }

    /// Maximum retained records or AST items.
    #[must_use]
    pub const fn max_items(self) -> usize {
        self.max_items
    }

    /// Maximum one string or font name in bytes.
    #[must_use]
    pub const fn max_string_bytes(self) -> usize {
        self.max_string_bytes
    }

    /// Maximum serialized bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

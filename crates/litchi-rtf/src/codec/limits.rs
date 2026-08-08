//! Finite resource profiles for RTF parsing.

use crate::DEFAULT_MAX_DECOMPRESSED_RTF_BYTES;

const MEBIBYTE: usize = 1_048_576;

/// Finite resource limits applied while opening and parsing an RTF document.
///
/// The convenience parsing methods use [`ParseLimits::default`]. Applications
/// may derive a tighter or larger profile with the `with_max_*` methods. Group
/// nesting and format integer bounds remain independently hard-limited.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct ParseLimits {
    max_source_bytes: usize,
    max_tokens: usize,
    max_binary_bytes: usize,
    max_total_binary_bytes: usize,
    max_decompressed_bytes: usize,
    max_opaque_nodes: usize,
    max_opaque_node_bytes: usize,
    max_total_opaque_bytes: usize,
}

impl ParseLimits {
    /// Default maximum encoded or uncompressed source size (256 MiB).
    pub const DEFAULT_MAX_SOURCE_BYTES: usize = 256 * MEBIBYTE;

    /// Default maximum number of lexer tokens.
    pub const DEFAULT_MAX_TOKENS: usize = 4 * 1_048_576;

    /// Default maximum size of one `binN` payload (256 MiB).
    pub const DEFAULT_MAX_BINARY_BYTES: usize = 256 * MEBIBYTE;

    /// Default maximum aggregate size of all `binN` payloads (256 MiB).
    pub const DEFAULT_MAX_TOTAL_BINARY_BYTES: usize = 256 * MEBIBYTE;

    /// Default maximum expanded size of compressed RTF (256 MiB).
    pub const DEFAULT_MAX_DECOMPRESSED_BYTES: usize = DEFAULT_MAX_DECOMPRESSED_RTF_BYTES;

    /// Default maximum number of unsupported syntax nodes retained per document.
    pub const DEFAULT_MAX_OPAQUE_NODES: usize = 65_536;

    /// Default maximum size of one retained unsupported syntax node (8 MiB).
    pub const DEFAULT_MAX_OPAQUE_NODE_BYTES: usize = 8 * MEBIBYTE;

    /// Default aggregate size of retained unsupported syntax (32 MiB).
    pub const DEFAULT_MAX_TOTAL_OPAQUE_BYTES: usize = 32 * MEBIBYTE;

    /// The production-safe default resource profile.
    pub const DEFAULT: Self = Self {
        max_source_bytes: Self::DEFAULT_MAX_SOURCE_BYTES,
        max_tokens: Self::DEFAULT_MAX_TOKENS,
        max_binary_bytes: Self::DEFAULT_MAX_BINARY_BYTES,
        max_total_binary_bytes: Self::DEFAULT_MAX_TOTAL_BINARY_BYTES,
        max_decompressed_bytes: Self::DEFAULT_MAX_DECOMPRESSED_BYTES,
        max_opaque_nodes: Self::DEFAULT_MAX_OPAQUE_NODES,
        max_opaque_node_bytes: Self::DEFAULT_MAX_OPAQUE_NODE_BYTES,
        max_total_opaque_bytes: Self::DEFAULT_MAX_TOTAL_OPAQUE_BYTES,
    };

    /// Create the default finite resource profile.
    pub const fn new() -> Self {
        Self::DEFAULT
    }

    /// Maximum encoded or uncompressed source bytes accepted before parsing.
    pub const fn max_source_bytes(self) -> usize {
        self.max_source_bytes
    }

    /// Maximum number of tokens the lexer may emit.
    pub const fn max_tokens(self) -> usize {
        self.max_tokens
    }

    /// Maximum bytes accepted in one `binN` payload.
    pub const fn max_binary_bytes(self) -> usize {
        self.max_binary_bytes
    }

    /// Maximum aggregate bytes accepted across all `binN` payloads.
    pub const fn max_total_binary_bytes(self) -> usize {
        self.max_total_binary_bytes
    }

    /// Maximum bytes produced by compressed-RTF expansion.
    pub const fn max_decompressed_bytes(self) -> usize {
        self.max_decompressed_bytes
    }

    /// Maximum number of unsupported syntax nodes retained per document.
    pub const fn max_opaque_nodes(self) -> usize {
        self.max_opaque_nodes
    }

    /// Maximum transport bytes retained by one unsupported syntax node.
    pub const fn max_opaque_node_bytes(self) -> usize {
        self.max_opaque_node_bytes
    }

    /// Maximum aggregate transport bytes retained as unsupported syntax.
    pub const fn max_total_opaque_bytes(self) -> usize {
        self.max_total_opaque_bytes
    }

    /// Return a profile with a different source-byte ceiling.
    pub const fn with_max_source_bytes(mut self, limit: usize) -> Self {
        self.max_source_bytes = limit;
        self
    }

    /// Return a profile with a different token ceiling.
    pub const fn with_max_tokens(mut self, limit: usize) -> Self {
        self.max_tokens = limit;
        self
    }

    /// Return a profile with a different per-payload binary ceiling.
    pub const fn with_max_binary_bytes(mut self, limit: usize) -> Self {
        self.max_binary_bytes = limit;
        self
    }

    /// Return a profile with a different aggregate binary ceiling.
    pub const fn with_max_total_binary_bytes(mut self, limit: usize) -> Self {
        self.max_total_binary_bytes = limit;
        self
    }

    /// Return a profile with a different compressed expansion ceiling.
    pub const fn with_max_decompressed_bytes(mut self, limit: usize) -> Self {
        self.max_decompressed_bytes = limit;
        self
    }

    /// Return a profile with a different unsupported-node count ceiling.
    pub const fn with_max_opaque_nodes(mut self, limit: usize) -> Self {
        self.max_opaque_nodes = limit;
        self
    }

    /// Return a profile with a different per-node unsupported-byte ceiling.
    pub const fn with_max_opaque_node_bytes(mut self, limit: usize) -> Self {
        self.max_opaque_node_bytes = limit;
        self
    }

    /// Return a profile with a different aggregate unsupported-byte ceiling.
    pub const fn with_max_total_opaque_bytes(mut self, limit: usize) -> Self {
        self.max_total_opaque_bytes = limit;
        self
    }
}

impl Default for ParseLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_and_derived_profiles_are_explicit() {
        assert_eq!(ParseLimits::new(), ParseLimits::default());
        let limits = ParseLimits::default()
            .with_max_source_bytes(1)
            .with_max_tokens(2)
            .with_max_binary_bytes(3)
            .with_max_total_binary_bytes(4)
            .with_max_decompressed_bytes(5)
            .with_max_opaque_nodes(6)
            .with_max_opaque_node_bytes(7)
            .with_max_total_opaque_bytes(8);
        assert_eq!(limits.max_source_bytes(), 1);
        assert_eq!(limits.max_tokens(), 2);
        assert_eq!(limits.max_binary_bytes(), 3);
        assert_eq!(limits.max_total_binary_bytes(), 4);
        assert_eq!(limits.max_decompressed_bytes(), 5);
        assert_eq!(limits.max_opaque_nodes(), 6);
        assert_eq!(limits.max_opaque_node_bytes(), 7);
        assert_eq!(limits.max_total_opaque_bytes(), 8);
    }
}

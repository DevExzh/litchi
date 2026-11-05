//! Finite physical-source profiles for RTF opening.

use crate::DEFAULT_MAX_DECOMPRESSED_RTF_BYTES;

use crate::leniency::Leniency;

const MEBIBYTE: usize = 1_048_576;

/// A physical RTF ingress dimension.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Encoded or uncompressed source bytes admitted at the transport boundary.
    InputBytes,
    /// Lexer events emitted while tokenizing the source.
    Events,
    /// Bytes in one `binN` payload.
    BinaryBytes,
    /// Aggregate bytes across all `binN` payloads.
    AggregateBinaryBytes,
    /// Bytes produced while expanding compressed RTF.
    DecompressedBytes,
    /// Unsupported syntax nodes retained by the lossless model.
    OpaqueNodes,
    /// Bytes retained in one unsupported syntax node.
    OpaqueNodeBytes,
    /// Aggregate bytes retained in unsupported syntax nodes.
    AggregateOpaqueBytes,
}

impl std::fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Events => "events",
            Self::BinaryBytes => "binary bytes",
            Self::AggregateBinaryBytes => "aggregate binary bytes",
            Self::DecompressedBytes => "decompressed bytes",
            Self::OpaqueNodes => "opaque nodes",
            Self::OpaqueNodeBytes => "opaque node bytes",
            Self::AggregateOpaqueBytes => "aggregate opaque bytes",
        })
    }
}

/// Finite resource limits applied while opening and parsing an RTF document.
///
/// Construct a profile with [`SourceLimits::builder`]. Zero is a valid,
/// explicit denial for every dimension. Values above the addressable signed
/// range are rejected so a valid profile never implies an unrepresentable
/// allocation or counter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
#[allow(
    clippy::struct_field_names,
    reason = "every source limit is a maximum, so the shared max_ prefix is intentional"
)]
pub struct SourceLimits {
    max_input_bytes: usize,
    max_events: usize,
    max_binary_bytes: usize,
    max_aggregate_binary_bytes: usize,
    max_decompressed_bytes: usize,
    max_opaque_nodes: usize,
    max_opaque_node_bytes: usize,
    max_aggregate_opaque_bytes: usize,
}

impl SourceLimits {
    /// Maximum value addressable by every source dimension.
    pub const MAX_VALUE: usize = isize::MAX as usize;

    /// Default maximum encoded or uncompressed source size (512 MiB).
    pub const DEFAULT_MAX_INPUT_BYTES: usize = 512 * MEBIBYTE;

    /// Default maximum number of lexer events.
    pub const DEFAULT_MAX_EVENTS: usize = 4 * 1_048_576;

    /// Default maximum size of one `binN` payload (256 MiB).
    pub const DEFAULT_MAX_BINARY_BYTES: usize = 256 * MEBIBYTE;

    /// Default maximum aggregate size of all `binN` payloads (256 MiB).
    pub const DEFAULT_MAX_AGGREGATE_BINARY_BYTES: usize = 256 * MEBIBYTE;

    /// Default maximum expanded size of compressed RTF (256 MiB).
    pub const DEFAULT_MAX_DECOMPRESSED_BYTES: usize = DEFAULT_MAX_DECOMPRESSED_RTF_BYTES;

    /// Default maximum number of unsupported syntax nodes retained per document.
    pub const DEFAULT_MAX_OPAQUE_NODES: usize = 65_536;

    /// Default maximum size of one retained unsupported syntax node (8 MiB).
    pub const DEFAULT_MAX_OPAQUE_NODE_BYTES: usize = 8 * MEBIBYTE;

    /// Default aggregate size of retained unsupported syntax (32 MiB).
    pub const DEFAULT_MAX_AGGREGATE_OPAQUE_BYTES: usize = 32 * MEBIBYTE;

    /// The production-safe default resource profile.
    pub const DEFAULT: Self = Self {
        max_input_bytes: Self::DEFAULT_MAX_INPUT_BYTES,
        max_events: Self::DEFAULT_MAX_EVENTS,
        max_binary_bytes: Self::DEFAULT_MAX_BINARY_BYTES,
        max_aggregate_binary_bytes: Self::DEFAULT_MAX_AGGREGATE_BINARY_BYTES,
        max_decompressed_bytes: Self::DEFAULT_MAX_DECOMPRESSED_BYTES,
        max_opaque_nodes: Self::DEFAULT_MAX_OPAQUE_NODES,
        max_opaque_node_bytes: Self::DEFAULT_MAX_OPAQUE_NODE_BYTES,
        max_aggregate_opaque_bytes: Self::DEFAULT_MAX_AGGREGATE_OPAQUE_BYTES,
    };

    /// Start a checked builder initialized with the standard defaults.
    pub const fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum encoded or uncompressed source bytes admitted before parsing.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum number of lexer events emitted while parsing.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Maximum bytes accepted in one `binN` payload.
    #[must_use]
    pub const fn max_binary_bytes(self) -> usize {
        self.max_binary_bytes
    }

    /// Maximum aggregate bytes accepted across all `binN` payloads.
    #[must_use]
    pub const fn max_aggregate_binary_bytes(self) -> usize {
        self.max_aggregate_binary_bytes
    }

    /// Maximum bytes produced by compressed-RTF expansion.
    #[must_use]
    pub const fn max_decompressed_bytes(self) -> usize {
        self.max_decompressed_bytes
    }

    /// Maximum number of unsupported syntax nodes retained per document.
    #[must_use]
    pub const fn max_opaque_nodes(self) -> usize {
        self.max_opaque_nodes
    }

    /// Maximum bytes retained by one unsupported syntax node.
    #[must_use]
    pub const fn max_opaque_node_bytes(self) -> usize {
        self.max_opaque_node_bytes
    }

    /// Maximum aggregate bytes retained as unsupported syntax.
    #[must_use]
    pub const fn max_aggregate_opaque_bytes(self) -> usize {
        self.max_aggregate_opaque_bytes
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for a checked [`SourceLimits`] profile.
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Set the encoded or uncompressed source-byte ceiling.
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::InputBytes, value)?;
        self.limits.max_input_bytes = value;
        Ok(self)
    }

    /// Set the lexer-event ceiling.
    pub fn max_events(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::Events, value)?;
        self.limits.max_events = value;
        Ok(self)
    }

    /// Set the per-payload binary ceiling.
    pub fn max_binary_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::BinaryBytes, value)?;
        self.limits.max_binary_bytes = value;
        Ok(self)
    }

    /// Set the aggregate binary-payload ceiling.
    pub fn max_aggregate_binary_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::AggregateBinaryBytes, value)?;
        self.limits.max_aggregate_binary_bytes = value;
        Ok(self)
    }

    /// Set the compressed-RTF expansion ceiling.
    pub fn max_decompressed_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::DecompressedBytes, value)?;
        self.limits.max_decompressed_bytes = value;
        Ok(self)
    }

    /// Set the unsupported-node count ceiling.
    pub fn max_opaque_nodes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::OpaqueNodes, value)?;
        self.limits.max_opaque_nodes = value;
        Ok(self)
    }

    /// Set the per-node unsupported-byte ceiling.
    pub fn max_opaque_node_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::OpaqueNodeBytes, value)?;
        self.limits.max_opaque_node_bytes = value;
        Ok(self)
    }

    /// Set the aggregate unsupported-byte ceiling.
    pub fn max_aggregate_opaque_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        validate(SourceLimitKind::AggregateOpaqueBytes, value)?;
        self.limits.max_aggregate_opaque_bytes = value;
        Ok(self)
    }

    /// Finalize this checked profile.
    pub fn build(self) -> Result<SourceLimits, crate::Error> {
        validate(SourceLimitKind::InputBytes, self.limits.max_input_bytes)?;
        validate(SourceLimitKind::Events, self.limits.max_events)?;
        validate(SourceLimitKind::BinaryBytes, self.limits.max_binary_bytes)?;
        validate(
            SourceLimitKind::AggregateBinaryBytes,
            self.limits.max_aggregate_binary_bytes,
        )?;
        validate(
            SourceLimitKind::DecompressedBytes,
            self.limits.max_decompressed_bytes,
        )?;
        validate(SourceLimitKind::OpaqueNodes, self.limits.max_opaque_nodes)?;
        validate(
            SourceLimitKind::OpaqueNodeBytes,
            self.limits.max_opaque_node_bytes,
        )?;
        validate(
            SourceLimitKind::AggregateOpaqueBytes,
            self.limits.max_aggregate_opaque_bytes,
        )?;
        Ok(self.limits)
    }
}

fn validate(kind: SourceLimitKind, value: usize) -> Result<(), crate::Error> {
    if value > SourceLimits::MAX_VALUE {
        return Err(crate::Error::InvalidLimit {
            kind,
            value,
            maximum: SourceLimits::MAX_VALUE,
        });
    }
    Ok(())
}

/// How a document is opened: its finite source profile and repair posture.
///
/// The default options pair [`SourceLimits::default`] with strict leniency.
/// Word-compatible repairs are applied and recorded under every leniency;
/// opt-in repairs are selected by [`OpenOptions::leniency`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct OpenOptions {
    source_limits: SourceLimits,
    leniency: Leniency,
}

impl OpenOptions {
    /// Construct strict options with the standard finite source profile.
    pub const fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            leniency: Leniency::Strict,
        }
    }

    /// Select a checked physical source profile.
    pub const fn source_limits(mut self, source_limits: SourceLimits) -> Self {
        self.source_limits = source_limits;
        self
    }

    /// Select the repair policy.
    pub const fn leniency(mut self, leniency: Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Return the selected physical source profile.
    pub const fn source_limits_value(self) -> SourceLimits {
        self.source_limits
    }

    /// Return the selected repair policy.
    #[must_use]
    pub const fn leniency_value(self) -> Leniency {
        self.leniency
    }
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl From<SourceLimits> for OpenOptions {
    fn from(source_limits: SourceLimits) -> Self {
        Self::new().source_limits(source_limits)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_and_checked_builder_expose_canonical_dimensions() {
        assert_eq!(SourceLimits::default(), SourceLimits::DEFAULT);
        let limits = SourceLimits::builder()
            .max_input_bytes(1)
            .expect("input limit")
            .max_events(2)
            .expect("event limit")
            .max_binary_bytes(3)
            .expect("binary limit")
            .max_aggregate_binary_bytes(4)
            .expect("aggregate binary limit")
            .max_decompressed_bytes(5)
            .expect("decompressed limit")
            .max_opaque_nodes(6)
            .expect("opaque-node limit")
            .max_opaque_node_bytes(7)
            .expect("opaque-node byte limit")
            .max_aggregate_opaque_bytes(8)
            .expect("aggregate opaque limit")
            .build()
            .expect("source limits");
        assert_eq!(limits.max_input_bytes(), 1);
        assert_eq!(limits.max_events(), 2);
        assert_eq!(limits.max_binary_bytes(), 3);
        assert_eq!(limits.max_aggregate_binary_bytes(), 4);
        assert_eq!(limits.max_decompressed_bytes(), 5);
        assert_eq!(limits.max_opaque_nodes(), 6);
        assert_eq!(limits.max_opaque_node_bytes(), 7);
        assert_eq!(limits.max_aggregate_opaque_bytes(), 8);
    }

    #[test]
    fn zero_is_an_explicit_denial_and_unaddressable_values_are_invalid() {
        let limits = SourceLimits::builder()
            .max_opaque_nodes(0)
            .expect("zero is a valid denial")
            .build()
            .expect("source limits");
        assert_eq!(limits.max_opaque_nodes(), 0);
        assert!(matches!(
            SourceLimits::builder().max_input_bytes(usize::MAX),
            Err(crate::Error::InvalidLimit {
                kind: SourceLimitKind::InputBytes,
                value,
                maximum,
            }) if value == usize::MAX && maximum == isize::MAX as usize
        ));
    }

    #[test]
    fn default_options_are_strict_and_retain_typed_values() {
        let options = OpenOptions::default();
        assert_eq!(options, OpenOptions::new());
        assert_eq!(options.source_limits_value(), SourceLimits::default());
        assert_eq!(options.leniency_value(), Leniency::Strict);
        let limits = SourceLimits::builder()
            .max_events(9)
            .expect("event limit")
            .build()
            .expect("source limits");
        assert_eq!(
            OpenOptions::new()
                .source_limits(limits)
                .source_limits_value(),
            limits
        );
    }
}

//! Bounded batch JSON ingress with content-free, typed failures.

use std::collections::TryReserveError;
use std::fmt;

use serde::de::{self, Deserialize, DeserializeSeed, Deserializer, SeqAccess, Visitor};

use super::BatchItem;
use crate::Result;

/// A physical resource consumed when decoding batch JSON.
///
/// ```
/// use litchi_ops::batch::SourceLimitKind;
/// assert_ne!(SourceLimitKind::InputBytes, SourceLimitKind::Items);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Complete JSON input, including whitespace, strings and properties.
    InputBytes,
    /// Number of command envelopes in the array.
    Items,
}

/// Typed batch ingress failure. No variant carries authored JSON or a
/// decoder message that might quote it.
///
/// Conversion to [`crate::Error`] retains this value in
/// [`crate::error::Source::Batch`].
///
/// ```
/// use litchi_ops::{Code, Error};
/// use litchi_ops::batch::{Error as BatchError, SourceLimitKind};
/// let error = Error::from(BatchError::Limit {
///     kind: SourceLimitKind::Items, observed: 11, maximum: 10,
/// });
/// assert_eq!(error.code(), Code::LimitExceeded);
/// ```
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A limit is zero or exceeds the supported allocation range.
    InvalidLimit {
        /// The dimension being configured.
        kind: SourceLimitKind,
        /// The rejected value.
        value: usize,
        /// The largest supported value.
        maximum: usize,
    },
    /// Input exceeded a caller-selected bound.
    Limit {
        /// The dimension consumed.
        kind: SourceLimitKind,
        /// The observed count (saturated on arithmetic overflow).
        observed: usize,
        /// The configured maximum.
        maximum: usize,
    },
    /// JSON syntax or a command envelope is invalid.
    Malformed {
        /// Decoder line, when available.
        line: usize,
        /// Decoder column, when available.
        column: usize,
    },
    /// Retaining the bounded item list failed.
    Allocation {
        /// The allocator's typed refusal.
        source: TryReserveError,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid batch {kind:?} limit {value}; expected 1..={maximum}"
            ),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "batch {kind:?} limit exceeded: {observed} > {maximum}"
            ),
            Self::Malformed { line, column } => write!(
                formatter,
                "invalid batch JSON or item envelope at line {line}, column {column}"
            ),
            Self::Allocation { .. } => formatter.write_str("cannot allocate the batch item list"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Allocation { source } => Some(source),
            _ => None,
        }
    }
}

/// Caller-selected physical bounds for a complete batch JSON array.
///
/// Defaults are 512 MiB input (the shared source-byte default) and 100,000
/// items. The byte bound covers property names/values and escaped text too;
/// it is checked before invoking the JSON decoder. No item beyond the item
/// bound is decoded. The decoder keeps its own bounded JSON nesting policy.
///
/// ```
/// use litchi_ops::batch::SourceLimits;
/// let limits = SourceLimits::builder().max_input_bytes(4096)?.max_items(20)?.build()?;
/// assert_eq!(limits.max_items(), 20);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: usize,
    items: usize,
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            input_bytes: 512 * 1024 * 1024,
            items: 100_000,
        }
    }
}

impl SourceLimits {
    /// Start from the default source limits.
    #[must_use]
    pub fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder(Self::default())
    }

    /// Maximum complete input bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Maximum command count.
    #[must_use]
    pub const fn max_items(self) -> usize {
        self.items
    }
}

/// Consuming builder of checked batch [`SourceLimits`].
///
/// ```
/// use litchi_ops::batch::SourceLimits;
/// assert!(SourceLimits::builder().max_items(0).is_err());
/// ```
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder(SourceLimits);

impl SourceLimitsBuilder {
    /// Set the maximum complete JSON input size.
    ///
    /// # Errors
    /// Returns an `invalid_argument` with [`Error::InvalidLimit`] for zero
    /// or a value larger than `isize::MAX`.
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self> {
        check_limit(SourceLimitKind::InputBytes, value, isize::MAX as usize)?;
        self.0.input_bytes = value;
        Ok(self)
    }

    /// Set the maximum item count.
    ///
    /// # Errors
    /// Returns an `invalid_argument` with [`Error::InvalidLimit`] for zero
    /// or a count whose contiguous item array cannot fit in `isize::MAX`.
    pub fn max_items(mut self, value: usize) -> Result<Self> {
        check_limit(
            SourceLimitKind::Items,
            value,
            isize::MAX as usize / size_of::<BatchItem<'_>>(),
        )?;
        self.0.items = value;
        Ok(self)
    }

    /// Finish the checked profile.
    ///
    /// # Errors
    /// Returns a typed invalid-limit error if a profile is outside the
    /// supported allocation range.
    pub fn build(self) -> Result<SourceLimits> {
        check_limit(
            SourceLimitKind::InputBytes,
            self.0.input_bytes,
            isize::MAX as usize,
        )?;
        check_limit(
            SourceLimitKind::Items,
            self.0.items,
            isize::MAX as usize / size_of::<BatchItem<'_>>(),
        )?;
        Ok(self.0)
    }
}

fn check_limit(kind: SourceLimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// Decode a complete batch JSON array under caller-selected physical
/// bounds ([`SourceLimits::default`] for the defaults). Unescaped strings
/// borrow from `bytes`; the result owns its item/property lists. A failed
/// decode returns no partial list. Trailing JSON is refused. No command
/// executes and no input is silently repaired.
///
/// # Errors
/// Returns `corrupt_batch` for invalid JSON/envelopes, `limit_exceeded` for
/// physical bounds, or a typed allocation error. Item failures carry their
/// zero-based index; an envelope or trailing-data failure has no item index.
///
/// An input-byte refusal occurs before parsing any item; an item-count
/// refusal occurs before decoding the excess item.
///
/// ```
/// use litchi_ops::batch::{self, SourceLimits};
/// let items = batch::from_bytes(
///     br#"[{"op":"set","path":"/body/p[1]","props":{"text":"Hello"}}]"#,
///     &SourceLimits::default(),
/// )?;
/// assert_eq!(items[0].command, "set");
/// let limits = SourceLimits::builder().max_input_bytes(100)?.max_items(1)?.build()?;
/// let items = batch::from_bytes(br#"[{"command":"meta","dumpVersion":2}]"#, &limits)?;
/// assert_eq!(items.len(), 1);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub fn from_bytes<'a>(bytes: &'a [u8], limits: &SourceLimits) -> Result<Vec<BatchItem<'a>>> {
    if bytes.len() > limits.input_bytes {
        return Err(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: bytes.len(),
            maximum: limits.input_bytes,
        }
        .into());
    }
    let mut state = State {
        failure: None,
        item: None,
    };
    let mut decoder = serde_json::Deserializer::from_slice(bytes);
    let result = (&mut decoder).deserialize_seq(ItemsVisitor {
        limits,
        state: &mut state,
    });
    match result {
        Ok(items) => {
            decoder.end().map_err(|error| {
                crate::Error::from(Error::Malformed {
                    line: error.line(),
                    column: error.column(),
                })
            })?;
            Ok(items)
        },
        Err(error) => {
            let error = state.failure.unwrap_or(Error::Malformed {
                line: error.line(),
                column: error.column(),
            });
            let error = crate::Error::from(error);
            Err(match state.item {
                Some(index) => error.with_item(index),
                None => error,
            })
        },
    }
}

struct State {
    failure: Option<Error>,
    item: Option<usize>,
}
struct ItemsVisitor<'a> {
    limits: &'a SourceLimits,
    state: &'a mut State,
}

impl<'de> Visitor<'de> for ItemsVisitor<'_> {
    type Value = Vec<BatchItem<'de>>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a batch array")
    }

    fn visit_seq<A: SeqAccess<'de>>(
        self,
        mut seq: A,
    ) -> std::result::Result<Self::Value, A::Error> {
        let mut items = Vec::new();
        loop {
            let index = items.len();
            self.state.item = None;
            let item = seq.next_element_seed(ItemSeed {
                index,
                maximum: self.limits.items,
                state: self.state,
            })?;
            let Some(item) = item else { return Ok(items) };
            if let Err(source) = items.try_reserve(1) {
                self.state.failure = Some(Error::Allocation { source });
                return Err(de::Error::custom("batch allocation failed"));
            }
            items.push(item);
        }
    }
}

struct ItemSeed<'a> {
    index: usize,
    maximum: usize,
    state: &'a mut State,
}
impl<'de> DeserializeSeed<'de> for ItemSeed<'_> {
    type Value = BatchItem<'de>;

    fn deserialize<D: Deserializer<'de>>(
        self,
        decoder: D,
    ) -> std::result::Result<Self::Value, D::Error> {
        self.state.item = Some(self.index);
        if self.index >= self.maximum {
            self.state.failure = Some(Error::Limit {
                kind: SourceLimitKind::Items,
                observed: self.index.checked_add(1).unwrap_or(usize::MAX),
                maximum: self.maximum,
            });
            return Err(de::Error::custom("batch item limit exceeded"));
        }
        BatchItem::deserialize(decoder)
    }
}

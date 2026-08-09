//! Immutable source-preserving toolbar-control snapshots.

use std::sync::Arc;

use super::codec;
use super::control::{Body, Control};
use super::validation::MAX_CONTROL_BYTES;
use super::{Error, Transaction};

/// A deterministic identity for one exact serialized toolbar control.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(crate) fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Return the compact source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

/// An immutable, source-preserving toolbar-control snapshot.
///
/// The projection is owned so a snapshot can outlive the caller's CFB stream,
/// while the exact source bytes remain one shareable allocation. Unsupported
/// control bodies and host-specific prefixes stay attached to the projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    control: Control<'static>,
    revision: Revision,
}

impl Snapshot {
    /// Parse one complete `TBCHeader` and its remaining body.
    ///
    /// Known controls are projected as `Data` when the complete remainder is
    /// a valid `TBCData`; otherwise it is retained as opaque bytes. Use
    /// [`Self::parse_with_prefix`] when the host format places an opaque
    /// command structure between the header and data.
    ///
    /// # Errors
    ///
    /// Returns an error if the control is malformed or exceeds the bounded
    /// control size.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self, Error> {
        Self::parse_shared(Arc::from(bytes.as_ref().to_vec().into_boxed_slice()))
    }

    /// Parse a complete control while retaining the source allocation.
    ///
    /// # Errors
    ///
    /// Returns an error if the control is malformed or exceeds the bounded
    /// control size.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self, Error> {
        Self::parse_inner(bytes, 0)
    }

    /// Parse a complete control with an explicit opaque prefix length.
    ///
    /// # Errors
    ///
    /// Returns an error if the control or prefix is malformed, or the control
    /// exceeds its bounded size.
    pub fn parse_with_prefix(bytes: impl AsRef<[u8]>, prefix_len: usize) -> Result<Self, Error> {
        Self::parse_with_prefix_shared(
            Arc::from(bytes.as_ref().to_vec().into_boxed_slice()),
            prefix_len,
        )
    }

    /// Parse a complete control with an explicit prefix without copying it.
    ///
    /// # Errors
    ///
    /// Returns an error if the control or prefix is malformed, or the control
    /// exceeds its bounded size.
    pub fn parse_with_prefix_shared(bytes: Arc<[u8]>, prefix_len: usize) -> Result<Self, Error> {
        Self::parse_inner(bytes, prefix_len)
    }

    /// Capture a typed control as an immutable source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` violates the shared toolbar invariants.
    pub fn from_control(value: Control<'_>) -> Result<Self, Error> {
        let owned = value.into_owned();
        let bytes = Arc::from(owned.to_bytes().into_boxed_slice());
        Self::from_parts(bytes, owned)
    }

    /// Borrow the complete typed or opaque control projection.
    #[must_use]
    pub const fn control(&self) -> &Control<'static> {
        &self.control
    }

    /// Borrow the fixed control header.
    #[must_use]
    pub const fn header(&self) -> &super::ControlHeader {
        self.control.header()
    }

    /// Borrow the control body.
    #[must_use]
    pub const fn body(&self) -> &Body<'static> {
        self.control.body()
    }

    /// Borrow the exact source or committed bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the shared source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// Return the source fingerprint.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Return the compact source fingerprint.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.revision.value()
    }

    /// Start an isolated typed edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    fn parse_inner(bytes: Arc<[u8]>, prefix_len: usize) -> Result<Self, Error> {
        if bytes.len() > MAX_CONTROL_BYTES {
            return Err(Error::invalid("toolbar control exceeds the bounded size"));
        }
        let (header, header_len) = super::ControlHeader::parse_prefix(bytes.as_ref())?;
        let body_start = header_len
            .checked_add(prefix_len)
            .ok_or_else(|| Error::invalid("toolbar prefix offset overflows usize"))?;
        let control = {
            let prefix = bytes
                .get(header_len..body_start)
                .ok_or(Error::Truncated("toolbar control prefix"))?;
            if prefix.len() > super::validation::MAX_PREFIX_BYTES {
                return Err(Error::invalid(
                    "toolbar control prefix exceeds the bounded limit",
                ));
            }
            let remainder = bytes
                .get(body_start..)
                .ok_or(Error::Truncated("toolbar control body"))?;
            let body = codec::parse_body(&header, remainder);
            Control::from_decoded(header, prefix, body)?.into_owned()
        };
        Self::from_parts(bytes, control)
    }

    pub(crate) fn from_parts(bytes: Arc<[u8]>, control: Control<'static>) -> Result<Self, Error> {
        if bytes.len() > MAX_CONTROL_BYTES {
            return Err(Error::invalid("toolbar control exceeds the bounded size"));
        }
        let encoded = control.to_bytes();
        if encoded.as_slice() != bytes.as_ref() {
            return Err(Error::invalid(
                "toolbar control projection is not source-preserving",
            ));
        }
        super::validation::validate_decoded(&control)?;
        Ok(Self {
            revision: Revision::from_bytes(&bytes),
            bytes,
            control,
        })
    }
}

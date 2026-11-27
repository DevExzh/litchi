use super::{IS_BLIP, IS_COMPLEX, Id, Value};

/// One ordered, lossless property-table descriptor and its decoded value.
#[derive(Debug)]
pub struct Prop<'data> {
    pub(in crate::prop) id: Id,
    pub(in crate::prop) raw_id: u16,
    pub(in crate::prop) blip: bool,
    pub(in crate::prop) complex: bool,
    pub(in crate::prop) raw_value: i32,
    pub(in crate::prop) value: Value<'data>,
}

impl<'data> Prop<'data> {
    /// Returns the typed identifier.
    #[must_use]
    pub const fn id(&self) -> Id {
        self.id
    }

    /// Returns the exact 14-bit identifier without flags.
    #[must_use]
    pub const fn raw_id(&self) -> u16 {
        self.raw_id
    }

    /// Reassembles the exact 16-bit identifier-and-flags field.
    #[must_use]
    pub const fn raw_opid(&self) -> u16 {
        self.raw_id
            | if self.blip { IS_BLIP } else { 0 }
            | if self.complex { IS_COMPLEX } else { 0 }
    }

    /// Returns whether `fBid` was set on the wire.
    #[must_use]
    pub const fn is_blip(&self) -> bool {
        self.blip
    }

    /// Returns whether `fComplex` was set on the wire.
    #[must_use]
    pub const fn is_complex(&self) -> bool {
        self.complex
    }

    /// Returns the exact signed four-byte `op` value.
    #[must_use]
    pub const fn raw_value(&self) -> i32 {
        self.raw_value
    }

    /// Borrows the decoded value.
    #[must_use]
    pub const fn value(&self) -> &Value<'data> {
        &self.value
    }
}

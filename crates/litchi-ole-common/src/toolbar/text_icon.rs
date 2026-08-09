use std::borrow::Cow;

use super::Error;

/// A borrowed or owned non-null-terminated UTF-16LE `WString`.
///
/// Decoded wire strings borrow their encoded bytes. Text construction owns a
/// compact UTF-16LE buffer, avoiding alignment-dependent casts and `unsafe`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WString<'a> {
    encoded: Cow<'a, [u8]>,
}

impl<'a> WString<'a> {
    /// Construct a string from valid Rust UTF-8 text.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` contains a NUL or exceeds the one-byte
    /// `WString` length limit.
    pub fn new(value: &str) -> Result<Self, Error> {
        if value.contains('\0') {
            return Err(Error::invalid("WString cannot contain NUL characters"));
        }
        let units = value.encode_utf16().collect::<Vec<_>>();
        Self::from_units(&units)
    }

    /// Construct a string from validated UTF-16 code units.
    ///
    /// # Errors
    ///
    /// Returns an error if `units` contains a NUL or unpaired surrogate, or
    /// exceeds the one-byte `WString` length limit.
    pub fn from_units(units: &[u16]) -> Result<Self, Error> {
        validate_units(units.iter().copied())?;
        let byte_len = units
            .len()
            .checked_mul(2)
            .ok_or_else(|| Error::invalid("WString byte length overflows usize"))?;
        if units.len() > usize::from(u8::MAX) {
            return Err(Error::invalid(
                "WString contains more than 255 UTF-16 code units",
            ));
        }
        let mut encoded = Vec::with_capacity(byte_len);
        for unit in units {
            encoded.extend_from_slice(&unit.to_le_bytes());
        }
        Ok(Self {
            encoded: Cow::Owned(encoded),
        })
    }

    pub(crate) fn from_wire(encoded: &'a [u8]) -> Result<Self, Error> {
        if !encoded.len().is_multiple_of(2) {
            return Err(Error::invalid("WString payload has an odd byte count"));
        }
        validate_units(encoded_units(encoded))?;
        Ok(Self {
            encoded: Cow::Borrowed(encoded),
        })
    }

    /// Return the number of UTF-16 code units in the string.
    #[must_use]
    pub fn len(&self) -> usize {
        self.encoded.len() / 2
    }

    /// Return whether the string is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.encoded.is_empty()
    }

    /// Return the original UTF-16LE payload without its one-byte length.
    #[must_use]
    pub fn encoded_bytes(&self) -> &[u8] {
        &self.encoded
    }

    /// Iterate over the validated UTF-16 code units without allocating.
    pub fn units(&self) -> impl Iterator<Item = u16> + '_ {
        encoded_units(&self.encoded)
    }

    /// Decode the string into owned Rust text.
    #[must_use]
    pub fn text(&self) -> String {
        let units = self.units().collect::<Vec<_>>();
        String::from_utf16_lossy(&units)
    }

    /// Copy a borrowed string into an owned representation.
    #[must_use]
    pub fn into_owned(self) -> WString<'static> {
        WString {
            encoded: Cow::Owned(self.encoded.into_owned()),
        }
    }

    pub(crate) fn encoded_len(&self) -> usize {
        self.encoded.len()
    }
}

/// Visibility mode stored in the two-bit `textIcon` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextIcon {
    /// Raw value `0`: icon on a basic toolbar, text and icon on a menu.
    IconOnly,
    /// Raw value `1`: icon on a basic toolbar, text only on a menu.
    MenuText,
    /// Raw value `2`: text only.
    TextOnly,
    /// Raw value `3`: text and icon.
    TextAndIcon,
}

impl TextIcon {
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::IconOnly => 0,
            Self::MenuText => 1,
            Self::TextOnly => 2,
            Self::TextAndIcon => 3,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value & 0x03 {
            0 => Self::IconOnly,
            1 => Self::MenuText,
            2 => Self::TextOnly,
            _ => Self::TextAndIcon,
        }
    }
}

impl TryFrom<u8> for TextIcon {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        if value > 3 {
            return Err(Error::invalid("textIcon exceeds two bits"));
        }
        Ok(Self::from_raw(value))
    }
}

/// Button state from the two-bit `state` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ButtonState {
    /// Raw value `0`.
    Up,
    /// Raw value `1`.
    Down,
    /// Raw value `3`.
    Mixed,
    /// An invalid or future raw value retained by a borrowed decode.
    Unknown(u8),
}

impl ButtonState {
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::Up => 0,
            Self::Down => 1,
            Self::Mixed => 3,
            Self::Unknown(value) => value & 0x03,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value & 0x03 {
            0 => Self::Up,
            1 => Self::Down,
            3 => Self::Mixed,
            unknown => Self::Unknown(unknown),
        }
    }
}

impl TryFrom<u8> for ButtonState {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let state = Self::from_raw(value);
        if matches!(state, Self::Unknown(_)) {
            return Err(Error::invalid("button state 0b10 is reserved"));
        }
        Ok(state)
    }
}

/// Hyperlink type from the two-bit `fHyperlinkType` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HyperlinkType {
    /// No hyperlink.
    None,
    /// Open the link in a browser.
    Browser,
    /// Link to an image.
    Image,
    /// An invalid or future raw value retained by a decode.
    Unknown(u8),
}

impl HyperlinkType {
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Browser => 1,
            Self::Image => 2,
            Self::Unknown(value) => value & 0x03,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value & 0x03 {
            0 => Self::None,
            1 => Self::Browser,
            2 => Self::Image,
            unknown => Self::Unknown(unknown),
        }
    }
}

impl TryFrom<u8> for HyperlinkType {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let kind = Self::from_raw(value);
        if matches!(kind, Self::Unknown(_)) {
            return Err(Error::invalid("hyperlink type 0b11 is reserved"));
        }
        Ok(kind)
    }
}

/// Button and expanding-grid flags from `TBCBSFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ButtonFlags {
    raw: u8,
}

impl ButtonFlags {
    /// Retain a raw wire value, including the reserved bit.
    #[must_use]
    pub const fn from_raw(raw: u8) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking the two reserved
    /// discriminants. The final reserved bit is advisory in [MS-OSHARED]
    /// and is therefore retained without rejection.
    ///
    /// # Errors
    ///
    /// Returns an error if either reserved two-bit discriminator is selected.
    pub fn try_from_raw(raw: u8) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag byte.
    #[must_use]
    pub const fn raw(self) -> u8 {
        self.raw
    }

    /// Return the advisory reserved bit exactly as stored.
    #[must_use]
    pub const fn reserved_bits(self) -> u8 {
        self.raw & 0x80
    }

    #[must_use]
    pub const fn state(self) -> ButtonState {
        ButtonState::from_raw(self.raw)
    }

    #[must_use]
    pub const fn accelerator(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    #[must_use]
    pub const fn custom_bitmap(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    #[must_use]
    pub const fn custom_button_face(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    #[must_use]
    pub const fn hyperlink(self) -> HyperlinkType {
        HyperlinkType::from_raw(self.raw >> 5)
    }

    #[must_use]
    pub fn with_state(mut self, value: ButtonState) -> Self {
        self.raw = (self.raw & !0x03) | value.raw();
        self
    }

    #[must_use]
    pub fn with_accelerator(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    #[must_use]
    pub fn with_custom_bitmap(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    #[must_use]
    pub fn with_custom_button_face(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

    #[must_use]
    pub fn with_hyperlink(mut self, value: HyperlinkType) -> Self {
        self.raw = (self.raw & !(0x03 << 5)) | (value.raw() << 5);
        self
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        ButtonState::try_from(self.raw & 0x03)?;
        HyperlinkType::try_from((self.raw >> 5) & 0x03)?;
        Ok(())
    }
}

trait BitOps:
    Copy
    + std::ops::BitOr<Output = Self>
    + std::ops::BitAnd<Output = Self>
    + std::ops::Not<Output = Self>
    + std::ops::Shl<u32, Output = Self>
{
    fn one() -> Self;
}

impl BitOps for u8 {
    fn one() -> Self {
        1
    }
}

fn encoded_units(bytes: &[u8]) -> impl Iterator<Item = u16> + '_ {
    bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
}

fn validate_units(units: impl Iterator<Item = u16>) -> Result<(), Error> {
    let mut pending_high = false;
    for unit in units {
        if unit == 0 {
            return Err(Error::invalid("WString cannot contain NUL code units"));
        }
        if pending_high {
            if !(0xDC00..=0xDFFF).contains(&unit) {
                return Err(Error::invalid(
                    "WString contains an unpaired UTF-16 surrogate",
                ));
            }
            pending_high = false;
        } else if (0xD800..=0xDBFF).contains(&unit) {
            pending_high = true;
        } else if (0xDC00..=0xDFFF).contains(&unit) {
            return Err(Error::invalid(
                "WString contains an unpaired UTF-16 surrogate",
            ));
        }
    }
    if pending_high {
        return Err(Error::invalid(
            "WString ends with an unpaired UTF-16 surrogate",
        ));
    }
    Ok(())
}

fn set_bit<T>(raw: T, bit: u32, value: bool) -> T
where
    T: BitOps,
{
    if value {
        raw | T::one() << bit
    } else {
        raw & !(T::one() << bit)
    }
}

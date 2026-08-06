use std::borrow::Cow;
use std::fmt;

/// A malformed or semantically invalid [MS-OSHARED] toolbar structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The input ended before the named field was complete.
    Truncated(&'static str),
    /// The structure violates a wire or semantic invariant.
    Invalid(String),
}

impl Error {
    pub(crate) fn invalid(message: impl Into<String>) -> Self {
        Self::Invalid(message.into())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated(field) => write!(formatter, "truncated toolbar {field}"),
            Self::Invalid(message) => write!(formatter, "invalid toolbar structure: {message}"),
        }
    }
}

impl std::error::Error for Error {}

/// A borrowed or owned non-null-terminated UTF-16LE `WString`.
///
/// Decoded wire strings borrow their encoded bytes.  Text construction owns a
/// compact UTF-16LE buffer, avoiding alignment-dependent casts and `unsafe`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WString<'a> {
    encoded: Cow<'a, [u8]>,
}

impl<'a> WString<'a> {
    /// Construct a string from valid Rust UTF-8 text.
    pub fn new(value: &str) -> Result<Self, Error> {
        if value.contains('\0') {
            return Err(Error::invalid("WString cannot contain NUL characters"));
        }
        let units = value.encode_utf16().collect::<Vec<_>>();
        Self::from_units(&units)
    }

    /// Construct a string from validated UTF-16 code units.
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
        if encoded.len() % 2 != 0 {
            return Err(Error::invalid("WString payload has an odd byte count"));
        }
        validate_units(encoded_units(encoded))?;
        Ok(Self {
            encoded: Cow::Borrowed(encoded),
        })
    }

    /// Return the number of UTF-16 code units in the string.
    pub fn len(&self) -> usize {
        self.encoded.len() / 2
    }

    /// Return whether the string is empty.
    pub fn is_empty(&self) -> bool {
        self.encoded.is_empty()
    }

    /// Return the original UTF-16LE payload without its one-byte length.
    pub fn encoded_bytes(&self) -> &[u8] {
        &self.encoded
    }

    /// Iterate over the validated UTF-16 code units without allocating.
    pub fn units(&self) -> impl Iterator<Item = u16> + '_ {
        encoded_units(&self.encoded)
    }

    /// Decode the string into owned Rust text.
    pub fn text(&self) -> String {
        let units = self.units().collect::<Vec<_>>();
        String::from_utf16_lossy(&units)
    }

    /// Copy a borrowed string into an owned representation.
    pub fn into_owned(self) -> WString<'static> {
        WString {
            encoded: Cow::Owned(self.encoded.into_owned()),
        }
    }

    pub(crate) fn encoded_len(&self) -> usize {
        self.encoded.len()
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

/// The two wire-defined toolbar kinds in `TBTRFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Type {
    /// A basic toolbar.
    Basic,
    /// A menu toolbar.
    Menu,
    /// An unrecognized future wire value.
    Unknown(u8),
}

impl Type {
    /// Return the raw eight-bit toolbar kind.
    pub const fn raw(self) -> u8 {
        match self {
            Self::Basic => 0,
            Self::Menu => 2,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value {
            0 => Self::Basic,
            2 => Self::Menu,
            value => Self::Unknown(value),
        }
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if matches!(self, Self::Unknown(_)) {
            return Err(Error::invalid(format!(
                "unsupported toolbar type 0x{:02X}",
                self.raw()
            )));
        }
        Ok(())
    }
}

impl TryFrom<u8> for Type {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let kind = Self::from_raw(value);
        kind.validate()?;
        Ok(kind)
    }
}

/// Toolbar type and restriction bits from `TBTRFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Restrictions {
    raw: u32,
}

impl Restrictions {
    const RESERVED_MASK: u32 = 0xFDFF_F600;

    /// Construct the canonical default restrictions for a toolbar kind.
    pub fn new(kind: Type) -> Result<Self, Error> {
        let mut raw = (kind.raw() as u32) << 24;
        if matches!(kind, Type::Menu) {
            raw |= (1 << 1) | (1 << 2) | (1 << 4) | (1 << 7) | (1 << 8) | (1 << 11);
        }
        let value = Self { raw };
        value.validate()?;
        Ok(value)
    }

    /// Retain a raw wire value, including reserved bits.
    pub const fn from_raw(raw: u32) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking required invariants.
    pub fn try_from_raw(raw: u32) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag word.
    pub const fn raw(self) -> u32 {
        self.raw
    }

    /// Return all reserved bits exactly as they appeared on the wire.
    pub const fn reserved_bits(self) -> u32 {
        self.raw & Self::RESERVED_MASK
    }

    /// Return the toolbar kind encoded in the most significant byte.
    pub const fn toolbar_type(self) -> Type {
        Type::from_raw((self.raw >> 24) as u8)
    }

    pub const fn no_add_delete_control(self) -> bool {
        self.raw & (1 << 0) != 0
    }

    pub const fn no_resize(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    pub const fn no_move(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn no_change_visible(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub const fn no_change_dock(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    pub const fn no_vertical_dock(self) -> bool {
        self.raw & (1 << 5) != 0
    }

    pub const fn no_horizontal_dock(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    pub const fn no_border(self) -> bool {
        self.raw & (1 << 7) != 0
    }

    pub const fn no_context_menu(self) -> bool {
        self.raw & (1 << 8) != 0
    }

    pub const fn not_top_level(self) -> bool {
        self.raw & (1 << 11) != 0
    }

    pub const fn popup_menu(self) -> bool {
        self.raw & (1 << 25) != 0
    }

    pub fn with_no_add_delete_control(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 0, value);
        self
    }

    pub fn with_no_resize(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 1, value);
        self
    }

    pub fn with_no_move(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_no_change_visible(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    pub fn with_no_change_dock(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

    pub fn with_no_vertical_dock(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 5, value);
        self
    }

    pub fn with_no_horizontal_dock(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 6, value);
        self
    }

    pub fn with_no_border(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 7, value);
        self
    }

    pub fn with_no_context_menu(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 8, value);
        self
    }

    pub fn with_not_top_level(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 11, value);
        self
    }

    pub fn with_popup_menu(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 25, value);
        self
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        self.toolbar_type().validate()?;
        if self.reserved_bits() != 0 {
            return Err(Error::invalid("TBTRFlags has nonzero reserved bits"));
        }
        let popup = self.popup_menu();
        if self.no_resize() != popup
            || self.no_move() != popup
            || self.no_change_dock() != popup
            || self.no_border() != popup
            || self.no_context_menu() != popup
            || self.not_top_level() != popup
        {
            return Err(Error::invalid(
                "TBTRFlags popup restrictions do not match TBTPopupMenu",
            ));
        }
        Ok(())
    }
}

/// Toolbar flags from `TBFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Flags {
    raw: u16,
}

impl Flags {
    const RESERVED_MASK: u16 = 0xFFE2;

    /// Retain a raw wire value, including undefined and reserved bits.
    pub const fn from_raw(raw: u16) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking required zero bits.
    pub fn try_from_raw(raw: u16) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag word.
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Return all reserved or undefined bits exactly as stored.
    pub const fn reserved_bits(self) -> u16 {
        self.raw & Self::RESERVED_MASK
    }

    pub const fn disabled(self) -> bool {
        self.raw & (1 << 0) != 0
    }

    pub const fn controls_modified(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn no_adaptive_menus(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub const fn needs_positioning(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    pub fn with_disabled(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 0, value);
        self
    }

    pub fn with_controls_modified(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_no_adaptive_menus(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    pub fn with_needs_positioning(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if self.raw & 0xFFE0 != 0 {
            return Err(Error::invalid("TBFlags has nonzero reserved bits"));
        }
        Ok(())
    }
}

/// Toolbar-control flags from `TBCFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ControlFlags {
    raw: u8,
}

impl ControlFlags {
    const RESERVED_MASK: u8 = 0xA0;

    /// Retain a raw wire value, including undefined and reserved bits.
    pub const fn from_raw(raw: u8) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking the required zero bit.
    pub fn try_from_raw(raw: u8) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag byte.
    pub const fn raw(self) -> u8 {
        self.raw
    }

    /// Return all reserved or undefined bits exactly as stored.
    pub const fn reserved_bits(self) -> u8 {
        self.raw & Self::RESERVED_MASK
    }

    pub const fn hidden(self) -> bool {
        self.raw & (1 << 0) != 0
    }

    pub const fn begin_group(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    pub const fn own_line(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn no_customize(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub const fn save_dimensions(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    pub const fn begin_line(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    pub fn with_hidden(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 0, value);
        self
    }

    pub fn with_begin_group(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 1, value);
        self
    }

    pub fn with_own_line(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_no_customize(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    pub fn with_save_dimensions(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

    pub fn with_begin_line(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 6, value);
        self
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if self.raw & 0x80 != 0 {
            return Err(Error::invalid("TBCFlags reserved2 must be zero"));
        }
        Ok(())
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

/// Toolbar-control settings from `TBCSFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SpecificFlags {
    raw: u32,
}

impl SpecificFlags {
    const UNUSED_MASK: u32 = 0x7900_FF00;
    const RESERVED_MASK: u32 = 0x8000_0000;

    /// Retain a raw wire value, including all unused and reserved bits.
    pub const fn from_raw(raw: u32) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking the required zero bit.
    pub fn try_from_raw(raw: u32) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag word.
    pub const fn raw(self) -> u32 {
        self.raw
    }

    /// Return all unused bits exactly as stored.
    pub const fn unused_bits(self) -> u32 {
        self.raw & Self::UNUSED_MASK
    }

    /// Return the required-zero reserved bit exactly as stored.
    pub const fn reserved_bits(self) -> u32 {
        self.raw & Self::RESERVED_MASK
    }

    pub const fn text_icon(self) -> TextIcon {
        TextIcon::from_raw(self.raw as u8)
    }

    pub const fn owner_draw(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn allow_resize(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub const fn one_state(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    pub const fn no_set_cursor(self) -> bool {
        self.raw & (1 << 5) != 0
    }

    pub const fn no_accelerator(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    pub const fn change_accelerator(self) -> bool {
        self.raw & (1 << 7) != 0
    }

    pub const fn always_enabled(self) -> bool {
        self.raw & (1 << 16) != 0
    }

    pub const fn always_visible(self) -> bool {
        self.raw & (1 << 17) != 0
    }

    pub const fn no_change_label(self) -> bool {
        self.raw & (1 << 18) != 0
    }

    pub const fn keep_label(self) -> bool {
        self.raw & (1 << 19) != 0
    }

    pub const fn no_query_tooltip(self) -> bool {
        self.raw & (1 << 20) != 0
    }

    pub const fn save_ui_strings(self) -> bool {
        self.raw & (1 << 21) != 0
    }

    pub const fn exclusive_popup(self) -> bool {
        self.raw & (1 << 22) != 0
    }

    pub const fn default_behavior(self) -> bool {
        self.raw & (1 << 23) != 0
    }

    pub const fn wrap_text(self) -> bool {
        self.raw & (1 << 25) != 0
    }

    pub const fn text_below(self) -> bool {
        self.raw & (1 << 26) != 0
    }

    pub fn with_text_icon(mut self, value: TextIcon) -> Self {
        self.raw = (self.raw & !0x03) | value.raw() as u32;
        self
    }

    pub fn with_owner_draw(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_allow_resize(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    pub fn with_one_state(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

    pub fn with_no_set_cursor(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 5, value);
        self
    }

    pub fn with_no_accelerator(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 6, value);
        self
    }

    pub fn with_change_accelerator(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 7, value);
        self
    }

    pub fn with_always_enabled(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 16, value);
        self
    }

    pub fn with_always_visible(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 17, value);
        self
    }

    pub fn with_no_change_label(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 18, value);
        self
    }

    pub fn with_keep_label(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 19, value);
        self
    }

    pub fn with_no_query_tooltip(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 20, value);
        self
    }

    pub fn with_save_ui_strings(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 21, value);
        self
    }

    pub fn with_exclusive_popup(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 22, value);
        self
    }

    pub fn with_default_behavior(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 23, value);
        self
    }

    pub fn with_wrap_text(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 25, value);
        self
    }

    pub fn with_text_below(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 26, value);
        self
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if self.reserved_bits() != 0 {
            return Err(Error::invalid("TBCSFlags reserved1 must be zero"));
        }
        Ok(())
    }
}

/// Toolbar-control general-information flags from `TBCGIFlags`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct GeneralFlags {
    raw: u8,
}

impl GeneralFlags {
    const UNUSED_MASK: u8 = 0xF0;

    /// Retain a raw wire value, including unused bits.
    pub const fn from_raw(raw: u8) -> Self {
        Self { raw }
    }

    /// Return the exact serialized flag byte.
    pub const fn raw(self) -> u8 {
        self.raw
    }

    /// Return unused bits exactly as stored.
    pub const fn unused_bits(self) -> u8 {
        self.raw & Self::UNUSED_MASK
    }

    pub const fn save_text(self) -> bool {
        self.raw & (1 << 0) != 0
    }

    pub const fn save_misc_ui_strings(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    pub const fn save_misc_custom(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn disabled(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub fn with_save_text(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 0, value);
        self
    }

    pub fn with_save_misc_ui_strings(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 1, value);
        self
    }

    pub fn with_save_misc_custom(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_disabled(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
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
            value => Self::Unknown(value),
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
            value => Self::Unknown(value),
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
    pub const fn from_raw(raw: u8) -> Self {
        Self { raw }
    }

    /// Construct from a raw value after checking the two reserved
    /// discriminants.  The final reserved bit is advisory in [MS-OSHARED]
    /// and is therefore retained without rejection.
    pub fn try_from_raw(raw: u8) -> Result<Self, Error> {
        let value = Self::from_raw(raw);
        value.validate()?;
        Ok(value)
    }

    /// Return the exact serialized flag byte.
    pub const fn raw(self) -> u8 {
        self.raw
    }

    /// Return the advisory reserved bit exactly as stored.
    pub const fn reserved_bits(self) -> u8 {
        self.raw & 0x80
    }

    pub const fn state(self) -> ButtonState {
        ButtonState::from_raw(self.raw)
    }

    pub const fn accelerator(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    pub const fn custom_bitmap(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    pub const fn custom_button_face(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    pub const fn hyperlink(self) -> HyperlinkType {
        HyperlinkType::from_raw(self.raw >> 5)
    }

    pub fn with_state(mut self, value: ButtonState) -> Self {
        self.raw = (self.raw & !0x03) | value.raw();
        self
    }

    pub fn with_accelerator(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 2, value);
        self
    }

    pub fn with_custom_bitmap(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 3, value);
        self
    }

    pub fn with_custom_button_face(mut self, value: bool) -> Self {
        self.raw = set_bit(self.raw, 4, value);
        self
    }

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

/// Optional width and height in a `ControlHeader`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Dimensions {
    width: u16,
    height: u16,
}

impl Dimensions {
    /// Construct a pair of unsigned pixel dimensions.
    pub const fn new(width: u16, height: u16) -> Self {
        Self { width, height }
    }

    /// Return the width in pixels.
    pub const fn width(self) -> u16 {
        self.width
    }

    /// Return the height in pixels.
    pub const fn height(self) -> u16 {
        self.height
    }
}

/// The toolbar-control types listed by `TBCHeader.tct`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ControlType {
    /// A push button.
    Button,
    /// An edit control.
    Edit,
    /// A drop-down control.
    DropDown,
    /// A combo box.
    ComboBox,
    /// A split drop-down.
    SplitDropDown,
    /// An OCX drop-down.
    OcxDropDown,
    /// A graphic drop-down.
    GraphicDropDown,
    /// A popup.
    Popup,
    /// A button popup.
    ButtonPopup,
    /// A split button popup.
    SplitButtonPopup,
    /// A split button MRU popup.
    SplitButtonMruPopup,
    /// A label.
    Label,
    /// An expanding grid.
    ExpandingGrid,
    /// A grid.
    Grid,
    /// A gauge.
    Gauge,
    /// A graphic combo.
    GraphicCombo,
    /// A pane.
    Pane,
    /// An ActiveX control.
    ActiveX,
    /// An unrecognized future wire value.
    Unknown(u8),
}

impl ControlType {
    pub const fn raw(self) -> u8 {
        match self {
            Self::Button => 0x01,
            Self::Edit => 0x02,
            Self::DropDown => 0x03,
            Self::ComboBox => 0x04,
            Self::SplitDropDown => 0x06,
            Self::OcxDropDown => 0x07,
            Self::GraphicDropDown => 0x09,
            Self::Popup => 0x0A,
            Self::ButtonPopup => 0x0C,
            Self::SplitButtonPopup => 0x0D,
            Self::SplitButtonMruPopup => 0x0E,
            Self::Label => 0x0F,
            Self::ExpandingGrid => 0x10,
            Self::Grid => 0x12,
            Self::Gauge => 0x13,
            Self::GraphicCombo => 0x14,
            Self::Pane => 0x15,
            Self::ActiveX => 0x16,
            Self::Unknown(value) => value,
        }
    }

    pub(crate) const fn from_raw(value: u8) -> Self {
        match value {
            0x01 => Self::Button,
            0x02 => Self::Edit,
            0x03 => Self::DropDown,
            0x04 => Self::ComboBox,
            0x06 => Self::SplitDropDown,
            0x07 => Self::OcxDropDown,
            0x09 => Self::GraphicDropDown,
            0x0A => Self::Popup,
            0x0C => Self::ButtonPopup,
            0x0D => Self::SplitButtonPopup,
            0x0E => Self::SplitButtonMruPopup,
            0x0F => Self::Label,
            0x10 => Self::ExpandingGrid,
            0x12 => Self::Grid,
            0x13 => Self::Gauge,
            0x14 => Self::GraphicCombo,
            0x15 => Self::Pane,
            0x16 => Self::ActiveX,
            value => Self::Unknown(value),
        }
    }

    pub(crate) fn validate(self) -> Result<(), Error> {
        if matches!(self, Self::Unknown(_)) {
            return Err(Error::invalid(format!(
                "unsupported toolbar-control type 0x{:02X}",
                self.raw()
            )));
        }
        Ok(())
    }
}

impl TryFrom<u8> for ControlType {
    type Error = Error;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        let kind = Self::from_raw(value);
        kind.validate()?;
        Ok(kind)
    }
}

/// The fixed and optional fields of `TBCHeader`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlHeader {
    control_type: ControlType,
    control_id: u16,
    flags: ControlFlags,
    specifics: SpecificFlags,
    priority: u8,
    dimensions: Option<Dimensions>,
}

impl ControlHeader {
    /// Construct a validated toolbar-control header.
    pub fn new(
        control_type: ControlType,
        control_id: u16,
        flags: ControlFlags,
        specifics: SpecificFlags,
        priority: u8,
        dimensions: Option<Dimensions>,
    ) -> Result<Self, Error> {
        let value = Self {
            control_type,
            control_id,
            flags,
            specifics,
            priority,
            dimensions,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        control_type: ControlType,
        control_id: u16,
        flags: ControlFlags,
        specifics: SpecificFlags,
        priority: u8,
        dimensions: Option<Dimensions>,
    ) -> Self {
        Self {
            control_type,
            control_id,
            flags,
            specifics,
            priority,
            dimensions,
        }
    }

    /// Return the toolbar-control type.
    pub const fn control_type(&self) -> ControlType {
        self.control_type
    }

    /// Return the format-specific toolbar-control identifier.
    pub const fn control_id(&self) -> u16 {
        self.control_id
    }

    /// Return the general toolbar-control flags.
    pub const fn flags(&self) -> ControlFlags {
        self.flags
    }

    /// Return the toolbar-control settings flags.
    pub const fn specifics(&self) -> SpecificFlags {
        self.specifics
    }

    /// Return the drop and wrap priority.
    pub const fn priority(&self) -> u8 {
        self.priority
    }

    /// Return optional saved dimensions.
    pub const fn dimensions(&self) -> Option<Dimensions> {
        self.dimensions
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        self.control_type.validate()?;
        self.flags.validate()?;
        self.specifics.validate()?;
        if self.priority > 7 {
            return Err(Error::invalid("TBCHeader priority exceeds 7"));
        }
        if self.flags.save_dimensions() != self.dimensions.is_some() {
            return Err(Error::invalid("TBCHeader dimensions must match fSaveDxy"));
        }
        Ok(())
    }
}

/// A `TB` toolbar header and its name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header<'a> {
    control_count: i16,
    restrictions: Restrictions,
    rows_default: u16,
    flags: Flags,
    name: WString<'a>,
}

impl<'a> Header<'a> {
    /// Construct a validated toolbar header.
    pub fn new(
        control_count: i16,
        restrictions: Restrictions,
        rows_default: u16,
        flags: Flags,
        name: WString<'a>,
    ) -> Result<Self, Error> {
        let value = Self {
            control_count,
            restrictions,
            rows_default,
            flags,
            name,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn from_decoded(
        control_count: i16,
        restrictions: Restrictions,
        rows_default: u16,
        flags: Flags,
        name: WString<'a>,
    ) -> Self {
        Self {
            control_count,
            restrictions,
            rows_default,
            flags,
            name,
        }
    }

    /// Return the signed `cCL` field exactly as decoded.
    pub const fn control_count(&self) -> i16 {
        self.control_count
    }

    /// Return toolbar type and restriction flags.
    pub const fn restrictions(&self) -> Restrictions {
        self.restrictions
    }

    /// Return the preferred row count exactly as decoded.
    pub const fn rows_default(&self) -> u16 {
        self.rows_default
    }

    /// Return toolbar flags.
    pub const fn flags(&self) -> Flags {
        self.flags
    }

    /// Return the borrowed or owned toolbar name.
    pub const fn name(&self) -> &WString<'a> {
        &self.name
    }

    /// Move a decoded toolbar header into an owned representation.
    ///
    /// This is used by format facades whose compound-file stream buffer is
    /// shorter-lived than the public workbook/document object.
    pub fn into_owned(self) -> Header<'static> {
        Header {
            control_count: self.control_count,
            restrictions: self.restrictions,
            rows_default: self.rows_default,
            flags: self.flags,
            name: self.name.into_owned(),
        }
    }

    pub(crate) fn validate(&self) -> Result<(), Error> {
        if self.control_count < 0 {
            return Err(Error::invalid("TB cCL cannot be negative"));
        }
        if self.rows_default > u16::from(u8::MAX) {
            return Err(Error::invalid("TB cRowsDefault exceeds 255"));
        }
        self.restrictions.validate()?;
        self.flags.validate()?;
        Ok(())
    }
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

impl BitOps for u16 {
    fn one() -> Self {
        1
    }
}

impl BitOps for u32 {
    fn one() -> Self {
        1
    }
}

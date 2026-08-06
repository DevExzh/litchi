use super::{Error, TextIcon};

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

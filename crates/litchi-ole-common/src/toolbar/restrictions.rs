use super::Error;

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

impl BitOps for u32 {
    fn one() -> Self {
        1
    }
}

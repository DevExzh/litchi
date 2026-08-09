/// A `CrtLayout12Mode` layout mode (MS-XLS 2.5.62): the meaning of the `x`,
/// `y`, `dx`, and `dy` fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum CrtLayout12Mode {
    /// 0x0000: position and dimension are determined by the application.
    Auto = 0x0000,
    /// 0x0001: `x`/`y` are offsets from the default position, `dx`/`dy` are
    /// dimensions, all as fractions of the chart area.
    Factor = 0x0001,
    /// 0x0002: `x`/`y` are the upper-left corner, `dx`/`dy` the bottom-right
    /// corner, all as fractions of the chart area.
    Edge = 0x0002,
}

/// Mask of the 4-bit `autolayouttype` field of `CrtLayout12` (MS-XLS 2.4.66).
const AUTO_LAYOUT_TYPE_MASK: u16 = 0x001E;
/// `CrtLayout12A` flag: the layout target is the inner plot area
/// (MS-XLS 2.4.67).
const LAYOUT_TARGET_INNER: u16 = 0x0001;

/// The four layout modes and values shared by `CrtLayout12` and
/// `CrtLayout12A`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct LayoutModes {
    pub(super) x_mode: CrtLayout12Mode,
    pub(super) y_mode: CrtLayout12Mode,
    pub(super) width_mode: CrtLayout12Mode,
    pub(super) height_mode: CrtLayout12Mode,
    /// Raw `x`/`y`/`dx`/`dy` Xnum bit patterns (MS-XLS 2.5.342).
    pub(super) x: f64,
    pub(super) y: f64,
    pub(super) dx: f64,
    pub(super) dy: f64,
}

/// Typed `CrtLayout12` record content (MS-XLS 2.4.66): layout information
/// for an attached label or legend.
///
/// The unused flag bit, the 11 `reserved1` bits, and the trailing `reserved2`
/// field (MUST be ignored) are preserved verbatim so the record round-trips
/// unchanged.
#[derive(Debug, Clone, PartialEq)]
pub struct CrtLayout12 {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub(super) frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    pub(super) frt_reserved: [u8; 8],
    /// Raw `dwCheckSum` of the layout values, preserved verbatim.
    pub(super) checksum: u32,
    /// Raw flags: unused bit, 4-bit `autolayouttype`, and 11 `reserved1` bits.
    pub(super) flags: u16,
    pub(super) modes: LayoutModes,
    /// Trailing `reserved2` field, preserved verbatim.
    pub(super) reserved2: u16,
}

impl CrtLayout12 {
    /// The automatic layout type of the legend (`autolayouttype`, 4 bits).
    /// MUST be ignored when the record is in an ATTACHEDLABEL rule sequence
    /// (MS-XLS 2.4.66); defined values are 0x0 through 0x4.
    #[must_use]
    pub fn auto_layout_type(&self) -> u8 {
        ((self.flags & AUTO_LAYOUT_TYPE_MASK) >> 1) as u8
    }

    /// Raw flags word, including the unused and `reserved1` bits.
    #[must_use]
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Raw `dwCheckSum` value, preserved verbatim.
    #[must_use]
    pub fn checksum(&self) -> u32 {
        self.checksum
    }

    /// Layout mode of `x` (`wXMode`).
    #[must_use]
    pub fn x_mode(&self) -> CrtLayout12Mode {
        self.modes.x_mode
    }

    /// Layout mode of `y` (`wYMode`).
    #[must_use]
    pub fn y_mode(&self) -> CrtLayout12Mode {
        self.modes.y_mode
    }

    /// Layout mode of `dx` (`wWidthMode`).
    #[must_use]
    pub fn width_mode(&self) -> CrtLayout12Mode {
        self.modes.width_mode
    }

    /// Layout mode of `dy` (`wHeightMode`).
    #[must_use]
    pub fn height_mode(&self) -> CrtLayout12Mode {
        self.modes.height_mode
    }

    /// Horizontal offset (`x`), interpreted per [`Self::x_mode`].
    #[must_use]
    pub fn x(&self) -> f64 {
        self.modes.x
    }

    /// Vertical offset (`y`), interpreted per [`Self::y_mode`].
    #[must_use]
    pub fn y(&self) -> f64 {
        self.modes.y
    }

    /// Width or horizontal offset (`dx`), interpreted per [`Self::width_mode`].
    #[must_use]
    pub fn dx(&self) -> f64 {
        self.modes.dx
    }

    /// Height or vertical offset (`dy`), interpreted per [`Self::height_mode`].
    #[must_use]
    pub fn dy(&self) -> f64 {
        self.modes.dy
    }
}

/// Typed `CrtLayout12A` record content (MS-XLS 2.4.67): layout information
/// for a plot area.
///
/// The 15 `reserved1` bits and the trailing `reserved2` field (MUST be
/// ignored) are preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq)]
pub struct CrtLayout12A {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub(super) frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    pub(super) frt_reserved: [u8; 8],
    /// `dwCheckSum`: 0x00000000 or 0x00000001 (MS-XLS 2.4.67).
    pub(super) checksum: u32,
    /// Raw flags: `fLayoutTargetInner` and 15 `reserved1` bits.
    pub(super) flags: u16,
    /// `xTL`: horizontal offset of the plot area's upper-left corner, in SPRC.
    pub(super) x_top_left: i16,
    /// `yTL`: vertical offset of the plot area's upper-left corner, in SPRC.
    pub(super) y_top_left: i16,
    /// `xBR`: width of the plot area, in SPRC.
    pub(super) x_bottom_right: i16,
    /// `yBR`: height of the plot area, in SPRC.
    pub(super) y_bottom_right: i16,
    pub(super) modes: LayoutModes,
    /// Trailing `reserved2` field, preserved verbatim.
    pub(super) reserved2: u16,
}

impl CrtLayout12A {
    /// The `dwCheckSum` value: 0x00000001 when the plot area layout is manual
    /// and not always automatically computed, 0x00000000 otherwise (derived
    /// from the `ShtProps` flags, MS-XLS 2.4.67).
    #[must_use]
    pub fn checksum(&self) -> u32 {
        self.checksum
    }

    /// Whether the layout target is the inner plot area (`fLayoutTargetInner`).
    #[must_use]
    pub fn is_layout_target_inner(&self) -> bool {
        self.flags & LAYOUT_TARGET_INNER != 0
    }

    /// Raw flags word, including the 15 `reserved1` bits.
    #[must_use]
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// Horizontal offset of the plot area's upper-left corner (`xTL`), in SPRC.
    #[must_use]
    pub fn x_top_left(&self) -> i16 {
        self.x_top_left
    }

    /// Vertical offset of the plot area's upper-left corner (`yTL`), in SPRC.
    #[must_use]
    pub fn y_top_left(&self) -> i16 {
        self.y_top_left
    }

    /// Width of the plot area (`xBR`), in SPRC.
    #[must_use]
    pub fn x_bottom_right(&self) -> i16 {
        self.x_bottom_right
    }

    /// Height of the plot area (`yBR`), in SPRC.
    #[must_use]
    pub fn y_bottom_right(&self) -> i16 {
        self.y_bottom_right
    }

    /// Layout mode of `x` (`wXMode`).
    #[must_use]
    pub fn x_mode(&self) -> CrtLayout12Mode {
        self.modes.x_mode
    }

    /// Layout mode of `y` (`wYMode`).
    #[must_use]
    pub fn y_mode(&self) -> CrtLayout12Mode {
        self.modes.y_mode
    }

    /// Layout mode of `dx` (`wWidthMode`).
    #[must_use]
    pub fn width_mode(&self) -> CrtLayout12Mode {
        self.modes.width_mode
    }

    /// Layout mode of `dy` (`wHeightMode`).
    #[must_use]
    pub fn height_mode(&self) -> CrtLayout12Mode {
        self.modes.height_mode
    }

    /// Horizontal offset (`x`), interpreted per [`Self::x_mode`].
    #[must_use]
    pub fn x(&self) -> f64 {
        self.modes.x
    }

    /// Vertical offset (`y`), interpreted per [`Self::y_mode`].
    #[must_use]
    pub fn y(&self) -> f64 {
        self.modes.y
    }

    /// Width or horizontal offset (`dx`), interpreted per [`Self::width_mode`].
    #[must_use]
    pub fn dx(&self) -> f64 {
        self.modes.dx
    }

    /// Height or vertical offset (`dy`), interpreted per [`Self::height_mode`].
    #[must_use]
    pub fn dy(&self) -> f64 {
        self.modes.dy
    }
}

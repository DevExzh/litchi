//! BIFF8 workbook color palette support.

use super::error::{Error, Result};
use super::leniency::{PaletteDefect, ToleranceReport};

const PALETTE_COLOR_COUNT: usize = 56;
const FIRST_PALETTE_INDEX: u16 = 0x0008;
const LAST_PALETTE_INDEX: u16 = 0x003f;
const PALETTE_RECORD_LENGTH: usize = 2 + PALETTE_COLOR_COUNT * 4;

/// An RGB color from a BIFF8 workbook color table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Color {
    red: u8,
    green: u8,
    blue: u8,
}

impl Color {
    const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    /// Red component.
    #[must_use]
    pub fn red(self) -> u8 {
        self.red
    }

    /// Green component.
    #[must_use]
    pub fn green(self) -> u8 {
        self.green
    }

    /// Blue component.
    #[must_use]
    pub fn blue(self) -> u8 {
        self.blue
    }
}

const BUILT_IN_COLORS: [Color; 8] = [
    Color::new(0x00, 0x00, 0x00),
    Color::new(0xff, 0xff, 0xff),
    Color::new(0xff, 0x00, 0x00),
    Color::new(0x00, 0xff, 0x00),
    Color::new(0x00, 0x00, 0xff),
    Color::new(0xff, 0xff, 0x00),
    Color::new(0xff, 0x00, 0xff),
    Color::new(0x00, 0xff, 0xff),
];

const DEFAULT_PALETTE: [Color; PALETTE_COLOR_COUNT] = [
    Color::new(0x00, 0x00, 0x00),
    Color::new(0xff, 0xff, 0xff),
    Color::new(0xff, 0x00, 0x00),
    Color::new(0x00, 0xff, 0x00),
    Color::new(0x00, 0x00, 0xff),
    Color::new(0xff, 0xff, 0x00),
    Color::new(0xff, 0x00, 0xff),
    Color::new(0x00, 0xff, 0xff),
    Color::new(0x80, 0x00, 0x00),
    Color::new(0x00, 0x80, 0x00),
    Color::new(0x00, 0x00, 0x80),
    Color::new(0x80, 0x80, 0x00),
    Color::new(0x80, 0x00, 0x80),
    Color::new(0x00, 0x80, 0x80),
    Color::new(0xc0, 0xc0, 0xc0),
    Color::new(0x80, 0x80, 0x80),
    Color::new(0x99, 0x99, 0xff),
    Color::new(0x99, 0x33, 0x66),
    Color::new(0xff, 0xff, 0xcc),
    Color::new(0xcc, 0xff, 0xff),
    Color::new(0x66, 0x00, 0x66),
    Color::new(0xff, 0x80, 0x80),
    Color::new(0x00, 0x66, 0xcc),
    Color::new(0xcc, 0xcc, 0xff),
    Color::new(0x00, 0x00, 0x80),
    Color::new(0xff, 0x00, 0xff),
    Color::new(0xff, 0xff, 0x00),
    Color::new(0x00, 0xff, 0xff),
    Color::new(0x80, 0x00, 0x80),
    Color::new(0x80, 0x00, 0x00),
    Color::new(0x00, 0x80, 0x80),
    Color::new(0x00, 0x00, 0xff),
    Color::new(0x00, 0xcc, 0xff),
    Color::new(0xcc, 0xff, 0xff),
    Color::new(0xcc, 0xff, 0xcc),
    Color::new(0xff, 0xff, 0x99),
    Color::new(0x99, 0xcc, 0xff),
    Color::new(0xff, 0x99, 0xcc),
    Color::new(0xcc, 0x99, 0xff),
    Color::new(0xff, 0xcc, 0x99),
    Color::new(0x33, 0x66, 0xff),
    Color::new(0x33, 0xcc, 0xcc),
    Color::new(0x99, 0xcc, 0x00),
    Color::new(0xff, 0xcc, 0x00),
    Color::new(0xff, 0x99, 0x00),
    Color::new(0xff, 0x66, 0x00),
    Color::new(0x66, 0x66, 0x99),
    Color::new(0x96, 0x96, 0x96),
    Color::new(0x00, 0x33, 0x66),
    Color::new(0x33, 0x99, 0x66),
    Color::new(0x00, 0x33, 0x00),
    Color::new(0x33, 0x33, 0x00),
    Color::new(0x99, 0x33, 0x00),
    Color::new(0x99, 0x33, 0x66),
    Color::new(0x33, 0x33, 0x99),
    Color::new(0x33, 0x33, 0x33),
];

/// The 56-color BIFF8 workbook palette.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Palette {
    colors: [Color; PALETTE_COLOR_COUNT],
    custom: bool,
}

impl Default for Palette {
    fn default() -> Self {
        Self {
            colors: DEFAULT_PALETTE,
            custom: false,
        }
    }
}

impl Palette {
    /// Resolve a built-in or workbook palette color-table index.
    #[must_use]
    pub fn color(&self, index: u16) -> Option<Color> {
        if index < FIRST_PALETTE_INDEX {
            return BUILT_IN_COLORS.get(usize::from(index)).copied();
        }
        if index > LAST_PALETTE_INDEX {
            return None;
        }
        self.colors
            .get(usize::from(index - FIRST_PALETTE_INDEX))
            .copied()
    }

    /// All 56 palette entries corresponding to color indices `0x08..=0x3f`.
    #[must_use]
    pub fn palette_colors(&self) -> &[Color; PALETTE_COLOR_COUNT] {
        &self.colors
    }

    /// Whether the workbook contained a custom `Palette` record.
    #[must_use]
    pub fn is_custom(&self) -> bool {
        self.custom
    }

    /// Parse a conforming `Palette` payload, refusing every deviation.
    #[cfg(test)]
    pub(crate) fn parse_record(data: &[u8]) -> Result<Self> {
        let mut report = ToleranceReport::default();
        let palette = Self::parse_tolerant(data, 0, &mut report)?;
        if report.is_clean() {
            Ok(palette)
        } else {
            Err(Error::InvalidData(
                "Palette record is non-conforming".to_string(),
            ))
        }
    }

    /// Parse a `Palette` payload, applying the Excel-compatible repairs.
    ///
    /// A nonzero `LongRGB.reserved` byte is ignored ([MS-XLS] 2.5.177: "MUST
    /// be zero, and MUST be ignored"), and a `ccv` other than 56 is accepted
    /// when the payload holds exactly `ccv` entries: the first 56 are used
    /// and any missing entry keeps its default color, as LibreOffice and POI
    /// read it. Each repair is recorded against `record_offset`. A payload
    /// whose length contradicts its own `ccv` is still malformed framing.
    pub(crate) fn parse_tolerant(
        data: &[u8],
        record_offset: u32,
        tolerance: &mut ToleranceReport,
    ) -> Result<Self> {
        if data.len() < 2 {
            return Err(Error::InvalidData(format!(
                "Palette record has {} bytes; expected {PALETTE_RECORD_LENGTH}",
                data.len()
            )));
        }
        let color_count = i16::from_le_bytes([data[0], data[1]]);
        let declared = usize::try_from(color_count).ok();
        if declared.is_none_or(|count| data.len() != 2 + count * 4) {
            return Err(Error::InvalidData(format!(
                "Palette record has {} bytes for {color_count} colors",
                data.len()
            )));
        }
        let declared = declared.unwrap_or_default();
        if declared != PALETTE_COLOR_COUNT {
            tolerance.record(
                PaletteDefect::ColorCount,
                record_offset,
                u32::from(color_count.unsigned_abs()),
            );
        }

        let mut colors = DEFAULT_PALETTE;
        let mut nonzero_reserved = 0_u32;
        for (color, entry) in colors.iter_mut().zip(data[2..].chunks_exact(4)) {
            if entry[3] != 0 {
                nonzero_reserved += 1;
            }
            *color = Color::new(entry[0], entry[1], entry[2]);
        }
        if nonzero_reserved != 0 {
            tolerance.record(PaletteDefect::ReservedByte, record_offset, nonzero_reserved);
        }
        Ok(Self {
            colors,
            custom: true,
        })
    }

    /// Parse one of possibly several `Palette` records. A repeat replaces
    /// the earlier palette — the last record wins, as in Excel and
    /// LibreOffice — and is recorded as [`PaletteDefect::Duplicate`].
    pub(crate) fn parse_unique_record(
        data: &[u8],
        record_offset: u32,
        seen: &mut u32,
        tolerance: &mut ToleranceReport,
    ) -> Result<Self> {
        let palette = Self::parse_tolerant(data, record_offset, tolerance)?;
        *seen = seen.saturating_add(1);
        if *seen > 1 {
            tolerance.record(PaletteDefect::Duplicate, record_offset, *seen);
        }
        Ok(palette)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn palette_payload() -> Vec<u8> {
        let mut data = Vec::with_capacity(PALETTE_RECORD_LENGTH);
        data.extend_from_slice(&(PALETTE_COLOR_COUNT as i16).to_le_bytes());
        for color in DEFAULT_PALETTE {
            data.extend_from_slice(&[color.red, color.green, color.blue, 0]);
        }
        data
    }

    #[test]
    fn parses_custom_palette_and_resolves_icv_indices() {
        let mut data = palette_payload();
        let first_offset = 2 + usize::from(0x12 - FIRST_PALETTE_INDEX) * 4;
        data[first_offset..first_offset + 3].copy_from_slice(&[101, 230, 100]);
        let second_offset = 2 + usize::from(0x3b - FIRST_PALETTE_INDEX) * 4;
        data[second_offset..second_offset + 3].copy_from_slice(&[0, 255, 52]);

        let palette = Palette::parse_record(&data).unwrap();
        assert!(palette.is_custom());
        assert_eq!(palette.color(0x12), Some(Color::new(101, 230, 100)));
        assert_eq!(palette.color(0x3b), Some(Color::new(0, 255, 52)));
        assert_eq!(palette.color(0x02), Some(Color::new(255, 0, 0)));
        assert_eq!(palette.color(0x40), None);
    }

    #[test]
    fn supplies_normative_default_palette() {
        let palette = Palette::default();
        assert!(!palette.is_custom());
        assert_eq!(palette.color(0x08), Some(Color::new(0, 0, 0)));
        assert_eq!(palette.color(0x0d), Some(Color::new(255, 255, 0)));
        assert_eq!(palette.color(0x3f), Some(Color::new(51, 51, 51)));
    }

    #[test]
    fn rejects_record_lengths_that_contradict_the_count() {
        let mut truncated = palette_payload();
        truncated.pop();
        assert!(Palette::parse_tolerant(&truncated, 0, &mut ToleranceReport::default()).is_err());

        let mut extra = palette_payload();
        extra.push(0);
        assert!(Palette::parse_tolerant(&extra, 0, &mut ToleranceReport::default()).is_err());

        let mut negative = palette_payload();
        negative[..2].copy_from_slice(&(-1i16).to_le_bytes());
        assert!(Palette::parse_tolerant(&negative, 0, &mut ToleranceReport::default()).is_err());
        assert!(Palette::parse_tolerant(&[0], 0, &mut ToleranceReport::default()).is_err());
    }

    #[test]
    fn a_short_palette_keeps_the_default_remainder() {
        let mut data = palette_payload();
        data.truncate(2 + 16 * 4);
        data[..2].copy_from_slice(&16i16.to_le_bytes());
        data[2..5].copy_from_slice(&[1, 2, 3]);
        let mut report = ToleranceReport::default();
        let palette = Palette::parse_tolerant(&data, 40, &mut report).unwrap();
        assert_eq!(palette.color(0x08), Some(Color::new(1, 2, 3)));
        assert_eq!(palette.color(0x3f), Some(Color::new(51, 51, 51)));
        assert_eq!(report.count(PaletteDefect::ColorCount), 1);
        assert_eq!(report.defects()[0].ordinal(), 40);
        assert_eq!(report.defects()[0].observed(), 16);
        // The strict parser still refuses it.
        assert!(Palette::parse_record(&data).is_err());
    }

    #[test]
    fn a_long_palette_uses_the_first_fifty_six_entries() {
        let mut data = palette_payload();
        data[..2].copy_from_slice(&57i16.to_le_bytes());
        data.extend_from_slice(&[9, 9, 9, 0]);
        let mut report = ToleranceReport::default();
        let palette = Palette::parse_tolerant(&data, 0, &mut report).unwrap();
        assert_eq!(palette.palette_colors(), &DEFAULT_PALETTE);
        assert_eq!(report.count(PaletteDefect::ColorCount), 1);
    }

    #[test]
    fn a_nonzero_long_rgb_reserved_byte_is_ignored_and_recorded() {
        let mut data = palette_payload();
        data[5] = 1;
        data[9] = 0xff;
        let mut report = ToleranceReport::default();
        let palette = Palette::parse_tolerant(&data, 7, &mut report).unwrap();
        assert_eq!(palette.palette_colors(), &DEFAULT_PALETTE);
        assert_eq!(report.total(), 1);
        assert_eq!(report.count(PaletteDefect::ReservedByte), 1);
        assert_eq!(report.defects()[0].observed(), 2, "two entries carried one");
        assert!(Palette::parse_record(&data).is_err());
    }

    #[test]
    fn a_repeated_palette_record_wins_and_is_recorded() {
        let first = palette_payload();
        let mut second = palette_payload();
        second[2..5].copy_from_slice(&[7, 8, 9]);
        let mut seen = 0;
        let mut report = ToleranceReport::default();
        Palette::parse_unique_record(&first, 10, &mut seen, &mut report).unwrap();
        assert!(report.is_clean());
        let palette = Palette::parse_unique_record(&second, 300, &mut seen, &mut report).unwrap();
        assert_eq!(palette.color(0x08), Some(Color::new(7, 8, 9)));
        assert_eq!(report.count(PaletteDefect::Duplicate), 1);
        assert_eq!(report.defects()[0].ordinal(), 300);
        assert_eq!(report.defects()[0].observed(), 2);
    }
}

//! BIFF8 workbook color palette support.

use super::error::{XlsError, XlsResult};

const PALETTE_COLOR_COUNT: usize = 56;
const FIRST_PALETTE_INDEX: u16 = 0x0008;
const LAST_PALETTE_INDEX: u16 = 0x003f;
const PALETTE_RECORD_LENGTH: usize = 2 + PALETTE_COLOR_COUNT * 4;

/// An RGB color from a BIFF8 workbook color table.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsColor {
    red: u8,
    green: u8,
    blue: u8,
}

impl XlsColor {
    const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    /// Red component.
    pub fn red(self) -> u8 {
        self.red
    }

    /// Green component.
    pub fn green(self) -> u8 {
        self.green
    }

    /// Blue component.
    pub fn blue(self) -> u8 {
        self.blue
    }
}

const BUILT_IN_COLORS: [XlsColor; 8] = [
    XlsColor::new(0x00, 0x00, 0x00),
    XlsColor::new(0xff, 0xff, 0xff),
    XlsColor::new(0xff, 0x00, 0x00),
    XlsColor::new(0x00, 0xff, 0x00),
    XlsColor::new(0x00, 0x00, 0xff),
    XlsColor::new(0xff, 0xff, 0x00),
    XlsColor::new(0xff, 0x00, 0xff),
    XlsColor::new(0x00, 0xff, 0xff),
];

const DEFAULT_PALETTE: [XlsColor; PALETTE_COLOR_COUNT] = [
    XlsColor::new(0x00, 0x00, 0x00),
    XlsColor::new(0xff, 0xff, 0xff),
    XlsColor::new(0xff, 0x00, 0x00),
    XlsColor::new(0x00, 0xff, 0x00),
    XlsColor::new(0x00, 0x00, 0xff),
    XlsColor::new(0xff, 0xff, 0x00),
    XlsColor::new(0xff, 0x00, 0xff),
    XlsColor::new(0x00, 0xff, 0xff),
    XlsColor::new(0x80, 0x00, 0x00),
    XlsColor::new(0x00, 0x80, 0x00),
    XlsColor::new(0x00, 0x00, 0x80),
    XlsColor::new(0x80, 0x80, 0x00),
    XlsColor::new(0x80, 0x00, 0x80),
    XlsColor::new(0x00, 0x80, 0x80),
    XlsColor::new(0xc0, 0xc0, 0xc0),
    XlsColor::new(0x80, 0x80, 0x80),
    XlsColor::new(0x99, 0x99, 0xff),
    XlsColor::new(0x99, 0x33, 0x66),
    XlsColor::new(0xff, 0xff, 0xcc),
    XlsColor::new(0xcc, 0xff, 0xff),
    XlsColor::new(0x66, 0x00, 0x66),
    XlsColor::new(0xff, 0x80, 0x80),
    XlsColor::new(0x00, 0x66, 0xcc),
    XlsColor::new(0xcc, 0xcc, 0xff),
    XlsColor::new(0x00, 0x00, 0x80),
    XlsColor::new(0xff, 0x00, 0xff),
    XlsColor::new(0xff, 0xff, 0x00),
    XlsColor::new(0x00, 0xff, 0xff),
    XlsColor::new(0x80, 0x00, 0x80),
    XlsColor::new(0x80, 0x00, 0x00),
    XlsColor::new(0x00, 0x80, 0x80),
    XlsColor::new(0x00, 0x00, 0xff),
    XlsColor::new(0x00, 0xcc, 0xff),
    XlsColor::new(0xcc, 0xff, 0xff),
    XlsColor::new(0xcc, 0xff, 0xcc),
    XlsColor::new(0xff, 0xff, 0x99),
    XlsColor::new(0x99, 0xcc, 0xff),
    XlsColor::new(0xff, 0x99, 0xcc),
    XlsColor::new(0xcc, 0x99, 0xff),
    XlsColor::new(0xff, 0xcc, 0x99),
    XlsColor::new(0x33, 0x66, 0xff),
    XlsColor::new(0x33, 0xcc, 0xcc),
    XlsColor::new(0x99, 0xcc, 0x00),
    XlsColor::new(0xff, 0xcc, 0x00),
    XlsColor::new(0xff, 0x99, 0x00),
    XlsColor::new(0xff, 0x66, 0x00),
    XlsColor::new(0x66, 0x66, 0x99),
    XlsColor::new(0x96, 0x96, 0x96),
    XlsColor::new(0x00, 0x33, 0x66),
    XlsColor::new(0x33, 0x99, 0x66),
    XlsColor::new(0x00, 0x33, 0x00),
    XlsColor::new(0x33, 0x33, 0x00),
    XlsColor::new(0x99, 0x33, 0x00),
    XlsColor::new(0x99, 0x33, 0x66),
    XlsColor::new(0x33, 0x33, 0x99),
    XlsColor::new(0x33, 0x33, 0x33),
];

/// The 56-color BIFF8 workbook palette.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsPalette {
    colors: [XlsColor; PALETTE_COLOR_COUNT],
    custom: bool,
}

impl Default for XlsPalette {
    fn default() -> Self {
        Self {
            colors: DEFAULT_PALETTE,
            custom: false,
        }
    }
}

impl XlsPalette {
    /// Resolve a built-in or workbook palette color-table index.
    pub fn color(&self, index: u16) -> Option<XlsColor> {
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
    pub fn palette_colors(&self) -> &[XlsColor; PALETTE_COLOR_COUNT] {
        &self.colors
    }

    /// Whether the workbook contained a custom `Palette` record.
    pub fn is_custom(&self) -> bool {
        self.custom
    }

    pub(crate) fn parse_record(data: &[u8]) -> XlsResult<Self> {
        if data.len() != PALETTE_RECORD_LENGTH {
            return Err(XlsError::InvalidData(format!(
                "Palette record has {} bytes; expected {PALETTE_RECORD_LENGTH}",
                data.len()
            )));
        }
        let color_count = i16::from_le_bytes([data[0], data[1]]);
        if color_count != PALETTE_COLOR_COUNT as i16 {
            return Err(XlsError::InvalidData(format!(
                "Palette color count is {color_count}; expected {PALETTE_COLOR_COUNT}"
            )));
        }

        let mut colors = [XlsColor::new(0, 0, 0); PALETTE_COLOR_COUNT];
        for (index, color) in colors.iter_mut().enumerate() {
            let offset = 2 + index * 4;
            let reserved = data[offset + 3];
            if reserved != 0 {
                return Err(XlsError::InvalidData(format!(
                    "Palette color {index} reserved byte is {reserved:#04x}; expected zero"
                )));
            }
            *color = XlsColor::new(data[offset], data[offset + 1], data[offset + 2]);
        }
        Ok(Self {
            colors,
            custom: true,
        })
    }

    pub(crate) fn parse_unique_record(data: &[u8], seen: &mut bool) -> XlsResult<Self> {
        if *seen {
            return Err(XlsError::InvalidData(
                "workbook globals contain duplicate Palette records".to_string(),
            ));
        }
        let palette = Self::parse_record(data)?;
        *seen = true;
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

        let palette = XlsPalette::parse_record(&data).unwrap();
        assert!(palette.is_custom());
        assert_eq!(palette.color(0x12), Some(XlsColor::new(101, 230, 100)));
        assert_eq!(palette.color(0x3b), Some(XlsColor::new(0, 255, 52)));
        assert_eq!(palette.color(0x02), Some(XlsColor::new(255, 0, 0)));
        assert_eq!(palette.color(0x40), None);
    }

    #[test]
    fn supplies_normative_default_palette() {
        let palette = XlsPalette::default();
        assert!(!palette.is_custom());
        assert_eq!(palette.color(0x08), Some(XlsColor::new(0, 0, 0)));
        assert_eq!(palette.color(0x0d), Some(XlsColor::new(255, 255, 0)));
        assert_eq!(palette.color(0x3f), Some(XlsColor::new(51, 51, 51)));
    }

    #[test]
    fn rejects_wrong_count_and_record_length() {
        let mut wrong_count = palette_payload();
        wrong_count[..2].copy_from_slice(&55i16.to_le_bytes());
        assert!(XlsPalette::parse_record(&wrong_count).is_err());

        let mut truncated = palette_payload();
        truncated.pop();
        assert!(XlsPalette::parse_record(&truncated).is_err());

        let mut extra = palette_payload();
        extra.push(0);
        assert!(XlsPalette::parse_record(&extra).is_err());
    }

    #[test]
    fn rejects_nonzero_long_rgb_reserved_byte() {
        let mut data = palette_payload();
        data[5] = 1;
        assert!(XlsPalette::parse_record(&data).is_err());
    }

    #[test]
    fn rejects_duplicate_palette_records() {
        let data = palette_payload();
        let mut seen = false;
        XlsPalette::parse_unique_record(&data, &mut seen).unwrap();
        assert!(XlsPalette::parse_unique_record(&data, &mut seen).is_err());
    }
}

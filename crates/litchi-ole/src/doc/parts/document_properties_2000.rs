//! Typed, lossless Word 2000 Document Properties extension.

use super::document_properties_97::{Dop97, DopExtensionError};

const DOP97_SIZE: usize = 500;
const DOP2000_SIZE: usize = 544;
const EXTENSION_SIZE: usize = DOP2000_SIZE - DOP97_SIZE;
const DOP97_COPTS80_OFFSET: usize = 84;
const FLAGS1: usize = 4;
const COPTS: usize = 8;
const COPTS_SIZE: usize = 32;
const VERSION_COMPATIBILITY: usize = 40;
const FLAGS2: usize = 42;

/// A lossless 256-bit Word compatibility option set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CompatibilityOptions {
    bytes: [u8; COPTS_SIZE],
}

impl CompatibilityOptions {
    pub const fn from_bytes(bytes: [u8; COPTS_SIZE]) -> Self {
        Self { bytes }
    }

    pub const fn as_bytes(&self) -> &[u8; COPTS_SIZE] {
        &self.bytes
    }

    pub fn is_set(&self, bit: u8) -> bool {
        let bit = usize::from(bit);
        self.bytes[bit / 8] & (1 << (bit % 8)) != 0
    }

    pub fn set(&mut self, bit: u8, enabled: bool) {
        let bit = usize::from(bit);
        let mask = 1 << (bit % 8);
        if enabled {
            self.bytes[bit / 8] |= mask;
        } else {
            self.bytes[bit / 8] &= !mask;
        }
    }
}

/// Target browser screen size for Word's Web export options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebScreenSize {
    Pixels544x376,
    Pixels640x480,
    Pixels720x512,
    Pixels800x600,
    Pixels1024x768,
    Pixels1152x882,
    Pixels1152x900,
    Pixels1280x1024,
    Pixels1600x1200,
    Pixels1800x1440,
    Pixels1920x1200,
}

impl WebScreenSize {
    fn parse(value: u8) -> Result<Self, DopExtensionError> {
        match value {
            0 => Ok(Self::Pixels544x376),
            1 => Ok(Self::Pixels640x480),
            2 => Ok(Self::Pixels720x512),
            3 => Ok(Self::Pixels800x600),
            4 => Ok(Self::Pixels1024x768),
            5 => Ok(Self::Pixels1152x882),
            6 => Ok(Self::Pixels1152x900),
            7 => Ok(Self::Pixels1280x1024),
            8 => Ok(Self::Pixels1600x1200),
            9 => Ok(Self::Pixels1800x1440),
            10 => Ok(Self::Pixels1920x1200),
            _ => Err(DopExtensionError::new(format!(
                "invalid Word 2000 Web screen-size value {value}"
            ))),
        }
    }

    const fn raw(self) -> u32 {
        match self {
            Self::Pixels544x376 => 0,
            Self::Pixels640x480 => 1,
            Self::Pixels720x512 => 2,
            Self::Pixels800x600 => 3,
            Self::Pixels1024x768 => 4,
            Self::Pixels1152x882 => 5,
            Self::Pixels1152x900 => 6,
            Self::Pixels1280x1024 => 7,
            Self::Pixels1600x1200 => 8,
            Self::Pixels1800x1440 => 9,
            Self::Pixels1920x1200 => 10,
        }
    }
}

/// Valid Word Web-export settings. Their absence means the stored bits are ignored.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WebExportOptions {
    pub rely_on_css: bool,
    pub rely_on_vml: bool,
    pub allow_png: bool,
    pub screen_size: WebScreenSize,
    pub organize_in_folder: bool,
    pub use_long_file_names: bool,
    pub pixels_per_inch: u16,
}

impl WebExportOptions {
    fn parse(flags: u32) -> Result<Self, DopExtensionError> {
        let pixels_per_inch = ((flags >> 18) & 0x3ff) as u16;
        if !(19..=480).contains(&pixels_per_inch) {
            return Err(DopExtensionError::new(format!(
                "Word 2000 Web pixels-per-inch {pixels_per_inch} is outside 19..=480"
            )));
        }
        Ok(Self {
            rely_on_css: flags & (1 << 9) != 0,
            rely_on_vml: flags & (1 << 10) != 0,
            allow_png: flags & (1 << 11) != 0,
            screen_size: WebScreenSize::parse(((flags >> 12) & 0xf) as u8)?,
            organize_in_folder: flags & (1 << 16) != 0,
            use_long_file_names: flags & (1 << 17) != 0,
            pixels_per_inch,
        })
    }
}

/// Desired legacy Word feature set. Unknown bits are retained and ignored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct LegacyFeatureSet(u16);

impl LegacyFeatureSet {
    pub const WORD_95: u16 = 0x0004;
    pub const WORD_97: u16 = 0x0008;
    pub const EAST_ASIAN_WORD_95: u16 = 0x0040;
    pub const WORD_2003: u16 = 0x0800;

    pub const fn from_raw(raw: u16) -> Self {
        Self(raw)
    }

    pub const fn raw(self) -> u16 {
        self.0
    }

    pub const fn contains(self, feature: u16) -> bool {
        self.0 & feature != 0
    }
}

/// Typed, lossless Word 2000 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dop2000 {
    raw: [u8; EXTENSION_SIZE],
    pub last_bullet_level: u8,
    pub last_number_level: u8,
    pub click_and_type_style: u16,
    pub language_detection_complete: bool,
    pub show_email_envelope: bool,
    pub maybe_tentative_lists: bool,
    pub maybe_fit_text: bool,
    pub format_consistency_complete: bool,
    pub web_options: Option<WebExportOptions>,
    pub maybe_east_asian_layout: bool,
    pub character_and_line_units: bool,
    pub compatibility: CompatibilityOptions,
    pub legacy_feature_set: LegacyFeatureSet,
    pub suppress_page_boundaries: bool,
    pub repaired_document: bool,
    pub save_uim: bool,
    pub filter_privacy_on_save: bool,
    pub repairs_seen: bool,
    pub has_custom_xml: bool,
    pub validate_custom_xml: bool,
    pub save_invalid_custom_xml: bool,
    pub show_custom_xml_errors: bool,
    pub merge_empty_xml_namespace: bool,
}

impl Dop2000 {
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2000_SIZE {
            return Err(DopExtensionError::new("Dop2000 is shorter than 544 bytes"));
        }
        Dop97::parse(dop)?;
        let extension = &dop[DOP97_SIZE..DOP2000_SIZE];
        let mut raw = [0u8; EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        if extension[0] > 9 || extension[1] > 9 {
            return Err(DopExtensionError::new(
                "Word 2000 toolbar list level exceeds 9",
            ));
        }
        let flags1 = le_u32(extension, FLAGS1);
        if flags1 & 0xf0 != 0 {
            return Err(DopExtensionError::new("Dop2000 empty1 bits are nonzero"));
        }
        let flags2 = le_u16(extension, FLAGS2);
        if flags2 & ((1 << 5) | (1 << 8)) != 0 {
            return Err(DopExtensionError::new(
                "Dop2000 empty2/empty3 bits are nonzero",
            ));
        }
        let mut copts_bytes = [0u8; COPTS_SIZE];
        copts_bytes.copy_from_slice(&extension[COPTS..COPTS + COPTS_SIZE]);
        if copts_bytes[..4] != dop[DOP97_COPTS80_OFFSET..DOP97_COPTS80_OFFSET + 4] {
            return Err(DopExtensionError::new(
                "Dop2000 Copts.copts80 does not mirror Dop97.copts80",
            ));
        }
        Ok(Self {
            raw,
            last_bullet_level: extension[0],
            last_number_level: extension[1],
            click_and_type_style: le_u16(extension, 2),
            language_detection_complete: flags1 & 1 != 0,
            show_email_envelope: flags1 & 2 != 0,
            maybe_tentative_lists: flags1 & 4 != 0,
            maybe_fit_text: flags1 & 8 != 0,
            format_consistency_complete: flags1 & (1 << 8) != 0,
            web_options: if flags1 & (1 << 28) != 0 {
                Some(WebExportOptions::parse(flags1)?)
            } else {
                None
            },
            maybe_east_asian_layout: flags1 & (1 << 29) != 0,
            character_and_line_units: flags1 & (1 << 30) != 0,
            compatibility: CompatibilityOptions::from_bytes(copts_bytes),
            legacy_feature_set: LegacyFeatureSet::from_raw(le_u16(
                extension,
                VERSION_COMPATIBILITY,
            )),
            suppress_page_boundaries: flags2 & 1 != 0,
            repaired_document: flags2 & (1 << 4) != 0,
            save_uim: flags2 & (1 << 6) != 0,
            filter_privacy_on_save: flags2 & (1 << 7) != 0,
            repairs_seen: flags2 & (1 << 9) != 0,
            has_custom_xml: flags2 & (1 << 10) != 0,
            validate_custom_xml: flags2 & (1 << 12) != 0,
            save_invalid_custom_xml: flags2 & (1 << 13) != 0,
            show_custom_xml_errors: flags2 & (1 << 14) != 0,
            merge_empty_xml_namespace: flags2 & (1 << 15) != 0,
        })
    }

    /// Validates the style index when the stylesheet slot count is available.
    pub fn validate_style_index(&self, style_count: usize) -> Result<(), DopExtensionError> {
        if usize::from(self.click_and_type_style) >= style_count {
            Err(DopExtensionError::new(format!(
                "Dop2000 click-and-type style {} exceeds stylesheet",
                self.click_and_type_style
            )))
        } else {
            Ok(())
        }
    }

    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2000_SIZE {
            return Err(DopExtensionError::new(
                "Dop2000 target is shorter than 544 bytes",
            ));
        }
        if self.last_bullet_level > 9 || self.last_number_level > 9 {
            return Err(DopExtensionError::new(
                "Word 2000 toolbar list level exceeds 9",
            ));
        }
        if self.compatibility.as_bytes()[..4] != dop[DOP97_COPTS80_OFFSET..DOP97_COPTS80_OFFSET + 4]
        {
            return Err(DopExtensionError::new(
                "Dop2000 Copts.copts80 does not mirror Dop97.copts80",
            ));
        }
        self.raw[0] = self.last_bullet_level;
        self.raw[1] = self.last_number_level;
        put_u16(&mut self.raw, 2, self.click_and_type_style);
        let old_flags1 = le_u32(&self.raw, FLAGS1);
        let mut flags1 = old_flags1 & 0x8000_0000;
        flags1 |= u32::from(self.language_detection_complete);
        flags1 |= u32::from(self.show_email_envelope) << 1;
        flags1 |= u32::from(self.maybe_tentative_lists) << 2;
        flags1 |= u32::from(self.maybe_fit_text) << 3;
        flags1 |= u32::from(self.format_consistency_complete) << 8;
        if let Some(web) = self.web_options {
            if !(19..=480).contains(&web.pixels_per_inch) {
                return Err(DopExtensionError::new(
                    "Word 2000 Web pixels-per-inch is outside 19..=480",
                ));
            }
            flags1 |= u32::from(web.rely_on_css) << 9;
            flags1 |= u32::from(web.rely_on_vml) << 10;
            flags1 |= u32::from(web.allow_png) << 11;
            flags1 |= web.screen_size.raw() << 12;
            flags1 |= u32::from(web.organize_in_folder) << 16;
            flags1 |= u32::from(web.use_long_file_names) << 17;
            flags1 |= u32::from(web.pixels_per_inch) << 18;
            flags1 |= 1 << 28;
        } else {
            flags1 |= old_flags1 & 0x0fff_fe00;
        }
        flags1 |= u32::from(self.maybe_east_asian_layout) << 29;
        flags1 |= u32::from(self.character_and_line_units) << 30;
        put_u32(&mut self.raw, FLAGS1, flags1);
        self.raw[COPTS..COPTS + COPTS_SIZE].copy_from_slice(self.compatibility.as_bytes());
        put_u16(
            &mut self.raw,
            VERSION_COMPATIBILITY,
            self.legacy_feature_set.raw(),
        );
        let old_flags2 = le_u16(&self.raw, FLAGS2);
        let mut flags2 = old_flags2 & 0x080e;
        flags2 |= u16::from(self.suppress_page_boundaries);
        flags2 |= u16::from(self.repaired_document) << 4;
        flags2 |= u16::from(self.save_uim) << 6;
        flags2 |= u16::from(self.filter_privacy_on_save) << 7;
        flags2 |= u16::from(self.repairs_seen) << 9;
        flags2 |= u16::from(self.has_custom_xml) << 10;
        flags2 |= u16::from(self.validate_custom_xml) << 12;
        flags2 |= u16::from(self.save_invalid_custom_xml) << 13;
        flags2 |= u16::from(self.show_custom_xml_errors) << 14;
        flags2 |= u16::from(self.merge_empty_xml_namespace) << 15;
        put_u16(&mut self.raw, FLAGS2, flags2);
        dop[DOP97_SIZE..DOP2000_SIZE].copy_from_slice(&self.raw);
        Ok(())
    }
}

fn le_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        data[offset..offset + 4]
            .try_into()
            .expect("fixed-width slice"),
    )
}

fn put_u16(data: &mut [u8], offset: usize, value: u16) {
    data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_dop2000() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2000_SIZE];
        let grid = [0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81];
        dop[0x190..0x19a].copy_from_slice(&grid);
        dop
    }

    #[test]
    fn parses_typed_web_xml_and_version_state() {
        let mut dop = valid_dop2000();
        let flags1 = 1 | (1 << 9) | (3 << 12) | (1 << 16) | (1 << 17) | (96 << 18) | (1 << 28);
        put_u32(&mut dop, DOP97_SIZE + FLAGS1, flags1);
        put_u16(
            &mut dop,
            DOP97_SIZE + VERSION_COMPATIBILITY,
            LegacyFeatureSet::WORD_97,
        );
        put_u16(
            &mut dop,
            DOP97_SIZE + FLAGS2,
            (1 << 7) | (1 << 10) | (1 << 12),
        );
        let value = Dop2000::parse(&dop).unwrap();
        assert!(value.language_detection_complete);
        assert_eq!(value.web_options.unwrap().pixels_per_inch, 96);
        assert!(value.legacy_feature_set.contains(LegacyFeatureSet::WORD_97));
        assert!(value.filter_privacy_on_save);
        assert!(value.has_custom_xml);
        assert!(value.validate_custom_xml);
    }

    #[test]
    fn round_trip_preserves_undefined_bits_and_bytes() {
        let mut dop = valid_dop2000();
        dop[DOP97_SIZE + FLAGS1 + 3] = 0x80;
        dop[DOP97_SIZE + FLAGS2] = 0x0e;
        dop[DOP97_SIZE + COPTS + 31] = 0xa5;
        let parsed = Dop2000::parse(&dop).unwrap();
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(output, dop);
    }

    #[test]
    fn rejects_must_constraint_violations() {
        let mut level = valid_dop2000();
        level[DOP97_SIZE] = 10;
        assert!(Dop2000::parse(&level).is_err());

        let mut empty = valid_dop2000();
        empty[DOP97_SIZE + FLAGS1] = 0x10;
        assert!(Dop2000::parse(&empty).is_err());

        let mut mirror = valid_dop2000();
        mirror[DOP97_SIZE + COPTS] = 1;
        assert!(Dop2000::parse(&mirror).is_err());

        let mut web = valid_dop2000();
        put_u32(&mut web, DOP97_SIZE + FLAGS1, (1 << 28) | (18 << 18));
        assert!(Dop2000::parse(&web).is_err());
    }

    #[test]
    fn validates_context_dependent_style_index() {
        let mut dop = valid_dop2000();
        put_u16(&mut dop, DOP97_SIZE + 2, 4);
        let parsed = Dop2000::parse(&dop).unwrap();
        assert!(parsed.validate_style_index(4).is_err());
        assert!(parsed.validate_style_index(5).is_ok());
    }
}

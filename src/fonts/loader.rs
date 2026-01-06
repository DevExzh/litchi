use font_kit::family_name::FamilyName;
use font_kit::handle::Handle;
use font_kit::properties::Properties;
use font_kit::source::SystemSource;

use crate::fonts::{FontData, FontError, FontProperties};

pub struct FontLoader {
    source: SystemSource,
}

impl FontLoader {
    pub fn new() -> Self {
        Self {
            source: SystemSource::new(),
        }
    }

    pub fn load_system_font(&self, family_name: &str) -> Result<FontData, FontError> {
        let handle = self
            .source
            .select_best_match(
                &[FamilyName::Title(family_name.to_string())],
                &Properties::new(),
            )
            .map_err(|_| FontError::NotFound(family_name.to_string()))?;

        match handle {
            Handle::Path { path, font_index } => {
                let data = std::fs::read(&path)?;
                let properties = Self::extract_font_properties(&data);
                Ok(FontData {
                    name: family_name.to_string(),
                    data,
                    index: font_index,
                    properties,
                })
            },
            Handle::Memory { bytes, font_index } => {
                let properties = Self::extract_font_properties(&bytes);
                Ok(FontData {
                    name: family_name.to_string(),
                    data: bytes.to_vec(),
                    index: font_index,
                    properties,
                })
            },
        }
    }

    /// Extract font properties from font data for Office embedding
    fn extract_font_properties(data: &[u8]) -> Option<FontProperties> {
        use allsorts::binary::read::ReadScope;
        use allsorts::tables::{FontTableProvider, OpenTypeFont};

        let scope = ReadScope::new(data);
        let font_file = scope.read::<OpenTypeFont<'_>>().ok()?;
        let provider = font_file.table_provider(0).ok()?;

        // Get OS/2 table raw bytes
        let os2_table = provider.table_data(allsorts::tag::OS_2).ok()??;

        // Need at least 70 bytes for Unicode ranges
        if os2_table.len() < 70 {
            return None;
        }

        // PANOSE is at offset 32 (0x20) in OS/2 table (10 bytes)
        let panose_bytes = &os2_table[32..42];
        let panose = format!(
            "{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
            panose_bytes[0],
            panose_bytes[1],
            panose_bytes[2],
            panose_bytes[3],
            panose_bytes[4],
            panose_bytes[5],
            panose_bytes[6],
            panose_bytes[7],
            panose_bytes[8],
            panose_bytes[9]
        );

        // Unicode ranges at offset 42-57 (4 DWORDs)
        let ul_unicode_range1 =
            u32::from_be_bytes([os2_table[42], os2_table[43], os2_table[44], os2_table[45]]);
        let ul_unicode_range2 =
            u32::from_be_bytes([os2_table[46], os2_table[47], os2_table[48], os2_table[49]]);
        let ul_unicode_range3 =
            u32::from_be_bytes([os2_table[50], os2_table[51], os2_table[52], os2_table[53]]);
        let ul_unicode_range4 =
            u32::from_be_bytes([os2_table[54], os2_table[55], os2_table[56], os2_table[57]]);

        // Code page ranges at offset 78-85 (2 DWORDs) - if version >= 1
        let (ul_code_page_range1, ul_code_page_range2) = if os2_table.len() >= 86 {
            let cp1 =
                u32::from_be_bytes([os2_table[78], os2_table[79], os2_table[80], os2_table[81]]);
            let cp2 =
                u32::from_be_bytes([os2_table[82], os2_table[83], os2_table[84], os2_table[85]]);
            (cp1, cp2)
        } else {
            (0, 0)
        };

        let sig = Some((
            format!("{:08X}", ul_unicode_range1),
            format!("{:08X}", ul_unicode_range2),
            format!("{:08X}", ul_unicode_range3),
            format!("{:08X}", ul_unicode_range4),
            format!("{:08X}", ul_code_page_range1),
            format!("{:08X}", ul_code_page_range2),
        ));

        // Charset - derive from code page ranges (ulCodePageRange1/2)
        // Format as decimal (Office uses decimal, not hex)
        // Common charsets:
        //   0 = ANSI/Western (Latin-1, CP1252)
        //   128 = Shift JIS (Japanese)
        //   134 = Chinese Simplified (GB2312) - appears as -122 when signed
        //   136 = Chinese Traditional (Big5)
        //   etc.
        let charset = if ul_code_page_range1 & 0x0000_0001 != 0 {
            // Bit 0: Latin 1 (CP1252) - use ANSI charset (0)
            "0".to_string()
        } else if ul_code_page_range1 & 0x0002_0000 != 0 {
            // Bit 17: Japanese Shift-JIS (CP932)
            "128".to_string()
        } else if ul_code_page_range1 & 0x0008_0000 != 0 {
            // Bit 19: Chinese Simplified (CP936/GBK)
            "134".to_string()
        } else if ul_code_page_range1 & 0x0010_0000 != 0 {
            // Bit 20: Korean Wansung (CP949)
            "129".to_string()
        } else if ul_code_page_range1 & 0x0020_0000 != 0 {
            // Bit 21: Chinese Traditional (CP950/Big5)
            "136".to_string()
        } else {
            // Default to ANSI if no specific code page is detected
            "0".to_string()
        };

        // Determine family type based on panose family kind (first byte)
        let family = match panose_bytes[0] {
            2 => "roman",  // Text and Display
            3 => "swiss",  // Script
            4 => "modern", // Decorative
            5 => "script", // Pictorial
            _ => "auto",
        };

        // Pitch - check if monospace (panose proportion field, byte 3)
        let pitch = if panose_bytes[3] == 9 {
            "fixed"
        } else {
            "variable"
        };

        Some(FontProperties {
            panose: Some(panose),
            charset: Some(charset),
            family: Some(family.to_string()),
            pitch: Some(pitch.to_string()),
            sig,
        })
    }
}

impl Default for FontLoader {
    fn default() -> Self {
        Self::new()
    }
}

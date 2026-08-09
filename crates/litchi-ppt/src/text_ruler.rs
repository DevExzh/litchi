//! `PowerPoint` `TextRulerAtom` parsing.

use litchi_core::binary::{read_i16_le, read_u16_le, read_u32_le};

use super::package::{Error, Result};
use super::records::Record;
use super::text_prop::TextTabStop;
use crate::consts::RecordType;

/// Margin and first-line indent overrides for one paragraph level.
#[allow(
    clippy::module_name_repetitions,
    reason = "`TextRulerLevel` is the established public API name; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct TextRulerLevel {
    /// Left margin in `PowerPoint` master units.
    pub left_margin: Option<i16>,
    /// First-line indent in `PowerPoint` master units.
    pub indent: Option<i16>,
}

/// Tabbing, margins, and indentation from a `TextRulerAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextRuler {
    /// Original presence mask, including ignored reserved bits.
    pub mask: u32,
    /// Number of style levels, when explicitly present.
    pub level_count: Option<i16>,
    /// Default tab size in `PowerPoint` master units.
    pub default_tab_size: Option<i16>,
    /// Whether the tab-stop array was explicitly present.
    pub tab_stops_present: bool,
    /// Explicit tab stops.
    pub tab_stops: Vec<TextTabStop>,
    /// Per-level ruler overrides for indent levels `0..=4`.
    pub levels: [TextRulerLevel; 5],
}

impl TextRuler {
    /// Parse the payload of a `TextRulerAtom`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(data: &[u8]) -> Result<Self> {
        require_bytes(data, 0, 4, "TextRuler mask")?;
        let mask = read_u32_le(data, 0).unwrap_or(0);
        if mask & !0x0000_1fff != 0 {
            return Err(Error::Corrupted(
                "TextRuler has nonzero reserved mask bits".to_string(),
            ));
        }
        let mut offset = 4usize;

        let level_count = read_optional_i16(data, &mut offset, mask & 0x0002 != 0, "cLevels")?;
        if level_count.is_some_and(|count| !(0..=5).contains(&count)) {
            return Err(Error::Corrupted(
                "TextRuler cLevels must be between 0 and 5".to_string(),
            ));
        }
        let default_tab_size =
            read_optional_i16(data, &mut offset, mask & 0x0001 != 0, "defaultTabSize")?;

        let tab_stops_present = mask & 0x0004 != 0;
        let mut tab_stops = Vec::new();
        if tab_stops_present {
            require_bytes(data, offset, 2, "TextRuler tab count")?;
            let count = read_u16_le(data, offset).unwrap_or(0) as usize;
            offset += 2;
            let byte_count = count
                .checked_mul(4)
                .ok_or_else(|| Error::Corrupted("TextRuler tab size overflow".to_string()))?;
            require_bytes(data, offset, byte_count, "TextRuler tab stops")?;
            tab_stops.reserve(count);
            for _ in 0..count {
                let position = read_i16_le(data, offset).unwrap_or(0);
                let alignment = read_u16_le(data, offset + 2).unwrap_or(0);
                if alignment > 3 {
                    return Err(Error::Corrupted(
                        "TextRuler has an invalid TextTabTypeEnum value".to_string(),
                    ));
                }
                tab_stops.push(TextTabStop {
                    position,
                    alignment,
                });
                offset += 4;
            }
        }

        let mut levels = [TextRulerLevel::default(); 5];
        for (level, value) in levels.iter_mut().enumerate() {
            value.left_margin = read_optional_i16(
                data,
                &mut offset,
                mask & (0x0008u32 << level) != 0,
                "left margin",
            )?;
            value.indent = read_optional_i16(
                data,
                &mut offset,
                mask & (0x0100u32 << level) != 0,
                "indent",
            )?;
        }

        if offset != data.len() {
            return Err(Error::Corrupted(
                "TextRulerAtom has trailing bytes".to_string(),
            ));
        }
        Ok(Self {
            mask,
            level_count,
            default_tab_size,
            tab_stops_present,
            tab_stops,
            levels,
        })
    }

    /// Parse a `DefaultRulerAtom` payload and require all default fields.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_default(data: &[u8]) -> Result<Self> {
        let ruler = Self::parse(data)?;
        if ruler.mask & 0x0000_1fff != 0x0000_1fff {
            return Err(Error::Corrupted(
                "DefaultRulerAtom omits required ruler fields".to_string(),
            ));
        }
        Ok(ruler)
    }
}

/// Discover and parse the document-wide `DefaultRulerAtom` below `root`.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
#[allow(
    clippy::module_name_repetitions,
    reason = "`parse_default_text_ruler` is the established public entry point of this module; renaming it would break downstream crates"
)]
pub fn parse_default_text_ruler(root: &Record) -> Result<Option<TextRuler>> {
    let mut records = Vec::new();
    collect_default_rulers(root, &mut records);
    if records.len() > 1 {
        return Err(Error::Corrupted(
            "Record tree contains multiple DefaultRulerAtom records".to_string(),
        ));
    }
    records
        .first()
        .map(|record| {
            if record.version != 0 || record.instance != 0 {
                return Err(Error::Corrupted(
                    "DefaultRulerAtom has an invalid record header".to_string(),
                ));
            }
            TextRuler::parse_default(&record.data)
        })
        .transpose()
}

fn collect_default_rulers<'a>(record: &'a Record, output: &mut Vec<&'a Record>) {
    if record.record_type == RecordType::DefaultRulerAtom {
        output.push(record);
        return;
    }
    for child in &record.children {
        collect_default_rulers(child, output);
    }
}

fn read_optional_i16(
    data: &[u8],
    offset: &mut usize,
    present: bool,
    field: &str,
) -> Result<Option<i16>> {
    if !present {
        return Ok(None);
    }
    require_bytes(data, *offset, 2, field)?;
    let value = read_i16_le(data, *offset).unwrap_or(0);
    *offset += 2;
    Ok(Some(value))
}

fn require_bytes(data: &[u8], offset: usize, size: usize, field: &str) -> Result<()> {
    let end = offset
        .checked_add(size)
        .ok_or_else(|| Error::Corrupted(format!("{field} offset overflow")))?;
    if end > data.len() {
        return Err(Error::Corrupted(format!("Truncated {field}")));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn parses_complete_text_ruler_in_spec_order() {
        let mask: u32 = 0x0001 | 0x0002 | 0x0004 | 0x0008 | 0x0020 | 0x0100 | 0x0400;
        let mut data = Vec::new();
        data.extend_from_slice(&mask.to_le_bytes());
        data.extend_from_slice(&5i16.to_le_bytes());
        data.extend_from_slice(&144i16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&(-20i16).to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&720i16.to_le_bytes());
        data.extend_from_slice(&3u16.to_le_bytes());
        data.extend_from_slice(&100i16.to_le_bytes());
        data.extend_from_slice(&(-50i16).to_le_bytes());
        data.extend_from_slice(&300i16.to_le_bytes());
        data.extend_from_slice(&(-150i16).to_le_bytes());

        let ruler = TextRuler::parse(&data).unwrap();

        assert_eq!(ruler.level_count, Some(5));
        assert_eq!(ruler.default_tab_size, Some(144));
        assert!(ruler.tab_stops_present);
        assert_eq!(ruler.tab_stops.len(), 2);
        assert_eq!(ruler.tab_stops[0].position, -20);
        assert_eq!(ruler.tab_stops[0].alignment, 1);
        assert_eq!(ruler.levels[0].left_margin, Some(100));
        assert_eq!(ruler.levels[0].indent, Some(-50));
        assert_eq!(ruler.levels[2].left_margin, Some(300));
        assert_eq!(ruler.levels[2].indent, Some(-150));
    }

    #[test]
    fn rejects_invalid_or_truncated_text_rulers() {
        assert!(TextRuler::parse(&[]).is_err());

        let mut invalid_levels = Vec::new();
        invalid_levels.extend_from_slice(&0x0002u32.to_le_bytes());
        invalid_levels.extend_from_slice(&6i16.to_le_bytes());
        assert!(TextRuler::parse(&invalid_levels).is_err());

        let mut invalid_tab = Vec::new();
        invalid_tab.extend_from_slice(&0x0004u32.to_le_bytes());
        invalid_tab.extend_from_slice(&1u16.to_le_bytes());
        invalid_tab.extend_from_slice(&0i16.to_le_bytes());
        invalid_tab.extend_from_slice(&4u16.to_le_bytes());
        assert!(TextRuler::parse(&invalid_tab).is_err());

        let mut trailing = 0u32.to_le_bytes().to_vec();
        trailing.push(0);
        assert!(TextRuler::parse(&trailing).is_err());

        assert!(TextRuler::parse(&0x2000u32.to_le_bytes()).is_err());
        assert!(TextRuler::parse_default(&0u32.to_le_bytes()).is_err());
    }

    #[test]
    fn parses_complete_default_ruler() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x1fffu32.to_le_bytes());
        data.extend_from_slice(&5i16.to_le_bytes());
        data.extend_from_slice(&720i16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for level in 0..5i16 {
            data.extend_from_slice(&(level * 100).to_le_bytes());
            data.extend_from_slice(&(level * 100 - 25).to_le_bytes());
        }

        let ruler = TextRuler::parse_default(&data).unwrap();
        assert_eq!(ruler.level_count, Some(5));
        assert_eq!(ruler.default_tab_size, Some(720));
        assert_eq!(ruler.levels[4].left_margin, Some(400));
        assert_eq!(ruler.levels[4].indent, Some(375));

        let atom = Record {
            record_type: RecordType::DefaultRulerAtom,
            record_type_raw: 4011,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
            data,
            children: Vec::new(),
        };
        let root = Record {
            record_type: RecordType::Environment,
            record_type_raw: 1010,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![atom],
        };
        assert_eq!(
            parse_default_text_ruler(&root)
                .unwrap()
                .unwrap()
                .default_tab_size,
            Some(720)
        );
    }
}

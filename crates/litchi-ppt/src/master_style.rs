//! `PowerPoint` `TxMasterStyleAtom` parsing.

use litchi_core::binary::{read_u16_le, read_u32_le};

use super::package::{Error, Result};
use super::text_prop::{
    TextPropCollection, TextPropType, character_property_size, paragraph_property_size,
    parse_character_properties, parse_paragraph_properties, require_style_bytes,
};

/// Formatting defaults for one master text indent level.
#[derive(Debug, Clone)]
pub struct TextMasterStyleLevel {
    /// Explicit level stored for text types 5 through 8.
    pub explicit_level: Option<u16>,
    /// Paragraph formatting defaults.
    pub paragraph: TextPropCollection,
    /// Character formatting defaults.
    pub character: TextPropCollection,
}

/// Parsed text defaults from a `TxMasterStyleAtom`.
#[allow(
    clippy::module_name_repetitions,
    reason = "`TextMasterStyle` is the established public API name for the `TxMasterStyleAtom` \
              model; renaming it would break downstream crates"
)]
#[derive(Debug, Clone)]
pub struct TextMasterStyle {
    /// `TextTypeEnum` value from the record instance.
    pub text_type: u16,
    /// Master formatting levels.
    pub levels: Vec<TextMasterStyleLevel>,
}

impl TextMasterStyle {
    /// Parse a `TxMasterStyleAtom` payload for the supplied record instance.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(data: &[u8], text_type: u16) -> Result<Self> {
        if !matches!(text_type, 0 | 1 | 2 | 4 | 5 | 6 | 7 | 8) {
            return Err(Error::Corrupted(
                "TxMasterStyleAtom has an invalid TextTypeEnum instance".to_string(),
            ));
        }
        require_style_bytes(data, 0, 2, "TxMasterStyleAtom level count")?;
        let level_count = read_u16_le(data, 0).unwrap_or(0);
        if level_count > 5 {
            return Err(Error::Corrupted(
                "TxMasterStyleAtom has more than five levels".to_string(),
            ));
        }

        let mut offset = 2usize;
        let mut levels = Vec::with_capacity(level_count as usize);
        let mut seen_levels = [false; 5];
        for logical_level in 0..level_count {
            let explicit_level = if text_type >= 5 {
                require_style_bytes(data, offset, 2, "TextMasterStyleLevel level")?;
                let level = read_u16_le(data, offset).unwrap_or(0);
                offset += 2;
                if level >= level_count || seen_levels[level as usize] {
                    return Err(Error::Corrupted(
                        "TextMasterStyleLevel has an invalid or duplicate level".to_string(),
                    ));
                }
                seen_levels[level as usize] = true;
                Some(level)
            } else {
                None
            };

            require_style_bytes(data, offset, 4, "TextMasterStyleLevel paragraph mask")?;
            let paragraph_mask = read_u32_le(data, offset).unwrap_or(0);
            offset += 4;
            let paragraph_size = paragraph_property_size(data, offset, paragraph_mask)?;
            require_style_bytes(
                data,
                offset,
                paragraph_size,
                "TextMasterStyleLevel paragraph",
            )?;
            let paragraph_end = offset + paragraph_size;
            let (paragraph_properties, tab_stops) =
                parse_paragraph_properties(data, &mut offset, paragraph_mask);
            if offset != paragraph_end {
                return Err(Error::Corrupted(
                    "TextMasterStyleLevel paragraph size mismatch".to_string(),
                ));
            }
            let effective_level = explicit_level.unwrap_or(logical_level);
            let mut paragraph = TextPropCollection::new(0, TextPropType::Paragraph);
            paragraph.indent_level = effective_level;
            paragraph.properties = paragraph_properties;
            paragraph.property_mask = paragraph_mask;
            paragraph.tab_stops = tab_stops;

            require_style_bytes(data, offset, 4, "TextMasterStyleLevel character mask")?;
            let character_mask = read_u32_le(data, offset).unwrap_or(0);
            offset += 4;
            let character_size = character_property_size(character_mask);
            require_style_bytes(
                data,
                offset,
                character_size,
                "TextMasterStyleLevel character",
            )?;
            let character_end = offset + character_size;
            let character_properties =
                parse_character_properties(data, &mut offset, character_mask);
            if offset != character_end {
                return Err(Error::Corrupted(
                    "TextMasterStyleLevel character size mismatch".to_string(),
                ));
            }
            let mut character = TextPropCollection::new(0, TextPropType::Character);
            character.properties = character_properties;
            character.property_mask = character_mask;

            levels.push(TextMasterStyleLevel {
                explicit_level,
                paragraph,
                character,
            });
        }

        if offset != data.len() {
            return Err(Error::Corrupted(
                "TxMasterStyleAtom has trailing bytes".to_string(),
            ));
        }
        Ok(Self { text_type, levels })
    }
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
    fn parses_full_and_explicit_master_style_levels() {
        let mut body = Vec::new();
        body.extend_from_slice(&1u16.to_le_bytes());
        body.extend_from_slice(&0x0800u32.to_le_bytes());
        body.extend_from_slice(&2u16.to_le_bytes());
        body.extend_from_slice(&0x0003_0001u32.to_le_bytes());
        body.extend_from_slice(&1u16.to_le_bytes());
        body.extend_from_slice(&65_535u16.to_le_bytes());
        body.extend_from_slice(&24i16.to_le_bytes());
        let style = TextMasterStyle::parse(&body, 1).unwrap();
        assert_eq!(style.levels.len(), 1);
        assert_eq!(style.levels[0].explicit_level, None);
        assert_eq!(style.levels[0].paragraph.get_value("alignment"), Some(2));
        assert_eq!(style.levels[0].character.get_value("char.flags"), Some(1));
        assert_eq!(
            style.levels[0].character.get_value("font.index"),
            Some(65_535)
        );
        assert_eq!(style.levels[0].character.get_value("font.size"), Some(24));

        let mut centered = Vec::new();
        centered.extend_from_slice(&1u16.to_le_bytes());
        centered.extend_from_slice(&0u16.to_le_bytes());
        centered.extend_from_slice(&0u32.to_le_bytes());
        centered.extend_from_slice(&0u32.to_le_bytes());
        let centered_style = TextMasterStyle::parse(&centered, 5).unwrap();
        assert_eq!(centered_style.levels[0].explicit_level, Some(0));
    }

    #[test]
    fn rejects_invalid_master_style_framing() {
        assert!(TextMasterStyle::parse(&[], 1).is_err());
        assert!(TextMasterStyle::parse(&0u16.to_le_bytes(), 3).is_err());
        assert!(TextMasterStyle::parse(&6u16.to_le_bytes(), 1).is_err());

        let mut duplicate = Vec::new();
        duplicate.extend_from_slice(&2u16.to_le_bytes());
        for _ in 0..2 {
            duplicate.extend_from_slice(&0u16.to_le_bytes());
            duplicate.extend_from_slice(&0u32.to_le_bytes());
            duplicate.extend_from_slice(&0u32.to_le_bytes());
        }
        assert!(TextMasterStyle::parse(&duplicate, 5).is_err());
    }
}

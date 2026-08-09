/// Text property parsing for `PowerPoint` `StyleTextPropAtom`.
///
/// Based on Apache POI's `TextPropCollection` and `TextProp` classes.
/// This module handles the complex structure of text styling in PPT files.
use litchi_core::binary::{read_i16_le, read_i32_le, read_u16_le, read_u32_le};

use super::package::{Error, Result};

/// Text property definition.
///
/// Based on Apache POI's `TextProp`. Each property has a size, mask, and value.
#[derive(Debug, Clone)]
pub struct TextProp {
    /// Name of the property
    pub name: &'static str,
    /// Size in bytes (0, 2, or 4)
    pub size: usize,
    /// Mask in the header field
    pub mask: u32,
    /// Value of the property
    pub value: i32,
}

impl TextProp {
    /// Create a new text property.
    #[must_use]
    pub fn new(name: &'static str, size: usize, mask: u32) -> Self {
        Self {
            name,
            size,
            mask,
            value: 0,
        }
    }
}

/// Text property collection type.
#[allow(
    clippy::module_name_repetitions,
    reason = "`TextPropType` is the established public API name; renaming it would break downstream crates"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextPropType {
    /// Paragraph properties
    Paragraph,
    /// Character properties
    Character,
}

/// A raw tab stop from a `PowerPoint` paragraph property run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextTabStop {
    /// Signed offset in `PowerPoint` master units.
    pub position: i16,
    /// Raw `TextTabTypeEnum` value.
    pub alignment: u16,
}

/// Collection of text properties for a run of characters.
///
/// Based on Apache POI's `TextPropCollection`.
#[allow(
    clippy::module_name_repetitions,
    reason = "`TextPropCollection` is the established public API name; renaming it would break downstream crates"
)]
#[derive(Debug, Clone)]
pub struct TextPropCollection {
    /// Number of characters this styling applies to
    pub characters_covered: u32,
    /// Unsigned paragraph indent level
    pub indent_level: u16,
    /// The properties in this collection
    pub properties: Vec<TextProp>,
    /// Original property-presence mask for this collection
    pub property_mask: u32,
    /// Tab stops carried by a paragraph property run
    pub tab_stops: Vec<TextTabStop>,
    /// Type of collection
    pub prop_type: TextPropType,
}

impl TextPropCollection {
    /// Create a new text property collection.
    #[must_use]
    pub fn new(characters_covered: u32, prop_type: TextPropType) -> Self {
        Self {
            characters_covered,
            indent_level: 0,
            properties: Vec::new(),
            property_mask: 0,
            tab_stops: Vec::new(),
            prop_type,
        }
    }

    /// Find a property by name.
    #[must_use]
    pub fn find_by_name(&self, name: &str) -> Option<&TextProp> {
        self.properties.iter().find(|p| p.name == name)
    }

    /// Get a property value by name.
    #[must_use]
    pub fn get_value(&self, name: &str) -> Option<i32> {
        self.find_by_name(name).map(|p| p.value)
    }
}

/// Parse paragraph text properties from binary data.
///
/// Based on POI's paragraph text property types.
pub fn parse_paragraph_properties(
    data: &[u8],
    offset: &mut usize,
    mask: u32,
) -> (Vec<TextProp>, Vec<TextTabStop>) {
    let mut props = Vec::new();
    let mut tab_stops = Vec::new();

    // The fields occur in TextPFException order, not numeric mask order.
    let prop_defs = [
        ("paragraph.flags", 2, 0x000F, false),
        ("bullet.char", 2, 0x0080, false),
        ("bullet.font", 2, 0x0010, false),
        ("bullet.size", 2, 0x0040, true),
        ("bullet.color", 4, 0x0020, false),
        ("alignment", 2, 0x0800, false),
        ("linespacing", 2, 0x1000, true),
        ("spacebefore", 2, 0x2000, true),
        ("spaceafter", 2, 0x4000, true),
        ("text.offset", 2, 0x0100, true),   // left margin
        ("bullet.offset", 2, 0x0400, true), // indent
        ("defaultTabSize", 2, 0x8000, true),
    ];

    for (name, size, prop_mask, signed) in prop_defs {
        if (mask & prop_mask) != 0 {
            if *offset + size > data.len() {
                *offset = data.len();
                return (props, tab_stops);
            }

            let value = match size {
                2 if signed => i32::from(read_i16_le(data, *offset).unwrap_or(0)),
                2 => i32::from(read_u16_le(data, *offset).unwrap_or(0)),
                4 => read_i32_le(data, *offset).unwrap_or(0),
                _ => 0,
            };

            let mut prop = TextProp::new(name, size, prop_mask);
            prop.value = value;
            props.push(prop);
            *offset += size;
        }
    }

    if (mask & 0x10_0000) != 0 {
        if *offset + 2 > data.len() {
            *offset = data.len();
            return (props, tab_stops);
        }
        let tab_stop_count = read_u16_le(data, *offset).unwrap_or(0);
        let count = usize::from(tab_stop_count);
        let Some(size) = count.checked_mul(4).and_then(|size| size.checked_add(2)) else {
            *offset = data.len();
            return (props, tab_stops);
        };
        if *offset + size > data.len() {
            *offset = data.len();
            return (props, tab_stops);
        }
        let mut prop = TextProp::new("tabStops", size, 0x10_0000);
        prop.value = i32::from(tab_stop_count);
        props.push(prop);
        let mut tab_offset = *offset + 2;
        for _ in 0..count {
            tab_stops.push(TextTabStop {
                position: read_i16_le(data, tab_offset).unwrap_or(0),
                alignment: read_u16_le(data, tab_offset + 2).unwrap_or(0),
            });
            tab_offset += 4;
        }
        *offset += size;
    }

    let trailing_defs = [
        ("fontAlignment", 2, 0x10000),
        ("wrapFlags", 2, 0xE0000),
        ("textDirection", 2, 0x20_0000),
    ];

    for (name, size, prop_mask) in trailing_defs {
        if (mask & prop_mask) != 0 {
            if *offset + size > data.len() {
                *offset = data.len();
                return (props, tab_stops);
            }

            let value = match size {
                2 => i32::from(read_u16_le(data, *offset).unwrap_or(0)),
                4 => read_i32_le(data, *offset).unwrap_or(0),
                _ => 0,
            };

            let mut prop = TextProp::new(name, size, prop_mask);
            prop.value = value;
            props.push(prop);
            *offset += size;
        }
    }

    (props, tab_stops)
}

/// Parse character text properties from binary data.
///
/// Based on POI's character text property types.
pub fn parse_character_properties(data: &[u8], offset: &mut usize, mask: u32) -> Vec<TextProp> {
    let mut props = Vec::new();

    // Character property definitions (from POI's TextPropCollection)
    let prop_defs = [
        ("char.flags", 2, 0xFFFF, false), // bold, italic, underline, etc.
        ("font.index", 2, 0x10000, false),
        ("asian.font.index", 2, 0x20_0000, false),
        ("ansi.font.index", 2, 0x40_0000, false),
        ("symbol.font.index", 2, 0x80_0000, false),
        ("font.size", 2, 0x20000, true),
        ("font.color", 4, 0x40000, false),
        ("superscript", 2, 0x80000, true),
    ];

    for (name, size, prop_mask, signed) in prop_defs {
        if (mask & prop_mask) != 0 {
            if *offset + size > data.len() {
                *offset = data.len();
                return props;
            }

            let value = match size {
                2 if signed => i32::from(read_i16_le(data, *offset).unwrap_or(0)),
                2 => i32::from(read_u16_le(data, *offset).unwrap_or(0)),
                4 => read_i32_le(data, *offset).unwrap_or(0),
                _ => 0,
            };

            let mut prop = TextProp::new(name, size, prop_mask);
            prop.value = value;
            props.push(prop);
            *offset += size;
        }
    }

    props
}

/// Parse `StyleTextPropAtom` data.
///
/// Based on Apache POI's `StyleTextPropAtom` parsing logic.
/// Returns (`paragraph_styles`, `character_styles`).
#[must_use]
pub fn parse_style_text_prop_atom(
    data: &[u8],
    text_length: usize,
) -> (Vec<TextPropCollection>, Vec<TextPropCollection>) {
    let mut paragraph_styles = Vec::new();
    let mut character_styles = Vec::new();

    if data.len() < 8 {
        return (paragraph_styles, character_styles);
    }

    let mut offset = 0;
    let style_length = u32::try_from(text_length)
        .unwrap_or(u32::MAX)
        .saturating_add(1);

    // Parse paragraph styles first
    let mut para_chars_covered = 0u32;
    while para_chars_covered < style_length && offset + 10 <= data.len() {
        // Read character count (4 bytes in POI's implementation)
        let char_count = read_u32_le(data, offset)
            .unwrap_or(0)
            .min(style_length - para_chars_covered);
        offset += 4;

        if char_count == 0 {
            break;
        }

        // Read indent level (2 bytes)
        let indent_level = read_u16_le(data, offset).unwrap_or(0);
        offset += 2;

        // Read mask (4 bytes)
        if offset + 4 > data.len() {
            break;
        }
        let mask = read_u32_le(data, offset).unwrap_or(0);
        offset += 4;

        // Parse properties based on mask
        let (properties, tab_stops) = parse_paragraph_properties(data, &mut offset, mask);

        let mut collection = TextPropCollection::new(char_count, TextPropType::Paragraph);
        collection.indent_level = indent_level;
        collection.properties = properties;
        collection.property_mask = mask;
        collection.tab_stops = tab_stops;
        paragraph_styles.push(collection);

        para_chars_covered += char_count;
    }

    // Parse character styles
    let mut char_chars_covered = 0u32;
    while char_chars_covered < style_length && offset + 8 <= data.len() {
        // Read character count (4 bytes)
        let char_count = read_u32_le(data, offset)
            .unwrap_or(0)
            .min(style_length - char_chars_covered);
        offset += 4;

        if char_count == 0 {
            break;
        }

        // Read mask (4 bytes) - no indent level for character styles
        if offset + 4 > data.len() {
            break;
        }
        let mask = read_u32_le(data, offset).unwrap_or(0);
        offset += 4;

        // Parse properties based on mask
        let properties = parse_character_properties(data, &mut offset, mask);

        let mut collection = TextPropCollection::new(char_count, TextPropType::Character);
        collection.properties = properties;
        collection.property_mask = mask;
        character_styles.push(collection);

        char_chars_covered += char_count;
    }

    (paragraph_styles, character_styles)
}

/// Parse `StyleTextPropAtom` data with strict MS-PPT framing validation.
///
/// Unlike [`parse_style_text_prop_atom`], this rejects zero-length runs,
/// truncated property payloads, coverage beyond or below `text_length + 1`,
/// and unexplained trailing bytes.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_style_text_prop_atom_strict(
    data: &[u8],
    text_length: usize,
) -> Result<(Vec<TextPropCollection>, Vec<TextPropCollection>)> {
    let style_length = u32::try_from(text_length)
        .map_err(|_err| Error::Corrupted("StyleTextPropAtom text length exceeds u32".to_string()))?
        .checked_add(1)
        .ok_or_else(|| Error::Corrupted("StyleTextPropAtom text length overflow".to_string()))?;
    let mut paragraph_styles = Vec::new();
    let mut character_styles = Vec::new();
    let mut offset = 0usize;
    let mut paragraph_coverage = 0u32;

    while paragraph_coverage < style_length {
        require_style_bytes(data, offset, 10, "TextPFRun header")?;
        let char_count = read_u32_le(data, offset).unwrap_or(0);
        if char_count == 0 || char_count > style_length - paragraph_coverage {
            return Err(Error::Corrupted(
                "TextPFRun has invalid character coverage".to_string(),
            ));
        }
        let indent_level = read_u16_le(data, offset + 4).unwrap_or(0);
        let mask = read_u32_le(data, offset + 6).unwrap_or(0);
        offset += 10;

        let property_size = paragraph_property_size(data, offset, mask)?;
        require_style_bytes(data, offset, property_size, "TextPFException")?;
        let property_end = offset + property_size;
        let (properties, tab_stops) = parse_paragraph_properties(data, &mut offset, mask);
        if offset != property_end {
            return Err(Error::Corrupted(
                "TextPFException property size mismatch".to_string(),
            ));
        }

        let mut collection = TextPropCollection::new(char_count, TextPropType::Paragraph);
        collection.indent_level = indent_level;
        collection.properties = properties;
        collection.property_mask = mask;
        collection.tab_stops = tab_stops;
        paragraph_styles.push(collection);
        paragraph_coverage += char_count;
    }

    let mut character_coverage = 0u32;
    while character_coverage < style_length {
        require_style_bytes(data, offset, 8, "TextCFRun header")?;
        let char_count = read_u32_le(data, offset).unwrap_or(0);
        if char_count == 0 || char_count > style_length - character_coverage {
            return Err(Error::Corrupted(
                "TextCFRun has invalid character coverage".to_string(),
            ));
        }
        let mask = read_u32_le(data, offset + 4).unwrap_or(0);
        offset += 8;

        let property_size = character_property_size(mask);
        require_style_bytes(data, offset, property_size, "TextCFException")?;
        let property_end = offset + property_size;
        let properties = parse_character_properties(data, &mut offset, mask);
        if offset != property_end {
            return Err(Error::Corrupted(
                "TextCFException property size mismatch".to_string(),
            ));
        }

        let mut collection = TextPropCollection::new(char_count, TextPropType::Character);
        collection.properties = properties;
        collection.property_mask = mask;
        character_styles.push(collection);
        character_coverage += char_count;
    }

    if offset != data.len() {
        return Err(Error::Corrupted(
            "StyleTextPropAtom has trailing bytes".to_string(),
        ));
    }
    Ok((paragraph_styles, character_styles))
}

pub(crate) fn paragraph_property_size(data: &[u8], offset: usize, mask: u32) -> Result<usize> {
    let mut size = 0usize;
    if mask & 0x000F != 0 {
        size += 2;
    }
    for (property_mask, property_size) in [
        (0x0080, 2usize),
        (0x0010, 2),
        (0x0040, 2),
        (0x0020, 4),
        (0x0800, 2),
        (0x1000, 2),
        (0x2000, 2),
        (0x4000, 2),
        (0x0100, 2),
        (0x0400, 2),
        (0x8000, 2),
    ] {
        if mask & property_mask != 0 {
            size += property_size;
        }
    }
    if mask & 0x0010_0000 != 0 {
        require_style_bytes(data, offset, size + 2, "TabStops count")?;
        let count = read_u16_le(data, offset + size).unwrap_or(0) as usize;
        size = size
            .checked_add(2)
            .and_then(|total| {
                count
                    .checked_mul(4)
                    .and_then(|tabs| total.checked_add(tabs))
            })
            .ok_or_else(|| Error::Corrupted("TabStops size overflow".to_string()))?;
    }
    if mask & 0x0001_0000 != 0 {
        size += 2;
    }
    if mask & 0x000E_0000 != 0 {
        size += 2;
    }
    if mask & 0x0020_0000 != 0 {
        size += 2;
    }
    Ok(size)
}

pub(crate) fn character_property_size(mask: u32) -> usize {
    let mut size = usize::from(mask & 0x0000_FFFF != 0) * 2;
    for (property_mask, property_size) in [
        (0x0001_0000, 2usize),
        (0x0020_0000, 2),
        (0x0040_0000, 2),
        (0x0080_0000, 2),
        (0x0002_0000, 2),
        (0x0004_0000, 4),
        (0x0008_0000, 2),
    ] {
        if mask & property_mask != 0 {
            size += property_size;
        }
    }
    size
}

pub(crate) fn require_style_bytes(
    data: &[u8],
    offset: usize,
    size: usize,
    field: &str,
) -> Result<()> {
    let end = offset
        .checked_add(size)
        .ok_or_else(|| Error::Corrupted(format!("{field} offset overflow")))?;
    if end > data.len() {
        return Err(Error::Corrupted(format!("Truncated {field}")));
    }
    Ok(())
}

/// Extract formatting from character flags.
///
/// Character flags (mask 0x0001) contains packed boolean properties:
/// - Bit 0: Bold
/// - Bit 1: Italic
/// - Bit 2: Underline
/// - Bit 4: Shadow
/// - Bit 8: Embossed
#[must_use]
pub fn extract_char_flags(flags: i32) -> (bool, bool, bool) {
    let bold = (flags & 0x0001) != 0;
    let italic = (flags & 0x0002) != 0;
    let underline = (flags & 0x0004) != 0;
    (bold, italic, underline)
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
    fn test_text_prop_creation() {
        let prop = TextProp::new("font.size", 2, 0x20000);
        assert_eq!(prop.name, "font.size");
        assert_eq!(prop.size, 2);
        assert_eq!(prop.mask, 0x20000);
    }

    #[test]
    fn test_text_prop_collection() {
        let collection = TextPropCollection::new(10, TextPropType::Character);
        assert_eq!(collection.characters_covered, 10);
        assert_eq!(collection.prop_type, TextPropType::Character);
    }

    #[test]
    fn test_extract_char_flags() {
        let (bold, italic, underline) = extract_char_flags(0x0007);
        assert!(bold);
        assert!(italic);
        assert!(underline);

        let (bold_flag, italic_flag, underline_flag) = extract_char_flags(0x0001);
        assert!(bold_flag);
        assert!(!italic_flag);
        assert!(!underline_flag);
    }

    #[test]
    fn parses_paragraph_fields_in_record_order_before_character_runs() {
        let mut data = Vec::new();
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(&0i16.to_le_bytes());
        data.extend_from_slice(&0x0010_08FFu32.to_le_bytes());
        data.extend_from_slice(&1i16.to_le_bytes());
        data.extend_from_slice(&0x2022i16.to_le_bytes());
        data.extend_from_slice(&2i16.to_le_bytes());
        data.extend_from_slice(&4i16.to_le_bytes());
        data.extend_from_slice(&0x0011_2233i32.to_le_bytes());
        data.extend_from_slice(&2i16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&720i16.to_le_bytes());
        data.extend_from_slice(&0i16.to_le_bytes());
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(&0x20002u32.to_le_bytes());
        data.extend_from_slice(&2i16.to_le_bytes());
        data.extend_from_slice(&20i16.to_le_bytes());

        let (paragraphs, characters) = parse_style_text_prop_atom(&data, 4);

        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].get_value("bullet.char"), Some(0x2022));
        assert_eq!(paragraphs[0].get_value("alignment"), Some(2));
        assert_eq!(paragraphs[0].get_value("tabStops"), Some(1));
        assert_eq!(paragraphs[0].property_mask, 0x0010_08FF);
        assert_eq!(
            paragraphs[0].tab_stops,
            vec![TextTabStop {
                position: 720,
                alignment: 0,
            }]
        );
        assert_eq!(characters.len(), 1);
        assert_eq!(characters[0].get_value("char.flags"), Some(2));
        assert_eq!(characters[0].get_value("font.size"), Some(20));
    }

    #[test]
    fn parses_font_references_as_unsigned_values() {
        let mut offset = 0;
        let properties = parse_character_properties(&[0xFF, 0xFF], &mut offset, 0x10000);

        assert_eq!(properties[0].value, 65_535);
    }

    fn minimal_style_data() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data
    }

    #[test]
    fn strict_style_parser_requires_exact_run_coverage() {
        let mut zero_count = minimal_style_data();
        zero_count[..4].copy_from_slice(&0u32.to_le_bytes());
        assert!(parse_style_text_prop_atom_strict(&zero_count, 1).is_err());

        let mut overlong = minimal_style_data();
        overlong[..4].copy_from_slice(&3u32.to_le_bytes());
        assert!(parse_style_text_prop_atom_strict(&overlong, 1).is_err());

        let mut underlong = minimal_style_data();
        underlong[..4].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_style_text_prop_atom_strict(&underlong, 1).is_err());
    }

    #[test]
    fn strict_style_parser_rejects_truncation_and_trailing_bytes() {
        let valid = minimal_style_data();
        assert!(parse_style_text_prop_atom_strict(&valid, 1).is_ok());

        for length in 0..valid.len() {
            assert!(parse_style_text_prop_atom_strict(&valid[..length], 1).is_err());
        }

        let mut trailing = valid;
        trailing.push(0);
        let error = parse_style_text_prop_atom_strict(&trailing, 1).unwrap_err();
        assert!(error.to_string().contains("trailing bytes"));
    }

    #[test]
    fn strict_style_parser_bounds_checks_tab_stop_arrays() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0x0010_0000u32.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&10i16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());

        let error = parse_style_text_prop_atom_strict(&data, 1).unwrap_err();
        assert!(error.to_string().contains("TextPFException"));
    }
}

/// Text property parsing for PowerPoint StyleTextPropAtom.
///
/// Based on Apache POI's TextPropCollection and TextProp classes.
/// This module handles the complex structure of text styling in PPT files.
use crate::ole::binary::read_u16_le;

/// Text property definition.
///
/// Based on Apache POI's TextProp. Each property has a size, mask, and value.
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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextPropType {
    /// Paragraph properties
    Paragraph,
    /// Character properties
    Character,
}

/// Collection of text properties for a run of characters.
///
/// Based on Apache POI's TextPropCollection.
#[derive(Debug, Clone)]
pub struct TextPropCollection {
    /// Number of characters this styling applies to
    pub characters_covered: u32,
    /// Indent level (for paragraphs, -1 if not set)
    pub indent_level: i16,
    /// The properties in this collection
    pub properties: Vec<TextProp>,
    /// Type of collection
    pub prop_type: TextPropType,
}

impl TextPropCollection {
    /// Create a new text property collection.
    pub fn new(characters_covered: u32, prop_type: TextPropType) -> Self {
        Self {
            characters_covered,
            indent_level: -1,
            properties: Vec::new(),
            prop_type,
        }
    }

    /// Find a property by name.
    pub fn find_by_name(&self, name: &str) -> Option<&TextProp> {
        self.properties.iter().find(|p| p.name == name)
    }

    /// Get a property value by name.
    pub fn get_value(&self, name: &str) -> Option<i32> {
        self.find_by_name(name).map(|p| p.value)
    }
}

/// Parse paragraph text properties from binary data.
///
/// Based on POI's paragraph text property types.
pub fn parse_paragraph_properties(data: &[u8], offset: &mut usize, mask: u32) -> Vec<TextProp> {
    let mut props = Vec::new();

    // Paragraph property definitions (from POI's TextPropCollection)
    let prop_defs = [
        ("alignment", 2, 0x0008),
        ("linespacing", 2, 0x1000),
        ("spacebefore", 2, 0x2000),
        ("spaceafter", 2, 0x4000),
        ("text.offset", 2, 0x0100),    // left margin
        ("bullet.offset", 2, 0x0400),   // indent
        ("defaultTabSize", 2, 0x8000),
        ("textDirection", 2, 0x200000),
    ];

    for (name, size, prop_mask) in &prop_defs {
        if (mask & prop_mask) != 0 {
            if *offset + size > data.len() {
                break;
            }

            let value = match size {
                2 => read_u16_le(data, *offset).unwrap_or(0) as i32,
                4 => {
                    if *offset + 4 <= data.len() {
                        i32::from_le_bytes([data[*offset], data[*offset + 1], data[*offset + 2], data[*offset + 3]])
                    } else {
                        0
                    }
                }
                _ => 0,
            };

            let mut prop = TextProp::new(name, *size, *prop_mask);
            prop.value = value;
            props.push(prop);

            *offset += size;
        }
    }

    props
}

/// Parse character text properties from binary data.
///
/// Based on POI's character text property types.
pub fn parse_character_properties(data: &[u8], offset: &mut usize, mask: u32) -> Vec<TextProp> {
    let mut props = Vec::new();

    // Character property definitions (from POI's TextPropCollection)
    let prop_defs = [
        ("char.flags", 2, 0x0001),        // bold, italic, underline, etc.
        ("font.index", 2, 0x10000),
        ("asian.font.index", 2, 0x200000),
        ("ansi.font.index", 2, 0x400000),
        ("symbol.font.index", 2, 0x800000),
        ("font.size", 2, 0x20000),
        ("font.color", 4, 0x40000),
        ("superscript", 2, 0x80000),
    ];

    for (name, size, prop_mask) in &prop_defs {
        if (mask & prop_mask) != 0 {
            if *offset + size > data.len() {
                break;
            }

            let value = match size {
                2 => read_u16_le(data, *offset).unwrap_or(0) as i32,
                4 => {
                    if *offset + 4 <= data.len() {
                        i32::from_le_bytes([data[*offset], data[*offset + 1], data[*offset + 2], data[*offset + 3]])
                    } else {
                        0
                    }
                }
                _ => 0,
            };

            let mut prop = TextProp::new(name, *size, *prop_mask);
            prop.value = value;
            props.push(prop);

            *offset += size;
        }
    }

    props
}

/// Parse StyleTextPropAtom data.
///
/// Based on Apache POI's StyleTextPropAtom parsing logic.
/// Returns (paragraph_styles, character_styles).
pub fn parse_style_text_prop_atom(data: &[u8], text_length: usize) -> (Vec<TextPropCollection>, Vec<TextPropCollection>) {
    let mut paragraph_styles = Vec::new();
    let mut character_styles = Vec::new();

    if data.len() < 10 {
        return (paragraph_styles, character_styles);
    }

    let mut offset = 0;

    // Parse paragraph styles first
    let mut para_chars_covered = 0u32;
    while para_chars_covered < text_length as u32 && offset + 6 <= data.len() {
        // Read character count (4 bytes in POI's implementation)
        let char_count = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]);
        offset += 4;

        if char_count == 0 {
            break;
        }

        // Read indent level (2 bytes)
        let indent_level = i16::from_le_bytes([data[offset], data[offset + 1]]);
        offset += 2;

        // Read mask (4 bytes)
        if offset + 4 > data.len() {
            break;
        }
        let mask = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]);
        offset += 4;

        // Parse properties based on mask
        let properties = parse_paragraph_properties(data, &mut offset, mask);

        let mut collection = TextPropCollection::new(char_count, TextPropType::Paragraph);
        collection.indent_level = indent_level;
        collection.properties = properties;
        paragraph_styles.push(collection);

        para_chars_covered += char_count;
    }

    // Parse character styles
    let mut char_chars_covered = 0u32;
    while char_chars_covered < text_length as u32 && offset + 6 <= data.len() {
        // Read character count (4 bytes)
        let char_count = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]);
        offset += 4;

        if char_count == 0 {
            break;
        }

        // Read mask (4 bytes) - no indent level for character styles
        if offset + 4 > data.len() {
            break;
        }
        let mask = u32::from_le_bytes([data[offset], data[offset + 1], data[offset + 2], data[offset + 3]]);
        offset += 4;

        // Parse properties based on mask
        let properties = parse_character_properties(data, &mut offset, mask);

        let mut collection = TextPropCollection::new(char_count, TextPropType::Character);
        collection.properties = properties;
        character_styles.push(collection);

        char_chars_covered += char_count;
    }

    (paragraph_styles, character_styles)
}

/// Extract formatting from character flags.
///
/// Character flags (mask 0x0001) contains packed boolean properties:
/// - Bit 0: Bold
/// - Bit 1: Italic
/// - Bit 2: Underline
/// - Bit 4: Shadow
/// - Bit 8: Embossed
pub fn extract_char_flags(flags: i32) -> (bool, bool, bool) {
    let bold = (flags & 0x0001) != 0;
    let italic = (flags & 0x0002) != 0;
    let underline = (flags & 0x0004) != 0;
    (bold, italic, underline)
}

#[cfg(test)]
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

        let (bold, italic, underline) = extract_char_flags(0x0001);
        assert!(bold);
        assert!(!italic);
        assert!(!underline);
    }
}


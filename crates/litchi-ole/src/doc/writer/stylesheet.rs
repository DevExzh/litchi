//! Validated Word 97+ stylesheet (`STSH`) generation.

use crate::doc::parts::styles::{StyleFlags, StyleKind, StylePost2000, StyleSheet};
use crate::sprm::parse_sprms;
use crate::sprm_operations::{SPRM_C_CNF, SPRM_P_CNF, get_sprm_type};

const MIN_STYLE_COUNT: usize = 15;
const MAX_STYLE_COUNT: usize = 0x0FFD;
const NIL_STYLE: u16 = 0x0FFF;
const USER_STYLE_ID: u16 = 0x0FFE;

/// A custom style appended after the fifteen fixed DOC style slots.
#[derive(Debug, Clone)]
pub struct DocStyleDefinition {
    /// Invariant built-in identifier, or `0x0FFE` for a user-defined style.
    pub invariant_id: u16,
    /// Paragraph, character, table, or numbering style.
    pub kind: StyleKind,
    /// Optional parent style index.
    pub base_style: Option<u16>,
    /// Style used for the paragraph following this style.
    pub next_style: u16,
    /// Primary style name.
    pub name: String,
    /// Alternate comma-free names.
    pub aliases: Vec<String>,
    /// Kind-specific UPX payloads in TAP/PAP/CHP order prescribed by MS-DOC.
    pub property_sets: Vec<Vec<u8>>,
    /// Optional Word 2000-and-later style metadata.
    pub post_2000: Option<StylePost2000>,
    /// Style behavior flags.
    pub flags: StyleFlags,
}

impl DocStyleDefinition {
    /// Create an empty user-defined style with the required UPX count.
    pub fn new(kind: StyleKind, name: impl Into<String>) -> Self {
        let property_count = match kind {
            StyleKind::Paragraph => 2,
            StyleKind::Character | StyleKind::Numbering => 1,
            StyleKind::Table => 3,
        };
        Self {
            invariant_id: USER_STYLE_ID,
            kind,
            base_style: None,
            next_style: 0,
            name: name.into(),
            aliases: Vec::new(),
            property_sets: vec![Vec::new(); property_count],
            post_2000: None,
            flags: StyleFlags::default(),
        }
    }

    /// Set the parent style index.
    pub fn with_base_style(mut self, index: u16) -> Self {
        self.base_style = Some(index);
        self
    }

    /// Set the following-paragraph style index.
    pub fn with_next_style(mut self, index: u16) -> Self {
        self.next_style = index;
        self
    }

    /// Add an alternate style name.
    pub fn with_alias(mut self, alias: impl Into<String>) -> Self {
        self.aliases.push(alias.into());
        self
    }

    /// Replace the kind-specific raw UPX payloads.
    pub fn with_property_sets(mut self, property_sets: Vec<Vec<u8>>) -> Self {
        self.property_sets = property_sets;
        self
    }

    /// Attach Word 2000-and-later metadata.
    pub fn with_post_2000(mut self, metadata: StylePost2000) -> Self {
        self.post_2000 = Some(metadata);
        self
    }
}

/// Error returned when a custom DOC stylesheet cannot be represented safely.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleWriteError(String);

impl std::fmt::Display for StyleWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for StyleWriteError {}

fn invalid(message: impl Into<String>) -> StyleWriteError {
    StyleWriteError(message.into())
}

fn required_property_count(style: &DocStyleDefinition) -> Result<usize, StyleWriteError> {
    let revision_marked = style
        .post_2000
        .as_ref()
        .is_some_and(|metadata| metadata.has_original_style);
    if revision_marked {
        return Err(invalid(
            "revision-marked DOC style emission requires typed revision metadata",
        ));
    }
    match (style.kind, revision_marked) {
        (StyleKind::Paragraph, false) => Ok(2),
        (StyleKind::Character, false) => Ok(1),
        (StyleKind::Table, false) => Ok(3),
        (StyleKind::Numbering, false) => Ok(1),
        (_, true) => unreachable!(),
    }
}

fn validate_grpprl(
    bytes: &[u8],
    expected_type: u8,
    forbidden_conditional: Option<u16>,
    description: &str,
) -> Result<(), StyleWriteError> {
    let sprms = parse_sprms(bytes);
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != bytes.len() {
        return Err(invalid(format!("{description} contains a truncated SPRM")));
    }
    if sprms
        .iter()
        .any(|sprm| get_sprm_type(sprm.opcode) != expected_type)
    {
        return Err(invalid(format!(
            "{description} contains an SPRM for the wrong property group"
        )));
    }
    if forbidden_conditional.is_some_and(|opcode| sprms.iter().any(|sprm| sprm.opcode == opcode)) {
        return Err(invalid(format!(
            "{description} contains conditional formatting outside a table style"
        )));
    }
    Ok(())
}

fn validate_style(style: &DocStyleDefinition, index: u16) -> Result<(), StyleWriteError> {
    if style.invariant_id > USER_STYLE_ID {
        return Err(invalid(
            "DOC style invariant identifiers cannot exceed 0x0FFE",
        ));
    }
    if style.name.is_empty()
        || style.name.contains(',')
        || style
            .aliases
            .iter()
            .any(|alias| alias.is_empty() || alias.contains(','))
    {
        return Err(invalid(
            "DOC style names and aliases must be nonempty and cannot contain commas",
        ));
    }
    if style.next_style > USER_STYLE_ID || style.base_style.is_some_and(|base| base > USER_STYLE_ID)
    {
        return Err(invalid("DOC style references cannot exceed 0x0FFE"));
    }
    let expected = required_property_count(style)?;
    if style.property_sets.len() != expected {
        return Err(invalid(format!(
            "DOC style {index} has {} UPX records; expected {expected}",
            style.property_sets.len()
        )));
    }
    if let Some(metadata) = &style.post_2000 {
        if metadata.priority > 99 || metadata.html_font_category > 7 {
            return Err(invalid(
                "DOC post-2000 style metadata is outside its bit-field range",
            ));
        }
        if metadata
            .linked_style
            .is_some_and(|linked| linked == 0 || linked > USER_STYLE_ID)
        {
            return Err(invalid(
                "DOC linked style index must be between 1 and 0x0FFE",
            ));
        }
    }
    match style.kind {
        StyleKind::Paragraph => {
            validate_grpprl(
                &style.property_sets[0],
                1,
                Some(SPRM_P_CNF),
                "paragraph-style UpxPapx",
            )?;
            validate_grpprl(
                &style.property_sets[1],
                2,
                Some(SPRM_C_CNF),
                "paragraph-style UpxChpx",
            )?;
        },
        StyleKind::Character => validate_grpprl(
            &style.property_sets[0],
            2,
            Some(SPRM_C_CNF),
            "character-style UpxChpx",
        )?,
        StyleKind::Table => {
            validate_grpprl(&style.property_sets[0], 5, None, "table-style UpxTapx")?;
            validate_grpprl(&style.property_sets[1], 1, None, "table-style UpxPapx")?;
            validate_grpprl(&style.property_sets[2], 2, None, "table-style UpxChpx")?;
        },
        StyleKind::Numbering => validate_grpprl(
            &style.property_sets[0],
            1,
            Some(SPRM_P_CNF),
            "numbering-style UpxPapx",
        )?,
    }
    Ok(())
}

fn kind_code(kind: StyleKind) -> u16 {
    match kind {
        StyleKind::Paragraph => 1,
        StyleKind::Character => 2,
        StyleKind::Table => 3,
        StyleKind::Numbering => 4,
    }
}

fn flags_word(flags: &StyleFlags) -> u16 {
    u16::from(flags.auto_redefine)
        | (u16::from(flags.hidden) << 1)
        | (u16::from(flags.legacy_languages_set) << 2)
        | (u16::from(flags.copy_language) << 3)
        | (u16::from(flags.personal_compose) << 4)
        | (u16::from(flags.personal_reply) << 5)
        | (u16::from(flags.personal) << 6)
        | (u16::from(flags.semi_hidden) << 8)
        | (u16::from(flags.locked) << 9)
        | (u16::from(flags.unhide_when_used) << 11)
        | (u16::from(flags.quick_format) << 12)
}

fn serialize_style(
    style: &DocStyleDefinition,
    index: u16,
    stdf_size: usize,
) -> Result<Vec<u8>, StyleWriteError> {
    validate_style(style, index)?;
    let property_count = style.property_sets.len();
    let mut bytes = vec![0u8; stdf_size];
    let info1 = style.invariant_id | (u16::from(style.flags.invalidate_height) << 13);
    let info2 = kind_code(style.kind) | (style.base_style.unwrap_or(NIL_STYLE) << 4);
    let info3 = property_count as u16 | (style.next_style << 4);
    bytes[0..2].copy_from_slice(&info1.to_le_bytes());
    bytes[2..4].copy_from_slice(&info2.to_le_bytes());
    bytes[4..6].copy_from_slice(&info3.to_le_bytes());
    bytes[8..10].copy_from_slice(&flags_word(&style.flags).to_le_bytes());
    if stdf_size == 18
        && let Some(metadata) = &style.post_2000
    {
        let post_info1 =
            metadata.linked_style.unwrap_or(0) | (u16::from(metadata.has_original_style) << 12);
        let post_info3 = u16::from(metadata.html_font_category) | (metadata.priority << 4);
        bytes[10..12].copy_from_slice(&post_info1.to_le_bytes());
        bytes[12..16].copy_from_slice(&metadata.revision_id.to_le_bytes());
        bytes[16..18].copy_from_slice(&post_info3.to_le_bytes());
    }

    let combined_name = std::iter::once(style.name.as_str())
        .chain(style.aliases.iter().map(String::as_str))
        .collect::<Vec<_>>()
        .join(",");
    let name = combined_name.encode_utf16().collect::<Vec<_>>();
    let name_len = u16::try_from(name.len()).map_err(|_| invalid("DOC style name is too long"))?;
    bytes.extend_from_slice(&name_len.to_le_bytes());
    bytes.extend(name.into_iter().flat_map(u16::to_le_bytes));
    bytes.extend_from_slice(&0u16.to_le_bytes());
    for property_set in &style.property_sets {
        let size = u16::try_from(property_set.len())
            .map_err(|_| invalid("DOC style UPX payload exceeds 65535 bytes"))?;
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(property_set);
        if size % 2 != 0 {
            bytes.push(0);
        }
    }
    let size = u16::try_from(bytes.len())
        .map_err(|_| invalid("DOC STD exceeds the 65535-byte representation limit"))?;
    if size > i16::MAX as u16 {
        return Err(invalid("DOC STD exceeds the signed LPStd size range"));
    }
    bytes[6..8].copy_from_slice(&size.to_le_bytes());
    Ok(bytes)
}

fn normal_style() -> DocStyleDefinition {
    let mut style = DocStyleDefinition::new(StyleKind::Paragraph, "Normal");
    style.invariant_id = 0;
    style.next_style = 0;
    style
}

fn default_font_style() -> DocStyleDefinition {
    let mut style = DocStyleDefinition::new(StyleKind::Character, "Default Paragraph Font");
    style.invariant_id = 65;
    style.next_style = 10;
    style
}

/// Generate a complete stylesheet containing required built-ins and custom styles.
pub fn generate_stylesheet(
    custom_styles: &[DocStyleDefinition],
) -> Result<Vec<u8>, StyleWriteError> {
    let style_count = MIN_STYLE_COUNT
        .checked_add(custom_styles.len())
        .ok_or_else(|| invalid("DOC stylesheet style count overflows"))?;
    if style_count > MAX_STYLE_COUNT {
        return Err(invalid("DOC stylesheet exceeds 4093 style slots"));
    }
    let stdf_size = if custom_styles.iter().any(|style| style.post_2000.is_some()) {
        18usize
    } else {
        10
    };
    let mut stsh = Vec::new();
    stsh.extend_from_slice(&18u16.to_le_bytes());
    stsh.extend_from_slice(&(style_count as u16).to_le_bytes());
    stsh.extend_from_slice(&(stdf_size as u16).to_le_bytes());
    stsh.extend_from_slice(&1u16.to_le_bytes());
    stsh.extend_from_slice(&15u16.to_le_bytes());
    stsh.extend_from_slice(&15u16.to_le_bytes());
    stsh.extend_from_slice(&0u16.to_le_bytes());
    stsh.extend_from_slice(&0i16.to_le_bytes());
    stsh.extend_from_slice(&0i16.to_le_bytes());
    stsh.extend_from_slice(&0i16.to_le_bytes());

    for index in 0..style_count {
        let style = match index {
            0 => Some(normal_style()),
            10 => Some(default_font_style()),
            15.. => Some(custom_styles[index - 15].clone()),
            _ => None,
        };
        if let Some(style) = style {
            let bytes = serialize_style(&style, index as u16, stdf_size)?;
            stsh.extend_from_slice(&(bytes.len() as u16).to_le_bytes());
            stsh.extend_from_slice(&bytes);
            if bytes.len() % 2 != 0 {
                stsh.push(0);
            }
        } else {
            stsh.extend_from_slice(&0u16.to_le_bytes());
        }
    }
    StyleSheet::parse_data(&stsh, 0)
        .map_err(|error| invalid(format!("generated DOC stylesheet is invalid: {error}")))?;
    Ok(stsh)
}

/// Generate the mandatory minimal Word 97+ stylesheet.
pub fn generate_minimal_stylesheet() -> Vec<u8> {
    generate_stylesheet(&[]).expect("the built-in DOC stylesheet is valid")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::parts::chp::CharacterProperties;
    use crate::doc::parts::pap::ParagraphProperties;
    use crate::doc::parts::tap::{
        TableConditionalFormatting, TableStyleCondition, TableStyleDefaults,
    };
    use crate::doc::writer::tap::generate_table_style_sprms_with_conditionals;
    use crate::sprm_operations::{SPRM_C_F_BOLD, SPRM_P_F_KEEP};

    fn conditional(opcode: u16, condition: u16, nested: &[u8]) -> Vec<u8> {
        let mut bytes = opcode.to_le_bytes().to_vec();
        bytes.push((nested.len() + 2) as u8);
        bytes.extend_from_slice(&condition.to_le_bytes());
        bytes.extend_from_slice(nested);
        bytes
    }

    #[test]
    fn minimal_stylesheet_round_trips() {
        let bytes = generate_minimal_stylesheet();
        let parsed = StyleSheet::parse_data(&bytes, 0).unwrap();
        assert_eq!(parsed.styles().len(), 15);
        assert_eq!(parsed.get(0).unwrap().name, "Normal");
        assert_eq!(parsed.get(10).unwrap().name, "Default Paragraph Font");
    }

    #[test]
    fn custom_table_style_round_trips_all_conditional_domains() {
        let tapx = generate_table_style_sprms_with_conditionals(
            &TableStyleDefaults::default(),
            &[TableConditionalFormatting {
                condition: TableStyleCondition::HeaderRow,
                properties: TableStyleDefaults {
                    no_wrap: Some(true),
                    ..TableStyleDefaults::default()
                },
                raw_grpprl: Vec::new(),
            }],
        )
        .unwrap();
        let pap_nested = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        let papx = conditional(SPRM_P_CNF, 0x0001, &pap_nested);
        let chp_nested = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        let chpx = conditional(SPRM_C_CNF, 0x0008, &chp_nested);
        let style = DocStyleDefinition::new(StyleKind::Table, "Grid Accent")
            .with_alias("Accent Grid")
            .with_property_sets(vec![tapx, papx, chpx])
            .with_post_2000(StylePost2000 {
                linked_style: None,
                has_original_style: false,
                revision_id: 0x1122_3344,
                html_font_category: 2,
                priority: 42,
            });

        let bytes = generate_stylesheet(&[style]).unwrap();
        let parsed = StyleSheet::parse_data(&bytes, 0).unwrap();
        assert_eq!(parsed.header().stdf_size, 18);
        let style = parsed.get(15).unwrap();
        assert_eq!(style.aliases, ["Accent Grid"]);
        assert_eq!(style.post_2000.as_ref().unwrap().priority, 42);
        let (_, table) = parsed.resolve_table_properties(15).unwrap();
        assert_eq!(table.conditional_formats.len(), 1);
        assert_eq!(table.conditional_formats[0].properties.no_wrap, Some(true));
        let paragraph =
            ParagraphProperties::from_sprm(style.paragraph_properties().unwrap()).unwrap();
        assert!(paragraph.conditional_formats[0].properties.keep_on_page);
        let character =
            CharacterProperties::from_sprm(style.character_properties().unwrap()).unwrap();
        assert_eq!(
            character.conditional_formats[0].properties.is_bold,
            Some(true)
        );
    }

    #[test]
    fn rejects_invalid_custom_styles() {
        let wrong_count =
            DocStyleDefinition::new(StyleKind::Table, "Wrong").with_property_sets(vec![Vec::new()]);
        assert!(generate_stylesheet(&[wrong_count]).is_err());

        let wrong_type = DocStyleDefinition::new(StyleKind::Table, "Wrong Type")
            .with_property_sets(vec![
                [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat(),
                Vec::new(),
                Vec::new(),
            ]);
        assert!(generate_stylesheet(&[wrong_type]).is_err());

        let conditional_paragraph = DocStyleDefinition::new(StyleKind::Paragraph, "Not Table")
            .with_property_sets(vec![conditional(SPRM_P_CNF, 1, &[]), Vec::new()]);
        assert!(generate_stylesheet(&[conditional_paragraph]).is_err());

        let self_based = DocStyleDefinition::new(StyleKind::Table, "Cycle").with_base_style(15);
        assert!(generate_stylesheet(&[self_based]).is_err());

        let duplicate = DocStyleDefinition::new(StyleKind::Table, "Normal");
        assert!(generate_stylesheet(&[duplicate]).is_err());

        let revision_marked = DocStyleDefinition::new(StyleKind::Character, "Revised")
            .with_post_2000(StylePost2000 {
                linked_style: None,
                has_original_style: true,
                revision_id: 1,
                html_font_category: 0,
                priority: 0,
            });
        assert!(generate_stylesheet(&[revision_marked]).is_err());
    }
}

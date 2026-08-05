//! Validated Word 97+ stylesheet (`STSH`) generation.

use crate::CommentDateTime;
use crate::parts::styles::{StyleFlags, StyleKind, StylePost2000, StyleSheet};
use std::collections::HashMap;

const MIN_STYLE_COUNT: usize = 15;
const MAX_STYLE_COUNT: usize = 0x0FFD;
const NIL_STYLE: u16 = 0x0FFF;
const USER_STYLE_ID: u16 = 0x0FFE;

/// Previous formatting and attribution retained by a revision-marked style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocStyleRevision {
    /// Revision author name, stored through the document's `SttbfRMark` table.
    pub author: String,
    /// Date and time at which the style was revision-marked.
    pub timestamp: Option<CommentDateTime>,
    /// Previous paragraph formatting; present only for a paragraph style.
    pub paragraph_properties: Option<Vec<u8>>,
    /// Previous character formatting.
    pub character_properties: Vec<u8>,
}

impl DocStyleRevision {
    /// Create revision state for a paragraph style.
    pub fn paragraph(
        author: impl Into<String>,
        paragraph_properties: Vec<u8>,
        character_properties: Vec<u8>,
    ) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            paragraph_properties: Some(paragraph_properties),
            character_properties,
        }
    }

    /// Create revision state for a character style.
    pub fn character(author: impl Into<String>, character_properties: Vec<u8>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            paragraph_properties: None,
            character_properties,
        }
    }

    /// Set the style-revision timestamp.
    pub fn with_timestamp(mut self, timestamp: CommentDateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }
}

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
    /// Previous formatting and attribution for a revision-marked style.
    pub revision: Option<DocStyleRevision>,
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
            revision: None,
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

    /// Mark this paragraph or character style as revised and retain its prior formatting.
    pub fn with_revision(mut self, revision: DocStyleRevision) -> Self {
        self.revision = Some(revision);
        self.post_2000
            .get_or_insert(StylePost2000 {
                linked_style: None,
                has_original_style: true,
                revision_id: 0,
                html_font_category: 0,
                priority: 0,
            })
            .has_original_style = true;
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
    if revision_marked != style.revision.is_some() {
        return Err(invalid(
            "DOC style revision metadata and fHasOriginalStyle must be present together",
        ));
    }
    match (style.kind, revision_marked) {
        (StyleKind::Paragraph, false) => Ok(2),
        (StyleKind::Paragraph, true) => Ok(3),
        (StyleKind::Character, false) => Ok(1),
        (StyleKind::Character, true) => Ok(2),
        (StyleKind::Table, false) => Ok(3),
        (StyleKind::Numbering, false) => Ok(1),
        (StyleKind::Table | StyleKind::Numbering, true) => Err(invalid(
            "DOC table and numbering styles cannot be revision-marked",
        )),
    }
}

fn current_property_count(kind: StyleKind) -> usize {
    match kind {
        StyleKind::Paragraph => 2,
        StyleKind::Character | StyleKind::Numbering => 1,
        StyleKind::Table => 3,
    }
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
    let expected_current = current_property_count(style.kind);
    if style.property_sets.len() != expected_current {
        return Err(invalid(format!(
            "DOC style {index} has {} current UPX records; expected {expected_current}",
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
            crate::parts::styles::validate_paragraph_style_sprms(
                &style.property_sets[0],
                index,
                false,
            )
            .map_err(|error| invalid(error.to_string()))?;
            crate::parts::styles::validate_character_style_sprms(&style.property_sets[1], false)
                .map_err(|error| invalid(error.to_string()))?;
        },
        StyleKind::Character => {
            crate::parts::styles::validate_character_style_sprms(&style.property_sets[0], false)
                .map_err(|error| invalid(error.to_string()))?;
        },
        StyleKind::Table => {
            crate::parts::styles::validate_table_style_sprms(&style.property_sets[0], index, false)
                .map_err(|error| invalid(error.to_string()))?;
            crate::parts::styles::validate_paragraph_style_sprms(
                &style.property_sets[1],
                index,
                true,
            )
            .map_err(|error| invalid(error.to_string()))?;
            crate::parts::styles::validate_character_style_sprms(&style.property_sets[2], true)
                .map_err(|error| invalid(error.to_string()))?;
        },
        StyleKind::Numbering => {
            crate::parts::styles::validate_numbering_style_sprms(&style.property_sets[0], index)
                .map_err(|error| invalid(error.to_string()))?;
        },
    }
    if let Some(revision) = &style.revision {
        if revision.author.is_empty() {
            return Err(invalid("DOC style revision author must not be empty"));
        }
        match (style.kind, &revision.paragraph_properties) {
            (StyleKind::Paragraph, Some(paragraph)) => {
                crate::parts::styles::validate_paragraph_style_sprms(paragraph, index, false)
                    .map_err(|error| invalid(error.to_string()))?;
            },
            (StyleKind::Character, None) => {},
            (StyleKind::Paragraph, None) => {
                return Err(invalid(
                    "paragraph style revision is missing prior paragraph formatting",
                ));
            },
            (StyleKind::Character, Some(_)) => {
                return Err(invalid(
                    "character style revision cannot contain paragraph formatting",
                ));
            },
            (StyleKind::Table | StyleKind::Numbering, _) => unreachable!(),
        }
        crate::parts::styles::validate_character_style_sprms(&revision.character_properties, false)
            .map_err(|error| invalid(error.to_string()))?;
    }
    debug_assert_eq!(
        expected,
        style.property_sets.len() + usize::from(style.revision.is_some())
    );
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
    revision_authors: Option<&HashMap<String, u16>>,
) -> Result<Vec<u8>, StyleWriteError> {
    validate_style(style, index)?;
    let property_count = required_property_count(style)?;
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
    if let Some(revision) = &style.revision {
        let author_index = revision_authors
            .and_then(|authors| authors.get(&revision.author))
            .copied()
            .ok_or_else(|| invalid("DOC style revision author was not indexed"))?;
        if author_index > i16::MAX as u16 {
            return Err(invalid(
                "DOC style revision author exceeds the signed author-index range",
            ));
        }
        let mut revision_payload = Vec::new();
        revision_payload.extend_from_slice(&6u16.to_le_bytes());
        revision_payload.extend_from_slice(
            &super::core::pack_dttm(revision.timestamp)
                .map_err(|error| invalid(error.to_string()))?
                .to_le_bytes(),
        );
        revision_payload.extend_from_slice(&(author_index as i16).to_le_bytes());
        if let Some(paragraph) = &revision.paragraph_properties {
            append_inner_upx(&mut revision_payload, paragraph)?;
        }
        append_inner_upx(&mut revision_payload, &revision.character_properties)?;
        debug_assert_eq!(revision_payload.len() % 2, 0);
        let size = u16::try_from(revision_payload.len())
            .map_err(|_| invalid("DOC style revision payload exceeds 65535 bytes"))?;
        bytes.extend_from_slice(&size.to_le_bytes());
        bytes.extend_from_slice(&revision_payload);
    }
    let size = u16::try_from(bytes.len())
        .map_err(|_| invalid("DOC STD exceeds the 65535-byte representation limit"))?;
    if size > i16::MAX as u16 {
        return Err(invalid("DOC STD exceeds the signed LPStd size range"));
    }
    bytes[6..8].copy_from_slice(&size.to_le_bytes());
    Ok(bytes)
}

fn append_inner_upx(output: &mut Vec<u8>, property_set: &[u8]) -> Result<(), StyleWriteError> {
    let size = u16::try_from(property_set.len())
        .map_err(|_| invalid("DOC revision-marked style UPX exceeds 65535 bytes"))?;
    output.extend_from_slice(&size.to_le_bytes());
    output.extend_from_slice(property_set);
    if size % 2 != 0 {
        output.push(0);
    }
    Ok(())
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
    revision_authors: Option<&HashMap<String, u16>>,
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
            let bytes = serialize_style(&style, index as u16, stdf_size, revision_authors)?;
            stsh.extend_from_slice(&(bytes.len() as u16).to_le_bytes());
            stsh.extend_from_slice(&bytes);
            if bytes.len() % 2 != 0 {
                stsh.push(0);
            }
        } else {
            stsh.extend_from_slice(&0u16.to_le_bytes());
        }
    }
    StyleSheet::parse_data(&stsh, 0, crate::DocLeniency::Strict)
        .map_err(|error| invalid(format!("generated DOC stylesheet is invalid: {error}")))?;
    Ok(stsh)
}

/// Generate the mandatory minimal Word 97+ stylesheet.
pub fn generate_minimal_stylesheet() -> Vec<u8> {
    generate_stylesheet(&[], None).expect("the built-in DOC stylesheet is valid")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parts::chp::CharacterProperties;
    use crate::parts::pap::ParagraphProperties;
    use crate::parts::tap::{TableConditionalFormatting, TableStyleCondition, TableStyleDefaults};
    use crate::sprm_operations::{SPRM_C_CNF, SPRM_C_F_BOLD, SPRM_P_CNF, SPRM_P_F_KEEP};
    use crate::writer::tap::generate_table_style_sprms_with_conditionals;

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
        let parsed = StyleSheet::parse_data(&bytes, 0, crate::DocLeniency::Strict).unwrap();
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

        let bytes = generate_stylesheet(&[style], None).unwrap();
        let parsed = StyleSheet::parse_data(&bytes, 0, crate::DocLeniency::Strict).unwrap();
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
    fn revision_marked_paragraph_style_round_trips_nested_upx() {
        let current_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        let current_chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        let previous_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[0]].concat();
        let previous_chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
        let timestamp = CommentDateTime {
            year: 2026,
            month: 7,
            day: 16,
            hour: 10,
            minute: 30,
            weekday: 4,
        };
        let style = DocStyleDefinition::new(StyleKind::Paragraph, "Tracked Body")
            .with_property_sets(vec![current_papx, current_chpx])
            .with_revision(
                DocStyleRevision::paragraph(
                    "Style Editor",
                    previous_papx.clone(),
                    previous_chpx.clone(),
                )
                .with_timestamp(timestamp),
            );
        let authors = HashMap::from([("Style Editor".to_string(), 3u16)]);

        let bytes = generate_stylesheet(&[style], Some(&authors)).unwrap();
        let parsed = StyleSheet::parse_data(&bytes, 0, crate::DocLeniency::Strict).unwrap();
        let style = parsed.get(15).unwrap();
        assert!(style.post_2000.as_ref().unwrap().has_original_style);
        assert_eq!(style.property_sets.len(), 3);
        let revision = style.revision.as_ref().unwrap();
        assert_eq!(revision.author_index, 3);
        assert_eq!(revision.timestamp, Some(timestamp));
        assert_eq!(
            revision.paragraph_properties.as_deref(),
            Some(previous_papx.as_slice())
        );
        assert_eq!(revision.character_properties, previous_chpx);
    }

    #[test]
    fn revision_marked_character_style_round_trips_nested_upx() {
        let previous = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
        let style =
            DocStyleDefinition::new(StyleKind::Character, "Tracked Emphasis").with_revision(
                DocStyleRevision::character("Style Editor", previous.clone()),
            );
        let authors = HashMap::from([("Style Editor".to_string(), 1u16)]);

        let bytes = generate_stylesheet(&[style], Some(&authors)).unwrap();
        let parsed = StyleSheet::parse_data(&bytes, 0, crate::DocLeniency::Strict).unwrap();
        let revision = parsed.get(15).unwrap().revision.as_ref().unwrap();
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.paragraph_properties, None);
        assert_eq!(revision.character_properties, previous);
    }

    #[test]
    fn rejects_invalid_custom_styles() {
        let wrong_count =
            DocStyleDefinition::new(StyleKind::Table, "Wrong").with_property_sets(vec![Vec::new()]);
        assert!(generate_stylesheet(&[wrong_count], None).is_err());

        let wrong_type = DocStyleDefinition::new(StyleKind::Table, "Wrong Type")
            .with_property_sets(vec![
                [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat(),
                Vec::new(),
                Vec::new(),
            ]);
        assert!(generate_stylesheet(&[wrong_type], None).is_err());

        let conditional_paragraph = DocStyleDefinition::new(StyleKind::Paragraph, "Not Table")
            .with_property_sets(vec![conditional(SPRM_P_CNF, 1, &[]), Vec::new()]);
        assert!(generate_stylesheet(&[conditional_paragraph], None).is_err());

        let self_based = DocStyleDefinition::new(StyleKind::Table, "Cycle").with_base_style(15);
        assert!(generate_stylesheet(&[self_based], None).is_err());

        let duplicate = DocStyleDefinition::new(StyleKind::Table, "Normal");
        assert!(generate_stylesheet(&[duplicate], None).is_err());

        let revision_marked = DocStyleDefinition::new(StyleKind::Character, "Revised")
            .with_post_2000(StylePost2000 {
                linked_style: None,
                has_original_style: true,
                revision_id: 1,
                html_font_category: 0,
                priority: 0,
            });
        assert!(generate_stylesheet(&[revision_marked], None).is_err());
    }
}

//! Typed style inheritance and property resolution.

use super::super::codec::{corrupted, read_u16};
use super::super::model::{StyleDefinition, StyleKind, StyleSheet};
use super::validation::{strip_paragraph_style_index, validate_style_sprms};
use crate::package::Result;
use crate::parts::tap::{TableProperties, TableStyleCondition};
use crate::sprm_operations::get_sprm_type;

impl StyleSheet {
    pub(crate) fn resolve_revision_authors(
        &mut self,
        authors: &crate::parts::revisions::RevisionAuthorTable,
    ) -> Result<()> {
        for style in self.styles.iter_mut().flatten() {
            let Some(revision) = &mut style.revision else {
                continue;
            };
            let index = u16::try_from(revision.author_index).map_err(|_| {
                corrupted("revision-marked style author index must not be negative")
            })?;
            let author = authors.get(index).ok_or_else(|| {
                corrupted("revision-marked style author index is outside SttbfRMark")
            })?;
            revision.author = Some(author.to_string());
        }
        Ok(())
    }

    /// Resolve the table-property differences for a requested table style.
    ///
    /// The returned index is the effective style index. MS-DOC defines an
    /// empty, missing, or non-table style reference as the default table style
    /// at fixed index 11. Property arrays are concatenated parent-first, then
    /// parsed from defaults so the derived style overrides its ancestors.
    pub fn resolve_table_properties(&self, requested_index: u16) -> Result<(u16, TableProperties)> {
        let effective_index = self.effective_table_style_index(requested_index);
        let mut chain = Vec::new();
        let mut current = self.get(effective_index);
        while let Some(style) = current {
            chain.push(style);
            current = style.base_style.and_then(|index| self.get(index));
        }
        chain.reverse();

        let byte_count = chain
            .iter()
            .filter_map(|style| style.table_properties())
            .try_fold(0usize, |total, properties| {
                total
                    .checked_add(properties.len())
                    .ok_or_else(|| corrupted("resolved table style is too large"))
            })?;
        let mut grpprl = Vec::with_capacity(byte_count);
        for properties in chain
            .into_iter()
            .filter_map(StyleDefinition::table_properties)
        {
            grpprl.extend_from_slice(properties);
        }

        let arena = bumpalo::Bump::new();
        let mut properties = crate::parts::tap_parser::TapParser::new(&arena).parse_tap(&grpprl)?;
        // A sprmTIstd inside UpxTapx is explicitly ignored by MS-DOC.
        properties.table_style_index = None;
        Ok((effective_index, properties))
    }

    /// Resolve the paragraph and character properties contributed by a table style.
    ///
    /// The returned index is the effective table style index, using fixed style
    /// 11 when the requested slot is empty, missing, or not a table style. Style
    /// properties are applied parent-first so a derived table style overrides its
    /// ancestors. Conditional `sprmPCnf` and `sprmCCnf` records remain available
    /// in the returned properties for the table layout pass to apply by cell.
    pub fn resolve_table_text_properties(
        &self,
        requested_index: u16,
    ) -> Result<(
        u16,
        crate::parts::pap::ParagraphProperties,
        crate::parts::chp::CharacterProperties,
    )> {
        let (effective_index, paragraph, character) =
            self.resolve_table_text_style_sprms(requested_index)?;
        Ok((
            effective_index,
            crate::parts::pap::ParagraphProperties::from_sprm(&paragraph)?,
            crate::parts::chp::CharacterProperties::from_sprm(&character)?,
        ))
    }

    pub(crate) fn resolve_table_text_style_sprms(
        &self,
        requested_index: u16,
    ) -> Result<(u16, Vec<u8>, Vec<u8>)> {
        let effective_index = self.effective_table_style_index(requested_index);
        let Some(style) = self
            .get(effective_index)
            .filter(|style| style.kind == StyleKind::Table)
        else {
            return Ok((effective_index, Vec::new(), Vec::new()));
        };

        let mut paragraph = Vec::new();
        let mut character = Vec::new();
        for style in self.style_chain(style) {
            if style.kind != StyleKind::Table {
                continue;
            }
            if let Some(properties) = style.paragraph_properties() {
                let properties = strip_paragraph_style_index(properties, style.index)?;
                validate_style_sprms(properties, 1, "table-style UpxPapx")?;
                paragraph.extend_from_slice(properties);
            }
            if let Some(properties) = style.character_properties() {
                validate_style_sprms(properties, 2, "table-style UpxChpx")?;
                character.extend_from_slice(properties);
            }
        }

        // Validate nested conditional payloads even when a particular table
        // position will not select them later.
        crate::parts::pap::ParagraphProperties::from_sprm(&paragraph)?;
        crate::parts::chp::CharacterProperties::from_sprm(&character)?;

        Ok((effective_index, paragraph, character))
    }

    pub(crate) fn resolve_table_text_style_sprms_for_conditions(
        &self,
        requested_index: u16,
        conditions: &[TableStyleCondition],
    ) -> Result<(u16, Vec<u8>, Vec<u8>)> {
        let (effective, paragraph, character) =
            self.resolve_table_text_style_sprms(requested_index)?;
        Ok((
            effective,
            flatten_conditional_style_sprms(
                &paragraph,
                crate::sprm_operations::SPRM_P_CNF,
                conditions,
            )?,
            flatten_conditional_style_sprms(
                &character,
                crate::sprm_operations::SPRM_C_CNF,
                conditions,
            )?,
        ))
    }

    /// Resolve the paragraph and character property differences contributed by
    /// a paragraph style, in parent-first application order.
    ///
    /// An invalid, empty, or non-paragraph style produces document defaults and
    /// therefore returns `None` with empty property arrays.
    pub fn resolve_paragraph_style_sprms(
        &self,
        requested_index: u16,
    ) -> Result<(Option<u16>, Vec<u8>, Vec<u8>)> {
        let Some(style) = self
            .get(requested_index)
            .filter(|style| style.kind == StyleKind::Paragraph)
        else {
            return Ok((None, Vec::new(), Vec::new()));
        };
        let chain = self.style_chain(style);
        let mut paragraph = Vec::new();
        let mut character = Vec::new();
        for style in chain {
            if let Some(properties) = style.paragraph_properties() {
                let properties = strip_paragraph_style_index(properties, style.index)?;
                validate_style_sprms(properties, 1, "UpxPapx")?;
                paragraph.extend_from_slice(properties);
            }
            if let Some(properties) = style.character_properties() {
                validate_style_sprms(properties, 2, "UpxChpx")?;
                character.extend_from_slice(properties);
            }
        }
        Ok((Some(requested_index), paragraph, character))
    }

    /// Resolve character property differences for a character style in
    /// parent-first application order.
    pub fn resolve_character_style_sprms(
        &self,
        requested_index: u16,
    ) -> Result<(Option<u16>, Vec<u8>)> {
        let Some(style) = self
            .get(requested_index)
            .filter(|style| style.kind == StyleKind::Character)
        else {
            return Ok((None, Vec::new()));
        };
        let mut character = Vec::new();
        for style in self.style_chain(style) {
            if let Some(properties) = style.character_properties() {
                validate_style_sprms(properties, 2, "UpxChpx")?;
                character.extend_from_slice(properties);
            }
        }
        Ok((Some(requested_index), character))
    }

    fn style_chain<'a>(&'a self, style: &'a StyleDefinition) -> Vec<&'a StyleDefinition> {
        let mut chain = Vec::new();
        let mut current = Some(style);
        while let Some(style) = current {
            chain.push(style);
            current = style.base_style.and_then(|index| self.get(index));
        }
        chain.reverse();
        chain
    }

    fn effective_table_style_index(&self, requested_index: u16) -> u16 {
        if self
            .get(requested_index)
            .is_some_and(|style| style.kind == StyleKind::Table)
        {
            requested_index
        } else {
            11
        }
    }
}

pub(in crate::parts::styles) fn flatten_conditional_style_sprms(
    properties: &[u8],
    conditional_opcode: u16,
    conditions: &[TableStyleCondition],
) -> Result<Vec<u8>> {
    let sprms = validate_style_sprms(
        properties,
        get_sprm_type(conditional_opcode),
        "conditional table-style property set",
    )?;
    let mut flattened = Vec::with_capacity(properties.len());
    for sprm in &sprms {
        if sprm.opcode != conditional_opcode {
            flattened.extend_from_slice(&properties[sprm.offset..sprm.offset + sprm.size]);
        }
    }
    for condition in conditions {
        for sprm in &sprms {
            if sprm.opcode != conditional_opcode {
                continue;
            }
            let operand = sprm.operand_bytes();
            let code = read_u16(operand, 0, "CNFOperand.cnfc")?;
            if code == condition.code() {
                flattened.extend_from_slice(&operand[2..]);
            }
        }
    }
    Ok(flattened)
}

//! Word 97+ stylesheet (`STSH`) parsing.

use std::collections::HashSet;

use super::super::leniency::{Leniency, StylesheetDefect, ToleranceReport};
use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use super::tap::TableProperties;
use crate::sprm_operations::{get_sprm_operation, get_sprm_type};

const STSH_POINTER_INDEX: usize = 1;
const STSHIF_SIZE: usize = 18;
const MIN_STYLE_COUNT: u16 = 0x000F;
const MAX_STYLE_COUNT: u16 = 0x0FFD;
const NIL_STYLE: u16 = 0x0FFF;

/// General stylesheet information stored in `Stshif`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleSheetHeader {
    /// Number of entries in [`StyleSheet::styles`].
    pub style_count: u16,
    /// Size of each fixed `Stdf` prefix (10 or 18 bytes).
    pub stdf_size: u16,
    /// Largest built-in style identifier known when the file was saved, plus one.
    pub max_builtin_style: u16,
    /// Count of fixed-index style slots. This is always 15 for Word 97+.
    pub fixed_style_count: u16,
    /// Built-in style-name version.
    pub builtin_name_version: u16,
    /// Default ASCII font index.
    pub ascii_font: i16,
    /// Default East Asian font index.
    pub east_asian_font: i16,
    /// Default non-ASCII font index.
    pub other_font: i16,
}

/// The four style kinds encoded by `StdfBase.stk`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleKind {
    /// Paragraph style.
    Paragraph,
    /// Character style.
    Character,
    /// Table style.
    Table,
    /// Numbering style.
    Numbering,
}

/// Miscellaneous flags stored in `StdfBase` and `GRFSTD`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct StyleFlags {
    /// Paragraph heights for this style need to be recalculated.
    pub invalidate_height: bool,
    /// User formatting is automatically merged into the style.
    pub auto_redefine: bool,
    /// The style is hidden from the application UI.
    pub hidden: bool,
    /// Legacy language compatibility properties have been applied.
    pub legacy_languages_set: bool,
    /// The legacy compatibility language represents no-proofing.
    pub copy_language: bool,
    /// Character style used for new e-mail messages.
    pub personal_compose: bool,
    /// Character style used for e-mail replies.
    pub personal_reply: bool,
    /// Character style used for e-mail senders.
    pub personal: bool,
    /// The style is hidden from the simplified styles UI.
    pub semi_hidden: bool,
    /// The style cannot be applied through the application UI.
    pub locked: bool,
    /// The style becomes visible after it is used.
    pub unhide_when_used: bool,
    /// The style is shown in the quick-style gallery.
    pub quick_format: bool,
}

/// Word 2000-and-later metadata appended to an `StdfBase`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StylePost2000 {
    /// Linked style index, or `None` when the encoded index is zero.
    pub linked_style: Option<u16>,
    /// Whether the style stores its pre-revision formatting.
    pub has_original_style: bool,
    /// Revision-save identifier of the last style modification.
    pub revision_id: u32,
    /// Legacy HTML font category.
    pub html_font_category: u8,
    /// UI ordering priority (0 through 99).
    pub priority: u16,
}

/// Previous formatting and attribution stored by a revision-marked style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleRevisionMark {
    /// Date and time at which the style was revision-marked.
    pub timestamp: Option<super::super::CommentDateTime>,
    /// Signed index into the document's `SttbfRMark` author table.
    pub author_index: i16,
    /// Resolved revision author when the stylesheet belongs to a complete document.
    pub author: Option<String>,
    /// Previous paragraph formatting for a paragraph style.
    pub paragraph_properties: Option<Vec<u8>>,
    /// Previous character formatting.
    pub character_properties: Vec<u8>,
}

/// One non-empty style definition from the stylesheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleDefinition {
    /// Zero-based index used by `sprmPIstd` and `sprmTIstd`.
    pub index: u16,
    /// Invariant built-in style identifier, or `0x0FFE` for a user style.
    pub invariant_id: u16,
    /// Style kind.
    pub kind: StyleKind,
    /// Parent style index, if this style inherits from another style.
    pub base_style: Option<u16>,
    /// Style automatically applied to the following paragraph.
    pub next_style: u16,
    /// Primary style name.
    pub name: String,
    /// Alternate names following the primary name in the `Xstz`.
    pub aliases: Vec<String>,
    /// Raw UPX payloads, in the kind-specific order prescribed by MS-DOC.
    pub property_sets: Vec<Vec<u8>>,
    /// Optional Word 2000-and-later metadata.
    pub post_2000: Option<StylePost2000>,
    /// Parsed revision attribution and previous formatting, when present.
    pub revision: Option<StyleRevisionMark>,
    /// Style behavior flags.
    pub flags: StyleFlags,
    /// Exact `STD` bytes, excluding the `LPStd` length and outer alignment byte.
    pub raw_std: Vec<u8>,
    /// Alignment byte following an odd-sized `STD`, when present.
    pub outer_padding: Option<u8>,
}

impl StyleDefinition {
    /// The table-property UPX for a table style.
    pub fn table_properties(&self) -> Option<&[u8]> {
        (self.kind == StyleKind::Table)
            .then(|| self.property_sets.first().map(Vec::as_slice))
            .flatten()
    }

    /// The current paragraph-property UPX for a paragraph or table style.
    pub fn paragraph_properties(&self) -> Option<&[u8]> {
        match self.kind {
            StyleKind::Paragraph => self.property_sets.first().map(Vec::as_slice),
            StyleKind::Table => self.property_sets.get(1).map(Vec::as_slice),
            StyleKind::Character | StyleKind::Numbering => None,
        }
    }

    /// The current character-property UPX for a paragraph, character, or table style.
    pub fn character_properties(&self) -> Option<&[u8]> {
        match self.kind {
            StyleKind::Paragraph => self.property_sets.get(1).map(Vec::as_slice),
            StyleKind::Character => self.property_sets.first().map(Vec::as_slice),
            StyleKind::Table => self.property_sets.get(2).map(Vec::as_slice),
            StyleKind::Numbering => None,
        }
    }
}

/// Parsed Word stylesheet with null style slots retained.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleSheet {
    /// Non-structural defects repaired during a lenient parse.
    tolerance: ToleranceReport,
    header: StyleSheetHeader,
    styles: Vec<Option<StyleDefinition>>,
    stshi_tail: Vec<u8>,
}

impl StyleSheet {
    /// Non-structural defects a lenient parse repaired.
    ///
    /// Always empty after a [`Leniency::Strict`] parse.
    #[inline]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance
    }

    /// Parse the mandatory Word 97+ stylesheet at FIB pointer index 1.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        Self::parse_with_leniency(fib, table_stream, Leniency::Strict)
    }

    /// Parse the stylesheet, optionally repairing non-structural defects.
    ///
    /// Under [`Leniency::Strict`] this behaves exactly like [`Self::parse`].
    pub fn parse_with_leniency(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        leniency: Leniency,
    ) -> Result<Self> {
        let (stream_offset, length) = fib
            .get_table_pointer(STSH_POINTER_INDEX)
            .filter(|(_, length)| *length != 0)
            .ok_or_else(|| {
                PackageError::Corrupted("FIB does not contain a stylesheet".to_string())
            })?;
        let start = usize::try_from(stream_offset)
            .map_err(|_| corrupted("stylesheet offset is too large"))?;
        let length =
            usize::try_from(length).map_err(|_| corrupted("stylesheet length is too large"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("stylesheet range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("stylesheet extends beyond the table stream"))?;
        Self::parse_data(data, start, leniency)
    }

    /// General stylesheet information.
    pub fn header(&self) -> &StyleSheetHeader {
        &self.header
    }

    /// All style slots, including required null fixed-index slots.
    pub fn styles(&self) -> &[Option<StyleDefinition>] {
        &self.styles
    }

    /// Resolve one style index to a non-empty definition.
    pub fn get(&self, index: u16) -> Option<&StyleDefinition> {
        self.styles.get(usize::from(index))?.as_ref()
    }

    /// Uninterpreted STSHI extension bytes following the 18-byte `Stshif`.
    pub fn stshi_tail(&self) -> &[u8] {
        &self.stshi_tail
    }

    pub(crate) fn resolve_revision_authors(
        &mut self,
        authors: &super::revisions::RevisionAuthorTable,
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
        let mut properties = super::tap_parser::TapParser::new(&arena).parse_tap(&grpprl)?;
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
        super::pap::ParagraphProperties,
        super::chp::CharacterProperties,
    )> {
        let (effective_index, paragraph, character) =
            self.resolve_table_text_style_sprms(requested_index)?;
        Ok((
            effective_index,
            super::pap::ParagraphProperties::from_sprm(&paragraph)?,
            super::chp::CharacterProperties::from_sprm(&character)?,
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
        super::pap::ParagraphProperties::from_sprm(&paragraph)?;
        super::chp::CharacterProperties::from_sprm(&character)?;

        Ok((effective_index, paragraph, character))
    }

    pub(crate) fn resolve_table_text_style_sprms_for_conditions(
        &self,
        requested_index: u16,
        conditions: &[super::tap::TableStyleCondition],
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

    pub(crate) fn parse_data(
        data: &[u8],
        stream_offset: usize,
        leniency: Leniency,
    ) -> Result<Self> {
        let cb_stshi = usize::from(read_u16(data, 0, "cbStshi")?);
        if cb_stshi < STSHIF_SIZE {
            return Err(corrupted("cbStshi is shorter than Stshif"));
        }
        let styles_offset = 2usize
            .checked_add(cb_stshi)
            .ok_or_else(|| corrupted("STSHI length overflows"))?;
        let stshi = data
            .get(2..styles_offset)
            .ok_or_else(|| corrupted("STSHI is truncated"))?;

        let style_count = read_u16(stshi, 0, "Stshif.cstd")?;
        if !(MIN_STYLE_COUNT..=MAX_STYLE_COUNT).contains(&style_count) {
            return Err(corrupted("Stshif.cstd is outside the valid range"));
        }
        let stdf_size = read_u16(stshi, 2, "Stshif.cbSTDBaseInFile")?;
        if stdf_size != 10 && stdf_size != 18 {
            return Err(corrupted("Stshif.cbSTDBaseInFile must be 10 or 18"));
        }
        let style_name_flags = read_u16(stshi, 4, "Stshif style-name flags")?;
        if style_name_flags != 1 {
            return Err(corrupted("Stshif style-name flags are invalid"));
        }
        let fixed_style_count = read_u16(stshi, 8, "Stshif.istdMaxFixedWhenSaved")?;
        if fixed_style_count != MIN_STYLE_COUNT {
            return Err(corrupted("Stshif fixed style count must be 15"));
        }
        let header = StyleSheetHeader {
            style_count,
            stdf_size,
            max_builtin_style: read_u16(stshi, 6, "Stshif.stiMaxWhenSaved")?,
            fixed_style_count,
            builtin_name_version: read_u16(stshi, 10, "Stshif name version")?,
            ascii_font: read_i16(stshi, 12, "Stshif.ftcAsci")?,
            east_asian_font: read_i16(stshi, 14, "Stshif.ftcFE")?,
            other_font: read_i16(stshi, 16, "Stshif.ftcOther")?,
        };

        let mut styles = Vec::with_capacity(usize::from(style_count));
        let mut offset = styles_offset;
        for index in 0..style_count {
            if stream_offset
                .checked_add(offset)
                .is_none_or(|absolute| absolute % 2 != 0)
            {
                return Err(corrupted("LPStd does not begin on an even-byte boundary"));
            }
            let cb_std = read_u16(data, offset, "LPStd.cbStd")?;
            if cb_std > i16::MAX as u16 {
                return Err(corrupted("LPStd.cbStd is negative"));
            }
            offset = offset
                .checked_add(2)
                .ok_or_else(|| corrupted("LPStd offset overflows"))?;
            if cb_std == 0 {
                styles.push(None);
                continue;
            }
            let std_end = offset
                .checked_add(usize::from(cb_std))
                .ok_or_else(|| corrupted("STD range overflows"))?;
            let std = data
                .get(offset..std_end)
                .ok_or_else(|| corrupted("STD is truncated"))?;
            let mut definition = parse_style(std, index, cb_std, stdf_size)?;
            offset = std_end;
            if cb_std % 2 != 0 {
                definition.outer_padding = Some(
                    *data
                        .get(offset)
                        .ok_or_else(|| corrupted("LPStd alignment byte is missing"))?,
                );
                offset += 1;
            }
            styles.push(Some(definition));
        }
        if offset != data.len() {
            return Err(corrupted("stylesheet has trailing bytes"));
        }

        let mut tolerance = ToleranceReport::default();
        validate_styles(&styles, leniency, &mut tolerance)?;
        Ok(Self {
            header,
            styles,
            tolerance,
            stshi_tail: stshi[STSHIF_SIZE..].to_vec(),
        })
    }
}

fn parse_style(std: &[u8], index: u16, cb_std: u16, stdf_size: u16) -> Result<StyleDefinition> {
    let stdf_size = usize::from(stdf_size);
    if std.len() < stdf_size {
        return Err(corrupted("STD is shorter than its Stdf prefix"));
    }
    let info1 = read_u16(std, 0, "StdfBase.info1")?;
    let info2 = read_u16(std, 2, "StdfBase.info2")?;
    let info3 = read_u16(std, 4, "StdfBase.info3")?;
    let bch_upe = read_u16(std, 6, "StdfBase.bchUpe")?;
    if bch_upe != cb_std {
        return Err(corrupted("StdfBase.bchUpe does not match LPStd.cbStd"));
    }
    let grfstd = read_u16(std, 8, "StdfBase.grfstd")?;
    if grfstd & 0xE080 != 0 {
        return Err(corrupted("GRFSTD contains reserved flags"));
    }

    let kind = match info2 & 0x000F {
        1 => StyleKind::Paragraph,
        2 => StyleKind::Character,
        3 => StyleKind::Table,
        4 => StyleKind::Numbering,
        _ => return Err(corrupted("StdfBase.stk is invalid")),
    };
    let base_index = info2 >> 4;
    let base_style = (base_index != NIL_STYLE).then_some(base_index);
    let property_count = info3 & 0x000F;
    let next_style = info3 >> 4;

    let post_2000 = if stdf_size == 18 {
        let post_info1 = read_u16(std, 10, "StdfPost2000.info1")?;
        if post_info1 & 0xE000 != 0 {
            return Err(corrupted("StdfPost2000 contains reserved flags"));
        }
        let post_info3 = read_u16(std, 16, "StdfPost2000.info3")?;
        if post_info3 & 0x0008 != 0 {
            return Err(corrupted("StdfPost2000 contains an invalid unused flag"));
        }
        let priority = post_info3 >> 4;
        if priority > 99 {
            return Err(corrupted("StdfPost2000 priority exceeds 99"));
        }
        let linked = post_info1 & 0x0FFF;
        Some(StylePost2000 {
            linked_style: (linked != 0).then_some(linked),
            has_original_style: post_info1 & 0x1000 != 0,
            revision_id: read_u32(std, 12, "StdfPost2000.rsid")?,
            html_font_category: (post_info3 & 7) as u8,
            priority,
        })
    } else {
        None
    };
    let revision_marked = post_2000
        .as_ref()
        .is_some_and(|post| post.has_original_style);
    let expected_count = match (kind, revision_marked) {
        (StyleKind::Paragraph, false) => 2,
        (StyleKind::Paragraph, true) => 3,
        (StyleKind::Character, false) => 1,
        (StyleKind::Character, true) => 2,
        (StyleKind::Table, false) => 3,
        (StyleKind::Numbering, false) => 1,
        (StyleKind::Table | StyleKind::Numbering, true) => {
            return Err(corrupted(
                "table and numbering styles cannot be revision-marked",
            ));
        },
    };
    if property_count != expected_count {
        return Err(corrupted("StdfBase.cupx does not match the style kind"));
    }

    let name_chars = usize::from(read_u16(std, stdf_size, "Xstz.cch")?);
    let name_start = stdf_size + 2;
    let name_bytes = name_chars
        .checked_mul(2)
        .ok_or_else(|| corrupted("style name length overflows"))?;
    let name_end = name_start
        .checked_add(name_bytes)
        .ok_or_else(|| corrupted("style name range overflows"))?;
    let terminator_end = name_end
        .checked_add(2)
        .ok_or_else(|| corrupted("style name terminator overflows"))?;
    let name_data = std
        .get(name_start..name_end)
        .ok_or_else(|| corrupted("style name is truncated"))?;
    if read_u16(std, name_end, "Xstz terminator")? != 0 {
        return Err(corrupted("style name is not null-terminated"));
    }
    let units = name_data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    let combined_name =
        String::from_utf16(&units).map_err(|_| corrupted("style name contains invalid UTF-16"))?;
    let mut names = combined_name.split(',');
    let name = names.next().unwrap_or_default().to_string();
    let aliases = names.map(str::to_string).collect::<Vec<_>>();
    if name.is_empty() || aliases.iter().any(String::is_empty) {
        return Err(corrupted("style names and aliases must not be empty"));
    }

    let mut offset = terminator_end;
    let mut property_sets = Vec::with_capacity(usize::from(property_count));
    for _ in 0..property_count {
        let size = usize::from(read_u16(std, offset, "LPUpx.cbUpx")?);
        offset = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("LPUpx offset overflows"))?;
        let end = offset
            .checked_add(size)
            .ok_or_else(|| corrupted("UPX range overflows"))?;
        property_sets.push(
            std.get(offset..end)
                .ok_or_else(|| corrupted("UPX is truncated"))?
                .to_vec(),
        );
        offset = end;
        if size % 2 != 0 {
            if std.get(offset).copied() != Some(0) {
                return Err(corrupted("UPX padding must be zero"));
            }
            offset += 1;
        }
    }
    if offset != std.len() {
        return Err(corrupted("STD has trailing bytes"));
    }
    let revision = if revision_marked {
        let revision_index = match kind {
            StyleKind::Paragraph => 2,
            StyleKind::Character => 1,
            StyleKind::Table | StyleKind::Numbering => unreachable!(),
        };
        Some(parse_style_revision(
            &property_sets[revision_index],
            kind,
            index,
        )?)
    } else {
        None
    };

    Ok(StyleDefinition {
        index,
        invariant_id: info1 & 0x0FFF,
        kind,
        base_style,
        next_style,
        name,
        aliases,
        property_sets,
        post_2000,
        revision,
        flags: StyleFlags {
            invalidate_height: info1 & 0x2000 != 0,
            auto_redefine: grfstd & 0x0001 != 0,
            hidden: grfstd & 0x0002 != 0,
            legacy_languages_set: grfstd & 0x0004 != 0,
            copy_language: grfstd & 0x0008 != 0,
            personal_compose: grfstd & 0x0010 != 0,
            personal_reply: grfstd & 0x0020 != 0,
            personal: grfstd & 0x0040 != 0,
            semi_hidden: grfstd & 0x0100 != 0,
            locked: grfstd & 0x0200 != 0,
            unhide_when_used: grfstd & 0x0800 != 0,
            quick_format: grfstd & 0x1000 != 0,
        },
        raw_std: std.to_vec(),
        outer_padding: None,
    })
}

fn parse_style_revision(
    data: &[u8],
    kind: StyleKind,
    style_index: u16,
) -> Result<StyleRevisionMark> {
    if !data.len().is_multiple_of(2) {
        return Err(corrupted(
            "revision-marked style payload is not even-length",
        ));
    }
    if read_u16(data, 0, "LPUpxRm.cbUpx")? != 6 {
        return Err(corrupted("LPUpxRm.cbUpx is not 6"));
    }
    let timestamp = super::super::revision::decode_dttm(read_u32(data, 2, "UpxRm.date")?)?;
    let author_index = read_i16(data, 6, "UpxRm.ibstAuthor")?;
    let mut offset = 8usize;
    let paragraph_properties = if kind == StyleKind::Paragraph {
        let paragraph = read_revision_property_set(data, &mut offset, "LPUpxPapxRM")?;
        let sprms = strip_paragraph_style_index(&paragraph, style_index)?;
        validate_style_sprms(sprms, 1, "UpxPapxRM")?;
        Some(paragraph)
    } else {
        None
    };
    let character_properties = read_revision_property_set(data, &mut offset, "LPUpxChpxRM")?;
    validate_style_sprms(&character_properties, 2, "UpxChpxRM")?;
    if offset != data.len() {
        return Err(corrupted(
            "revision-marked style payload has trailing bytes",
        ));
    }
    Ok(StyleRevisionMark {
        timestamp,
        author_index,
        author: None,
        paragraph_properties,
        character_properties,
    })
}

fn read_revision_property_set(data: &[u8], offset: &mut usize, structure: &str) -> Result<Vec<u8>> {
    let size = usize::from(read_u16(data, *offset, &format!("{structure}.cbUpx"))?);
    *offset = offset
        .checked_add(2)
        .ok_or_else(|| corrupted("revision-marked style offset overflows"))?;
    let end = offset
        .checked_add(size)
        .ok_or_else(|| corrupted("revision-marked style property range overflows"))?;
    let properties = data
        .get(*offset..end)
        .ok_or_else(|| corrupted("revision-marked style property set is truncated"))?
        .to_vec();
    *offset = end;
    if size % 2 != 0 {
        if data.get(*offset).copied() != Some(0) {
            return Err(corrupted(
                "revision-marked style property padding must be zero",
            ));
        }
        *offset += 1;
    }
    Ok(properties)
}

fn strip_paragraph_style_index(properties: &[u8], style_index: u16) -> Result<&[u8]> {
    if properties.len() >= 2 {
        let prefix = read_u16(properties, 0, "UpxPapx.istd")?;
        if prefix == style_index {
            return Ok(&properties[2..]);
        }
    }
    Ok(properties)
}

fn flatten_conditional_style_sprms(
    properties: &[u8],
    conditional_opcode: u16,
    conditions: &[super::tap::TableStyleCondition],
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

fn validate_style_sprms(
    properties: &[u8],
    expected_type: u8,
    structure: &str,
) -> Result<Vec<crate::sprm::Sprm>> {
    let sprms = crate::sprm::parse_sprms(properties)?;
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != properties.len()
        || sprms
            .iter()
            .any(|sprm| get_sprm_type(sprm.opcode) != expected_type)
    {
        return Err(corrupted(&format!(
            "{structure} contains malformed or wrong-type SPRMs"
        )));
    }
    Ok(sprms)
}

pub(crate) fn validate_character_style_sprms(
    properties: &[u8],
    conditional_table_style: bool,
) -> Result<()> {
    let sprms = validate_style_sprms(properties, 2, "UpxChpx")?;
    // [MS-DOC] UpxChpx: explicit exclusions plus every property that sprmCIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x09, 0x0A, 0x0C, 0x11, 0x15, 0x16, 0x17,
        0x18, 0x1A, 0x30, 0x31, 0x33, 0x47, 0x55, 0x56, 0x57, 0x5A, 0x62, 0x63, 0x64, 0x67, 0x6F,
        0x79, 0x82, 0x83, 0x86, 0x87, 0x88, 0x89, 0x90,
    ];
    if let Some(sprm) = sprms.iter().find(|sprm| {
        let operation = get_sprm_operation(sprm.opcode);
        FORBIDDEN.contains(&operation) || (!conditional_table_style && operation == 0x85)
    }) {
        return Err(corrupted(&format!(
            "UpxChpx contains disallowed style SPRM {:#06x}",
            sprm.opcode
        )));
    }
    if sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x85)
    {
        super::chp::CharacterProperties::from_sprm(properties)?;
    }
    Ok(())
}

pub(crate) fn validate_paragraph_style_sprms(
    properties: &[u8],
    style_index: u16,
    conditional_table_style: bool,
) -> Result<()> {
    let properties = strip_paragraph_style_index(properties, style_index)?;
    let sprms = validate_style_sprms(properties, 1, "UpxPapx")?;
    // [MS-DOC] UpxPapx: explicit exclusions plus every property that sprmPIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x00, 0x01, 0x02, 0x10, 0x15, 0x16, 0x17, 0x2C, 0x3F, 0x43, 0x45, 0x46, 0x49, 0x4B, 0x4C,
        0x5A, 0x5F, 0x62, 0x64, 0x65, 0x67, 0x69, 0x6B, 0x6C, 0x6F,
    ];
    if let Some(sprm) = sprms.iter().find(|sprm| {
        let operation = get_sprm_operation(sprm.opcode);
        FORBIDDEN.contains(&operation) || (!conditional_table_style && operation == 0x66)
    }) {
        return Err(corrupted(&format!(
            "UpxPapx contains disallowed style SPRM {:#06x}",
            sprm.opcode
        )));
    }
    if sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x66)
    {
        super::pap::ParagraphProperties::from_sprm(properties)?;
    }
    Ok(())
}

pub(crate) fn validate_numbering_style_sprms(properties: &[u8], style_index: u16) -> Result<()> {
    let properties = strip_paragraph_style_index(properties, style_index)?;
    let sprms = validate_style_sprms(properties, 1, "numbering-style UpxPapx")?;
    if sprms
        .iter()
        .any(|sprm| sprm.opcode != crate::sprm_operations::SPRM_P_ILFO)
    {
        return Err(corrupted(
            "numbering-style UpxPapx contains an SPRM other than sprmPIlfo",
        ));
    }
    Ok(())
}

pub(crate) fn validate_table_style_sprms(
    properties: &[u8],
    style_index: u16,
    inside_conditional: bool,
) -> Result<()> {
    let sprms = crate::sprm::parse_sprms(properties)?;
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != properties.len() || sprms.iter().any(|sprm| get_sprm_type(sprm.opcode) != 5) {
        return Err(corrupted("UpxTapx contains malformed or wrong-type SPRMs"));
    }
    // [MS-DOC] UpxTapx: explicit exclusions plus every property that sprmTIstd preserves.
    const FORBIDDEN: &[u16] = &[
        0x01, 0x02, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F, 0x10, 0x11, 0x12, 0x14,
        0x15, 0x16, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F, 0x20, 0x21, 0x22, 0x23, 0x24,
        0x25, 0x29, 0x2B, 0x2C, 0x2F, 0x32, 0x35, 0x36, 0x39, 0x42, 0x60, 0x62, 0x64, 0x65, 0x67,
        0x68, 0x69, 0x70, 0x71, 0x72, 0x79,
    ];
    let has_conditional = sprms
        .iter()
        .any(|sprm| get_sprm_operation(sprm.opcode) == 0x6A);
    for sprm in sprms {
        let operation = get_sprm_operation(sprm.opcode);
        let conditional_border = (0x7F..=0x84).contains(&operation);
        if FORBIDDEN.contains(&operation) || (conditional_border && !inside_conditional) {
            return Err(corrupted("UpxTapx contains a disallowed style SPRM"));
        }
        if operation == 0x17
            && (inside_conditional || style_index != 11 || sprm.operand_bytes() != [3, 0, 0])
        {
            return Err(corrupted(
                "sprmTWidthBefore is invalid for this table style",
            ));
        }
        if operation == 0x6A {
            if inside_conditional {
                return Err(corrupted("sprmTCnf cannot be nested recursively"));
            }
            let operand = sprm.operand_bytes();
            let nested = operand
                .get(2..)
                .ok_or_else(|| corrupted("sprmTCnf operand is truncated"))?;
            validate_table_style_sprms(nested, style_index, true)?;
        }
    }
    if has_conditional {
        let arena = bumpalo::Bump::new();
        super::tap_parser::TapParser::new(&arena).parse_tap(properties)?;
    }
    Ok(())
}

fn validate_styles(
    styles: &[Option<StyleDefinition>],
    leniency: Leniency,
    tolerance: &mut ToleranceReport,
) -> Result<()> {
    for required_empty in [13usize, 14] {
        if styles.get(required_empty).is_some_and(Option::is_some) {
            return Err(corrupted("reserved fixed-index style is not empty"));
        }
    }
    const FIXED_IDS: [u16; 13] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 65, 105, 107];
    for (index, expected_id) in FIXED_IDS.into_iter().enumerate() {
        if let Some(style) = styles.get(index).and_then(Option::as_ref) {
            if style.invariant_id != expected_id {
                return Err(corrupted("fixed-index style has the wrong invariant ID"));
            }
            let expected_kind = match index {
                0..=9 => StyleKind::Paragraph,
                10 => StyleKind::Character,
                11 => StyleKind::Table,
                12 => StyleKind::Numbering,
                _ => unreachable!(),
            };
            if style.kind != expected_kind {
                return Err(corrupted("fixed-index style has the wrong kind"));
            }
        }
    }

    let mut names = HashSet::new();
    for style in styles.iter().flatten() {
        for name in std::iter::once(&style.name).chain(style.aliases.iter()) {
            if !names.insert(name.as_str()) {
                // MS-DOC 2.9 requires uniqueness, but the name only labels the
                // style: every stored reference resolves by index, so a
                // duplicate cannot make a lookup ambiguous. Rejecting costs the
                // caller the whole document.
                if !leniency.tolerates_stylesheet_defects() {
                    return Err(corrupted("style names and aliases must be unique"));
                }
                tolerance.record(StylesheetDefect::DuplicateStyleName, style.index);
            }
        }
        if let Some(base) = style.base_style
            && (base == style.index || styles.get(usize::from(base)).is_none_or(Option::is_none))
        {
            return Err(corrupted("style has an invalid base style"));
        }
        if styles
            .get(usize::from(style.next_style))
            .is_none_or(Option::is_none)
        {
            return Err(corrupted("style has an invalid next style"));
        }
        if let Some(linked) = style.post_2000.as_ref().and_then(|post| post.linked_style)
            && styles.get(usize::from(linked)).is_none_or(Option::is_none)
        {
            return Err(corrupted("style has an invalid linked style"));
        }
    }

    for style in styles.iter().flatten() {
        let mut visited = HashSet::new();
        let mut current = Some(style.index);
        while let Some(index) = current {
            if !visited.insert(index) {
                return Err(corrupted("style inheritance contains a cycle"));
            }
            current = styles
                .get(usize::from(index))
                .and_then(Option::as_ref)
                .and_then(|definition| definition.base_style);
        }
    }
    Ok(())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(&format!("invalid {field}: {error}")))
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    Ok(read_u16(data, offset, field)? as i16)
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(&format!("invalid {field}: {error}")))
}

fn corrupted(message: &str) -> PackageError {
    PackageError::Corrupted(format!("invalid stylesheet: {message}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(data: &[u8]) -> Result<StyleSheet> {
        StyleSheet::parse_data(data, 0, Leniency::Strict)
    }

    fn std_record(
        invariant_id: u16,
        kind: u16,
        base: u16,
        next: u16,
        name: &str,
        property_sets: &[&[u8]],
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&invariant_id.to_le_bytes());
        data.extend_from_slice(&(kind | (base << 4)).to_le_bytes());
        data.extend_from_slice(&((property_sets.len() as u16) | (next << 4)).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        let units = name.encode_utf16().collect::<Vec<_>>();
        data.extend_from_slice(&(units.len() as u16).to_le_bytes());
        data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        data.extend_from_slice(&0u16.to_le_bytes());
        for property_set in property_sets {
            data.extend_from_slice(&(property_set.len() as u16).to_le_bytes());
            data.extend_from_slice(property_set);
            if property_set.len() % 2 != 0 {
                data.push(0);
            }
        }
        let size = data.len() as u16;
        data[6..8].copy_from_slice(&size.to_le_bytes());
        data
    }

    fn stylesheet(mut slots: Vec<Option<Vec<u8>>>) -> Vec<u8> {
        if slots.len() < 15 {
            slots.resize(15, None);
        }
        let mut data = Vec::new();
        data.extend_from_slice(&18u16.to_le_bytes());
        data.extend_from_slice(&(slots.len() as u16).to_le_bytes());
        data.extend_from_slice(&10u16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&15u16.to_le_bytes());
        data.extend_from_slice(&15u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0i16.to_le_bytes());
        data.extend_from_slice(&0i16.to_le_bytes());
        data.extend_from_slice(&0i16.to_le_bytes());
        for slot in slots {
            if let Some(std) = slot {
                data.extend_from_slice(&(std.len() as u16).to_le_bytes());
                data.extend_from_slice(&std);
                if std.len() % 2 != 0 {
                    data.push(0xA5);
                }
            } else {
                data.extend_from_slice(&0u16.to_le_bytes());
            }
        }
        data
    }

    fn with_post_2000(std: Vec<u8>, info1: u16, revision_id: u32, info3: u16) -> Vec<u8> {
        let mut extended = Vec::with_capacity(std.len() + 8);
        extended.extend_from_slice(&std[..10]);
        extended.extend_from_slice(&info1.to_le_bytes());
        extended.extend_from_slice(&revision_id.to_le_bytes());
        extended.extend_from_slice(&info3.to_le_bytes());
        extended.extend_from_slice(&std[10..]);
        let size = extended.len() as u16;
        extended[6..8].copy_from_slice(&size.to_le_bytes());
        extended
    }

    fn valid_stylesheet() -> Vec<u8> {
        let normal = std_record(0, 1, NIL_STYLE, 0, "Normal,正文", &[&[], &[]]);
        let default_font = std_record(65, 2, NIL_STYLE, 10, "Default Paragraph Font", &[&[]]);
        let mut slots = vec![None; 15];
        slots[0] = Some(normal);
        slots[10] = Some(default_font);
        stylesheet(slots)
    }

    #[test]
    fn parses_styles_and_preserves_raw_upx() {
        let tapx = [
            crate::sprm_operations::SPRM_T_JC.to_le_bytes().as_slice(),
            &1u16.to_le_bytes(),
        ]
        .concat();
        let papx = [
            crate::sprm_operations::SPRM_P_F_KEEP
                .to_le_bytes()
                .as_slice(),
            &[1],
        ]
        .concat();
        let chpx = [
            crate::sprm_operations::SPRM_C_F_BOLD
                .to_le_bytes()
                .as_slice(),
            &[1],
        ]
        .concat();
        let mut slots = vec![None; 15];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Default Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            0,
            0,
            "Table Normal,Grid Alias",
            &[&tapx, &papx, &chpx],
        ));
        let parsed = parse(&stylesheet(slots)).unwrap();
        assert_eq!(parsed.header().style_count, 15);
        assert_eq!(parsed.styles().len(), 15);
        let normal = parsed.get(0).unwrap();
        assert_eq!(normal.name, "Normal");
        assert_eq!(normal.paragraph_properties(), Some([].as_slice()));
        let table = parsed.get(11).unwrap();
        assert_eq!(table.kind, StyleKind::Table);
        assert_eq!(table.base_style, Some(0));
        assert_eq!(table.aliases, ["Grid Alias"]);
        assert_eq!(table.table_properties(), Some(tapx.as_slice()));
        assert_eq!(table.paragraph_properties(), Some(papx.as_slice()));
        assert_eq!(table.character_properties(), Some(chpx.as_slice()));
    }

    #[test]
    fn parses_the_writer_stylesheet() {
        let data = crate::writer::stylesheet::generate_minimal_stylesheet();
        let parsed = parse(&data).unwrap();
        assert_eq!(parsed.styles().len(), 15);
        assert_eq!(parsed.get(0).unwrap().invariant_id, 0);
        assert_eq!(parsed.get(10).unwrap().invariant_id, 65);
        assert!(parsed.get(13).is_none());
        assert!(parsed.get(14).is_none());
    }

    #[test]
    fn parses_post_2000_metadata_and_preserves_stshi_extensions() {
        let revision = [
            6, 0, // LPUpxRm.cbUpx
            0, 0, 0, 0, // UpxRm.date
            0, 0, // UpxRm.ibstAuthor
            0, 0, // LPUpxPapxRM.cbUpx
            0, 0, // LPUpxChpxRM.cbUpx
        ];
        let normal = std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[], &revision]);
        let normal = with_post_2000(normal, 10 | 0x1000, 0x1234_5678, (42 << 4) | 5);
        let default_font = std_record(65, 2, NIL_STYLE, 10, "Default Font", &[&[]]);
        let default_font = with_post_2000(default_font, 0, 0, 0);
        let mut slots = vec![None; 15];
        slots[0] = Some(normal);
        slots[10] = Some(default_font);
        let mut data = stylesheet(slots);
        data[4..6].copy_from_slice(&18u16.to_le_bytes());
        data.splice(20..20, [4, 0, 0xAA, 0x55]);
        data[0..2].copy_from_slice(&22u16.to_le_bytes());

        let parsed = parse(&data).unwrap();
        assert_eq!(parsed.stshi_tail(), [4, 0, 0xAA, 0x55]);
        let post = parsed.get(0).unwrap().post_2000.as_ref().unwrap();
        assert_eq!(post.linked_style, Some(10));
        assert!(post.has_original_style);
        assert_eq!(post.revision_id, 0x1234_5678);
        assert_eq!(post.html_font_category, 5);
        assert_eq!(post.priority, 42);
    }

    #[test]
    fn validates_revision_marked_style_nested_records() {
        use crate::sprm_operations::{SPRM_C_F_BOLD, SPRM_P_F_KEEP};

        let papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        let chpx = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[0]].concat();
        let mut revision = Vec::new();
        revision.extend_from_slice(&6u16.to_le_bytes());
        revision.extend_from_slice(&0u32.to_le_bytes());
        revision.extend_from_slice(&2i16.to_le_bytes());
        revision.extend_from_slice(&(papx.len() as u16).to_le_bytes());
        revision.extend_from_slice(&papx);
        revision.push(0);
        revision.extend_from_slice(&(chpx.len() as u16).to_le_bytes());
        revision.extend_from_slice(&chpx);
        revision.push(0);

        let parsed = parse_style_revision(&revision, StyleKind::Paragraph, 15).unwrap();
        assert_eq!(parsed.author_index, 2);
        assert_eq!(parsed.author, None);
        assert_eq!(parsed.timestamp, None);
        assert_eq!(parsed.paragraph_properties, Some(papx));
        assert_eq!(parsed.character_properties, chpx);

        let mut wrong_rm_size = revision.clone();
        wrong_rm_size[0..2].copy_from_slice(&5u16.to_le_bytes());
        assert!(parse_style_revision(&wrong_rm_size, StyleKind::Paragraph, 15).is_err());

        let mut bad_inner_padding = revision.clone();
        bad_inner_padding[8 + 2 + 3] = 0xA5;
        assert!(parse_style_revision(&bad_inner_padding, StyleKind::Paragraph, 15).is_err());

        let mut trailing = revision;
        trailing.extend_from_slice(&[0, 0]);
        assert!(parse_style_revision(&trailing, StyleKind::Paragraph, 15).is_err());
    }

    #[test]
    fn enforces_kind_specific_style_sprm_restrictions() {
        use crate::sprm_operations::{
            SPRM_C_CNF, SPRM_C_F_BOLD, SPRM_C_ISTD, SPRM_P_F_IN_TABLE, SPRM_P_F_KEEP, SPRM_P_ILFO,
        };

        let bold = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        assert!(validate_character_style_sprms(&bold, false).is_ok());
        let character_style = [SPRM_C_ISTD.to_le_bytes().as_slice(), &15u16.to_le_bytes()].concat();
        assert!(validate_character_style_sprms(&character_style, false).is_err());
        let conditional_character = [
            SPRM_C_CNF.to_le_bytes().as_slice(),
            &[2],
            &1u16.to_le_bytes(),
        ]
        .concat();
        assert!(validate_character_style_sprms(&conditional_character, false).is_err());
        assert!(validate_character_style_sprms(&conditional_character, true).is_ok());

        let keep = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        assert!(validate_paragraph_style_sprms(&keep, 15, false).is_ok());
        let table_state = [SPRM_P_F_IN_TABLE.to_le_bytes().as_slice(), &[1]].concat();
        assert!(validate_paragraph_style_sprms(&table_state, 15, true).is_err());
        let list = [SPRM_P_ILFO.to_le_bytes().as_slice(), &1u16.to_le_bytes()].concat();
        assert!(validate_numbering_style_sprms(&list, 15).is_ok());
        assert!(validate_numbering_style_sprms(&keep, 15).is_err());

        let forbidden_table_position =
            [0x9601u16.to_le_bytes().as_slice(), &0i16.to_le_bytes()].concat();
        assert!(validate_table_style_sprms(&forbidden_table_position, 15, false).is_err());
        let width_before = [0xF617u16.to_le_bytes().as_slice(), &[3, 0, 0]].concat();
        assert!(validate_table_style_sprms(&width_before, 15, false).is_err());
        assert!(validate_table_style_sprms(&width_before, 11, false).is_ok());

        let border = [0xD47Fu16.to_le_bytes().as_slice(), &[8], &[0; 8]].concat();
        assert!(validate_table_style_sprms(&border, 15, false).is_err());
        let conditional_table = [
            0xD66Au16.to_le_bytes().as_slice(),
            &[(border.len() + 2) as u8],
            &1u16.to_le_bytes(),
            border.as_slice(),
        ]
        .concat();
        assert!(validate_table_style_sprms(&conditional_table, 15, false).is_ok());
        assert!(validate_table_style_sprms(&conditional_table, 15, true).is_err());
    }

    #[test]
    fn rejects_malformed_record_framing() {
        let valid = valid_stylesheet();
        assert!(parse(&valid[..valid.len() - 1]).is_err());

        let mut short_header = valid.clone();
        short_header[0..2].copy_from_slice(&16u16.to_le_bytes());
        assert!(parse(&short_header).is_err());

        let mut negative_std = valid.clone();
        negative_std[20..22].copy_from_slice(&0x8000u16.to_le_bytes());
        assert!(parse(&negative_std).is_err());

        let mut wrong_bch = valid.clone();
        wrong_bch[28..30].copy_from_slice(&0u16.to_le_bytes());
        assert!(parse(&wrong_bch).is_err());

        let mut trailing = valid;
        trailing.push(0);
        assert!(parse(&trailing).is_err());
    }

    #[test]
    fn rejects_invalid_semantics_and_padding() {
        let mut slots = vec![None; 15];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Same", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Same", &[&[]]));
        assert!(parse(&stylesheet(slots)).is_err());

        let mut slots = vec![None; 15];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[11] = Some(std_record(105, 3, 11, 0, "Table", &[&[], &[], &[]]));
        assert!(parse(&stylesheet(slots)).is_err());

        let mut slots = vec![None; 15];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        let mut table = std_record(105, 3, 0, 0, "Table", &[&[1], &[], &[]]);
        let padding = table
            .windows(4)
            .rposition(|bytes| bytes == [1, 0, 1, 0])
            .unwrap()
            + 3;
        table[padding] = 1;
        slots[11] = Some(table);
        assert!(parse(&stylesheet(slots)).is_err());
    }

    #[test]
    fn resolves_table_style_inheritance_and_default_fallback() {
        let mut slots = vec![None; 17];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            NIL_STYLE,
            0,
            "Normal Table",
            &[&[0x00, 0x54, 0x01, 0x00], &[], &[]],
        ));
        slots[15] = Some(std_record(
            0x0FFE,
            3,
            11,
            0,
            "Base Table",
            &[&[0x7D, 0x34, 0x01], &[], &[]],
        ));
        slots[16] = Some(std_record(
            0x0FFE,
            3,
            15,
            0,
            "Derived Table",
            &[&[0x00, 0x54, 0x02, 0x00], &[], &[]],
        ));
        let parsed = parse(&stylesheet(slots)).unwrap();

        let (effective, properties) = parsed.resolve_table_properties(16).unwrap();
        assert_eq!(effective, 16);
        assert_eq!(
            properties.justification,
            super::super::tap::TableJustification::Right
        );
        assert_eq!(properties.style_defaults.no_wrap, Some(true));

        let (effective, fallback) = parsed.resolve_table_properties(999).unwrap();
        assert_eq!(effective, 11);
        assert_eq!(
            fallback.justification,
            super::super::tap::TableJustification::Center
        );
    }

    #[test]
    fn resolves_table_text_style_inheritance_conditions_and_fallback() {
        use crate::sprm_operations::{
            SPRM_C_CNF, SPRM_C_F_BOLD, SPRM_C_F_ITALIC, SPRM_P_CNF, SPRM_P_F_KEEP,
            SPRM_P_F_KEEP_FOLLOW,
        };

        fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.extend_from_slice(operand);
        }

        fn conditional(opcode: u16, condition: u16, nested: &[u8]) -> Vec<u8> {
            let mut grpprl = opcode.to_le_bytes().to_vec();
            grpprl.push((nested.len() + 2) as u8);
            grpprl.extend_from_slice(&condition.to_le_bytes());
            grpprl.extend_from_slice(nested);
            grpprl
        }

        let mut normal_papx = Vec::new();
        append(&mut normal_papx, SPRM_P_F_KEEP, &[1]);
        let normal_conditional = [SPRM_P_F_KEEP_FOLLOW.to_le_bytes().as_slice(), &[1]].concat();
        normal_papx.extend_from_slice(&conditional(SPRM_P_CNF, 0x0001, &normal_conditional));
        let mut normal_chpx = Vec::new();
        append(&mut normal_chpx, SPRM_C_F_BOLD, &[1]);
        let normal_character_conditional =
            [SPRM_C_F_ITALIC.to_le_bytes().as_slice(), &[1]].concat();
        normal_chpx.extend_from_slice(&conditional(
            SPRM_C_CNF,
            0x0001,
            &normal_character_conditional,
        ));

        let mut derived_papx = 15u16.to_le_bytes().to_vec();
        append(&mut derived_papx, SPRM_P_F_KEEP, &[0]);
        let derived_conditional = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        derived_papx.extend_from_slice(&conditional(SPRM_P_CNF, 0x0008, &derived_conditional));
        let mut derived_chpx = Vec::new();
        append(&mut derived_chpx, SPRM_C_F_BOLD, &[0]);
        let derived_character_conditional = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[1]].concat();
        derived_chpx.extend_from_slice(&conditional(
            SPRM_C_CNF,
            0x0008,
            &derived_character_conditional,
        ));

        let mut slots = vec![None; 16];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            NIL_STYLE,
            0,
            "Normal Table",
            &[&[], &normal_papx, &normal_chpx],
        ));
        slots[15] = Some(std_record(
            0x0FFE,
            3,
            11,
            0,
            "Derived Table",
            &[&[], &derived_papx, &derived_chpx],
        ));
        let stylesheet = parse(&stylesheet(slots)).unwrap();

        let (effective, paragraph, character) =
            stylesheet.resolve_table_text_properties(15).unwrap();
        assert_eq!(effective, 15);
        assert!(!paragraph.keep_on_page);
        assert_eq!(paragraph.conditional_formats.len(), 2);
        assert_eq!(
            paragraph.conditional_formats[0].condition,
            super::super::tap::TableStyleCondition::HeaderRow
        );
        assert!(paragraph.conditional_formats[0].properties.keep_with_next);
        assert_eq!(
            paragraph.conditional_formats[1].condition,
            super::super::tap::TableStyleCondition::LastColumn
        );
        assert!(paragraph.conditional_formats[1].properties.keep_on_page);
        assert_eq!(character.is_bold, Some(false));
        assert_eq!(character.conditional_formats.len(), 2);
        assert_eq!(
            character.conditional_formats[0].condition,
            super::super::tap::TableStyleCondition::HeaderRow
        );
        assert_eq!(
            character.conditional_formats[0].properties.is_italic,
            Some(true)
        );
        assert_eq!(
            character.conditional_formats[1].condition,
            super::super::tap::TableStyleCondition::LastColumn
        );
        assert_eq!(
            character.conditional_formats[1].properties.is_bold,
            Some(true)
        );

        let (_, table_papx, _) = stylesheet.resolve_table_text_style_sprms(15).unwrap();
        let direct_papx = [SPRM_P_F_KEEP.to_le_bytes().as_slice(), &[1]].concat();
        let cascaded = super::super::pap::ParagraphProperties::cascade_table_style(
            &table_papx,
            Some(0),
            &direct_papx,
            &stylesheet,
        )
        .unwrap();
        assert!(cascaded.keep_on_page);
        assert_eq!(cascaded.style_index, Some(0));
        assert_eq!(cascaded.conditional_formats.len(), 2);

        let (effective, paragraph, character) =
            stylesheet.resolve_table_text_properties(999).unwrap();
        assert_eq!(effective, 11);
        assert!(paragraph.keep_on_page);
        assert_eq!(paragraph.conditional_formats.len(), 1);
        assert_eq!(character.is_bold, Some(true));
        assert_eq!(character.conditional_formats.len(), 1);
    }

    #[test]
    fn rejects_malformed_table_text_style_property_sets_when_resolved() {
        let mut slots = vec![None; 16];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            NIL_STYLE,
            0,
            "Normal Table",
            &[&[], &[], &[]],
        ));
        slots[15] = Some(std_record(
            0x0FFE,
            3,
            11,
            0,
            "Malformed Table",
            &[&[], &[0x35, 0x08, 1], &[]],
        ));
        let stylesheet = parse(&stylesheet(slots)).unwrap();
        assert!(stylesheet.resolve_table_text_properties(15).is_err());
    }

    #[test]
    fn flattens_table_text_conditions_in_position_precedence_order() {
        use crate::sprm_operations::{SPRM_C_CNF, SPRM_C_F_BOLD};

        fn conditional(condition: u16, value: u8) -> Vec<u8> {
            let nested = [SPRM_C_F_BOLD.to_le_bytes().as_slice(), &[value]].concat();
            let mut grpprl = SPRM_C_CNF.to_le_bytes().to_vec();
            grpprl.push((nested.len() + 2) as u8);
            grpprl.extend_from_slice(&condition.to_le_bytes());
            grpprl.extend_from_slice(&nested);
            grpprl
        }

        // Source order is deliberately the reverse of positional precedence.
        let source = [conditional(0x0001, 0), conditional(0x0040, 1)].concat();
        let flattened = flatten_conditional_style_sprms(
            &source,
            SPRM_C_CNF,
            &[
                super::super::tap::TableStyleCondition::OddRowBand,
                super::super::tap::TableStyleCondition::HeaderRow,
            ],
        )
        .unwrap();
        let properties = super::super::chp::CharacterProperties::from_sprm(&flattened).unwrap();
        assert_eq!(properties.is_bold, Some(false));
        assert!(properties.conditional_formats.is_empty());
    }

    #[test]
    fn applies_table_styles_in_direct_sprm_order_and_preserves_sizing() {
        fn append(grpprl: &mut Vec<u8>, opcode: u16, operand: &[u8]) {
            grpprl.extend_from_slice(&opcode.to_le_bytes());
            grpprl.extend_from_slice(operand);
        }

        let mut slots = vec![None; 16];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            NIL_STYLE,
            0,
            "Normal Table",
            &[&[], &[], &[]],
        ));
        let mut style_tapx = Vec::new();
        append(&mut style_tapx, 0x548A, &[2, 0]);
        append(&mut style_tapx, 0x3404, &[1]);
        append(&mut style_tapx, 0x3403, &[1]);
        append(&mut style_tapx, 0x347D, &[1]);
        slots[15] = Some(std_record(
            0x0FFE,
            3,
            11,
            0,
            "Applied Table",
            &[&style_tapx, &[], &[]],
        ));
        let stylesheet = parse(&stylesheet(slots)).unwrap();

        let mut direct = Vec::new();
        append(&mut direct, 0x548A, &[1, 0]);
        append(&mut direct, 0x3404, &[0]);
        append(&mut direct, 0x3403, &[0]);
        append(&mut direct, 0x9407, &[0x20, 0x03]);
        append(&mut direct, 0xF614, &[3, 0xE8, 0x03]);
        append(&mut direct, 0x3615, &[1]);
        append(&mut direct, 0x5664, &[1, 0]);
        append(&mut direct, 0x7479, &[0x78, 0x56, 0x34, 0x12]);
        append(&mut direct, 0x563A, &[15, 0]);

        let arena = bumpalo::Bump::new();
        let parser = super::super::tap_parser::TapParser::new(&arena);
        let styled = parser
            .parse_tap_with_stylesheet(&direct, &stylesheet)
            .unwrap();
        assert_eq!(styled.table_style_index, Some(15));
        assert_eq!(
            styled.justification,
            super::super::tap::TableJustification::Right
        );
        assert!(styled.is_header_row);
        assert!(!styled.allow_row_break);
        assert_eq!(styled.style_defaults.no_wrap, Some(true));
        assert_eq!(styled.row_height, Some(800));
        assert_eq!(styled.preferred_width.unwrap().value, 1000);
        assert!(styled.auto_fit);
        assert!(styled.right_to_left);
        assert_eq!(styled.revision_save_id, Some(0x1234_5678));

        append(&mut direct, 0x548A, &[0, 0]);
        append(&mut direct, 0x3404, &[0]);
        append(&mut direct, 0x3403, &[0]);
        let overridden = parser
            .parse_tap_with_stylesheet(&direct, &stylesheet)
            .unwrap();
        assert_eq!(
            overridden.justification,
            super::super::tap::TableJustification::Left
        );
        assert!(!overridden.is_header_row);
        assert!(overridden.allow_row_break);
        assert_eq!(overridden.row_height, Some(800));
        assert_eq!(overridden.preferred_width.unwrap().value, 1000);
        assert!(overridden.auto_fit);
        assert!(overridden.right_to_left);
        assert_eq!(overridden.revision_save_id, Some(0x1234_5678));
    }

    #[test]
    fn resolves_paragraph_and_character_style_property_arrays() {
        let mut slots = vec![None; 19];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Font", &[&[]]));
        slots[15] = Some(std_record(
            0x0FFE,
            1,
            0,
            0,
            "Base Paragraph",
            &[&[15, 0, 0x03, 0x24, 2], &[0x35, 0x08, 1]],
        ));
        slots[16] = Some(std_record(
            0x0FFE,
            1,
            15,
            0,
            "Derived Paragraph",
            &[&[0x03, 0x24, 1], &[0x36, 0x08, 1]],
        ));
        slots[17] = Some(std_record(
            0x0FFE,
            2,
            10,
            17,
            "Base Character",
            &[&[0x35, 0x08, 1]],
        ));
        slots[18] = Some(std_record(
            0x0FFE,
            2,
            17,
            18,
            "Derived Character",
            &[&[0x36, 0x08, 1]],
        ));
        let stylesheet = parse(&stylesheet(slots)).unwrap();

        let (effective, paragraph, character) =
            stylesheet.resolve_paragraph_style_sprms(16).unwrap();
        assert_eq!(effective, Some(16));
        assert_eq!(paragraph, [0x03, 0x24, 2, 0x03, 0x24, 1]);
        assert_eq!(character, [0x35, 0x08, 1, 0x36, 0x08, 1]);
        let styled = super::super::pap::ParagraphProperties::from_sprm(&paragraph).unwrap();
        assert_eq!(
            styled.justification,
            super::super::pap::Justification::Center
        );
        let mut direct = paragraph.clone();
        direct.extend_from_slice(&[0x03, 0x24, 0]);
        let overridden = super::super::pap::ParagraphProperties::from_sprm(&direct).unwrap();
        assert_eq!(
            overridden.justification,
            super::super::pap::Justification::Left
        );

        let (effective, character) = stylesheet.resolve_character_style_sprms(18).unwrap();
        assert_eq!(effective, Some(18));
        assert_eq!(character, [0x35, 0x08, 1, 0x36, 0x08, 1]);

        let direct_grpprl = [0x30, 0x4A, 18, 0, 0x35, 0x08, 0];
        let direct = super::super::chp::CharacterProperties::from_sprm(&direct_grpprl).unwrap();
        let cascaded = super::super::paragraph_extractor::cascade_character_properties(
            Some(&stylesheet),
            &[0x35, 0x08, 1, 0x36, 0x08, 0],
            &direct,
            &direct_grpprl,
        )
        .unwrap();
        assert_eq!(cascaded.style_index, Some(18));
        assert_eq!(cascaded.is_bold, Some(false));
        assert_eq!(cascaded.is_italic, Some(true));

        let ordered_grpprl = [
            0x35, 0x08, 0, // direct bold before the style: reset by sprmCIstd
            0x0C, 0x2A, 7, // highlight: explicitly preserved by sprmCIstd
            0x55, 0x08, 1, // fSpec: explicitly preserved by sprmCIstd
            0x30, 0x4A, 18, 0, // derived character style
            0x36, 0x08, 0, // direct italic after the style: authoritative
        ];
        let ordered_direct =
            super::super::chp::CharacterProperties::from_sprm(&ordered_grpprl).unwrap();
        let ordered = super::super::paragraph_extractor::cascade_character_properties(
            Some(&stylesheet),
            &[0x35, 0x08, 1],
            &ordered_direct,
            &ordered_grpprl,
        )
        .unwrap();
        assert_eq!(ordered.style_index, Some(18));
        assert_eq!(ordered.is_bold, Some(true));
        assert_eq!(ordered.is_italic, Some(false));
        assert!(ordered.is_spec);
        assert_eq!(
            ordered.highlight,
            Some(super::super::chp::HighlightColor::Yellow)
        );

        assert_eq!(
            stylesheet.resolve_paragraph_style_sprms(18).unwrap(),
            (None, Vec::new(), Vec::new())
        );
        assert_eq!(
            stylesheet.resolve_character_style_sprms(16).unwrap(),
            (None, Vec::new())
        );

        let mut ordered_papx = vec![
            0x03, 0x24, 0, // direct left alignment before style switch
            0x16, 0x24, 1, // table membership is preserved
            0x5A, 0x24, 1, // open cell-mark display state is preserved
            0x64, 0x26, 1, // paragraph revision wall is preserved
            0x65, 0x64, 7, 0, 0, 0, // PGPInfo identity is preserved
            0x67, 0x64, 0x78, 0x56, 0x34, 0x12, // paragraph RSID is preserved
            0x00, 0x46, 15, 0, // switch back to Base Paragraph (right)
        ];
        let switched = super::super::pap::ParagraphProperties::cascade_styles(
            Some(16),
            &ordered_papx,
            &stylesheet,
        )
        .unwrap();
        assert_eq!(switched.style_index, Some(15));
        assert_eq!(
            switched.justification,
            super::super::pap::Justification::Right
        );
        assert!(switched.in_table);
        assert!(switched.open_table_cell_mark);
        assert!(switched.properties_preserved_for_revision);
        assert_eq!(switched.paragraph_group_id, Some(7));
        assert_eq!(switched.revision_save_id, Some(0x1234_5678));

        ordered_papx.extend_from_slice(&[0x03, 0x24, 1]);
        let overridden = super::super::pap::ParagraphProperties::cascade_styles(
            Some(16),
            &ordered_papx,
            &stylesheet,
        )
        .unwrap();
        assert_eq!(
            overridden.justification,
            super::super::pap::Justification::Center
        );

        let permuted = super::super::pap::ParagraphProperties::cascade_styles(
            Some(16),
            &[
                0x01, 0xC6, 7, // sprmPIstdPermute and SPPOperand length
                0, 16, 0, 16, 0, 15, 0, // style 16 maps to style 15
            ],
            &stylesheet,
        )
        .unwrap();
        assert_eq!(permuted.style_index, Some(15));
        assert_eq!(
            permuted.justification,
            super::super::pap::Justification::Right
        );
    }
}

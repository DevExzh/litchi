//! Word 97+ stylesheet (`STSH`) parsing.

use std::collections::HashSet;

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use super::tap::TableProperties;

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
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
    header: StyleSheetHeader,
    styles: Vec<Option<StyleDefinition>>,
    stshi_tail: Vec<u8>,
}

impl StyleSheet {
    /// Parse the mandatory Word 97+ stylesheet at FIB pointer index 1.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let (stream_offset, length) = fib
            .get_table_pointer(STSH_POINTER_INDEX)
            .filter(|(_, length)| *length != 0)
            .ok_or_else(|| DocError::Corrupted("FIB does not contain a stylesheet".to_string()))?;
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
        Self::parse_data(data, start)
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

    fn parse_data(data: &[u8], stream_offset: usize) -> Result<Self> {
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

        validate_styles(&styles)?;
        Ok(Self {
            header,
            styles,
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

fn strip_paragraph_style_index(properties: &[u8], style_index: u16) -> Result<&[u8]> {
    if properties.len() >= 2 {
        let prefix = read_u16(properties, 0, "UpxPapx.istd")?;
        if prefix == style_index {
            return Ok(&properties[2..]);
        }
    }
    Ok(properties)
}

fn validate_style_sprms(properties: &[u8], expected_type: u8, structure: &str) -> Result<()> {
    let sprms = crate::sprm::parse_sprms(properties);
    let consumed = sprms.last().map_or(0, |sprm| sprm.offset + sprm.size);
    if consumed != properties.len()
        || sprms
            .iter()
            .any(|sprm| crate::sprm_operations::get_sprm_type(sprm.opcode) != expected_type)
    {
        return Err(corrupted(&format!(
            "{structure} contains malformed or wrong-type SPRMs"
        )));
    }
    Ok(())
}

fn validate_styles(styles: &[Option<StyleDefinition>]) -> Result<()> {
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
                return Err(corrupted("style names and aliases must be unique"));
            }
        }
        if let Some(base) = style.base_style {
            if base == style.index || styles.get(usize::from(base)).is_none_or(Option::is_none) {
                return Err(corrupted("style has an invalid base style"));
            }
        }
        if styles
            .get(usize::from(style.next_style))
            .is_none_or(Option::is_none)
        {
            return Err(corrupted("style has an invalid next style"));
        }
        if let Some(linked) = style.post_2000.as_ref().and_then(|post| post.linked_style) {
            if styles.get(usize::from(linked)).is_none_or(Option::is_none) {
                return Err(corrupted("style has an invalid linked style"));
            }
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

fn corrupted(message: &str) -> DocError {
    DocError::Corrupted(format!("invalid stylesheet: {message}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(data: &[u8]) -> Result<StyleSheet> {
        StyleSheet::parse_data(data, 0)
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
        let mut slots = vec![None; 15];
        slots[0] = Some(std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[], &[]]));
        slots[10] = Some(std_record(65, 2, NIL_STYLE, 10, "Default Font", &[&[]]));
        slots[11] = Some(std_record(
            105,
            3,
            0,
            0,
            "Table Normal,Grid Alias",
            &[&[0x11, 0x22, 0x33], &[0x44, 0x55], &[0x66]],
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
        assert_eq!(
            table.table_properties(),
            Some([0x11, 0x22, 0x33].as_slice())
        );
        assert_eq!(table.paragraph_properties(), Some([0x44, 0x55].as_slice()));
        assert_eq!(table.character_properties(), Some([0x66].as_slice()));
    }

    #[test]
    fn parses_the_writer_stylesheet() {
        let data = crate::doc::writer::stylesheet::generate_minimal_stylesheet();
        let parsed = parse(&data).unwrap();
        assert_eq!(parsed.styles().len(), 15);
        assert_eq!(parsed.get(0).unwrap().invariant_id, 0);
        assert_eq!(parsed.get(10).unwrap().invariant_id, 65);
        assert!(parsed.get(13).is_none());
        assert!(parsed.get(14).is_none());
    }

    #[test]
    fn parses_post_2000_metadata_and_preserves_stshi_extensions() {
        let normal = std_record(0, 1, NIL_STYLE, 0, "Normal", &[&[1], &[2], &[3]]);
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

        assert_eq!(
            stylesheet.resolve_paragraph_style_sprms(18).unwrap(),
            (None, Vec::new(), Vec::new())
        );
        assert_eq!(
            stylesheet.resolve_character_style_sprms(16).unwrap(),
            (None, Vec::new())
        );
    }
}

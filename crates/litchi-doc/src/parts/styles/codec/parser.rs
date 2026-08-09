//! Binary STSH/STD/UPX parser.

use super::super::model::{
    StyleDefinition, StyleFlags, StyleKind, StylePost2000, StyleRevisionMark, StyleSheet,
    StyleSheetHeader,
};
use super::{corrupted, read_i16, read_u16, read_u32};
use crate::leniency::{Leniency, ToleranceReport};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use crate::parts::styles::semantic::{
    strip_paragraph_style_index, validate_style_sprms, validate_styles,
};

const STSH_POINTER_INDEX: usize = 1;
const STSHIF_SIZE: usize = 18;
const MIN_STYLE_COUNT: u16 = 0x000F;
const MAX_STYLE_COUNT: u16 = 0x0FFD;
pub(in crate::parts::styles) const NIL_STYLE: u16 = 0x0FFF;

impl StyleSheet {
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
}

impl StyleSheet {
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

pub(in crate::parts::styles) fn parse_style_revision(
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
    let timestamp = crate::revision::decode_dttm(read_u32(data, 2, "UpxRm.date")?)?;
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

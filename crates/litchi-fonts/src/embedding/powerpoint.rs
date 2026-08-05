//! PowerPoint EOT publication.

use allsorts::{
    binary::read::ReadScope,
    tables::{
        FontTableProvider, HeadTable, NameTable, OpenTypeData, OpenTypeFont,
        os2::{FsSelectionFlag, Os2},
    },
};

use super::Prepared;
use crate::FontError;

type Result<T> = std::result::Result<T, FontError>;

fn invalid(message: impl Into<String>) -> FontError {
    FontError::EmbeddingFailed(message.into())
}

/// Wrap one selected, uncompressed OpenType face in the EOT container used by
/// PowerPoint's `application/x-fontdata` parts.
///
/// This deliberately emits EOT 1.0 without proprietary MicroType compression
/// or XOR processing. All integers in the wrapper are little-endian; the
/// adopted OpenType program remains byte-for-byte unchanged.
pub fn data(font: &mut Prepared) -> Result<Vec<u8>> {
    const EOT_VERSION_1: u32 = 0x0001_0000;
    const EOT_MAGIC: u16 = 0x504C;
    const EOT_FIXED_BYTES: usize = 82;
    const SUBSET: u32 = 0x0000_0001;

    let scope = ReadScope::new(&font.data);
    let file = scope
        .read::<OpenTypeFont<'_>>()
        .map_err(|error| invalid(format!("invalid OpenType font: {error}")))?;
    if !matches!(&file.data, OpenTypeData::Single(_)) {
        return Err(invalid(
            "PowerPoint EOT authoring currently requires one standalone OpenType face",
        ));
    }
    let provider = file
        .table_provider(0)
        .map_err(|error| invalid(format!("invalid standalone OpenType face: {error}")))?;

    let os2_data = provider
        .table_data(allsorts::tag::OS_2)
        .map_err(|error| invalid(format!("invalid OS/2 table: {error}")))?
        .ok_or_else(|| invalid("OpenType font has no OS/2 table"))?;
    let os2_bytes: &[u8] = os2_data.as_ref();
    let os2 = ReadScope::new(os2_bytes)
        .read_dep::<Os2>(os2_bytes.len())
        .map_err(|error| invalid(format!("invalid OS/2 table: {error}")))?;
    let license = font.properties.license();
    super::validate_license(&font.name, license)?;

    let head_data = provider
        .table_data(allsorts::tag::HEAD)
        .map_err(|error| invalid(format!("invalid head table: {error}")))?
        .ok_or_else(|| invalid("OpenType font has no head table"))?;
    let head = ReadScope::new(head_data.as_ref())
        .read::<HeadTable>()
        .map_err(|error| invalid(format!("invalid head table: {error}")))?;

    let name_data = provider
        .table_data(allsorts::tag::NAME)
        .map_err(|error| invalid(format!("invalid name table: {error}")))?
        .ok_or_else(|| invalid("OpenType font has no name table"))?;
    let names = ReadScope::new(name_data.as_ref())
        .read::<NameTable<'_>>()
        .map_err(|error| invalid(format!("invalid name table: {error}")))?;
    let family = names
        .string_for_id(NameTable::FONT_FAMILY_NAME)
        .unwrap_or_else(|| font.name.clone());
    let style = names
        .string_for_id(NameTable::FONT_SUBFAMILY_NAME)
        .unwrap_or_else(|| "Regular".into());
    let version = names
        .string_for_id(NameTable::VERSION_STRING)
        .unwrap_or_default();
    let full = names
        .string_for_id(NameTable::FULL_FONT_NAME)
        .unwrap_or_else(|| font.name.clone());
    let names = [&family, &style, &version, &full];

    let name_bytes = names
        .iter()
        .enumerate()
        .try_fold(0usize, |total, (index, value)| {
            let bytes = utf16_bytes(value)?;
            // EOT 1.0 has a two-byte size before every name and padding after
            // Family, Style, and Version only. FontData follows FullName directly.
            let overhead = if index + 1 == names.len() { 2 } else { 4 };
            total
                .checked_add(overhead)
                .and_then(|total| total.checked_add(bytes))
                .ok_or_else(|| invalid("EOT name data is too large"))
        })?;
    let eot_size = EOT_FIXED_BYTES
        .checked_add(name_bytes)
        .and_then(|size| size.checked_add(font.data.len()))
        .ok_or_else(|| invalid("EOT payload size overflow"))?;
    let eot_size =
        u32::try_from(eot_size).map_err(|_| invalid("EOT payload exceeds the 32-bit format"))?;
    let font_size = u32::try_from(font.data.len())
        .map_err(|_| invalid("OpenType payload exceeds the 32-bit EOT format"))?;

    let allocation_size =
        usize::try_from(eot_size).map_err(|_| invalid("EOT size does not fit this platform"))?;
    let header_size = allocation_size
        .checked_sub(font.data.len())
        .ok_or_else(|| invalid("EOT header size underflow"))?;
    let mut header = Vec::new();
    header
        .try_reserve_exact(header_size)
        .map_err(|source| FontError::Allocation {
            resource: "PowerPoint EOT header",
            source,
        })?;
    push_u32(&mut header, eot_size);
    push_u32(&mut header, font_size);
    push_u32(&mut header, EOT_VERSION_1);
    push_u32(&mut header, if font.subsetted { SUBSET } else { 0 });
    header.extend_from_slice(&os2.panose);
    header.push(font.properties.charset().map_or(1, crate::Charset::code));
    header.push(u8::from(os2.fs_selection.contains(FsSelectionFlag::ITALIC)));
    push_u32(&mut header, u32::from(os2.us_weight_class));
    push_u16(&mut header, license.bits());
    push_u16(&mut header, EOT_MAGIC);
    for range in [
        os2.ul_unicode_range1,
        os2.ul_unicode_range2,
        os2.ul_unicode_range3,
        os2.ul_unicode_range4,
    ] {
        push_u32(&mut header, range);
    }
    let (code_page1, code_page2) = os2.version1.as_ref().map_or((0, 0), |value| {
        (value.ul_code_page_range1, value.ul_code_page_range2)
    });
    push_u32(&mut header, code_page1);
    push_u32(&mut header, code_page2);
    push_u32(&mut header, head.check_sum_adjustment);
    for _ in 0..4 {
        push_u32(&mut header, 0);
    }
    push_u16(&mut header, 0);
    for (index, value) in names.into_iter().enumerate() {
        push_utf16(&mut header, value)?;
        if index + 1 != names.len() {
            push_u16(&mut header, 0);
        }
    }
    if header.len() != header_size {
        return Err(invalid("constructed EOT header has an inconsistent size"));
    }

    // EOT requires one contiguous buffer. Reuse the source Vec, shift its font
    // bytes once in place, and copy only the bounded header into the prefix.
    font.data
        .try_reserve_exact(header_size)
        .map_err(|source| FontError::Allocation {
            resource: "PowerPoint EOT payload",
            source,
        })?;
    let mut output = std::mem::take(&mut font.data);
    let source_size = output.len();
    output.resize(allocation_size, 0);
    output.copy_within(0..source_size, header_size);
    output[..header_size].copy_from_slice(&header);
    if output.len() != allocation_size {
        return Err(invalid("constructed EOT payload has an inconsistent size"));
    }
    Ok(output)
}

fn utf16_bytes(value: &str) -> Result<usize> {
    let units = value.encode_utf16().count();
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| invalid("EOT name length overflow"))?;
    u16::try_from(bytes)
        .map(|_| bytes)
        .map_err(|_| invalid("EOT name exceeds 65535 bytes"))
}

fn push_utf16(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let bytes = utf16_bytes(value)?;
    push_u16(
        output,
        u16::try_from(bytes).map_err(|_| invalid("EOT name exceeds 65535 bytes"))?,
    );
    for unit in value.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn creates_uncompressed_eot_by_moving_one_standalone_face() {
        let sfnt = test_sfnt();
        let signature = crate::Signature::new([1, 2, 3, 4], [5, 6]);
        let properties = crate::FontProperties::new(
            crate::License::new(0).unwrap(),
            crate::Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
            Some(crate::Charset::ANSI),
            crate::Family::Roman,
            crate::Pitch::Variable,
            signature,
        );
        let mut font = Prepared {
            name: "Litchi Test".into(),
            style: crate::Style::Regular,
            data: sfnt.clone(),
            properties,
            subsetted: true,
        };

        let eot = data(&mut font).unwrap();
        assert!(font.data.is_empty());
        assert_eq!(u32_at(&eot, 0), eot.len() as u32);
        assert_eq!(u32_at(&eot, 4), sfnt.len() as u32);
        assert_eq!(u32_at(&eot, 8), 0x0001_0000);
        assert_eq!(u32_at(&eot, 12), 1);
        assert_eq!(&eot[34..36], &0x504Cu16.to_le_bytes());
        assert_eq!(eot_v1_font_offset(&eot), eot.len() - sfnt.len());
        assert!(eot.ends_with(&sfnt));
    }

    fn test_sfnt() -> Vec<u8> {
        let mut os2 = vec![0; 96];
        set_u16(&mut os2, 0, 2);
        set_u16(&mut os2, 4, 400);
        set_u16(&mut os2, 6, 5);
        os2[32..42].copy_from_slice(&[2, 11, 6, 4, 2, 2, 2, 2, 2, 4]);
        set_u32(&mut os2, 42, 1);
        set_u32(&mut os2, 78, 1);

        let mut head = vec![0; 54];
        set_u32(&mut head, 0, 0x0001_0000);
        set_u32(&mut head, 8, 0x1234_5678);
        set_u32(&mut head, 12, 0x5F0F_3CF5);
        set_u16(&mut head, 18, 1000);

        let name = name_table(&[
            (NameTable::FONT_FAMILY_NAME, "Litchi Test"),
            (NameTable::FONT_SUBFAMILY_NAME, "Regular"),
            (NameTable::FULL_FONT_NAME, "Litchi Test Regular"),
            (NameTable::VERSION_STRING, "Version 1.0"),
        ]);
        sfnt(&[(b"OS/2", os2), (b"head", head), (b"name", name)])
    }

    fn name_table(values: &[(u16, &str)]) -> Vec<u8> {
        let records_bytes = values.len() * 12;
        let string_offset = 6 + records_bytes;
        let mut strings = Vec::new();
        let mut records = Vec::new();
        for (id, value) in values {
            let offset = strings.len();
            for unit in value.encode_utf16() {
                strings.extend_from_slice(&unit.to_be_bytes());
            }
            records.push((*id, offset, strings.len() - offset));
        }
        let mut output = vec![0; string_offset];
        set_u16(&mut output, 2, values.len() as u16);
        set_u16(&mut output, 4, string_offset as u16);
        for (index, (id, offset, length)) in records.into_iter().enumerate() {
            let start = 6 + index * 12;
            set_u16(&mut output, start, 3);
            set_u16(&mut output, start + 2, 1);
            set_u16(&mut output, start + 4, 0x0409);
            set_u16(&mut output, start + 6, id);
            set_u16(&mut output, start + 8, length as u16);
            set_u16(&mut output, start + 10, offset as u16);
        }
        output.extend_from_slice(&strings);
        output
    }

    fn sfnt(tables: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
        let directory = 12 + tables.len() * 16;
        let mut offsets = Vec::new();
        let mut length = directory;
        for (_, table) in tables {
            length = (length + 3) & !3;
            offsets.push(length);
            length += table.len();
        }
        let mut output = vec![0; length];
        set_u32(&mut output, 0, 0x0001_0000);
        set_u16(&mut output, 4, tables.len() as u16);
        for (index, ((tag, table), offset)) in tables.iter().zip(offsets).enumerate() {
            let record = 12 + index * 16;
            output[record..record + 4].copy_from_slice(*tag);
            set_u32(&mut output, record + 8, offset as u32);
            set_u32(&mut output, record + 12, table.len() as u32);
            output[offset..offset + table.len()].copy_from_slice(table);
        }
        output
    }

    fn set_u16(bytes: &mut [u8], offset: usize, value: u16) {
        bytes[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
    }

    fn set_u32(bytes: &mut [u8], offset: usize, value: u32) {
        bytes[offset..offset + 4].copy_from_slice(&value.to_be_bytes());
    }

    fn u32_at(bytes: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
    }

    fn eot_v1_font_offset(bytes: &[u8]) -> usize {
        let mut offset = 82;
        for index in 0..4 {
            let size = usize::from(u16::from_le_bytes(
                bytes[offset..offset + 2].try_into().unwrap(),
            ));
            offset += 2 + size;
            if index != 3 {
                assert_eq!(&bytes[offset..offset + 2], &[0, 0]);
                offset += 2;
            }
        }
        offset
    }
}

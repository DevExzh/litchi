//! PowerPoint EOT publication.

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

    let sfnt = Sfnt::parse(&font.data)?;
    let os2 = Os2::parse(
        sfnt.table(*b"OS/2")
            .ok_or_else(|| invalid("OpenType font has no OS/2 table"))?,
    )?;
    let license = font.properties.license();
    super::validate_license(&font.name, license)?;

    let head = sfnt
        .table(*b"head")
        .ok_or_else(|| invalid("OpenType font has no head table"))?;
    let check_sum_adjustment = be_u32(head, 8, "head checkSumAdjustment")?;
    let names = sfnt
        .table(*b"name")
        .ok_or_else(|| invalid("OpenType font has no name table"))?;
    let family = name_string(names, 1).unwrap_or_else(|| font.name.clone());
    let style = name_string(names, 2).unwrap_or_else(|| "Regular".into());
    let version = name_string(names, 5).unwrap_or_default();
    let full = name_string(names, 4).unwrap_or_else(|| font.name.clone());
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
    header.extend_from_slice(os2.panose());
    header.push(font.properties.charset().map_or(1, crate::Charset::code));
    header.push(u8::from(os2.italic()));
    push_u32(&mut header, u32::from(os2.weight()));
    push_u16(&mut header, license.bits());
    push_u16(&mut header, EOT_MAGIC);
    for range in [
        os2.unicode_ranges()[0],
        os2.unicode_ranges()[1],
        os2.unicode_ranges()[2],
        os2.unicode_ranges()[3],
    ] {
        push_u32(&mut header, range);
    }
    let (code_page1, code_page2) = os2.code_pages();
    push_u32(&mut header, code_page1);
    push_u32(&mut header, code_page2);
    push_u32(&mut header, check_sum_adjustment);
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

struct Sfnt<'a> {
    data: &'a [u8],
    directory: usize,
    table_count: usize,
}

impl<'a> Sfnt<'a> {
    fn parse(data: &'a [u8]) -> Result<Self> {
        let signature = data
            .get(..4)
            .ok_or_else(|| invalid("OpenType font is missing an sfnt signature"))?;
        if signature == b"ttcf" {
            return Err(invalid(
                "PowerPoint EOT authoring currently requires one standalone OpenType face",
            ));
        }
        if !matches!(signature, b"\0\x01\0\0" | b"OTTO" | b"true" | b"typ1") {
            return Err(invalid("invalid standalone OpenType sfnt signature"));
        }
        let table_count = usize::from(be_u16(data, 4, "sfnt table count")?);
        let directory_len = table_count
            .checked_mul(16)
            .and_then(|length| 12usize.checked_add(length))
            .ok_or_else(|| invalid("sfnt table directory overflows"))?;
        if data.len() < directory_len {
            return Err(invalid("truncated sfnt table directory"));
        }
        Ok(Self {
            data,
            directory: 12,
            table_count,
        })
    }

    fn table(&self, wanted: [u8; 4]) -> Option<&'a [u8]> {
        for index in 0..self.table_count {
            let record = self.directory + index * 16;
            if self.data.get(record..record + 4)? != wanted {
                continue;
            }
            let offset =
                usize::try_from(be_u32(self.data, record + 8, "sfnt table offset").ok()?).ok()?;
            let length =
                usize::try_from(be_u32(self.data, record + 12, "sfnt table length").ok()?).ok()?;
            let end = offset.checked_add(length)?;
            return self.data.get(offset..end);
        }
        None
    }
}

struct Os2<'a> {
    bytes: &'a [u8],
    version: u16,
}

impl<'a> Os2<'a> {
    fn parse(bytes: &'a [u8]) -> Result<Self> {
        let version = be_u16(bytes, 0, "OS/2 version")?;
        let minimum = match version {
            0 => 78,
            1 => 86,
            2..=4 => 96,
            5 => 100,
            _ => {
                return Err(invalid(format!(
                    "unsupported OpenType OS/2 table version {version}"
                )));
            },
        };
        if bytes.len() < minimum {
            return Err(invalid(format!(
                "truncated OpenType OS/2 version {version} table"
            )));
        }
        Ok(Self { bytes, version })
    }

    fn panose(&self) -> &'a [u8] {
        &self.bytes[32..42]
    }

    fn italic(&self) -> bool {
        u16::from_be_bytes(
            self.bytes[62..64]
                .try_into()
                .expect("validated OS/2 length"),
        ) & 1
            != 0
    }

    fn weight(&self) -> u16 {
        u16::from_be_bytes(self.bytes[4..6].try_into().expect("validated OS/2 length"))
    }

    fn unicode_ranges(&self) -> [u32; 4] {
        [
            u32::from_be_bytes(
                self.bytes[42..46]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[46..50]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[50..54]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[54..58]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
        ]
    }

    fn code_pages(&self) -> (u32, u32) {
        if self.version == 0 {
            return (0, 0);
        }
        (
            u32::from_be_bytes(
                self.bytes[78..82]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[82..86]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
        )
    }
}

fn name_string(table: &[u8], wanted: u16) -> Option<String> {
    let count = usize::from(be_u16(table, 2, "name record count").ok()?);
    let strings = usize::from(be_u16(table, 4, "name string offset").ok()?);
    let records_end = 6usize.checked_add(count.checked_mul(12)?)?;
    if records_end > table.len() || strings > table.len() {
        return None;
    }

    let mut selected = None;
    for index in 0..count {
        let record = 6 + index * 12;
        let platform = be_u16(table, record, "name platform").ok()?;
        let encoding = be_u16(table, record + 2, "name encoding").ok()?;
        if be_u16(table, record + 6, "name id").ok()? != wanted {
            continue;
        }
        let length = usize::from(be_u16(table, record + 8, "name string length").ok()?);
        let offset = usize::from(be_u16(table, record + 10, "name string offset").ok()?);
        let start = strings.checked_add(offset)?;
        let end = start.checked_add(length)?;
        let value = table.get(start..end)?;
        let rank = match platform {
            3 if encoding <= 10 => 0,
            0 => 1,
            _ => 2,
        };
        let decoded = match platform {
            0 | 3 if value.len() % 2 == 0 => {
                let units = value
                    .chunks_exact(2)
                    .map(|bytes| u16::from_be_bytes([bytes[0], bytes[1]]))
                    .collect::<Vec<_>>();
                String::from_utf16(&units).ok()
            },
            _ => std::str::from_utf8(value).ok().map(str::to_owned),
        }?;
        if selected.as_ref().is_none_or(|(best, _)| rank < *best) {
            selected = Some((rank, decoded));
        }
    }
    selected.map(|(_, value)| value)
}

fn be_u16(bytes: &[u8], offset: usize, field: &str) -> Result<u16> {
    let value = bytes
        .get(
            offset
                ..offset
                    .checked_add(2)
                    .ok_or_else(|| invalid(format!("{field} offset overflows")))?,
        )
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u16::from_be_bytes(
        value.try_into().expect("two-byte slice"),
    ))
}

fn be_u32(bytes: &[u8], offset: usize, field: &str) -> Result<u32> {
    let value = bytes
        .get(
            offset
                ..offset
                    .checked_add(4)
                    .ok_or_else(|| invalid(format!("{field} offset overflows")))?,
        )
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u32::from_be_bytes(
        value.try_into().expect("four-byte slice"),
    ))
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
            (1, "Litchi Test"),
            (2, "Regular"),
            (4, "Litchi Test Regular"),
            (5, "Version 1.0"),
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

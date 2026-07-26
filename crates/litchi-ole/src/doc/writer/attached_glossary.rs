//! Assembly of a secondary glossary FIB and its independent text graph.

use super::core::DocWriteError;

const FIB_SIZE: usize = 1248;
const FIB_PN_NEXT_OFFSET: usize = 8;
const FIB_FLAGS_OFFSET: usize = 10;
const FIB_FC_MIN_OFFSET: usize = 24;
const FIB_FC_MAC_OFFSET: usize = 28;
const FIB_CB_MAC_OFFSET: usize = 64;
const FIB_POINTERS_OFFSET: usize = 154;
const FIB_POINTER_COUNT: usize = 136;
const FIB_POINTER_SIZE: usize = 8;
const PLCF_SED_INDEX: usize = 6;
const PLCF_BTE_CHPX_INDEX: usize = 12;
const PLCF_BTE_PAPX_INDEX: usize = 13;
const CLX_INDEX: usize = 33;
const WORD_PAGE_SIZE: usize = 512;
const REGULAR_STREAM_ALIGNMENT: usize = 4096;
const FLAG_TEMPLATE: u16 = 0x0001;
const FLAG_GLOSSARY: u16 = 0x0002;
const COMPRESSED_FC_FLAG: u32 = 0x4000_0000;

#[derive(Clone, Copy)]
struct BinEntry {
    start_fc: u32,
    end_fc: u32,
    page_number: u32,
}

fn invalid(message: impl Into<String>) -> DocWriteError {
    DocWriteError::InvalidData(message.into())
}

fn read_u16(data: &[u8], offset: usize, label: &str) -> Result<u16, DocWriteError> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| invalid(format!("{label} is truncated")))?;
    Ok(u16::from_le_bytes(
        bytes.try_into().expect("two-byte slice"),
    ))
}

fn read_u32(data: &[u8], offset: usize, label: &str) -> Result<u32, DocWriteError> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| invalid(format!("{label} is truncated")))?;
    Ok(u32::from_le_bytes(
        bytes.try_into().expect("four-byte slice"),
    ))
}

fn write_u16(data: &mut [u8], offset: usize, value: u16, label: &str) -> Result<(), DocWriteError> {
    let target = data
        .get_mut(offset..offset + 2)
        .ok_or_else(|| invalid(format!("{label} is truncated")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_u32(data: &mut [u8], offset: usize, value: u32, label: &str) -> Result<(), DocWriteError> {
    let target = data
        .get_mut(offset..offset + 4)
        .ok_or_else(|| invalid(format!("{label} is truncated")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn checked_add(value: u32, delta: u32, label: &str) -> Result<u32, DocWriteError> {
    value
        .checked_add(delta)
        .ok_or_else(|| invalid(format!("{label} exceeds 32-bit DOC address space")))
}

fn pointer(fib: &[u8], index: usize) -> Result<(u32, u32), DocWriteError> {
    let offset = FIB_POINTERS_OFFSET
        .checked_add(
            index
                .checked_mul(FIB_POINTER_SIZE)
                .ok_or_else(|| invalid("DOC FIB pointer index overflows"))?,
        )
        .ok_or_else(|| invalid("DOC FIB pointer offset overflows"))?;
    Ok((
        read_u32(fib, offset, "DOC FIB table pointer")?,
        read_u32(fib, offset + 4, "DOC FIB table length")?,
    ))
}

fn set_pointer(
    fib: &mut [u8],
    index: usize,
    offset: u32,
    length: u32,
) -> Result<(), DocWriteError> {
    let position = FIB_POINTERS_OFFSET + index * FIB_POINTER_SIZE;
    write_u32(fib, position, offset, "DOC FIB table pointer")?;
    write_u32(fib, position + 4, length, "DOC FIB table length")
}

fn table_range(
    table: &[u8],
    offset: u32,
    length: u32,
    label: &str,
) -> Result<std::ops::Range<usize>, DocWriteError> {
    let start =
        usize::try_from(offset).map_err(|_| invalid(format!("{label} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| invalid(format!("{label} length is too large")))?;
    let end = start
        .checked_add(length)
        .filter(|end| *end <= table.len())
        .ok_or_else(|| invalid(format!("{label} lies outside the DOC table stream")))?;
    Ok(start..end)
}

fn parse_bin_table(
    fib: &[u8],
    table: &[u8],
    index: usize,
    label: &str,
) -> Result<Vec<BinEntry>, DocWriteError> {
    let (offset, length) = pointer(fib, index)?;
    let range = table_range(table, offset, length, label)?;
    let bytes = &table[range];
    if bytes.len() < 12 || (bytes.len() - 4) % 8 != 0 {
        return Err(invalid(format!("{label} has an invalid PLCF length")));
    }
    let count = (bytes.len() - 4) / 8;
    let page_numbers_offset = (count + 1) * 4;
    let mut entries = Vec::with_capacity(count);
    for index in 0..count {
        let start_fc = read_u32(bytes, index * 4, label)?;
        let end_fc = read_u32(bytes, (index + 1) * 4, label)?;
        let page_number = read_u32(bytes, page_numbers_offset + index * 4, label)?;
        if start_fc > end_fc {
            return Err(invalid(format!("{label} FC boundaries are not ordered")));
        }
        entries.push(BinEntry {
            start_fc,
            end_fc,
            page_number,
        });
    }
    Ok(entries)
}

fn relocate_fkp_pages(
    word_document: &mut [u8],
    entries: &mut [BinEntry],
    word_delta: u32,
    label: &str,
) -> Result<(), DocWriteError> {
    let page_delta = word_delta / WORD_PAGE_SIZE as u32;
    let mut previous_page = None;
    for entry in entries {
        entry.start_fc = checked_add(entry.start_fc, word_delta, label)?;
        entry.end_fc = checked_add(entry.end_fc, word_delta, label)?;
        if previous_page != Some(entry.page_number) {
            let page_offset = usize::try_from(entry.page_number)
                .ok()
                .and_then(|page| page.checked_mul(WORD_PAGE_SIZE))
                .ok_or_else(|| invalid(format!("{label} page offset overflows")))?;
            let page = word_document
                .get_mut(page_offset..page_offset + WORD_PAGE_SIZE)
                .ok_or_else(|| invalid(format!("{label} references a missing FKP page")))?;
            let run_count = usize::from(page[WORD_PAGE_SIZE - 1]);
            let fc_bytes = (run_count + 1)
                .checked_mul(4)
                .filter(|length| *length <= WORD_PAGE_SIZE - 1)
                .ok_or_else(|| invalid(format!("{label} FKP run count is invalid")))?;
            for offset in (0..fc_bytes).step_by(4) {
                let value = read_u32(page, offset, label)?;
                write_u32(page, offset, checked_add(value, word_delta, label)?, label)?;
            }
            previous_page = Some(entry.page_number);
        }
        entry.page_number = checked_add(entry.page_number, page_delta, label)?;
    }
    Ok(())
}

fn relocate_clx(fib: &[u8], table: &mut [u8], word_delta: u32) -> Result<(), DocWriteError> {
    let (offset, length) = pointer(fib, CLX_INDEX)?;
    let range = table_range(table, offset, length, "attached glossary CLX")?;
    let clx = &mut table[range];
    if clx.first() != Some(&0x02) || clx.len() < 9 {
        return Err(invalid(
            "attached glossary writer emitted an unsupported CLX layout",
        ));
    }
    let plc_length = usize::try_from(read_u32(clx, 1, "attached glossary CLX")?)
        .map_err(|_| invalid("attached glossary CLX length is too large"))?;
    if plc_length.checked_add(5) != Some(clx.len()) || plc_length < 4 || (plc_length - 4) % 12 != 0
    {
        return Err(invalid(
            "attached glossary CLX has an invalid PlcPcd length",
        ));
    }
    let piece_count = (plc_length - 4) / 12;
    let descriptors_offset = 5 + (piece_count + 1) * 4;
    for index in 0..piece_count {
        let fc_offset = descriptors_offset + index * 8 + 2;
        let encoded = read_u32(clx, fc_offset, "attached glossary PCD")?;
        let relocated = if encoded & COMPRESSED_FC_FLAG == 0 {
            checked_add(encoded, word_delta, "attached glossary text FC")?
        } else {
            let physical = (encoded & !COMPRESSED_FC_FLAG) / 2;
            let relocated = checked_add(physical, word_delta, "attached glossary text FC")?;
            relocated
                .checked_mul(2)
                .filter(|value| value & COMPRESSED_FC_FLAG == 0)
                .map(|value| value | COMPRESSED_FC_FLAG)
                .ok_or_else(|| invalid("attached glossary compressed FC overflows"))?
        };
        write_u32(clx, fc_offset, relocated, "attached glossary PCD")?;
    }
    Ok(())
}

fn relocate_section_table(
    fib: &[u8],
    table: &mut [u8],
    word_delta: u32,
) -> Result<(), DocWriteError> {
    let (offset, length) = pointer(fib, PLCF_SED_INDEX)?;
    let range = table_range(table, offset, length, "attached glossary PlcfSed")?;
    let plc = &mut table[range];
    if plc.len() < 20 || (plc.len() - 4) % 16 != 0 {
        return Err(invalid("attached glossary PlcfSed has an invalid length"));
    }
    let count = (plc.len() - 4) / 16;
    let descriptors_offset = (count + 1) * 4;
    for index in 0..count {
        let descriptor = descriptors_offset + index * 12;
        for relative_offset in [2usize, 8] {
            let value = read_u32(plc, descriptor + relative_offset, "attached glossary SED")?;
            if value != 0 && value != u32::MAX {
                write_u32(
                    plc,
                    descriptor + relative_offset,
                    checked_add(value, word_delta, "attached glossary SED FC")?,
                    "attached glossary SED",
                )?;
            }
        }
    }
    Ok(())
}

fn relocate_table_pointers(
    fib: &mut [u8],
    table: &[u8],
    table_delta: u32,
) -> Result<(), DocWriteError> {
    for index in 0..FIB_POINTER_COUNT {
        let (offset, length) = pointer(fib, index)?;
        if length == 0 {
            continue;
        }
        table_range(table, offset, length, "attached glossary FIB table data")?;
        set_pointer(
            fib,
            index,
            checked_add(offset, table_delta, "attached glossary table FC")?,
            length,
        )?;
    }
    Ok(())
}

fn generate_bin_table(entries: &[BinEntry], label: &str) -> Result<Vec<u8>, DocWriteError> {
    if entries.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    let mut bytes = Vec::with_capacity(entries.len() * 8 + 4);
    let mut previous_start = None;
    for entry in entries {
        if previous_start.is_some_and(|start| entry.start_fc < start) {
            return Err(invalid(format!("{label} is not ordered by FC")));
        }
        bytes.extend_from_slice(&entry.start_fc.to_le_bytes());
        previous_start = Some(entry.start_fc);
    }
    bytes.extend_from_slice(
        &entries
            .last()
            .expect("nonempty bin table")
            .end_fc
            .to_le_bytes(),
    );
    for entry in entries {
        bytes.extend_from_slice(&entry.page_number.to_le_bytes());
    }
    Ok(bytes)
}

fn set_flags(fib: &mut [u8], set: u16, clear: u16, label: &str) -> Result<(), DocWriteError> {
    let flags = read_u16(fib, FIB_FLAGS_OFFSET, label)?;
    write_u16(fib, FIB_FLAGS_OFFSET, (flags | set) & !clear, label)
}

fn pad_regular_stream(stream: &mut Vec<u8>) {
    let remainder = stream.len() % REGULAR_STREAM_ALIGNMENT;
    if remainder != 0 {
        stream.resize(stream.len() + REGULAR_STREAM_ALIGNMENT - remainder, 0);
    }
}

/// Merge a separately generated glossary-only DOC into a template story graph.
pub(super) fn merge_attached_glossary(
    main_word: &mut Vec<u8>,
    main_table: &mut Vec<u8>,
    glossary_word: &mut Vec<u8>,
    glossary_table: &mut Vec<u8>,
) -> Result<(), DocWriteError> {
    if main_word.len() < FIB_SIZE || glossary_word.len() < FIB_SIZE {
        return Err(invalid("DOC writer emitted a truncated FIB"));
    }
    if main_word.len() % WORD_PAGE_SIZE != 0 {
        return Err(invalid("main WordDocument stream is not page-aligned"));
    }

    let word_delta = u32::try_from(main_word.len())
        .map_err(|_| invalid("main WordDocument stream exceeds 32-bit DOC address space"))?;
    let next_page = u16::try_from(main_word.len() / WORD_PAGE_SIZE)
        .map_err(|_| invalid("attached glossary FIB page exceeds FibBase.pnNext"))?;
    let table_delta = u32::try_from(main_table.len())
        .map_err(|_| invalid("main DOC table stream exceeds 32-bit address space"))?;

    let main_fib = main_word[..FIB_SIZE].to_vec();
    let glossary_fib = glossary_word[..FIB_SIZE].to_vec();
    let mut main_chpx = parse_bin_table(
        &main_fib,
        main_table,
        PLCF_BTE_CHPX_INDEX,
        "main PlcfBteChpx",
    )?;
    let mut main_papx = parse_bin_table(
        &main_fib,
        main_table,
        PLCF_BTE_PAPX_INDEX,
        "main PlcfBtePapx",
    )?;
    let mut glossary_chpx = parse_bin_table(
        &glossary_fib,
        glossary_table,
        PLCF_BTE_CHPX_INDEX,
        "attached glossary PlcfBteChpx",
    )?;
    let mut glossary_papx = parse_bin_table(
        &glossary_fib,
        glossary_table,
        PLCF_BTE_PAPX_INDEX,
        "attached glossary PlcfBtePapx",
    )?;

    relocate_fkp_pages(
        glossary_word,
        &mut glossary_chpx,
        word_delta,
        "attached glossary CHPX",
    )?;
    relocate_fkp_pages(
        glossary_word,
        &mut glossary_papx,
        word_delta,
        "attached glossary PAPX",
    )?;
    relocate_clx(&glossary_fib, glossary_table, word_delta)?;
    relocate_section_table(&glossary_fib, glossary_table, word_delta)?;

    let glossary_fc_min = checked_add(
        read_u32(glossary_word, FIB_FC_MIN_OFFSET, "attached glossary fcMin")?,
        word_delta,
        "attached glossary fcMin",
    )?;
    let glossary_fc_mac = checked_add(
        read_u32(glossary_word, FIB_FC_MAC_OFFSET, "attached glossary fcMac")?,
        word_delta,
        "attached glossary fcMac",
    )?;
    write_u32(
        glossary_word,
        FIB_FC_MIN_OFFSET,
        glossary_fc_min,
        "attached glossary fcMin",
    )?;
    write_u32(
        glossary_word,
        FIB_FC_MAC_OFFSET,
        glossary_fc_mac,
        "attached glossary fcMac",
    )?;
    relocate_table_pointers(&mut glossary_word[..FIB_SIZE], glossary_table, table_delta)?;

    main_chpx.append(&mut glossary_chpx);
    main_papx.append(&mut glossary_papx);
    let merged_chpx = generate_bin_table(&main_chpx, "shared PlcfBteChpx")?;
    let merged_papx = generate_bin_table(&main_papx, "shared PlcfBtePapx")?;

    main_word.extend_from_slice(glossary_word);
    main_table.extend_from_slice(glossary_table);
    let shared_chpx_offset = u32::try_from(main_table.len())
        .map_err(|_| invalid("combined DOC table stream exceeds 32-bit address space"))?;
    main_table.extend_from_slice(&merged_chpx);
    let shared_papx_offset = u32::try_from(main_table.len())
        .map_err(|_| invalid("combined DOC table stream exceeds 32-bit address space"))?;
    main_table.extend_from_slice(&merged_papx);

    let logical_word_size = checked_add(
        word_delta,
        read_u32(&glossary_fib, FIB_CB_MAC_OFFSET, "attached glossary cbMac")?,
        "combined WordDocument size",
    )?;
    let secondary_offset = usize::try_from(word_delta)
        .map_err(|_| invalid("attached glossary FIB offset is too large"))?;
    let (main_fib_mut, secondary_and_after) = main_word.split_at_mut(secondary_offset);
    let secondary_fib_mut = secondary_and_after
        .get_mut(..FIB_SIZE)
        .ok_or_else(|| invalid("attached glossary FIB is truncated after assembly"))?;

    write_u16(main_fib_mut, FIB_PN_NEXT_OFFSET, next_page, "main pnNext")?;
    set_flags(main_fib_mut, FLAG_TEMPLATE, FLAG_GLOSSARY, "main FIB flags")?;
    write_u32(
        main_fib_mut,
        FIB_CB_MAC_OFFSET,
        logical_word_size,
        "main cbMac",
    )?;
    set_pointer(
        main_fib_mut,
        PLCF_BTE_CHPX_INDEX,
        shared_chpx_offset,
        u32::try_from(merged_chpx.len()).map_err(|_| invalid("shared PlcfBteChpx is too large"))?,
    )?;
    set_pointer(
        main_fib_mut,
        PLCF_BTE_PAPX_INDEX,
        shared_papx_offset,
        u32::try_from(merged_papx.len()).map_err(|_| invalid("shared PlcfBtePapx is too large"))?,
    )?;

    write_u16(
        secondary_fib_mut,
        FIB_PN_NEXT_OFFSET,
        0,
        "attached glossary pnNext",
    )?;
    set_flags(
        secondary_fib_mut,
        FLAG_GLOSSARY,
        FLAG_TEMPLATE,
        "attached glossary FIB flags",
    )?;
    write_u32(
        secondary_fib_mut,
        FIB_CB_MAC_OFFSET,
        logical_word_size,
        "attached glossary cbMac",
    )?;
    set_pointer(
        secondary_fib_mut,
        PLCF_BTE_CHPX_INDEX,
        shared_chpx_offset,
        u32::try_from(merged_chpx.len()).map_err(|_| invalid("shared PlcfBteChpx is too large"))?,
    )?;
    set_pointer(
        secondary_fib_mut,
        PLCF_BTE_PAPX_INDEX,
        shared_papx_offset,
        u32::try_from(merged_papx.len()).map_err(|_| invalid("shared PlcfBtePapx is too large"))?,
    )?;

    pad_regular_stream(main_word);
    pad_regular_stream(main_table);
    Ok(())
}

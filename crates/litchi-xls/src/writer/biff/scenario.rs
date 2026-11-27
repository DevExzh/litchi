use std::io::Write;

use crate::scenario::ScenarioManager;
use crate::writer::biff::write_record_header;
use crate::{Error, Result};

const MAX_RECORD_DATA: usize = 8224;

pub(crate) fn write_scenario_manager<W: Write>(
    writer: &mut W,
    manager: &ScenarioManager,
) -> Result<()> {
    manager.validate_for_write()?;
    let manager_len = 8usize + manager.result_ranges.len() * 8;
    write_record_header(
        writer,
        0x00AE,
        crate::utils::truncate_usize_to_u16(manager_len),
    )?;
    writer.write_all(&crate::utils::wrap_usize_to_i16(manager.scenarios.len()).to_le_bytes())?;
    write_index(writer, manager.current_scenario)?;
    write_index(writer, manager.shown_scenario)?;
    writer
        .write_all(&crate::utils::wrap_usize_to_i16(manager.result_ranges.len()).to_le_bytes())?;
    for range in &manager.result_ranges {
        writer.write_all(&range.first_row.to_le_bytes())?;
        writer.write_all(&range.last_row.to_le_bytes())?;
        writer.write_all(&u16::from(range.first_column).to_le_bytes())?;
        writer.write_all(&u16::from(range.last_column).to_le_bytes())?;
    }
    for scenario in &manager.scenarios {
        write_scenario(writer, scenario)?;
    }
    Ok(())
}

fn write_index<W: Write>(writer: &mut W, index: Option<usize>) -> Result<()> {
    let value = index.map_or(-1, crate::utils::wrap_usize_to_i16);
    writer.write_all(&value.to_le_bytes())?;
    Ok(())
}

fn write_scenario<W: Write>(writer: &mut W, scenario: &crate::scenario::Scenario) -> Result<()> {
    let mut prefix = Vec::new();
    prefix.extend_from_slice(
        &crate::utils::truncate_usize_to_u16(scenario.cells.len()).to_le_bytes(),
    );
    prefix.push(u8::from(scenario.locked));
    prefix.push(u8::from(scenario.hidden));
    prefix.push(crate::utils::truncate_usize_to_u8(
        scenario.name.encode_utf16().count(),
    ));
    prefix.push(scenario.comment.as_ref().map_or(0, |value| {
        crate::utils::truncate_usize_to_u8(value.encode_utf16().count())
    }));
    prefix.push(scenario.creator.as_ref().map_or(0, |value| {
        crate::utils::truncate_usize_to_u8(value.encode_utf16().count())
    }));
    encode_no_cch(&mut prefix, &scenario.name);
    if let Some(creator) = &scenario.creator {
        encode_unicode(&mut prefix, creator)?;
    }
    if let Some(comment) = &scenario.comment {
        encode_unicode(&mut prefix, comment)?;
    }
    for cell in &scenario.cells {
        prefix.extend_from_slice(&cell.row.to_le_bytes());
        let flags = u16::from(cell.column) | if cell.deleted { 0x4000 } else { 0 };
        prefix.extend_from_slice(&flags.to_le_bytes());
    }
    if prefix.len() > MAX_RECORD_DATA {
        return Err(Error::InvalidData(
            "Scenario fixed fields exceed BIFF8 record limit".to_string(),
        ));
    }

    let mut chunks = vec![prefix];
    for cell in &scenario.cells {
        let mut encoded = Vec::new();
        encode_unicode(&mut encoded, &cell.value)?;
        push_component(&mut chunks, encoded)?;
    }
    push_component(&mut chunks, vec![0; scenario.cells.len() * 2])?;
    for (index, chunk) in chunks.iter().enumerate() {
        write_record_header(
            writer,
            if index == 0 { 0x00AF } else { 0x003C },
            crate::utils::truncate_usize_to_u16(chunk.len()),
        )?;
        writer.write_all(chunk)?;
    }
    Ok(())
}

fn push_component(chunks: &mut Vec<Vec<u8>>, component: Vec<u8>) -> Result<()> {
    if component.len() > MAX_RECORD_DATA {
        return Err(Error::InvalidData(
            "individual Scenario value exceeds BIFF8 continuation-safe limit".to_string(),
        ));
    }
    if chunks.last().unwrap().len() + component.len() > MAX_RECORD_DATA {
        chunks.push(component);
    } else {
        chunks.last_mut().unwrap().extend_from_slice(&component);
    }
    Ok(())
}

fn encode_unicode(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let count = value.encode_utf16().count();
    if count > u16::MAX as usize {
        return Err(Error::InvalidData(
            "Scenario string exceeds 65535 UTF-16 code units".to_string(),
        ));
    }
    output.extend_from_slice(&crate::utils::truncate_usize_to_u16(count).to_le_bytes());
    encode_no_cch(output, value);
    Ok(())
}

fn encode_no_cch(output: &mut Vec<u8>, value: &str) {
    if value.chars().all(|character| u32::from(character) <= 0xFF) {
        output.push(0);
        output.extend(value.chars().map(|character| character as u8));
    } else {
        output.push(1);
        for unit in value.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

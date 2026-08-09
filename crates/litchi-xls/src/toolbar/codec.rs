//! Binary codec for the bounded XLS Office Toolbars (`XCB`) stream.

use crate::{Error, Result};
use litchi_ole_common::toolbar::{ControlHeader, Data, Error as SharedError, Header};

use super::model::{APPLICATION_TOOLBAR_ID, VISUAL_DATA_LEN};
use super::validation::{self, MAX_CONTROLS};
use super::{Command, Control, Toolbar, ToolbarSet, VisualData, Wrapper};

const CTBS_LEN: usize = 14;
const APPLICATION_ID_LEN: usize = 4;

/// Parse exactly one `CTBWRAPPER`/`XCB` stream.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse(data: &[u8]) -> Result<Wrapper<'_>> {
    let (toolbar_set, offset) = parse_toolbar_set(data)?;
    let count = usize::from(toolbar_set.toolbar_count());
    let (toolbars, consumed) = parse_sequence(data, offset, &toolbar_set, count)?;
    if consumed != data.len() {
        return Err(validation::invalid(format!(
            "XCB has {} trailing bytes",
            data.len() - consumed
        )));
    }
    Wrapper::from_parts(toolbar_set, toolbars)
}

/// Serialize one complete `CTBWRAPPER`/`XCB` stream.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn to_bytes(value: &Wrapper<'_>) -> Result<Vec<u8>> {
    value.validate()?;
    let mut output = Vec::new();
    encode_toolbar_set(&mut output, value.toolbar_set());
    for toolbar in value.toolbars() {
        encode_toolbar(&mut output, toolbar);
    }
    Ok(output)
}

fn parse_toolbar_set(data: &[u8]) -> Result<(ToolbarSet, usize)> {
    let bytes = data
        .get(..CTBS_LEN)
        .ok_or_else(|| Error::UnexpectedEndOfStream("XCB CTBS header".to_string()))?;
    let toolbar_set = ToolbarSet::from_parts(
        bytes[0],
        bytes[1],
        read_u16(bytes, 2),
        read_u16(bytes, 4),
        read_u16(bytes, 6),
        read_u16(bytes, 8),
        read_u16(bytes, 10),
        read_u16(bytes, 12),
    )?;
    Ok((toolbar_set, CTBS_LEN))
}

fn parse_sequence<'a>(
    data: &'a [u8],
    offset: usize,
    toolbar_set: &ToolbarSet,
    remaining: usize,
) -> Result<(Vec<Toolbar<'a>>, usize)> {
    if remaining == 0 {
        return Ok((Vec::new(), offset));
    }

    let candidates = parse_toolbar_candidates(data, offset)?;
    let mut success: Option<(Vec<Toolbar<'a>>, usize)> = None;
    let mut first_error = None;

    for (toolbar, next_offset) in candidates {
        match parse_sequence(data, next_offset, toolbar_set, remaining - 1) {
            Ok((mut rest, end)) => {
                if success.is_some() {
                    return Err(validation::unsupported(
                        "CTB optional rVisualData has ambiguous boundaries",
                    ));
                }
                let mut toolbars = Vec::with_capacity(rest.len() + 1);
                toolbars.push(toolbar);
                toolbars.append(&mut rest);
                success = Some((toolbars, end));
            },
            Err(error) => {
                if first_error.is_none() {
                    first_error = Some(error);
                }
            },
        }
    }

    success.ok_or_else(|| {
        first_error.unwrap_or_else(|| validation::invalid("CTB record has no valid boundary"))
    })
}

fn parse_toolbar_candidates(data: &[u8], offset: usize) -> Result<Vec<(Toolbar<'_>, usize)>> {
    let bytes = data
        .get(offset..)
        .ok_or_else(|| Error::UnexpectedEndOfStream("XCB CTB toolbar".to_string()))?;
    let (header, header_len) = Header::parse_prefix(bytes).map_err(map_shared)?;
    let count = header.control_count();
    if count < 0 {
        return Err(validation::invalid("TB cCL must not be negative"));
    }
    let count = usize::try_from(count).map_err(|_error| validation::invalid("TB cCL overflows"))?;
    if count > MAX_CONTROLS {
        return Err(validation::invalid("TB cCL exceeds the bounded limit"));
    }

    let toolbar_data = offset
        .checked_add(header_len)
        .ok_or_else(|| validation::invalid("CTB toolbar offset overflows"))?;
    let mut candidates = Vec::new();
    let mut first_error = None;
    for visual_len in [0usize, VISUAL_DATA_LEN] {
        let app_offset = toolbar_data
            .checked_add(visual_len)
            .ok_or_else(|| validation::invalid("CTB visual-data offset overflows"))?;
        let end = app_offset
            .checked_add(APPLICATION_ID_LEN)
            .ok_or_else(|| validation::invalid("CTB ectbid offset overflows"))?;
        let Some(application_bytes) = data.get(app_offset..end) else {
            continue;
        };
        let application_id = i32::from_le_bytes([
            application_bytes[0],
            application_bytes[1],
            application_bytes[2],
            application_bytes[3],
        ]);
        if application_id != APPLICATION_TOOLBAR_ID {
            continue;
        }
        let visual_data = if visual_len == 0 {
            None
        } else {
            let visual = data
                .get(toolbar_data..app_offset)
                .ok_or_else(|| Error::UnexpectedEndOfStream("XCB rVisualData".to_string()))?;
            Some(VisualData::new(visual.try_into().map_err(|_error| {
                validation::invalid("XCB rVisualData must be 60 bytes")
            })?))
        };
        match parse_controls(data, end, count) {
            Ok(control_candidates) => {
                for (controls, cursor) in control_candidates {
                    let toolbar =
                        Toolbar::from_parts(header.clone(), visual_data, application_id, controls)?;
                    candidates.push((toolbar, cursor));
                    if candidates.len() > validation::MAX_BOUNDARY_CANDIDATES {
                        return Err(validation::unsupported(
                            "XCB TBC boundaries exceed the bounded candidate limit",
                        ));
                    }
                }
            },
            Err(error) => {
                if first_error.is_none() {
                    first_error = Some(error);
                }
            },
        }
    }

    if candidates.is_empty() {
        let minimum = toolbar_data
            .checked_add(APPLICATION_ID_LEN)
            .ok_or_else(|| validation::invalid("CTB ectbid offset overflows"))?;
        if data.len() < minimum {
            return Err(Error::UnexpectedEndOfStream("XCB CTB ectbid".to_string()));
        }
        return Err(first_error.unwrap_or_else(|| {
            validation::invalid(
                "CTB ectbid must be 0x00000FFF or the optional visual-data block is truncated",
            )
        }));
    }
    Ok(candidates)
}

fn parse_controls(
    data: &[u8],
    offset: usize,
    remaining: usize,
) -> Result<Vec<(Vec<Control<'_>>, usize)>> {
    if remaining == 0 {
        return Ok(vec![(Vec::new(), offset)]);
    }

    let control_bytes = data
        .get(offset..)
        .ok_or_else(|| Error::UnexpectedEndOfStream("XCB TBC header".to_string()))?;
    let (control_header, consumed) =
        ControlHeader::parse_prefix(control_bytes).map_err(map_shared)?;
    let data_offset = offset
        .checked_add(consumed)
        .ok_or_else(|| validation::invalid("TBC header offset overflows"))?;

    if control_header.control_type().raw() == 0x16 {
        let control = Control::from_decoded(control_header, None, None)?;
        let tails = parse_controls(data, data_offset, remaining - 1)?;
        return Ok(prepend(control, tails));
    }

    let mut candidates = Vec::new();
    let mut first_error = None;

    let command_count = if validation::command_allowed(&control_header) {
        2
    } else {
        1
    };
    for command_index in 0..command_count {
        let command_len = if command_index == 0 { 0 } else { 4 };
        let data_start = data_offset
            .checked_add(command_len)
            .ok_or_else(|| validation::invalid("TBCCmd offset overflows"))?;
        let command = if command_len == 0 {
            None
        } else {
            let command_bytes: [u8; 4] = if let Some(bytes) = data.get(data_offset..data_start) {
                if let Ok(bytes) = bytes.try_into() {
                    bytes
                } else {
                    if first_error.is_none() {
                        first_error = Some(validation::invalid("TBCCmd must be four bytes"));
                    }
                    continue;
                }
            } else {
                if first_error.is_none() {
                    first_error = Some(Error::UnexpectedEndOfStream("XCB TBCCmd".to_string()));
                }
                continue;
            };
            match Command::from_bytes(command_bytes) {
                Ok(command) => Some(command),
                Err(error) => {
                    if first_error.is_none() {
                        first_error = Some(error);
                    }
                    continue;
                },
            }
        };

        let boundaries = find_boundaries(data, data_start, remaining > 1)?;
        for boundary in boundaries {
            let payload = if let Some(payload) = data.get(data_start..boundary) {
                payload
            } else {
                if first_error.is_none() {
                    first_error = Some(validation::invalid("TBCData boundary is out of range"));
                }
                continue;
            };
            let control_data = match Data::parse(payload) {
                Ok(value) => value,
                Err(error) => {
                    if first_error.is_none() {
                        first_error = Some(map_shared(error));
                    }
                    continue;
                },
            };
            let tails = match parse_controls(data, boundary, remaining - 1) {
                Ok(tails) => tails,
                Err(error) => {
                    if first_error.is_none() {
                        first_error = Some(error);
                    }
                    continue;
                },
            };
            let control = match Control::from_decoded(
                control_header.clone(),
                command,
                Some(control_data.clone()),
            ) {
                Ok(control) => control,
                Err(error) => {
                    if first_error.is_none() {
                        first_error = Some(error);
                    }
                    continue;
                },
            };
            for (tail, end) in tails {
                let mut controls = Vec::with_capacity(tail.len() + 1);
                controls.push(control.clone());
                controls.extend(tail);
                candidates.push((controls, end));
                if candidates.len() > validation::MAX_BOUNDARY_CANDIDATES {
                    return Err(validation::unsupported(
                        "XCB TBC boundaries exceed the bounded candidate limit",
                    ));
                }
            }
        }
    }

    if candidates.is_empty() {
        Err(first_error.unwrap_or_else(|| {
            validation::invalid("TBCData has no unambiguous TBCHeader boundary")
        }))
    } else {
        Ok(candidates)
    }
}

fn prepend<'a>(
    control: Control<'a>,
    tails: Vec<(Vec<Control<'a>>, usize)>,
) -> Vec<(Vec<Control<'a>>, usize)> {
    tails
        .into_iter()
        .map(|(tail, end)| {
            let mut controls = Vec::with_capacity(tail.len() + 1);
            controls.push(control.clone());
            controls.extend(tail);
            (controls, end)
        })
        .collect()
}

fn find_boundaries(data: &[u8], start: usize, has_next_control: bool) -> Result<Vec<usize>> {
    let signature = if has_next_control { 0x03 } else { 0x02 };
    let mut boundaries = Vec::new();
    if !has_next_control {
        boundaries.push(data.len());
    }

    let end = data.len().saturating_sub(1);
    for position in start..end {
        if data[position] != signature || data[position + 1] != 0x01 {
            continue;
        }
        let valid = if has_next_control {
            ControlHeader::parse_prefix(&data[position..]).is_ok()
        } else {
            Header::parse_prefix(&data[position..]).is_ok()
        };
        if valid {
            boundaries.push(position);
            if boundaries.len() > validation::MAX_BOUNDARY_CANDIDATES {
                return Err(validation::unsupported(
                    "XCB TBC boundary signatures exceed the bounded candidate limit",
                ));
            }
        }
    }
    Ok(boundaries)
}

fn encode_toolbar_set(output: &mut Vec<u8>, value: &ToolbarSet) {
    output.push(value.signature());
    output.push(value.version());
    output.extend_from_slice(&value.reserved1().to_le_bytes());
    output.extend_from_slice(&value.reserved2().to_le_bytes());
    output.extend_from_slice(&value.reserved3().to_le_bytes());
    output.extend_from_slice(&value.toolbar_count().to_le_bytes());
    output.extend_from_slice(&value.view_count().to_le_bytes());
    output.extend_from_slice(&value.active_view().to_le_bytes());
}

fn encode_toolbar(output: &mut Vec<u8>, value: &Toolbar<'_>) {
    output.extend_from_slice(&value.header().to_bytes());
    if let Some(visual_data) = value.visual_data() {
        output.extend_from_slice(visual_data.bytes());
    }
    output.extend_from_slice(&value.application_id().to_le_bytes());
    for control in value.controls() {
        output.extend_from_slice(&control.header().to_bytes());
        if let Some(command) = control.command() {
            output.extend_from_slice(&command.bytes());
        }
        if let Some(data) = control.data() {
            output.extend_from_slice(&data.to_bytes());
        }
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn map_shared(error: SharedError) -> Error {
    match error {
        SharedError::Truncated(field) => Error::UnexpectedEndOfStream(format!("XCB {field}")),
        SharedError::Invalid(message) => Error::InvalidData(format!("XCB {message}")),
    }
}

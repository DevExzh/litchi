//! Lossless binary codec and FIB/table-stream seam for command bars.

mod control;

use super::model::*;
use super::validation::validate_count;
use crate::package::{Error as PackageError, Result};
use litchi_ole_common::toolbar::Header;
use std::borrow::Cow;

const TCG_VERSION: u8 = 0xFF;
const TCG_TERMINATOR: u8 = 0x40;
const PLF_MCD: u8 = 0x01;
const PLF_ACD: u8 = 0x02;
const PLF_KME: u8 = 0x03;
const PLF_KME_MISMATCHED: u8 = 0x04;
const CTB_WRAPPER: u8 = 0x12;
const MCD_SIZE: usize = 24;
const ACD_SIZE: usize = 4;
const KME_SIZE: usize = 14;
const TB_DELTA_SIZE: usize = 18;
const CTB_DATA_OVERHEAD: usize = 112;
const MAX_BITMAP_SIZE: i32 = 65_576;

/// Parse exactly one complete `Tcg` payload.
pub fn parse_bytes<'a>(data: &'a [u8]) -> Result<CommandBars<'a>> {
    let mut reader = Reader::new(data);
    let version = reader.u8("Tcg nTcgVer")?;
    if version != TCG_VERSION {
        return Err(corrupted("Tcg nTcgVer must be 0xFF"));
    }

    let mut entries = Vec::new();
    let mut terminated = false;
    while !reader.is_empty() {
        let tag = reader.peek_u8("Tcg255 record tag")?;
        if tag == TCG_TERMINATOR {
            reader.advance(1, "Tcg255 terminator")?;
            terminated = true;
            break;
        }
        validate_count(entries.len(), "Tcg255 record")?;
        let entry = match tag {
            PLF_MCD => Entry::MacroCommands(parse_macro_commands(&mut reader)?),
            PLF_ACD => Entry::AllocatedCommands(parse_allocated_commands(&mut reader)?),
            PLF_KME | PLF_KME_MISMATCHED => Entry::KeyMaps(parse_key_maps(&mut reader)?),
            CTB_WRAPPER => Entry::Toolbar(parse_toolbar_wrapper(&mut reader)?),
            0x10 | 0x11 => {
                return Err(unsupported(
                    "TcgSttbf/MacroNames records are not decoded because their nested string-table boundaries are not safely inferable in this bounded owner",
                ));
            },
            other => {
                return Err(unsupported(format!(
                    "Tcg255 record tag 0x{other:02X} is not safely bounded"
                )));
            },
        };
        entries.push(entry);
    }
    if !terminated {
        return Err(corrupted("Tcg255 is missing chTerminator"));
    }
    reader.finish()?;

    let value = CommandBars {
        version,
        entries,
        terminator: TCG_TERMINATOR,
    };
    value.validate()?;
    Ok(value)
}

/// Serialize one complete `Tcg` payload.
pub fn to_bytes(value: &CommandBars<'_>) -> Result<Vec<u8>> {
    value.validate()?;
    let mut output = Vec::new();
    output.push(value.version);
    for entry in &value.entries {
        match entry {
            Entry::MacroCommands(commands) => encode_macro_commands(&mut output, commands)?,
            Entry::AllocatedCommands(commands) => encode_allocated_commands(&mut output, commands)?,
            Entry::KeyMaps(maps) => encode_key_maps(&mut output, maps)?,
            Entry::Toolbar(wrapper) => encode_toolbar_wrapper(&mut output, wrapper)?,
        }
    }
    output.push(value.terminator);
    Ok(output)
}

impl<'a> XString<'a> {
    /// Parse one `Xst` and return the number of bytes consumed.
    pub fn parse_prefix(data: &'a [u8]) -> Result<(Self, usize)> {
        let mut reader = Reader::new(data);
        let count = usize::from(reader.u16("Xst cch")?);
        let byte_len = count
            .checked_mul(2)
            .ok_or_else(|| corrupted("Xst byte length overflows"))?;
        let encoded = reader.take(byte_len, "Xst rgtchar")?;
        Ok((Self::from_wire(encoded)?, reader.position()))
    }

    /// Parse exactly one `Xst`.
    pub fn parse(data: &'a [u8]) -> Result<Self> {
        let (value, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(corrupted("Xst has trailing bytes"));
        }
        Ok(value)
    }

    /// Serialize the two-byte length and the exact UTF-16 payload.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut output = Vec::with_capacity(2 + self.encoded.len());
        output.extend_from_slice(&(self.len() as u16).to_le_bytes());
        output.extend_from_slice(&self.encoded);
        output
    }
}

fn parse_macro_commands(reader: &mut Reader<'_>) -> Result<MacroCommands> {
    reader.expect(PLF_MCD, "PlfMcd tag")?;
    let count = count_from_i32(reader.i32("PlfMcd iMac")?, "PlfMcd iMac")?;
    validate_count(count, "PlfMcd iMac")?;
    let required = count
        .checked_mul(MCD_SIZE)
        .ok_or_else(|| corrupted("PlfMcd size overflows"))?;
    reader.ensure(required, "PlfMcd rgmcd")?;
    let mut commands = Vec::with_capacity(count);
    for _ in 0..count {
        let reserved1 = reader.u8("Mcd reserved1")?;
        let reserved2 = reader.u8("Mcd reserved2")?;
        let macro_name_index = reader.u16("Mcd ibst")?;
        let command_name_index = reader.u16("Mcd ibstName")?;
        commands.push(MacroCommand::new(
            macro_name_index,
            command_name_index,
            reserved1,
            reserved2,
            reader.u16("Mcd reserved3")?,
            reader.u32("Mcd reserved4")?,
            reader.u32("Mcd reserved5")?,
            reader.u32("Mcd reserved6")?,
            reader.u32("Mcd reserved7")?,
        ));
    }
    MacroCommands::new(commands)
}

fn parse_allocated_commands(reader: &mut Reader<'_>) -> Result<AllocatedCommands> {
    reader.expect(PLF_ACD, "PlfAcd tag")?;
    let count = count_from_i32(reader.i32("PlfAcd iMac")?, "PlfAcd iMac")?;
    validate_count(count, "PlfAcd iMac")?;
    let required = count
        .checked_mul(ACD_SIZE)
        .ok_or_else(|| corrupted("PlfAcd size overflows"))?;
    reader.ensure(required, "PlfAcd rgacd")?;
    let mut commands = Vec::with_capacity(count);
    for _ in 0..count {
        let argument_index = reader.u16("Acd ibst")?;
        let packed = reader.u16("Acd fciBasedOn")?;
        commands.push(AllocatedCommand::new(
            argument_index,
            packed & 0x1FFF,
            (packed >> 13) as u8,
        ));
    }
    AllocatedCommands::new(commands)
}

fn parse_key_maps(reader: &mut Reader<'_>) -> Result<KeyMaps> {
    let kind = KeyMapKind::from_raw(reader.peek_u8("PlfKme ch")?)?;
    reader.advance(1, "PlfKme ch")?;
    let count = count_from_i32(reader.i32("PlfKme iMac")?, "PlfKme iMac")?;
    validate_count(count, "PlfKme iMac")?;
    let required = count
        .checked_mul(KME_SIZE)
        .ok_or_else(|| corrupted("PlfKme size overflows"))?;
    reader.ensure(required, "PlfKme rgkme")?;
    let mut entries = Vec::with_capacity(count);
    for _ in 0..count {
        let reserved1 = reader.u16("Kme reserved1")?;
        let reserved2 = reader.u16("Kme reserved2")?;
        let primary_key = reader.u16("Kme kcm1")?;
        let secondary_key = reader.u16("Kme kcm2")?;
        let action = Action::from_raw(reader.u16("Kme kt")?);
        let parameter = reader.u32("Kme param")?;
        entries.push(KeyMap::new(
            primary_key,
            secondary_key,
            action,
            parameter,
            reserved1,
            reserved2,
        ));
    }
    KeyMaps::new(kind, entries)
}

fn parse_toolbar_wrapper<'a>(reader: &mut Reader<'a>) -> Result<ToolbarWrapper<'a>> {
    reader.expect(CTB_WRAPPER, "CTBWRAPPER tag")?;
    let reserved1 = reader.u8("CTBWRAPPER reserved1")?;
    let reserved2 = reader.u16("CTBWRAPPER reserved2")?;
    let reserved3 = reader.u8("CTBWRAPPER reserved3")?;
    let reserved4 = reader.u16("CTBWRAPPER reserved4")?;
    let reserved5 = reader.u16("CTBWRAPPER reserved5")?;
    let cb_tbd = reader.u16("CTBWRAPPER cbTBD")?;
    let customization_count = positive_count(reader.i16("CTBWRAPPER cCust")?, "CTBWRAPPER cCust")?;
    let controls_len = nonnegative_i32(reader.i32("CTBWRAPPER cbDTBC")?, "cbDTBC")?;
    let controls = reader.take(controls_len, "CTBWRAPPER rtbdc")?;
    let delta_controls = control::parse_many(controls)?;

    let mut customizations = Vec::with_capacity(customization_count);
    for _ in 0..customization_count {
        customizations.push(parse_customization(reader, customization_count)?);
    }
    let value = ToolbarWrapper {
        reserved1,
        reserved2,
        reserved3,
        reserved4,
        reserved5,
        cb_tbd,
        toolbar_controls: Cow::Borrowed(controls),
        delta_controls,
        customizations,
    };
    value.validate()?;
    Ok(value)
}

fn parse_customization<'a>(
    reader: &mut Reader<'a>,
    customization_count: usize,
) -> Result<Customization<'a>> {
    let toolbar_id = reader.u32("Customization tbidForTBD")?;
    let reserved = reader.u16("Customization reserved1")?;
    let delta_count = reader.u16("Customization ctbds")?;
    let data = if toolbar_id == 0 {
        if delta_count != 0 {
            return Err(corrupted("Customization CTB has nonzero ctbds"));
        }
        let (toolbar, consumed) = parse_toolbar(reader.remaining(), customization_count)?;
        reader.advance(consumed, "Customization CTB")?;
        CustomizationData::Toolbar(toolbar)
    } else {
        let count = usize::from(delta_count);
        validate_count(count, "Customization ctbds")?;
        let required = count
            .checked_mul(TB_DELTA_SIZE)
            .ok_or_else(|| corrupted("Customization delta size overflows"))?;
        reader.ensure(required, "CustomizationData TBDelta array")?;
        let mut deltas = Vec::with_capacity(count);
        for _ in 0..count {
            deltas.push(parse_delta(reader)?);
        }
        CustomizationData::Deltas(deltas)
    };
    let value = Customization {
        toolbar_id,
        reserved,
        delta_count,
        data,
    };
    validate_count(customization_count, "CTBWRAPPER cCust")?;
    Ok(value)
}

fn parse_toolbar<'a>(data: &'a [u8], customization_count: usize) -> Result<(Toolbar<'a>, usize)> {
    let (name, name_size) = XString::parse_prefix(data)?;
    let mut reader = Reader::new(&data[name_size..]);
    let cb_tbd = nonnegative_i32(reader.i32("CTB cbTBData")?, "cbTBData")?;
    if cb_tbd < CTB_DATA_OVERHEAD {
        return Err(corrupted("CTB cbTBData is smaller than its fixed fields"));
    }
    let body_len = cb_tbd
        .checked_sub(4)
        .ok_or_else(|| corrupted("CTB cbTBData underflows"))?;
    let body = reader.take(body_len, "CTB TB data")?;
    let expected_toolbar_len = cb_tbd
        .checked_sub(CTB_DATA_OVERHEAD)
        .ok_or_else(|| corrupted("CTB toolbar length underflows"))?;
    let (header, header_len) = Header::parse_prefix(body)
        .map_err(|error| corrupted(format!("invalid CTB TB header: {error}")))?;
    if header_len != expected_toolbar_len {
        return Err(corrupted(format!(
            "CTB cbTBData declares {expected_toolbar_len} TB bytes, decoded {header_len}"
        )));
    }

    let mut body_reader = Reader::new(&body[header_len..]);
    let visual_data = body_reader.array::<100>("CTB rVisualData")?;
    let toolbar_index = body_reader.i32("CTB iWCTB")?;
    let reserved = body_reader.u16("CTB reserved")?;
    let unused = body_reader.u16("CTB unused")?;
    body_reader.finish()?;

    let control_count = nonnegative_i32(reader.i32("CTB cCtls")?, "CTB cCtls")?;
    validate_count(control_count, "CTB cCtls")?;
    let mut controls = Vec::with_capacity(control_count);
    for _ in 0..control_count {
        controls.push(control::parse_one(&mut reader)?);
    }
    let control_bytes = reader.position();

    let value = Toolbar {
        name,
        header,
        visual_data,
        toolbar_index,
        reserved,
        unused,
        control_count: i32::try_from(control_count)
            .map_err(|_| corrupted("CTB cCtls exceeds i32::MAX"))?,
        controls,
    };
    validate_count(customization_count, "CTB cCust")?;
    Ok((value, name_size + control_bytes))
}

fn parse_delta(reader: &mut Reader<'_>) -> Result<ToolbarDelta> {
    let packed = reader.u16("TBDelta operation")?;
    let operation = Operation::from_raw((packed & 0x0003) as u8)?;
    let at_end = packed & 0x0004 != 0;
    let reserved_flags = ((packed >> 3) & 0x001F) as u8;
    let control_index = (packed >> 8) as u8;
    Ok(ToolbarDelta::new(
        operation,
        at_end,
        reserved_flags,
        control_index,
        reader.u32("TBDelta cidNext")?,
        reader.u32("TBDelta cid")?,
        reader.u32("TBDelta fc")?,
        reader.u16("TBDelta state")?,
        reader.u16("TBDelta cbTBC")?,
    ))
}

fn encode_macro_commands(output: &mut Vec<u8>, value: &MacroCommands) -> Result<()> {
    output.push(PLF_MCD);
    output.extend_from_slice(&count_i32(value.commands.len(), "PlfMcd iMac")?.to_le_bytes());
    for command in &value.commands {
        output.push(command.reserved1);
        output.push(command.reserved2);
        output.extend_from_slice(&command.macro_name_index.to_le_bytes());
        output.extend_from_slice(&command.command_name_index.to_le_bytes());
        output.extend_from_slice(&command.reserved3.to_le_bytes());
        output.extend_from_slice(&command.reserved4.to_le_bytes());
        output.extend_from_slice(&command.reserved5.to_le_bytes());
        output.extend_from_slice(&command.reserved6.to_le_bytes());
        output.extend_from_slice(&command.reserved7.to_le_bytes());
    }
    Ok(())
}

fn encode_allocated_commands(output: &mut Vec<u8>, value: &AllocatedCommands) -> Result<()> {
    output.push(PLF_ACD);
    output.extend_from_slice(&count_i32(value.commands.len(), "PlfAcd iMac")?.to_le_bytes());
    for command in &value.commands {
        output.extend_from_slice(&command.argument_index.to_le_bytes());
        let packed = (command.command & 0x1FFF) | (u16::from(command.flags & 0x07) << 13);
        output.extend_from_slice(&packed.to_le_bytes());
    }
    Ok(())
}

fn encode_key_maps(output: &mut Vec<u8>, value: &KeyMaps) -> Result<()> {
    output.push(value.kind.raw());
    output.extend_from_slice(&count_i32(value.entries.len(), "PlfKme iMac")?.to_le_bytes());
    for entry in &value.entries {
        output.extend_from_slice(&entry.reserved1.to_le_bytes());
        output.extend_from_slice(&entry.reserved2.to_le_bytes());
        output.extend_from_slice(&entry.primary_key.to_le_bytes());
        output.extend_from_slice(&entry.secondary_key.to_le_bytes());
        output.extend_from_slice(&entry.action.raw().to_le_bytes());
        output.extend_from_slice(&entry.parameter.to_le_bytes());
    }
    Ok(())
}

fn encode_toolbar_wrapper(output: &mut Vec<u8>, value: &ToolbarWrapper<'_>) -> Result<()> {
    value.validate()?;
    output.push(CTB_WRAPPER);
    output.push(value.reserved1);
    output.extend_from_slice(&value.reserved2.to_le_bytes());
    output.push(value.reserved3);
    output.extend_from_slice(&value.reserved4.to_le_bytes());
    output.extend_from_slice(&value.reserved5.to_le_bytes());
    output.extend_from_slice(&value.cb_tbd.to_le_bytes());
    output.extend_from_slice(
        &i16::try_from(value.customizations.len())
            .map_err(|_| corrupted("CTBWRAPPER cCust exceeds i16::MAX"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(
        &u32::try_from(value.toolbar_controls.len())
            .map_err(|_| corrupted("CTBWRAPPER cbDTBC exceeds u32::MAX"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(&value.toolbar_controls);
    for customization in &value.customizations {
        encode_customization(output, customization)?;
    }
    Ok(())
}

fn encode_customization(output: &mut Vec<u8>, value: &Customization<'_>) -> Result<()> {
    output.extend_from_slice(&value.toolbar_id.to_le_bytes());
    output.extend_from_slice(&value.reserved.to_le_bytes());
    output.extend_from_slice(&value.delta_count.to_le_bytes());
    match &value.data {
        CustomizationData::Toolbar(toolbar) => encode_toolbar(output, toolbar)?,
        CustomizationData::Deltas(deltas) => {
            for delta in deltas {
                encode_delta(output, delta);
            }
        },
    }
    Ok(())
}

fn encode_toolbar(output: &mut Vec<u8>, value: &Toolbar<'_>) -> Result<()> {
    value.validate(1)?;
    let toolbar = value.header.to_bytes();
    let cb_tbd = toolbar
        .len()
        .checked_add(CTB_DATA_OVERHEAD)
        .ok_or_else(|| corrupted("CTB cbTBData overflows"))?;
    let cb_tbd = i32::try_from(cb_tbd).map_err(|_| corrupted("CTB cbTBData exceeds i32::MAX"))?;
    output.extend_from_slice(&value.name.to_bytes());
    output.extend_from_slice(&cb_tbd.to_le_bytes());
    output.extend_from_slice(&toolbar);
    output.extend_from_slice(&value.visual_data);
    output.extend_from_slice(&value.toolbar_index.to_le_bytes());
    output.extend_from_slice(&value.reserved.to_le_bytes());
    output.extend_from_slice(&value.unused.to_le_bytes());
    output.extend_from_slice(
        &i32::try_from(value.controls.len())
            .map_err(|_| corrupted("CTB cCtls exceeds i32::MAX"))?
            .to_le_bytes(),
    );
    for control in &value.controls {
        encode_control(output, control)?;
    }
    Ok(())
}

/// Serialize one complete DOC toolbar-control record.
pub fn to_control_bytes(value: &Control<'_>) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    encode_control(&mut output, value)?;
    Ok(output)
}

impl<'a> Control<'a> {
    /// Serialize one complete DOC toolbar-control record.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_control_bytes(self)
    }
}

impl<'a> ToolbarWrapper<'a> {
    /// Attach typed rtbdc controls while retaining their serialized bytes.
    pub fn with_delta_controls(mut self, controls: Vec<Control<'a>>) -> Result<Self> {
        validate_count(controls.len(), "CTBWRAPPER rtbdc")?;
        let mut encoded = Vec::new();
        for control in &controls {
            encode_control(&mut encoded, control)?;
        }
        self.toolbar_controls = Cow::Owned(encoded);
        self.delta_controls = controls;
        self.validate()?;
        Ok(self)
    }
}

fn encode_control(output: &mut Vec<u8>, value: &Control<'_>) -> Result<()> {
    value.validate()?;
    output.extend_from_slice(&value.header.to_bytes());
    if let Some(command) = value.command {
        output.extend_from_slice(&command.raw().to_le_bytes());
    }
    if let Some(data) = &value.data {
        output.extend_from_slice(&data.to_bytes());
    }
    Ok(())
}

fn encode_delta(output: &mut Vec<u8>, value: &ToolbarDelta) {
    let packed = u16::from(value.operation.raw())
        | (u16::from(value.at_end) << 2)
        | (u16::from(value.reserved_flags & 0x1F) << 3)
        | (u16::from(value.control_index) << 8);
    output.extend_from_slice(&packed.to_le_bytes());
    output.extend_from_slice(&value.next_command.to_le_bytes());
    output.extend_from_slice(&value.command.to_le_bytes());
    output.extend_from_slice(&value.file_offset.to_le_bytes());
    output.extend_from_slice(&value.state.to_le_bytes());
    output.extend_from_slice(&value.control_size.to_le_bytes());
}

fn count_from_i32(value: i32, field: &str) -> Result<usize> {
    usize::try_from(value).map_err(|_| corrupted(format!("{field} must be nonnegative")))
}

fn count_i32(value: usize, field: &str) -> Result<i32> {
    validate_count(value, field)?;
    i32::try_from(value).map_err(|_| corrupted(format!("{field} exceeds i32::MAX")))
}

fn positive_count(value: i16, field: &str) -> Result<usize> {
    if value <= 0 {
        return Err(corrupted(format!("{field} must be greater than zero")));
    }
    Ok(value as usize)
}

fn nonnegative_i32(value: i32, field: &str) -> Result<usize> {
    usize::try_from(value).map_err(|_| corrupted(format!("{field} must be nonnegative")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn unsupported(message: impl Into<String>) -> PackageError {
    PackageError::InvalidFormat(format!(
        "unsupported DOC command-bar record: {}",
        message.into()
    ))
}

struct Reader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    const fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn position(&self) -> usize {
        self.offset
    }

    fn remaining(&self) -> &'a [u8] {
        &self.data[self.offset..]
    }

    fn is_empty(&self) -> bool {
        self.offset == self.data.len()
    }

    fn take(&mut self, length: usize, field: &str) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
        let value = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
        self.offset = end;
        Ok(value)
    }

    fn ensure(&self, length: usize, field: &str) -> Result<()> {
        self.offset
            .checked_add(length)
            .filter(|end| *end <= self.data.len())
            .ok_or_else(|| corrupted(format!("{field} is truncated")))
            .map(|_| ())
    }

    fn advance(&mut self, length: usize, field: &str) -> Result<()> {
        self.take(length, field).map(|_| ())
    }

    fn finish(&self) -> Result<()> {
        if self.is_empty() {
            Ok(())
        } else {
            Err(corrupted("command-bar record has trailing bytes"))
        }
    }

    fn peek_u8(&self, field: &str) -> Result<u8> {
        self.data
            .get(self.offset)
            .copied()
            .ok_or_else(|| corrupted(format!("{field} is truncated")))
    }

    fn expect(&mut self, expected: u8, field: &str) -> Result<()> {
        let actual = self.u8(field)?;
        if actual == expected {
            Ok(())
        } else {
            Err(corrupted(format!(
                "{field} must be 0x{expected:02X}, got 0x{actual:02X}"
            )))
        }
    }

    fn u8(&mut self, field: &str) -> Result<u8> {
        self.take(1, field).map(|bytes| bytes[0])
    }

    fn u16(&mut self, field: &str) -> Result<u16> {
        let bytes = self.take(2, field)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn i16(&mut self, field: &str) -> Result<i16> {
        Ok(self.u16(field)? as i16)
    }

    fn u32(&mut self, field: &str) -> Result<u32> {
        let bytes = self.take(4, field)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn i32(&mut self, field: &str) -> Result<i32> {
        Ok(self.u32(field)? as i32)
    }

    fn array<const N: usize>(&mut self, field: &str) -> Result<[u8; N]> {
        self.take(N, field).and_then(|bytes| {
            bytes
                .try_into()
                .map_err(|_| corrupted(format!("{field} has an invalid length")))
        })
    }
}

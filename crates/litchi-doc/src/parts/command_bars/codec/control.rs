//! Boundary-aware TBC codecs.

use super::super::model::*;
use super::super::validation::validate_count;
use super::{MAX_BITMAP_SIZE, Reader, corrupted, unsupported};
use crate::package::Result;
use litchi_ole_common::toolbar::{ControlHeader, ControlType, Data, GeneralInfo, WString};

pub(super) fn parse_many<'a>(data: &'a [u8]) -> Result<Vec<Control<'a>>> {
    let mut reader = Reader::new(data);
    let mut controls = Vec::new();
    while !reader.is_empty() {
        validate_count(controls.len(), "TBC control")?;
        controls.push(parse_one(&mut reader)?);
    }
    Ok(controls)
}

pub(super) fn parse_one<'a>(reader: &mut Reader<'a>) -> Result<Control<'a>> {
    let (header, header_len) = ControlHeader::parse_prefix(reader.remaining())
        .map_err(|error| corrupted(format!("invalid TBCHeader: {error}")))?;
    reader.advance(header_len, "TBCHeader")?;

    let command = if header.control_id() == 0x0001 || header.control_id() == 0x1051 {
        None
    } else {
        Some(CommandId::new(reader.u32("TBC Cid")?)?)
    };

    let data = if matches!(header.control_type(), ControlType::ActiveX) {
        None
    } else {
        let remaining = reader.remaining();
        let specific_len = specific_length(remaining, header.control_type(), header.control_id())?;
        let (data, consumed) = Data::parse_prefix(remaining, specific_len)
            .map_err(|error| corrupted(format!("invalid TBCData: {error}")))?;
        reader.advance(consumed, "TBCData")?;
        Some(data)
    };

    Control::new(header, command, data)
}

fn specific_length(data: &[u8], control_type: ControlType, control_id: u16) -> Result<usize> {
    let (_, general_len) = GeneralInfo::parse_prefix(data)
        .map_err(|error| corrupted(format!("invalid TBCGeneralInfo: {error}")))?;
    let specific = &data[general_len..];
    let specific_len = match control_type {
        ControlType::Button | ControlType::ExpandingGrid => button_specific_length(specific)?,
        ControlType::Popup
        | ControlType::ButtonPopup
        | ControlType::SplitButtonPopup
        | ControlType::SplitButtonMruPopup => menu_specific_length(specific)?,
        ControlType::Edit
        | ControlType::ComboBox
        | ControlType::GraphicCombo
        | ControlType::DropDown
        | ControlType::SplitDropDown
        | ControlType::GraphicDropDown => {
            if control_id == 0x0001 {
                combo_specific_length(specific)?
            } else {
                0
            }
        },
        ControlType::OcxDropDown
        | ControlType::Label
        | ControlType::Grid
        | ControlType::Gauge
        | ControlType::Pane => 0,
        ControlType::ActiveX | ControlType::Unknown(_) => {
            return Err(unsupported(format!(
                "TBCData is not defined for control type 0x{:02X}",
                control_type.raw()
            )));
        },
    };
    Ok(specific_len)
}

fn button_specific_length(data: &[u8]) -> Result<usize> {
    let mut reader = Reader::new(data);
    let flags = reader.u8("TBCBSpecific bFlags")?;
    if flags & (1 << 3) != 0 {
        scan_bitmap(&mut reader, "TBCBSpecific icon")?;
        scan_bitmap(&mut reader, "TBCBSpecific iconMask")?;
    }
    if flags & (1 << 4) != 0 {
        reader.advance(2, "TBCBSpecific iBtnFace")?;
    }
    if flags & (1 << 2) != 0 {
        scan_wstring(&mut reader, "TBCBSpecific wstrAcc")?;
    }
    Ok(reader.position())
}

fn menu_specific_length(data: &[u8]) -> Result<usize> {
    let mut reader = Reader::new(data);
    let toolbar_id = reader.u32("TBCMenuSpecific tbid")?;
    if toolbar_id == 1 {
        scan_wstring(&mut reader, "TBCMenuSpecific name")?;
    }
    Ok(reader.position())
}

fn combo_specific_length(data: &[u8]) -> Result<usize> {
    let mut reader = Reader::new(data);
    let item_count = positive_count(reader.i16("TBCCDData cwstrItems")?, "TBCCDData cwstrItems")?;
    validate_count(item_count, "TBCCDData cwstrItems")?;
    for _ in 0..item_count {
        scan_wstring(&mut reader, "TBCCDData wstrList")?;
    }
    let mru_count = reader.i16("TBCCDData cwstrMRU")?;
    if mru_count < -1 {
        return Err(corrupted("TBCCDData cwstrMRU must be -1 or nonnegative"));
    }
    let selected = reader.i16("TBCCDData iSel")?;
    if selected < -1
        || (selected >= 0 && usize::try_from(selected).unwrap_or(usize::MAX) >= item_count)
    {
        return Err(corrupted("TBCCDData iSel is outside wstrList"));
    }
    if reader.i16("TBCCDData cLines")? < 0 {
        return Err(corrupted("TBCCDData cLines must be nonnegative"));
    }
    if reader.i16("TBCCDData dxWidth")? < -1 {
        return Err(corrupted("TBCCDData dxWidth must be -1 or nonnegative"));
    }
    scan_wstring(&mut reader, "TBCCDData wstrEdit")?;
    Ok(reader.position())
}

fn scan_wstring(reader: &mut Reader<'_>, field: &str) -> Result<()> {
    let (_, consumed) = WString::parse_prefix(reader.remaining())
        .map_err(|error| corrupted(format!("{field} is invalid: {error}")))?;
    reader.advance(consumed, field)
}

fn scan_bitmap(reader: &mut Reader<'_>, field: &str) -> Result<()> {
    let cb_dib = reader.i32(&format!("{field} cbDIB"))?;
    if !(40..=MAX_BITMAP_SIZE).contains(&cb_dib) {
        return Err(corrupted(format!(
            "{field} cbDIB must be between 40 and {MAX_BITMAP_SIZE}"
        )));
    }
    let after_length =
        usize::try_from(cb_dib - 10).map_err(|_| corrupted(format!("{field} cbDIB underflows")))?;
    if after_length < 30 {
        return Err(corrupted(format!("{field} bitmap header is truncated")));
    }
    reader.advance(after_length, field)
}

fn positive_count(value: i16, field: &str) -> Result<usize> {
    if value <= 0 {
        return Err(corrupted(format!("{field} must be greater than zero")));
    }
    Ok(value as usize)
}

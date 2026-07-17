//! Typed outline and slide-sorter view-information records.

use super::package::{PptError, Result};
use super::records::PptRecord;
use super::view_info::{PowerPointRatio, PowerPointViewOrigin};
use crate::consts::PptRecordType;

const OUTLINE_VIEW_INFO_TYPE: u16 = 0x0407;
const SORTER_VIEW_INFO_TYPE: u16 = 0x0408;
const VIEW_INFO_ATOM_TYPE: u16 = 0x03FD;
const CONTAINER_HEADER_LEN: usize = 8;
const ATOM_HEADER_LEN: usize = 8;
const ATOM_DATA_LEN: usize = 52;
const ATOM_TOTAL_LEN: usize = ATOM_HEADER_LEN + ATOM_DATA_LEN;
const MAX_CONTAINER_DATA: usize = ATOM_TOTAL_LEN;

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn normalized_ratio(ratio: PowerPointRatio) -> (i64, i64) {
    let numerator = i64::from(ratio.numerator());
    let denominator = i64::from(ratio.denominator());
    if denominator < 0 {
        (-numerator, -denominator)
    } else {
        (numerator, denominator)
    }
}

fn validate_scale(x: PowerPointRatio, y: PowerPointRatio) -> Result<()> {
    let (x_numerator, x_denominator) = normalized_ratio(x);
    let (y_numerator, y_denominator) = normalized_ratio(y);
    if x_numerator * y_denominator != y_numerator * x_denominator {
        return Err(corrupted(
            "NoZoomViewInfoAtom x and y scale ratios must be equal",
        ));
    }
    if x_numerator * 5 < x_denominator || x_numerator > x_denominator {
        return Err(corrupted(
            "NoZoomViewInfoAtom scale must be between 20 and 100 percent",
        ));
    }
    Ok(())
}

/// The presentation view represented by an outline/sorter view container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointNonZoomViewKind {
    Outline,
    Sorter,
}

impl PowerPointNonZoomViewKind {
    fn record_type(self) -> (PptRecordType, u16) {
        match self {
            Self::Outline => (PptRecordType::OutlineViewInfo, OUTLINE_VIEW_INFO_TYPE),
            Self::Sorter => (PptRecordType::SorterViewInfo, SORTER_VIEW_INFO_TYPE),
        }
    }

    fn from_raw(record_type: u16) -> Result<Self> {
        match record_type {
            OUTLINE_VIEW_INFO_TYPE => Ok(Self::Outline),
            SORTER_VIEW_INFO_TYPE => Ok(Self::Sorter),
            _ => Err(corrupted(format!(
                "Unexpected outline/sorter view record type {record_type}"
            ))),
        }
    }
}

/// Fixed-size `NoZoomViewInfoAtom` data with ignored bytes preserved losslessly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointNoZoomViewInfo {
    x_scale: PowerPointRatio,
    y_scale: PowerPointRatio,
    ignored1: [u8; 24],
    origin: PowerPointViewOrigin,
    ignored2: u8,
    draft_mode: bool,
    ignored3: [u8; 2],
}

impl PowerPointNoZoomViewInfo {
    pub fn new(
        x_scale: PowerPointRatio,
        y_scale: PowerPointRatio,
        ignored1: [u8; 24],
        origin: PowerPointViewOrigin,
        ignored2: u8,
        draft_mode: bool,
        ignored3: [u8; 2],
    ) -> Result<Self> {
        validate_scale(x_scale, y_scale)?;
        Ok(Self {
            x_scale,
            y_scale,
            ignored1,
            origin,
            ignored2,
            draft_mode,
            ignored3,
        })
    }

    pub fn x_scale(&self) -> PowerPointRatio { self.x_scale }
    pub fn y_scale(&self) -> PowerPointRatio { self.y_scale }
    pub fn ignored1(&self) -> &[u8; 24] { &self.ignored1 }
    pub fn origin(&self) -> PowerPointViewOrigin { self.origin }
    pub fn ignored2(&self) -> u8 { self.ignored2 }
    pub fn draft_mode(&self) -> bool { self.draft_mode }
    pub fn ignored3(&self) -> &[u8; 2] { &self.ignored3 }

    fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != ATOM_DATA_LEN {
            return Err(corrupted(format!(
                "NoZoomViewInfoAtom payload length must be {ATOM_DATA_LEN}, got {}",
                data.len()
            )));
        }
        let x_scale = PowerPointRatio::new(read_i32(data, 0), read_i32(data, 4))?;
        let y_scale = PowerPointRatio::new(read_i32(data, 8), read_i32(data, 12))?;
        let ignored1 = data[16..40].try_into().unwrap();
        let origin = PowerPointViewOrigin::new(read_i32(data, 40), read_i32(data, 44));
        let draft_mode = match data[49] {
            0 => false,
            1 => true,
            value => {
                return Err(corrupted(format!(
                    "NoZoomViewInfoAtom fDraftMode must be 0 or 1, got {value}"
                )))
            }
        };
        Self::new(
            x_scale,
            y_scale,
            ignored1,
            origin,
            data[48],
            draft_mode,
            data[50..52].try_into().unwrap(),
        )
    }

    fn write_atom(&self, bytes: &mut Vec<u8>) -> Result<()> {
        validate_scale(self.x_scale, self.y_scale)?;
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&VIEW_INFO_ATOM_TYPE.to_le_bytes());
        bytes.extend_from_slice(&(ATOM_DATA_LEN as u32).to_le_bytes());
        bytes.extend_from_slice(&self.x_scale.numerator().to_le_bytes());
        bytes.extend_from_slice(&self.x_scale.denominator().to_le_bytes());
        bytes.extend_from_slice(&self.y_scale.numerator().to_le_bytes());
        bytes.extend_from_slice(&self.y_scale.denominator().to_le_bytes());
        bytes.extend_from_slice(&self.ignored1);
        bytes.extend_from_slice(&self.origin.x().to_le_bytes());
        bytes.extend_from_slice(&self.origin.y().to_le_bytes());
        bytes.push(self.ignored2);
        bytes.push(u8::from(self.draft_mode));
        bytes.extend_from_slice(&self.ignored3);
        Ok(())
    }
}

/// An `OutlineViewInfoContainer` or `SorterViewInfoContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointOutlineSorterViewInfo {
    kind: PowerPointNonZoomViewKind,
    view_info: Option<PowerPointNoZoomViewInfo>,
}

impl PowerPointOutlineSorterViewInfo {
    pub const fn new(
        kind: PowerPointNonZoomViewKind,
        view_info: Option<PowerPointNoZoomViewInfo>,
    ) -> Self {
        Self { kind, view_info }
    }

    pub fn kind(&self) -> PowerPointNonZoomViewKind { self.kind }
    pub fn view_info(&self) -> Option<&PowerPointNoZoomViewInfo> { self.view_info.as_ref() }

    /// Parse a complete container record, including its eight-byte record header.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() < CONTAINER_HEADER_LEN {
            return Err(corrupted("Truncated outline/sorter view record header"));
        }
        let version_instance = read_u16(bytes, 0);
        let version = version_instance & 0xF;
        let instance = version_instance >> 4;
        let record_type = read_u16(bytes, 2);
        let declared = usize::try_from(read_u32(bytes, 4))
            .map_err(|_| corrupted("Outline/sorter view length does not fit memory"))?;
        let total = CONTAINER_HEADER_LEN
            .checked_add(declared)
            .ok_or_else(|| corrupted("Outline/sorter view length overflow"))?;
        if total != bytes.len() {
            return Err(corrupted("Outline/sorter view declared length mismatch"));
        }
        Self::parse_container(version, instance, record_type, &bytes[CONTAINER_HEADER_LEN..])
    }

    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        let declared = usize::try_from(record.data_length)
            .map_err(|_| corrupted("Outline/sorter view length does not fit memory"))?;
        if declared != record.data.len() {
            return Err(corrupted("Outline/sorter view declared length mismatch"));
        }
        let kind = PowerPointNonZoomViewKind::from_raw(record.record_type_raw)?;
        let (expected_type, _) = kind.record_type();
        if record.record_type != expected_type {
            return Err(corrupted("Outline/sorter view mapped record type mismatch"));
        }
        Self::parse_container(record.version, record.instance, record.record_type_raw, &record.data)
    }

    fn parse_container(
        version: u16,
        instance: u16,
        record_type: u16,
        data: &[u8],
    ) -> Result<Self> {
        let kind = PowerPointNonZoomViewKind::from_raw(record_type)?;
        if version != 0xF || instance != 1 {
            return Err(corrupted(
                "Outline/sorter view header must have version 0xF and instance 1",
            ));
        }
        if data.len() > MAX_CONTAINER_DATA {
            return Err(corrupted("Outline/sorter view container exceeds 60-byte cap"));
        }
        if data.is_empty() {
            return Ok(Self::new(kind, None));
        }
        if data.len() != ATOM_TOTAL_LEN {
            return Err(corrupted(
                "Outline/sorter view container must be empty or contain one 60-byte atom",
            ));
        }
        let version_instance = read_u16(data, 0);
        let atom_version = version_instance & 0xF;
        let atom_instance = version_instance >> 4;
        let atom_type = read_u16(data, 2);
        let atom_length = usize::try_from(read_u32(data, 4))
            .map_err(|_| corrupted("NoZoomViewInfoAtom length does not fit memory"))?;
        if atom_version != 0 || atom_instance != 0 || atom_type != VIEW_INFO_ATOM_TYPE {
            return Err(corrupted("Invalid NoZoomViewInfoAtom record header"));
        }
        if atom_length != ATOM_DATA_LEN {
            return Err(corrupted(format!(
                "NoZoomViewInfoAtom declared length must be {ATOM_DATA_LEN}, got {atom_length}"
            )));
        }
        let view_info = PowerPointNoZoomViewInfo::parse(&data[ATOM_HEADER_LEN..])?;
        Ok(Self::new(kind, Some(view_info)))
    }

    /// Serialize a complete container record deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let (_, record_type) = self.kind.record_type();
        let data_len = if self.view_info.is_some() { ATOM_TOTAL_LEN } else { 0 };
        let mut bytes = Vec::with_capacity(CONTAINER_HEADER_LEN + data_len);
        bytes.extend_from_slice(&0x001Fu16.to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&(data_len as u32).to_le_bytes());
        if let Some(view_info) = &self.view_info {
            view_info.write_atom(&mut bytes)?;
        }
        Ok(bytes)
    }
}

/// Outline and slide-sorter view settings exposed by a presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointOutlineSorterViewInformation {
    outline: Option<PowerPointOutlineSorterViewInfo>,
    sorter: Option<PowerPointOutlineSorterViewInfo>,
}

impl PowerPointOutlineSorterViewInformation {
    pub fn outline(&self) -> Option<&PowerPointOutlineSorterViewInfo> { self.outline.as_ref() }
    pub fn sorter(&self) -> Option<&PowerPointOutlineSorterViewInfo> { self.sorter.as_ref() }

    pub(crate) fn parse_records(records: &[&PptRecord]) -> Result<Self> {
        let mut information = Self::default();
        for record in records {
            let slot = match record.record_type {
                PptRecordType::OutlineViewInfo => &mut information.outline,
                PptRecordType::SorterViewInfo => &mut information.sorter,
                _ => continue,
            };
            if slot.is_some() {
                return Err(corrupted("Duplicate outline/sorter view container"));
            }
            *slot = Some(PowerPointOutlineSorterViewInfo::parse_record(record)?);
        }
        Ok(information)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_OUTLINE: [u8; 68] = [
        0x1f, 0x00, 0x07, 0x04, 0x3c, 0, 0, 0, 0, 0, 0xfd, 0x03, 0x34, 0, 0, 0,
        0x21, 0, 0, 0, 0x64, 0, 0, 0, 0x21, 0, 0, 0, 0x64, 0, 0, 0, 0xa3, 0x8b,
        0x07, 0x30, 0xdc, 0xba, 0x12, 0, 0, 0, 0, 0, 0xfc, 0xe6, 0x29, 0x06, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0x12, 0,
    ];

    const POI_SORTER: [u8; 68] = [
        0x1f, 0x00, 0x08, 0x04, 0x3c, 0, 0, 0,
        0, 0, 0xfd, 0x03, 0x34, 0, 0, 0,
        0x64, 0, 0, 0, 0x64, 0, 0, 0,
        0x64, 0, 0, 0, 0x64, 0, 0, 0,
        0x6c, 0x42, 0xbb, 0x04, 0, 0, 0, 0,
        0xc8, 0xb7, 0x12, 0, 1, 0, 0, 0,
        0xa6, 0x17, 0, 0, 0x34, 0x0e, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0xff, 0xff,
    ];

    #[test]
    fn parses_and_round_trips_poi_reference_records() {
        let outline = PowerPointOutlineSorterViewInfo::from_bytes(&POI_OUTLINE).unwrap();
        assert_eq!(outline.kind(), PowerPointNonZoomViewKind::Outline);
        let view = outline.view_info().unwrap();
        assert_eq!((view.x_scale().numerator(), view.x_scale().denominator()), (33, 100));
        assert_eq!(view.x_scale(), view.y_scale());
        assert_eq!(view.origin(), PowerPointViewOrigin::new(0, 0));
        assert!(view.draft_mode());
        assert_eq!(outline.to_bytes().unwrap(), POI_OUTLINE);

        let sorter = PowerPointOutlineSorterViewInfo::from_bytes(&POI_SORTER).unwrap();
        assert_eq!(sorter.kind(), PowerPointNonZoomViewKind::Sorter);
        assert_eq!(sorter.view_info().unwrap().x_scale().numerator(), 100);
        assert!(!sorter.view_info().unwrap().draft_mode());
        assert_eq!(sorter.to_bytes().unwrap(), POI_SORTER);
    }

    #[test]
    fn accepts_and_serializes_optional_empty_container() {
        let bytes = [0x1f, 0, 0x07, 0x04, 0, 0, 0, 0];
        let parsed = PowerPointOutlineSorterViewInfo::from_bytes(&bytes).unwrap();
        assert!(parsed.view_info().is_none());
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_headers_and_lengths() {
        let mut bytes = POI_OUTLINE;
        bytes[0] = 0x0f;
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[4] = 59;
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[8] = 1;
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[12] = 51;
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&POI_OUTLINE[..67]).is_err());
    }

    #[test]
    fn rejects_invalid_scale_and_boolean_constraints() {
        let mut bytes = POI_OUTLINE;
        bytes[20..24].copy_from_slice(&0i32.to_le_bytes());
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[16..20].copy_from_slice(&19i32.to_le_bytes());
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[16..20].copy_from_slice(&101i32.to_le_bytes());
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[24..28].copy_from_slice(&34i32.to_le_bytes());
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
        let mut bytes = POI_OUTLINE;
        bytes[65] = 2;
        assert!(PowerPointOutlineSorterViewInfo::from_bytes(&bytes).is_err());
    }
}

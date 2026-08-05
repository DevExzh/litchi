use super::model::{
    PowerPointNormalViewSet, PowerPointNormalViewSetInfo, PowerPointNormalViewSetPayload,
    PowerPointNotesTextViewInfo, PowerPointViewBarState,
};
use crate::consts::PptRecordType;
use crate::non_zoom_view::PowerPointNoZoomViewInfo;
use crate::package::{PptError, Result};
use crate::records::record::PptRecord;
use crate::view_info::PowerPointRatio;

/// `RT_NormalViewSetInfo9` record type.
pub(super) const NORMAL_VIEW_SET_INFO_TYPE: u16 = 0x0414;
/// `RT_NormalViewSetInfo9Atom` record type.
pub(super) const NORMAL_VIEW_SET_INFO_ATOM_TYPE: u16 = 0x0415;
/// `RT_NotesTextViewInfo9` record type.
pub(super) const NOTES_TEXT_VIEW_INFO_TYPE: u16 = 0x0413;
/// Byte length of a `NormalViewSetInfo9Atom` payload.
const ATOM_LEN: usize = 20;
/// Flag bits defined by `NormalViewSetInfo9Atom`.
const KNOWN_FLAGS: u8 = 0x03;

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

impl PowerPointViewBarState {
    fn from_byte(value: u8) -> Result<Self> {
        Ok(match value {
            0x00 => Self::Minimized,
            0x01 => Self::Restored,
            0x02 => Self::Maximized,
            value => {
                return Err(corrupted(format!(
                    "NormalViewSetBarStates value {value} is undefined"
                )));
            },
        })
    }

    const fn byte(self) -> u8 {
        match self {
            Self::Minimized => 0x00,
            Self::Restored => 0x01,
            Self::Maximized => 0x02,
        }
    }
}

fn strict_bool(value: u8, name: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        value => Err(corrupted(format!("{name} must be 0 or 1, got {value}"))),
    }
}

fn check_unit_portion(ratio: PowerPointRatio, name: &str) -> Result<()> {
    if ratio.denominator() <= 0 || ratio.numerator() < 0 || ratio.numerator() > ratio.denominator()
    {
        return Err(corrupted(format!(
            "NormalViewSetInfo9Atom {name} must be between 0 and 1"
        )));
    }
    Ok(())
}

impl PowerPointNormalViewSetInfo {
    pub(super) fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != ATOM_LEN {
            return Err(corrupted(format!(
                "NormalViewSetInfo9Atom payload length must be {ATOM_LEN}, got {}",
                data.len()
            )));
        }
        let left_portion = PowerPointRatio::new(read_i32(data, 0), read_i32(data, 4))?;
        let top_portion = PowerPointRatio::new(read_i32(data, 8), read_i32(data, 12))?;
        check_unit_portion(left_portion, "leftPortion")?;
        check_unit_portion(top_portion, "topPortion")?;
        let flags = data[19];
        if flags & !KNOWN_FLAGS != 0 {
            return Err(corrupted(
                "NormalViewSetInfo9Atom reserved flag bits must be zero",
            ));
        }
        Ok(Self {
            left_portion,
            top_portion,
            vert_bar_state: PowerPointViewBarState::from_byte(data[16])?,
            horiz_bar_state: PowerPointViewBarState::from_byte(data[17])?,
            prefer_single_set: strict_bool(data[18], "NormalViewSetInfo9Atom.fPreferSingleSet")?,
            hide_thumbnails: flags & 0x01 != 0,
            bar_snapped: flags & 0x02 != 0,
        })
    }

    fn to_bytes(self) -> Result<Vec<u8>> {
        check_unit_portion(self.left_portion, "leftPortion")?;
        check_unit_portion(self.top_portion, "topPortion")?;
        let mut data = Vec::with_capacity(ATOM_LEN);
        for ratio in [self.left_portion, self.top_portion] {
            data.extend_from_slice(&ratio.numerator().to_le_bytes());
            data.extend_from_slice(&ratio.denominator().to_le_bytes());
        }
        data.push(self.vert_bar_state.byte());
        data.push(self.horiz_bar_state.byte());
        data.push(u8::from(self.prefer_single_set));
        data.push(u8::from(self.hide_thumbnails) | (u8::from(self.bar_snapped) << 1));
        Ok(data)
    }
}

impl PowerPointNormalViewSet {
    /// Parse a complete container record.
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        let declared = usize::try_from(record.data_length)
            .map_err(|_| corrupted("NormalViewSetInfo9 length does not fit memory"))?;
        if record.record_type != PptRecordType::NormalViewSetInfo9
            || record.record_type_raw != NORMAL_VIEW_SET_INFO_TYPE
            || record.version != 0xF
            || declared != record.data.len()
        {
            return Err(corrupted(
                "NormalViewSetInfo9 container has an invalid header or size",
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "NormalViewSetInfo9")?;
        let [atom] = children.as_slice() else {
            return Err(corrupted(
                "NormalViewSetInfo9 must contain exactly one NormalViewSetInfo9Atom",
            ));
        };
        if atom.record_type != PptRecordType::NormalViewSetInfo9Atom || atom.version != 0 {
            return Err(corrupted(
                "NormalViewSetInfo9 child is not a NormalViewSetInfo9Atom",
            ));
        }
        let payload = PowerPointNormalViewSetInfo::parse(&atom.data).map_or_else(
            |_| PowerPointNormalViewSetPayload::Other(atom.data.clone()),
            PowerPointNormalViewSetPayload::Layout,
        );
        Ok(Self { payload })
    }

    /// Serialize the complete container record, including its header.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let payload = match &self.payload {
            PowerPointNormalViewSetPayload::Layout(layout) => layout.to_bytes()?,
            PowerPointNormalViewSetPayload::Other(data) => data.clone(),
        };
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&NORMAL_VIEW_SET_INFO_ATOM_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(&payload);
        let mut record = Vec::with_capacity(8 + data.len());
        record.extend_from_slice(&(0x0010u16 | 0xF).to_le_bytes());
        record.extend_from_slice(&NORMAL_VIEW_SET_INFO_TYPE.to_le_bytes());
        record.extend_from_slice(&(data.len() as u32).to_le_bytes());
        record.extend_from_slice(&data);
        Ok(record)
    }
}

impl PowerPointNotesTextViewInfo {
    /// Parse a complete container record.
    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        let declared = usize::try_from(record.data_length)
            .map_err(|_| corrupted("NotesTextViewInfo9 length does not fit memory"))?;
        if record.record_type != PptRecordType::NotesTextViewInfo9
            || record.record_type_raw != NOTES_TEXT_VIEW_INFO_TYPE
            || record.version != 0xF
            || declared != record.data.len()
        {
            return Err(corrupted(
                "NotesTextViewInfo9 container has an invalid header or size",
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "NotesTextViewInfo9")?;
        let [atom] = children.as_slice() else {
            return Err(corrupted(
                "NotesTextViewInfo9 must contain exactly one zoom atom",
            ));
        };
        if atom.record_type != PptRecordType::ViewInfoAtom || atom.version != 0 {
            return Err(corrupted("NotesTextViewInfo9 child is not a ViewInfoAtom"));
        }
        Ok(Self {
            view_info: PowerPointNoZoomViewInfo::parse(&atom.data)?,
        })
    }
}

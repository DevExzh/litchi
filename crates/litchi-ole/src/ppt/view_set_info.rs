//! `NormalViewSetInfo9` and `NotesTextViewInfo9` view display preferences
//! (MS-PPT 2.4.21.2-2.4.21.4 and 2.4.21.7).
//!
//! `NormalViewSetInfo9` describes the pane splitter state of the normal
//! three-pane view; `NotesTextViewInfo9` the scaling of the notes-text view.
//! Both are inert display state and never drive any layout.

use super::non_zoom_view::PowerPointNoZoomViewInfo;
use super::records::record::PptRecord;
use super::view_info::PowerPointRatio;
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};

/// `RT_NormalViewSetInfo9` record type.
const NORMAL_VIEW_SET_INFO_TYPE: u16 = 0x0414;
/// `RT_NormalViewSetInfo9Atom` record type.
const NORMAL_VIEW_SET_INFO_ATOM_TYPE: u16 = 0x0415;
/// `RT_NotesTextViewInfo9` record type.
const NOTES_TEXT_VIEW_INFO_TYPE: u16 = 0x0413;
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

/// State of one pane splitter bar (`NormalViewSetBarStates`, MS-PPT 2.13.16).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PowerPointViewBarState {
    /// The region occupies a minimal area of the view.
    Minimized,
    /// The region has an intermediate size.
    Restored,
    /// The region occupies a maximal area of the view.
    Maximized,
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
    if ratio.denominator() <= 0 || ratio.numerator() < 0 || ratio.numerator() > ratio.denominator() {
        return Err(corrupted(format!(
            "NormalViewSetInfo9Atom {name} must be between 0 and 1"
        )));
    }
    Ok(())
}

/// Pane splitter state of the normal three-pane view (`NormalViewSetInfo9Atom`,
/// MS-PPT 2.4.21.3).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointNormalViewSetInfo {
    left_portion: PowerPointRatio,
    top_portion: PowerPointRatio,
    vert_bar_state: PowerPointViewBarState,
    horiz_bar_state: PowerPointViewBarState,
    prefer_single_set: bool,
    hide_thumbnails: bool,
    bar_snapped: bool,
}

impl PowerPointNormalViewSetInfo {
    /// Width of the side content pane as a fraction of the view width.
    pub const fn left_portion(&self) -> PowerPointRatio {
        self.left_portion
    }
    /// Height of the slide pane as a fraction of the view height.
    pub const fn top_portion(&self) -> PowerPointRatio {
        self.top_portion
    }
    /// State of the vertical splitter bar.
    pub const fn vert_bar_state(&self) -> PowerPointViewBarState {
        self.vert_bar_state
    }
    /// State of the horizontal splitter bar.
    pub const fn horiz_bar_state(&self) -> PowerPointViewBarState {
        self.horiz_bar_state
    }
    /// Whether the view consists of only the slide pane.
    pub const fn prefer_single_set(&self) -> bool {
        self.prefer_single_set
    }
    /// Whether the side content pane shows comments instead of thumbnails.
    pub const fn hide_thumbnails(&self) -> bool {
        self.hide_thumbnails
    }
    /// Whether the vertical bar snaps to specific positions when resized.
    pub const fn bar_snapped(&self) -> bool {
        self.bar_snapped
    }

    fn parse(data: &[u8]) -> Result<Self> {
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
            return Err(corrupted("NormalViewSetInfo9Atom reserved flag bits must be zero"));
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

/// The payload of a `NormalViewSetInfo9Atom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointNormalViewSetPayload {
    /// The specification pane layout (MS-PPT 2.4.21.3).
    Layout(PowerPointNormalViewSetInfo),
    /// An opaque payload preserved verbatim. POI's undocumented
    /// `SheetPropertiesAtom` (document timestamps) occupies the same record
    /// type in many files and falls into this variant.
    Other(Vec<u8>),
}

/// A `NormalViewSetInfo9Container` (MS-PPT 2.4.21.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointNormalViewSet {
    payload: PowerPointNormalViewSetPayload,
}

impl PowerPointNormalViewSet {
    /// The pane-layout payload.
    pub const fn payload(&self) -> &PowerPointNormalViewSetPayload {
        &self.payload
    }

    /// The pane layout, when the atom carries the specification payload.
    pub const fn layout(&self) -> Option<&PowerPointNormalViewSetInfo> {
        match &self.payload {
            PowerPointNormalViewSetPayload::Layout(layout) => Some(layout),
            PowerPointNormalViewSetPayload::Other(_) => None,
        }
    }

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
        let payload = PowerPointNormalViewSetInfo::parse(&atom.data)
            .map_or_else(|_| PowerPointNormalViewSetPayload::Other(atom.data.clone()),
                PowerPointNormalViewSetPayload::Layout);
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

/// A `NotesTextViewInfo9Container` (MS-PPT 2.4.21.4): scaling of the
/// notes-text view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointNotesTextViewInfo {
    view_info: PowerPointNoZoomViewInfo,
}

impl PowerPointNotesTextViewInfo {
    /// The notes-text view scaling and origin.
    pub const fn view_info(&self) -> &PowerPointNoZoomViewInfo {
        &self.view_info
    }

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

#[cfg(test)]
mod tests {
    use super::*;

    fn atom_record(data: &[u8]) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::NormalViewSetInfo9Atom,
            record_type_raw: NORMAL_VIEW_SET_INFO_ATOM_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    fn container_record(atom_data: &[u8]) -> PptRecord {
        let atom = atom_record(atom_data);
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&NORMAL_VIEW_SET_INFO_ATOM_TYPE.to_le_bytes());
        data.extend_from_slice(&(atom.data.len() as u32).to_le_bytes());
        data.extend_from_slice(&atom.data);
        PptRecord {
            record_type: PptRecordType::NormalViewSetInfo9,
            record_type_raw: NORMAL_VIEW_SET_INFO_TYPE,
            version: 0xF,
            instance: 1,
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        }
    }

    fn pane_atom() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&1i32.to_le_bytes());
        data.extend_from_slice(&4i32.to_le_bytes());
        data.extend_from_slice(&2i32.to_le_bytes());
        data.extend_from_slice(&3i32.to_le_bytes());
        data.push(1); // BS_Restored
        data.push(0); // BS_Minimized
        data.push(1); // fPreferSingleSet
        data.push(0x03); // fHideThumbnails | fBarSnapped
        data
    }

    #[test]
    fn parses_pane_layout_and_round_trips() {
        let container = PowerPointNormalViewSet::parse_record(&container_record(&pane_atom())).unwrap();
        let layout = container.layout().unwrap();
        assert_eq!(layout.left_portion().numerator(), 1);
        assert_eq!(layout.left_portion().denominator(), 4);
        assert_eq!(layout.vert_bar_state(), PowerPointViewBarState::Restored);
        assert_eq!(layout.horiz_bar_state(), PowerPointViewBarState::Minimized);
        assert!(layout.prefer_single_set());
        assert!(layout.hide_thumbnails());
        assert!(layout.bar_snapped());
        assert_eq!(
            container.to_bytes().unwrap()[8..],
            container_record(&pane_atom()).data[..]
        );
    }

    #[test]
    fn preserves_opaque_poi_sheet_properties_payloads() {
        let mut payload = pane_atom();
        // Out-of-range ratios are not a spec pane layout; POI timestamps land here.
        payload[0..8].copy_from_slice(&0x3B9A_CA00_F6B0_93BAu64.to_le_bytes());
        let container = PowerPointNormalViewSet::parse_record(&container_record(&payload)).unwrap();
        assert!(container.layout().is_none());
        let PowerPointNormalViewSetPayload::Other(raw) = container.payload() else {
            panic!()
        };
        assert_eq!(raw, &payload);
        assert_eq!(
            container.to_bytes().unwrap()[8..],
            container_record(&payload).data[..]
        );
    }

    #[test]
    fn rejects_malformed_layouts() {
        // Truncated atom.
        assert!(PowerPointNormalViewSetInfo::parse(&pane_atom()[..12]).is_err());
        // Ratio above 1.
        let mut bad = pane_atom();
        bad[0..4].copy_from_slice(&5i32.to_le_bytes());
        assert!(PowerPointNormalViewSetInfo::parse(&bad).is_err());
        // Undefined bar state.
        let mut bad = pane_atom();
        bad[16] = 3;
        assert!(PowerPointNormalViewSetInfo::parse(&bad).is_err());
        // Reserved flag bits set.
        let mut bad = pane_atom();
        bad[19] = 0xFC;
        assert!(PowerPointNormalViewSetInfo::parse(&bad).is_err());
        // Two atoms in one container.
        let mut data = container_record(&pane_atom()).data;
        data.extend_from_slice(&container_record(&pane_atom()).data);
        let record = PptRecord {
            record_type: PptRecordType::NormalViewSetInfo9,
            record_type_raw: NORMAL_VIEW_SET_INFO_TYPE,
            version: 0xF,
            instance: 1,
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        };
        assert!(PowerPointNormalViewSet::parse_record(&record).is_err());
    }
}

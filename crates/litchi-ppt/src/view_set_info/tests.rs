use super::codec::{NORMAL_VIEW_SET_INFO_ATOM_TYPE, NORMAL_VIEW_SET_INFO_TYPE};
use super::*;
use crate::consts::RecordType;
use crate::records::record::Record;

fn atom_record(data: &[u8]) -> Record {
    Record {
        record_type: RecordType::NormalViewSetInfo9Atom,
        record_type_raw: NORMAL_VIEW_SET_INFO_ATOM_TYPE,
        version: 0,
        instance: 0,
        data_length: data.len() as u32,
        data: data.to_vec(),
        children: Vec::new(),
    }
}

fn container_record(atom_data: &[u8]) -> Record {
    let atom = atom_record(atom_data);
    let mut data = Vec::new();
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&NORMAL_VIEW_SET_INFO_ATOM_TYPE.to_le_bytes());
    data.extend_from_slice(&(atom.data.len() as u32).to_le_bytes());
    data.extend_from_slice(&atom.data);
    Record {
        record_type: RecordType::NormalViewSetInfo9,
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
    let container = NormalViewSet::parse_record(&container_record(&pane_atom())).unwrap();
    let layout = container.layout().unwrap();
    assert_eq!(layout.left_portion().numerator(), 1);
    assert_eq!(layout.left_portion().denominator(), 4);
    assert_eq!(layout.vert_bar_state(), ViewBarState::Restored);
    assert_eq!(layout.horiz_bar_state(), ViewBarState::Minimized);
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
    let container = NormalViewSet::parse_record(&container_record(&payload)).unwrap();
    assert!(container.layout().is_none());
    let NormalViewSetPayload::Other(raw) = container.payload() else {
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
    assert!(NormalViewSetInfo::parse(&pane_atom()[..12]).is_err());
    // Ratio above 1.
    let mut bad = pane_atom();
    bad[0..4].copy_from_slice(&5i32.to_le_bytes());
    assert!(NormalViewSetInfo::parse(&bad).is_err());
    // Undefined bar state.
    let mut bad = pane_atom();
    bad[16] = 3;
    assert!(NormalViewSetInfo::parse(&bad).is_err());
    // Reserved flag bits set.
    let mut bad = pane_atom();
    bad[19] = 0xFC;
    assert!(NormalViewSetInfo::parse(&bad).is_err());
    // Two atoms in one container.
    let mut data = container_record(&pane_atom()).data;
    data.extend_from_slice(&container_record(&pane_atom()).data);
    let record = Record {
        record_type: RecordType::NormalViewSetInfo9,
        record_type_raw: NORMAL_VIEW_SET_INFO_TYPE,
        version: 0xF,
        instance: 1,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    };
    assert!(NormalViewSet::parse_record(&record).is_err());
}

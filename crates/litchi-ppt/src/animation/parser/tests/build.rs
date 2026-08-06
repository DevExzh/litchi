//! PowerPoint 2002 build-list round trips and validation.
use super::super::*;
use super::support::*;

#[test]
fn round_trips_exact_powerpoint_2002_build_lists() {
    assert_eq!(RecordType::BuildList.as_u16(), 0x2B02);
    assert_eq!(RecordType::LevelInfoAtom.as_u16(), 0x2B0A);
    let bytes = write_build_list(&sample_build_list()).unwrap();
    assert_eq!(RecordType::TimeNode.as_u16(), 0xF127);
    assert_eq!(bytes.len(), 216);
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(record.record_type, RecordType::BuildList);
    assert_eq!(record.children.len(), 3);

    let parsed = parse_build_list(&record).unwrap();
    assert_eq!(parsed.builds.len(), 3);
    let BuildListEntry::Paragraph(paragraph) = &parsed.builds[0] else {
        panic!("expected paragraph build");
    };
    assert_eq!(paragraph.atom.build_id, 10);
    assert_eq!(paragraph.atom.shape_id_ref, 100);
    assert_eq!(paragraph.paragraph.build_type, ParagraphBuildType::AsAWhole);
    assert_eq!(paragraph.paragraph.delay_time_ms, 750);
    assert_eq!(paragraph.levels.len(), 1);
    assert_eq!(paragraph.levels[0].level, 0);
    assert_eq!(paragraph.levels[0].time_node.atom, TimeNodeAtom::default());

    let BuildListEntry::Chart(chart) = &parsed.builds[1] else {
        panic!("expected chart build");
    };
    assert_eq!(chart.chart.build_type, ChartBuildType::ByElementInCategory);
    assert!(chart.chart.animate_background);

    let BuildListEntry::Diagram(diagram) = &parsed.builds[2] else {
        panic!("expected diagram build");
    };
    assert_eq!(
        diagram.diagram.build_type,
        DiagramBuildType::CounterClockwiseOut
    );
}

#[test]
fn rejects_malformed_powerpoint_2002_build_lists() {
    let bytes = write_build_list(&sample_build_list()).unwrap();
    let (valid, _) = Record::parse(&bytes, 0).unwrap();

    let mut truncated = bytes.clone();
    let claimed_length = u32::from_le_bytes(truncated[4..8].try_into().unwrap()) + 1;
    truncated[4..8].copy_from_slice(&claimed_length.to_le_bytes());
    let (truncated, _) = Record::parse(&truncated, 0).unwrap();
    assert_eq!(truncated.data_length, claimed_length);
    assert!(parse_build_list(&truncated).is_err());

    let mut malformed = Vec::new();
    let mut wrong_header = valid.clone();
    wrong_header.version = 0;
    malformed.push(wrong_header);

    let mut wrong_bool = valid.clone();
    wrong_bool.children[1].children[1].data[4] = 2;
    malformed.push(wrong_bool);

    let mut wrong_kind = valid.clone();
    wrong_kind.children[2].children[0].data[0..4].copy_from_slice(&1u32.to_le_bytes());
    malformed.push(wrong_kind);

    let mut wrong_level = valid.clone();
    wrong_level.children[0].children[2].data[0..4].copy_from_slice(&10u32.to_le_bytes());
    malformed.push(wrong_level);

    let mut duplicate = valid.clone();
    duplicate.children[2].children[0].data[4..8].copy_from_slice(&11u32.to_le_bytes());
    duplicate.children[2].children[0].data[8..12].copy_from_slice(&101u32.to_le_bytes());
    malformed.push(duplicate);

    let mut wrong_order = valid.clone();
    wrong_order.children[1].children.swap(0, 1);
    malformed.push(wrong_order);

    for record in malformed {
        assert!(parse_build_list(&record).is_err());
    }
}
#[test]
fn test_build_info_default() {
    let build_info = BuildInfo::default();
    assert!(build_info.builds.is_empty());
}

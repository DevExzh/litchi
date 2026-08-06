use super::super::write;
use crate::chart::plot_area::{Lines, PieTypeGroup, TypeGroup};
use crate::chart::{Chart, DataLabels};

#[test]
fn chart_style_record_is_emitted_without_reordering_chart_space() {
    let mut chart = Chart::new();
    chart.style = Some(48);

    let mut xml = Vec::new();
    write(&mut xml, &chart).unwrap();
    let xml = std::str::from_utf8(&xml).unwrap();
    assert!(xml.contains(r#"<c:style val="48"/>"#));
    assert!(xml.find(r#"<c:style val="48"/>"#).unwrap() < xml.find("<c:chart>").unwrap());
}

#[test]
fn data_label_leader_line_records_are_written_together() {
    let mut chart = Chart::new();
    let mut group = PieTypeGroup::new();
    let mut labels = DataLabels::new();
    labels.show_leader_lines = true;
    labels.leader_lines = Some(Lines::new());
    group.common.data_labels = Some(labels);
    chart.plot_area.type_groups.push(TypeGroup::Pie(group));

    let mut xml = Vec::new();
    write(&mut xml, &chart).unwrap();
    let xml = std::str::from_utf8(&xml).unwrap();
    assert!(xml.contains(r#"<c:showLeaderLines val="1"/>"#));
    assert!(xml.contains("<c:leaderLines/>"));
}

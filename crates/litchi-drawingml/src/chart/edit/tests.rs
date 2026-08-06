//! Focused source-preserving chart transaction tests.

use super::{DataLabelFlag, Snapshot};
use crate::chart::data::TitleText;
use crate::chart::types::{AxisPosition, DisplayBlanks};

const CHART: &[u8] = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:example:future">
  <c:lang val="en-US"/><c:style val="2"/>
  <c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
    <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Old</c:v></c:tx>
      <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f></c:strRef></c:cat>
      <c:val><c:numRef><c:f>Sheet1!$B$1:$B$2</c:f></c:numRef></c:val>
      <c:dLbls><c:showVal val="0"/><c:extLst><c:ext uri="opaque"><x:future/></c:ext></c:extLst></c:dLbls>
    </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>
  </c:plotArea><c:dispBlanksAs val="gap"/></c:chart>
  <c:extLst><c:ext uri="future"><x:payload keep="1"/></c:ext></c:extLst>
</c:chartSpace>"#;

#[test]
fn no_op_replays_source_and_changed_edits_keep_opaque_xml() {
    let snapshot = Snapshot::from_xml(CHART).expect("chart fixture must parse");
    let unchanged = snapshot.edit().commit().expect("no-op must commit");
    assert_eq!(unchanged.snapshot().xml_bytes(), CHART);
    assert_eq!(unchanged.patch().before_xml(), CHART);
    assert_eq!(unchanged.patch().after_xml(), CHART);

    let mut edit = snapshot.edit();
    edit.set_language(Some("zh-Hant"))
        .unwrap()
        .set_style(Some(8))
        .unwrap()
        .set_display_blanks(DisplayBlanks::Span)
        .unwrap()
        .set_series_title(0, Some(TitleText::from_string("New")))
        .unwrap()
        .set_series_data_label_flag(0, DataLabelFlag::ShowValue, true)
        .unwrap()
        .set_series_data_label_separator(0, Some(" | ".into()))
        .unwrap()
        .set_axis_range(2, Some(0.0), Some(100.0))
        .unwrap()
        .set_axis_position(1, AxisPosition::Top)
        .unwrap();
    let commit = edit.commit().expect("typed chart edit must commit");
    let output = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(output.contains(r#"val="zh-Hant""#));
    assert!(output.contains(r#"val="8""#));
    assert!(output.contains("New"));
    assert!(output.contains(" | "));
    assert!(output.contains(r#"uri="opaque""#));
    assert!(output.contains("keep=\"1\""));
    assert_eq!(
        commit.snapshot().value().language.as_deref(),
        Some("zh-Hant")
    );
    assert_eq!(commit.snapshot().value().style, Some(8));
    assert_eq!(
        commit.snapshot().value().display_blanks_as,
        DisplayBlanks::Span
    );

    let restored = commit
        .patch()
        .clone()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse patch must restore source");
    assert_eq!(restored.xml_bytes(), CHART);
}

#[test]
fn stale_and_invalid_edits_are_rejected_atomically() {
    let snapshot = Snapshot::from_xml(CHART).expect("chart fixture must parse");
    let mut changed_edit = snapshot.edit();
    changed_edit.set_style(Some(9)).unwrap();
    let changed = changed_edit.commit().unwrap();
    assert!(changed.patch().apply(&snapshot).is_ok());

    let other = Snapshot::from_xml(
        std::str::from_utf8(CHART)
            .unwrap()
            .replace("en-US", "fr-FR")
            .into_bytes(),
    )
    .unwrap();
    assert!(changed.patch().apply(&other).is_err());

    let mut edit = snapshot.edit();
    assert!(edit.set_style(Some(49)).is_err());
    assert!(!edit.is_changed());
    assert_eq!(edit.value().language.as_deref(), Some("en-US"));
    assert!(edit.set_axis_range(2, Some(10.0), Some(1.0)).is_err());
    assert!(!edit.is_changed());
    assert!(
        edit.set_series_title(99, Some(TitleText::from_string("bad")))
            .is_err()
    );
    assert!(!edit.is_changed());
}

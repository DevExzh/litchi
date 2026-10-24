//! Focused source-preserving chart transaction tests.

use super::{DataLabelFlag, Package};
use crate::Error;
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
    let snapshot = Package::from_xml(CHART).expect("chart fixture must parse");
    let unchanged = snapshot.edit().commit().expect("no-op must commit");
    assert_eq!(unchanged.package().xml_bytes(), CHART);
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
    let output = std::str::from_utf8(commit.package().xml_bytes()).unwrap();
    assert!(output.contains(r#"val="zh-Hant""#));
    assert!(output.contains(r#"val="8""#));
    assert!(output.contains("New"));
    assert!(output.contains(" | "));
    assert!(output.contains(r#"uri="opaque""#));
    assert!(output.contains("keep=\"1\""));
    assert_eq!(
        commit.package().value().language.as_deref(),
        Some("zh-Hant")
    );
    assert_eq!(commit.package().value().style, Some(8));
    assert_eq!(
        commit.package().value().display_blanks_as,
        DisplayBlanks::Span
    );
    let replayed = snapshot
        .apply(commit.patch())
        .expect("package replay must publish a commit");
    assert!(replayed.diagnostics().changed());
    assert_eq!(replayed.package().xml_bytes(), commit.package().xml_bytes());

    let inverse = commit.patch().clone().inverse();
    let restored = commit
        .package()
        .apply(&inverse)
        .expect("inverse patch must restore source");
    assert_eq!(restored.package().xml_bytes(), CHART);
}

#[test]
fn stale_and_invalid_edits_are_rejected_atomically() {
    let snapshot = Package::from_xml(CHART).expect("chart fixture must parse");
    let mut changed_edit = snapshot.edit();
    changed_edit.set_style(Some(9)).unwrap();
    let changed = changed_edit.commit().unwrap();
    assert!(snapshot.apply(changed.patch()).is_ok());

    let other = Package::from_xml(
        std::str::from_utf8(CHART)
            .unwrap()
            .replace("en-US", "fr-FR")
            .into_bytes(),
    )
    .unwrap();
    assert!(matches!(
        other.apply(changed.patch()),
        Err(Error::PatchConflict)
    ));

    let mut edit = snapshot.edit();
    assert!(edit.set_style(Some(49)).is_err());
    assert_eq!(edit.value().language.as_deref(), Some("en-US"));
    assert!(edit.set_axis_range(2, Some(10.0), Some(1.0)).is_err());
    assert!(
        edit.set_series_title(99, Some(TitleText::from_string("bad")))
            .is_err()
    );
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn series_colors_replace_insert_and_clear_fills() {
    use crate::chart::authoring::{Kind, Series, Spec, Variant};
    let color = |xml: &[u8]| {
        crate::chart::authoring::read(xml, &crate::chart::Limits::DEFAULT)
            .unwrap()
            .spec
            .series()[0]
            .color
            .clone()
    };
    // Inserted into a series without shape properties, then replaced, then
    // cleared; the opaque extension of the fixture survives.
    let snapshot = Package::from_xml(CHART).unwrap();
    let mut edit = snapshot.edit();
    edit.set_series_color(0, Some("c00000")).unwrap();
    let inserted = edit.commit().unwrap().into_package();
    assert!(String::from_utf8_lossy(inserted.xml_bytes()).contains("<x:payload keep=\"1\"/>"));
    let mut edit = inserted.edit();
    edit.set_series_color(0, Some("8000FF00")).unwrap();
    let replaced = edit.commit().unwrap().into_package();
    let text = String::from_utf8_lossy(replaced.xml_bytes()).into_owned();
    assert!(text.contains(r#"<a:srgbClr val="00FF00"><a:alpha val="50196"/></a:srgbClr>"#));
    assert_eq!(text.matches("solidFill>").count(), 2, "{text}");
    let mut edit = replaced.edit();
    edit.set_series_color(0, None).unwrap();
    assert!(
        !String::from_utf8_lossy(edit.commit().unwrap().package().xml_bytes()).contains("srgbClr")
    );
    assert!(snapshot.edit().set_series_color(0, Some("red")).is_err());
    assert!(snapshot.edit().set_series_color(9, Some("FF0000")).is_err());

    // A line series takes the color on its line.
    for kind in [Kind::Line, Kind::Column] {
        let spec = Spec::new(Variant::new(kind))
            .with_series(Series::new("S", vec![1.0, 2.0]).with_color("112233"));
        let snapshot = Package::from_xml(spec.to_xml(None).unwrap()).unwrap();
        let mut edit = snapshot.edit();
        edit.set_series_color(0, Some("445566")).unwrap();
        let xml = edit.commit().unwrap().into_package().xml_bytes().to_vec();
        assert_eq!(color(&xml).as_deref(), Some("445566"), "{kind:?}");
        // The edited part reads as the authored chart of the new color.
        let parsed = crate::chart::authoring::read(&xml, &crate::chart::Limits::DEFAULT).unwrap();
        assert!(
            parsed.is_authored(&xml, &crate::chart::Limits::DEFAULT),
            "{kind:?}"
        );
    }
}

/// Record 0290 F2: a source-preserving edit finds an axis whose producer ID
/// is spelled as a negative signed integer.
#[test]
fn edits_find_axes_with_negative_producer_ids() {
    let source = std::str::from_utf8(CHART)
        .unwrap()
        .replace(r#"<c:axId val="1"/>"#, r#"<c:axId val="-1884094432"/>"#)
        .replace(
            r#"<c:crossAx val="1"/>"#,
            r#"<c:crossAx val="-1884094432"/>"#,
        );
    let snapshot = Package::from_xml(source.as_bytes()).expect("chart fixture must parse");
    let axis = (-1_884_094_432_i32).cast_unsigned();
    let mut edit = snapshot.edit();
    edit.set_axis_position(axis, AxisPosition::Top)
        .expect("the negative axis ID must resolve");
    let commit = edit.commit().expect("edit must commit");
    let after = std::str::from_utf8(commit.patch().after_xml()).unwrap();
    assert!(after.contains(r#"<c:axId val="-1884094432"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="t"/>"#), "{after}");
}

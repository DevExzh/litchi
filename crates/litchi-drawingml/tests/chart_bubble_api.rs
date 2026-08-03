use litchi_drawingml::chart::bubble::{Scale, Size};
use litchi_drawingml::chart::plot_area::{BubbleTypeGroup, TypeGroup};
use litchi_drawingml::chart::series::Series;
use litchi_drawingml::chart::{Chart, PlotArea};

#[test]
fn typed_bubble_settings_round_trip_through_the_public_api() {
    let scale = Scale::new(125).expect("125 is a valid bubble scale");
    let group = BubbleTypeGroup::new()
        .with_scale(scale)
        .with_size(Size::Width);

    assert_eq!(group.scale(), scale);
    assert_eq!(group.size(), Size::Width);

    let mut chart = Chart::new();
    chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Bubble(group));

    let mut xml = Vec::new();
    litchi_drawingml::chart::writer::write(&mut xml, &chart)
        .expect("typed bubble settings must serialize");
    let xml_text = std::str::from_utf8(&xml).expect("chart XML must be UTF-8");
    assert!(xml_text.contains(r#"<c:bubbleScale val="125"/>"#));
    assert!(xml_text.contains(r#"<c:sizeRepresents val="w"/>"#));

    let parsed =
        litchi_drawingml::chart::reader::read(xml.as_slice()).expect("serialized chart must parse");
    let [TypeGroup::Bubble(parsed)] = parsed.plot_area.type_groups.as_slice() else {
        panic!("expected one bubble chart group");
    };
    assert_eq!(parsed.scale(), scale);
    assert_eq!(parsed.size(), Size::Width);
}

#[test]
fn parser_rejects_invalid_bubble_settings() {
    for setting in [
        r#"<c:bubbleScale val="301"/>"#,
        r#"<c:sizeRepresents val="diameter"/>"#,
    ] {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bubbleChart>{setting}</c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#
        );
        assert!(litchi_drawingml::chart::reader::read(xml.as_bytes()).is_err());
    }
}

#[test]
fn office_compatible_output_keeps_bubble_3d_at_series_level() {
    let mut series = Series::new(0);
    series.bubble_3d = true;
    let mut group = BubbleTypeGroup::new();
    group.common.series.push(series);

    let mut chart = Chart::new();
    chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Bubble(group));
    let mut xml = Vec::new();
    litchi_drawingml::chart::writer::write(&mut xml, &chart)
        .expect("typed bubble chart must serialize");
    let xml = String::from_utf8(xml).expect("chart XML must be UTF-8");

    assert_eq!(xml.matches("<c:bubble3D").count(), 1);
    assert!(
        xml.find("<c:bubble3D").expect("series bubble3D")
            < xml.find("</c:ser>").expect("series end")
    );

    // ECMA-376 permits a group-level value that desktop Excel rejects. The
    // reader accepts it and projects its semantic effect onto each series.
    let legacy = xml.replacen(r#"<c:bubble3D val="1"/>"#, "", 1).replacen(
        "<c:bubbleScale",
        r#"<c:bubble3D val="1"/><c:bubbleScale"#,
        1,
    );
    let parsed = litchi_drawingml::chart::reader::read(legacy.as_bytes())
        .expect("standard group-level bubble3D must remain readable");
    let [TypeGroup::Bubble(parsed)] = parsed.plot_area.type_groups.as_slice() else {
        panic!("expected one bubble chart group");
    };
    assert!(parsed.common.series[0].bubble_3d);
}

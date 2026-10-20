#![allow(
    clippy::unwrap_used,
    reason = "tests panic on unexpected errors for concise assertions"
)]

use litchi_odc::{
    AxisSpec, Builder, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, DataPointSpec,
    Definition, History, OpenOptions, Package, SemanticLimits, SeriesSpec, StyleTarget,
    chart::Dimension,
};

fn definition() -> Definition {
    let mut value = Definition::new(ChartClass::bar());
    let mut axis = AxisSpec::new(Dimension::X);
    axis.name = Some("categories".into());
    value.plot_area.axes.push(axis);
    value.plot_area.series.push(SeriesSpec {
        attached_axis: Some("categories".into()),
        values_cell_range_address: Some("Data.$B$2:.$B$4".into()),
        data_points: vec![DataPointSpec::default()],
        ..SeriesSpec::default()
    });
    let mut table = CachedTable::new("Data", 2);
    table
        .rows
        .push(CachedRow::new(vec![CachedCell::new(CachedValue::Float(
            1.0,
        ))]));
    value.cached_table = Some(table);
    value
}

#[test]
fn granular_definition_crud_composes_and_survives_history() {
    let source = Builder::new()
        .with_definition(definition())
        .build()
        .unwrap();
    let mut first_edit = source.edit();
    let definition_edit = first_edit.definition_edit().unwrap();
    definition_edit
        .insert_axis(1, AxisSpec::new(Dimension::Y))
        .unwrap();
    definition_edit
        .insert_series(1, SeriesSpec::default())
        .unwrap();
    definition_edit
        .insert_data_point(1, 0, DataPointSpec::default())
        .unwrap();
    definition_edit
        .set_style(StyleTarget::Series(1), Some("second-series".into()))
        .unwrap();
    definition_edit
        .insert_cached_cell(
            false,
            0,
            1,
            CachedCell::new(CachedValue::String("one".into())),
        )
        .unwrap();
    let first = first_edit.commit().unwrap();
    assert!(first.diagnostics().changed());
    assert!(!first.patch().definition_changes().is_empty());

    let mut second_edit = first.package().edit();
    let definition_edit = second_edit.definition_edit().unwrap();
    definition_edit.remove_data_point(0, 0).unwrap();
    definition_edit.remove_series(0).unwrap();
    definition_edit.remove_axis(0).unwrap();
    definition_edit
        .update_cached_row(
            false,
            0,
            CachedRow::new(vec![CachedCell::new(CachedValue::Float(2.0))]),
        )
        .unwrap();
    let second = second_edit.commit().unwrap();

    let composed = first.patch().compose(second.patch()).unwrap();
    let target = source.apply(&composed).unwrap().into_package();
    assert_eq!(
        target.chart().definition().unwrap(),
        second.package().chart().definition().unwrap()
    );
    assert_eq!(
        target
            .apply(&composed.inverse())
            .unwrap()
            .package()
            .chart()
            .definition()
            .unwrap(),
        source.chart().definition().unwrap()
    );
    assert!(second.package().apply(&composed).is_err());

    let mut history = History::new(source.clone());
    history.record(&first).unwrap();
    history.record(&second).unwrap();
    assert!(history.undo().unwrap());
    assert_eq!(history.current().as_bytes(), first.package().as_bytes());
    assert!(history.redo().unwrap());
    assert_eq!(history.current().as_bytes(), second.package().as_bytes());
}

#[test]
fn formulas_ranges_and_caller_limits_refuse_invalid_publication() {
    let mut value = definition();
    value.cached_table.as_mut().unwrap().rows[0].cells[0].formula =
        Some("of:=SUM([.B2:.B4])".into());
    assert!(
        Builder::new()
            .with_definition(value.clone())
            .build()
            .is_ok()
    );

    value.cached_table.as_mut().unwrap().rows[0].cells[0].formula =
        Some("of:=SUM([.b2:.B4])".into());
    assert!(Builder::new().with_definition(value).build().is_err());

    let tight_axes = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_axes(1)
            .unwrap()
            .build()
            .unwrap(),
    );
    let mut too_many = definition();
    too_many.plot_area.axes.push(AxisSpec::new(Dimension::Y));
    assert!(
        Builder::new()
            .with_options(tight_axes)
            .with_definition(too_many)
            .build()
            .is_err()
    );

    let bytes = Builder::new()
        .with_definition(definition())
        .build()
        .unwrap()
        .to_bytes()
        .unwrap();
    let tight_package = litchi_odc::SourceLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).unwrap())
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(tight_package)),
        Err(litchi_odc::Error::Limit { .. })
    ));
}

#[test]
fn package_styles_and_resource_crud_reopens_composes_and_reverses() {
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.4"><office:styles/></office:document-styles>"#;
    let source = Builder::new()
        .with_definition(definition())
        .build()
        .unwrap();
    let mut first_edit = source.edit();
    first_edit.set_styles_xml(styles);
    first_edit
        .add_resource("Pictures/chart.bin", "application/octet-stream", b"one")
        .unwrap();
    let first = first_edit.commit().unwrap();
    assert_eq!(first.package().chart().styles_xml(), Some(styles));
    assert_eq!(first.package().chart().resources().len(), 1);
    assert_eq!(first.package().chart().resource_bytes(0).unwrap(), b"one");
    let reopened = Package::from_bytes(first.package().as_bytes()).unwrap();
    assert_eq!(reopened.chart().resource_bytes(0).unwrap(), b"one");

    let mut second_edit = first.package().edit();
    second_edit
        .update_resource(0, "application/octet-stream", b"two")
        .unwrap();
    second_edit.remove_styles_xml();
    let second = second_edit.commit().unwrap();
    let composed = first.patch().compose(second.patch()).unwrap();
    assert_eq!(
        source.apply(&composed).unwrap().into_package().as_bytes(),
        second.package().as_bytes()
    );
    assert_eq!(
        second
            .package()
            .apply(&composed.inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        source.as_bytes()
    );
    assert!(first.package().apply(&composed).is_err());

    let mut remove = second.package().edit();
    remove.remove_resource(0).unwrap();
    let removed = remove.commit().unwrap();
    assert!(removed.package().chart().resources().is_empty());
    assert!(removed.package().apply(&removed.patch().inverse()).is_ok());
}

#[test]
fn malformed_style_resource_and_bounded_history_are_refused() {
    let source = Builder::new()
        .with_definition(definition())
        .build()
        .unwrap();
    let mut bad_style = source.edit();
    bad_style.set_styles_xml("<office:document-styles>");
    assert!(bad_style.commit().is_err());

    let mut bad_media = source.edit();
    assert!(
        bad_media
            .add_resource("Pictures/a.bin", "not a media type", b"x")
            .is_err()
    );
    assert!(bad_media.remove_resource(0).is_err());

    let one_step = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_history(1)
            .unwrap()
            .build()
            .unwrap(),
    );
    let initial = Builder::new()
        .with_definition(definition())
        .with_options(one_step)
        .build()
        .unwrap();
    assert_eq!(initial.options(), one_step);
    let mut first_edit = initial.edit();
    first_edit
        .definition_edit()
        .unwrap()
        .set_style(StyleTarget::Chart, Some("one".into()))
        .unwrap();
    let first = first_edit.commit().unwrap();
    let mut second_edit = first.package().edit();
    second_edit
        .definition_edit()
        .unwrap()
        .set_style(StyleTarget::Chart, Some("two".into()))
        .unwrap();
    let second = second_edit.commit().unwrap();
    let mut history = History::new(initial);
    history.record(&first).unwrap();
    assert!(history.record(&second).is_err());
}

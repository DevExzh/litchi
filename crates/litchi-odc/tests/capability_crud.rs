#![allow(
    clippy::unwrap_used,
    reason = "tests panic on unexpected errors for concise assertions"
)]

use litchi_odc::{
    AxisSpec, Builder, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, DataPointSpec,
    Definition, DefinitionHistory, DefinitionSnapshot, Limits, Package, SeriesSpec, StyleTarget,
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
    let source = DefinitionSnapshot::with_default_limits(definition()).unwrap();
    let mut first_edit = source.edit();
    first_edit
        .insert_axis(1, AxisSpec::new(Dimension::Y))
        .unwrap();
    first_edit.insert_series(1, SeriesSpec::default()).unwrap();
    first_edit
        .insert_data_point(1, 0, DataPointSpec::default())
        .unwrap();
    first_edit
        .set_style(StyleTarget::Series(1), Some("second-series".into()))
        .unwrap();
    first_edit
        .insert_cached_cell(
            false,
            0,
            1,
            CachedCell::new(CachedValue::String("one".into())),
        )
        .unwrap();
    let first = first_edit.commit().unwrap();

    let mut second_edit = first.snapshot().edit();
    second_edit.remove_data_point(0, 0).unwrap();
    second_edit.remove_series(0).unwrap();
    second_edit.remove_axis(0).unwrap();
    second_edit
        .update_cached_row(
            false,
            0,
            CachedRow::new(vec![CachedCell::new(CachedValue::Float(2.0))]),
        )
        .unwrap();
    let second = second_edit.commit().unwrap();

    let composed = first.patch().compose(second.patch()).unwrap();
    let target = composed.apply(&source).unwrap();
    assert_eq!(target.definition(), second.snapshot().definition());
    assert_eq!(
        composed.inverse().apply(&target).unwrap().definition(),
        source.definition()
    );
    assert!(composed.apply(second.snapshot()).is_err());

    let mut history = DefinitionHistory::new(source.clone());
    history.record(first.patch()).unwrap();
    history.record(second.patch()).unwrap();
    assert!(history.undo().unwrap());
    assert_eq!(
        history.current().definition(),
        first.snapshot().definition()
    );
    assert!(history.redo().unwrap());
    assert_eq!(
        history.current().definition(),
        second.snapshot().definition()
    );
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

    let tight_axes = Limits::new().with_axes(1).unwrap();
    let mut too_many = definition();
    too_many.plot_area.axes.push(AxisSpec::new(Dimension::Y));
    assert!(
        Builder::new()
            .with_limits(tight_axes)
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
        Package::from_bytes_with_options(
            &bytes,
            litchi_odc::OpenOptions::new().source_limits(tight_package)
        ),
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

    let limits = Limits::new().with_history(1).unwrap();
    let initial = DefinitionSnapshot::new(definition(), limits).unwrap();
    let mut first_edit = initial.edit();
    first_edit
        .set_style(StyleTarget::Chart, Some("one".into()))
        .unwrap();
    let first = first_edit.commit().unwrap();
    let mut second_edit = first.snapshot().edit();
    second_edit
        .set_style(StyleTarget::Chart, Some("two".into()))
        .unwrap();
    let second = second_edit.commit().unwrap();
    let mut history = DefinitionHistory::new(initial);
    history.record(first.patch()).unwrap();
    assert!(history.record(second.patch()).is_err());
}

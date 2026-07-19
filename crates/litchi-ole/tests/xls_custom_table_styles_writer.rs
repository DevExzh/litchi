use std::io::Cursor;

use litchi_ole::xls::writer::{XlsCustomTableStyles, XlsWriter};
use litchi_ole::xls::{
    XlsDifferentialFormat, XlsDifferentialFormatId, XlsTableStyle, XlsTableStyleElement,
    XlsTableStyleRegion, XlsTableStyles, XlsWorkbook, XlsXfProperty,
};

fn configured_styles(number_format_id: u16) -> XlsCustomTableStyles {
    let differential_formats = vec![
        XlsDifferentialFormat::try_new(
            false,
            vec![XlsXfProperty::NumberFormatId(number_format_id)],
        )
        .unwrap(),
        XlsDifferentialFormat::try_new(false, vec![XlsXfProperty::Locked(true)]).unwrap(),
    ];
    let style = XlsTableStyle::try_new(
        "WriterCustom",
        true,
        false,
        vec![
            XlsTableStyleElement::try_new(
                XlsTableStyleRegion::WholeTable,
                XlsDifferentialFormatId::new(0),
            )
            .unwrap(),
            XlsTableStyleElement::try_with_stripe_size(
                XlsTableStyleRegion::FirstRowStripe,
                2,
                XlsDifferentialFormatId::new(1),
            )
            .unwrap(),
        ],
    )
    .unwrap();
    XlsCustomTableStyles::try_from_styles(
        differential_formats,
        "WriterCustom",
        "PivotStyleLight16",
        vec![style],
    )
    .unwrap()
}

#[test]
fn custom_table_styles_round_trip_from_writer_to_reader() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Data").unwrap();
    let number_format_id = writer.register_number_format("0.0000\"kg\"");
    writer
        .set_custom_table_styles(configured_styles(number_format_id))
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();

    assert_eq!(workbook.differential_formats().len(), 2);
    assert_eq!(
        workbook.differential_formats()[0].properties().properties(),
        &[XlsXfProperty::NumberFormatId(number_format_id)]
    );
    let catalog = workbook.table_styles().unwrap();
    assert_eq!(catalog.default_table_style(), "WriterCustom");
    assert_eq!(catalog.custom_style_count(), 1);
    let style = &catalog.custom_styles()[0];
    assert_eq!(style.name(), "WriterCustom");
    assert_eq!(style.elements().len(), 2);
    assert_eq!(style.elements()[1].stripe_size(), Some(2));
}

#[test]
fn custom_table_style_configuration_rejects_hostile_relationships() {
    let dxf = XlsDifferentialFormat::try_new(false, vec![XlsXfProperty::Locked(true)]).unwrap();

    let incomplete_catalog =
        XlsTableStyles::try_new(145, "TableStyleMedium2", "PivotStyleLight16").unwrap();
    assert!(XlsCustomTableStyles::try_new(vec![dxf.clone()], incomplete_catalog).is_err());

    let dangling_style = XlsTableStyle::try_new(
        "Dangling",
        true,
        false,
        vec![
            XlsTableStyleElement::try_new(
                XlsTableStyleRegion::WholeTable,
                XlsDifferentialFormatId::new(1),
            )
            .unwrap(),
        ],
    )
    .unwrap();
    assert!(
        XlsCustomTableStyles::try_from_styles(
            vec![dxf],
            "Dangling",
            "PivotStyleLight16",
            vec![dangling_style],
        )
        .is_err()
    );

    let duplicate_region = XlsTableStyle::try_new(
        "Duplicate",
        true,
        false,
        vec![
            XlsTableStyleElement::try_new(
                XlsTableStyleRegion::WholeTable,
                XlsDifferentialFormatId::new(0),
            )
            .unwrap(),
            XlsTableStyleElement::try_new(
                XlsTableStyleRegion::WholeTable,
                XlsDifferentialFormatId::new(0),
            )
            .unwrap(),
        ],
    );
    assert!(duplicate_region.is_err());
    assert!(XlsTableStyle::try_new("", true, false, Vec::new()).is_err());
}

#[test]
fn replacement_is_atomic_and_clear_restores_unset_bytes() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Data").unwrap();
    let valid = configured_styles(0);
    writer.set_custom_table_styles(valid.clone()).unwrap();

    let undefined_number_format = configured_styles(164);
    assert!(
        writer
            .set_custom_table_styles(undefined_number_format)
            .is_err()
    );

    let mut retained = Cursor::new(Vec::new());
    writer.write_to(&mut retained).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(retained.into_inner())).unwrap();
    assert_eq!(
        workbook.differential_formats()[0].properties().properties(),
        &[XlsXfProperty::NumberFormatId(0)]
    );

    writer.clear_custom_table_styles();
    let mut cleared = Cursor::new(Vec::new());
    writer.write_to(&mut cleared).unwrap();

    let mut fresh = XlsWriter::new();
    fresh.add_worksheet("Data").unwrap();
    let mut baseline = Cursor::new(Vec::new());
    fresh.write_to(&mut baseline).unwrap();
    assert_eq!(cleared.into_inner(), baseline.into_inner());
}

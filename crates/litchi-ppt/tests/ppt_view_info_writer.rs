use std::io::Cursor;

use litchi_ppt::writer::{
    PowerPointGuide, PowerPointGuideOrientation, PowerPointRatio, PowerPointSlideViewInfo,
    PowerPointSlideViewPreferences, PowerPointViewKind, PowerPointViewOrigin,
    PowerPointZoomViewInfo, PptWriter,
};
use litchi_ppt::{Package, PptRecord};

fn slide_view() -> PowerPointSlideViewInfo {
    let scale = PowerPointRatio::new(3, 4).unwrap();
    PowerPointSlideViewInfo::new(
        PowerPointViewKind::Slide,
        PowerPointSlideViewPreferences::new(true, false),
        Some(
            PowerPointZoomViewInfo::new(
                scale,
                scale,
                PowerPointViewOrigin::new(-100, 200),
                true,
                false,
            )
            .unwrap(),
        ),
        vec![
            PowerPointGuide::new(PowerPointGuideOrientation::Horizontal, 2160).unwrap(),
            PowerPointGuide::new(PowerPointGuideOrientation::Vertical, 2880).unwrap(),
        ],
    )
    .unwrap()
}

fn notes_view() -> PowerPointSlideViewInfo {
    PowerPointSlideViewInfo::new(
        PowerPointViewKind::Notes,
        PowerPointSlideViewPreferences::new(false, true),
        None,
        vec![PowerPointGuide::new(PowerPointGuideOrientation::Horizontal, 720).unwrap()],
    )
    .unwrap()
}

fn write(writer: &mut PptWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn view_records(bytes: &[u8]) -> Vec<PptRecord> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let (document, _) = PptRecord::parse(&stream, 0).unwrap();
    let doc_info = document
        .children
        .iter()
        .find(|record| record.record_type_raw == 2000)
        .unwrap();
    doc_info
        .children
        .iter()
        .filter(|record| record.record_type_raw == 1018)
        .cloned()
        .collect()
}

#[test]
fn emits_exact_slide_and_notes_view_record_structure() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer.set_slide_view_info(slide_view()).unwrap();
    writer.set_notes_view_info(notes_view()).unwrap();

    let bytes = write(&mut writer);
    let views = view_records(&bytes);
    assert_eq!(views.len(), 2);
    assert_eq!((views[0].version, views[0].instance), (0xF, 0));
    assert_eq!((views[1].version, views[1].instance), (0xF, 1));
    assert_eq!(
        views[0]
            .children
            .iter()
            .map(|record| (record.record_type_raw, record.instance, record.data.len()))
            .collect::<Vec<_>>(),
        vec![(1022, 0, 3), (1021, 0, 52), (1019, 7, 8), (1019, 7, 8)]
    );
    assert_eq!(
        views[1]
            .children
            .iter()
            .map(|record| (record.record_type_raw, record.instance, record.data.len()))
            .collect::<Vec<_>>(),
        vec![(1022, 0, 3), (1019, 7, 8)]
    );
}

#[test]
fn simultaneous_slide_and_notes_views_round_trip_through_writer_and_reader() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer.set_slide_view_info(slide_view()).unwrap();
    writer.set_notes_view_info(notes_view()).unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let presentation = package.presentation().unwrap();
    let information = presentation.slide_view_information().unwrap();
    let slide = information.slide().unwrap();
    let notes = information.notes().unwrap();
    assert!(slide.preferences().snap_to_grid());
    assert!(!slide.preferences().snap_to_shape());
    assert_eq!(slide.zoom().unwrap().x_scale().numerator(), 3);
    assert_eq!(slide.guides().len(), 2);
    assert!(!notes.preferences().snap_to_grid());
    assert!(notes.preferences().snap_to_shape());
    assert!(notes.zoom().is_none());
    assert_eq!(notes.guides()[0].position(), 720);
}

#[test]
fn setters_reject_wrong_kinds_atomically_and_models_enforce_limits() {
    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer.set_slide_view_info(slide_view()).unwrap();
    assert!(writer.set_slide_view_info(notes_view()).is_err());
    assert_eq!(
        writer.slide_view_info().unwrap().kind(),
        PowerPointViewKind::Slide
    );
    assert!(writer.set_notes_view_info(slide_view()).is_err());
    assert!(writer.notes_view_info().is_none());

    assert!(PowerPointGuide::new(PowerPointGuideOrientation::Vertical, 32_256).is_err());
    let too_many =
        vec![PowerPointGuide::new(PowerPointGuideOrientation::Horizontal, 0).unwrap(); 9];
    assert!(
        PowerPointSlideViewInfo::new(
            PowerPointViewKind::Slide,
            PowerPointSlideViewPreferences::new(false, false),
            None,
            too_many,
        )
        .is_err()
    );
    assert!(
        PowerPointZoomViewInfo::new(
            PowerPointRatio::new(1, 2).unwrap(),
            PowerPointRatio::new(3, 4).unwrap(),
            PowerPointViewOrigin::new(0, 0),
            false,
            false,
        )
        .is_err()
    );

    writer.clear_slide_view_info();
    writer.clear_notes_view_info();
    assert!(writer.slide_view_info().is_none());
    assert!(writer.notes_view_info().is_none());
    let views = view_records(&write(&mut writer));
    assert_eq!(views.len(), 1);
    assert_eq!(views[0].instance, 0);
    assert_eq!(views[0].children.len(), 1);
}

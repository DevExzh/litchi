//! Native BIFF8 worksheet metadata and opaque image records; no rendering.
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::{
    BackgroundImage, BackgroundImageFormat, CellWatch, ConsolidationFunction, ConsolidationSource,
    FormulaErrorChecks, FormulaErrorFeature, HeaderFooterPicture, Package, raw::Reader,
};
use std::{collections::BTreeMap, io::Cursor};
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-metadata.xls"),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let n = p.iter().map(String::as_str).collect::<Vec<_>>();
            let b = ole.open_stream(&n).unwrap();
            (p, b)
        })
        .collect()
}
fn records(bytes: &[u8]) -> Vec<(u16, &[u8])> {
    let mut out = Vec::new();
    let mut at = 0;
    while at < bytes.len() {
        let k = u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            bytes[at + 2..at + 4].try_into().unwrap(),
        ));
        out.push((k, &bytes[at + 4..at + 4 + len]));
        at += len + 4;
    }
    assert_eq!(at, bytes.len());
    out
}
fn caches(r: &Reader<Cursor<&Vec<u8>>>) {
    let s = r.xls_worksheet(0).unwrap();
    for (row, text, value) in [
        (3, "=(A3+1)", CellValue::Error("#VALUE!".into())),
        (4, "=(1/0)", CellValue::Error("#DIV/0!".into())),
        (6, "=SUM(10,20)", CellValue::Float(30.)),
    ] {
        let c = s.get_cell(row, 0).unwrap();
        assert_eq!(c.formula(), Some(text));
        assert_eq!(c.value(), &value);
    }
}

#[test]
fn native_cell_watch_and_error_check_records_round_trip_with_saved_errors() {
    let bytes = native();
    let r = Reader::new(Cursor::new(&bytes)).unwrap();
    assert_eq!(r.sheets().len(), 5);
    caches(&r);
    let sheet = r.xls_worksheet(0).unwrap();
    let features = sheet.formula_error_features();
    assert_eq!(features.len(), 1);
    let feature = &features[0];
    assert_eq!(feature.checks().bits(), 1);
    assert!(feature.checks().calculation_errors());
    assert_eq!(feature.ranges().len(), 1);
    let range = feature.ranges()[0];
    assert_eq!(
        (
            range.first_row(),
            range.last_row(),
            range.first_column(),
            range.last_column()
        ),
        (3, 4, 0, 0)
    );
    let parts = streams(&bytes);
    let recs = records(&parts[&vec!["Workbook".into()]]);
    let watches = recs
        .iter()
        .filter(|(k, _)| *k == 0x086c)
        .collect::<Vec<_>>();
    assert_eq!(watches.len(), 1);
    let watch = CellWatch::parse(watches[0].1).unwrap();
    assert_eq!(watch.flags(), 1);
    assert_eq!(
        (
            watch.row_first(),
            watch.row_last(),
            watch.column_first(),
            watch.column_last()
        ),
        (6, 6, 0, 0)
    );
    assert_eq!(watch.to_payload(), watches[0].1);
    assert!(
        recs.iter()
            .any(|(k, p)| *k == 0x0868 && *p == feature.to_payload().unwrap())
    );
    assert_eq!(
        FormulaErrorFeature::parse_payload(&feature.to_payload().unwrap()).unwrap(),
        *feature
    );
    let detached =
        FormulaErrorFeature::try_new(feature.ranges().to_vec(), FormulaErrorChecks::from_bits(5))
            .unwrap();
    assert!(detached.checks().calculation_errors() && detached.checks().numbers_stored_as_text());
    assert_eq!(
        FormulaErrorFeature::parse_payload(&detached.to_payload().unwrap()).unwrap(),
        detached
    );
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().to_bytes().unwrap(),
        bytes
    );
}

#[test]
fn native_background_and_header_footer_images_keep_continuations_and_opaque_payloads() {
    let bytes = native();
    let parts = streams(&bytes);
    let recs = records(&parts[&vec!["Workbook".into()]]);
    let index = recs.iter().position(|(k, _)| *k == 0x00e9).unwrap();
    let mut combined = recs[index].1.to_vec();
    let mut continued = 0;
    for (k, p) in &recs[index + 1..] {
        if *k != 0x003c {
            break;
        }
        continued += 1;
        combined.extend_from_slice(p);
    }
    assert_eq!(continued, 21);
    let image = BackgroundImage::parse(&combined).unwrap();
    assert_eq!(image.format(), BackgroundImageFormat::Bitmap);
    assert_eq!(image.reserved(), 1);
    assert_eq!(image.image().len(), 172812);
    assert_eq!(
        &image.image()[..12],
        &[12, 0, 0, 0, 64, 1, 180, 0, 1, 0, 24, 0]
    );
    assert_eq!(image.to_payload(), combined);
    let pictures = recs
        .iter()
        .filter(|(k, _)| *k == 0x0866)
        .map(|(_, p)| (*p, HeaderFooterPicture::parse(p).unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(pictures.len(), 2);
    for ((payload, picture), (size, group)) in pictures.iter().zip([(1905, true), (290, false)]) {
        assert_eq!(payload.len(), size);
        assert_eq!(picture.is_drawing_group(), group);
        assert_eq!(picture.is_drawing(), !group);
        assert!(!picture.is_continuation());
        assert_eq!(picture.to_payload(), *payload);
    }
    let drawing = pictures[0].1.drawing();
    let png = drawing
        .windows(8)
        .position(|w| w == b"\x89PNG\r\n\x1a\n")
        .unwrap();
    assert_eq!(png, 149);
    assert_eq!(
        &drawing[png + 1698 - 12..png + 1698],
        b"\0\0\0\0IEND\xaeB`\x82"
    );
    assert_eq!(pictures[1].1.drawing().len(), 276);
}

#[test]
fn native_consolidation_metadata_and_binary_pictures_survive_property_publication() {
    let bytes = native();
    let r = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = r.xls_worksheet(3).unwrap();
    let c = sheet.consolidation().unwrap();
    assert_eq!(c.function(), ConsolidationFunction::Sum);
    assert!(!c.uses_left_labels(), "native DCon stores fLeftCat=0");
    assert!(c.uses_top_labels() && c.creates_links());
    assert_eq!(c.sources().len(), 2);
    for (s, name) in c
        .sources()
        .iter()
        .zip(["\u{2}Source_A_015", "\u{2}Source_B_015"])
    {
        let ConsolidationSource::CellRange { range, file } = s else {
            panic!("local range")
        };
        assert_eq!(
            (
                range.first_row(),
                range.last_row(),
                range.first_column(),
                range.last_column()
            ),
            (0, 2, 0, 1)
        );
        assert_eq!(file.encoded_path(), name);
    }
    for (row, v) in (3..=8).zip([10., 5., 15., 20., 7., 27.]) {
        assert_eq!(
            sheet.get_cell(row, 2).unwrap().value(),
            &CellValue::Float(v)
        );
    }
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_metadata().unwrap();
    e.properties_mut()
        .unwrap()
        .set("Title", "Native worksheet metadata preservation")
        .unwrap();
    let commit = e.commit().unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(f.path()).unwrap();
    let reopened = Package::open(f.path()).unwrap();
    assert_eq!(
        reopened.metadata().unwrap().title.as_deref(),
        Some("Native worksheet metadata preservation")
    );
    let after = reopened.to_bytes().unwrap();
    let a = streams(&bytes);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (path, data) in &a {
        if path != &vec!["\u{5}SummaryInformation".to_owned()] {
            assert!(b[path] == *data, "changed {path:?}");
        }
    }
    let read = Reader::new(Cursor::new(&after)).unwrap();
    caches(&read);
    assert_eq!(read.xls_worksheet(3).unwrap().consolidation(), Some(c));
    assert_eq!(
        reopened
            .apply_metadata(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

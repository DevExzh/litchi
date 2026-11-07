//! Native BIFF framing, BOF metadata and future-record preservation.
use litchi_biff::{Encoder, Error, Limits, Record, Records, Resource};
use litchi_core::sheet::{CellValue, WorkbookTrait};
use litchi_xls::cell_values::{Reference, Value};
use litchi_xls::{
    Package,
    raw::Reader,
    records::{BiffVersion, BofRecord},
};
use std::{collections::BTreeMap, io::Cursor};
const META: &str = "excel-metadata.xls";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
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

fn workbook(bytes: &[u8]) -> Vec<u8> {
    streams(bytes)[&vec!["Workbook".into()]].clone()
}
fn frames(bytes: &[u8]) -> Vec<litchi_biff::RecordRef<'_>> {
    Records::new(bytes, Limits::default())
        .unwrap()
        .collect::<Result<Vec<_>, _>>()
        .unwrap()
}
#[test]
fn native_biff_frames_preserve_all_offsets_headers_payloads_and_continuations() {
    for (name, length, count, eof) in [
        (META, 195402, 564, 6),
        ("excel-date1904-breaks.xls", 18386, 415, 2),
        ("excel-conditional-drawings.xls", 24253, 406, 2),
        ("excel-recorded-macro.xlt", 16193, 334, 2),
    ] {
        let original = native(name);
        let stream = workbook(&original);
        assert_eq!(stream.len(), length);
        let records = frames(&stream);
        assert_eq!(records.len(), count);
        assert_eq!(records.iter().filter(|r| r.kind().get() == 10).count(), eof);
        let mut out = Encoder::new(Limits {
            max_records: count,
            max_input_bytes: length,
            max_output_bytes: length,
            ..Default::default()
        })
        .unwrap();
        let mut offset = 0;
        for r in &records {
            assert_eq!(r.offset(), offset);
            assert_eq!(r.encoded(), &stream[offset..offset + 4 + r.payload().len()]);
            let owned = r.own(Limits::default()).unwrap();
            assert_eq!(owned.as_bytes(), r.encoded());
            assert_eq!(
                Record::open(owned.as_bytes().to_vec(), Limits::default())
                    .unwrap()
                    .as_bytes(),
                r.encoded()
            );
            out.push_ref(*r).unwrap();
            offset += r.encoded().len();
        }
        assert_eq!(offset, length);
        assert_eq!(out.count(), count);
        assert!(out.finish() == stream);
        if name == META {
            assert!(
                records
                    .iter()
                    .any(|r| r.kind().get() == 0x3c && r.payload().len() == 8212)
            );
            assert_eq!(
                records.iter().filter(|r| r.kind().get() == 0x896).count(),
                1
            );
        }
        assert!(Package::from_bytes(&original).unwrap().to_bytes().unwrap() == original);
    }
}
#[test]
fn native_bof_substream_types_and_workbook_date_modes_are_independent() {
    for (name, date1904, sheets) in [(META, false, 5), ("excel-date1904-breaks.xls", true, 1)] {
        let bytes = native(name);
        let stream = workbook(&bytes);
        let records = frames(&stream);
        let starts = records
            .iter()
            .filter(|r| r.kind().get() == 0x809)
            .collect::<Vec<_>>();
        assert_eq!(starts.len(), sheets + 1);
        for (i, r) in starts.iter().enumerate() {
            let bof = BofRecord::parse(r.payload()).unwrap();
            assert_eq!(bof.version, BiffVersion::Biff8);
            assert_eq!(bof.substream_type, if i == 0 { 5 } else { 0x10 });
            assert_eq!(&r.payload()[4..8], &[0x5a, 0x4f, 0xcd, 7]);
        }
        let date = records
            .iter()
            .filter(|r| r.kind().get() == 0x22)
            .collect::<Vec<_>>();
        assert_eq!(date.len(), 1);
        assert_eq!(date[0].payload(), [u8::from(date1904), 0]);
        let reader = Reader::new(Cursor::new(&bytes)).unwrap();
        assert_eq!(reader.is_1904_date_system(), date1904);
        assert_eq!(reader.sheets().len(), sheets);
    }
}
#[test]
fn native_biff_limits_and_truncation_fail_at_exact_record_boundaries() {
    let bytes = native(META);
    let stream = workbook(&bytes);
    let records = frames(&stream);
    let short = Limits {
        max_input_bytes: stream.len() - 1,
        ..Default::default()
    };
    assert!(matches!(
        Records::new(&stream, short),
        Err(Error::LimitExceeded {
            resource: Resource::InputBytes,
            ..
        })
    ));
    let mut read = Records::new(
        &stream,
        Limits {
            max_records: records.len() - 1,
            ..Default::default()
        },
    )
    .unwrap();
    for _ in 0..records.len() - 1 {
        read.next().unwrap().unwrap();
    }
    assert!(matches!(
        read.next(),
        Some(Err(Error::LimitExceeded {
            resource: Resource::RecordCount,
            ..
        }))
    ));
    assert!(read.next().is_none());
    assert!(read.next().is_none());
    let mut read = Records::new(
        &stream,
        Limits {
            max_record_bytes: 8219,
            ..Default::default()
        },
    )
    .unwrap();
    let error = read.find_map(Result::err).unwrap();
    assert!(matches!(
        error,
        Error::LimitExceeded {
            resource: Resource::RecordBytes,
            observed: 8220,
            maximum: 8219
        }
    ));
    let mut out = Encoder::new(Limits {
        max_output_bytes: stream.len() - 1,
        ..Default::default()
    })
    .unwrap();
    for r in &records {
        let before = out.len();
        if let Err(error) = out.push_ref(*r) {
            assert!(matches!(
                error,
                Error::LimitExceeded {
                    resource: Resource::OutputBytes,
                    ..
                }
            ));
            assert_eq!(out.len(), before);
            assert_eq!(out.as_bytes(), &stream[..r.offset()]);
            break;
        }
    }
    assert_eq!(out.count(), records.len() - 1);
    assert_eq!(
        out.len(),
        stream.len() - records.last().unwrap().encoded().len()
    );
    let first = records[0].encoded();
    assert!(matches!(
        Record::open(first[..first.len() - 1].to_vec(), Limits::default()),
        Err(Error::TruncatedPayload {
            offset: 0,
            declared: 16,
            available: 15,
            ..
        })
    ));
    assert!(matches!(
        Record::open(first[..3].to_vec(), Limits::default()),
        Err(Error::TruncatedHeader {
            offset: 0,
            available: 3
        })
    ));
    assert_eq!(native(META), bytes);
}
#[test]
fn native_scalar_edit_keeps_every_future_frame_and_all_unowned_streams_exact() {
    let bytes = native(META);
    let p = Package::from_bytes(&bytes).unwrap();
    let at = Reference::new(1, 1).unwrap();
    assert_eq!(
        p.workbook()
            .worksheet("Print_015")
            .unwrap()
            .unwrap()
            .cell_value(at)
            .unwrap(),
        Some(CellValue::Float(42.))
    );
    let before = streams(&bytes);
    let key = vec!["Workbook".into()];
    let a = frames(&before[&key]);
    let future = a
        .iter()
        .filter(|r| r.kind().get() >= 0x800 && r.kind().get() != 0x809)
        .map(|r| (r.kind().get(), r.encoded().to_vec()))
        .collect::<Vec<_>>();
    assert!(future.len() > 100);
    for (kind, count) in [(0x863, 1), (0x87d, 62), (0x892, 47), (0x896, 1)] {
        assert_eq!(future.iter().filter(|(k, _)| *k == kind).count(), count);
    }
    let mut e = p.edit().unwrap();
    e.set_value("Print_015", at, Value::Number(43.)).unwrap();
    let c = e.commit().unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    c.package().save(file.path()).unwrap();
    let out = Package::open(file.path()).unwrap();
    let after = streams(&out.to_bytes().unwrap());
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, data) in &before {
        if name != &key {
            assert!(after[name] == *data, "changed {name:?}");
        }
    }
    let b = frames(&after[&key]);
    assert_eq!(a.len(), b.len());
    assert_eq!(
        b.iter()
            .filter(|r| r.kind().get() >= 0x800 && r.kind().get() != 0x809)
            .map(|r| (r.kind().get(), r.encoded().to_vec()))
            .collect::<Vec<_>>(),
        future
    );
    let changed = a
        .iter()
        .zip(&b)
        .filter(|(a, b)| a.encoded() != b.encoded())
        .collect::<Vec<_>>();
    assert_eq!(changed.len(), 1);
    assert_eq!(changed[0].0.offset(), changed[0].1.offset());
    assert_eq!(changed[0].0.kind(), changed[0].1.kind());
    assert!([0x203, 0x27e, 0xbd].contains(&changed[0].0.kind().get()));
    assert_eq!(
        out.workbook()
            .worksheet("Print_015")
            .unwrap()
            .unwrap()
            .cell_value(at)
            .unwrap(),
        Some(CellValue::Float(43.))
    );
    assert!(
        out.apply(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

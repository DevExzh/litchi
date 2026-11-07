//! Explicit inert BIFF metadata on native Excel parents; no RTD provider or phonetic rendering.
use litchi_biff::{Limits, Records};
use litchi_core::sheet::Cell as _;
use litchi_xls::{Package, real_time_data as rtd};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
const PARENTS: [&str; 3] = [
    "excel-metadata.xls",
    "excel-recorded-macro.xlt",
    "excel-date1904-breaks.xls",
];
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|name| {
            let data = ole
                .open_stream(&name.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (name, data)
        })
        .collect()
}
fn workbook(p: &Package) -> Vec<u8> {
    streams(&p.to_bytes().unwrap())[&vec!["Workbook".into()]].clone()
}
fn frame(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut r = kind.to_le_bytes().to_vec();
    r.extend(u16::try_from(payload.len()).unwrap().to_le_bytes());
    r.extend_from_slice(payload);
    r
}
fn records(bytes: &[u8]) -> Vec<(usize, u16, Vec<u8>)> {
    Records::new(bytes, Limits::default())
        .unwrap()
        .map(|r| {
            let r = r.unwrap();
            (r.offset(), r.kind().get(), r.payload().to_vec())
        })
        .collect()
}
fn rtd_boundary(data: &[u8]) -> usize {
    records(data).iter().find(|r| r.1 == 0x01c1).unwrap().0
}

/// The chosen fixtures use BoundSheet8, Index and ExtSST absolute offsets.
/// Keep every other source byte intact when installing the explicit record
/// derivative. This is test setup via the public CFB owner, not a package RTD API.
fn publish_range(p: &Package, start: usize, end: usize, replacement: &[u8]) -> Package {
    use litchi_ole_common::object;
    let before = p.to_bytes().unwrap();
    let original = workbook(p);
    let mut output = original[..start].to_vec();
    output.extend_from_slice(replacement);
    output.extend_from_slice(&original[end..]);
    let delta = i64::try_from(replacement.len()).unwrap() - i64::try_from(end - start).unwrap();
    let mut checks = Vec::new();
    for (at, kind, payload) in records(&output) {
        let fields = match kind {
            0x0085 => vec![0],
            0x020b => (12..payload.len()).step_by(4).collect(),
            0x00ff => (2..payload.len()).step_by(8).collect(),
            _ => Vec::new(),
        };
        for field in fields {
            let old = u32::from_le_bytes(payload[field..field + 4].try_into().unwrap()) as usize;
            if old == 0 {
                continue;
            }
            assert!(
                old < start || old >= end,
                "a source offset targets replaced metadata"
            );
            assert!(old + 4 <= original.len());
            let new = if old >= end {
                usize::try_from(i64::try_from(old).unwrap() + delta).unwrap()
            } else {
                old
            };
            output[at + 4 + field..at + 8 + field]
                .copy_from_slice(&u32::try_from(new).unwrap().to_le_bytes());
            checks.push((old, new));
        }
    }
    for (old, new) in checks {
        assert_eq!(&output[new..new + 4], &original[old..old + 4]);
    }
    let source = object::Package::open(before.clone()).unwrap();
    let mut edit = source.edit().unwrap();
    edit.put_stream(&["Workbook".into()], output.clone())
        .unwrap();
    let c = edit.commit().unwrap();
    let q = Package::from_bytes(c.patch().after()).unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    q.save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    assert_eq!(workbook(&q), output);
    let after = q.to_bytes().unwrap();
    assert_eq!(
        cfb_metadata::directory(&before),
        cfb_metadata::directory(&after)
    );
    let a = streams(&before);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if name != ["Workbook"] {
            assert!(b.get(&name) == Some(&data), "changed stream {name:?}");
        }
    }
    let original_cell = p
        .reader_ref()
        .unwrap()
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .unwrap()
        .value()
        .clone();
    assert_eq!(
        q.reader_ref()
            .unwrap()
            .xls_worksheet(0)
            .unwrap()
            .get_cell(0, 0)
            .unwrap()
            .value(),
        &original_cell
    );
    assert_eq!(
        q.reader_ref().unwrap().date_system(),
        p.reader_ref().unwrap().date_system()
    );
    assert_eq!(c.patch().apply(&source).unwrap().patch().after(), after);
    let target = object::Package::open(after).unwrap();
    assert_eq!(
        c.patch().inverse().apply(&target).unwrap().patch().after(),
        before
    );
    q
}
fn payload(record: &rtd::Record) -> Vec<u8> {
    let mut bytes = vec![0x13, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    bytes.extend(record.common_prefix_len.to_le_bytes());
    let units = record
        .topic_segments
        .iter()
        .map(|s| s.encode_utf16().count() + 1)
        .sum::<usize>();
    bytes.extend(u32::try_from(units).unwrap().to_le_bytes());
    bytes.push(1);
    for s in &record.topic_segments {
        let units = s.encode_utf16().collect::<Vec<_>>();
        bytes.extend(u16::try_from(units.len()).unwrap().to_le_bytes());
        bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
    match &record.value {
        rtd::Value::Text(value) => {
            let units = value.encode_utf16().collect::<Vec<_>>();
            bytes.extend((if units.len() < 256 { 2u32 } else { 0x1000 }).to_le_bytes());
            bytes.extend(u32::try_from(units.len()).unwrap().to_le_bytes());
            bytes.push(1);
            bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
        },
        rtd::Value::Integer(value) => {
            bytes.extend(0x800u32.to_le_bytes());
            bytes.extend(value.to_le_bytes());
        },
        rtd::Value::Number(value) => {
            bytes.extend(1u32.to_le_bytes());
            bytes.extend(value.to_le_bytes());
        },
        rtd::Value::Boolean(value) => {
            bytes.extend(4u32.to_le_bytes());
            bytes.extend(u32::from(*value).to_le_bytes());
        },
        rtd::Value::Error(value) => {
            bytes.extend(0x10u32.to_le_bytes());
            bytes.extend(value.to_le_bytes());
        },
    }
    for c in &record.cells {
        bytes.extend(c.row.to_le_bytes());
        bytes.extend(u16::from(c.column).to_le_bytes());
        bytes.extend(c.sheet_index.to_le_bytes());
    }
    bytes
}
fn framed(record: &rtd::Record, chunk: usize) -> Vec<u8> {
    let bytes = payload(record);
    let mut output = Vec::new();
    for (i, data) in bytes.chunks(chunk).enumerate() {
        if i == 0 {
            output.extend(frame(0x0813, data));
        } else {
            let mut continued = vec![0x12, 0x08, 0, 0x80];
            continued.extend_from_slice(data);
            output.extend(frame(0x0812, &continued));
        }
    }
    output
}
fn topic(value: rtd::Value) -> rtd::Record {
    rtd::Record {
        common_prefix_len: 0,
        topic_segments: vec!["Native.Server".into(), String::new(), "東京🍎A".into()],
        topic: "Native.Server東京🍎A".into(),
        value,
        cells: vec![rtd::Cell::new(120, 25, 0).unwrap()],
    }
}
fn publish_rtd(p: &Package, c: &rtd::Commit) -> Package {
    let before = rtd::Snapshot::parse(&workbook(p)[..rtd_boundary(&workbook(p))]).unwrap();
    assert_eq!(
        c.patch().apply(&before).unwrap().bytes(),
        c.package().bytes()
    );
    assert_eq!(
        c.patch().inverse().apply(c.package()).unwrap().bytes(),
        before.bytes()
    );
    let unowned = |s: &rtd::Snapshot| {
        s.unknown_records()
            .map(|r| (r.record_type(), r.payload().to_vec()))
            .collect::<Vec<_>>()
    };
    assert_eq!(unowned(&before), unowned(c.package()));
    publish_range(p, 0, before.bytes().len(), c.package().bytes())
}

fn phonetic_bytes(p: &Package) -> Vec<u8> {
    let frames = records(&workbook(p));
    let first = frames.iter().position(|r| r.1 == 0x00ef).unwrap();
    let mut bytes = frame(frames[first].1, &frames[first].2);
    for r in &frames[first + 1..] {
        if r.1 != 0x003c {
            break;
        }
        bytes.extend(frame(r.1, &r.2));
    }
    bytes
}

#[test]
fn native_xls_rtd_continuation_headers_and_cached_values_read_without_refresh() {
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.reader_ref().unwrap().real_time_data().is_empty());
        let record = topic(rtd::Value::Text("値🍎".repeat(1800)));
        let original = workbook(&p);
        let at = rtd_boundary(&original);
        let mut prefix = original[..at].to_vec();
        prefix.extend(framed(&record, 123));
        let snapshot = rtd::Snapshot::parse(&prefix).unwrap();
        assert_eq!(snapshot.records(), std::slice::from_ref(&record));
        let q = publish_range(&p, 0, at, &prefix);
        assert_eq!(
            q.reader_ref().unwrap().real_time_data(),
            std::slice::from_ref(&record)
        );
        let source = rtd::Snapshot::parse(&workbook(&q)[..rtd_boundary(&workbook(&q))]).unwrap();
        let mut unchanged = source.edit();
        unchanged.set_value(0, record.value.clone()).unwrap();
        let unchanged = unchanged.commit().unwrap();
        assert!(unchanged.patch().is_noop());
        assert_eq!(unchanged.package().bytes(), source.bytes());
        let mut edit = source.edit();
        edit.set_value(0, rtd::Value::Integer(76)).unwrap();
        let mut invalid = edit.records()[0].clone();
        invalid.common_prefix_len = 1;
        assert!(
            matches!(edit.replace(0,invalid),Err(litchi_xls::Error::InvalidRecord {record_type:0x0813,message}) if message=="first RealTimeData record declares a shared prefix")
        );
        assert_eq!(edit.records()[0].value, rtd::Value::Integer(76));
        let c = edit.commit().unwrap();
        let changed = publish_rtd(&q, &c);
        assert_eq!(
            changed.reader_ref().unwrap().real_time_data()[0].value,
            rtd::Value::Integer(76)
        );
        let source =
            rtd::Snapshot::parse(&workbook(&changed)[..rtd_boundary(&workbook(&changed))]).unwrap();
        let mut e = source.edit();
        e.remove(0).unwrap();
        let c = e.commit().unwrap();
        let removed = publish_rtd(&changed, &c);
        assert!(removed.reader_ref().unwrap().real_time_data().is_empty());
        assert_eq!(workbook(&removed), original);
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_xls_rtd_shared_prefix_uses_utf16_units_and_rejects_half_surrogates() {
    let bytes = native(PARENTS[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    let first = topic(rtd::Value::Boolean(true));
    let prefix = "Native.Server東京🍎";
    let mut second = rtd::Record {
        common_prefix_len: u32::try_from(prefix.encode_utf16().count()).unwrap(),
        topic_segments: vec![String::new(), String::new(), "B".into()],
        topic: format!("{prefix}B"),
        value: rtd::Value::Number(42.5),
        cells: Vec::new(),
    };
    assert_eq!(
        rtd::Record::parse(&payload(&second), Some(&first.topic)).unwrap(),
        second
    );
    let original = workbook(&p);
    let at = rtd_boundary(&original);
    let mut source = original[..at].to_vec();
    source.extend(framed(&first, 8224));
    source.extend(framed(&second, 8224));
    let q = publish_range(&p, 0, at, &source);
    assert_eq!(
        q.reader_ref().unwrap().real_time_data(),
        [first.clone(), second.clone()]
    );
    second.common_prefix_len -= 1;
    assert!(
        matches!(rtd::Record::parse(&payload(&second),Some(&first.topic)),Err(litchi_xls::Error::InvalidRecord {record_type:0x0813,message}) if message=="RealTimeData ichSamePrefix splits a UTF-16 surrogate pair")
    );
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_xls_rtd_authoring_emits_real_continuefrt_headers() {
    let bytes = native(PARENTS[1]);
    let p = Package::from_bytes(&bytes).unwrap();
    let original = workbook(&p);
    let at = rtd_boundary(&original);
    let source = rtd::Snapshot::parse(&original[..at]).unwrap();
    let mut edit = source.edit();
    let value = topic(rtd::Value::Text("値🍎".repeat(1800)));
    edit.push(value.clone()).unwrap();
    let c = edit.commit().unwrap();
    let mut continues = 0;
    for (_, kind, bytes) in records(c.package().bytes()) {
        if kind == 0x0812 {
            continues += 1;
            assert_eq!(&bytes[..4], &[0x12, 0x08, 0, 0]);
            assert!(bytes.len() <= 8224);
        }
    }
    assert!(continues > 0);
    let q = publish_rtd(&p, &c);
    assert_eq!(
        q.reader_ref().unwrap().real_time_data(),
        std::slice::from_ref(&value)
    );
    let mut builder = litchi_xls::writer::Builder::new();
    builder.add_worksheet("Sheet1").unwrap();
    builder.add_real_time_data(value.clone()).unwrap();
    let authored = builder.build().unwrap();
    let mut data = Vec::new();
    authored.write_to(&mut data).unwrap();
    let authored = Package::from_bytes(&data).unwrap();
    let mut count = 0;
    for (_, kind, body) in records(&workbook(&authored)) {
        if kind == 0x0812 {
            count += 1;
            assert_eq!(&body[..4], &[0x12, 0x08, 0, 0]);
        }
    }
    assert!(count > 0);
    assert_eq!(
        authored.reader_ref().unwrap().real_time_data(),
        std::slice::from_ref(&value)
    );
    assert_eq!(native(PARENTS[1]), bytes);
}

#[test]
fn native_xls_rtd_malformed_continuation_headers_are_not_interpreted_as_topic_bytes() {
    let bytes = native(PARENTS[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    let book = workbook(&p);
    let mut source = book[..rtd_boundary(&book)].to_vec();
    source.extend(framed(
        &topic(rtd::Value::Text("stored text".repeat(400))),
        64,
    ));
    assert_eq!(rtd::Snapshot::parse(&source).unwrap().records().len(), 1);
    let valid = publish_range(&p, 0, rtd_boundary(&book), &source);
    let valid_book = workbook(&valid);
    let valid_owner = litchi_ole_common::object::Package::open(valid.to_bytes().unwrap()).unwrap();
    let at = records(&source).iter().find(|r| r.1 == 0x0812).unwrap().0 + 4;
    for (header, message) in [
        (
            [0x13, 0x08, 0, 0],
            "ContinueFrt header has the wrong record type",
        ),
        (
            [0x12, 0x08, 1, 0],
            "ContinueFrt reference and alert flags must be zero",
        ),
    ] {
        let mut invalid = source.clone();
        invalid[at..at + 4].copy_from_slice(&header);
        assert!(
            matches!(rtd::Snapshot::parse(&invalid),Err(litchi_xls::Error::InvalidRecord {record_type:0x0812,message:m}) if m==message)
        );
        let mut invalid_book = valid_book.clone();
        invalid_book[at..at + 4].copy_from_slice(&header);
        let mut edit = valid_owner.edit().unwrap();
        edit.put_stream(&["Workbook".into()], invalid_book).unwrap();
        let malformed = edit.commit().unwrap();
        assert!(
            matches!(Package::from_bytes(malformed.patch().after()),Err(litchi_xls::package::Error::Reader(litchi_xls::Error::InvalidRecord {record_type:0x0812,message:m})) if m==message)
        );
    }
    assert_eq!(p.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_xls_phonetic_info_continuations_survive_ordinary_cell_publication() {
    use litchi_xls::{
        PhoneticAlignment, PhoneticFormat, PhoneticInfo, PhoneticRange, PhoneticType,
    };
    let expected = PhoneticInfo::try_new(
        PhoneticFormat::new(0, PhoneticType::Hiragana, PhoneticAlignment::Center),
        (0..4000u16)
            .map(|r| PhoneticRange::new(r, r, 0, 4).unwrap())
            .collect(),
    )
    .unwrap();
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Carrier").unwrap();
    builder
        .set_phonetic_info(sheet, Some(expected.clone()))
        .unwrap();
    let mut bytes = Vec::new();
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    let carrier = Package::from_bytes(&bytes).unwrap();
    let mut addition = phonetic_bytes(&carrier);
    assert!(records(&addition).len() >= 3);
    let flags = u16::from_le_bytes(addition[6..8].try_into().unwrap()) | 0xa500;
    addition[6..8].copy_from_slice(&flags.to_le_bytes());
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(
            p.reader_ref()
                .unwrap()
                .xls_worksheet(0)
                .unwrap()
                .phonetic_info()
                .is_none()
        );
        let before = workbook(&p);
        let frames = records(&before);
        let at = frames.iter().filter(|r| r.1 == 0x000a).nth(1).unwrap().0;
        let q = publish_range(&p, at, at, &addition);
        let read = q
            .reader_ref()
            .unwrap()
            .xls_worksheet(0)
            .unwrap()
            .phonetic_info()
            .unwrap()
            .clone();
        assert_eq!(read.ranges(), expected.ranges());
        assert_eq!(
            (
                read.format().font_index(),
                read.format().phonetic_type(),
                read.format().alignment()
            ),
            (0, PhoneticType::Hiragana, PhoneticAlignment::Center)
        );
        assert_eq!(phonetic_bytes(&q), addition);
        let mut e = q.edit().unwrap();
        let address = litchi_xls::cell_values::Reference::new(120, 25).unwrap();
        e.insert_cell(
            "Sheet1",
            address,
            litchi_xls::cell_values::Value::Number(76.0),
        )
        .unwrap();
        let c = e.commit().unwrap();
        let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        c.package().save(output.path()).unwrap();
        let changed = Package::open(output.path()).unwrap();
        assert_eq!(
            changed
                .reader_ref()
                .unwrap()
                .xls_worksheet(0)
                .unwrap()
                .phonetic_info(),
            Some(&read)
        );
        assert_eq!(phonetic_bytes(&changed), addition);
        assert_eq!(
            changed
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(address)
                .unwrap(),
            Some(litchi_core::sheet::CellValue::Float(76.0))
        );
        let a = streams(&q.to_bytes().unwrap());
        let b = streams(&changed.to_bytes().unwrap());
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (name, data) in a {
            if name != ["Workbook"] {
                assert!(b.get(&name) == Some(&data));
            }
        }
        assert_eq!(
            changed
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            q.to_bytes().unwrap()
        );
        assert_eq!(native(file), bytes);
    }
}

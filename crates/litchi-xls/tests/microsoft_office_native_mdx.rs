//! Inert MDX derivatives on native XLS parents; no OLAP connections or cube evaluation.
use litchi_xls::{Package, mdx_metadata::*};
#[path = "../../../tools/test_support/native_xls_workbook.rs"]
mod native_xls_workbook;
use native_xls_workbook::{cfb_metadata, frame, native, publish_range, records, streams, workbook};

const PARENTS: [&str; 3] = [
    "excel-metadata.xls",
    "excel-recorded-macro.xlt",
    "excel-date1904-breaks.xls",
];
fn header(kind: u16) -> Vec<u8> {
    let mut bytes = vec![0; 12];
    bytes[..2].copy_from_slice(&kind.to_le_bytes());
    bytes
}
fn wide(text: &str) -> Vec<u8> {
    let units = text.encode_utf16().collect::<Vec<_>>();
    let mut out = u16::try_from(units.len()).unwrap().to_le_bytes().to_vec();
    out.extend(units.into_iter().flat_map(u16::to_le_bytes));
    out
}
fn metadata() -> MdxMetadata {
    let mut m = MdxMetadata::default();
    m.add_info(MdtInfo {
        flags: MdtInfoFlags::from_bits(0x9000_0034),
        name: "XLDAPR".into(),
    });
    for s in [
        "NeverConnected".to_owned(),
        format!("[Sales].[{}]", "東京🍎".repeat(1800)),
        "Caption".into(),
        "[KPI].[Revenue]".into(),
        "[Measures].[Revenue]".into(),
        "{[Products].[A], [Products].[B]}".into(),
    ] {
        m.add_string(s).unwrap();
    }
    m.add_record(MdxMetadataRecord::Tuple(MdxTuple {
        connection_name_index: 0,
        function: CubeFunction::CubeMember,
        string_indexes: vec![1],
    }))
    .unwrap();
    m.add_record(MdxMetadataRecord::Set(MdxSet {
        connection_name_index: 0,
        function: CubeFunction::CubeSet,
        sort_order: MdxSetSortOrder::AlphaAscending,
        set_definition_index: 5,
        string_indexes: vec![1],
    }))
    .unwrap();
    m.add_record(MdxMetadataRecord::Prop(MdxProp {
        connection_name_index: 0,
        function: CubeFunction::CubeMemberProperty,
        member_index: 1,
        property_index: 2,
    }))
    .unwrap();
    m.add_record(MdxMetadataRecord::Kpi(MdxKpi {
        connection_name_index: 0,
        function: CubeFunction::CubeKpiProperty,
        kpi_property: KpiProperty::Trend,
        kpi_name_index: 3,
        member_kpi_index: 4,
    }))
    .unwrap();
    for index in 0..4 {
        m.add_block(Mdb {
            entries: vec![MdxMetadataDir {
                info_index: 1,
                metadata_index: index,
            }],
        })
        .unwrap();
    }
    m
}
/// Independent MS-XLS wire data, deliberately not obtained from the writer under test.
fn logical_records(m: &MdxMetadata) -> Vec<(u16, Vec<u8>)> {
    let mut out = Vec::new();
    let mut info = header(0x0884);
    info.extend(0x9000_0034u32.to_le_bytes());
    info.extend(wide("XLDAPR"));
    out.push((0x0884, info));
    for s in m.strings() {
        let mut p = header(0x0885);
        p.extend(wide(s));
        out.push((0x0885, p));
    }
    for (kind, function, extra, indexes) in [
        (0x0886, 1, vec![], vec![1, 1]),
        (0x0887, 3, vec![3], vec![5, 1, 1]),
        (0x0888, 6, vec![], vec![1, 2]),
        (0x0889, 7, vec![4], vec![3, 4]),
    ] {
        let mut p = header(kind);
        p.extend(0i32.to_le_bytes());
        p.push(function);
        p.extend(extra);
        for index in indexes {
            p.extend(i32::to_le_bytes(index));
        }
        out.push((kind, p));
    }
    for index in 0..4u32 {
        let mut p = header(0x088a);
        p.extend(1i32.to_le_bytes());
        p.extend(index.to_le_bytes());
        out.push((0x088a, p));
    }
    out
}
fn encoded(m: &MdxMetadata) -> Vec<u8> {
    let mut out = Vec::new();
    for (kind, logical) in logical_records(m) {
        // Nonuniform continuation boundaries can split a UTF-16 code unit.
        let first = logical.len().min(8223);
        out.extend(frame(kind, &logical[..first]));
        for chunk in logical[first..].chunks(8211) {
            let mut body = header(0x087f);
            body[3] = 0x80; // FrtFlags reserved bits MUST be ignored on read.
            body.extend_from_slice(chunk);
            out.extend(frame(0x087f, &body));
        }
    }
    out
}
fn boundary(data: &[u8]) -> usize {
    let frames = records(data);
    let last = frames.iter().rposition(|r| r.1 == 0x0085).unwrap();
    frames[last + 1].0
}
fn mdx_bytes(p: &Package) -> Vec<u8> {
    let data = workbook(p);
    let frames = records(&data);
    let first = frames.iter().position(|r| r.1 == 0x0884).unwrap();
    let last = frames[first..]
        .iter()
        .position(|r| !matches!(r.1, 0x0884..=0x088a | 0x087f))
        .unwrap()
        + first;
    data[frames[first].0..frames[last].0].to_vec()
}
fn add(p: &Package, bytes: &[u8]) -> Package {
    let at = boundary(&workbook(p));
    publish_range(p, at, at, bytes)
}
fn preserve_other_streams(a: &Package, b: &Package) {
    let a = a.to_bytes().unwrap();
    let b = b.to_bytes().unwrap();
    assert_eq!(cfb_metadata::directory(&a), cfb_metadata::directory(&b));
    let a = streams(&a);
    let b = streams(&b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name != ["Workbook"] {
            assert_eq!(b[&name], bytes);
        }
    }
}

#[test]
fn native_xls_mdx_continuations_reconstruct_all_record_kinds_and_preserve_cell_edits() {
    let expected = metadata();
    let addition = encoded(&expected);
    for file in PARENTS {
        let bytes = native(file);
        let parent = Package::from_bytes(&bytes).unwrap();
        assert!(parent.reader_ref().unwrap().mdx_metadata().is_empty());
        let source = add(&parent, &addition);
        assert_eq!(source.reader_ref().unwrap().mdx_metadata(), &expected);
        assert_eq!(mdx_bytes(&source), addition);
        let flags = source.reader_ref().unwrap().mdx_metadata().infos()[0].flags;
        assert!(flags.preserved_on_edit() && flags.copied_with_cell() && flags.paste_all());
        assert!(flags.preserved_on_coerce());
        assert!(!flags.is_cell_metadata());
        assert_eq!(flags.bits(), 0x9000_0034);
        let mut edit = source.edit().unwrap();
        let address = litchi_xls::cell_values::Reference::new(120, 25).unwrap();
        edit.insert_cell(
            "Sheet1",
            address,
            litchi_xls::cell_values::Value::Number(88.0),
        )
        .unwrap();
        let changed = edit.commit().unwrap();
        let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        changed.package().save(output.path()).unwrap();
        let reopened = Package::open(output.path()).unwrap();
        assert_eq!(reopened.reader_ref().unwrap().mdx_metadata(), &expected);
        assert_eq!(mdx_bytes(&reopened), addition);
        assert_eq!(
            reopened
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(address)
                .unwrap(),
            Some(litchi_core::sheet::CellValue::Float(88.0))
        );
        preserve_other_streams(&source, &reopened);
        assert_eq!(
            source
                .apply(changed.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            reopened.to_bytes().unwrap()
        );
        assert_eq!(
            reopened
                .apply(&changed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            source.to_bytes().unwrap()
        );
        assert_eq!(native(file), bytes);
        assert_eq!(parent.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_xls_mdx_builder_emits_continuefrt12_headers() {
    let expected = metadata();
    let mut builder = litchi_xls::writer::Builder::new();
    builder.add_worksheet("Carrier").unwrap();
    builder.set_mdx_metadata(Some(expected.clone()));
    let mut bytes = Vec::new();
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    let built = Package::from_bytes(&bytes).unwrap();
    let addition = mdx_bytes(&built);
    let frames = records(&addition);
    assert!(frames.iter().any(|r| r.1 == 0x087f));
    let mut logical = Vec::<(u16, Vec<u8>)>::new();
    for (_, kind, body) in frames {
        if kind == 0x087f {
            assert_eq!(&body[..12], header(0x087f));
            assert!(body.len() <= 8224);
            logical.last_mut().unwrap().1.extend_from_slice(&body[12..]);
        } else {
            logical.push((kind, body));
        }
    }
    assert_eq!(logical, logical_records(&expected));
    assert_eq!(built.reader_ref().unwrap().mdx_metadata(), &expected);
    let parent = Package::from_bytes(&native(PARENTS[1])).unwrap();
    let derivative = add(&parent, &addition);
    assert_eq!(derivative.reader_ref().unwrap().mdx_metadata(), &expected);
}

#[test]
fn native_xls_mdx_builder_places_metadata_after_sheet_catalog_before_country() {
    let mut builder = litchi_xls::writer::Builder::new();
    builder.add_worksheet("Carrier").unwrap();
    let mut expected = MdxMetadata::default();
    expected.add_info(MdtInfo {
        flags: MdtInfoFlags::from_bits(0x34),
        name: "XLDAPR".into(),
    });
    builder.set_mdx_metadata(Some(expected.clone()));
    let mut bytes = Vec::new();
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    let built = Package::from_bytes(&bytes).unwrap();
    let frames = records(&workbook(&built));
    let catalog = frames.iter().rposition(|r| r.1 == 0x0085).unwrap();
    let mdx = frames.iter().position(|r| r.1 == 0x0884).unwrap();
    let country = frames.iter().position(|r| r.1 == 0x008c).unwrap();
    assert_eq!(mdx, catalog + 1);
    assert!(mdx < country);
    let parent = Package::from_bytes(&native(PARENTS[0])).unwrap();
    let derivative = add(&parent, &mdx_bytes(&built));
    assert_eq!(derivative.reader_ref().unwrap().mdx_metadata(), &expected);
}

fn built(m: MdxMetadata) -> Package {
    let mut builder = litchi_xls::writer::Builder::new();
    builder.add_worksheet("Carrier").unwrap();
    builder.set_mdx_metadata(Some(m));
    let mut bytes = Vec::new();
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    Package::from_bytes(&bytes).unwrap()
}

#[test]
fn native_xls_mdx_continuation_sizes_include_the_twelve_byte_header() {
    let bytes = native(PARENTS[0]);
    let parent = Package::from_bytes(&bytes).unwrap();
    for (units, continuation_count) in [(4105, 0), (4106, 1), (8211, 1), (8212, 2)] {
        let mut expected = MdxMetadata::default();
        expected.add_info(MdtInfo {
            flags: MdtInfoFlags::from_bits(0x34),
            name: "XLDAPR".into(),
        });
        expected.add_string("a".repeat(units)).unwrap();
        let carrier = built(expected.clone());
        let addition = mdx_bytes(&carrier);
        let frames = records(&addition);
        assert_eq!(
            frames.iter().filter(|r| r.1 == 0x087f).count(),
            continuation_count
        );
        let mut logical = Vec::new();
        for (_, kind, data) in frames {
            assert!(data.len() <= 8224);
            if kind == 0x0885 {
                logical.extend(data);
            } else if kind == 0x087f {
                assert_eq!(&data[..12], header(0x087f));
                logical.extend_from_slice(&data[12..]);
            }
        }
        let mut exact = header(0x0885);
        exact.extend(wide(&"a".repeat(units)));
        assert_eq!(logical, exact);
        let derivative = add(&parent, &addition);
        assert_eq!(derivative.reader_ref().unwrap().mdx_metadata(), &expected);
    }
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_xls_mdx_metadata_lifecycle_keeps_indexes_and_rejects_invalid_staging() {
    for file in PARENTS {
        let bytes = native(file);
        let parent = Package::from_bytes(&bytes).unwrap();
        let source = add(&parent, &encoded(&metadata()));
        let mut expected = source.reader_ref().unwrap().mdx_metadata().clone();
        expected.add_string("[Products].[追加🍎]".into()).unwrap();
        expected
            .add_record(MdxMetadataRecord::Tuple(MdxTuple {
                connection_name_index: 0,
                function: CubeFunction::CubeValue,
                string_indexes: vec![6],
            }))
            .unwrap();
        expected
            .add_block(Mdb {
                entries: vec![MdxMetadataDir {
                    info_index: 1,
                    metadata_index: 4,
                }],
            })
            .unwrap();
        let staged = expected.clone();
        assert!(
            matches!(expected.add_record(MdxMetadataRecord::Tuple(MdxTuple {
            connection_name_index: 99, function: CubeFunction::CubeValue, string_indexes: vec![1],
        })), Err(litchi_xls::Error::InvalidRecord { record_type: 0x0885, message })
            if message == "MDXStrIndex 99 is out of range of the MDXStr collection (7 strings)")
        );
        for (entry, message) in [
            (
                MdxMetadataDir {
                    info_index: 0,
                    metadata_index: 0,
                },
                "MDir.imdt 0 is out of range of the MDTInfo collection",
            ),
            (
                MdxMetadataDir {
                    info_index: 1,
                    metadata_index: 5,
                },
                "MDir.mdd 5 is out of range of the MDX metadata record collection",
            ),
        ] {
            assert!(matches!(expected.add_block(Mdb { entries: vec![entry] }),
                Err(litchi_xls::Error::InvalidRecord { record_type: 0x088a, message: m }) if m == message));
        }
        assert_eq!(expected, staged);
        let start = boundary(&workbook(&source));
        let changed = publish_range(
            &source,
            start,
            start + mdx_bytes(&source).len(),
            &mdx_bytes(&built(expected.clone())),
        );
        assert_eq!(changed.reader_ref().unwrap().mdx_metadata(), &expected);
        let start = boundary(&workbook(&changed));
        let removed = publish_range(&changed, start, start + mdx_bytes(&changed).len(), &[]);
        assert!(removed.reader_ref().unwrap().mdx_metadata().is_empty());
        assert_eq!(workbook(&removed), workbook(&parent));
        assert_eq!(native(file), bytes);
    }
}

fn with_workbook(source: &Package, bytes: Vec<u8>) -> Vec<u8> {
    let owner = litchi_ole_common::object::Package::open(source.to_bytes().unwrap()).unwrap();
    let mut edit = owner.edit().unwrap();
    edit.put_stream(&["Workbook".into()], bytes).unwrap();
    edit.commit().unwrap().patch().after().to_vec()
}

#[test]
fn native_xls_mdx_continuation_headers_and_cross_record_indexes_are_validated() {
    let original = native(PARENTS[0]);
    let parent = Package::from_bytes(&original).unwrap();
    let expected = metadata();
    let source = add(&parent, &encoded(&expected));
    let source_bytes = source.to_bytes().unwrap();
    let book = workbook(&source);
    let frames = records(&book);
    let continuation = frames.iter().find(|r| r.1 == 0x087f).unwrap().0 + 4;
    // An active range is checked; an inactive range and reserved flags are ignored.
    for active in [true, false] {
        let mut modified = book.clone();
        modified[continuation + 2] = u8::from(active);
        let range = if active {
            [0, 0, 3, 0, 0, 0, 2, 0]
        } else {
            [0xff; 8]
        };
        modified[continuation + 4..continuation + 12].copy_from_slice(&range);
        let q = Package::from_bytes(&with_workbook(&source, modified)).unwrap();
        assert_eq!(q.reader_ref().unwrap().mdx_metadata(), &expected);
    }
    let mut cases = Vec::new();
    let mut bad = book.clone();
    bad[continuation] = 0x80;
    cases.push((
        bad,
        0x087f,
        "ContinueFrt12 header has the wrong record type",
    ));
    let mut bad = book.clone();
    bad[continuation + 2] = 2;
    cases.push((bad, 0x087f, "ContinueFrt12 alert flag must be zero"));
    for range in [
        [5, 0, 4, 0, 0, 0, 2, 0],
        [0, 0, 3, 0, 3, 0, 2, 0],
        [0, 0, 3, 0, 0, 0, 0, 1],
    ] {
        let mut bad = book.clone();
        bad[continuation + 2] = 1;
        bad[continuation + 4..continuation + 12].copy_from_slice(&range);
        cases.push((
            bad,
            0x087f,
            "ContinueFrt12 reference is outside the BIFF8 cell grid",
        ));
    }
    for (kind, offset, value, error_kind, message) in [
        (
            0x0886,
            12,
            99u32,
            0x0885,
            "MDXStrIndex 99 is out of range of the MDXStr collection (6 strings)",
        ),
        (
            0x088a,
            12,
            0,
            0x088a,
            "MDir.imdt 0 is out of range of the MDTInfo collection",
        ),
        (
            0x088a,
            16,
            4,
            0x088a,
            "MDir.mdd 4 is out of range of the MDX metadata record collection",
        ),
    ] {
        let start = frames.iter().find(|r| r.1 == kind).unwrap().0 + 4 + offset;
        let mut bad = book.clone();
        bad[start..start + 4].copy_from_slice(&value.to_le_bytes());
        cases.push((bad, error_kind, message));
    }
    for (bad, kind, message) in cases {
        assert!(matches!(Package::from_bytes(&with_workbook(&source, bad)),
            Err(litchi_xls::package::Error::Reader(litchi_xls::Error::InvalidRecord { record_type, message: m })) if record_type == kind && m == message));
    }
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    assert_eq!(native(PARENTS[0]), original);
}

//! Explicit inert DDE/OLE streams on native parents, published through the public OPC owner.
use litchi_core::sheet::{Cell as _, WorkbookTrait, Worksheet};
use litchi_opc::constants::relationship_type as rel;
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI, TargetMode};
use litchi_xlsb::raw::{Encoder, Kind, Limits, Records, kind};
use litchi_xlsb::{Package, SourceBackedWorkbook, external_link as link};
use std::{collections::BTreeMap, io::Cursor, path::Path};
const PARENTS: [&str; 3] = [
    "excel-metadata.xlsb",
    "excel-recorded-macro.xlsb",
    "excel-links-objects.xlsb",
];
const DDE: &str = "/xl/externalLinks/nativeDde.bin";
const OLE: &str = "/xl/externalLinks/nativeOle.bin";
const BOOK: &str = "/xl/workbook.bin";
const OLE_RID: &str = "NativeOleTarget";
const TARGET: &str = "https://example.invalid/native-source.ole";
const TYPE: &str = "application/vnd.ms-excel.externalLink";
const POLICY: link::ExternalLinkLimits = link::ExternalLinkLimits::DEFAULT;
fn uri(s: &str) -> PackURI {
    PackURI::new(s).unwrap()
}
fn native(f: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(f),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn wide(s: &str) -> Vec<u8> {
    let u = s.encode_utf16().collect::<Vec<_>>();
    let mut b = u32::try_from(u.len()).unwrap().to_le_bytes().to_vec();
    b.extend(u.into_iter().flat_map(u16::to_le_bytes));
    b
}
fn framed(k: Kind, b: &[u8]) -> Vec<u8> {
    let mut v = Vec::new();
    Encoder::new(&mut v, Limits::DEFAULT)
        .write_record(k, b)
        .unwrap();
    v
}
fn matrix() -> link::ValueMatrix {
    use link::CachedValue::*;
    link::ValueMatrix::new(
        2,
        3,
        vec![
            Empty,
            Number(42.5),
            Boolean(true),
            Error(link::ErrorValue::DivisionByZero),
            String("東京🍎".into()),
            Number(0.0),
        ],
    )
    .unwrap()
}
fn model(ole: bool) -> link::Link {
    if ole {
        link::Link::ole_with_items(
            TARGET,
            "Example.Document.1",
            vec![
                link::OleItem::new("Report")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_icon(true)
                    .with_cached_values(matrix()),
            ],
        )
        .unwrap()
    } else {
        link::Link::dde_with_items(
            "Native.Server",
            "Topic 東京",
            vec![
                link::DdeItem::new("StdDocumentName")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_ole_support(true)
                    .with_cached_values(matrix()),
            ],
        )
        .unwrap()
    }
}
/// Independent MS-XLSB record payloads, including ignored high flag bits.
fn wire(ole: bool) -> Vec<u8> {
    let mut begin = (if ole { 2u16 } else { 1 }).to_le_bytes().to_vec();
    begin.extend(wide(if ole { OLE_RID } else { "Native.Server" }));
    begin.extend(wide(if ole {
        "Example.Document.1"
    } else {
        "Topic 東京"
    }));
    let mut b = framed(kind::BEGIN_SUP_BOOK, &begin);
    b.extend(framed(Kind::new(0x3ffe).unwrap(), b"opaque-link-record"));
    b.extend(framed(
        kind::SUP_NAME_START,
        &wide(if ole { "Report" } else { "StdDocumentName" }),
    ));
    b.extend(framed(
        kind::SUP_NAME_BITS,
        &[if ole { 0xf6 } else { 0xce }, 0xa5, 0, 0, 0, 0, 0x81],
    ));
    b.extend(framed(
        kind::SUP_NAME_VALUE_START,
        &[2u32.to_le_bytes(), 3u32.to_le_bytes()].concat(),
    ));
    b.extend(framed(kind::SUP_NAME_NIL, &[]));
    b.extend(framed(kind::SUP_NAME_NUM, &42.5f64.to_le_bytes()));
    b.extend(framed(kind::SUP_NAME_BOOL, &[1]));
    b.extend(framed(kind::SUP_NAME_ERROR, &[7]));
    b.extend(framed(kind::SUP_NAME_STRING, &wide("東京🍎")));
    b.extend(framed(kind::SUP_NAME_NUM, &0f64.to_le_bytes()));
    b.extend(framed(kind::SUP_NAME_VALUE_END, &[]));
    b.extend(framed(kind::SUP_NAME_END, &[]));
    b.extend(framed(kind::END_SUP_BOOK, &[]));
    b
}
type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph(p: &OpcPackage) -> Graph {
    let edges = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().into(),
                    r.reltype().into(),
                    r.target_ref().into(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    let mut g = BTreeMap::from([("/".into(), (String::new(), edges(p.rels())))]);
    for part in p.iter_parts() {
        g.insert(
            part.partname().as_str().into(),
            (part.content_type().into(), edges(part.rels())),
        );
    }
    g
}
fn publish(p: &Package, c: &litchi_xlsb::opc::Commit, allowed: &[&str]) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    c.package().save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    let a = members(p);
    let b = members(&q);
    for n in a.keys().chain(b.keys()) {
        if !allowed.contains(&n.as_str()) {
            assert_eq!(a.get(n), b.get(n), "changed {n}");
        }
    }
    assert_eq!(members(c.package()), b);
    assert_eq!(members(p.apply_opc(c.patch()).unwrap().package()), b);
    let inverse = q.apply_opc(&c.patch().inverse()).unwrap();
    assert_eq!(members(inverse.package()), a);
    assert_eq!(
        graph(inverse.package().opc_package()),
        graph(p.opc_package())
    );
    let before = p.workbook().unwrap();
    let after = q.workbook().unwrap();
    assert_eq!(before.worksheet_names(), after.worksheet_names());
    for name in before.worksheet_names() {
        let s = before.worksheet_by_name(name).unwrap();
        let t = after.worksheet_by_name(name).unwrap();
        for c in s.cells() {
            let c = c.unwrap();
            let c = c.as_ref();
            let d = t.get_cell(c.row(), c.column()).unwrap();
            assert_eq!(c.value(), d.value());
            assert_eq!(c.formula_bytes(), d.formula_bytes());
            assert_eq!(c.style_id(), d.style_id());
        }
    }
    q
}
fn append_sources(data: &[u8]) -> Vec<u8> {
    let frames = Records::new(data, Limits::DEFAULT)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    let mut extra = framed(kind::SUP_BOOK_SRC, &wide("rIdNativeDde"));
    extra.extend(framed(kind::SUP_BOOK_SRC, &wide("rIdNativeOle")));
    let at = if let Some(r) = frames.iter().find(|r| r.kind() == kind::EXTERN_SHEET) {
        r.offset()
    } else {
        let mut block = framed(kind::BEGIN_EXTERNALS, &[]);
        block.extend(extra);
        block.extend(framed(kind::EXTERN_SHEET, &0u32.to_le_bytes()));
        block.extend(framed(kind::END_EXTERNALS, &[]));
        extra = block;
        let i = frames
            .iter()
            .position(|r| r.kind() == kind::END_BUNDLE_SHS)
            .unwrap();
        frames[i + 1].offset()
    };
    [data[..at].to_vec(), extra, data[at..].to_vec()].concat()
}
fn seed(file: &str) -> Package {
    let bytes = native(file);
    let p = Package::from_bytes(&bytes).unwrap();
    let old = graph(p.opc_package());
    let book = append_sources(p.opc_package().get_part(&uri(BOOK)).unwrap().blob());
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.try_add_part(Box::new(BlobPart::new(uri(DDE), TYPE.into(), wire(false))))?;
        let mut ole = BlobPart::new(uri(OLE), TYPE.into(), wire(true));
        ole.rels_mut().try_add_relationship(
            rel::OLE_OBJECT.into(),
            TARGET.into(),
            OLE_RID.into(),
            TargetMode::External,
        )?;
        opc.try_add_part(Box::new(ole))?;
        let owner = opc.get_part_mut(&uri(BOOK))?;
        owner.set_blob(book.clone());
        owner.rels_mut().try_add_relationship(
            rel::EXTERNAL_LINK.into(),
            "externalLinks/nativeDde.bin".into(),
            "rIdNativeDde".into(),
            TargetMode::Internal,
        )?;
        owner.rels_mut().try_add_relationship(
            rel::EXTERNAL_LINK.into(),
            "externalLinks/nativeOle.bin".into(),
            "rIdNativeOle".into(),
            TargetMode::Internal,
        )?;
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let q = publish(
        &p,
        &c,
        &[
            "[Content_Types].xml",
            "xl/workbook.bin",
            "xl/_rels/workbook.bin.rels",
            "xl/externalLinks/nativeDde.bin",
            "xl/externalLinks/nativeOle.bin",
            "xl/externalLinks/_rels/nativeOle.bin.rels",
        ],
    );
    let mut expected = old;
    let edges = &mut expected.get_mut(BOOK).unwrap().1;
    edges.extend([
        (
            "rIdNativeDde".into(),
            rel::EXTERNAL_LINK.into(),
            "externalLinks/nativeDde.bin".into(),
            false,
        ),
        (
            "rIdNativeOle".into(),
            rel::EXTERNAL_LINK.into(),
            "externalLinks/nativeOle.bin".into(),
            false,
        ),
    ]);
    edges.sort();
    expected.insert(DDE.into(), (TYPE.into(), vec![]));
    expected.insert(
        OLE.into(),
        (
            TYPE.into(),
            vec![(OLE_RID.into(), rel::OLE_OBJECT.into(), TARGET.into(), true)],
        ),
    );
    assert_eq!(graph(q.opc_package()), expected);
    assert_eq!(q.opc_package().get_part(&uri(BOOK)).unwrap().blob(), book);
    let mut links = p.workbook().unwrap().external_links();
    links.extend([model(false), model(true)]);
    check_links(&q, &links);
    assert_eq!(native(file), bytes);
    q
}
fn check_links(p: &Package, expected: &[link::Link]) {
    assert_eq!(p.workbook().unwrap().external_links(), expected);
    let lazy = SourceBackedWorkbook::from_reader(Cursor::new(p.to_bytes().unwrap())).unwrap();
    let links = lazy.external_links().unwrap();
    assert_eq!(
        links.iter().map(|v| v.metadata()).collect::<Vec<_>>(),
        expected.iter().collect::<Vec<_>>()
    );
}
fn snapshot(p: &Package, path: &str) -> link::Snapshot {
    link::Snapshot::read(p.opc_package().get_part(&uri(path)).unwrap().blob(), POLICY).unwrap()
}
fn publish_stream(
    p: &Package,
    path: &str,
    source: &link::Snapshot,
    after: &link::Snapshot,
    patch: &link::Patch,
) -> Package {
    assert_eq!(
        patch.apply(source.source_bytes(), POLICY).unwrap(),
        after.source_bytes()
    );
    assert_eq!(
        patch.inverse().apply(after.source_bytes(), POLICY).unwrap(),
        source.source_bytes()
    );
    assert_eq!(source.unknown_records(), after.unknown_records());
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(&uri(path))?
            .set_blob(after.source_bytes().to_vec());
        Ok(())
    })
    .unwrap();
    let c = edit.commit().unwrap();
    let q = publish(p, &c, &[&path[1..]]);
    assert_eq!(graph(p.opc_package()), graph(q.opc_package()));
    let mut expected = p.workbook().unwrap().external_links();
    let i = expected.len() - if path == DDE { 2 } else { 1 };
    expected[i] = if path == DDE {
        after.link().clone()
    } else {
        after.resolved_link(TARGET).unwrap()
    };
    check_links(&q, &expected);
    q
}

#[test]
fn native_xlsb_dde_ole_derivatives_read_exact_caches_and_preserve_noops() {
    for file in PARENTS {
        let p = seed(file);
        for (path, ole) in [(DDE, false), (OLE, true)] {
            let s = snapshot(&p, path);
            assert_eq!(s.unknown_records().len(), 1);
            assert_eq!(s.unknown_records()[0].payload(), b"opaque-link-record");
            if ole {
                assert_eq!(s.relationship_id(), Some(OLE_RID));
                assert_eq!(s.resolved_link(TARGET).unwrap(), model(true));
            } else {
                assert!(s.relationship_id().is_none());
                assert_eq!(s.link(), &model(false));
            }
            let (noop, patch) = s.draft().finish().unwrap();
            assert!(patch.is_noop());
            assert_eq!(noop.source_bytes(), s.source_bytes());
            let canonical = link::write_external_link_stream(
                &model(ole),
                if ole { Some(OLE_RID) } else { None },
                POLICY,
            )
            .unwrap();
            let read = link::Snapshot::read(&canonical, POLICY).unwrap();
            assert_eq!(read.link(), s.link());
        }
    }
}

#[test]
fn native_xlsb_dde_item_cache_lifecycle_keeps_opaque_records_and_failed_staging() {
    for file in PARENTS {
        let p = seed(file);
        let s = snapshot(&p, DDE);
        let mut edit = s.draft();
        edit.set_source("Native.Changed").unwrap();
        edit.set_dde_topic("Topic 更新🍎").unwrap();
        edit.set_dde_flags("stddocumentname", false, true, true)
            .unwrap();
        edit.set_dde_cache(
            "StdDocumentName",
            Some(
                link::ValueMatrix::new(
                    1,
                    1,
                    vec![link::CachedValue::String("東京🍎".repeat(1024))],
                )
                .unwrap(),
            ),
        )
        .unwrap();
        edit.upsert_dde_item(
            link::DdeItem::new("Status")
                .unwrap()
                .with_cached_values(matrix()),
        )
        .unwrap();
        let staged = edit.link().clone();
        assert!(matches!(
            edit.set_relationship_id("no"),
            Err(link::Error::InvalidFormula(
                "DDE external links do not have a relationship identifier"
            ))
        ));
        assert_eq!(edit.link(), &staged);
        let (changed, patch) = edit.finish().unwrap();
        let q = publish_stream(&p, DDE, &s, &changed, &patch);
        let bits = Records::new(changed.source_bytes(), Limits::DEFAULT)
            .map(Result::unwrap)
            .find(|r| r.kind() == kind::SUP_NAME_BITS)
            .unwrap();
        assert_eq!(bits.payload(), [0xcc, 0xa5, 0, 0, 0, 0, 0x81]);
        assert!(matches!(
            patch.apply(changed.source_bytes(), POLICY),
            Err(link::Error::PatchConflict)
        ));
        let mut edit = changed.draft();
        edit.set_dde_cache("StdDocumentName", None).unwrap();
        assert!(edit.remove_dde_item("STATUS").unwrap().is_some());
        let (cleared, patch) = edit.finish().unwrap();
        let q = publish_stream(&q, DDE, &changed, &cleared, &patch);
        assert_eq!(cleared.link().dde_items().len(), 1);
        assert!(cleared.link().dde_items()[0].cached_values().is_none());
        let mut edit = cleared.draft();
        let item = edit.remove_dde_item("StdDocumentName").unwrap().unwrap();
        let (empty, patch) = edit.finish().unwrap();
        let q = publish_stream(&q, DDE, &cleared, &empty, &patch);
        assert!(empty.link().dde_items().is_empty());
        let mut edit = empty.draft();
        edit.upsert_dde_item(item).unwrap();
        let (restored, patch) = edit.finish().unwrap();
        publish_stream(&q, DDE, &empty, &restored, &patch);
        assert_eq!(restored.source_bytes(), cleared.source_bytes());
    }
}

#[test]
fn native_xlsb_ole_item_cache_lifecycle_keeps_external_relationships_inert() {
    for file in PARENTS {
        let p = seed(file);
        let s = snapshot(&p, OLE);
        let mut edit = s.draft();
        edit.set_ole_program_id("Example.Document.2").unwrap();
        edit.set_ole_flags("report", false, true, false).unwrap();
        edit.set_ole_cache(
            "Report",
            Some(
                link::ValueMatrix::new(
                    1,
                    2,
                    vec![
                        link::CachedValue::Number(76.0),
                        link::CachedValue::String("Stored 🍎".into()),
                    ],
                )
                .unwrap(),
            ),
        )
        .unwrap();
        edit.upsert_ole_item(
            link::OleItem::new("Added")
                .unwrap()
                .with_cached_values(matrix()),
        )
        .unwrap();
        let staged = edit.link().clone();
        assert!(matches!(
            edit.set_dde_topic("wrong kind"),
            Err(link::Error::InvalidFormula(
                "DDE topic can only be edited on a DDE link"
            ))
        ));
        assert_eq!(edit.link(), &staged);
        let (changed, patch) = edit.finish().unwrap();
        let q = publish_stream(&p, OLE, &s, &changed, &patch);
        let bits = Records::new(changed.source_bytes(), Limits::DEFAULT)
            .map(Result::unwrap)
            .find(|r| r.kind() == kind::SUP_NAME_BITS)
            .unwrap();
        assert_eq!(bits.payload(), [0xd4, 0xa5, 0, 0, 0, 0, 0x81]);
        let mut edit = changed.draft();
        edit.set_ole_cache("Report", None).unwrap();
        assert!(edit.remove_ole_item("ADDED").unwrap().is_some());
        let (cleared, patch) = edit.finish().unwrap();
        let q = publish_stream(&q, OLE, &changed, &cleared, &patch);
        assert_eq!(cleared.link().ole_items().len(), 1);
        assert!(cleared.link().ole_items()[0].cached_values().is_none());
        let mut edit = cleared.draft();
        let item = edit.remove_ole_item("Report").unwrap().unwrap();
        let (empty, patch) = edit.finish().unwrap();
        let q = publish_stream(&q, OLE, &cleared, &empty, &patch);
        assert!(empty.link().ole_items().is_empty());
        let mut edit = empty.draft();
        edit.upsert_ole_item(item).unwrap();
        let (restored, patch) = edit.finish().unwrap();
        publish_stream(&q, OLE, &empty, &restored, &patch);
        assert_eq!(restored.source_bytes(), cleared.source_bytes());
    }
}

#[test]
fn native_xlsb_dde_ole_host_relationship_errors_refuse_publication_atomically() {
    use litchi_xlsb::package::PackageError;
    let p = seed(PARENTS[2]);
    let before = members(&p);
    for (case, expected) in [
        (0, "DDE external link must not contain relationships"),
        (1, "external data relationship is missing"),
        (2, "external data relationship is internal"),
        (3, "relationship has invalid type"),
    ] {
        let mut edit = p.edit_opc().unwrap();
        edit.update(|opc| {
            if case == 0 {
                opc.get_part_mut(&uri(DDE))?
                    .rels_mut()
                    .try_add_relationship(
                        rel::OLE_OBJECT.into(),
                        TARGET.into(),
                        "Unexpected".into(),
                        TargetMode::External,
                    )?;
            } else {
                let rels = opc.get_part_mut(&uri(OLE))?.rels_mut();
                rels.remove(OLE_RID).unwrap();
                if case == 2 {
                    rels.try_add_relationship(
                        rel::OLE_OBJECT.into(),
                        "../worksheets/sheet1.bin".into(),
                        OLE_RID.into(),
                        TargetMode::Internal,
                    )?;
                }
                if case == 3 {
                    rels.try_add_relationship(
                        rel::EXTERNAL_LINK_PATH.into(),
                        TARGET.into(),
                        OLE_RID.into(),
                        TargetMode::External,
                    )?;
                }
            }
            Ok(())
        })
        .unwrap();
        assert!(matches!(edit.commit(),Err(PackageError::InvalidFormula(m)) if m==expected));
        assert_eq!(members(&p), before);
    }
    // A syntactically valid detached ID edit needs the corresponding host relationship.
    let original = snapshot(&p, OLE);
    let mut edit = original.draft();
    edit.set_relationship_id("MissingTarget").unwrap();
    let (changed, _) = edit.finish().unwrap();
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(&uri(OLE))?
            .set_blob(changed.source_bytes().to_vec());
        Ok(())
    })
    .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(PackageError::InvalidFormula(
            "external data relationship is missing"
        ))
    ));
    assert_eq!(members(&p), before);
}

#[test]
fn native_xlsb_dde_ole_cached_matrix_count_errors_refuse_host_publication() {
    use litchi_xlsb::package::PackageError;
    let p = seed(PARENTS[0]);
    let before = members(&p);
    for path in [DDE, OLE] {
        let original = snapshot(&p, path);
        let mut corrupt = Vec::new();
        let mut writer = Encoder::new(&mut corrupt, Limits::DEFAULT);
        for r in Records::new(original.source_bytes(), Limits::DEFAULT) {
            let r = r.unwrap();
            let mut b = r.payload().to_vec();
            if r.kind() == kind::SUP_NAME_VALUE_START {
                b[4..8].copy_from_slice(&4u32.to_le_bytes());
            }
            writer.write_record(r.kind(), &b).unwrap();
        }
        assert!(matches!(
            link::Snapshot::read(&corrupt, POLICY),
            Err(link::Error::InvalidFormula(
                "invalid cached external value matrix"
            ))
        ));
        let mut edit = p.edit_opc().unwrap();
        edit.update(|opc| {
            opc.get_part_mut(&uri(path))?.set_blob(corrupt);
            Ok(())
        })
        .unwrap();
        assert!(matches!(
            edit.commit(),
            Err(PackageError::InvalidFormula(
                "invalid cached external value matrix"
            ))
        ));
        assert_eq!(members(&p), before);
    }
}

#[test]
fn native_xlsb_dde_ole_fresh_writer_streams_reopen_on_macro_parent() {
    use litchi_xlsb::builder::{Builder, Worksheet};
    let mut builder = Builder::new();
    builder.add_worksheet(Worksheet::new("Carrier"));
    builder.add_external_link(model(false)).unwrap();
    builder.add_external_link(model(true)).unwrap();
    let carrier = builder.build().unwrap();
    check_links(&carrier, &[model(false), model(true)]);
    let p = seed(PARENTS[1]);
    let mut edit = p.edit_opc().unwrap();
    for (i, path) in [(1, DDE), (2, OLE)] {
        let part = carrier
            .opc_package()
            .get_part(&uri(&format!("/xl/externalLinks/externalLink{i}.bin")))
            .unwrap();
        let source = link::Snapshot::read(part.blob(), POLICY).unwrap();
        assert!(source.unknown_records().is_empty());
        let mut draft = source.draft();
        if path == OLE {
            draft.set_relationship_id(OLE_RID).unwrap();
        }
        let (authored, _) = draft.finish().unwrap();
        let canonical = link::write_external_link_stream(
            &model(path == OLE),
            if path == OLE { Some(OLE_RID) } else { None },
            POLICY,
        )
        .unwrap();
        assert_eq!(authored.source_bytes(), canonical);
        edit.update(|opc| {
            opc.get_part_mut(&uri(path))?.set_blob(canonical);
            Ok(())
        })
        .unwrap();
    }
    let c = edit.commit().unwrap();
    let q = publish(&p, &c, &[&DDE[1..], &OLE[1..]]);
    check_links(&q, &[model(false), model(true)]);
    assert_eq!(graph(p.opc_package()), graph(q.opc_package()));
}

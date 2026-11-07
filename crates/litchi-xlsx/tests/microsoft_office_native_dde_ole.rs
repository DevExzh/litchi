//! Inert DDE/OLE catalog derivatives; native external workbook and embedded content stay exact.
use litchi_xlsx::Package;
use litchi_xlsx::external_links::{
    Dde, DdeItem, DdeValue, DdeValueType, DdeValues, ItemSource, Link, Ole, OleItem, Target,
};
use std::collections::BTreeMap;
const FILE: &str = "excel-links-objects.xlsx";
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(FILE),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
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
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}

type Relationships = BTreeMap<String, (String, String, bool)>;
fn graph(p: &Package) -> BTreeMap<String, (String, Relationships)> {
    let opc = p.clone().into_plain_opc();
    let edges = |rels: &litchi_opc::Relationships| -> Relationships {
        rels.iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    (
                        r.reltype().to_owned(),
                        r.target_ref().to_owned(),
                        r.is_external(),
                    ),
                )
            })
            .collect()
    };
    let mut out = BTreeMap::new();
    out.insert("/".into(), (String::new(), edges(opc.rels())));
    for part in opc.iter_parts() {
        out.insert(
            part.partname().to_string(),
            (part.content_type().to_owned(), edges(part.rels())),
        );
    }
    out
}

fn values() -> DdeValues {
    DdeValues {
        rows: 2,
        columns: 2,
        values: vec![
            DdeValue {
                value_type: DdeValueType::Number,
                raw_value: "42".into(),
            },
            DdeValue {
                value_type: DdeValueType::String,
                raw_value: "cached & inert".into(),
            },
            DdeValue {
                value_type: DdeValueType::Boolean,
                raw_value: "1".into(),
            },
            DdeValue {
                value_type: DdeValueType::Error,
                raw_value: "#N/A".into(),
            },
        ],
    }
}
fn dde() -> Link {
    Link::Dde(Dde {
        service: "Litchi.InertService".into(),
        topic: "NativeSource".into(),
        items: vec![DdeItem {
            name: Some("Native.Item".into()),
            use_ole: false,
            advise: false,
            prefer_picture: false,
            values: Some(values()),
        }],
    })
}
fn ole() -> Link {
    Link::Ole(Ole {
        target: Target {
            relationship_id: "rIdInertOle".into(),
            target: "https://example.invalid/inert.docx".into(),
            relationship_type: litchi_opc::constants::relationship_type::OLE_OBJECT.into(),
        },
        program_id: "Word.Document.12".into(),
        items: vec![OleItem {
            source: ItemSource::Office2010,
            name: "NativeOleItem".into(),
            icon: true,
            advise: false,
            prefer_picture: true,
            values: Some(values()),
        }],
    })
}
#[test]
fn native_external_catalog_append_update_remove_preserves_original_link_and_all_payloads() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let original = p.edit_external_links().unwrap().entries().unwrap().to_vec();
    assert_eq!(original.len(), 1);
    assert!(matches!(original[0].link, Link::Workbook(_)));
    let mut e = p.edit_external_links().unwrap();
    assert_eq!(e.insert(dde()).unwrap(), 1);
    assert_eq!(e.insert(ole()).unwrap(), 2);
    let created = e.commit().unwrap();
    let out = save(created.package());
    let after = out.to_bytes().unwrap();
    assert_eq!(
        changed(&bytes, &after),
        [
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/externalLinks/_rels/externalLink3.xml.rels",
            "xl/externalLinks/externalLink2.xml",
            "xl/externalLinks/externalLink3.xml",
            "xl/workbook.xml"
        ]
    );
    let entries = out
        .edit_external_links()
        .unwrap()
        .entries()
        .unwrap()
        .to_vec();
    assert_eq!(entries.len(), 3);
    assert_eq!(
        out.workbook().external_reference_ids(),
        entries
            .iter()
            .map(|e| e.relationship_id.clone())
            .collect::<Vec<_>>()
    );
    assert_eq!(entries[0], original[0]);
    assert_eq!(entries[1].link, dde());
    assert_eq!(entries[2].link, ole());
    assert!(
        members(
            &out.apply_external_links(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&bytes)
    );
    assert!(matches!(
        out.apply_external_links(created.patch()),
        Err(litchi_xlsx::Error::PatchConflict { .. })
    ));
    let mut e = out.edit_external_links().unwrap();
    assert!(!e.set(1, dde()).unwrap());
    assert!(e.commit().unwrap().package().to_bytes().unwrap() == after);
    let mut changed_dde = dde();
    let Link::Dde(ref mut v) = changed_dde else {
        unreachable!()
    };
    v.topic = "UpdatedTopic".into();
    v.items[0].values.as_mut().unwrap().values[0].raw_value = "43".into();
    let mut changed_ole = ole();
    let Link::Ole(ref mut v) = changed_ole else {
        unreachable!()
    };
    v.items[0].icon = false;
    v.items[0].values.as_mut().unwrap().values[1].raw_value = "updated cache".into();
    let mut e = out.edit_external_links().unwrap();
    e.set(1, changed_dde.clone()).unwrap();
    e.set(2, changed_ole.clone()).unwrap();
    let updated = e.commit().unwrap();
    let newer = save(updated.package());
    let newbytes = newer.to_bytes().unwrap();
    assert_eq!(
        changed(&after, &newbytes),
        [
            "xl/externalLinks/externalLink2.xml",
            "xl/externalLinks/externalLink3.xml"
        ]
    );
    let entries = newer
        .edit_external_links()
        .unwrap()
        .entries()
        .unwrap()
        .to_vec();
    assert_eq!(entries[0], original[0]);
    assert_eq!(entries[1].link, changed_dde);
    assert_eq!(entries[2].link, changed_ole);
    let mut e = newer.edit_external_links().unwrap();
    assert!(e.remove(2).unwrap().is_some());
    assert!(e.remove(1).unwrap().is_some());
    assert!(e.remove(1).unwrap().is_none());
    let removed = e.commit().unwrap();
    let restored = save(removed.package());
    assert_eq!(
        restored.edit_external_links().unwrap().entries().unwrap(),
        original
    );
    assert_eq!(
        changed(&bytes, &restored.to_bytes().unwrap()),
        ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]
    );
    assert_eq!(graph(&restored), graph(&p));
    assert!(
        members(
            &restored
                .apply_external_links(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&newbytes)
    );
    assert!(
        members(
            &newer
                .apply_external_links(&updated.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&after)
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_external_catalog_invalid_dde_matrix_and_ole_metadata_leave_prior_staging_intact() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_external_links().unwrap();
    e.insert(dde()).unwrap();
    let staged = e.entries().unwrap().to_vec();
    let mut invalid = dde();
    let Link::Dde(ref mut v) = invalid else {
        unreachable!()
    };
    v.items[0].values.as_mut().unwrap().values.pop();
    assert!(e.set(1, invalid).is_err());
    assert_eq!(e.entries().unwrap(), staged);
    let mut invalid = ole();
    let Link::Ole(ref mut v) = invalid else {
        unreachable!()
    };
    v.program_id.clear();
    assert!(e.insert(invalid).is_err());
    assert_eq!(e.entries().unwrap(), staged);
    let mut invalid = ole();
    let Link::Ole(ref mut v) = invalid else {
        unreachable!()
    };
    v.items[0].source = ItemSource::SpreadsheetMl;
    assert!(e.insert(invalid).is_err());
    assert_eq!(e.entries().unwrap(), staged);
    let out = save(e.commit().unwrap().package());
    assert_eq!(
        out.edit_external_links().unwrap().entries().unwrap(),
        staged
    );
    assert_eq!(
        changed(&bytes, &out.to_bytes().unwrap()),
        [
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/externalLinks/externalLink2.xml",
            "xl/workbook.xml"
        ]
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

#[test]
fn native_referenced_link_removal_and_stale_workbook_source_are_refused() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_external_links().unwrap();
    e.remove(0).unwrap();
    assert!(
        matches!(e.commit(),Err(litchi_xlsx::Error::Refused(ref r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent)
    );
    assert!(p.to_bytes().unwrap() == bytes);
    let mut e = p.edit_external_links().unwrap();
    e.insert(dde()).unwrap();
    let c = e.commit().unwrap();
    let mut opc = p.clone().into_plain_opc();
    let name = opc.main_document_part().unwrap().partname().clone();
    let part = opc.get_part_mut(&name).unwrap();
    let xml = std::str::from_utf8(part.blob()).unwrap();
    assert_eq!(xml.matches("</workbook>").count(), 1);
    part.set_blob(
        xml.replace("</workbook>", "<!-- changed source --></workbook>")
            .into_bytes(),
    );
    let changed = Package::from_opc(opc).unwrap();
    assert_eq!(
        changed.workbook().external_reference_ids(),
        p.workbook().external_reference_ids()
    );
    assert!(matches!(
        changed.apply_external_links(c.patch()),
        Err(litchi_xlsx::Error::PatchConflict { .. })
    ));
    assert!(p.to_bytes().unwrap() == bytes);
}

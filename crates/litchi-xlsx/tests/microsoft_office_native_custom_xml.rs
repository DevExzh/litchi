//! Shared Custom XML graph edits on native Excel parents; no XML mapping or schema execution.
use litchi_ooxml_common::custom_xml::{Conformance, Edit, NewItem, NewProps, Props, Snapshot};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceLimits};
use litchi_xlsx::Package;
use std::{collections::BTreeMap, path::Path};
const ID_A: &str = "{64000000-1111-4111-8111-111111111111}";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn members(p: &OpcPackage) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        PackageWriter::to_bytes(p).unwrap(),
        SourceLimits::default(),
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
fn equivalent(a: &OpcPackage, b: &OpcPackage) {
    let rels = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    assert_eq!(rels(a.rels()), rels(b.rels()));
    let x = a
        .iter_parts()
        .map(|p| {
            (
                p.partname().as_str().to_owned(),
                (p.content_type().to_owned(), rels(p.rels())),
            )
        })
        .collect::<BTreeMap<_, _>>();
    let y = b
        .iter_parts()
        .map(|p| {
            (
                p.partname().as_str().to_owned(),
                (p.content_type().to_owned(), rels(p.rels())),
            )
        })
        .collect::<BTreeMap<_, _>>();
    assert_eq!(x, y);
    for p in a.iter_parts() {
        assert!(
            a.get_part(p.partname()).unwrap().blob() == b.get_part(p.partname()).unwrap().blob(),
            "changed {}",
            p.partname()
        );
    }
}
fn restored(source: &OpcPackage, restored: &OpcPackage) {
    equivalent(source, restored);
    let a = members(source);
    let b = members(restored);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if !matches!(
            name.as_str(),
            "[Content_Types].xml" | "xl/_rels/workbook.xml.rels"
        ) {
            assert!(b.get(&name) == Some(&data), "changed {name}");
        }
    }
}
fn item(label: &str, id: &str, conformance: Conformance) -> NewItem {
    NewItem {
        source: PackURI::new("/xl/workbook.xml").unwrap(),
        rel_id: format!("rIdNative{label}"),
        part: PackURI::new(format!("/customXml/native{label}.xml")).unwrap(),
        content_type: "application/xml".into(),
        xml: format!("<n:item xmlns:n=\"urn:litchi:native-store\" value=\"{label}\"/>")
            .into_bytes(),
        props: Some(NewProps {
            part: PackURI::new(format!("/customXml/native{label}Props.xml")).unwrap(),
            rel_id: "rIdProperties".into(),
            value: Props {
                id: id.into(),
                schemas: vec!["https://example.invalid/inert-schema.xsd".into()],
            },
        }),
        conformance,
    }
}
#[test]
fn native_custom_xml_rejected_duplicate_id_preserves_prior_valid_staging() {
    let bytes = native("excel-metadata.xlsx");
    let p = Package::from_bytes(&bytes).unwrap().into_plain_opc();
    let before = members(&p);
    assert!(Snapshot::load(&p).unwrap().items().is_empty());
    let mut e = Edit::new(&p).unwrap();
    assert_eq!(
        e.insert(item("A", ID_A, Conformance::Transitional))
            .unwrap(),
        0
    );
    let staged = e.items().to_vec();
    let source = e.before().clone();
    let mut duplicate = item("B", ID_A, Conformance::Transitional);
    duplicate.source = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    assert!(
        matches!(e.insert(duplicate),Err(litchi_ooxml_common::Error::Invalid(message)) if message==format!("duplicate custom XML itemID '{ID_A}'"))
    );
    assert_eq!(
        e.items().len(),
        1,
        "a rejected insertion must not remain staged"
    );
    assert!(e.items() == staged.as_slice());
    assert!(
        e.before() == &source,
        "a rejected new owner must not widen the source precondition"
    );
    let c = e.commit().unwrap();
    let out = Snapshot::load(c.package()).unwrap();
    assert_eq!(out.items().len(), 1);
    assert_eq!(out.items()[0].props().unwrap().id, ID_A);
    assert!(members(&p) == before);
    assert!(native("excel-metadata.xlsx") == bytes);
}

#[test]
fn native_custom_xml_catalog_order_is_independent_of_staging_order() {
    let bytes = native("excel-recorded-macro.xlsm");
    let p = Package::from_bytes(&bytes).unwrap().into_plain_opc();
    let before = members(&p);
    let id_b = "{65000000-2222-4222-8222-222222222222}";
    let mut e = Edit::new(&p).unwrap();
    assert_eq!(
        e.insert(item("B", id_b, Conformance::Transitional))
            .unwrap(),
        0
    );
    assert_eq!(
        e.insert(item("A", ID_A, Conformance::Transitional))
            .unwrap(),
        1
    );
    // Staged selectors retain their ordinal; published catalogs have stable source/relationship order.
    e.set_xml(0, b"<updatedB/>".to_vec()).unwrap();
    let c = e.commit().unwrap();
    let out = Snapshot::load(c.package()).unwrap();
    assert_eq!(
        out.items()
            .iter()
            .map(|i| i.props().unwrap().id.as_str())
            .collect::<Vec<_>>(),
        [ID_A, id_b]
    );
    assert_eq!(out.items()[1].xml(), b"<updatedB/>");
    restored(
        &p,
        c.patch().inverse().apply(c.package()).unwrap().package(),
    );
    assert!(members(&p) == before);
    assert!(native("excel-recorded-macro.xlsm") == bytes);
}

fn published(p: &OpcPackage, c: &litchi_ooxml_common::custom_xml::Commit) -> OpcPackage {
    let f = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    Package::from_opc(c.package().clone())
        .unwrap()
        .save(f.path())
        .unwrap();
    let q = Package::open(f.path()).unwrap().into_plain_opc();
    equivalent(&q, c.package());
    equivalent(c.patch().apply(p).unwrap().package(), &q);
    restored(p, c.patch().inverse().apply(&q).unwrap().package());
    q
}
fn native_parts_unchanged(p: &OpcPackage, q: &OpcPackage) {
    let a = members(p);
    let b = members(q);
    for (name, data) in a {
        if !matches!(
            name.as_str(),
            "[Content_Types].xml" | "xl/_rels/workbook.xml.rels"
        ) {
            assert!(b.get(&name) == Some(&data), "changed {name}");
        }
    }
    let rels = |r: &litchi_opc::Relationships| {
        let mut v = r
            .iter()
            .filter(|r| {
                !matches!(
                    r.reltype(),
                    litchi_ooxml_common::custom_xml::TRANSITIONAL_RELATIONSHIP
                        | litchi_ooxml_common::custom_xml::STRICT_RELATIONSHIP
                )
            })
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    };
    assert_eq!(rels(p.rels()), rels(q.rels()));
    for part in p.iter_parts() {
        let other = q.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), other.content_type());
        assert_eq!(rels(part.rels()), rels(other.rels()));
    }
}
#[test]
fn native_transitional_strict_and_macro_custom_xml_lifecycle_preserves_all_native_parts() {
    for (file, conformance) in [
        ("excel-metadata.xlsx", Conformance::Transitional),
        ("excel-recorded-macro.xlsm", Conformance::Transitional),
        ("excel-visual-analytics-strict.xlsx", Conformance::Strict),
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap().into_plain_opc();
        let before = members(&p);
        assert!(Snapshot::load(&p).unwrap().is_empty());
        let id_b = "{65000000-2222-4222-8222-222222222222}";
        let mut e = Edit::new(&p).unwrap();
        e.insert(item("A", ID_A, conformance)).unwrap();
        e.insert(item("B", id_b, conformance)).unwrap();
        let c = e.commit().unwrap();
        let added = published(&p, &c);
        native_parts_unchanged(&p, &added);
        let snapshot = Snapshot::load(&added).unwrap();
        assert_eq!(snapshot.items().len(), 2);
        for (value, id) in snapshot.items().iter().zip([ID_A, id_b]) {
            assert_eq!(value.props().unwrap().id, id);
            assert_eq!(value.root().namespace, "urn:litchi:native-store");
            assert_eq!(value.root().local_name, "item");
            assert_eq!(
                value.source_relationship().relationship_type,
                conformance.relationship()
            );
            assert!(
                std::str::from_utf8(value.props_xml().unwrap())
                    .unwrap()
                    .contains(conformance.namespace())
            );
        }
        let mut e = Edit::new(&added).unwrap();
        let properties=String::from_utf8(snapshot.items()[0].props_xml().unwrap().to_vec()).unwrap().replace("</ds:datastoreItem>","<!--keep--><x:future xmlns:x=\"urn:opaque\" value=\"A &amp; B\"/></ds:datastoreItem>");
        assert!(properties.contains("<!--keep-->"));
        e.set_properties_xml(0, properties.into_bytes()).unwrap();
        e.set_properties(
            0,
            Props {
                id: ID_A.into(),
                schemas: vec![
                    "urn:native:updated".into(),
                    "https://example.invalid/never-loaded.xsd".into(),
                ],
            },
        )
        .unwrap();
        let payload=b"\xef\xbb\xbf<?xml version=\"1.0\"?>\r\n<n:item xmlns:n=\"urn:litchi:native-store\">caf\xc3\xa9 &amp; &#x1F642;</n:item>";
        e.set_xml(0, payload.to_vec()).unwrap();
        let c = e.commit().unwrap();
        let updated = published(&added, &c);
        native_parts_unchanged(&p, &updated);
        let snapshot = Snapshot::load(&updated).unwrap();
        assert_eq!(snapshot.items()[0].xml(), payload);
        let props = std::str::from_utf8(snapshot.items()[0].props_xml().unwrap()).unwrap();
        assert!(props.contains("<!--keep-->") && props.contains("value=\"A &amp; B\""));
        assert_eq!(
            snapshot.items()[0].props().unwrap().schemas,
            [
                "urn:native:updated",
                "https://example.invalid/never-loaded.xsd"
            ]
        );
        let old = members(&added);
        let new = members(&updated);
        assert_eq!(
            old.keys()
                .filter(|n| old.get(*n) != new.get(*n))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            ["customXml/nativeA.xml", "customXml/nativeAProps.xml"]
        );
        assert!(matches!(
            c.patch().apply(&updated),
            Err(litchi_ooxml_common::Error::PatchConflict)
        ));
        let noop = Edit::new(&updated).unwrap().commit().unwrap();
        assert!(noop.patch().is_noop());
        assert!(members(noop.package()) == new);
        let mut e = Edit::new(&updated).unwrap();
        assert!(e.remove_properties(1).unwrap());
        assert!(
            e.items()[1].props_part().is_none()
                && e.items()[1].props().is_none()
                && e.items()[1].props_xml().is_none()
        );
        let c = e.commit().unwrap();
        let bare = published(&updated, &c);
        let snapshot = Snapshot::load(&bare).unwrap();
        assert!(snapshot.items()[1].props().is_none());
        assert!(
            bare.get_part(&PackURI::new("/customXml/nativeBProps.xml").unwrap())
                .is_err()
        );
        native_parts_unchanged(&p, &bare);
        let mut e = Edit::new(&bare).unwrap();
        assert!(e.remove(1).unwrap().is_some());
        assert!(e.remove(0).unwrap().is_some());
        assert!(e.remove(0).unwrap().is_none());
        let c = e.commit().unwrap();
        let removed = published(&bare, &c);
        assert!(Snapshot::load(&removed).unwrap().is_empty());
        restored(&p, &removed);
        assert!(members(&p) == before);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_custom_xml_payload_budget_preserves_an_exact_limit_staged_value() {
    let bytes = native("excel-metadata.xlsx");
    let native = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let maximum = members(&native).values().map(Vec::len).max().unwrap();
    let limits = SourceLimits::builder()
        .max_part_bytes(u64::try_from(maximum).unwrap())
        .unwrap()
        .build()
        .unwrap();
    let p = OpcPackage::from_bytes(&bytes, limits).unwrap();
    let before = members(&p);
    let mut e = Edit::new(&p).unwrap();
    e.insert(item("A", ID_A, Conformance::Transitional))
        .unwrap();
    let exact = format!("<r>{}</r>", "a".repeat(maximum - 7));
    assert_eq!(exact.len(), maximum);
    e.set_xml(0, exact.as_bytes().to_vec()).unwrap();
    let over = format!("<r>{}</r>", "a".repeat(maximum - 6));
    assert!(
        matches!(e.set_xml(0,over.into_bytes()),Err(litchi_ooxml_common::Error::Limit{kind:litchi_ooxml_common::LimitKind::CustomXml(litchi_ooxml_common::custom_xml::LimitKind::PayloadBytes),observed,maximum:m}) if observed==maximum+1&&m==maximum)
    );
    assert_eq!(e.items()[0].xml(), exact.as_bytes());
    let c = e.commit().unwrap();
    let q = published(&p, &c);
    assert_eq!(
        Snapshot::load(&q).unwrap().items()[0].xml(),
        exact.as_bytes()
    );
    assert!(members(&p) == before);
}

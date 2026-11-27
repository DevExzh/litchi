//! Shared test support for native DOC/XLS parents; uses public store and CFB owners.
use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
use litchi_ole_common::{custom_xml as xml, object};
use std::{collections::BTreeMap, io::Cursor, path::Path};
const STORE: &str = "native-xml";
const OPAQUE: &str = "NativeOpaque064";
const ID_A: &str = "{64000000-1111-4111-8111-111111111111}";
const ID_B: &str = "{65000000-2222-4222-8222-222222222222}";
pub(crate) fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn target() -> object::Target {
    object::Target::new(STORE, [xml::STORE_STORAGE]).unwrap()
}
pub(crate) fn open(bytes: &[u8]) -> object::Package {
    let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    let mut targets = Vec::new();
    if ole.directory_exists(&[xml::STORE_STORAGE]) {
        targets.push(target());
    }
    if ole.directory_exists(&[OPAQUE]) {
        targets.push(object::Target::new("opaque", [OPAQUE]).unwrap());
    }
    object::Package::open_with_options(
        bytes.to_vec(),
        object::OpenOptions::new().targets(object::Targets::new(targets).unwrap()),
    )
    .unwrap()
}
pub(crate) fn bytes(p: &object::Package) -> Vec<u8> {
    p.source_shared().as_ref().clone()
}
pub(crate) fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let b = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, b)
        })
        .collect()
}
type Metadata = (u8, String, u32, u64, u64);
pub(crate) fn directory(bytes: &[u8]) -> BTreeMap<Vec<String>, Metadata> {
    let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    let mut out = BTreeMap::new();
    fn metadata(e: &litchi_cfb::DirectoryEntry) -> Metadata {
        (
            e.entry_type,
            e.clsid.clone(),
            e.state_bits,
            e.creation_time,
            e.modified_time,
        )
    }
    fn walk(
        ole: &OleFile<Cursor<&[u8]>>,
        path: Vec<String>,
        out: &mut BTreeMap<Vec<String>, Metadata>,
    ) {
        for e in ole
            .list_directory_entries(&path.iter().map(String::as_str).collect::<Vec<_>>())
            .unwrap()
        {
            let mut p = path.clone();
            p.push(e.name.clone());
            out.insert(p.clone(), metadata(e));
            if e.entry_type == 1 {
                walk(ole, p, out);
            }
        }
    }
    out.insert(Vec::new(), metadata(ole.root_entry().unwrap()));
    walk(&ole, Vec::new(), &mut out);
    out
}
pub(crate) fn snapshot(bytes: &[u8]) -> Option<xml::Snapshot> {
    let mut ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    xml::Snapshot::load(&mut ole, xml::Limits::default()).unwrap()
}
fn item(id: &str, label: &str) -> xml::Item {
    let id: xml::ItemId = id.parse().unwrap();
    xml::Item::new(
        id.storage_name(),
        format!("<n:item xmlns:n=\"urn:litchi:native-store\" label=\"{label}\"/>").into_bytes(),
        xml::Properties {
            item_id: id,
            schema_references: vec!["https://example.invalid/inert-schema.xsd".into()],
        },
    )
    .unwrap()
}
fn compound(store: &xml::Store) -> Vec<u8> {
    let mut writer = OleWriter::new();
    xml::write(&mut writer, store, &xml::Limits::default()).unwrap();
    let mut b = Cursor::new(Vec::new());
    writer.write_to(&mut b).unwrap();
    let p = object::Package::open_with_options(
        b.into_inner(),
        object::OpenOptions::new().targets(object::Targets::one(target())),
    )
    .unwrap();
    p.objects().get(STORE).unwrap().compound().to_vec()
}
fn opaque() -> Vec<u8> {
    let mut w = OleWriter::new();
    w.set_root_clsid([0x64; 16]);
    w.create_storage(&["Empty"]).unwrap();
    w.set_storage_clsid(&["Empty"], [0x65; 16]).unwrap();
    w.set_storage_metadata(
        &["Empty"],
        0x1020_3040,
        132_537_600_000_000_000,
        132_537_600_100_000_000,
    )
    .unwrap();
    w.create_stream(&["Payload"], b"OPAQUE_NATIVE_PARENT_064\0\xff")
        .unwrap();
    let mut out = Cursor::new(Vec::new());
    w.write_to(&mut out).unwrap();
    out.into_inner()
}
fn outside_store(before: &[u8], after: &[u8]) {
    let a = streams(before);
    let b = streams(after);
    for (p, v) in a {
        if p.first().is_none_or(|n| n != xml::STORE_STORAGE) {
            assert!(b.get(&p) == Some(&v), "changed native stream {p:?}");
        }
    }
    let after_directory = directory(after);
    for (p, v) in directory(before) {
        if p.first().is_none_or(|n| n != xml::STORE_STORAGE) {
            assert_eq!(
                after_directory.get(&p),
                Some(&v),
                "changed directory metadata {p:?}"
            );
        }
    }
}
fn publish(p: &object::Package, c: &object::Commit) -> object::Package {
    let file = tempfile::NamedTempFile::new().unwrap();
    c.package()
        .write_to(std::fs::File::create(file.path()).unwrap())
        .unwrap();
    let q = open(&std::fs::read(file.path()).unwrap());
    assert!(bytes(&q) == bytes(c.package()));
    // Object patches replace artifact bytes; target keys are host-supplied
    // open policy. Use an unselected view when replay changes that catalog,
    // then reopen the result with the targets valid for its new directory.
    let before_artifact = object::Package::open(bytes(p)).unwrap();
    let after_artifact = object::Package::open(bytes(&q)).unwrap();
    assert!(bytes(before_artifact.apply(c.patch()).unwrap().package()) == bytes(&q));
    assert!(
        bytes(
            after_artifact
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
        ) == bytes(p)
    );
    outside_store(&bytes(p), &bytes(&q));
    q
}
/// Qualify the shared store model and its publication through the public CFB owner.
/// Return the created native-parent derivative for a subsequent format-specific edit.
pub(crate) fn exercise(native: &[u8]) -> Vec<u8> {
    assert!(
        snapshot(native).is_none(),
        "native store evidence changed; review the fixture"
    );
    let base = open(native);
    let original_streams = streams(native);
    let original_directory = directory(native);
    let empty = xml::Snapshot::from_store(
        xml::Store::new(xml::Promotion::Unspecified, vec![]).unwrap(),
        xml::Limits::default(),
    )
    .unwrap();
    let mut e = empty.edit();
    assert_eq!(e.insert(item(ID_A, "A")).unwrap(), 0);
    let c = e.commit().unwrap();
    let model = c.package();
    assert_eq!(
        model.item(0).unwrap().root_name().namespace.as_deref(),
        Some("urn:litchi:native-store")
    );
    assert_eq!(model.item(0).unwrap().root_name().local_name, "item");
    assert_eq!(model.item(0).unwrap().kind(), xml::ItemKind::Other);
    assert_eq!(c.patch().apply(&empty).unwrap(), *model);
    assert_eq!(c.patch().inverse().apply(model).unwrap(), empty);
    let mut e = base.edit().unwrap();
    e.add_storage(target(), compound(model.store())).unwrap();
    e.add_storage(object::Target::new("opaque", [OPAQUE]).unwrap(), opaque())
        .unwrap();
    let c = e.commit().unwrap();
    let added = publish(&base, &c);
    let seeded = bytes(&added);
    assert_eq!(streams(&seeded).len(), original_streams.len() + 3);
    assert_eq!(directory(&seeded).len(), original_directory.len() + 7);
    assert_eq!(
        streams(&seeded)[&vec![OPAQUE.into(), "Payload".into()]],
        b"OPAQUE_NATIVE_PARENT_064\0\xff"
    );
    let opaque_metadata = directory(&opaque());
    let seeded_metadata = directory(&seeded);
    assert_eq!(
        seeded_metadata[&vec![OPAQUE.into()]].1,
        opaque_metadata[&Vec::new()].1
    );
    assert_eq!(
        seeded_metadata[&vec![OPAQUE.into(), "Empty".into()]],
        opaque_metadata[&vec!["Empty".into()]]
    );

    for (p, b) in original_streams {
        assert!(streams(&seeded).get(&p) == Some(&b));
    }
    for (p, d) in original_directory {
        assert_eq!(directory(&seeded).get(&p), Some(&d));
    }
    let first = snapshot(&seeded).unwrap();
    let noop = first.edit().commit().unwrap();
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package(), &first);
    assert_eq!(first.items().len(), 1);
    assert_eq!(
        first.item(0).unwrap().properties().item_id.to_string(),
        ID_A
    );
    let mut e = first.edit();
    let replacement=b"<?xml version=\"1.0\"?>\r\n<n:item xmlns:n=\"urn:litchi:native-store\">caf\xc3\xa9 &amp; intact</n:item>";
    e.set_xml(0, replacement.to_vec()).unwrap();
    let staged = e.items().to_vec();
    assert!(matches!(
        e.insert(item(ID_A, "duplicate")),
        Err(xml::Error::Invalid(_))
    ));
    assert!(e.items() == staged.as_slice());
    for invalid in [
        b"<r>&unknown;</r>".as_slice(),
        b"<r>&#0;</r>",
        b"<!DOCTYPE r><r/>",
    ] {
        assert!(e.set_xml(0, invalid.to_vec()).is_err());
        assert_eq!(e.items()[0].xml(), replacement);
    }
    let properties = String::from_utf8(e.items()[0].properties_xml().to_vec()).unwrap();
    let properties = properties.replace(
        "</ds:datastoreItem>",
        "<x:payload xmlns:x=\"urn:opaque\">A &amp; B &#x1F642;</x:payload></ds:datastoreItem>",
    );
    assert!(properties.contains("A &amp; B"));
    e.set_properties_xml(0, properties.into_bytes()).unwrap();
    e.set_schema_references(
        0,
        ["urn:native:one", "https://example.invalid/never-loaded.xsd"],
    )
    .unwrap();
    e.insert(item(ID_B, "B")).unwrap();
    let changed = e.commit().unwrap();
    assert_eq!(changed.package().items().len(), 2);
    assert_eq!(changed.package().item(0).unwrap().xml(), replacement);
    let mut publish_edit = added.edit().unwrap();
    publish_edit
        .replace(STORE, compound(changed.package().store()))
        .unwrap();
    let c = publish_edit.commit().unwrap();
    let updated = publish(&added, &c);
    let second = snapshot(&bytes(&updated)).unwrap();
    assert_eq!(second, *changed.package());
    assert!(
        std::str::from_utf8(second.item(0).unwrap().properties_xml())
            .unwrap()
            .contains("A &amp; B &#x1F642;")
    );
    assert_eq!(changed.patch().inverse().apply(&second).unwrap(), first);
    assert!(
        matches!(changed.patch().apply(&second),Err(xml::Error::Invalid(m)) if m=="custom XML patch source does not match its base snapshot")
    );
    let limited = xml::Snapshot::from_store(
        first.store().clone(),
        xml::Limits {
            max_items: 1,
            ..Default::default()
        },
    )
    .unwrap();
    let mut e = limited.edit();
    assert!(
        matches!(e.insert(item(ID_B,"B")),Err(xml::Error::ResourceLimit(m)) if m=="item count exceeds its limit")
    );
    assert_eq!(e.items(), limited.items());
    let mut e = second.edit();
    let removed = e.remove(1).unwrap();
    assert_eq!(removed.properties().item_id.to_string(), ID_B);
    let changed = e.commit().unwrap();
    let mut pe = updated.edit().unwrap();
    pe.replace(STORE, compound(changed.package().store()))
        .unwrap();
    let c = pe.commit().unwrap();
    let one = publish(&updated, &c);
    assert_eq!(snapshot(&bytes(&one)).unwrap().items().len(), 1);
    let mut e = one.edit().unwrap();
    e.remove_storage(STORE).unwrap();
    let c = e.commit().unwrap();
    let gone = publish(&one, &c);
    assert!(snapshot(&bytes(&gone)).is_none());
    assert!(directory(&bytes(&gone)).contains_key(&vec![OPAQUE.into(), "Empty".into()]));
    assert!(bytes(&base) == native);
    seeded
}
/// The native format edit may change its own streams, but never the added XML/opaque owners.
pub(crate) fn preserved_owned(before: &[u8], after: &[u8]) {
    let a = streams(before);
    let b = streams(after);
    for (p, v) in a {
        if p.first()
            .is_some_and(|n| n == xml::STORE_STORAGE || n == OPAQUE)
        {
            assert!(b.get(&p) == Some(&v), "changed retained stream {p:?}");
        }
    }
    let d = directory(after);
    for (p, v) in directory(before) {
        if p.first()
            .is_some_and(|n| n == xml::STORE_STORAGE || n == OPAQUE)
        {
            assert_eq!(d.get(&p), Some(&v), "changed retained directory {p:?}");
        }
    }
    assert_eq!(snapshot(after), snapshot(before));
}

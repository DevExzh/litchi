//! Native legacy record metadata: typed reads and bounded source-preserving edits.
use litchi_core::{Position, edit::RefusalKind};
use litchi_ppt::client_data::{self, Child, ChildKind, Limits};
use litchi_ppt::text_edit::Target;
use litchi_ppt::{
    CustomTableStylesPlacement as Placement, Package, Record, RecordParser, RecordType,
};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, io::Cursor};
const TEXT: &str = "powerpoint-text-table.ppt";
const MEDIA: &str = "powerpoint-modern-media.ppt";
const STYLE: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let sequence = NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "litchi-native-record-metadata-{}-{nonce}-{sequence}.ppt",
            std::process::id()
        ));
        std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)
            .unwrap();
        Self(path)
    }
    fn path(&self) -> &std::path::Path {
        &self.0
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut cfb =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    cfb.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let bytes = cfb.open_stream(&names).unwrap();
            (path, bytes)
        })
        .collect()
}
fn document(bytes: &[u8]) -> Record {
    let parts = streams(bytes);
    let stream = &parts[&vec!["PowerPoint Document".into()]];
    let mut parser = RecordParser::new();
    parser
        .parse_document(stream, &litchi_ppt::SourceLimits::default())
        .unwrap();
    let mapping = litchi_ppt::PersistMapping::build_from_records_ref(&parser.find_records_ref());
    let p = Package::from_bytes(bytes).unwrap();
    let presentation = p.presentation().unwrap();
    let id = presentation.slide_directory().document_persist_id();
    let (r, _) = Record::parse(
        stream,
        mapping.get_offset(id).unwrap() as usize,
        &litchi_ppt::SourceLimits::default(),
    )
    .unwrap();
    assert_eq!(r.record_type, RecordType::Document);
    r
}
fn encoded(record: &Record) -> Vec<u8> {
    let mut out = ((record.instance << 4) | record.version)
        .to_le_bytes()
        .to_vec();
    out.extend_from_slice(&record.record_type_raw.to_le_bytes());
    out.extend_from_slice(&record.data_length.to_le_bytes());
    out.extend_from_slice(&record.data);
    out
}

fn package(name: &str) -> Package {
    Package::from_bytes(&native(name)).unwrap()
}
fn save(p: &Package) -> Package {
    let f = TempPpt::new();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn styles(p: &Package) -> litchi_ppt::document_properties::CustomTableStyles {
    p.presentation()
        .unwrap()
        .powerpoint12_document_properties()
        .unwrap()
        .custom_table_styles
        .unwrap()
}
fn style_atom(p: &Package) -> Vec<u8> {
    let d = document(&p.to_bytes().unwrap());
    encoded(
        d.children
            .iter()
            .find(|r| r.record_type == RecordType::RoundTripCustomTableStyles12Atom)
            .unwrap(),
    )
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    p.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = p.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn preserved_streams(a: &[u8], b: &[u8]) {
    let a = streams(a);
    let b = streams(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in &a {
        if name != &["PowerPoint Document"] && name != &["Current User"] {
            assert!(data == &b[name], "changed {name:?}");
        }
    }
    assert!(
        b[&vec!["PowerPoint Document".into()]].starts_with(&a[&vec!["PowerPoint Document".into()]])
    );
}
fn preserved_document_children(a: &[u8], b: &[u8]) {
    let a = document(a);
    let b = document(b);
    let left = a
        .children
        .iter()
        .filter(|r| r.record_type != RecordType::RoundTripCustomTableStyles12Atom)
        .collect::<Vec<_>>();
    let right = b
        .children
        .iter()
        .filter(|r| r.record_type != RecordType::RoundTripCustomTableStyles12Atom)
        .collect::<Vec<_>>();
    assert!(left == right, "unowned Document child changed");
}
fn preserved_slides(a: &Package, b: &Package) {
    let p = a.presentation().unwrap();
    let q = b.presentation().unwrap();
    let a = p.slides().unwrap();
    let b = q.slides().unwrap();
    assert_eq!(a.len(), b.len());
    for (a, b) in a.iter().zip(&b) {
        assert_eq!(a.slide_id(), b.slide_id());
        assert_eq!(a.text().unwrap(), b.text().unwrap());
        assert!(a.record() == b.record());
    }
}

fn persisted_records(bytes: &[u8]) -> BTreeMap<u32, Vec<u8>> {
    let parts = streams(bytes);
    let stream = &parts[&vec!["PowerPoint Document".into()]];
    let limits = litchi_ppt::SourceLimits::default();
    let mut parser = RecordParser::new();
    parser.parse_document(stream, &limits).unwrap();
    let mapping = litchi_ppt::PersistMapping::build_from_records_ref(&parser.find_records_ref());
    mapping
        .get_persist_ids()
        .into_iter()
        .map(|id| {
            let (record, _) =
                Record::parse(stream, mapping.get_offset(id).unwrap() as usize, &limits).unwrap();
            (id, encoded(&record))
        })
        .collect()
}

fn preserved_persisted_owners(a: &[u8], b: &[u8], owned: u32) {
    let a = persisted_records(a);
    let b = persisted_records(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (id, bytes) in &a {
        if *id != owned {
            assert!(bytes == &b[id], "unowned persist {id} changed");
        }
    }
}

fn replay(p: &Package, c: &litchi_ppt::package::edit::Commit) -> Package {
    let original = p.to_bytes().unwrap();
    let out = save(c.package());
    let bytes = out.to_bytes().unwrap();
    preserved_streams(&original, &bytes);
    preserved_persisted_owners(
        &original,
        &bytes,
        p.presentation()
            .unwrap()
            .slide_directory()
            .document_persist_id(),
    );
    preserved_document_children(&original, &bytes);
    preserved_slides(p, &out);
    assert!(p.apply(c.patch()).unwrap().package().to_bytes().unwrap() == bytes);
    assert!(matches!(
        out.apply(c.patch()),
        Err(litchi_ppt::package::edit::Error::PatchConflict)
    ));
    assert!(
        out.apply(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == original
    );
    out
}
fn refusal(e: &litchi_ppt::package::edit::Error) -> RefusalKind {
    let litchi_ppt::package::edit::Error::Native(litchi_ppt::slide_order::Error::Refused(r)) = e
    else {
        panic!("typed refusal: {e}")
    };
    r.kind()
}
#[test]
fn native_custom_table_style_records_distinguish_populated_and_empty_packages() {
    for (file, size, count) in [
        (TEXT, 2393, 1),
        (MEDIA, 1986, 0),
        ("powerpoint-fonts-math-handout.ppt", 1986, 0),
        ("powerpoint-navigation-review.ppt", 1986, 0),
    ] {
        let p = package(file);
        let v = styles(&p);
        assert_eq!(v.record_version, 0);
        assert_eq!(v.package.data.len(), size);
        assert_eq!(v.package.part_count, 1);
        assert_eq!(v.package.xml_part_name, "/tableStyles.xml");
        let m = members(&v.package.data);
        assert_eq!(
            m.keys().map(String::as_str).collect::<Vec<_>>(),
            ["[Content_Types].xml", "_rels/.rels", "tableStyles.xml"]
        );
        let xml = std::str::from_utf8(&m["tableStyles.xml"]).unwrap();
        assert!(xml.contains(&format!(r#"def="{STYLE}""#)));
        let mut r = quick_xml::Reader::from_reader(m["tableStyles.xml"].as_slice());
        let mut found = 0;
        loop {
            match r.read_event().unwrap() {
                quick_xml::events::Event::Start(e) if e.name().as_ref() == b"a:tblStyle" => {
                    found += 1;
                    let attrs = e
                        .attributes()
                        .map(|a| {
                            let a = a.unwrap();
                            (
                                a.key.as_ref().to_vec(),
                                a.decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Explicit1_0,
                                    r.decoder(),
                                )
                                .unwrap()
                                .into_owned(),
                            )
                        })
                        .collect::<BTreeMap<_, _>>();
                    assert_eq!(attrs[b"styleId".as_slice()], STYLE);
                    assert_eq!(attrs[b"styleName".as_slice()], "Medium Style 2 - Accent 1");
                },
                quick_xml::events::Event::Eof => break,
                _ => {},
            }
        }
        assert_eq!(found, count);
        let d = document(&p.to_bytes().unwrap());
        let at = d
            .children
            .iter()
            .position(|r| r.record_type == RecordType::RoundTripCustomTableStyles12Atom)
            .unwrap();
        assert_eq!(d.children[at + 1].record_type, RecordType::EndDocument);
        assert_eq!(d.children[at].data, v.package.data);
        assert_eq!(
            p.presentation()
                .unwrap()
                .document_structure()
                .unwrap()
                .custom_table_styles,
            Some(Placement::BeforeEndDocument)
        );
        let mut e = p.edit().unwrap();
        e.replace_custom_table_styles(&encoded(&d.children[at]))
            .unwrap();
        assert!(e.commit().unwrap().package().to_bytes().unwrap() == native(file));
    }
}
#[test]
fn native_table_style_atom_move_remove_insert_keeps_live_catalog_and_every_other_owner() {
    let p = package(MEDIA);
    let before = p.to_bytes().unwrap();
    let atom = style_atom(&p);
    let mut e = p.edit().unwrap();
    e.move_custom_table_styles(Placement::AfterEndDocument)
        .unwrap();
    let moved = e.commit().unwrap();
    let out = replay(&p, &moved);
    assert_eq!(
        out.presentation()
            .unwrap()
            .document_structure()
            .unwrap()
            .custom_table_styles,
        Some(Placement::AfterEndDocument)
    );
    assert_eq!(style_atom(&out), atom);
    let mut e = out.edit().unwrap();
    e.move_custom_table_styles(Placement::AfterEndDocument)
        .unwrap();
    e.replace_custom_table_styles(&atom).unwrap();
    assert!(!e.commit().unwrap().diagnostics().changed());
    let mut e = out.edit().unwrap();
    assert_eq!(e.remove_custom_table_styles().unwrap(), atom);
    assert_eq!(
        refusal(&e.remove_custom_table_styles().unwrap_err()),
        RefusalKind::NotFound
    );
    let removed = e.commit().unwrap();
    let empty = replay(&out, &removed);
    assert_eq!(
        empty
            .presentation()
            .unwrap()
            .document_structure()
            .unwrap()
            .custom_table_styles,
        None
    );
    assert!(
        empty
            .presentation()
            .unwrap()
            .powerpoint12_document_properties()
            .unwrap()
            .custom_table_styles
            .is_none()
    );
    let mut e = empty.edit().unwrap();
    e.insert_custom_table_styles(&atom, Placement::BeforeEndDocument)
        .unwrap();
    assert_eq!(
        refusal(
            &e.insert_custom_table_styles(&atom, Placement::BeforeEndDocument)
                .unwrap_err()
        ),
        RefusalKind::UnsupportedContent
    );
    let inserted = e.commit().unwrap();
    let restored = replay(&empty, &inserted);
    assert_eq!(styles(&restored), styles(&p));
    assert_eq!(style_atom(&restored), atom);
    assert!(p.to_bytes().unwrap() == before);
}
#[test]
fn native_table_style_payload_replacement_uses_a_valid_native_donor_and_rejects_bad_envelopes() {
    let p = package(MEDIA);
    let donor = package(TEXT);
    let atom = style_atom(&donor);
    let original = style_atom(&p);
    let mut e = p.edit().unwrap();
    assert_eq!(e.replace_custom_table_styles(&atom).unwrap(), original);
    let mut invalid = atom.clone();
    invalid[2] = 0;
    assert_eq!(
        refusal(&e.replace_custom_table_styles(&invalid).unwrap_err()),
        RefusalKind::InvalidValue
    );
    let c = e.commit().unwrap();
    let out = replay(&p, &c);
    assert_eq!(styles(&out), styles(&donor));
    assert_eq!(style_atom(&out), atom);
    assert_eq!(
        Sha256::digest(styles(&out).package.data),
        Sha256::digest(styles(&donor).package.data)
    );
    assert_eq!(native(TEXT), donor.to_bytes().unwrap());
}
fn descendants<'a>(r: &'a Record, kind: RecordType, out: &mut Vec<&'a Record>) {
    if r.record_type == kind {
        out.push(r);
    }
    for c in &r.children {
        descendants(c, kind, out);
    }
}
#[test]
fn native_text_language_and_spelling_defaults_round_trip_without_claiming_outline_bodies() {
    for file in [
        TEXT,
        MEDIA,
        "powerpoint-fonts-math-handout.ppt",
        "powerpoint-navigation-review.ppt",
    ] {
        let p = package(file);
        let presentation = p.presentation().unwrap();
        let value = presentation.text_special_info_defaults().unwrap().unwrap();
        assert_eq!(value.language(), Some(0x0409));
        assert_eq!(value.alternate_language(), Some(0));
        let spelling = value.spelling().unwrap();
        assert!(!spelling.error() && !spelling.clean());
        let d = document(&p.to_bytes().unwrap());
        let mut records = vec![];
        descendants(&d, RecordType::TextSpecialInfoDefaultAtom, &mut records);
        assert_eq!(records.len(), 1);
        assert_eq!(value.to_bytes().unwrap(), encoded(records[0]));
        assert_eq!(records[0].data, [7, 0, 0, 0, 0, 0, 9, 4, 0, 0]);
        assert!(presentation.outline_text_refs().unwrap().is_empty());
        for slide in presentation.slides().unwrap() {
            assert!(slide.outline_text_bodies().is_empty());
            assert!(!slide.text().unwrap().is_empty());
        }
        assert!(p.to_bytes().unwrap() == native(file));
    }
}
fn target(slide: usize, shape: usize) -> Target {
    Target::new(Position::new(slide), Position::new(shape))
}
#[test]
fn native_client_data_reads_exact_placeholder_roundtrip_and_external_object_children() {
    let p = package(MEDIA);
    let e = p.edit_client_data(target(0, 0)).unwrap();
    let children = e.children().unwrap();
    assert_eq!(
        children.iter().map(Child::kind).collect::<Vec<_>>(),
        [
            ChildKind::Placeholder,
            ChildKind::ProgrammableTags,
            ChildKind::RoundTripShapeId12
        ]
    );
    assert_eq!(children[0].payload(), [255, 255, 255, 255, 15, 0, 0, 0]);
    assert_eq!(children[2].round_trip_shape_id(), Some(2));
    assert_eq!(children[1].payload().len(), 58);
    let encoded = e.client_data().unwrap().to_bytes().unwrap();
    let snapshot = client_data::Snapshot::parse(&encoded, &Limits::default()).unwrap();
    assert_eq!(snapshot.client_data(), e.client_data().unwrap());
    for (slide, id) in [(2, 4), (3, 5)] {
        let e = p.edit_client_data(target(slide, 1)).unwrap();
        assert_eq!(
            e.client_data()
                .unwrap()
                .external_object_reference()
                .unwrap()
                .external_object_id(),
            Some(id)
        );
        assert_eq!(
            e.client_data()
                .unwrap()
                .child(ChildKind::RoundTripShapeId12)
                .unwrap()
                .round_trip_shape_id(),
            Some(3)
        );
    }
    let n = package("powerpoint-navigation-review.ppt");
    for (shape, id) in [(2, 7), (3, 9), (4, 8)] {
        let e = n.edit_client_data(target(0, shape)).unwrap();
        assert_eq!(e.children().unwrap().len(), 1);
        assert_eq!(
            e.children().unwrap()[0].round_trip_placeholder_id(),
            Some(id)
        );
    }
    assert!(e.commit().unwrap().package().to_bytes().unwrap() == native(MEDIA));
}
fn drawing_leaves(bytes: &[u8], out: &mut Vec<(u16, Vec<u8>)>) {
    let mut at = 0;
    while at < bytes.len() {
        let header = u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap());
        let kind = u16::from_le_bytes(bytes[at + 2..at + 4].try_into().unwrap());
        let len = u32::from_le_bytes(bytes[at + 4..at + 8].try_into().unwrap()) as usize;
        let end = at + 8 + len;
        if header & 15 == 15 && kind != 0xf011 {
            drawing_leaves(&bytes[at + 8..end], out);
        } else {
            out.push((kind, bytes[at..end].to_vec()));
        }
        at = end;
    }
    assert_eq!(at, bytes.len());
}
fn preserve_client_edit(a: &Package, b: &Package) {
    let left = a.to_bytes().unwrap();
    let right = b.to_bytes().unwrap();
    preserved_streams(&left, &right);
    assert!(document(&left) == document(&right));
    let p = a.presentation().unwrap();
    let q = b.presentation().unwrap();
    preserved_persisted_owners(&left, &right, p.slide_directory().entries()[0].persist_id());
    let a = p.slides().unwrap();
    let b = q.slides().unwrap();
    assert_eq!(a.len(), b.len());
    for (i, (x, y)) in a.iter().zip(&b).enumerate() {
        assert_eq!(x.text().unwrap(), y.text().unwrap());
        if i != 0 {
            assert!(x.record() == y.record());
            continue;
        }
        let left = x
            .record()
            .children
            .iter()
            .filter(|r| r.record_type != RecordType::PPDrawing)
            .collect::<Vec<_>>();
        let right = y
            .record()
            .children
            .iter()
            .filter(|r| r.record_type != RecordType::PPDrawing)
            .collect::<Vec<_>>();
        assert!(left == right);
        let mut l = vec![];
        let mut r = vec![];
        drawing_leaves(
            &x.record().find_child(RecordType::PPDrawing).unwrap().data,
            &mut l,
        );
        drawing_leaves(
            &y.record().find_child(RecordType::PPDrawing).unwrap().data,
            &mut r,
        );
        assert_eq!(l.len(), r.len());
        let selected = l.iter().position(|(k, _)| *k == 0xf011).unwrap();
        assert_eq!(r[selected].0, 0xf011);
        for (n, (l, r)) in l.iter().zip(&r).enumerate() {
            if n != selected {
                assert!(l == r, "changed unowned drawing leaf {n}");
            }
        }
    }
}
#[test]
fn native_client_data_opaque_child_crud_is_atomic_and_preserves_every_other_record() {
    let p = package(MEDIA);
    let bytes = p.to_bytes().unwrap();
    let t = target(0, 0);
    let original = p.edit_client_data(t).unwrap().children().unwrap().to_vec();
    let mut e = p.edit_client_data(t).unwrap();
    let first = Child::opaque(0, 0, 0x7770, b"native-derived-note".to_vec()).unwrap();
    e.append(first.clone()).unwrap();
    assert!(matches!(
        e.append(original[2].clone()),
        Err(client_data::Error::Refused(ref r)) if r.kind()==RefusalKind::InvalidValue
    ));
    assert_eq!(e.children().unwrap().len(), 4);
    let c = e.commit().unwrap();
    let out = save(c.package());
    preserve_client_edit(&p, &out);
    assert_eq!(
        out.edit_client_data(t).unwrap().children().unwrap(),
        [
            original[0].clone(),
            original[1].clone(),
            original[2].clone(),
            first.clone()
        ]
    );
    assert!(matches!(
        out.apply_client_data(c.patch()),
        Err(client_data::Error::PatchConflict)
    ));
    assert!(
        p.apply_client_data(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == out.to_bytes().unwrap()
    );
    assert!(
        out.apply_client_data(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    let mut e = out.edit_client_data(t).unwrap();
    let changed = Child::opaque(0, 0, 0x7770, b"updated-native-derived-note".to_vec()).unwrap();
    assert_eq!(e.replace(3, changed.clone()).unwrap(), first);
    assert!(
        matches!(e.remove(99),Err(client_data::Error::Refused(ref r)) if r.kind()==RefusalKind::NotFound)
    );
    let updated = e.commit().unwrap();
    let after = save(updated.package());
    preserve_client_edit(&out, &after);
    assert_eq!(
        after.edit_client_data(t).unwrap().children().unwrap()[3],
        changed
    );
    let mut e = after.edit_client_data(t).unwrap();
    e.replace(3, changed.clone()).unwrap();
    assert!(!e.commit().unwrap().diagnostics().changed());
    let mut e = after.edit_client_data(t).unwrap();
    assert_eq!(e.remove(3).unwrap(), changed);
    let removed = e.commit().unwrap();
    let restored = save(removed.package());
    preserve_client_edit(&after, &restored);
    assert_eq!(
        restored.edit_client_data(t).unwrap().children().unwrap(),
        original
    );
    assert!(
        restored
            .apply_client_data(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == after.to_bytes().unwrap()
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

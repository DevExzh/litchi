//! Native notes/handout master metadata without rendering or synthesis.

use litchi_core::{Position, edit::RefusalKind};
use litchi_ppt::{
    Package, Record, RecordType,
    master::{Kind, metadata::notes_styles::Styles},
};
use std::{collections::BTreeMap, io::Cursor};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}

struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        Self(std::env::temp_dir().join(format!(
            "litchi-native-master-{}-{nonce}.ppt",
            std::process::id()
        )))
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

fn master(package: &Package, kind: Kind) -> Record {
    let p = package.presentation().unwrap();
    let masters = p.masters().unwrap();
    match kind {
        Kind::Notes => masters.notes().unwrap().record().clone(),
        Kind::Handout => masters.handout().unwrap().record().clone(),
        _ => panic!("test selects notes or handout masters"),
    }
}

fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let bytes = ole.open_stream(&names).unwrap();
            (path, bytes)
        })
        .collect()
}

fn preserved(before: &[u8], after: &[u8], kind: Kind, selected: fn(&Record) -> bool) {
    let a = streams(before);
    let b = streams(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (path, bytes) in &a {
        if path != &["PowerPoint Document"] && path != &["Current User"] {
            assert!(bytes == &b[path], "changed {path:?}");
        }
    }
    assert!(
        b[&vec!["PowerPoint Document".into()]].starts_with(&a[&vec!["PowerPoint Document".into()]])
    );
    let a = Package::from_bytes(before).unwrap();
    let b = Package::from_bytes(after).unwrap();
    let before_master = master(&a, kind);
    let after_master = master(&b, kind);
    let left = before_master
        .children
        .iter()
        .filter(|r| !selected(r))
        .collect::<Vec<_>>();
    let right = after_master
        .children
        .iter()
        .filter(|r| !selected(r))
        .collect::<Vec<_>>();
    assert!(left == right, "changed unrelated master child");
    let p = a.presentation().unwrap();
    let q = b.presentation().unwrap();
    let left = p.slides().unwrap();
    let right = q.slides().unwrap();
    assert_eq!(left.len(), right.len());
    for (a, b) in left.iter().zip(&right) {
        assert_eq!(a.slide_id(), b.slide_id());
        assert!(a.record().data == b.record().data);
    }
    let left = p.masters().unwrap();
    let right = q.masters().unwrap();
    assert_eq!(left.main().len(), right.main().len());
    assert_eq!(left.title().len(), right.title().len());
    for (a, b) in left.main().iter().zip(right.main()) {
        assert_eq!(a.identity(), b.identity());
        assert!(a.record().data == b.record().data);
    }
}

#[test]
fn native_notes_and_handout_masters_resolve_exact_persist_owners() {
    for (name, kind, record_type, bytes) in [
        ("powerpoint-navigation-review.ppt", Kind::Notes, 1008, 19869),
        (
            "powerpoint-fonts-math-handout.ppt",
            Kind::Handout,
            4041,
            10872,
        ),
    ] {
        let package = Package::from_bytes(&native(name)).unwrap();
        let p = package.presentation().unwrap();
        let masters = p.masters().unwrap();
        assert_eq!(masters.main().len(), 1);
        assert!(masters.title().is_empty());
        let selected = master(&package, kind);
        assert_eq!(selected.record_type_raw, record_type);
        assert_eq!(selected.data.len(), bytes);
        if kind == Kind::Notes {
            assert_eq!(masters.notes().unwrap().persist().id(), 2);
            assert!(masters.handout().is_none());
        } else {
            assert_eq!(masters.handout().unwrap().persist().id(), 2);
            assert!(masters.notes().is_none());
        }
        assert!(
            package
                .edit_master_metadata(kind, Position::new(0))
                .unwrap()
                .name()
                .unwrap()
                .is_none()
        );
    }
    let package = Package::from_bytes(&native("powerpoint-navigation-review.ppt")).unwrap();
    let edit = package
        .edit_master_metadata(Kind::Notes, Position::new(0))
        .unwrap();
    let styles = edit.notes_styles().unwrap().unwrap();
    assert_eq!(styles.package().part_count(), 1);
    assert_eq!(
        styles.package().xml_part_name(),
        "/drs/slideMasters/slideMaster1.xml"
    );
    assert_eq!(styles.bytes().len(), 1285);
    use sha2::{Digest, Sha256};
    assert_eq!(
        Sha256::digest(styles.bytes())
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>(),
        "a8877df932aa27d106ecee8523b841063aa29063cd09536733987fa1a2359889"
    );
}

#[test]
fn native_notes_and_handout_master_name_edits_are_isolated_and_reversible() {
    for (file, kind, name) in [
        (
            "powerpoint-navigation-review.ppt",
            Kind::Notes,
            "Native notes metadata",
        ),
        (
            "powerpoint-fonts-math-handout.ppt",
            Kind::Handout,
            "Native handout metadata",
        ),
    ] {
        let before = native(file);
        let package = Package::from_bytes(&before).unwrap();
        let mut edit = package
            .edit_master_metadata(kind, Position::new(0))
            .unwrap();
        assert!(!edit.clear_name().unwrap());
        edit.set_name(name).unwrap();
        // Notes-style metadata is only valid on a notes master. A refused
        // handout request must preserve the earlier valid name change.
        if kind == Kind::Handout {
            assert!(matches!(edit.set_notes_styles(Styles::empty().unwrap()),
                Err(litchi_ppt::master::metadata::Error::Refused(refusal))
                if refusal.kind() == RefusalKind::InvalidValue));
        }
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let after = updated.to_bytes().unwrap();
        preserved(&before, &after, kind, |r| {
            r.record_type == RecordType::CString && r.instance == 3
        });
        let output = TempPpt::new();
        updated.save(&output.0).unwrap();
        let reopened = Package::open(&output.0).unwrap();
        assert_eq!(
            reopened
                .edit_master_metadata(kind, Position::new(0))
                .unwrap()
                .name()
                .unwrap()
                .unwrap()
                .as_str(),
            name
        );
        let mut noop = reopened
            .edit_master_metadata(kind, Position::new(0))
            .unwrap();
        noop.set_name(name).unwrap();
        assert!(noop.commit().unwrap().patch().is_noop());
        let mut remove = reopened
            .edit_master_metadata(kind, Position::new(0))
            .unwrap();
        assert!(remove.clear_name().unwrap());
        assert!(!remove.clear_name().unwrap());
        let removed = remove.commit().unwrap().into_package();
        assert!(master(&removed, kind).data == master(&package, kind).data);
        assert!(
            reopened
                .apply_master_metadata(&patch.inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
                == before
        );
        assert!(package.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_notes_style_package_replacement_preserves_other_master_and_slide_records() {
    let before = native("powerpoint-navigation-review.ppt");
    let package = Package::from_bytes(&before).unwrap();
    let original = package
        .edit_master_metadata(Kind::Notes, Position::new(0))
        .unwrap()
        .notes_styles()
        .unwrap()
        .unwrap();
    let mut noop = package
        .edit_master_metadata(Kind::Notes, Position::new(0))
        .unwrap();
    noop.set_notes_styles(original).unwrap();
    assert!(noop.commit().unwrap().patch().is_noop());
    let xml = br#"<p:txStyles xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle></p:txStyles>"#;
    let expected = Styles::from_xml(xml).unwrap();
    let mut edit = package
        .edit_master_metadata(Kind::Notes, Position::new(0))
        .unwrap();
    edit.set_notes_styles_xml(xml).unwrap();
    let rejected = edit.set_notes_styles_xml(b"<notStyles/>");
    assert!(
        matches!(
            rejected,
            Err(litchi_ppt::master::metadata::Error::Package(
                litchi_ppt::package::Error::Corrupted(_)
            ))
        ),
        "{rejected:?}"
    );

    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    preserved(&before, &after, Kind::Notes, |r| {
        r.record_type == RecordType::RoundTripNotesMasterTextStyles12Atom
    });
    let output = TempPpt::new();
    updated.save(&output.0).unwrap();
    let reopened = Package::open(&output.0).unwrap();
    let actual = reopened
        .edit_master_metadata(Kind::Notes, Position::new(0))
        .unwrap()
        .notes_styles()
        .unwrap()
        .unwrap();
    assert_eq!(actual, expected);
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        actual.bytes().to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive
            .read_member("drs/slideMasters/slideMaster1.xml")
            .unwrap(),
        xml
    );
    assert!(
        reopened
            .apply_master_metadata(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == before
    );
}

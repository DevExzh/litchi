use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::*;
use crate::time::Offset;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
fn slide_name() -> PackURI {
    PackURI::new("/ppt/slides/slide1.xml").unwrap()
}

fn fixture() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        slide_name(),
        ct::PML_SLIDE.into(),
        format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:x="urn:future"><p:cSld/><p:extLst><p:ext uri="{{opaque}}"><x:future><x:data>keep</x:data></x:future></p:ext></p:extLst></p:sld>"#
        )
        .into_bytes(),
    )));
    package
}

fn drafts() -> Vec<Draft> {
    vec![
        Draft::trigger(Trigger::OnClick, Offset::secs(1), 6),
        Draft::seek(Offset::secs(2), 4, Offset::secs(3)),
    ]
}

#[test]
fn create_read_noop_and_typed_update_are_source_checked() {
    let name = slide_name();
    let mut package = fixture();
    store(&mut package, &name, &drafts()).unwrap();
    let source = load_snapshot(&package, &name).unwrap().unwrap();
    assert_eq!(source.events().len(), 2);

    let no_op = source.edit().commit().unwrap();
    assert!(!no_op.is_changed());
    assert!(no_op.patch().is_empty());
    apply_commit(&mut package, no_op).unwrap();
    assert_eq!(package.get_part(&name).unwrap().blob(), source.source_xml());

    let mut edit = source.edit();
    edit.set_time(0, Offset::ms(1500)).unwrap();
    edit.set_object_id(1, 7).unwrap();
    edit.set_kind(1, Kind::Play).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains("time=\"1500\""));
    assert!(output.contains("objId=\"7\""));
    assert!(output.contains("<p14:playEvt"));
    assert!(output.contains("<x:data>keep</x:data>"));
    apply_commit(&mut package, commit).unwrap();
    assert_eq!(load(&package, &name).unwrap().unwrap()[0].object_id(), 6);
}

#[test]
fn unknown_event_attributes_and_opaque_extensions_survive_attribute_edits() {
    let name = slide_name();
    let mut package = fixture();
    store(&mut package, &name, &drafts()).unwrap();
    let part = package.get_part_mut(&name).unwrap();
    let mut xml = part.blob().to_vec();
    let needle = br#"<p14:triggerEvt"#;
    let at = xml
        .windows(needle.len())
        .position(|value| value == needle)
        .unwrap();
    xml.splice(
        at..at + needle.len(),
        br#"<p14:triggerEvt future="keep""#.iter().copied(),
    );
    part.set_blob(xml);

    let source = load_snapshot(&package, &name).unwrap().unwrap();
    let mut edit = source.edit();
    edit.set_time(0, Offset::ms(2500)).unwrap();
    let commit = edit.commit().unwrap();
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains(r#"future="keep""#));
    assert!(output.contains("<x:data>keep</x:data>"));
}

#[test]
fn list_crud_inverse_stale_rejection_and_remove_are_atomic() {
    let name = slide_name();
    let mut package = fixture();
    store(&mut package, &name, &drafts()).unwrap();
    let original = load_snapshot(&package, &name).unwrap().unwrap();

    let mut edit = original.edit();
    edit.insert(1, Draft::null(Offset::ms(4), 9)).unwrap();
    let removed = edit.remove(1).unwrap();
    assert_eq!(removed.kind(), &Kind::Null);
    edit.push(Draft::play(Offset::secs(5), 4)).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();

    let part = package.get_part_mut(&name).unwrap();
    let mut stale = part.blob().to_vec();
    stale.extend_from_slice(b" ");
    part.set_blob(stale.clone());
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(package.get_part(&name).unwrap().blob(), stale.as_slice());

    let mut restored_package = fixture();
    store(&mut restored_package, &name, &drafts()).unwrap();
    patch.apply(&mut restored_package).unwrap();
    patch.inverse().apply(&mut restored_package).unwrap();
    assert_eq!(
        restored_package.get_part(&name).unwrap().blob(),
        original.source_xml()
    );
    let removed = remove(&mut restored_package, &name).unwrap().unwrap();
    assert_eq!(removed.source_xml(), original.source_xml());
    assert!(load_snapshot(&restored_package, &name).unwrap().is_none());
    let xml = String::from_utf8_lossy(restored_package.get_part(&name).unwrap().blob());
    assert!(xml.contains("<x:data>keep</x:data>"));
}

#[test]
fn failed_edits_leave_the_staged_sequence_unchanged() {
    let name = slide_name();
    let mut package = fixture();
    store(&mut package, &name, &drafts()).unwrap();
    let source = load_snapshot(&package, &name).unwrap().unwrap();
    let mut edit = source.edit();
    let before = edit.events().to_vec();
    assert!(edit.replace(Vec::new()).is_err());
    assert_eq!(edit.events(), before.as_slice());
    assert!(edit.set_time(99, Offset::ZERO).is_err());
    assert_eq!(edit.events(), before.as_slice());
}

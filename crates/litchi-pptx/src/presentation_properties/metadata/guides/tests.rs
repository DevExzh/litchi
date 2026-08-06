use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::{Color, ColorKind, Guide, Guides, ListKind, Orientation, Snapshot};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_URI: &str = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}";
const NOTES_URI: &str = "{2D200454-40CA-4A62-9FC3-DE9A4176ACB9}";

fn source() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:p15="{P15}" xmlns:r="{R}">
  <p:extLst>
    <p:ext uri="{{future}}"><v:opaque xmlns:v="urn:future" r:id="rIdOpaque">retain</v:opaque></p:ext>
    <p:ext uri="{SLIDE_URI}">
      <p15:sldGuideLst>
        <p15:guide id="1" name="Middle" orient="horz" pos="2160" userDrawn="1">
          <p15:clr><a:srgbClr val="AABBCC"/></p15:clr>
          <p:extLst><p:ext uri="urn:guide"><v:data xmlns:v="urn:future">keep-guide</v:data></p:ext></p:extLst>
        </p15:guide>
        <p:extLst><p:ext uri="urn:list"><v:data xmlns:v="urn:future">keep-list</v:data></p:ext></p:extLst>
      </p15:sldGuideLst>
    </p:ext>
    <p:ext uri="{NOTES_URI}"><p15:notesGuideLst/></p:ext>
  </p:extLst>
</p:presentation>"#
    )
    .into_bytes()
}

fn presentation_name() -> PackURI {
    PackURI::new("/ppt/presentation.xml").expect("valid presentation part name")
}

fn package() -> OpcPackage {
    let name = presentation_name();
    let mut package = OpcPackage::new();
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(BlobPart::new(
        name,
        ct::PML_PRESENTATION_MAIN.into(),
        source(),
    )));
    package
}

fn color() -> Color {
    Color {
        kind: ColorKind::Srgb,
        xml: format!(r#"<a:srgbClr xmlns:a="{A}" val="DDEEFF"/>"#).into_bytes(),
    }
}

fn guide(id: u32) -> Guide {
    Guide {
        id,
        name: Some("New".into()),
        orientation: Some(Orientation::Vertical),
        position: Some(12),
        user_drawn: Some(false),
        color: color(),
        extension_xml: None,
    }
}

#[test]
fn no_op_commit_keeps_exact_source_and_raw_payloads() {
    let source = source();
    let snapshot = Snapshot::from_xml(&source).expect("guide source parses");
    let commit = snapshot.edit().commit().expect("no-op commit");

    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), source.as_slice());
    assert!(
        std::str::from_utf8(
            commit.snapshot().guides().slide.as_ref().unwrap().guides[0]
                .extension_xml
                .as_ref()
                .unwrap()
        )
        .unwrap()
        .contains("keep-guide")
    );
}

#[test]
fn typed_edits_preserve_opaque_extensions_and_update_only_guides() {
    let snapshot = Snapshot::from_xml(&source()).expect("guide source parses");
    let mut edit = snapshot.edit();
    edit.set_id(ListKind::Slide, 0, 7)
        .expect("ID edit is valid");
    edit.edit_guide(ListKind::Slide, 0, |guide| {
        guide.position = Some(2400);
        Ok(())
    })
    .expect("position edit is valid");
    edit.push(ListKind::Notes, guide(8))
        .expect("notes guide insertion is valid");
    let commit = edit.commit().expect("typed guide commit");
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());

    assert!(output.contains(r#"uri="{future}"#));
    assert!(output.contains(r#"r:id="rIdOpaque"#));
    assert!(output.contains("keep-guide"));
    assert!(output.contains("keep-list"));
    assert!(output.contains(r#"id="7"#));
    assert!(output.contains(r#"pos="2400"#));
    assert!(output.matches(r#"<p15:guide"#).count() == 2);

    let parsed = Guides::from_xml(commit.snapshot().source_xml()).expect("committed XML parses");
    assert_eq!(parsed, *commit.snapshot().guides());
}

#[test]
fn invalid_ids_and_bounds_leave_the_staged_value_unchanged() {
    let snapshot = Snapshot::from_xml(&source()).expect("guide source parses");
    let mut edit = snapshot.edit();
    let before = edit.guides().clone();

    assert!(edit.push(ListKind::Slide, guide(1)).is_err());
    assert_eq!(edit.guides(), &before);
    assert!(edit.insert(ListKind::Slide, 9, guide(2)).is_err());
    assert_eq!(edit.guides(), &before);
    assert!(edit.set_id(ListKind::Slide, 9, 2).is_err());
    assert_eq!(edit.guides(), &before);
    assert!(edit.remove(ListKind::Slide, 9).is_err());
    assert_eq!(edit.guides(), &before);
}

#[test]
fn stale_publication_is_atomic_and_inverse_restores_the_source() {
    let name = presentation_name();
    let mut original = package();
    let snapshot = super::load_snapshot(&original).expect("package snapshot");
    let mut edit = snapshot.edit();
    edit.set_id(ListKind::Slide, 0, 7)
        .expect("ID edit is valid");
    let patch = edit.commit().expect("guide commit").into_patch();

    let mut stale = package();
    let mut stale_xml = stale.get_part(&name).unwrap().blob().to_vec();
    stale_xml.extend_from_slice(b" ");
    stale
        .get_part_mut(&name)
        .unwrap()
        .set_blob(stale_xml.clone());
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(stale.get_part(&name).unwrap().blob(), stale_xml.as_slice());

    patch.apply(&mut original).expect("forward guide patch");
    patch
        .inverse()
        .apply(&mut original)
        .expect("inverse guide patch");
    assert_eq!(
        original.get_part(&name).unwrap().blob(),
        source().as_slice()
    );
}

#[test]
fn absent_extension_lists_are_created_and_xml_inverse_is_exact() {
    let source = format!(r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}"/>"#).into_bytes();
    let snapshot = Snapshot::from_xml(&source).expect("empty presentation parses");
    let mut edit = snapshot.edit();
    edit.push(ListKind::Slide, guide(3))
        .expect("slide guide insertion is valid");
    let patch = edit.commit().expect("guide insertion commit").into_patch();

    let mut updated = source.clone();
    patch.apply_xml(&mut updated).expect("forward XML patch");
    let parsed = Guides::from_xml(&updated).expect("created guide extension parses");
    assert_eq!(parsed.slide.as_ref().unwrap().guides[0].id, 3);
    assert!(String::from_utf8_lossy(&updated).contains("sldGuideLst"));

    patch
        .inverse()
        .apply_xml(&mut updated)
        .expect("inverse XML patch");
    assert_eq!(updated, source);
}

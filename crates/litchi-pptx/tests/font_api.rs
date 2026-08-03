use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_pptx::font::{self, Data, Face, Font, Fonts, Style};
use std::sync::Arc;

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let presentation = PackURI::new("/ppt/presentation.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        presentation,
        ct::PML_PRESENTATION_MAIN.into(),
        br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst/><p:defaultTextStyle/></p:presentation>"#.to_vec(),
    )));
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "ppt/presentation.xml".into(),
        "rId1".into(),
        false,
    );
    package
}

fn eot() -> Vec<u8> {
    let mut value = vec![0; 96];
    value[0..4].copy_from_slice(&108u32.to_le_bytes());
    value[4..8].copy_from_slice(&12u32.to_le_bytes());
    value[8..12].copy_from_slice(&0x0001_0000u32.to_le_bytes());
    value[34..36].copy_from_slice(&0x504Cu16.to_le_bytes());
    value.extend_from_slice(&[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    value
}

#[test]
fn semantic_crud_is_concise_checked_and_zero_copy_shared() {
    let shared = Data::powerpoint(eot()).unwrap();
    let font = Font::from_face("Example Sans", Face::new(Style::Regular, shared.clone()))
        .unwrap()
        .with(Face::new(Style::Bold, shared))
        .unwrap();
    let mut fonts = Fonts::new();
    fonts.add(font).unwrap();

    assert_eq!(fonts.get("example sans").unwrap().name(), "Example Sans");
    assert_eq!(fonts.get(0_usize).unwrap().faces().len(), 2);
    assert!(fonts.get(1_usize).is_err());
    assert!(fonts.add(Font::new("EXAMPLE SANS").unwrap()).is_err());

    let mut package = package();
    assert!(font::put(&mut package, fonts).unwrap());
    let loaded = font::load(&package).unwrap().unwrap();
    let loaded = loaded.get("Example Sans").unwrap();
    let regular = loaded
        .get(Style::Regular)
        .unwrap()
        .data()
        .clone()
        .into_shared();
    let bold = loaded
        .get(Style::Bold)
        .unwrap()
        .data()
        .clone()
        .into_shared();
    assert!(Arc::ptr_eq(&regular, &bold));
    let unchanged = font::load(&package).unwrap().unwrap();
    assert!(!font::put(&mut package, unchanged).unwrap());

    let removed = font::remove(&mut package).unwrap().unwrap();
    assert_eq!(removed.len(), 1);
    assert!(font::load(&package).unwrap().is_none());
}

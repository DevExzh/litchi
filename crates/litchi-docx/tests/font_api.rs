use litchi_docx::font::{
    Charset, Conformance, Embed, Family, Font, Key, Pitch, Resource, Signature, Style, Table, put,
    read, remove, write,
};
use litchi_opc::{OpcPackage, PackURI, XmlPart};

const KEY: &str = "{00112233-4455-6677-8899-AABBCCDDEEFF}";

#[test]
fn downstream_surface_is_short_safe_and_move_first() -> litchi_docx::Result<()> {
    let face = Embed::new(Style::Regular, KEY, Resource::new(vec![0; 32])?)?;
    let font = Font::new("Example")?
        .with_charset(Charset::Ansi)
        .with_family(Family::Swiss)
        .with_pitch(Pitch::Variable)
        .with_signature(Signature::new([0; 4], [0; 2]))
        .with_embed(face)?;

    let mut table = Table::new();
    table.add(font)?;
    assert_eq!(table.get("example")?.map(Font::name), Some("Example"));
    assert_eq!(table.get(Key::Index(0))?.map(Font::name), Some("Example"));
    assert!(table.get(usize::MAX)?.is_none());

    let old = table.replace("Example", Font::new("Replacement")?)?;
    assert_eq!(
        old.map(|font| font.name().to_owned()),
        Some("Example".into())
    );
    assert!(write(&table, Conformance::Strict)?.starts_with(b"<?xml"));

    let mut package = OpcPackage::new();
    let document = PackURI::new("/word/document.xml").expect("test URI");
    package.add_part(Box::new(XmlPart::new(
        document,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
    )));
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".into(),
        "word/document.xml".into(),
        "rId1".into(),
        false,
    );
    put(&mut package, table, Conformance::Transitional)?;
    assert!(read(&package)?.is_some());
    assert!(remove(&mut package)?);
    assert!(!remove(&mut package)?);
    Ok(())
}

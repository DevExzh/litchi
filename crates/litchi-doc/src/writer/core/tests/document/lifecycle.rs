use super::super::support::*;

#[test]
fn test_create_writer() {
    let writer = Builder::new();
    assert_eq!(writer.paragraphs.len(), 0);
    assert_eq!(writer.tables.len(), 0);
}

#[test]
fn file_and_seekable_outputs_are_byte_identical() {
    let mut writer = Builder::new();
    writer
        .set_metadata(litchi_core::Metadata {
            title: Some("Canonical output".to_owned()),
            ..litchi_core::Metadata::default()
        })
        .unwrap();
    writer.add_paragraph("One output assembly path").unwrap();

    let mut memory = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut memory).unwrap();

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-output-equivalence-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.build().unwrap().save(&path).unwrap();
    let file = std::fs::read(&path).unwrap();
    std::fs::remove_file(path).unwrap();

    assert_eq!(file, memory.into_inner());
}

#[test]
fn test_set_metadata() {
    let mut writer = Builder::new();
    writer
        .set_metadata(litchi_core::Metadata {
            title: Some("Test Document".to_owned()),
            author: Some("Test Author".to_owned()),
            ..litchi_core::Metadata::default()
        })
        .unwrap();
    let summary = writer.summary_properties().summary_information();
    assert_eq!(summary.title(), Some("Test Document"));
    assert_eq!(summary.author(), Some("Test Author"));
    let refused = writer.set_metadata(litchi_core::Metadata {
        identifier: Some("urn:x".to_owned()),
        ..litchi_core::Metadata::default()
    });
    assert!(matches!(refused, Err(WriteError::Metadata(_))));
    assert_eq!(
        writer.summary_properties().summary_information().title(),
        Some("Test Document")
    );
}

#[test]
fn test_write_to_memory() {
    let mut writer = Builder::new();
    writer.add_paragraph("Test paragraph").unwrap();
    let mut cursor = Cursor::new(Vec::new());
    let result = writer.build().unwrap().write_to(&mut cursor);
    assert!(result.is_ok());
    assert!(!cursor.into_inner().is_empty());
}

#[test]
fn test_empty_document_write() {
    let mut writer = Builder::new();
    let mut cursor = Cursor::new(Vec::new());
    let result = writer.build().unwrap().write_to(&mut cursor);
    assert!(result.is_ok());
    let data = cursor.into_inner();
    assert!(!data.is_empty());
    let mut package = crate::raw::Reader::from_reader(Cursor::new(data)).unwrap();
    assert_eq!(package.document().unwrap().text().unwrap(), "\r");
}

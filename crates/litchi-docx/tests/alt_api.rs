use litchi_docx::alt::{Chunk, Conformance, Data, Import, Kind, Rel, Uri};

#[test]
fn downstream_api_is_short_typed_and_move_owned() {
    let bytes = vec![1, 2, 3, 4];
    let pointer = bytes.as_ptr();
    let import = Import::data(Data::Docx(bytes));
    let Import::Data(data) = import else {
        panic!("expected an internal payload")
    };
    assert_eq!(
        data.media_type(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    );
    assert_eq!(Kind::from_media_type(data.media_type()), Kind::Docx);
    let moved = data.into_bytes();
    assert_eq!(moved.as_ptr(), pointer);

    let relationship = Rel::new("rIdAlt1").unwrap();
    let chunk = Chunk::new(relationship, Some(false));
    assert_eq!(chunk.relationship().as_str(), "rIdAlt1");
    assert!(chunk.xml(Conformance::Strict).contains("w:val=\"0\""));

    let uri = Uri::new("https://example.invalid/import.xhtml").unwrap();
    assert_eq!(uri.as_str(), "https://example.invalid/import.xhtml");
}

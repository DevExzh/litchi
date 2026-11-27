//! The facade exposes the same typed unified read API as the operations owner.
#![cfg(any(feature = "docx", feature = "xlsx"))]

#[cfg(feature = "docx")]
#[test]
fn facade_file_resolves_nodes_and_streams_queries() {
    let file = litchi::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/docx/drawing.docx"
    ))
    .unwrap();
    assert!(
        file.capabilities()
            .verb(litchi::capabilities::Verb::Get)
            .is_supported()
    );
    let path = litchi::path::Path::parse("/body").unwrap();
    assert_eq!(file.get(&path).unwrap().kind(), "body");
    let query = litchi::query::Query::parse("p").unwrap();
    let node: litchi::node::Node<'_> = file.query(&query).unwrap().first().unwrap().unwrap();
    assert_eq!(file.get(node.path().to_path()).unwrap().kind(), node.kind());
    assert!(!file.view(litchi::view::Mode::Stats).unwrap().is_empty());
}

#[cfg(feature = "xlsx")]
#[test]
fn facade_xlsx_reads_use_the_same_positional_node_and_query_api() {
    let file = litchi::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsx/SimpleNormal.xlsx"
    ))
    .unwrap();
    assert_eq!(file.root().unwrap().kind(), "workbook");
    assert!(
        file.capabilities()
            .verb(litchi::capabilities::Verb::Query)
            .is_supported()
    );
    let sheet = file.get("/sheet[1]").unwrap();
    assert_eq!(sheet.path().to_string(), "/sheet[1]");
    let cell: litchi::node::Node<'_> = file.query("cell").unwrap().first().unwrap().unwrap();
    assert_eq!(file.get(cell.path().to_path()).unwrap().path(), cell.path());
    assert!(!file.view(litchi::view::Mode::Stats).unwrap().is_empty());
}

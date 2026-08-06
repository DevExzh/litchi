#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_odb::{Builder, Database, connection::Connection, query::Query};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    let query = Query::new("recent", "SELECT * FROM records");
    assert_eq!(query.name(), "recent");
    assert_eq!(query.command(), "SELECT * FROM records");
    assert!(matches!(
        Connection::File("records.db".into()),
        Connection::File(_)
    ));

    let bytes = Builder::new().build().unwrap();
    let database = Database::from_bytes(bytes).unwrap();
    assert!(database.content_xml().contains("<office:database"));
}

use litchi_odf::{
    DatabaseDocument, OdfDatabaseQueries, OdfDatabaseQuery, OdfDatabaseQueryCollection,
    OdfDatabaseQueryItem, OdfDatabaseStatement,
};
use std::io::{Cursor, Read, Write};

const MIME: &str = "application/vnd.oasis.opendocument.base";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

fn content(queries: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:db="{DB}" xmlns:xlink="{XLINK}"><office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href="sdbc:embedded:firebird" xlink:type="simple"/></db:connection-data></db:data-source><db:forms/><!--preserve-->{queries}<db:table-representations/><db:schema-definition><db:table-definitions/></db:schema-definition></office:database></office:body></office:document-content>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    let manifest = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="database/" manifest:media-type=""/><manifest:file-entry manifest:full-path="database/data" manifest:media-type="application/x-firebird"/><manifest:file-entry manifest:full-path="forms/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="forms/Form1/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="reports/" manifest:media-type="application/vnd.sun.xml.report"/><manifest:file-entry manifest:full-path="reports/Report1/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Basic/" manifest:media-type="application/vnd.sun.star.basic-library"/><manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Dialogs/" manifest:media-type="application/vnd.sun.star.dialog-library"/><manifest:file-entry manifest:full-path="Dialogs/dialog-lc.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Configurations2/" manifest:media-type="application/vnd.sun.xml.ui.configuration"/><manifest:file-entry manifest:full-path="unknown.bin" manifest:media-type="application/x-litchi-unknown"/></manifest:manifest>"#
    );
    let entries: [(&str, &[u8]); 11] = [
        ("content.xml", content.as_bytes()),
        ("settings.xml", b"<settings keep='yes'/>") ,
        ("styles.xml", b"<styles keep='yes'/>") ,
        ("meta.xml", b"<meta keep='yes'/>") ,
        ("database/data", b"opaque database bytes"),
        ("forms/Form1/content.xml", b"<form keep='yes'/>") ,
        ("reports/Report1/content.xml", b"<report keep='yes'/>") ,
        ("Basic/Standard/script-lb.xml", b"<basic inert='yes'/>") ,
        ("Dialogs/dialog-lc.xml", b"<dialogs keep='yes'/>") ,
        ("unknown.bin", b"unknown bytes"),
        ("META-INF/manifest.xml", manifest.as_bytes()),
    ];
    let mut output = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut output);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(MIME.as_bytes()).unwrap();
        for (path, bytes) in entries {
            zip.start_file(path, deflated).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    output.into_inner()
}

fn remove_empty_scripts_from_fixture(bytes: Vec<u8>) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut entries = Vec::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        if entry.name() == "content.xml" {
            let xml = String::from_utf8(bytes).unwrap();
            bytes = xml.replace("<office:scripts/>", "").into_bytes();
        }
        entries.push((entry.name().to_string(), bytes, entry.compression()));
    }
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut output);
        for (name, bytes, compression) in entries {
            let options = zip::write::SimpleFileOptions::default().compression_method(compression);
            if name.ends_with('/') {
                writer.add_directory(name, options).unwrap();
            } else {
                writer.start_file(name, options).unwrap();
                writer.write_all(&bytes).unwrap();
            }
        }
        writer.finish().unwrap();
    }
    output.into_inner()
}

fn recursive_queries() -> OdfDatabaseQueries {
    let mut query = OdfDatabaseQuery::new("Recent", "SELECT * FROM records");
    query.escape_processing = Some(false);
    query.order_statement = Some(OdfDatabaseStatement::new("id DESC"));
    query.filter_statement = Some(OdfDatabaseStatement::new("active = TRUE"));
    let mut collection = OdfDatabaseQueryCollection::new("Reports");
    collection.items.push(OdfDatabaseQueryItem::Query(query));
    OdfDatabaseQueries {
        items: vec![OdfDatabaseQueryItem::Collection(collection)],
    }
}

#[test]
fn packaged_query_insert_replace_remove_preserves_every_unrelated_part() {
    let original = package(&content(""));
    let mut database = DatabaseDocument::from_bytes(original).unwrap();
    let preserved = [
        "settings.xml",
        "styles.xml",
        "meta.xml",
        "database/data",
        "forms/Form1/content.xml",
        "reports/Report1/content.xml",
        "Basic/Standard/script-lb.xml",
        "Dialogs/dialog-lc.xml",
        "unknown.bin",
    ]
    .map(|path| (path, database.get_file(path).unwrap()));

    assert_eq!(database.set_queries(Some(&recursive_queries())).unwrap(), None);
    let parsed = database.queries().unwrap().unwrap();
    assert_eq!(parsed, recursive_queries());
    let xml = String::from_utf8(database.get_file("content.xml").unwrap()).unwrap();
    assert!(xml.contains("<!--preserve--><db:queries"));
    assert!(xml.find("<db:queries").unwrap() < xml.find("<db:table-representations").unwrap());

    for (path, bytes) in preserved {
        assert_eq!(database.get_file(path).unwrap(), bytes, "changed {path}");
    }
    let manifest = String::from_utf8(database.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    for entry in [
        "database/",
        "forms/",
        "reports/",
        "Basic/",
        "Dialogs/",
        "Configurations2/",
        "application/x-firebird",
        "application/x-litchi-unknown",
    ] {
        assert!(manifest.contains(entry), "manifest lost {entry}");
    }

    let previous = database.set_queries(None).unwrap().unwrap();
    assert_eq!(previous, recursive_queries());
    assert!(database.queries().unwrap().is_none());
}

#[test]
fn mutates_real_libreoffice_query_package_without_touching_embedded_resources() {
    let bytes = include_bytes!(
        "../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb"
    )
    .to_vec();
    let bytes = remove_empty_scripts_from_fixture(bytes);
    let mut database = DatabaseDocument::from_bytes(bytes).unwrap();
    let embedded = database
        .embedded_database_files()
        .unwrap()
        .into_iter()
        .map(|path| {
            let bytes = database.get_file(&path).unwrap();
            (path, bytes)
        })
        .collect::<Vec<_>>();
    let settings = database.get_file("settings.xml").unwrap();
    let old = database.queries().unwrap().unwrap();

    database.set_queries(Some(&recursive_queries())).unwrap();
    assert_ne!(database.queries().unwrap().unwrap(), old);
    for (path, bytes) in embedded {
        assert_eq!(database.get_file(&path).unwrap(), bytes);
    }
    assert_eq!(database.get_file("settings.xml").unwrap(), settings);
    let manifest = String::from_utf8(database.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    assert!(manifest.contains("Configurations2/"));
    DatabaseDocument::from_bytes(database.to_bytes()).unwrap();
}

#[test]
fn invalid_query_update_is_atomic_and_commands_remain_inert() {
    let mut database = DatabaseDocument::from_bytes(package(&content(""))).unwrap();
    let before = database.to_bytes();
    let mut invalid = recursive_queries();
    let OdfDatabaseQueryItem::Collection(collection) = &mut invalid.items[0] else {
        unreachable!()
    };
    let OdfDatabaseQueryItem::Query(query) = &mut collection.items[0] else {
        unreachable!()
    };
    query.command = "x".repeat(1024 * 1024 + 1);
    assert!(database.set_queries(Some(&invalid)).is_err());
    assert_eq!(database.to_bytes(), before);
    assert!(database.queries().unwrap().is_none());

    let mut inert = recursive_queries();
    let OdfDatabaseQueryItem::Collection(collection) = &mut inert.items[0] else {
        unreachable!()
    };
    let OdfDatabaseQueryItem::Query(query) = &mut collection.items[0] else {
        unreachable!()
    };
    query.command = "SELECT * FROM http://127.0.0.1:9/never-fetch".to_string();
    database.set_queries(Some(&inert)).unwrap();
    assert_eq!(database.queries().unwrap().unwrap(), inert);
}

#[test]
fn spoofed_and_dtd_database_xml_are_rejected_before_mutation() {
    let spoofed = content("").replace(DB, "urn:spoofed-database");
    assert!(DatabaseDocument::from_bytes(package(&spoofed)).is_err());
    let dtd = format!("<!DOCTYPE x>{}", content(""));
    let mut database = DatabaseDocument::from_bytes(package(&dtd)).unwrap();
    let before = database.to_bytes();
    assert!(database.set_queries(Some(&recursive_queries())).is_err());
    assert_eq!(database.to_bytes(), before);
}

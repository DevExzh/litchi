//! Shared-parser compatibility over compact in-memory ODT packages.

#![cfg_attr(
    test,
    allow(
        clippy::unwrap_used,
        reason = "Fixed in-memory package fixtures use direct assertion setup."
    )
)]

use litchi_core::Error;
use litchi_odf_common::core::{Package, PackageWriter};
use litchi_odf_common::style::{data, master};
use litchi_odf_common::{calculation, embedded, media};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.0"><office:body><office:text><text:p>body</text:p><draw:frame draw:name="image" svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>aQ==</office:binary-data></draw:image></draw:frame></office:text></office:body></office:document-content>"#;
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.0"><office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Header</text:p></style:header></style:master-page></office:master-styles></office:document-styles>"#;
const META: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.0"><office:meta><dc:title>Compact</dc:title></office:meta></office:document-meta>"#;

fn package(content: &str) -> Result<Vec<u8>, Error> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
    writer.add_file("content.xml", content.as_bytes())?;
    writer.add_file("styles.xml", STYLES.as_bytes())?;
    writer.add_file("meta.xml", META.as_bytes())?;
    writer.finish_to_bytes()
}

fn assert_invalid<T>(result: Result<T, Error>, case: &str) {
    match result {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("{case} produced a non-format error: {error:?}"),
        Ok(_) => panic!("{case} unexpectedly parsed"),
    }
}

#[test]
fn compact_odt_exercises_shared_package_and_xml_parsers() -> Result<(), Error> {
    let package = Package::from_bytes(
        package(CONTENT)?,
        "application/vnd.oasis.opendocument.text",
        "<office:text",
        "ODT",
    )?;
    let archive = package.package().package()?;
    assert_eq!(
        media::scan_package(package.content_xml(), package.styles_xml(), &archive)?.len(),
        1
    );
    drop(embedded::scan_package(
        package.content_xml(),
        package.styles_xml(),
        &archive,
    )?);
    drop(calculation::parse(package.content_xml())?);
    drop(data::parse_package(
        package.styles_xml(),
        package.content_xml(),
    )?);
    let Some(styles) = package.styles_xml() else {
        panic!("compact package has no styles.xml");
    };
    assert!(!master::reader::read(styles)?.is_empty());
    Ok(())
}

#[test]
fn compact_xxe_probe_is_rejected_by_shared_scanners() -> Result<(), Error> {
    let hostile = r#"<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE office [<!ENTITY nasty SYSTEM "externalcontent.txt">]><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>&nasty;</text:p></office:text></office:body></office:document-content>"#;
    let package = Package::from_bytes(
        package(hostile)?,
        "application/vnd.oasis.opendocument.text",
        "<office:text",
        "ODT",
    )?;
    let archive = package.package().package()?;
    assert_invalid(
        media::scan_package(package.content_xml(), package.styles_xml(), &archive),
        "XXE image scan",
    );
    assert_invalid(
        embedded::scan_package(package.content_xml(), package.styles_xml(), &archive),
        "XXE object scan",
    );
    Ok(())
}

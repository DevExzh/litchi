//! Public Data Model opaque-extension context and inverse regressions.
#![allow(clippy::unwrap_used, reason = "fixture assertions intentionally panic")]

use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::Package;

const SOURCE: &[u8] = include_bytes!("data/data_model/tdf167689_x15_namespace.xlsx");
const WORKBOOK: &str = "/xl/workbook.xml";
const SOURCE_BASE: &str = "https://source.example/root/";
const DESTINATION_BASE: &str = "https://destination.example/root/";
const MAX_EXTENSION_BYTES: usize = 4 * 1024 * 1024;

#[test]
fn public_parser_rejects_oversized_opaque_source_before_child_dom() {
    let text = "x".repeat(MAX_EXTENSION_BYTES);
    let xml = format!(
        "<m:dataModel xmlns:m='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'><m:extLst>{text}</m:extLst></m:dataModel>"
    );
    assert!(litchi_xlsx::parse_data_model(xml.as_bytes()).is_err());
}

#[test]
fn public_parser_rejects_oversized_in_scope_namespace_before_unescape() {
    let namespace = "x".repeat(MAX_EXTENSION_BYTES);
    let xml = format!(
        "<m:dataModel xmlns:m='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main' xmlns:q='{namespace}'><m:extLst/></m:dataModel>"
    );
    assert!(litchi_xlsx::parse_data_model(xml.as_bytes()).is_err());
}

#[test]
fn public_parser_validates_opaque_attribute_entities_without_rewriting_them() {
    let xml = "<m:dataModel xmlns:m='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'><m:extLst><v:link xmlns:v='urn:v&amp;x' href='a&amp;b' code='&#x41;'/></m:extLst></m:dataModel>";
    let definition = litchi_xlsx::parse_data_model(xml.as_bytes()).unwrap();
    let extension = definition.extension_list.unwrap().xml;
    let extension = String::from_utf8(extension).unwrap();
    assert!(extension.contains("xmlns:v='urn:v&amp;x'"));
    assert!(extension.contains("href='a&amp;b'"));
    assert!(extension.contains("code='&#x41;'"));
}

#[test]
fn public_parser_rejects_undeclared_or_malformed_opaque_attribute_entities() {
    for value in ["a&bogus;", "a&broken"] {
        let xml = format!(
            "<m:dataModel xmlns:m='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'><m:extLst><v:link xmlns:v='urn:v' href='{value}'/></m:extLst></m:dataModel>"
        );
        assert!(
            litchi_xlsx::parse_data_model(xml.as_bytes()).is_err(),
            "accepted opaque attribute value {value:?}"
        );
    }

    let xml = "<m:dataModel xmlns:m='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'><m:extLst><v:link xmlns:v='urn:v&bogus;'/></m:extLst></m:dataModel>";
    assert!(litchi_xlsx::parse_data_model(xml.as_bytes()).is_err());
}

#[test]
fn public_model_import_closes_inherited_xml_context_without_rebasing_nested_overrides() {
    let source_bytes = source_with_context(SOURCE, SOURCE_BASE);
    let source = Package::from_bytes(source_bytes.clone()).unwrap();
    let model = source.data_model().unwrap().model().unwrap().to_owned();
    let extension = String::from_utf8(
        model
            .definition
            .extension_list
            .as_ref()
            .unwrap()
            .xml
            .clone(),
    )
    .unwrap();
    assert!(extension.contains("xml:base=\"https://source.example/root/\""));
    assert!(extension.contains("xml:lang=\"fr\""));
    assert!(extension.contains("xml:space=\"preserve\""));
    assert!(extension.contains("xml:base=\"../other/\""));

    let mut without_model = source;
    without_model.remove_data_model().unwrap();
    let mut destination_bytes = Vec::new();
    without_model.write_to(&mut destination_bytes).unwrap();
    let destination_bytes = replace_workbook(&destination_bytes, |xml| {
        xml.replace(SOURCE_BASE, DESTINATION_BASE)
    });
    let mut destination = Package::from_bytes(destination_bytes).unwrap();
    destination.put_data_model(model).unwrap();

    let imported = String::from_utf8(
        destination
            .data_model()
            .unwrap()
            .model()
            .unwrap()
            .definition()
            .extension_list
            .as_ref()
            .unwrap()
            .xml
            .clone(),
    )
    .unwrap();
    assert!(imported.contains("xml:base=\"https://source.example/root/\""));
    assert!(!imported.contains("xml:base=\"https://destination.example/root/\""));
    assert!(imported.contains("xml:base=\"../other/\""));
}

#[test]
fn public_model_context_capture_survives_descriptor_edit_and_exact_inverse() {
    let source_bytes = source_with_context(SOURCE, SOURCE_BASE);
    let original = OpcPackage::from_bytes(&source_bytes).unwrap();
    let original_workbook = original
        .get_part(&PackURI::new(WORKBOOK).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let mut package = Package::from_bytes(source_bytes).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.edit_definition(|definition| {
        definition.min_version_load += 1;
        Ok(())
    })
    .unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    let mut raw = package.into_plain_opc();
    patch.inverse().apply(&mut raw).unwrap();
    assert_eq!(
        raw.get_part(&PackURI::new(WORKBOOK).unwrap())
            .unwrap()
            .blob(),
        original_workbook
    );
}

fn source_with_context(source: &[u8], base: &str) -> Vec<u8> {
    replace_workbook(source, |xml| {
        let xml = xml.replace(
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
            &format!(
                "<workbook xml:base=\"{base}\" xml:lang=\"fr\" xml:space=\"preserve\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
            ),
        );
        xml.replace(
            "</x15:modelTables>",
            "</x15:modelTables><x15:extLst><v:link xmlns:v=\"urn:v\" href=\"item.xml\"/><v:nested xmlns:v=\"urn:v\" xml:base=\"../other/\" href=\"nested.xml\"/></x15:extLst>",
        )
    })
}

fn replace_workbook(source: &[u8], replace: impl FnOnce(String) -> String) -> Vec<u8> {
    let archive = soapberry_zip::office::ArchiveReader::new(source).unwrap();
    let workbook =
        replace(String::from_utf8(archive.read("xl/workbook.xml").unwrap()).unwrap()).into_bytes();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let bytes = if name == "xl/workbook.xml" {
            workbook.clone()
        } else {
            archive.read(name).unwrap()
        };
        writer.write_deflated(name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

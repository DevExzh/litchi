//! Native XLS parents with an explicitly authored XML stream, using public codec/CFB owners.
use litchi_ole_common::object;
use litchi_xls::{Package, xml_map};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
const XML: &str = r#"<MapInfo SelectionNamespaces="xmlns:n='urn:litchi:native-map:071'"><Schema ID="schema-71" Namespace="urn:litchi:native-map:071"><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="urn:litchi:native-map:071"><xs:import namespace="urn:external" schemaLocation="https://example.invalid/never-fetched.xsd"/><xs:element name="root" type="xs:string"/></xs:schema></Schema><Map ID="71" Name="Native Map 🍎" RootElement="root" SchemaID="schema-71" ShowImportExportValidationErrors="true" AutoFit="false" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"><DataBinding DataBindingName="binding-71" FileBinding="file:///never-opened/source.xml" FileBindingName="source-71" DataBindingLoadMode="1"><n:retained xmlns:n="urn:litchi:binding" value="A &amp; B"/></DataBinding></Map></MapInfo>"#;
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let b = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, b)
        })
        .collect()
}
fn preserved(original: &[u8], after: &[u8]) {
    let a = streams(original);
    let b = streams(after);
    for (p, v) in a {
        if p != ["XML"] {
            assert!(b.get(&p) == Some(&v), "changed {}", p.join("/"));
        }
    }
    let a = cfb_metadata::directory(original);
    let b = cfb_metadata::directory(after);
    for (p, v) in a {
        if p != ["XML"] {
            assert_eq!(b.get(&p), Some(&v));
        }
    }
}
fn publish(source: &Package, value: Option<&xml_map::MapInfo>) -> Package {
    let bytes = source.to_bytes().unwrap();
    let p = object::Package::open(bytes.clone()).unwrap();
    let mut e = p.edit().unwrap();
    if let Some(value) = value {
        let path = vec!["XML".into()];
        let data = xml_map::write(value).unwrap();
        if e.stream(&path).is_some() {
            e.put_stream(&path, data).unwrap();
        } else {
            e.add_stream(path, data).unwrap();
        }
    } else {
        e.remove_stream(&["XML".into()]).unwrap();
    }
    let c = e.commit().unwrap();
    let output = c.patch().after();
    preserved(&bytes, output);
    assert!(c.patch().apply(&p).unwrap().patch().after() == output);
    let q = object::Package::open(output.to_vec()).unwrap();
    assert!(c.patch().inverse().apply(&q).unwrap().patch().after() == bytes);
    let p = Package::from_bytes(output).unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    p.save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(q.to_bytes().unwrap() == output);
    assert_eq!(q.reader_ref().unwrap().xml_map(), value);
    q
}
#[test]
fn native_xls_xml_map_codec_and_stream_lifecycle_preserve_cells_macros_and_directory() {
    for file in ["excel-metadata.xls", "excel-recorded-macro.xlt"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.reader_ref().unwrap().xml_map().is_none());
        let value = xml_map::parse(XML.as_bytes()).unwrap();
        let added = publish(&p, Some(&value));
        let map = added
            .reader_ref()
            .unwrap()
            .xml_map()
            .unwrap()
            .map(xml_map::MapId::new(71).unwrap())
            .unwrap();
        assert_eq!(map.name(), "Native Map 🍎");
        assert_eq!(map.schema_id().as_str(), "schema-71");
        let binding = map.data_binding().unwrap();
        assert_eq!(binding.file_binding(), "file:///never-opened/source.xml");
        assert_eq!(binding.load_mode(), xml_map::LoadMode::Normal);
        let updated = xml_map::parse(
            XML.replace("Native Map 🍎", "Renamed Map 日本語")
                .replace("DataBindingLoadMode=\"1\"", "DataBindingLoadMode=\"2\"")
                .as_bytes(),
        )
        .unwrap();
        let changed = publish(&added, Some(&updated));
        assert_eq!(
            changed.reader_ref().unwrap().xml_map().unwrap().maps()[0]
                .data_binding()
                .unwrap()
                .load_mode(),
            xml_map::LoadMode::DelayLoad
        );
        assert_eq!(
            changed.reader_ref().unwrap().xml_map().unwrap().schemas(),
            value.schemas()
        );
        let at = litchi_xls::cell_values::Reference::new(120, 25).unwrap();
        let mut e = changed.edit().unwrap();
        e.insert_cell("Sheet1", at, litchi_xls::cell_values::Value::Number(71.0))
            .unwrap();
        let c = e.commit().unwrap();
        let output_file = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        c.package().save(output_file.path()).unwrap();
        let reopened = Package::open(output_file.path()).unwrap();
        let cell_changed = &reopened;
        assert_eq!(
            cfb_metadata::directory(&changed.to_bytes().unwrap()),
            cfb_metadata::directory(&cell_changed.to_bytes().unwrap())
        );
        assert_eq!(
            cell_changed
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            Some(litchi_core::sheet::CellValue::Float(71.0))
        );
        assert_eq!(cell_changed.reader_ref().unwrap().xml_map(), Some(&updated));
        let before = streams(&changed.to_bytes().unwrap());
        let after = streams(&cell_changed.to_bytes().unwrap());
        for (p, v) in before {
            if p != ["Workbook"] {
                assert!(after.get(&p) == Some(&v));
            }
        }
        assert!(
            cell_changed
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == changed.to_bytes().unwrap()
        );
        let removed = publish(&changed, None);
        assert!(removed.reader_ref().unwrap().xml_map().is_none());
        assert_eq!(streams(&removed.to_bytes().unwrap()), streams(&bytes));
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_xls_malformed_map_reference_is_reported_by_codec_and_package_reader() {
    let bytes = native("excel-metadata.xls");
    let malformed = XML.replace("SchemaID=\"schema-71\"", "SchemaID=\"missing\"");
    assert!(
        matches!(xml_map::parse(malformed.as_bytes()),Err(litchi_xls::Error::InvalidData(m)) if m=="XML map: Map SchemaID points to a missing Schema ID")
    );
    let p = object::Package::open(bytes.clone()).unwrap();
    let mut e = p.edit().unwrap();
    e.add_stream(vec!["XML".into()], malformed.into_bytes())
        .unwrap();
    let c = e.commit().unwrap();
    assert!(
        matches!(Package::from_bytes(c.patch().after()),Err(litchi_xls::package::Error::Reader(litchi_xls::Error::InvalidData(m))) if m=="XML map: Map SchemaID points to a missing Schema ID")
    );
    assert!(native("excel-metadata.xls") == bytes);
}

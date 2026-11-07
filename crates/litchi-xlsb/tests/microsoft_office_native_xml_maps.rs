//! Inert XML-map derivatives of native Excel parents; no XML import or XPath evaluation.
use litchi_opc::OpcPackage;
use litchi_xlsb::package::PackageError as Error;
use litchi_xlsb::{
    Package,
    xml_maps::{
        self, CellReference, ColumnBinding, SingleCellBinding, XPath, XmlDataType, XmlMapLimits,
    },
};
use std::{collections::BTreeMap, path::Path};
const META: &str = "excel-metadata.xlsb";
const VALUES: &str = "excel-values-formulas-chart.xlsb";
const XML:&[u8]=br#"<MapInfo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" SelectionNamespaces="xmlns:e='urn:litchi:native-map'">
<Schema ID="schema-7" SchemaRef="https://example.invalid/native-map.xsd" Namespace="urn:litchi:native-map"><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="urn:litchi:native-map" elementFormDefault="qualified"><xs:element name="orders"><xs:complexType><xs:sequence><xs:element name="amount" type="xs:string"/><xs:element name="quantity" type="xs:string"/><xs:element name="total" type="xs:string"/></xs:sequence><xs:attribute name="amount" type="xs:string"/></xs:complexType></xs:element></xs:schema></Schema>
<Map ID="7" Name="NativeMap" RootElement="orders" SchemaID="schema-7" ShowImportExportValidationErrors="true" AutoFit="false" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"><DataBinding DataBindingName="inert" FileBinding="false" DataBindingLoadMode="1"/></Map></MapInfo>"#;
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn package(name: &str) -> Package {
    Package::from_bytes(&native(name)).unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let b = p.to_bytes().unwrap();
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        b,
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
type Relation = (String, String, String, String);
fn graph(p: &OpcPackage) -> BTreeMap<String, (String, Vec<Relation>)> {
    fn rels(r: &litchi_opc::Relationships) -> Vec<Relation> {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().to_string(),
                    r.reltype().to_string(),
                    r.target_ref().to_string(),
                    format!("{:?}", r.target_mode()),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    }
    let mut g = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for part in p.iter_parts() {
        g.insert(
            part.partname().as_str().into(),
            (part.content_type().into(), rels(part.rels())),
        );
    }
    g
}
fn equivalent(a: &Package, b: &Package) {
    assert_eq!(graph(a.opc_package()), graph(b.opc_package()));
    for part in a.opc_package().iter_parts() {
        assert!(
            a.opc_package().get_part(part.partname()).unwrap().blob()
                == b.opc_package().get_part(part.partname()).unwrap().blob(),
            "changed {}",
            part.partname()
        );
    }
}
fn catalog() -> xml_maps::XmlMapInfo {
    xml_maps::parse_xml_map_info(XML, &XmlMapLimits::DEFAULT).unwrap()
}

fn binding(id: u32, map: u32, path: &str) -> SingleCellBinding {
    SingleCellBinding::new(
        id,
        1,
        CellReference::new(1, 1).unwrap(),
        map,
        XmlDataType::new(1).unwrap(),
        XPath::new(path).unwrap(),
    )
    .unwrap()
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn assert_commit(p: &Package, c: &xml_maps::Commit, owner_rels: &str) -> Package {
    let r = save(c.package());
    assert!(!c.patch().is_noop());
    equivalent(&r, c.package());
    assert!(members(&p.apply_xml_maps(c.patch()).unwrap().into_package()) == members(c.package()));
    let restored = r
        .apply_xml_maps(&c.patch().inverse())
        .unwrap()
        .into_package();
    equivalent(&restored, p);
    let a = members(p);
    let b = members(&restored);
    let changed = a
        .keys()
        .filter(|name| a.get(*name) != b.get(*name))
        .cloned()
        .collect::<Vec<_>>();
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for name in changed {
        assert!(
            name == "[Content_Types].xml"
                || name == "xl/_rels/workbook.bin.rels"
                || name == owner_rels,
            "unexpected inverse change {name}"
        );
    }
    r
}
fn existing_payloads_unchanged(a: &Package, b: &Package, allowed: &[&str]) {
    let a = members(a);
    let b = members(b);
    for (n, v) in a {
        if !matches!(
            n.as_str(),
            "[Content_Types].xml" | "xl/_rels/workbook.bin.rels"
        ) && !allowed.contains(&n.as_str())
        {
            assert!(b.get(&n) == Some(&v), "changed {n}");
        }
    }
}

fn saved_print_value(p: &Package) {
    use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet};
    let workbook = p.workbook().unwrap();
    let sheet = workbook.worksheet_by_name("Print_015").unwrap();
    let cell = sheet.cell_by_coordinate("B2").unwrap();
    assert_eq!(cell.coordinate(), "B2");
    assert_eq!(cell.value(), &CellValue::Int(42));
}

#[test]
fn native_metadata_xml_map_lifecycle_keeps_every_worksheet_payload_inert() {
    let p = package(META);
    let original = native(META);
    saved_print_value(&p);
    let s = p.xml_maps().unwrap();
    assert!(s.map_info().is_none());
    assert!(s.mapped_tables().is_empty());
    assert!(s.single_cell_tables().is_empty());
    assert!(
        p.edit_xml_maps()
            .unwrap()
            .commit()
            .unwrap()
            .patch()
            .is_noop()
    );
    let mut e = p.edit_xml_maps().unwrap();
    e.set_catalog(catalog()).unwrap();
    e.put_single_cell_binding(4, binding(91, 7, "/e:orders/e:amount"))
        .unwrap();
    assert!(matches!(
        e.remove_catalog(),
        Err(Error::InvalidFormat(
            "cannot remove the XML Maps catalog while binary bindings depend on it"
        ))
    ));
    let c = e.commit().unwrap();
    let added = assert_commit(&p, &c, "xl/worksheets/_rels/sheet5.bin.rels");
    saved_print_value(&added);
    assert_eq!(
        added.xml_maps().unwrap().single_cell_bindings(4).unwrap(),
        &[binding(91, 7, "/e:orders/e:amount")]
    );
    existing_payloads_unchanged(&p, &added, &["xl/worksheets/_rels/sheet5.bin.rels"]);
    let mut e = added.edit_xml_maps().unwrap();
    e.set_map_name(7, "Reviewed 日本語").unwrap();
    e.put_single_cell_binding(4, binding(91, 7, "/e:orders/@amount"))
        .unwrap();
    let c = e.commit().unwrap();
    let updated = assert_commit(&added, &c, "xl/worksheets/_rels/sheet5.bin.rels");
    saved_print_value(&updated);
    let a = members(&added);
    let b = members(&updated);
    assert_eq!(
        a.keys()
            .filter(|n| a.get(*n) != b.get(*n))
            .map(String::as_str)
            .collect::<Vec<_>>(),
        ["xl/tables/tableSingleCells1.bin", "xl/xmlMaps.xml"]
    );
    assert_eq!(
        updated.xml_maps().unwrap().maps()[0].name,
        "Reviewed 日本語"
    );
    assert_eq!(
        updated.xml_maps().unwrap().single_cell_bindings(4).unwrap(),
        &[binding(91, 7, "/e:orders/@amount")]
    );
    let mut e = updated.edit_xml_maps().unwrap();
    e.remove_single_cell_binding(4, 91).unwrap();
    e.remove_catalog().unwrap();
    let c = e.commit().unwrap();
    let removed = assert_commit(&updated, &c, "xl/worksheets/_rels/sheet5.bin.rels");
    equivalent(&removed, &p);
    assert!(removed.xml_maps().unwrap().map_info().is_none());
    assert!(native(META) == original);
}

#[test]
fn native_range_table_refusal_and_single_bindings_keep_worksheet_order_after_a_chart() {
    let p = package(VALUES);
    let before = p.xml_maps().unwrap();
    assert!(before.map_info().is_none());
    assert!(before.mapped_tables().is_empty());
    let mut e = p.edit_xml_maps().unwrap();
    e.set_catalog(catalog()).unwrap();
    let column = ColumnBinding::new(
        2,
        7,
        XmlDataType::new(1).unwrap(),
        XPath::new("/e:orders/e:quantity").unwrap(),
        false,
    )
    .unwrap();
    let table = litchi_xlsb::package::table::parse_table_part(
        &members(&p)["xl/tables/table1.bin"],
        litchi_xlsb::raw::Limits::default(),
    )
    .unwrap();
    assert_eq!(table.id, 1);
    assert_eq!(table.table_type, litchi_xlsb::package::table::Type::Range);
    assert!(matches!(
        e.put_table_column_binding(1, column),
        Err(litchi_xlsb::package::PackageError::InvalidFormat(
            "table ID is not an ordinary LTXML Table"
        ))
    ));
    e.put_single_cell_binding(1, binding(92, 7, "/e:orders/e:total"))
        .unwrap();
    let c = e.commit().unwrap();
    let r = assert_commit(&p, &c, "xl/worksheets/_rels/sheet2.bin.rels");
    let s = r.xml_maps().unwrap();
    assert!(s.mapped_tables().is_empty());
    assert!(s.single_cell_bindings(0).is_none());
    assert_eq!(
        s.single_cell_bindings(1).unwrap(),
        &[binding(92, 7, "/e:orders/e:total")]
    );
    let m = members(&r);
    assert!(m.contains_key("xl/worksheets/_rels/sheet2.bin.rels"));
    existing_payloads_unchanged(&p, &r, &["xl/worksheets/_rels/sheet2.bin.rels"]);
    let mut e = r.edit_xml_maps().unwrap();

    e.remove_single_cell_binding(1, 92).unwrap();
    e.remove_catalog().unwrap();
    let c = e.commit().unwrap();
    let removed = assert_commit(&r, &c, "xl/worksheets/_rels/sheet2.bin.rels");
    equivalent(&removed, &p);
}

#[test]
fn native_xml_map_invalid_staging_and_stale_patches_do_not_change_the_source() {
    let p = package(META);
    let before = p.to_bytes().unwrap();
    let mut e = p.edit_xml_maps().unwrap();
    e.set_catalog(catalog()).unwrap();
    let staged = e.catalog().unwrap().unwrap().clone();
    let mut duplicate = staged.clone();
    duplicate.maps.push(duplicate.maps[0].clone());
    assert!(
        matches!(e.set_catalog(duplicate), Err(Error::Common(litchi_ooxml_common::Error::SpreadsheetXmlMaps(ref message))) if message == "duplicate Map ID")
    );
    assert_eq!(e.catalog().unwrap(), Some(&staged));
    assert!(matches!(
        e.put_single_cell_binding(5, binding(91, 7, "/e:orders/e:amount")),
        Err(Error::WorksheetNotFound(Some(5)))
    ));
    assert!(matches!(
        e.put_single_cell_binding(4, binding(91, 8, "/e:orders/e:amount")),
        Err(Error::InvalidFormat(_))
    ));
    assert_eq!(e.catalog().unwrap(), Some(&staged));
    e.put_single_cell_binding(4, binding(91, 7, "/e:orders/e:amount"))
        .unwrap();
    let c = e.commit().unwrap();
    let added = assert_commit(&p, &c, "xl/worksheets/_rels/sheet5.bin.rels");
    let mut e = added.edit_xml_maps().unwrap();
    e.set_map_name(7, "first").unwrap();
    let c = e.commit().unwrap();
    let mut other = added.edit_xml_maps().unwrap();
    other.set_map_name(7, "different").unwrap();
    let changed = other.commit().unwrap().into_package();
    assert!(matches!(
        changed.apply_xml_maps(c.patch()),
        Err(Error::PatchConflict { part: "XML Maps" })
    ));
    assert!(p.to_bytes().unwrap() == before);

    let limits = xml_maps::SemanticLimits {
        max_total_bindings: 1,
        ..xml_maps::SemanticLimits::DEFAULT
    };
    let mut limited = p
        .edit_xml_maps_with_options(xml_maps::EditOptions::default().limits(limits))
        .unwrap();
    limited.set_catalog(catalog()).unwrap();
    limited
        .put_single_cell_binding(4, binding(91, 7, "/e:orders/e:amount"))
        .unwrap();
    assert!(matches!(
        limited.put_single_cell_binding(0, binding(92, 7, "/e:orders/e:total")),
        Err(Error::InvalidFormat("aggregate XML binding count exceeds"))
    ));
    let c = limited.commit().unwrap();
    assert!(
        c.package()
            .xml_maps()
            .unwrap()
            .single_cell_bindings(0)
            .is_none()
    );
    assert_eq!(
        c.package()
            .xml_maps()
            .unwrap()
            .single_cell_bindings(4)
            .unwrap()
            .len(),
        1
    );
    assert!(p.to_bytes().unwrap() == before);
}

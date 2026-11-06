//! Native Excel binary external-data metadata and source-preserving changes.
use litchi_xlsb::Package;
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-connections.xlsb"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let value = reader.read_member(&name).unwrap();
            (name, value)
        })
        .collect()
}
#[test]
fn native_text_import_connection_exposes_the_stored_property_block() {
    let package = Package::from_bytes(&native()).unwrap();
    let owner = package.connections().unwrap();
    let connections = owner.connections().unwrap();
    assert_eq!(connections.connections.len(), 2);
    let text = connections.by_id(1).unwrap();
    assert_eq!(
        text.source_type,
        litchi_xlsb::package::connections::SourceType::Text
    );
    use litchi_xlsb::package::connections::{
        Properties, TextCodePage, TextFieldType, TextQualifier,
    };
    assert!(text.data_file.is_none() && text.connection_file.is_none());
    let Properties::Text(options) = &text.properties else {
        panic!("native text import properties")
    };
    assert_eq!(options.code_page, TextCodePage::WindowsAnsi);
    assert_eq!(options.application_code_page, 45);
    assert!(options.delimited && options.comma && options.prompt_for_file);
    assert!(
        !options.tab
            && !options.space
            && !options.semicolon
            && !options.consecutive
            && !options.custom_delimiter
    );
    assert_eq!(options.qualifier, TextQualifier::DoubleQuote);
    assert_eq!(
        (
            options.delimiter,
            options.first_row,
            options.decimal_separator,
            options.thousands_separator
        ),
        (0, 1, b'.', b',')
    );
    assert_eq!(options.ignored_flags, 1 << 20);
    assert_eq!(
        options.source_file,
        "/Users/ryker/CodeProjects/litchi/test-data/microsoft-office/assets/native-query-010.csv"
    );
    assert_eq!(options.fields.as_ref().unwrap().len(), 3);
    assert!(
        options
            .fields
            .as_ref()
            .unwrap()
            .iter()
            .all(|f| f.field_type == TextFieldType::General && f.start == 0)
    );
    assert_eq!(
        owner
            .unknown_records()
            .iter()
            .map(|r| (r.kind(), r.connection_id(), r.after_known()))
            .collect::<Vec<_>>(),
        [
            (37, None, 1),
            (3072, None, 1),
            (38, None, 1),
            (37, None, 2),
            (3072, None, 2),
            (38, None, 2)
        ]
    );
    let database = connections.by_id(2).unwrap();
    assert_eq!(database.name, "Query - native-query-010");
    let Properties::Database(properties) = &database.properties else {
        panic!("native database properties")
    };
    assert_eq!(
        properties.command.as_deref(),
        Some("SELECT * FROM [native-query-010]")
    );
    assert_eq!(
        properties.connection_string,
        "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=native-query-010;Extended Properties=\"\""
    );
}
fn outside_changed_header(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    litchi_xlsb::raw::Records::new(bytes, litchi_xlsb::raw::Limits::DEFAULT)
        .map(|r| r.unwrap())
        .filter(|r| {
            !(r.kind() == litchi_xlsb::raw::kind::BEGIN_EXT_CONNECTION
                && r.payload().get(18..22) == Some(2u32.to_le_bytes().as_slice()))
        })
        .map(|r| (r.kind().get(), r.payload().to_vec()))
        .collect()
}
#[test]
fn native_connection_description_edit_keeps_text_records_in_their_original_scope() {
    let bytes = native();
    let before = members(&bytes);
    let package = Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_connections().unwrap();
    edit.edit(2, |c| {
        c.description = Some("NATIVE_XLSB_DESCRIPTION_021".into());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("description.xlsb");
    commit.package().save(&path).unwrap();
    let saved = std::fs::read(path).unwrap();
    let after = members(&saved);
    let reopened = Package::from_bytes(&saved).unwrap();
    assert_eq!(
        reopened
            .connections()
            .unwrap()
            .connections()
            .unwrap()
            .by_id(2)
            .unwrap()
            .description
            .as_deref(),
        Some("NATIVE_XLSB_DESCRIPTION_021")
    );
    let restored = reopened
        .apply_connections(&commit.patch().inverse())
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    assert!(members(&restored) == before);
    assert!(package.to_bytes().unwrap() == bytes);
    assert!(
        outside_changed_header(&before["xl/connections.bin"])
            == outside_changed_header(&after["xl/connections.bin"]),
        "unchanged records moved or changed"
    );
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, value) in before {
        if name != "xl/connections.bin" {
            assert!(after[&name] == value, "changed {name}");
        }
    }
}

#[test]
fn native_text_options_and_fields_edit_without_changing_cached_import_rows() {
    use litchi_xlsb::package::connections::{Properties, TextFieldType};
    let bytes = native();
    let before = members(&bytes);
    let package = Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_connections().unwrap();
    edit.edit(1, |c| {
        let Properties::Text(t) = &mut c.properties else {
            panic!("text")
        };
        t.comma = false;
        t.semicolon = true;
        t.fields.as_mut().unwrap()[1].field_type = TextFieldType::Text;
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let saved = commit.package().to_bytes().unwrap();
    let after = members(&saved);
    let records = |bytes: &[u8]| {
        litchi_xlsb::raw::Records::new(bytes, litchi_xlsb::raw::Limits::DEFAULT)
            .map(|r| {
                let r = r.unwrap();
                (r.kind().get(), r.payload().to_vec())
            })
            .collect::<Vec<_>>()
    };
    let old = records(&before["xl/connections.bin"]);
    let new = records(&after["xl/connections.bin"]);
    assert_eq!(old.len(), 21);
    assert_eq!(new.len(), old.len());
    assert_eq!(
        old.iter()
            .zip(&new)
            .enumerate()
            .filter(|(_, (a, b))| a != b)
            .map(|(i, _)| i)
            .collect::<Vec<_>>(),
        [5, 8]
    );
    for (name, data) in &before {
        if name != "xl/connections.bin" {
            assert!(after[name] == *data, "changed {name}");
        }
    }
    let reopened = Package::from_bytes(&saved).unwrap();
    let snapshot = reopened.connections().unwrap();
    let Properties::Text(t) = &snapshot.connections().unwrap().by_id(1).unwrap().properties else {
        panic!("text")
    };
    assert!(!t.comma && t.semicolon);
    assert_eq!(
        t.fields.as_ref().unwrap()[1].field_type,
        TextFieldType::Text
    );
    assert!(
        members(
            &reopened
                .apply_connections(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == before
    );
}

#[test]
fn native_text_field_add_remove_is_reversible_and_root_identity_order_is_guarded() {
    use litchi_xlsb::package::connections::{Properties, TextField, TextFieldType};
    let bytes = native();
    let before = members(&bytes);
    let package = Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_connections().unwrap();
    edit.edit(1, |c| {
        let Properties::Text(t) = &mut c.properties else {
            panic!("text")
        };
        t.fields.as_mut().unwrap().push(TextField {
            field_type: TextFieldType::Skip,
            start: 0,
        });
        Ok(())
    })
    .unwrap();
    let added = edit.commit().unwrap();
    let saved = added.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&saved).unwrap();
    let snapshot = reopened.connections().unwrap();
    let Properties::Text(t) = &snapshot.connections().unwrap().by_id(1).unwrap().properties else {
        panic!("text")
    };
    assert_eq!(t.fields.as_ref().unwrap().len(), 4);
    let mut edit = reopened.edit_connections().unwrap();
    edit.edit(1, |c| {
        let Properties::Text(t) = &mut c.properties else {
            panic!("text")
        };
        t.fields.as_mut().unwrap().pop();
        Ok(())
    })
    .unwrap();
    let removed = edit.commit().unwrap();
    assert!(members(&removed.package().to_bytes().unwrap()) == before);
    assert!(
        members(
            &reopened
                .apply_connections(&added.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == before
    );
    let mut edit = package.edit_connections().unwrap();
    let mut changed = edit.connections().unwrap().unwrap().clone();
    changed.connections.reverse();
    edit.replace(Some(changed)).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(litchi_xlsb::package::error::Error::UnsupportedFeature(
            "connection layout changes around opaque records"
        ))
    ));
    let mut single = package.edit_connections().unwrap();
    single.remove(1).unwrap();
    assert!(matches!(
        single.commit(),
        Err(litchi_xlsb::package::error::Error::UnsupportedFeature(
            "removing connections with query-table dependencies"
        ))
    ));
    let mut removal = package.edit_connections().unwrap();
    removal.replace(None).unwrap();
    assert!(matches!(
        removal.commit(),
        Err(litchi_xlsb::package::error::Error::UnsupportedFeature(
            "removing connections with query-table dependencies"
        ))
    ));
    assert!(package.to_bytes().unwrap() == bytes);
}

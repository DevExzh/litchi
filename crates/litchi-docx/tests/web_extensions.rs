use std::collections::BTreeMap;
use std::io::Cursor;

use litchi_docx::{Package, web_extensions as web};
use litchi_opc::PackageWriter;
use litchi_opc::constants::relationship_type as rt;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

#[path = "web_extensions/durable.rs"]
mod durable;

fn custom_functions() -> web::CustomFunctions {
    let mut functions = web::CustomFunctions::new();
    functions
        .set_contains_custom_functions(Some(web::ContainsCustomFunctions::new(Some(true))))
        .set_background_app_data(Some(
            web::BackgroundAppData::new(7, "runtime-public").unwrap(),
        ));
    let mut ids = web::CustomFunctionList::new();
    ids.push_id("CONTOSO.ADDIN.FUNCTION")
        .unwrap()
        .push_id("CONTOSO.ADDIN.SECOND")
        .unwrap();
    functions.set_custom_function_list(Some(ids));
    functions
}

fn authored_panes() -> (web::Panes, web::CustomFunctions) {
    let metadata = custom_functions();
    let reference = web::Reference::new("public-addin", "1.0.0.0", web::Store::Omex).unwrap();
    let mut add_in = web::AddIn::new("public-instance", reference)
        .unwrap()
        .bind(web::Binding::new("binding", "matrix", "public-app").unwrap())
        .unwrap();
    add_in.set_custom_functions(Some(metadata.clone())).unwrap();

    let pane = web::Pane::new(add_in)
        .show(false)
        .dock(web::Dock::Left)
        .unwrap()
        .width(420.0)
        .unwrap();
    let mut panes = web::Panes::new();
    panes.push(pane).unwrap();
    (panes, metadata)
}

fn package_bytes(package: &Package) -> Vec<u8> {
    PackageWriter::to_bytes(package.opc_package()).unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}

fn replace_once(bytes: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    let offset = bytes
        .windows(from.len())
        .position(|window| window == from)
        .unwrap_or_else(|| panic!("missing fixture marker {from:?}"));
    let mut output = Vec::with_capacity(bytes.len() + to.len().saturating_sub(from.len()));
    output.extend_from_slice(&bytes[..offset]);
    output.extend_from_slice(to);
    output.extend_from_slice(&bytes[offset + from.len()..]);
    output
}

fn insert_before(bytes: &[u8], marker: &[u8], insertion: &[u8]) -> Vec<u8> {
    let offset = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .unwrap_or_else(|| panic!("missing fixture marker {marker:?}"));
    let mut output = Vec::with_capacity(bytes.len() + insertion.len());
    output.extend_from_slice(&bytes[..offset]);
    output.extend_from_slice(insertion);
    output.extend_from_slice(&bytes[offset..]);
    output
}

fn replace_byte(bytes: &[u8], from: u8, to: u8) -> Vec<u8> {
    bytes
        .iter()
        .map(|byte| if *byte == from { to } else { *byte })
        .collect()
}

fn rewrite_member<F>(bytes: &[u8], selected: &str, mut rewrite: F) -> Vec<u8>
where
    F: FnMut(&[u8]) -> Vec<u8>,
{
    let archive = ArchiveReader::new(bytes).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let content = archive.read(name).unwrap();
        let content = if name == selected {
            rewrite(&content)
        } else {
            content
        };
        writer.write_stored(name, &content).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn add_member(bytes: &[u8], added: &str, content: &[u8]) -> Vec<u8> {
    let archive = ArchiveReader::new(bytes).unwrap();
    assert!(!archive.file_names().any(|name| name == added));
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        writer
            .write_stored(name, &archive.read(name).unwrap())
            .unwrap();
    }
    writer.write_stored(added, content).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn loaded_custom_functions(package: &Package) -> web::CustomFunctions {
    package
        .task_panes()
        .unwrap()
        .unwrap()
        .get(0usize)
        .unwrap()
        .add_in()
        .custom_functions()
        .unwrap()
        .clone()
}

fn changed_custom_functions() -> web::CustomFunctions {
    let mut value = custom_functions();
    value.set_background_app_data(Some(
        web::BackgroundAppData::new(8, "runtime-changed").unwrap(),
    ));
    value
}

fn changed_custom_functions_with_id() -> web::CustomFunctions {
    let mut value = changed_custom_functions();
    let mut ids = web::CustomFunctionList::new();
    ids.push_id("CONTOSO.ADDIN.REWRITTEN").unwrap();
    ids.push_id("CONTOSO.ADDIN.SECOND").unwrap();
    value.set_custom_function_list(Some(ids));
    value
}

fn custom_functions_empty_list() -> web::CustomFunctions {
    let mut value = custom_functions();
    value.set_custom_function_list(Some(web::CustomFunctionList::new()));
    value
}

fn explicit_false_custom_functions() -> web::CustomFunctions {
    let mut value = custom_functions();
    value.set_contains_custom_functions(Some(web::ContainsCustomFunctions::new(Some(false))));
    value
}

fn panes_with_custom_functions(value: web::CustomFunctions) -> web::Panes {
    let reference = web::Reference::new("public-addin", "1.0.0.0", web::Store::Omex).unwrap();
    let mut add_in = web::AddIn::new("public-instance", reference)
        .unwrap()
        .bind(web::Binding::new("binding", "matrix", "public-app").unwrap())
        .unwrap();
    add_in.set_custom_functions(Some(value)).unwrap();
    let mut panes = web::Panes::new();
    panes.push(web::Pane::new(add_in)).unwrap();
    panes
}

fn two_authored_panes() -> web::Panes {
    let (mut panes, _) = authored_panes();
    let reference = web::Reference::new("second-addin", "2.0.0.0", web::Store::Omex).unwrap();
    let mut add_in = web::AddIn::new("second-instance", reference)
        .unwrap()
        .bind(web::Binding::new("second-binding", "table", "second-app").unwrap())
        .unwrap();
    add_in
        .set_custom_functions(Some(changed_custom_functions()))
        .unwrap();
    panes
        .push(web::Pane::new(add_in).dock(web::Dock::Right).unwrap())
        .unwrap();
    panes
}

fn extension_member_name<'a>(
    package_members: &'a BTreeMap<String, Vec<u8>>,
    instance_id: &str,
) -> &'a str {
    let marker = format!("id=\"{instance_id}\"");
    package_members
        .iter()
        .find_map(|(name, content)| {
            (name.starts_with("webextensions/webextension")
                && content
                    .windows(marker.len())
                    .any(|window| window == marker.as_bytes()))
            .then_some(name.as_str())
        })
        .unwrap_or_else(|| panic!("missing web extension instance {instance_id}"))
}

fn panes_with_metadata(package: &Package, value: web::CustomFunctions) -> web::Panes {
    let mut panes = package.task_panes().unwrap().unwrap();
    assert!(
        panes
            .edit("public-instance", |pane| {
                pane.add_in_mut()
                    .set_custom_functions(Some(value.clone()))?;
                Ok(())
            })
            .unwrap()
    );
    panes
}

fn noncanonical_web_member(content: &[u8], root: &[u8], marker: &[u8]) -> Vec<u8> {
    let quoted = replace_byte(content, b'"', b'\'');
    insert_before(&quoted, root, marker)
}

fn changed_panes(package: &Package) -> web::Panes {
    let mut panes = package.task_panes().unwrap().unwrap();
    assert!(
        panes
            .edit("public-instance", |pane| {
                pane.set_visible(true);
                pane.set_row(3);
                Ok(())
            })
            .unwrap()
    );
    panes
}

fn assert_unchanged_after_error(package: &mut Package, operation: impl FnOnce(&mut Package)) {
    let before = members(&package_bytes(package));
    operation(package);
    assert_eq!(members(&package_bytes(package)), before);
}

#[test]
fn publicly_constructed_custom_function_payloads_round_trip_in_both_conformances() {
    for conformance in [web::Conformance::Transitional, web::Conformance::Strict] {
        let (panes, metadata) = authored_panes();
        let mut package = Package::new().unwrap();
        package.put_task_panes(panes.clone(), conformance).unwrap();

        assert_eq!(loaded_custom_functions(&package), metadata);
        let first = package_bytes(&package);

        let mut second = Package::new().unwrap();
        second.put_task_panes(panes, conformance).unwrap();
        assert_eq!(members(&first), members(&package_bytes(&second)));

        let first_members = members(&first);
        let extension_xml = first_members
            .get("webextensions/webextension1.xml")
            .unwrap();
        assert!(
            extension_xml
                .windows(b"containsCustomFunctions".len())
                .any(|window| { window == b"containsCustomFunctions" })
        );
        assert!(
            extension_xml
                .windows(b"backgroundAppData".len())
                .any(|window| window == b"backgroundAppData")
        );
        assert!(
            extension_xml
                .windows(b"customFunctionList".len())
                .any(|window| window == b"customFunctionList")
        );
    }
}

#[test]
fn public_task_pane_plan_is_nonmutating_and_inverse_restores_saved_members_exactly() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    let original = members(&package_bytes(&package));
    let patch = package
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    assert!(!patch.is_empty());
    assert_eq!(members(&package_bytes(&package)), original);

    assert!(package.apply_task_panes_patch(&patch).unwrap());
    let published = package_bytes(&package);
    let published_members = members(&published);
    for (name, content) in &original {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert_eq!(published_members.get(name), Some(content), "member {name}");
        }
    }

    let mut reopened = Package::from_reader(Cursor::new(published)).unwrap();
    assert_eq!(loaded_custom_functions(&reopened), custom_functions());
    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&reopened)), original);
}

#[test]
fn changed_and_exact_noop_public_patches_have_expected_effects() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let before_noop = members(&package_bytes(&package));
    let loaded = package.task_panes().unwrap().unwrap();
    let noop = package
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(noop.is_empty());
    assert!(!package.apply_task_panes_patch(&noop).unwrap());
    assert_eq!(members(&package_bytes(&package)), before_noop);

    let changed = changed_panes(&package);
    let patch = package
        .plan_task_panes(changed.clone(), web::Conformance::Strict)
        .unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    let loaded = package.task_panes().unwrap().unwrap();
    assert!(loaded.get(0usize).unwrap().visible());
    assert_eq!(loaded.get(0usize).unwrap().row(), 3);
    assert_eq!(loaded.get(0usize).unwrap().add_in().id(), "public-instance");
    assert_eq!(package.task_panes().unwrap(), Some(changed));
}

#[test]
fn stale_content_types_and_relationship_lexical_members_refuse_atomically() {
    let (panes, _) = authored_panes();
    let mut source = Package::new().unwrap();
    source
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let patch = source.plan_remove_task_panes().unwrap();
    let original = package_bytes(&source);

    let content_types_variant = rewrite_member(&original, "[Content_Types].xml", |content| {
        replace_once(
            content,
            b"</Types>",
            b"<!-- stale content-types lexical token -->\n</Types>",
        )
    });
    let relationship_variant = rewrite_member(&original, "_rels/.rels", |content| {
        replace_once(
            content,
            b"</Relationships>",
            b"<!-- stale rel lexical token -->\n</Relationships>",
        )
    });

    for (name, expected, variant) in [
        (
            "content types",
            "source content-types changed",
            content_types_variant,
        ),
        (
            "root relationships",
            "source relationship metadata changed",
            relationship_variant,
        ),
    ] {
        let mut stale = Package::from_reader(Cursor::new(variant)).unwrap();
        let before = members(&package_bytes(&stale));
        let error = stale.apply_task_panes_patch(&patch).unwrap_err();
        assert!(
            error.to_string().contains(expected),
            "{name}: unexpected stale error: {error}"
        );
        assert_eq!(members(&package_bytes(&stale)), before);
    }
}

#[test]
fn signed_noop_is_preserved_but_changed_publication_requires_explicit_unsign() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    package
        .edit_opc(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .unwrap();
    assert!(package.is_signed());

    let before = members(&package_bytes(&package));
    let loaded = package.task_panes().unwrap().unwrap();
    let noop = package
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(noop.is_empty());
    assert!(!package.apply_task_panes_patch(&noop).unwrap());
    assert!(package.is_signed());
    assert_eq!(members(&package_bytes(&package)), before);

    let signed_patch = package
        .plan_task_panes(changed_panes(&package), web::Conformance::Transitional)
        .unwrap();
    let error = package.apply_task_panes_patch(&signed_patch).unwrap_err();
    assert!(error.to_string().contains("explicit Package::unsign"));
    assert!(package.is_signed());
    assert_eq!(members(&package_bytes(&package)), before);

    package.unsign();
    let after_unsign = members(&package_bytes(&package));
    let error = package.apply_task_panes_patch(&signed_patch).unwrap_err();
    assert!(
        error.to_string().to_ascii_lowercase().contains("source"),
        "signed patch did not become stale after unsign: {error}"
    );
    assert_eq!(members(&package_bytes(&package)), after_unsign);

    let patch = package
        .plan_task_panes(changed_panes(&package), web::Conformance::Transitional)
        .unwrap();
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    assert!(!package.is_signed());
}

#[test]
fn dirty_document_facade_refuses_task_pane_planning_without_publication() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("unmaterialized document edit");
    let before = members(&package_bytes(&package));
    let error = package
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap_err();
    assert!(error.to_string().contains("unmaterialized changes"));
    assert_eq!(members(&package_bytes(&package)), before);
}

#[test]
fn public_limits_and_malformed_namespace_or_attribute_fixtures_refuse_atomically() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    let mut limits = web::Limits::standard();
    limits.items = 0;
    let plan_error = package
        .plan_task_panes_with_limits(panes.clone(), web::Conformance::Transitional, &limits)
        .expect_err("item limit should reject task-pane planning");
    assert!(plan_error.to_string().contains("task pane"));
    assert_unchanged_after_error(&mut package, |package| {
        let error = package
            .put_task_panes_with_limits(panes.clone(), web::Conformance::Transitional, &limits)
            .err()
            .expect("item limit should reject task-pane publication");
        assert!(error.to_string().contains("task pane"));
    });

    let mut valid = Package::new().unwrap();
    valid
        .put_task_panes(panes.clone(), web::Conformance::Transitional)
        .unwrap();
    let valid_bytes = package_bytes(&valid);
    let mut read_limits = web::Limits::standard();
    read_limits.xml_bytes = 1;
    let limited = Package::from_reader(Cursor::new(valid_bytes.clone())).unwrap();
    let error = limited.task_panes_with_limits(&read_limits).unwrap_err();
    assert!(error.to_string().contains("XML") || error.to_string().contains("xml"));

    let mut remove_limits = web::Limits::standard();
    remove_limits.part_deletions = 0;
    let plan_remove_limited = Package::from_reader(Cursor::new(valid_bytes.clone())).unwrap();
    let before_plan_remove = members(&package_bytes(&plan_remove_limited));
    let error = plan_remove_limited
        .plan_remove_task_panes_with_limits(&remove_limits)
        .expect_err("part-deletion limit should reject task-pane removal planning");
    assert!(error.to_string().contains("deletion"));
    assert_eq!(
        members(&package_bytes(&plan_remove_limited)),
        before_plan_remove
    );

    let mut remove_limited = Package::from_reader(Cursor::new(valid_bytes.clone())).unwrap();
    let before_remove = members(&package_bytes(&remove_limited));
    let error = remove_limited
        .remove_task_panes_with_limits(&remove_limits)
        .expect_err("part-deletion limit should reject task-pane removal");
    assert!(error.to_string().contains("deletion"));
    assert_eq!(members(&package_bytes(&remove_limited)), before_remove);

    let malformed_attribute =
        rewrite_member(&valid_bytes, "webextensions/webextension1.xml", |content| {
            replace_once(
                content,
                b"<we:containsCustomFunctions",
                b"<we:containsCustomFunctions unexpected=\"1\"",
            )
        });
    let malformed_namespace =
        rewrite_member(&valid_bytes, "webextensions/webextension1.xml", |content| {
            replace_once(
                content,
                b"http://schemas.microsoft.com/office/webextensions/webextension/2010/11",
                b"urn:foreign-web-extension",
            )
        });
    for fixture in [malformed_attribute, malformed_namespace] {
        let malformed = Package::from_reader(Cursor::new(fixture)).unwrap();
        let error = malformed.task_panes().unwrap_err();
        assert!(error.to_string().contains("web") || error.to_string().contains("XML"));
    }
}

#[test]
fn public_remove_and_inverse_restore_task_panes_after_save_and_reopen() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Strict)
        .unwrap();
    let original = members(&package_bytes(&package));
    let patch = package.plan_remove_task_panes().unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    assert_eq!(package.task_panes().unwrap(), None);

    let published = package_bytes(&package);
    let mut reopened = Package::from_reader(Cursor::new(published)).unwrap();
    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
    assert!(reopened.task_panes().unwrap().is_some());
    assert_eq!(members(&package_bytes(&reopened)), original);

    let (direct_panes, _) = authored_panes();
    let mut direct = Package::new().unwrap();
    direct
        .put_task_panes(direct_panes, web::Conformance::Transitional)
        .unwrap();
    assert!(direct.remove_task_panes().unwrap());
    assert_eq!(direct.task_panes().unwrap(), None);
}

#[test]
fn legacy_taskpane_child_name_is_refused_without_claiming_broad_schema_compatibility() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let legacy = rewrite_member(&source, "webextensions/taskpanes.xml", |content| {
        replace_once(content, b"<wetp:webextensionref", b"<wetp:webextension")
    });
    let error = match Package::from_reader(Cursor::new(legacy)) {
        Ok(package) => package.task_panes().unwrap_err(),
        Err(error) => error,
    };
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("webextensionref")
            || message.contains("taskpane")
            || message.contains("unexpected"),
        "legacy task-pane child refusal was not identifiable: {error}"
    );
}

#[test]
fn custom_function_rewrite_targets_later_payload_extension_after_vendor_extension() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let vendor_extension = br#"<a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:litchi:test:web-extension" uri="urn:litchi:vendor"><v:opaque/></a:ext>"#;
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        insert_before(content, b"<a:ext", vendor_extension)
    });

    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    assert_eq!(loaded_custom_functions(&package), custom_functions());
    let desired = panes_with_metadata(&package, changed_custom_functions());
    let patch = package
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());

    let extension = members(&package_bytes(&package));
    let extension = extension.get("webextensions/webextension1.xml").unwrap();
    let vendor_position = extension
        .windows(b"uri=\"urn:litchi:vendor\"".len())
        .position(|window| window == b"uri=\"urn:litchi:vendor\"")
        .expect("vendor extension should remain in the source order");
    let payload_position = extension
        .windows(b"backgroundAppData".len())
        .position(|window| window == b"backgroundAppData")
        .expect("typed payload should remain present");
    assert!(vendor_position < payload_position);
    assert!(
        extension
            .windows(b"<v:opaque/>".len())
            .any(|window| window == b"<v:opaque/>")
    );
    assert_eq!(
        loaded_custom_functions(&package),
        changed_custom_functions()
    );
}

#[test]
fn custom_function_semantic_noop_preserves_noncanonical_addin_taskpane_xml_and_signature() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    authored
        .edit_opc(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .unwrap();
    let source = package_bytes(&authored);
    let noncanonical = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        noncanonical_web_member(
            content,
            b"<we:webextension",
            b"<!-- noncanonical add-in --><?keep-addin value?>",
        )
    });
    let noncanonical = rewrite_member(&noncanonical, "webextensions/taskpanes.xml", |content| {
        noncanonical_web_member(
            content,
            b"<wetp:taskpanes",
            b"<!-- noncanonical task panes --><?keep-panes value?>",
        )
    });

    let mut package = Package::from_reader(Cursor::new(noncanonical)).unwrap();
    let before = members(&package_bytes(&package));
    let loaded = package.task_panes().unwrap().unwrap();
    let patch = package
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(patch.is_empty());
    assert!(!package.apply_task_panes_patch(&patch).unwrap());
    assert!(package.is_signed());
    assert_eq!(members(&package_bytes(&package)), before);
}

#[test]
fn omitted_contains_boolean_is_distinct_from_explicit_false_and_inverse_restores_source() {
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(
            panes_with_custom_functions(explicit_false_custom_functions()),
            web::Conformance::Transitional,
        )
        .unwrap();
    let source = package_bytes(&authored);
    let omitted = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        replace_once(content, b" val=\"false\"", b"")
    });
    let original = members(&omitted);
    let mut package = Package::from_reader(Cursor::new(omitted)).unwrap();
    let loaded = loaded_custom_functions(&package);
    assert_eq!(
        loaded.contains_custom_functions().unwrap().explicit_value(),
        None
    );
    assert!(!loaded.contains_custom_functions().unwrap().value());

    let patch = package
        .plan_task_panes(
            panes_with_metadata(&package, explicit_false_custom_functions()),
            web::Conformance::Transitional,
        )
        .unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    let published = members(&package_bytes(&package));
    assert!(
        published
            .get("webextensions/webextension1.xml")
            .unwrap()
            .windows(b"val=\"false\"".len())
            .any(|window| window == b"val=\"false\"")
    );

    let mut reopened = Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
    assert_eq!(
        loaded_custom_functions(&reopened)
            .contains_custom_functions()
            .unwrap()
            .explicit_value(),
        Some(false)
    );
    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&reopened)), original);
}

#[test]
fn known_custom_function_payload_comments_and_processing_instructions_are_retained_or_refused_precisely()
 {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        insert_before(
            content,
            b"<we:backgroundAppData",
            b"<!-- known-payload comment --><?known-payload keep?>",
        )
    });
    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    let before = members(&package_bytes(&package));
    let desired = panes_with_metadata(&package, changed_custom_functions());
    let planned = package.plan_task_panes(desired, web::Conformance::Transitional);
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("custom-function") || message.contains("custom function"),
                "known-payload refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => {
            assert!(!patch.is_empty());
            match package.apply_task_panes_patch(&patch) {
                Err(error) => {
                    let message = error.to_string().to_ascii_lowercase();
                    assert!(
                        message.contains("custom-function") || message.contains("custom function"),
                        "known-payload refusal was not precise: {error}"
                    );
                    assert_eq!(members(&package_bytes(&package)), before);
                },
                Ok(changed) => {
                    assert!(changed);
                    let extension = members(&package_bytes(&package));
                    let extension = extension.get("webextensions/webextension1.xml").unwrap();
                    assert!(
                        extension
                            .windows(b"<!-- known-payload comment -->".len())
                            .any(|window| window == b"<!-- known-payload comment -->")
                    );
                    assert!(
                        extension
                            .windows(b"<?known-payload keep?>".len())
                            .any(|window| window == b"<?known-payload keep?>")
                    );
                },
            }
        },
    }
}

#[test]
fn internal_custom_payload_markup_survives_changed_attributes_or_refuses_atomically() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let background = b"<we:backgroundAppData xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" state=\"7\" runtimeId=\"runtime-public\"/>";
    let expanded_background = b"<we:backgroundAppData xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" state=\"7\" runtimeId=\"runtime-public\"><!--inside-background--><?inside-background keep?></we:backgroundAppData>";
    let first_id = b"<we:customFunctionIds>CONTOSO.ADDIN.FUNCTION</we:customFunctionIds>";
    let marked_id = b"<we:customFunctionIds><!--inside-id--><?inside-id keep?>CONTOSO.ADDIN.FUNCTION</we:customFunctionIds>";
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        let content = replace_once(content, background, expanded_background);
        replace_once(&content, first_id, marked_id)
    });
    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    let before = members(&package_bytes(&package));
    let desired = panes_with_metadata(&package, changed_custom_functions_with_id());
    let planned = package.plan_task_panes(desired, web::Conformance::Transitional);
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("custom-function") || message.contains("custom function"),
                "internal-payload refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => {
            assert!(!patch.is_empty());
            match package.apply_task_panes_patch(&patch) {
                Err(error) => {
                    let message = error.to_string().to_ascii_lowercase();
                    assert!(
                        message.contains("custom-function") || message.contains("custom function"),
                        "internal-payload refusal was not precise: {error}"
                    );
                    assert_eq!(members(&package_bytes(&package)), before);
                },
                Ok(changed) => {
                    assert!(changed);
                    let published = members(&package_bytes(&package));
                    let extension = published.get("webextensions/webextension1.xml").unwrap();
                    for marker in [
                        b"<!--inside-background-->".as_slice(),
                        b"<?inside-background keep?>".as_slice(),
                        b"<!--inside-id-->".as_slice(),
                        b"<?inside-id keep?>".as_slice(),
                    ] {
                        assert!(
                            extension
                                .windows(marker.len())
                                .any(|window| window == marker),
                            "selected payload edit dropped {marker:?}"
                        );
                    }
                    assert!(
                        extension
                            .windows(b"CONTOSO.ADDIN.REWRITTEN".len())
                            .any(|window| window == b"CONTOSO.ADDIN.REWRITTEN")
                    );

                    let mut reopened =
                        Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
                    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
                    assert_eq!(members(&package_bytes(&reopened)), before);
                },
            }
        },
    }
}

#[test]
fn empty_custom_function_list_grows_without_dropping_internal_markup_or_siblings() {
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(
            panes_with_custom_functions(custom_functions_empty_list()),
            web::Conformance::Transitional,
        )
        .unwrap();
    let source = package_bytes(&authored);
    let empty_list = b"<we:customFunctionList xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\"/>";
    let marked_empty_list = b"<we:customFunctionList xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\"><!--empty-list--><?empty-list keep?></we:customFunctionList>";
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        replace_once(content, empty_list, marked_empty_list)
    });
    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    let before = members(&package_bytes(&package));
    let desired = panes_with_metadata(&package, custom_functions());
    let planned = package.plan_task_panes(desired, web::Conformance::Transitional);
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("custom-function") || message.contains("custom function"),
                "empty-list refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => {
            assert!(!patch.is_empty());
            assert!(package.apply_task_panes_patch(&patch).unwrap());
            let published = members(&package_bytes(&package));
            let extension = published.get("webextensions/webextension1.xml").unwrap();
            for marker in [
                b"<!--empty-list-->".as_slice(),
                b"<?empty-list keep?>".as_slice(),
            ] {
                assert!(
                    extension
                        .windows(marker.len())
                        .any(|window| window == marker),
                    "empty-list edit dropped {marker:?}"
                );
            }
            assert!(
                extension
                    .windows(b"CONTOSO.ADDIN.FUNCTION".len())
                    .any(|window| window == b"CONTOSO.ADDIN.FUNCTION")
            );

            let mut reopened = Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
            assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
            assert_eq!(members(&package_bytes(&reopened)), before);
        },
    }
}

#[test]
fn changed_custom_metadata_preserves_noncanonical_owner_bytes_and_inverse_after_reopen() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        let content = noncanonical_web_member(
            content,
            b"<we:webextension",
            b"<!-- owner-comment --><?owner-keep value?>",
        );
        insert_before(
            &content,
            b"</we:extLst>",
            b"<a:ext xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:v='urn:litchi:owner' uri='urn:litchi:owner'><v:opaque keep='yes'/></a:ext>",
        )
    });
    let mut package = Package::from_reader(Cursor::new(variant.clone())).unwrap();
    let before = members(&package_bytes(&package));
    let before_extension = before
        .get("webextensions/webextension1.xml")
        .unwrap()
        .clone();
    let desired = panes_with_metadata(&package, changed_custom_functions());
    let patch = package
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    let after = members(&package_bytes(&package));
    for (name, content) in &before {
        if name != "webextensions/webextension1.xml" {
            assert_eq!(after.get(name), Some(content), "unselected member {name}");
        }
    }
    let expected_extension = replace_once(
        &replace_once(&before_extension, b"state='7'", b"state='8'"),
        b"runtimeId='runtime-public'",
        b"runtimeId='runtime-changed'",
    );
    assert_eq!(
        after.get("webextensions/webextension1.xml"),
        Some(&expected_extension)
    );

    let second_desired = panes_with_metadata(&package, changed_custom_functions_with_id());
    let second_patch = package
        .plan_task_panes(second_desired, web::Conformance::Transitional)
        .unwrap();
    assert!(!second_patch.is_empty());
    assert!(package.apply_task_panes_patch(&second_patch).unwrap());
    assert!(
        package
            .apply_task_panes_patch(&second_patch.inverse())
            .unwrap()
    );
    assert!(package.apply_task_panes_patch(&patch.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&package)), before);

    let mut reopened = Package::from_reader(Cursor::new(variant)).unwrap();
    assert!(reopened.apply_task_panes_patch(&patch).unwrap());
    let mut reopened_after = Package::from_reader(Cursor::new(package_bytes(&reopened))).unwrap();
    assert!(
        reopened_after
            .apply_task_panes_patch(&patch.inverse())
            .unwrap()
    );
    assert_eq!(members(&package_bytes(&reopened_after)), before);
}

#[test]
fn two_pane_custom_metadata_edit_is_local_to_selected_extension() {
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(two_authored_panes(), web::Conformance::Transitional)
        .unwrap();
    let before = members(&package_bytes(&package));
    let first_name = extension_member_name(&before, "public-instance").to_owned();
    let second_name = extension_member_name(&before, "second-instance").to_owned();
    let second_before = before.get(&second_name).unwrap().clone();
    let desired = panes_with_metadata(&package, changed_custom_functions());
    let patch = package
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    assert!(!patch.is_empty());
    assert!(package.apply_task_panes_patch(&patch).unwrap());
    let after = members(&package_bytes(&package));
    assert_eq!(after.get(&second_name), Some(&second_before));
    for (name, content) in &before {
        if name != &first_name {
            assert_eq!(after.get(name), Some(content), "unselected member {name}");
        }
    }
    assert!(
        after
            .get("webextensions/taskpanes.xml")
            .is_some_and(|content| {
                content
                    .windows(b"rIdAddIn2".len())
                    .any(|window| window == b"rIdAddIn2")
            })
    );

    let mut reopened = Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&reopened)), before);
}

#[test]
fn noncustom_two_pane_edit_preserves_opaque_unselected_addin_or_refuses_atomically() {
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(two_authored_panes(), web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let source_members = members(&source);
    let second_name = extension_member_name(&source_members, "second-instance").to_owned();
    let variant = rewrite_member(&source, &second_name, |content| {
        insert_before(
            content,
            b"</we:extLst>",
            b"<!-- untouched second add-in --><?keep-second-addin value?>",
        )
    });
    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    let before = members(&package_bytes(&package));
    let desired = changed_panes(&package);
    let planned = package.plan_task_panes(desired, web::Conformance::Transitional);
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("source")
                    || message.contains("opaque")
                    || message.contains("markup")
                    || message.contains("extension"),
                "unselected opaque add-in refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => match package.apply_task_panes_patch(&patch) {
            Err(error) => {
                let message = error.to_string().to_ascii_lowercase();
                assert!(
                    message.contains("source")
                        || message.contains("opaque")
                        || message.contains("markup")
                        || message.contains("extension"),
                    "unselected opaque add-in refusal was not precise: {error}"
                );
                assert_eq!(members(&package_bytes(&package)), before);
            },
            Ok(changed) => {
                assert!(changed);
                let after = members(&package_bytes(&package));
                assert_eq!(
                    after.get(&second_name),
                    Some(before.get(&second_name).unwrap())
                );
            },
        },
    }
}

#[test]
fn noncustom_edit_with_taskpane_opaque_marker_preserves_it_or_refuses_atomically() {
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(two_authored_panes(), web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let variant = rewrite_member(&source, "webextensions/taskpanes.xml", |content| {
        insert_before(
            content,
            b"</wetp:taskpanes>",
            b"<vendor:unknown xmlns:vendor=\"urn:litchi:unknown-task-pane\" marker=\"keep\"/>",
        )
    });
    let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
    let before = members(&package_bytes(&package));
    if let Err(error) = package.task_panes() {
        let message = error.to_string().to_ascii_lowercase();
        assert!(
            message.contains("task-pane")
                || message.contains("task pane")
                || message.contains("unknown")
                || message.contains("unexpected")
                || message.contains("xml"),
            "task-pane unknown-marker refusal was not precise: {error}"
        );
        assert_eq!(members(&package_bytes(&package)), before);
        return;
    }
    let desired = changed_panes(&package);
    let planned = package.plan_task_panes(desired, web::Conformance::Transitional);
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("source")
                    || message.contains("opaque")
                    || message.contains("markup")
                    || message.contains("task-pane")
                    || message.contains("task pane"),
                "task-pane opaque refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => match package.apply_task_panes_patch(&patch) {
            Err(error) => {
                let message = error.to_string().to_ascii_lowercase();
                assert!(
                    message.contains("source")
                        || message.contains("opaque")
                        || message.contains("markup")
                        || message.contains("task-pane")
                        || message.contains("task pane"),
                    "task-pane opaque refusal was not precise: {error}"
                );
                assert_eq!(members(&package_bytes(&package)), before);
            },
            Ok(changed) => {
                assert!(changed);
                let after = members(&package_bytes(&package));
                let task_panes = after.get("webextensions/taskpanes.xml").unwrap();
                let marker = b"<vendor:unknown xmlns:vendor=\"urn:litchi:unknown-task-pane\" marker=\"keep\"/>";
                assert!(
                    task_panes
                        .windows(marker.len())
                        .any(|window| window == marker),
                    "successful edit dropped task-pane unknown marker"
                );
            },
        },
    }
}

#[test]
fn conformance_noop_ignores_missing_unused_prefix_and_vendor_other_dialect_text() {
    for (conformance, own_relationships, other_relationships) in [
        (
            web::Conformance::Transitional,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        ),
        (
            web::Conformance::Strict,
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
    ] {
        let (panes, _) = authored_panes();
        let mut authored = Package::new().unwrap();
        authored.put_task_panes(panes, conformance).unwrap();
        let source = package_bytes(&authored);
        let missing_unused_relationship_prefix =
            rewrite_member(&source, "webextensions/webextension1.xml", |content| {
                let marker = format!(" xmlns:r=\"{own_relationships}\"");
                replace_once(content, marker.as_bytes(), b"")
            });
        let vendor_other_dialect = rewrite_member(
            &source,
            "webextensions/webextension1.xml",
            |content| {
                let marker = format!(
                    "<a:ext xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:unusedOther=\"{other_relationships}\" uri=\"urn:litchi:vendor\"><v:opaque xmlns:v=\"urn:litchi:vendor\">{other_relationships}</v:opaque></a:ext>"
                );
                insert_before(content, b"</we:extLst>", marker.as_bytes())
            },
        );

        for (name, variant) in [
            (
                "missing unused r namespace",
                missing_unused_relationship_prefix,
            ),
            ("vendor other-dialect text", vendor_other_dialect),
        ] {
            let mut package = Package::from_reader(Cursor::new(variant)).unwrap();
            let before = members(&package_bytes(&package));
            let loaded = package.task_panes().unwrap().unwrap();
            let patch = package.plan_task_panes(loaded, conformance).unwrap();
            assert!(patch.is_empty(), "{name} incorrectly became a changed plan");
            assert!(!package.apply_task_panes_patch(&patch).unwrap());
            assert_eq!(members(&package_bytes(&package)), before, "{name}");
        }
    }
}

#[test]
fn inactive_mce_branch_is_retained_or_rejected_precisely_during_custom_edit() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let alternate = br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:litchi:unsupported" xmlns:v="urn:litchi:mce"><mc:Choice Requires="u"><v:inactive marker="choice"/></mc:Choice><mc:Fallback><v:inactive marker="fallback"/></mc:Fallback></mc:AlternateContent>"#;
    let variant = rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        insert_before(content, b"</a:ext>", alternate)
    });
    let mut package = match Package::from_reader(Cursor::new(variant)) {
        Ok(package) => package,
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("mce")
                    || message.contains("alternate")
                    || message.contains("markup-compatibility")
                    || message.contains("markup compatibility"),
                "MCE refusal was not precise: {error}"
            );
            return;
        },
    };
    let before = members(&package_bytes(&package));
    match package.task_panes() {
        Ok(Some(_)) => {},
        Ok(None) => panic!("MCE fixture lost task panes"),
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("mce")
                    || message.contains("alternate")
                    || message.contains("markup-compatibility")
                    || message.contains("markup compatibility"),
                "MCE refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
            return;
        },
    };
    let planned = package.plan_task_panes(
        panes_with_metadata(&package, changed_custom_functions()),
        web::Conformance::Transitional,
    );
    match planned {
        Err(error) => {
            let message = error.to_string().to_ascii_lowercase();
            assert!(
                message.contains("mce")
                    || message.contains("alternate")
                    || message.contains("markup-compatibility")
                    || message.contains("markup compatibility"),
                "MCE refusal was not precise: {error}"
            );
            assert_eq!(members(&package_bytes(&package)), before);
        },
        Ok(patch) => {
            assert!(!patch.is_empty());
            match package.apply_task_panes_patch(&patch) {
                Err(error) => {
                    let message = error.to_string().to_ascii_lowercase();
                    assert!(
                        message.contains("mce")
                            || message.contains("alternate")
                            || message.contains("markup-compatibility")
                            || message.contains("markup compatibility"),
                        "MCE refusal was not precise: {error}"
                    );
                    assert_eq!(members(&package_bytes(&package)), before);
                },
                Ok(changed) => {
                    assert!(changed);
                    let output = members(&package_bytes(&package));
                    let extension = output.get("webextensions/webextension1.xml").unwrap();
                    assert!(
                        extension
                            .windows(alternate.len())
                            .any(|window| window == alternate)
                    );
                    assert!(
                        extension
                            .windows(b"marker=\"choice\"".len())
                            .any(|window| { window == b"marker=\"choice\"" })
                    );
                    assert!(
                        extension
                            .windows(b"marker=\"fallback\"".len())
                            .any(|window| { window == b"marker=\"fallback\"" })
                    );
                    let mut reopened =
                        Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
                    assert!(reopened.apply_task_panes_patch(&patch.inverse()).unwrap());
                    assert_eq!(members(&package_bytes(&reopened)), before);
                },
            }
        },
    }
}

#[test]
fn semantic_noop_patch_rejects_semantically_stale_source_atomically() {
    let (panes, _) = authored_panes();
    let mut source = Package::new().unwrap();
    source
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let loaded = source.task_panes().unwrap().unwrap();
    let noop = source
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(noop.is_empty());
    let original = package_bytes(&source);
    let stale_bytes = rewrite_member(&original, "webextensions/webextension1.xml", |content| {
        replace_once(content, b"runtime-public", b"runtime-stale")
    });
    let root_relationship_bytes = rewrite_member(&original, "_rels/.rels", |content| {
        replace_once(
            content,
            b"</Relationships>",
            b"<!-- stale root relationship token --></Relationships>",
        )
    });
    let pane_relationship_bytes = rewrite_member(
        &original,
        "webextensions/_rels/taskpanes.xml.rels",
        |content| {
            replace_once(
                content,
                b"</Relationships>",
                b"<!-- stale task-pane relationship token --></Relationships>",
            )
        },
    );
    let empty_add_in_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
    let add_in_relationship_bytes = add_member(
        &original,
        "webextensions/_rels/webextension1.xml.rels",
        empty_add_in_relationships,
    );
    let add_in_relationship_source =
        Package::from_reader(Cursor::new(add_in_relationship_bytes.clone())).unwrap();
    let add_in_noop = add_in_relationship_source
        .plan_task_panes(
            add_in_relationship_source.task_panes().unwrap().unwrap(),
            web::Conformance::Transitional,
        )
        .unwrap();
    assert!(add_in_noop.is_empty());
    let add_in_relationship_lexical_bytes = rewrite_member(
        &add_in_relationship_bytes,
        "webextensions/_rels/webextension1.xml.rels",
        |content| {
            replace_once(
                content,
                b"</Relationships>",
                b"<!-- stale add-in relationship token --></Relationships>",
            )
        },
    );
    for (name, stale_bytes) in [
        ("semantic add-in XML", stale_bytes),
        ("root relationship XML", root_relationship_bytes),
        ("task-pane relationship XML", pane_relationship_bytes),
        ("new add-in relationship member", add_in_relationship_bytes),
    ] {
        let mut stale = Package::from_reader(Cursor::new(stale_bytes)).unwrap();
        let before = members(&package_bytes(&stale));
        let error = stale.apply_task_panes_patch(&noop).unwrap_err();
        let message = error.to_string().to_ascii_lowercase();
        assert!(
            message.contains("source")
                || message.contains("stale")
                || message.contains("relationship"),
            "{name} stale no-op refusal was not identifiable: {error}"
        );
        assert_eq!(members(&package_bytes(&stale)), before, "{name}");
    }
    let mut stale = Package::from_reader(Cursor::new(add_in_relationship_lexical_bytes)).unwrap();
    let before = members(&package_bytes(&stale));
    let error = stale.apply_task_panes_patch(&add_in_noop).unwrap_err();
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("source") || message.contains("stale") || message.contains("relationship"),
        "add-in relationship lexical stale refusal was not identifiable: {error}"
    );
    assert_eq!(members(&package_bytes(&stale)), before);
}

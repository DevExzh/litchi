mod support;

use litchi_odf_common::package::raw_identical_members;
use litchi_ods::{
    Spreadsheet,
    content_validation::{self, Definition, DisplayList, Error},
    document::Snapshot as DocumentSnapshot,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn document(inner: &str) -> String {
    format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{TEXT}"><o:body><o:spreadsheet>{inner}</o:spreadsheet></o:body></o:document-content>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        (
            "Pictures/opaque.bin",
            b"opaque-member".as_slice(),
            "application/octet-stream",
        ),
    ])
}

#[test]
fn absent_owner_crud_is_clone_staged_reversible_and_exact_source_checked() {
    let source = document(r#"<t:table t:name="Sheet"/>"#);
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();

    let mut first = Definition::new("whole").unwrap();
    first
        .set_condition(Some("cell-content-is-whole-number()"))
        .unwrap();
    first.set_allow_empty_cell(Some(true));
    edit.add(first).unwrap();
    assert!(matches!(
        edit.add(Definition::new("whole").unwrap()),
        Err(Error::DuplicateDefinition(name)) if name == "whole"
    ));

    let mut decimal = Definition::new("decimal").unwrap();
    decimal.set_display_list(Some(DisplayList::Unsorted));
    assert!(edit.set(decimal).unwrap().is_none());
    edit.update("decimal", |definition| {
        definition.set_base_cell_address(Some("$Sheet.$A$1"))
    })
    .unwrap();
    let replacement = {
        let mut value = Definition::new("decimal").unwrap();
        value.set_display_list(Some(DisplayList::SortAscending));
        value
    };
    assert_eq!(
        edit.replace("decimal", replacement).unwrap().name(),
        "decimal"
    );
    assert!(matches!(
        edit.replace("decimal", Definition::new("renamed").unwrap()),
        Err(Error::UnsafeRename { .. })
    ));
    assert_eq!(edit.remove("missing").unwrap(), None);

    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let target = content_validation::Snapshot::parse(commit.content_xml()).unwrap();
    assert_eq!(target.definitions().len(), 2);
    assert_eq!(
        target.definition("decimal").unwrap().base_cell_address(),
        None
    );
    assert_eq!(
        target.definition("decimal").unwrap().display_list(),
        Some(DisplayList::SortAscending)
    );

    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied.content_xml(), commit.content_xml());
    let foreign_xml = source.replace("Sheet", "Foreign");
    let foreign = content_validation::Snapshot::parse(&foreign_xml).unwrap();
    assert_eq!(
        commit.patch().apply(&foreign).unwrap_err(),
        Error::SourceMismatch
    );

    let inverse = commit.patch().inverse().unwrap();
    let restored = inverse.apply(&target).unwrap();
    assert_eq!(restored.content_xml(), source);
    assert_eq!(
        inverse
            .apply(&content_validation::Snapshot::parse(&source).unwrap())
            .unwrap_err(),
        Error::SourceMismatch
    );
}

#[test]
fn rollback_and_semantic_noops_retain_the_exact_source() {
    let source = document(
        r#"<t:content-validations><t:content-validation t:name="v" t:allow-empty-cell="true"/></t:content-validations>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();
    let same = snapshot.definition("v").unwrap().clone();
    assert_eq!(edit.set(same).unwrap().unwrap().name(), "v");
    edit.update("v", |_definition| Ok(())).unwrap();
    edit.add(Definition::new("temporary").unwrap()).unwrap();
    edit.rollback().unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.content_xml(), source);
}

#[test]
fn repeated_and_covered_bindings_close_removal_and_clear() {
    let source = document(
        r#"<t:content-validations><t:content-validation t:name="unused"/><t:content-validation t:name="bound"/></t:content-validations><t:table t:name="Sheet"><t:table-row t:number-rows-repeated="1000"><t:table-cell t:number-columns-repeated="50" t:content-validation-name="bound"/><t:covered-table-cell t:number-columns-repeated="7" t:content-validation-name="bound"/></t:table-row></t:table>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();
    assert!(matches!(
        edit.remove("bound"),
        Err(Error::DefinitionReferenced { bindings: 2, .. })
    ));
    assert!(matches!(
        edit.clear(),
        Err(Error::DefinitionReferenced { bindings: 2, .. })
    ));
    assert_eq!(edit.remove("unused").unwrap().unwrap().name(), "unused");
    let commit = edit.commit().unwrap();
    let target = content_validation::Snapshot::parse(commit.content_xml()).unwrap();
    assert_eq!(target.dangling_binding_count(), 0);
    assert_eq!(target.sheets()[0].bindings().len(), 2);
    assert_eq!(target.sheets()[0].bindings()[0].row_count(), 1000);
    assert!(target.sheets()[0].bindings()[1].is_covered_cell());
}

#[test]
fn dangling_bindings_can_only_be_repaired_before_commit() {
    let source = document(
        r#"<t:table t:name="Sheet"><t:table-row><t:table-cell t:content-validation-name="missing"/></t:table-row></t:table>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    assert!(!snapshot.edit().unwrap().commit().unwrap().changed());
    let mut repair = snapshot.edit().unwrap();
    repair.add(Definition::new("missing").unwrap()).unwrap();
    assert_eq!(
        content_validation::Snapshot::parse(repair.commit().unwrap().content_xml())
            .unwrap()
            .dangling_binding_count(),
        0
    );
}

#[test]
fn absent_self_closing_spreadsheet_expands_and_empty_catalog_fails_closed() {
    let source = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet/></o:body></o:document-content>"#
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.add(Definition::new("v").unwrap()).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.content_xml().contains("<o:spreadsheet>"));
    assert!(commit.content_xml().contains("table:name=\"v\""));

    let empty = document("<t:content-validations/>");
    assert!(content_validation::Snapshot::parse(&empty).is_err());
}

#[test]
fn opaque_mce_and_dtd_owners_never_enter_mutation() {
    for source in [
        document(
            r#"<t:content-validations><!--opaque--><t:content-validation t:name="v"/></t:content-validations>"#,
        ),
        document(
            r#"<t:content-validations><t:content-validation t:name="v"><t:help-message><x:p>help</x:p></t:help-message></t:content-validation></t:content-validations>"#,
        ),
        document(
            r#"<t:content-validations><t:content-validation t:name="v"/></t:content-validations><t:table t:name="S"><t:table-row><t:table-cell t:content-validation-name="v"><x:p>bound</x:p></t:table-cell></t:table-row></t:table>"#,
        ),
    ] {
        let snapshot = content_validation::Snapshot::parse(&source).unwrap();
        assert_eq!(
            snapshot
                .edit()
                .unwrap()
                .add(Definition::new("new").unwrap())
                .unwrap_err(),
            Error::OpaqueOwner
        );
    }

    let mce = document(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback/></mc:AlternateContent>"#,
    );
    assert_eq!(
        content_validation::Snapshot::parse(&mce).unwrap_err(),
        Error::UnsupportedMarkupCompatibility
    );
    let dtd = format!(
        r#"<!DOCTYPE o:document-content><o:document-content xmlns:o="{OFFICE}"><o:body><o:spreadsheet/></o:body></o:document-content>"#
    );
    assert!(content_validation::Snapshot::parse(&dtd).is_err());
}

#[test]
fn unified_publication_preserves_members_and_enforces_security_only_at_commit() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="Sheet"/></office:spreadsheet></office:body></office:document-content>"#
    );
    let source = package(&content);
    let snapshot = DocumentSnapshot::from_bytes(source.clone()).unwrap();
    let mut edit = snapshot.edit();
    edit.content_validations(|validations| validations.add(Definition::new("v")?))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert!(
        reopened
            .content_validations()
            .unwrap()
            .definition("v")
            .is_some()
    );
    let identical = raw_identical_members(&source, commit.snapshot().as_bytes()).unwrap();
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));

    let signed_source = support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        (
            "META-INF/documentsignatures.xml",
            br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
            "text/xml",
        ),
    ]);
    let signed = DocumentSnapshot::from_bytes(signed_source.clone()).unwrap();
    assert_eq!(
        signed.edit().commit().unwrap().snapshot().as_bytes(),
        signed_source
    );
    let mut changed = signed.edit();
    changed
        .content_validations(|validations| validations.add(Definition::new("v")?))
        .unwrap();
    assert!(changed.commit().is_err());
}

#[test]
fn unified_patch_apply_is_stale_safe_reversible_and_failure_atomic() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="Sheet"/></office:spreadsheet></office:body></office:document-content>"#
    );
    let source = package(&content);
    let semantic_source = content_validation::Snapshot::parse(&content).unwrap();
    let mut semantic_edit = semantic_source.edit().unwrap();
    semantic_edit.add(Definition::new("v").unwrap()).unwrap();
    let semantic_commit = semantic_edit.commit().unwrap();
    let patch = semantic_commit.patch().clone();

    let document_source = DocumentSnapshot::from_bytes(source.clone()).unwrap();
    let mut document_edit = document_source.edit();
    document_edit
        .apply_content_validation_patch(&patch)
        .unwrap();
    let target = document_edit.commit().unwrap();
    let target_content = Spreadsheet::from_bytes(target.snapshot().as_bytes().to_vec())
        .unwrap()
        .content_xml()
        .to_string();
    let semantic_target = content_validation::Snapshot::parse(&target_content).unwrap();
    let inverse = patch.inverse().unwrap();
    assert_eq!(
        inverse.apply(&semantic_target).unwrap().content_xml(),
        content
    );

    let mut restore = target.snapshot().edit();
    restore.apply_content_validation_patch(&inverse).unwrap();
    let restored = restore.commit().unwrap();
    assert_eq!(
        Spreadsheet::from_bytes(restored.snapshot().as_bytes().to_vec())
            .unwrap()
            .content_xml(),
        content
    );
    assert!(
        raw_identical_members(&source, restored.snapshot().as_bytes())
            .unwrap()
            .contains("Pictures/opaque.bin")
    );

    let foreign_content = content.replace("Sheet", "Foreign");
    let foreign = package(&foreign_content);
    let foreign_snapshot = DocumentSnapshot::from_bytes(foreign.clone()).unwrap();
    let mut stale = foreign_snapshot.edit();
    assert!(stale.apply_content_validation_patch(&patch).is_err());
    assert_eq!(stale.as_bytes(), foreign);
}

#[test]
fn typed_operation_and_output_limits_are_failure_atomic() {
    let source = document(r#"<t:table t:name="Sheet"/>"#);
    let operation_snapshot = content_validation::Snapshot::parse_with_limits(
        &source,
        content_validation::Limits::default().with_operations(0),
    )
    .unwrap();
    let mut operation_edit = operation_snapshot.edit().unwrap();
    assert!(matches!(
        operation_edit.add(Definition::new("v").unwrap()),
        Err(Error::LimitExceeded {
            kind: content_validation::LimitKind::Operations,
            observed: 1,
            maximum: 0,
        })
    ));
    assert!(operation_edit.definitions().is_empty());

    let output_snapshot = content_validation::Snapshot::parse_with_limits(
        &source,
        content_validation::Limits::default().with_output_bytes(source.len()),
    )
    .unwrap();
    let mut output_edit = output_snapshot.edit().unwrap();
    output_edit.add(Definition::new("v").unwrap()).unwrap();
    assert!(matches!(
        output_edit.commit(),
        Err(Error::LimitExceeded {
            kind: content_validation::LimitKind::OutputBytes,
            maximum,
            ..
        }) if maximum == source.len()
    ));
}

#[test]
fn tight_source_input_and_larger_output_support_add_and_patch_readback() {
    let source = document(r#"<t:table t:name="Sheet"/>"#);
    let limits = content_validation::Limits::default()
        .with_input_bytes(source.len())
        .with_output_bytes(source.len() + 1024);
    let snapshot = content_validation::Snapshot::parse_with_limits(&source, limits).unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.add(Definition::new("v").unwrap()).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.content_xml().len() > source.len());
    assert!(commit.patch().apply(&snapshot).unwrap().changed());
}

#[test]
fn tiny_output_budget_refuses_before_rendering_large_escaped_values() {
    let source = document(r#"<t:table t:name="Sheet"/>"#);
    let limits = content_validation::Limits::default()
        .with_input_bytes(source.len())
        .with_output_bytes(source.len() + 10);
    let snapshot = content_validation::Snapshot::parse_with_limits(&source, limits).unwrap();
    let mut definition = Definition::new("large").unwrap();
    let large = "&".repeat(2 * 1024 * 1024);
    definition.set_condition(Some(&large)).unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.add(definition).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::LimitExceeded {
            kind: content_validation::LimitKind::OutputBytes,
            maximum,
            ..
        }) if maximum == source.len() + 10
    ));
}

#[test]
fn inverse_handles_mixed_replacement_deletion_and_addition_offsets() {
    let source = document(
        r#"<t:content-validations><t:content-validation t:name="change" t:allow-empty-cell="false"/><t:content-validation t:name="remove"/><t:content-validation t:name="keep"/></t:content-validations>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.update("change", |definition| {
        definition.set_allow_empty_cell(Some(true));
        Ok(())
    })
    .unwrap();
    edit.remove("remove").unwrap();
    edit.add(Definition::new("addition-with-a-longer-name").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    let target = content_validation::Snapshot::parse(commit.content_xml()).unwrap();
    let restored = commit.patch().inverse().unwrap().apply(&target).unwrap();
    assert_eq!(restored.content_xml(), source);
}

#[test]
fn clear_removes_an_unbound_catalog_and_reparses_as_absent() {
    let source = document(
        r#"<t:content-validations><t:content-validation t:name="one"/><t:content-validation t:name="two"/></t:content-validations><t:table t:name="Sheet"/>"#,
    );
    let snapshot = content_validation::Snapshot::parse(&source).unwrap();
    let mut edit = snapshot.edit().unwrap();
    assert_eq!(edit.clear().unwrap(), 2);
    let commit = edit.commit().unwrap();
    let target = content_validation::Snapshot::parse(commit.content_xml()).unwrap();
    assert!(target.definitions().is_empty());
    assert!(!commit.content_xml().contains("content-validations"));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .unwrap()
            .apply(&target)
            .unwrap()
            .content_xml(),
        source
    );
}

#[test]
fn many_colliding_namespace_prefixes_are_indexed_once_and_get_a_fresh_prefix() {
    use std::fmt::Write as _;

    let mut declarations = String::new();
    for suffix in 0..200 {
        let prefix = if suffix == 0 {
            "litchicv".to_string()
        } else {
            format!("litchicv{suffix}")
        };
        write!(declarations, " xmlns:{prefix}=\"urn:test:{suffix}\"").unwrap();
    }
    let source = document(&format!(
        r#"<t:content-validations><t:content-validation t:name="v"{declarations}/></t:content-validations>"#
    ));
    let snapshot = content_validation::Snapshot::parse_with_limits(
        &source,
        content_validation::Limits::default().with_attributes(210),
    )
    .unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.update("v", |definition| {
        definition.set_allow_empty_cell(Some(true));
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.content_xml().contains("xmlns:litchicv200="));
    assert!(
        content_validation::Snapshot::parse(commit.content_xml())
            .unwrap()
            .definition("v")
            .is_some()
    );
}

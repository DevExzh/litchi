//! Native VBA storage and inert project-source inspection.
#![cfg(feature = "vba-inspection")]

use litchi_vba::{LimitKind, Limits, Payload, project::Project};
use std::path::PathBuf;

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
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
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn changed(before: &[u8], after: &[u8]) -> Vec<String> {
    let a = members(before);
    let b = members(after);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

fn assert_same_graph(before: &[u8], after: &[u8]) {
    let a =
        litchi_opc::OpcPackage::from_bytes(before, litchi_opc::SourceLimits::default()).unwrap();
    let b = litchi_opc::OpcPackage::from_bytes(after, litchi_opc::SourceLimits::default()).unwrap();
    let relationships = |rels: &litchi_opc::Relationships| {
        rels.iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    (
                        r.reltype().to_owned(),
                        r.target_ref().to_owned(),
                        r.is_external(),
                    ),
                )
            })
            .collect::<std::collections::BTreeMap<_, _>>()
    };
    assert_eq!(relationships(a.rels()), relationships(b.rels()));
    for part in a.iter_parts() {
        let restored = b.get_part(part.partname()).unwrap();
        assert_eq!(restored.content_type(), part.content_type());
        assert_eq!(relationships(restored.rels()), relationships(part.rels()));
    }
}

fn sha256(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
const EXCEL_SOURCE: &str = r#"Attribute VB_Name = "Module1"
Sub FixtureBoldSelection()
Attribute FixtureBoldSelection.VB_Description = "Synthetic fixture macro recorded from a bold-formatting action. No automatic execution."
Attribute FixtureBoldSelection.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FixtureBoldSelection Macro
' Synthetic fixture macro recorded from a bold-formatting action. No automatic execution.
'

'
    Selection.Font.Bold = True
End Sub
"#;
fn assert_project(project: &Project) {
    assert_eq!(project.name(), "VBAProject");
    assert_eq!(project.page_id(), 10000);
    assert_eq!(project.modules().len(), 1);
    let module = &project.modules()[0];
    assert_eq!(module.name(), "Module1");
    assert_eq!(module.stream_name(), "Module1");
    assert_eq!(module.kind(), litchi_vba::dir::Kind::Procedural);
    assert_eq!(module.text_offset(), 1428);
    assert!(!module.is_read_only() && !module.is_private());
    assert!(!module.source().had_decode_errors());
    assert_eq!(module.source().text(), EXCEL_SOURCE);
    assert_eq!(module.source().raw(), EXCEL_SOURCE.as_bytes());
    assert_eq!(
        project
            .references()
            .iter()
            .map(|r| r.name().unwrap())
            .collect::<Vec<_>>(),
        ["MSForms", "Office"]
    );
    assert!(project.references().iter().all(|r| r.raw().is_some()));
    assert!(project.designer_frames().is_empty() && project.license_info().is_none());
    let limits = Limits::default();
    let text = project.project_text(&limits).unwrap();
    assert_eq!(text.project_name(), Some("VBAProject"));
    assert_eq!(
        project.encode_project_text(text, &limits).unwrap(),
        project.project_properties().raw()
    );
    let maps = project.project_wm(&limits).unwrap().unwrap();
    assert_eq!(maps.len(), 1);
    assert_eq!(
        (maps[0].mbcs_name(), maps[0].unicode_name()),
        ("Module1", "Module1")
    );
    assert_eq!(
        project.encode_project_wm(maps, &limits).unwrap(),
        project.project_wm_raw().unwrap()
    );
    let tight = Limits {
        max_modules: 0,
        ..limits
    };
    assert!(matches!(
        project.project_text(&tight),
        Err(litchi_vba::Error::Limit {
            kind: LimitKind::Modules,
            observed: 1,
            maximum: 0
        })
    ));
    assert!(matches!(
        project.project_wm(&tight),
        Err(litchi_vba::Error::Limit {
            kind: LimitKind::Modules,
            observed: 1,
            maximum: 0
        })
    ));
}
#[test]
fn native_excel_binary_vba_graph_source_and_limits_are_explicit() {
    let bytes = native("excel-recorded-macro.xlsb");
    let package = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let metadata = package.vba().unwrap().unwrap();
    assert_eq!(metadata.source_part_name().as_str(), "/xl/workbook.bin");
    assert_eq!(metadata.relationship_id(), "rId5");
    assert_eq!(metadata.project_part_name().as_str(), "/xl/vbaProject.bin");
    assert!(metadata.signatures().is_empty());
    let part = package
        .opc_package()
        .get_part(metadata.project_part_name())
        .unwrap();
    assert_eq!(part.content_type(), "application/vnd.ms-office.vbaProject");
    assert_eq!(part.blob().len(), 10752);
    assert_eq!(
        sha256(part.blob()),
        "ce2ef3b47a7c28ce3cbf3287477685da2897004ec267dab8af28606220075839"
    );
    let project = metadata
        .project(package.opc_package(), &Limits::default())
        .unwrap();
    assert_project(&project);
    assert!(project.project_root_path().is_empty());
    assert!(project.project_text(&Limits::default()).unwrap().records().iter()
        .any(|r| matches!(r, litchi_vba::project::ProjectTextRecord::Id(value) if value == "{86708EE3-4AFE-5D4F-B718-F9FAF26C517C}")));
    let tight = Limits {
        max_modules: 0,
        ..Limits::default()
    };
    assert!(matches!(
        metadata.project(package.opc_package(), &tight),
        Err(litchi_xlsb::package::PackageError::Common(
            litchi_ooxml_common::Error::Vba(litchi_vba::Error::Limit {
                kind: LimitKind::Modules,
                observed: 1,
                maximum: 0
            })
        ))
    ));
}

#[test]
fn native_excel_binary_vba_noop_remove_and_inverse_preserve_all_other_members() {
    let bytes = native("excel-recorded-macro.xlsb");
    let package = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let payload = members(&bytes)["xl/vbaProject.bin"].clone();
    let mut noop = package.edit_vba().unwrap();
    noop.put(Payload::read(payload, &Limits::default()).unwrap())
        .unwrap();
    let commit = noop.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert!(commit.package().to_bytes().unwrap() == bytes);
    let mut edit = package.edit_vba().unwrap();
    assert!(edit.clear().unwrap());
    assert!(!edit.clear().unwrap());
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    assert_eq!(
        changed(&bytes, &after),
        [
            "[Content_Types].xml",
            "xl/_rels/workbook.bin.rels",
            "xl/vbaProject.bin"
        ]
    );
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    updated.save(output.path()).unwrap();
    let reopened = litchi_xlsb::Package::open(output.path()).unwrap();
    assert!(reopened.vba().unwrap().is_none());
    assert_eq!(
        reopened
            .opc_package()
            .main_document_part()
            .unwrap()
            .content_type(),
        litchi_opc::constants::content_type::XLSB_BIN
    );
    let restored = reopened.apply_vba(&patch.inverse()).unwrap().into_package();
    assert!(
        changed(&bytes, &restored.to_bytes().unwrap()).is_empty(),
        "inverse changed {:?}",
        changed(&bytes, &restored.to_bytes().unwrap())
    );
    let metadata = restored.vba().unwrap().unwrap();
    assert_project(
        &metadata
            .project(restored.opc_package(), &Limits::default())
            .unwrap(),
    );
    assert!(package.to_bytes().unwrap() == bytes);
}

#[test]
fn native_excel_binary_vba_replacement_keeps_source_modules_and_unrelated_parts() {
    let bytes = native("excel-recorded-macro.xlsb");
    let replacement = members(&native("excel-recorded-macro.xlsm"))["xl/vbaProject.bin"].clone();
    let package = litchi_xlsb::Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_vba().unwrap();
    edit.put(Payload::read(replacement.clone(), &Limits::default()).unwrap())
        .unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    let changes = changed(&bytes, &after);
    assert!(
        changes.iter().all(|n| [
            "[Content_Types].xml",
            "xl/_rels/workbook.bin.rels",
            "xl/vbaProject.bin"
        ]
        .contains(&n.as_str())),
        "{changes:?}"
    );
    assert_eq!(members(&after)["xl/vbaProject.bin"], replacement);
    assert_same_graph(&bytes, &after);
    let reopened = litchi_xlsb::Package::from_bytes(&after).unwrap();
    let metadata = reopened.vba().unwrap().unwrap();
    assert_eq!(metadata.relationship_id(), "rId5");
    assert_project(
        &metadata
            .project(reopened.opc_package(), &Limits::default())
            .unwrap(),
    );
    let restored = reopened.apply_vba(&patch.inverse()).unwrap().into_package();
    assert!(changed(&bytes, &restored.to_bytes().unwrap()).is_empty());
}

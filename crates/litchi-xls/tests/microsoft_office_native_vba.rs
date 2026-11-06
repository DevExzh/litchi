//! Native Excel VBA metadata/source inspection without macro execution.

use litchi_vba::{LimitKind, Limits, project::Project};
use std::{io::Cursor, path::PathBuf};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

fn streams(bytes: &[u8]) -> std::collections::BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names: Vec<_> = path.iter().map(String::as_str).collect();
            let bytes = ole.open_stream(&names).unwrap();
            (path, bytes)
        })
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
fn native_excel_biff_vba_markers_storage_and_source_agree() {
    for name in ["excel-recorded-macro.xls", "excel-recorded-macro.xlt"] {
        let bytes = native(name);
        let mut reader = litchi_xls::raw::Reader::new(Cursor::new(&bytes)).unwrap();
        let metadata = reader.vba_metadata();
        assert!(metadata.has_project_marker() && metadata.has_project_storage());
        assert!(!metadata.has_no_macros_marker());
        assert!(metadata.markers_are_consistent() && metadata.may_contain_executable_code());
        assert_eq!(metadata.workbook_code_name(), None);
        assert_eq!(reader.xls_worksheet(0).unwrap().vba_code_name(), None);
        let storage = reader.vba_project_storage().unwrap();
        assert_eq!(storage.root_storage_path(), ["_VBA_PROJECT_CUR"]);
        assert_eq!(
            storage.vba_storage_path().unwrap(),
            ["_VBA_PROJECT_CUR", "VBA"]
        );
        assert!(storage.is_structurally_complete() && storage.may_contain_macro_code());
        assert!(storage.has_project_wm_stream() && !storage.has_project_lk_stream());
        assert_eq!(storage.candidate_module_stream_names(), ["Module1"]);
        assert_eq!(
            storage.srp_stream_names(),
            ["__SRP_0", "__SRP_1", "__SRP_2", "__SRP_3"]
        );
        let project = reader.vba(&Limits::default()).unwrap().unwrap();
        assert_eq!(project.project_root_path(), ["_VBA_PROJECT_CUR"]);
        assert_project(&project);
        let tight = Limits {
            max_modules: 0,
            ..Limits::default()
        };
        assert!(matches!(
            reader.vba(&tight),
            Err(litchi_xls::Error::Vba(litchi_vba::Error::Limit {
                kind: LimitKind::Modules,
                observed: 1,
                maximum: 0
            }))
        ));
    }
}

#[test]
fn native_excel_biff_cell_edit_preserves_and_reopens_the_vba_project() {
    for name in ["excel-recorded-macro.xls", "excel-recorded-macro.xlt"] {
        let bytes = native(name);
        let package = litchi_xls::Package::from_bytes(&bytes).unwrap();
        let at = litchi_xls::cell_values::Reference::new(1, 1).unwrap();
        let mut edit = package.edit().unwrap();
        edit.set_value("Sheet1", at, litchi_xls::cell_values::Value::Number(43.0))
            .unwrap();
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        updated.save(output.path()).unwrap();
        let after = std::fs::read(output.path()).unwrap();
        let a = streams(&bytes);
        let b = streams(&after);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        let changed = a
            .iter()
            .filter(|(path, data)| **data != b[*path])
            .map(|(path, _)| path.clone())
            .collect::<Vec<_>>();
        assert_eq!(changed, [vec!["Workbook".to_owned()]]);
        let mut reader = litchi_xls::raw::Reader::new(Cursor::new(&after)).unwrap();
        let metadata = reader.vba_metadata();
        assert!(metadata.has_project_marker() && metadata.has_project_storage());
        assert!(!metadata.has_no_macros_marker());
        assert_eq!(metadata.workbook_code_name(), None);
        assert_eq!(reader.xls_worksheet(0).unwrap().vba_code_name(), None);
        assert_project(&reader.vba(&Limits::default()).unwrap().unwrap());
        let reopened = litchi_xls::Package::from_bytes(&after).unwrap();
        assert_eq!(
            reopened
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            Some(litchi_core::sheet::CellValue::Float(43.0))
        );
        let restored = reopened.apply(&patch.inverse()).unwrap().into_package();
        assert!(restored.to_bytes().unwrap() == bytes);
    }
}

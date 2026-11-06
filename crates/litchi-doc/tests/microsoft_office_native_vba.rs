//! Native Word VBA metadata and source are inspected, never executed.

use litchi_vba::{LimitKind, Limits, project::Project};
use std::{io::Cursor, path::PathBuf};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
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
const DOCUMENT_SOURCE: &str = r#"Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "1Normal.ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = True
Attribute VB_Customizable = True
"#;
const MACRO_SOURCE: &str = r#"Attribute VB_Name = "NewMacros"
Sub FixtureBoldSelection()
Attribute FixtureBoldSelection.VB_Description = "Synthetic fixture macro: applies bold to the selected text. No automatic execution."
Attribute FixtureBoldSelection.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.FixtureBoldSelection"
'
' FixtureBoldSelection Macro
' Synthetic fixture macro: applies bold to the selected text. No automatic execution.
'
    Selection.Font.Bold = wdToggle
End Sub
"#;
fn assert_project(project: &Project) {
    use litchi_vba::dir::Kind;
    assert_eq!(project.name(), "Project");
    assert_eq!(project.page_id(), 10000);
    assert_eq!(project.modules().len(), 2);
    assert_eq!(
        project
            .references()
            .iter()
            .map(|r| r.name().unwrap())
            .collect::<Vec<_>>(),
        ["MSForms", "Normal", "Office"]
    );
    assert!(project.references().iter().all(|r| r.raw().is_some()));
    for (module, name, offset, kind, expected) in [
        (
            &project.modules()[0],
            "ThisDocument",
            931,
            Kind::DocumentClassOrDesigner,
            DOCUMENT_SOURCE,
        ),
        (
            &project.modules()[1],
            "NewMacros",
            1420,
            Kind::Procedural,
            MACRO_SOURCE,
        ),
    ] {
        assert_eq!(module.name(), name);
        assert_eq!(module.stream_name(), name);
        assert_eq!(module.text_offset(), offset);
        assert_eq!(module.kind(), kind);
        assert!(!module.is_read_only() && !module.is_private());
        assert!(!module.source().had_decode_errors());
        assert_eq!(module.source().text(), expected);
        assert_eq!(module.source().raw(), expected.as_bytes());
    }
    assert!(project.designer_frames().is_empty());
    assert!(project.license_info().is_none());
    let limits = Limits::default();
    let text = project.project_text(&limits).unwrap();
    assert_eq!(text.project_name(), Some("Project"));
    assert_eq!(
        project.encode_project_text(text, &limits).unwrap(),
        project.project_properties().raw()
    );
    let maps = project.project_wm(&limits).unwrap().unwrap();
    assert_eq!(
        maps.iter()
            .map(|m| (m.mbcs_name(), m.unicode_name()))
            .collect::<Vec<_>>(),
        [("ThisDocument", "ThisDocument"), ("NewMacros", "NewMacros")]
    );
    assert_eq!(
        project.encode_project_wm(maps, &limits).unwrap(),
        project.project_wm_raw().unwrap()
    );
    let tight = Limits {
        max_modules: 1,
        ..limits
    };
    assert!(matches!(
        project.project_text(&tight),
        Err(litchi_vba::Error::Limit {
            kind: LimitKind::Modules,
            observed: 2,
            maximum: 1
        })
    ));
    assert!(matches!(
        project.project_wm(&tight),
        Err(litchi_vba::Error::Limit {
            kind: LimitKind::Modules,
            observed: 2,
            maximum: 1
        })
    ));
}
#[test]
fn native_word_vba_storage_and_decoded_modules_are_bounded_and_inert() {
    for name in ["word-recorded-macro.doc", "word-recorded-macro.dot"] {
        let bytes = native(name);
        let mut reader = litchi_doc::raw::Reader::from_reader(Cursor::new(&bytes)).unwrap();
        let storage = reader.vba_project_storage().unwrap();
        assert_eq!(reader.vba_project_storages().len(), 1);
        assert_eq!(storage.project_root_path(), ["Macros"]);
        assert_eq!(storage.vba_storage_path(), ["Macros", "VBA"]);
        assert!(storage.is_structurally_complete() && storage.may_contain_macro_code());
        assert!(storage.has_project_wm_stream() && !storage.has_project_lk_stream());
        assert_eq!(
            storage.candidate_module_stream_names(),
            ["NewMacros", "ThisDocument"]
        );
        assert_eq!(
            storage.srp_stream_names(),
            ["__SRP_0", "__SRP_1", "__SRP_2", "__SRP_3"]
        );
        let project = reader.vba(&Limits::default()).unwrap().unwrap();
        assert_eq!(project.project_root_path(), ["Macros"]);
        assert_project(&project);
        let tight = Limits {
            max_modules: 1,
            ..Limits::default()
        };
        assert!(matches!(
            reader.vba(&tight),
            Err(litchi_doc::package::Error::Vba(litchi_vba::Error::Limit {
                kind: LimitKind::Modules,
                observed: 2,
                maximum: 1
            }))
        ));
        assert_project(&reader.vba(&Limits::default()).unwrap().unwrap());
    }
}

#[test]
fn native_word_body_edit_keeps_every_vba_stream_and_reopens_typed_source() {
    for name in ["word-recorded-macro.doc", "word-recorded-macro.dot"] {
        let bytes = native(name);
        let package = litchi_doc::Package::from_bytes(&bytes).unwrap();
        let mut edit = package.edit().unwrap();
        edit.replace_paragraph(
            litchi_core::Position::new(0),
            "Edited native macro fixture title",
        )
        .unwrap();
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let output = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
        updated.save(output.path()).unwrap();
        let after = std::fs::read(output.path()).unwrap();
        let a = streams(&bytes);
        let b = streams(&after);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (path, data) in &a {
            if !matches!(path.as_slice(), [name] if matches!(name.as_str(), "WordDocument" | "0Table" | "1Table"))
            {
                assert!(data == &b[path], "changed unrelated stream {path:?}");
            }
        }
        let mut reader = litchi_doc::raw::Reader::from_reader(Cursor::new(&after)).unwrap();
        assert_project(&reader.vba(&Limits::default()).unwrap().unwrap());
        let reopened = litchi_doc::Package::from_bytes(&after).unwrap();
        assert_eq!(
            reopened
                .document()
                .body_paragraph_text(litchi_core::Position::new(0))
                .unwrap(),
            Some("Edited native macro fixture title")
        );
        let restored = reopened.apply(&patch.inverse()).unwrap().into_package();
        assert!(restored.to_bytes().unwrap() == bytes);
    }
}

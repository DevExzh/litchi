//! Public operation contracts on native Microsoft Office specimens.
#![cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "pptx",
    feature = "ppt",
    feature = "xlsx",
    feature = "xlsb",
    feature = "xls"
))]

use litchi_ops::{OpenOptions, SourceMode, props};

#[cfg(feature = "ppt")]
#[test]
fn native_ppt_queries_and_atomic_title_visibility_edits_preserve_timing() {
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-master-motion.ppt"
    );
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("native.ppt");
    std::fs::write(&path, bytes).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = if mode == SourceMode::Auto {
            litchi_ops::open_with(&path, &options).unwrap()
        } else {
            litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap()
        };
        let found = source
            .query("placeholder[text=PPT_MASTER_MOTION_006]")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(found.len(), 1);
        let title = found[0].path().to_string();
        assert_eq!(title, "/slide[2]/shape[1]");
        let committed = source
            .edit(|edit| {
                edit.set(title.as_str(), props! {"text"=>"Reviewed native title"})?;
                edit.set("/slide[4]", props! {"hidden"=>false})
            })
            .unwrap();
        let after = committed.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(after.clone(), &options).unwrap();
        assert_eq!(
            reopened
                .get(title.as_str())
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Reviewed native title")
        );
        let owner = litchi_ppt::Package::from_bytes(&after).unwrap();
        let p = owner.presentation().unwrap();
        assert!(!p.slide_at(3).unwrap().unwrap().timing().unwrap().hidden);
        assert_eq!(
            p.slide_at(1)
                .unwrap()
                .unwrap()
                .transition()
                .unwrap()
                .unwrap()
                .advance_time_ms,
            Some(3000)
        );
        assert!(
            p.slide_at(1)
                .unwrap()
                .unwrap()
                .animation_extension()
                .unwrap()
                .is_some()
        );
        assert!(
            committed
                .patch()
                .inverse()
                .apply(committed.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(source.to_bytes().unwrap() == bytes);
    }
    assert!(std::fs::read(path).unwrap() == bytes);
}

#[cfg(feature = "xls")]
#[test]
fn native_xls_formula_queries_and_cell_edits_keep_array_caches_on_both_owners() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-formula-owners.xls");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let matches = source
            .query("cell[value=206]")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].path().to_string(), "/sheet[1]/D103");
        assert_eq!(
            source
                .get("Sheet1!D103")
                .unwrap()
                .view()
                .format()
                .unwrap()
                .find(|p| p.key() == "formula")
                .unwrap()
                .to_string(),
            "(ROW()*2)"
        );
        let committed = source
            .edit(|edit| {
                edit.set("Sheet1!Z121", props! {"value"=>42})?;
                edit.set("Sheet2!Z121", props! {"value"=>"=inert native text"})
            })
            .unwrap();
        let after = committed.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(after, &options).unwrap();
        for (path, text) in [
            ("Sheet1!Z121", "42"),
            ("Sheet2!Z121", "=inert native text"),
            ("Sheet1!C7", "200"),
            ("Sheet2!F4", "50"),
        ] {
            assert_eq!(
                reopened
                    .get(path)
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some(text)
            );
        }
        let restored = committed.patch().inverse().apply(committed.file()).unwrap();
        assert!(restored.to_bytes().unwrap() == bytes);
        assert!(source.to_bytes().unwrap() == bytes);
    }
}

#[cfg(feature = "xlsb")]
#[test]
fn native_xlsb_formula_queries_and_scalar_crud_keep_stored_caches_on_both_owners() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-formula-owners.xlsb");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let nodes = source
            .query("cell[value=206]")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(nodes.len(), 1);
        assert_eq!(nodes[0].path().to_string(), "/sheet[1]/D103");
        assert_eq!(
            source
                .get("Sheet1!D103")
                .unwrap()
                .view()
                .format()
                .unwrap()
                .find(|p| p.key() == "formula")
                .unwrap()
                .to_string(),
            "(ROW()*2)"
        );
        let committed = source
            .edit(|edit| {
                edit.set("Sheet1!Z100", props! {"value"=>42})?;
                edit.set("Sheet2!Z100", props! {"value"=>"=inert native text"})
            })
            .unwrap_or_else(|error| {
                let mut detail = error.to_string();
                let mut cause = std::error::Error::source(&error);
                while let Some(value) = cause {
                    detail.push_str(&format!("; {value}"));
                    cause = value.source();
                }
                panic!("native XLSB {mode:?} edit failed: {detail}")
            });
        let after = committed.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(after, &options).unwrap();
        for (path, expected) in [
            ("Sheet1!Z100", "42"),
            ("Sheet2!Z100", "=inert native text"),
            ("Sheet1!D103", "206"),
            ("Sheet2!F4", "50"),
        ] {
            assert_eq!(
                reopened
                    .get(path)
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some(expected)
            );
        }
        let updated = reopened
            .edit(|edit| edit.set("Sheet1!Z100", props! {"value"=>43}))
            .unwrap();
        assert_eq!(
            updated
                .file()
                .get("Sheet1!Z100")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("43")
        );
        let removed = updated
            .file()
            .edit(|edit| edit.remove("Sheet1!Z100"))
            .unwrap();
        assert_eq!(
            removed
                .file()
                .get("Sheet1!Z100")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("(empty)")
        );
        let removed_bytes = removed.file().to_bytes().unwrap();
        let owner = litchi_xlsb::Package::from_bytes(&removed_bytes).unwrap();
        assert!(
            owner
                .workbook()
                .unwrap()
                .grid(0_usize)
                .unwrap()
                .unwrap()
                .worksheet()
                .get_cell(99, 25)
                .is_none()
        );
        assert!(
            committed
                .patch()
                .inverse()
                .apply(committed.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(source.to_bytes().unwrap() == bytes);
    }
}

#[cfg(feature = "pptx")]
#[test]
fn native_ink_deck_missing_notes_have_the_same_first_error_in_all_edit_forms() {
    use litchi_ops::{BatchItem, BatchMode, Code, Props};
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-ink-replay.pptx");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let file =
            litchi_ops::open_bytes_with(bytes.to_vec(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let before = file.to_bytes().unwrap();
        let path = "/slide[1]/notes/paragraph[1]";
        assert_eq!(
            file.set(path, props! {"align"=>"right"})
                .unwrap_err()
                .code(),
            Code::PathNotFound
        );
        let mut first = BatchItem::new("set");
        first.path = Some(path.into());
        first.props = Some(Props::new().with("align", "right"));
        let mut second = BatchItem::new("set");
        second.path = Some("/slide[1]/shape[1]".into());
        second.props = Some(Props::new().with("text", "must not publish"));
        let error = file.batch([first, second], BatchMode::Atomic).unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(error.item(), Some(0));
        let error = file
            .edit(|edit| {
                edit.set(path, props! {"align"=>"right"})?;
                edit.set("/slide[1]/shape[1]", props! {"text"=>"must not publish"})
            })
            .unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
        assert_eq!(error.item(), Some(0));
        assert!(file.to_bytes().unwrap() == before);
    }
}

#[cfg(feature = "docx")]
#[test]
fn native_word_nested_table_paths_edit_and_restore_on_both_owners() {
    let bytes = include_bytes!("../../../test-data/microsoft-office/word/word-tables-layout.docx");
    let cell = "/body/tbl[1]/tr[3]/tc[4]/p[1]";
    let nested = "/body/tbl[1]/tr[4]/tc[4]/tbl[1]/tr[2]/tc[2]/p[1]/r[1]";
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let file = litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let before = file.to_bytes().unwrap();
        assert_eq!(
            file.get(cell).unwrap().view().text().unwrap().as_deref(),
            Some("Unicode 中文")
        );
        assert_eq!(
            file.get(nested).unwrap().view().text().unwrap().as_deref(),
            Some("20")
        );
        let commit = file
            .edit(|edit| {
                edit.set(cell, props! {"text"=>"Unicode 中文 revised"})?;
                edit.set(nested, props! {"text"=>"21"})
            })
            .unwrap();
        let published = commit.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(published.clone(), &options).unwrap();
        assert_eq!(
            reopened
                .get(cell)
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Unicode 中文 revised")
        );
        assert_eq!(
            reopened
                .get(nested)
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("21")
        );
        let original_members = native_word_members(&before);
        let changed_members = native_word_members(&published);
        assert_eq!(
            original_members.keys().collect::<Vec<_>>(),
            changed_members.keys().collect::<Vec<_>>()
        );
        for (name, data) in &original_members {
            if name != "word/document.xml" {
                assert!(changed_members[name] == *data, "changed {name}");
            }
        }
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap();
        let restored_members = native_word_members(&restored);
        assert_eq!(
            original_members.keys().collect::<Vec<_>>(),
            restored_members.keys().collect::<Vec<_>>()
        );
        for (name, data) in original_members {
            assert!(restored_members[&name] == data, "inverse changed {name}");
        }
        assert!(file.to_bytes().unwrap() == before);
    }
}

#[cfg(feature = "docx")]
fn native_word_members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = archive.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

#[cfg(feature = "doc")]
#[test]
fn native_word_binary_template_queries_and_atomic_edits_preserve_autotext() {
    let bytes = include_bytes!("../../../test-data/microsoft-office/word/word-glossary.dot");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let nodes = source
            .query("p")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(nodes.len(), 4);
        assert_eq!(nodes[0].positional_path().to_string(), "/body/p[1]");
        assert_eq!(
            source
                .get("/body/p[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("WORD_GLOSSARY_014")
        );
        let commit = source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"text"=>"NATIVE_TEMPLATE_EDIT"})?;
                edit.set("/body/p[2]", props! {"text"=>"Template metadata edited"})
            })
            .unwrap();
        let published = commit.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(published.clone(), &options).unwrap();
        assert_eq!(
            reopened
                .get("/body/p[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("NATIVE_TEMPLATE_EDIT")
        );
        assert_eq!(
            reopened
                .get("/body/p[2]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Template metadata edited")
        );
        let package = litchi_doc::Package::from_bytes(&published).unwrap();
        let document = package.document();
        let glossary = document.attached_glossary().unwrap().unwrap();
        assert_eq!(glossary.metadata().items()[0].name(), "LITCHI_AUTOTEXT_014");
        assert_eq!(
            glossary.item_text(0),
            Some("AUTOTEXT_BODY_014: reusable native Word building block.\r")
        );
        assert!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(source.to_bytes().unwrap() == bytes);
    }
}

#[cfg(feature = "xlsx")]
#[test]
fn native_excel_formula_queries_and_atomic_scalar_edits_work_on_both_owners() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/excel/excel-formula-owners.xlsx");
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let nodes = source
            .query("cell[value=206]")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(nodes.len(), 1);
        assert_eq!(nodes[0].path().to_string(), "/sheet[1]/D103");
        let cell = source.get("Sheet1!D103").unwrap();
        assert_eq!(
            cell.view()
                .format()
                .unwrap()
                .find(|p| p.key() == "formula")
                .unwrap()
                .to_string(),
            "ROW()*2"
        );
        let commit = source
            .edit(|edit| {
                edit.set("Sheet1!Z100", props! {"value"=>42})?;
                edit.set("Sheet2!Z100", props! {"value"=>"=inert native text"})
            })
            .unwrap();
        let after = commit.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(after.clone(), &options).unwrap();
        assert_eq!(
            reopened
                .get("Sheet1!Z100")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("42")
        );
        assert_eq!(
            reopened
                .get("Sheet2!Z100")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("=inert native text")
        );
        let package = litchi_xlsx::Package::from_bytes(&after).unwrap();
        assert!(
            matches!(package.workbook().sheet("Sheet2").unwrap().unwrap().cell("Z100").unwrap().stored(),Some(litchi_xlsx::Cell::Value(litchi_xlsx::Value::Text(value))) if value.as_str()=="=inert native text")
        );
        assert!(source.to_bytes().unwrap() == bytes);
        assert!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
    }
}

#[cfg(feature = "pptx")]
#[test]
fn native_pptx_table_queries_and_cell_row_edits_preserve_other_slides() {
    let bytes =
        include_bytes!("../../../test-data/microsoft-office/powerpoint/powerpoint-text-table.pptx");
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("native.pptx");
    std::fs::write(&path, bytes).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = if mode == SourceMode::Auto {
            litchi_ops::open_with(&path, &options).unwrap()
        } else {
            litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap()
        };
        let tables = source
            .query("table")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].path().to_string(), "/slide[3]/table[1]");
        let cell = "/slide[3]/table[1]/tr[2]/tc[1]";
        assert_eq!(
            source.get(cell).unwrap().view().text().unwrap().as_deref(),
            Some("Apples")
        );
        let changed = source
            .edit(|edit| {
                edit.set(cell, props! {"text"=>"Oranges"})?;
                edit.set("/slide[3]/table[1]/tr[2]", props! {"height"=>"0.5in"})
            })
            .unwrap();
        let after = changed.file().to_bytes().unwrap();
        let reopened = litchi_ops::open_bytes_with(after.clone(), &options).unwrap();
        assert_eq!(
            reopened
                .get(cell)
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Oranges")
        );
        assert_eq!(
            reopened
                .get("/slide[3]/table[1]/tr[3]/tc[3]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("12.50")
        );
        assert_eq!(
            reopened
                .get("/slide[3]/table[1]/tr[2]")
                .unwrap()
                .view()
                .format()
                .unwrap()
                .find(|p| p.key() == "height")
                .unwrap()
                .value()
                .to_string(),
            "1.27cm"
        );
        let before_owner = litchi_pptx::Package::from_bytes(bytes).unwrap();
        let after_owner = litchi_pptx::Package::from_bytes(&after).unwrap();
        for uri in [
            "/ppt/slides/slide1.xml",
            "/ppt/slides/slide2.xml",
            "/ppt/tableStyles.xml",
        ] {
            let uri = litchi_opc::PackURI::new(uri).unwrap();
            assert_eq!(
                before_owner.opc().unwrap().get_part(&uri).unwrap().blob(),
                after_owner.opc().unwrap().get_part(&uri).unwrap().blob()
            );
        }
        assert_eq!(
            changed
                .patch()
                .inverse()
                .apply(changed.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
    }
    assert_eq!(std::fs::read(path).unwrap(), bytes);
}

#[cfg(feature = "pptx")]
#[test]
fn native_pptx_transition_edits_preserve_zoom_owners_through_public_modes() {
    use litchi_pptx::transition::{self, Kind, Morph, Side};
    let bytes = include_bytes!(
        "../../../test-data/microsoft-office/powerpoint/powerpoint-master-motion.pptx"
    );
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("native.pptx");
    std::fs::write(&path, bytes).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let source = if mode == SourceMode::Auto {
            litchi_ops::open_with(&path, &options).unwrap()
        } else {
            litchi_ops::open_bytes_with(bytes.to_vec(), &options).unwrap()
        };
        let committed = source
            .edit(|edit| {
                edit.set("/slide[2]", props! {"transition"=>"push-left-1750"})?;
                edit.set("/slide[3]", props! {"transition"=>"morph-byword-1750"})
            })
            .unwrap();
        let after = committed.file().to_bytes().unwrap();
        let owner = litchi_pptx::Package::from_bytes(&after).unwrap();
        let p = owner.presentation().unwrap();
        let slides = p.slides().unwrap();
        for (index, kind) in [(1, Kind::Push(Side::Left)), (2, Kind::Morph(Morph::ByWord))] {
            let value = transition::read(
                slides[index].part().part().blob(),
                transition::Limits::DEFAULT,
            )
            .unwrap()
            .unwrap();
            assert_eq!(value.kind(), &kind);
            assert_eq!(value.duration().unwrap().get(), 1750);
            let original = litchi_pptx::Package::from_bytes(bytes).unwrap();
            let q = original.presentation().unwrap();
            let original = q.slides().unwrap();
            let a = original[index].zooms().unwrap();
            let b = slides[index].zooms().unwrap();
            assert!(a.iter().eq(b.iter()));
        }
        let cleared = committed
            .file()
            .set("/slide[3]", props! {"transition"=>"none"})
            .unwrap();
        let cleared_owner = litchi_pptx::Package::from_bytes(&cleared.to_bytes().unwrap()).unwrap();
        assert!(
            transition::read(
                cleared_owner.presentation().unwrap().slides().unwrap()[2]
                    .part()
                    .part()
                    .blob(),
                transition::Limits::DEFAULT
            )
            .unwrap()
            .is_none()
        );
        let restored = cleared
            .set("/slide[3]", props! {"transition"=>"morph-byword-1750"})
            .unwrap();
        let reopened = litchi_ops::open_bytes_with(restored.to_bytes().unwrap(), &options).unwrap();
        let restored_owner =
            litchi_pptx::Package::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        let restored_value = transition::read(
            restored_owner.presentation().unwrap().slides().unwrap()[2]
                .part()
                .part()
                .blob(),
            transition::Limits::DEFAULT,
        )
        .unwrap()
        .unwrap();
        assert_eq!(restored_value.kind(), &Kind::Morph(Morph::ByWord));
        assert_eq!(restored_value.duration().unwrap().get(), 1750);
        let original_members = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
            bytes.to_vec(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        let changed_members = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
            after,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        assert_eq!(
            original_members.member_names().unwrap(),
            changed_members.member_names().unwrap()
        );
        for name in original_members.member_names().unwrap() {
            if !["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"].contains(&name.as_str()) {
                assert!(
                    original_members.read_member(&name).unwrap()
                        == changed_members.read_member(&name).unwrap(),
                    "changed {name}"
                );
            }
        }
        assert!(
            committed
                .patch()
                .inverse()
                .apply(committed.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(source.to_bytes().unwrap() == bytes);
    }
    assert!(std::fs::read(path).unwrap() == bytes);
}

#[cfg(feature = "docx")]
#[test]
fn native_word_layout_extension_values_survive_public_title_edits() {
    use litchi_docx::{paragraph::Collapsed, section::footnote_columns::Layout};
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word");
    for name in ["word-layout-extensions.docx", "word-layout-extensions.dotx"] {
        let path = root.join(name);
        let bytes = std::fs::read(&path).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let options = OpenOptions::new().source_mode(mode);
            let file = if mode == SourceMode::Auto {
                litchi_ops::open_with(&path, &options).unwrap()
            } else {
                litchi_ops::open_bytes_with(bytes.clone(), &options).unwrap()
            };
            let found = file
                .query("p[text=COLLAPSED_HEADING_016]")
                .unwrap()
                .collect::<litchi_ops::Result<Vec<_>>>()
                .unwrap();
            assert_eq!(found.len(), 1);
            assert_eq!(found[0].path().to_string(), "/body/p[3]");
            let commit = file
                .edit(|edit| edit.set("/body/p[1]/r[1]", props! {"text"=>"LAYOUT_PUBLIC_REVIEWED"}))
                .unwrap();
            let after = commit.file().to_bytes().unwrap();
            let reopened = litchi_ops::open_bytes_with(after.clone(), &options).unwrap();
            assert_eq!(
                reopened
                    .get("/body/p[1]")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("LAYOUT_PUBLIC_REVIEWED")
            );
            let owner = litchi_docx::Package::from_bytes(&after).unwrap();
            let document = owner.document().unwrap();
            assert_eq!(
                document.paragraphs().unwrap()[2].collapsed().unwrap(),
                Some(Collapsed::Enabled)
            );
            assert_eq!(
                document
                    .sections()
                    .unwrap()
                    .iter_mut()
                    .next()
                    .unwrap()
                    .footnote_columns()
                    .unwrap(),
                Some(Layout::new(2).unwrap())
            );
            let original = litchi_docx::Package::from_bytes(&bytes).unwrap();
            for part in original.opc_package().iter_parts() {
                if part.partname().as_str() != "/word/document.xml" {
                    let after = owner.opc_package().get_part(part.partname()).unwrap();
                    let before = original.opc_package().get_part(part.partname()).unwrap();
                    assert!(after.blob() == before.blob(), "changed {}", part.partname());
                }
            }
            assert!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap()
                    == bytes
            );
            assert!(file.to_bytes().unwrap() == bytes);
        }
        assert!(std::fs::read(path).unwrap() == bytes);
    }
}

#[cfg(feature = "docx")]
#[test]
fn native_word_mail_merge_and_form_payloads_survive_public_title_edits() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word");
    for (name, title) in [
        ("word-mailmerge-fields.docx", "WORD_MAILMERGE_009"),
        ("word-legacy-form-unprotected.docx", "WORD_LEGACY_FORM_010"),
    ] {
        let path = root.join(name);
        let bytes = std::fs::read(&path).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let options = OpenOptions::new().source_mode(mode);
            let file = litchi_ops::open_with(&path, &options).unwrap();
            assert_eq!(
                file.get("/body/p[1]")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some(title)
            );
            let changed = file
                .set("/body/p[1]/r[1]", props! {"text"=>"NATIVE_FIELDS_REVIEWED"})
                .unwrap();
            let directory = tempfile::tempdir().unwrap();
            let destination = directory.path().join(name);
            changed.save_as(&destination).unwrap();
            let reopened = litchi_ops::open_with(&destination, &options).unwrap();
            assert_eq!(
                reopened
                    .get("/body/p[1]")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("NATIVE_FIELDS_REVIEWED")
            );
            let original = litchi_docx::Package::from_bytes(&bytes).unwrap();
            let owner = litchi_docx::Package::open(&destination).unwrap();
            assert_eq!(
                owner.mail_merge_settings().unwrap(),
                original.mail_merge_settings().unwrap()
            );
            for metadata in original.opc_package().iter_parts() {
                let before = original
                    .opc_package()
                    .get_part(metadata.partname())
                    .unwrap();
                let after = owner.opc_package().get_part(metadata.partname()).unwrap();
                if metadata.partname().as_str() == "/word/document.xml" {
                    // Exact rest-of-document equality includes ffData, nested
                    // field instructions, cached results and paired bookmarks.
                    let expected = std::str::from_utf8(before.blob()).unwrap().replacen(
                        title,
                        "NATIVE_FIELDS_REVIEWED",
                        1,
                    );
                    assert!(after.blob() == expected.as_bytes());
                } else {
                    assert!(
                        after.blob() == before.blob(),
                        "changed {}",
                        metadata.partname()
                    );
                }
            }
            assert!(file.to_bytes().unwrap() == bytes);
        }
        assert!(std::fs::read(path).unwrap() == bytes);
    }
}

#[cfg(feature = "docx")]
#[test]
fn native_word_smartart_outline_edits_preserve_extent_and_other_visuals() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word/word-vector-diagram.docx");
    let bytes = std::fs::read(&path).unwrap();
    let original = litchi_docx::Package::from_bytes(&bytes).unwrap();
    let uri = litchi_opc::PackURI::new("/word/document.xml").unwrap();
    let source_xml =
        std::str::from_utf8(original.opc_package().get_part(&uri).unwrap().blob()).unwrap();
    let start = source_xml.find("<mc:AlternateContent>").unwrap();
    let end = source_xml[start..].find("</mc:AlternateContent>").unwrap()
        + start
        + "</mc:AlternateContent>".len();
    let converted_group = &source_xml[start..end];
    for (mode, node_only) in [
        (SourceMode::Auto, true),
        (SourceMode::Eager, true),
        (SourceMode::Auto, false),
        (SourceMode::Eager, false),
    ] {
        let options = OpenOptions::new().source_mode(mode);
        let file = litchi_ops::open_with(&path, &options).unwrap();
        let found = file
            .query("smartart")
            .unwrap()
            .collect::<litchi_ops::Result<Vec<_>>>()
            .unwrap();
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].path().to_string(), "/smartart[1]");
        assert_eq!(
            found[0].view().text().unwrap().as_deref(),
            Some("ACQUIRE_013\nTRANSFORM_013\nVERIFY_013")
        );
        let commit = file
            .edit(|e| {
                if node_only {
                    e.set("/smartart[1]/node[2]",props! {"text"=>"TRANSFORM_REVIEWED"})
                } else { e.set(
                    "/smartart[1]",
                    props! {"layout"=>"process","text"=>"ACQUIRE_013\nTRANSFORM_REVIEWED\nVERIFY_013"},
                ) }
            })
            .unwrap();
        let directory = tempfile::tempdir().unwrap();
        let destination = directory.path().join("smartart.docx");
        commit.file().save_as(&destination).unwrap();
        let reopened = litchi_docx::Package::open(&destination).unwrap();
        let diagrams = reopened.smart_arts().unwrap();
        assert_eq!(diagrams.len(), 1);
        assert_eq!(
            diagrams[0].text(),
            "ACQUIRE_013\nTRANSFORM_REVIEWED\nVERIFY_013"
        );
        assert_eq!(diagrams[0].extent, Some((5486400, 3200400)));
        let drawings = reopened.document().unwrap().drawing_objects().unwrap();
        assert_eq!(
            drawings.iter().map(|d| d.name()).collect::<Vec<_>>(),
            ["Graphic 1", "Graphic 3", "Diagram 2"]
        );
        assert_eq!(drawings[1].kind(), litchi_docx::drawing::Kind::Other);
        assert_eq!(
            (drawings[1].width_emu(), drawings[1].height_emu()),
            (3048000, 1714500)
        );
        assert_eq!(
            diagrams[0].diagram_type(),
            litchi_drawingml::diagram::DiagramType::Process
        );
        let after_xml =
            std::str::from_utf8(reopened.opc_package().get_part(&uri).unwrap().blob()).unwrap();
        assert!(after_xml.contains(converted_group));
        for metadata in original.opc_package().iter_parts() {
            let should_preserve = if node_only {
                !["/word/diagrams/data1.xml", "/word/diagrams/drawing1.xml"]
                    .contains(&metadata.partname().as_str())
            } else {
                metadata.partname() != &uri
                    && !metadata.partname().as_str().starts_with("/word/diagrams/")
            };
            if should_preserve {
                assert!(
                    original
                        .opc_package()
                        .get_part(metadata.partname())
                        .unwrap()
                        .blob()
                        == reopened
                            .opc_package()
                            .get_part(metadata.partname())
                            .unwrap()
                            .blob(),
                    "changed {}",
                    metadata.partname()
                );
            }
        }
        assert!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(file.to_bytes().unwrap() == bytes);
    }
}

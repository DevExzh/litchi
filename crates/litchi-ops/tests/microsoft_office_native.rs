//! Public operation contracts on native Microsoft Office specimens.
#![cfg(any(feature = "doc", feature = "docx", feature = "pptx"))]

use litchi_ops::{OpenOptions, SourceMode, props};

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

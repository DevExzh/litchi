//! Public query/path/transaction coverage of Word's retained RTF comments.
#![cfg(feature = "rtf")]

use litchi::{File, FileFormat, OpenOptions, SourceMode, props};

const ROOT_COMMENT: &str = "MODERN_COMMENT_003: review fixture comment.";
const REPLY: &str = "MODERN_REPLY_003: native reply retained.";

fn comments(file: &File, text: [&str; 2]) {
    assert_eq!(file.format(), Some(FileFormat::Rtf));
    let document = file.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 9);
    assert_eq!(document.paragraphs().unwrap().len(), 9);
    assert_eq!(
        document.paragraph_text(7usize).unwrap().as_deref(),
        Some("")
    );
    assert_eq!(
        document.paragraph(8usize).unwrap().unwrap().text().unwrap(),
        "REVIEW_CONTROLS_END"
    );
    assert!(document.paragraph(9usize).unwrap().is_none());
    assert_eq!(document.elements().unwrap().len(), 9);
    let paragraphs = file
        .query("p")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(paragraphs.len(), 9);
    assert_eq!(
        file.get("/body/p[8]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("")
    );
    assert_eq!(
        file.get("/body/p[9]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("REVIEW_CONTROLS_END")
    );
    let nodes = file
        .query("comment")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(nodes.len(), 2);
    for (index, node) in nodes.iter().enumerate() {
        let path = format!("/comments/comment[{}]", index + 1);
        assert_eq!(node.path().to_string(), path);
        let direct = file.get(path.as_str()).unwrap();
        assert_eq!(direct.path(), node.path());
        for comment in [node, &direct] {
            assert_eq!(comment.view().text().unwrap().as_deref(), Some(text[index]));
            let view = comment.view();
            let properties = view
                .format()
                .unwrap()
                .map(|p| (p.key().to_owned(), p.to_string()))
                .collect::<std::collections::BTreeMap<_, _>>();
            assert_eq!(
                properties.get("author").map(String::as_str),
                Some("Ryker Zhu")
            );
            assert_eq!(properties.get("initials").map(String::as_str), Some("RZ"));
            assert_eq!(
                properties.get("date").map(String::as_str),
                Some(["-2014676110", "-2014676109"][index])
            );
        }
    }
    assert_eq!(
        file.query("comment[author=\"Ryker Zhu\"]")
            .unwrap()
            .collect::<litchi::Result<Vec<_>>>()
            .unwrap()
            .len(),
        2
    );
}

#[test]
fn native_rtf_comments_have_stable_public_paths_and_reversible_batched_text_edits() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word/word-review-controls.rtf");
    let bytes = std::fs::read(&path).unwrap();
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let options = OpenOptions::new().source_mode(mode);
        let file = litchi::open_with(&path, &options).unwrap();
        comments(&file, [ROOT_COMMENT, REPLY]);
        let before_text = file.text().unwrap();
        let commit = file
            .edit(|edit| {
                edit.set(
                    "/comments/comment[1]",
                    props! {"text" => "PUBLIC_ROOT_EDITED"},
                )?;
                edit.set(
                    "/comments/comment[2]",
                    props! {"text" => "PUBLIC_REPLY_EDITED"},
                )
            })
            .unwrap();
        let output = commit.file().to_bytes().unwrap();
        assert_eq!(
            output,
            std::str::from_utf8(&bytes)
                .unwrap()
                .replace(ROOT_COMMENT, "PUBLIC_ROOT_EDITED")
                .replace(REPLY, "PUBLIC_REPLY_EDITED")
                .as_bytes()
        );
        let directory = tempfile::tempdir().unwrap();
        let saved = directory.path().join("review.rtf");
        commit.file().save_as(&saved).unwrap();
        assert_eq!(std::fs::read(&saved).unwrap(), output);
        let reopened = litchi::open_with(&saved, &options).unwrap();
        comments(&reopened, ["PUBLIC_ROOT_EDITED", "PUBLIC_REPLY_EDITED"]);
        assert_eq!(reopened.text().unwrap(), before_text);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(&reopened)
                .unwrap()
                .to_bytes()
                .unwrap(),
            bytes
        );
        assert_eq!(file.to_bytes().unwrap(), bytes);
    }
    assert_eq!(std::fs::read(path).unwrap(), bytes);
}

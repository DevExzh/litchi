//! Comment replies, resolution and anchors staged with main-document edits
//! (record 0345): a reply is threaded through `commentsExtended` by the
//! comment paragraphs' `w14:paraId`, `done` is the entry's `w15:done`, and
//! the range can start after a run or be a point reference. Each publishes
//! in one commit and inverts exactly.

use litchi_core::Position;
use litchi_docx::document::{CommentChange, ElementPath, ElementStep, NewComment};
use litchi_docx::{Builder, Package};

fn source() -> Package {
    let mut builder = Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("Reviewed text");
    builder.build().unwrap()
}

fn paragraph(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

fn main_xml(package: &Package) -> String {
    String::from_utf8(
        package
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
}

#[test]
fn a_reply_is_threaded_and_resolution_is_recorded() {
    let package = source();
    let mut edit = package.edit().unwrap();
    let root = edit
        .add_comment_at(&paragraph(0), None, &NewComment::new("Ada", "Check this"))
        .unwrap();
    let reply = edit
        .add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Grace", "Done").with_parent(root),
        )
        .unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    let package = commit.into_package();
    let reopened = Package::from_bytes(&package.clone().to_bytes().unwrap()).unwrap();
    let document = reopened.document().unwrap();
    let comments = document.comments().unwrap();
    let entries = document.comments_extended().unwrap();
    let root_paragraph = comments[0].paragraph_id().expect("the root has a paraId");
    let reply_paragraph = comments[1].paragraph_id().expect("the reply has a paraId");
    assert_ne!(root_paragraph, reply_paragraph);
    assert_eq!(comments[1].id(), reply);
    let find = |paragraph: u32| {
        entries
            .iter()
            .find(|entry| entry.paragraph_id == paragraph)
            .expect("an entry")
    };
    assert_eq!(find(root_paragraph).parent_paragraph_id, None);
    assert!(!find(root_paragraph).done);
    assert_eq!(
        find(reply_paragraph).parent_paragraph_id,
        Some(root_paragraph)
    );

    // Resolve the root; an equal state is a no-op.
    let mut edit = reopened.edit().unwrap();
    assert!(
        edit.set_comment(root, &CommentChange::new().with_done(true))
            .unwrap()
    );
    let resolved = edit.commit().unwrap().into_package();
    let entries = resolved.document().unwrap().comments_extended().unwrap();
    assert!(
        entries
            .iter()
            .any(|entry| entry.paragraph_id == root_paragraph && entry.done)
    );
    let mut edit = resolved.edit().unwrap();
    assert!(
        !edit
            .set_comment(root, &CommentChange::new().with_done(true))
            .unwrap()
    );

    // The inverse of the first commit removes both comments and the part.
    let restored = package.apply(&inverse).unwrap().into_package();
    assert!(restored.document().unwrap().comments().unwrap().is_empty());
    assert!(
        restored
            .document()
            .unwrap()
            .comments_extended()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn a_reply_to_an_unknown_comment_is_refused() {
    let mut edit = source().edit().unwrap();
    edit.add_comment_at(&paragraph(0), None, &NewComment::new("Ada", "x"))
        .unwrap();
    assert!(
        edit.add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Grace", "y").with_parent(42)
        )
        .is_err()
    );
}

#[test]
fn ranges_start_after_a_run_or_are_points() {
    let mut builder = Builder::new().unwrap();
    {
        let document = builder.document().unwrap();
        let paragraph = document.add_paragraph();
        paragraph.add_run_with_text("one ");
        paragraph.add_run_with_text("two");
    }
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    let ranged = edit
        .add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Ada", "after one").with_start_after(Position::new(0)),
        )
        .unwrap();
    let point = edit
        .add_comment_at(
            &paragraph(0),
            None,
            &NewComment::new("Ada", "point").with_range(false),
        )
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let xml = main_xml(&package);
    let start = xml
        .find(&format!(r#"<w:commentRangeStart w:id="{ranged}"/>"#))
        .expect("range start");
    let one = xml.find("one ").expect("first run");
    let two = xml.find(">two<").expect("second run");
    assert!(one < start && start < two, "{xml}");
    assert!(!xml.contains(&format!(r#"<w:commentRangeStart w:id="{point}"/>"#)));
    assert!(xml.contains(&format!(r#"<w:commentReference w:id="{point}"/>"#)));
}

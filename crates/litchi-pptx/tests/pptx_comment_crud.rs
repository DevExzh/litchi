#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
//! Legacy and modern comment CRUD through the package comment families
//! (record 0356a).

use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, Part};
use litchi_pptx::Package;
use litchi_pptx::comments::{
    Author as LegacyAuthor, Comment as LegacyComment, Comments as LegacyComments,
    Conformance as LegacyConformance, List as LegacyList,
};
use litchi_pptx::modern_comments::{Author, Comment, Progress, Reply};

const SLIDE: &str = "/ppt/slides/slide1.xml";
const SLIDE_POSITION: usize = 0;
const AUTHOR_A: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
const AUTHOR_B: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
const COMMENT_A: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
const COMMENT_B: &str = "{ABCDEF12-3456-4ABC-8DEF-1234567890AB}";
const REPLY_A: &str = "{E524A04C-CF22-45D7-A60D-09322EA5A80D}";

/// A one-slide deck; its slide is `SLIDE`.
fn package() -> Package {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    builder.build().unwrap()
}

fn legacy_author(id: u32, name: &str) -> LegacyAuthor {
    LegacyAuthor {
        id,
        name: name.into(),
        initials: name.chars().next().unwrap().to_string(),
        last_index: 1,
        color_index: id,
    }
}

fn legacy_comment(author_id: u32, index: u32, text: &str) -> LegacyComment {
    LegacyComment {
        author_id,
        date_time: Some("2026-07-19T12:00:00Z".into()),
        index,
        x: 100,
        y: 200,
        text: text.into(),
    }
}

fn modern_author(id: &str, name: &str) -> Author {
    Author {
        id: id.into(),
        name: name.into(),
        initials: Some(name.chars().next().unwrap().to_string()),
        user_id: format!("{}@example.test", name.to_ascii_lowercase()),
        provider_id: "local-test".into(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    }
}

fn modern_comment(id: &str, author_id: &str, title: &str) -> Comment {
    Comment {
        id: id.into(),
        author_id: author_id.into(),
        status: None,
        created: "2026-07-19T12:00:00Z".into(),
        start_date: None,
        due_date: None,
        assigned_to: None,
        complete: None,
        title: Some(title.into()),
        namespace_declarations: Vec::new(),
        anchors: Vec::new(),
        position: None,
        reply_list_namespace_declarations: Vec::new(),
        replies: Vec::new(),
        reply_list_present: false,
        text_body_xml: None,
        extension_xml: None,
    }
}

fn modern_reply(id: &str, author_id: &str) -> Reply {
    Reply {
        id: id.into(),
        author_id: author_id.into(),
        status: None,
        created: "2026-07-19T12:01:00Z".into(),
        namespace_declarations: Vec::new(),
        text_body_xml: None,
        extension_xml: None,
    }
}

#[test]
fn legacy_author_and_slide_comment_crud_preserves_shared_target() {
    let package = package();
    let mut edit = package.edit_comments().unwrap();
    let graph = LegacyComments {
        author_relationship_id: "rIdLegacyAuthors".into(),
        author_part_name: "/ppt/commentAuthors.xml".into(),
        authors: vec![legacy_author(1, "Ada")],
        slides: vec![LegacyList {
            slide_part_name: SLIDE.into(),
            relationship_id: "rIdLegacyComments".into(),
            part_name: "/ppt/comments/comment1.xml".into(),
            comments: vec![legacy_comment(1, 1, "first")],
        }],
    };
    edit.set(&graph, LegacyConformance::Transitional).unwrap();

    edit.add_author(legacy_author(2, "Grace"), LegacyConformance::Transitional)
        .unwrap();
    edit.add(
        SLIDE_POSITION,
        legacy_comment(2, 1, "second"),
        LegacyConformance::Transitional,
    )
    .unwrap();
    let mut changed = legacy_comment(2, 1, "updated");
    changed.x = -50;
    edit.update(
        SLIDE_POSITION,
        2,
        1,
        changed.clone(),
        LegacyConformance::Strict,
    )
    .unwrap();
    edit.update(
        SLIDE_POSITION,
        2,
        1,
        changed,
        LegacyConformance::Transitional,
    )
    .unwrap();
    edit.reorder(
        SLIDE_POSITION,
        &[(2, 1), (1, 1)],
        LegacyConformance::Transitional,
    )
    .unwrap();
    edit.reorder_authors(&[2, 1], LegacyConformance::Transitional)
        .unwrap();
    let mut grace = edit.author(2).unwrap().unwrap();
    grace.name = "Grace Hopper".into();
    edit.update_author(2, grace.clone(), LegacyConformance::Transitional)
        .unwrap();
    edit.update_author(2, grace, LegacyConformance::Transitional)
        .unwrap();
    assert_eq!(
        edit.comment(SLIDE_POSITION, 2, 1).unwrap().unwrap().text,
        "updated"
    );
    assert!(
        edit.remove_author(2, LegacyConformance::Transitional)
            .is_err()
    );

    let package = edit.commit().unwrap().into_package();
    let comment_part = PackURI::new("/ppt/comments/comment1.xml").unwrap();
    let mut raw = package.edit_opc().unwrap();
    raw.update(|opc| {
        let mut shared_owner = BlobPart::new(
            PackURI::new("/ppt/shared-owner.xml").unwrap(),
            "application/xml".into(),
            Vec::new(),
        );
        shared_owner.relate_to("comments/comment1.xml", "urn:test:shared");
        opc.add_part(Box::new(shared_owner));
        Ok(())
    })
    .unwrap();
    let package = raw.commit().unwrap().into_package();
    let mut edit = package.edit_comments().unwrap();
    assert!(
        edit.remove(SLIDE_POSITION, 2, 1, LegacyConformance::Transitional)
            .unwrap()
    );
    assert!(
        edit.remove_author(2, LegacyConformance::Transitional)
            .unwrap()
    );
    assert!(
        edit.remove(SLIDE_POSITION, 1, 1, LegacyConformance::Transitional)
            .unwrap()
    );
    let package = edit.commit().unwrap().into_package();
    assert!(package.opc().unwrap().get_part(&comment_part).is_ok());
}

#[test]
fn modern_author_comment_and_reply_crud_is_graph_checked() {
    let package = package();
    let slide = SLIDE_POSITION;
    let mut edit = package.edit_modern_comments().unwrap();
    edit.add_author(modern_author(AUTHOR_A, "Ada")).unwrap();
    edit.add_author(modern_author(AUTHOR_B, "Grace")).unwrap();
    assert!(
        edit.add(
            slide,
            modern_comment(COMMENT_A, "{11111111-1111-4111-8111-111111111111}", "bad")
        )
        .is_err()
    );
    assert!(edit.comment(slide, COMMENT_A).unwrap().is_none());

    edit.add(slide, modern_comment(COMMENT_A, AUTHOR_A, "first"))
        .unwrap();
    edit.add(slide, modern_comment(COMMENT_B, AUTHOR_B, "second"))
        .unwrap();
    assert!(
        edit.add(slide, modern_comment(COMMENT_A, AUTHOR_A, "duplicate"))
            .is_err()
    );
    edit.add_reply(slide, COMMENT_A, modern_reply(REPLY_A, AUTHOR_B))
        .unwrap();
    assert!(
        edit.add_reply(slide, COMMENT_B, modern_reply(REPLY_A, AUTHOR_A))
            .is_err()
    );
    edit.update(slide, COMMENT_A, |comment| {
        comment.title = Some("updated".into());
        comment.assigned_to = Some(vec![AUTHOR_B.into()]);
        comment.complete = Some(Progress::new(75).unwrap());
    })
    .unwrap();
    assert_eq!(
        edit.comment(slide, COMMENT_A).unwrap().unwrap().complete,
        Some(Progress::new(75).unwrap())
    );
    edit.update_reply(slide, COMMENT_A, REPLY_A, |reply| {
        reply.created = "2026-07-19T12:02:00Z".into();
    })
    .unwrap();
    let replacement_reply = modern_reply(REPLY_A, AUTHOR_B);
    edit.replace_reply(slide, COMMENT_A, REPLY_A, replacement_reply)
        .unwrap();
    let mut replacement = edit.comment(slide, COMMENT_B).unwrap().unwrap();
    replacement.title = Some("replacement".into());
    edit.replace(slide, COMMENT_B, replacement).unwrap();
    edit.reorder(slide, &[COMMENT_B.into(), COMMENT_A.into()])
        .unwrap();
    edit.reorder_authors(&[AUTHOR_B.into(), AUTHOR_A.into()])
        .unwrap();
    edit.update_author(AUTHOR_A, |author| {
        author.name = "Ada Lovelace".into();
    })
    .unwrap();
    let mut ada = modern_author(AUTHOR_A, "Ada Lovelace");
    ada.initials = Some("AL".into());
    edit.replace_author(AUTHOR_A, ada).unwrap();
    assert_eq!(
        edit.reply(slide, COMMENT_A, REPLY_A)
            .unwrap()
            .unwrap()
            .author_id,
        AUTHOR_B
    );
    assert!(edit.remove_author(AUTHOR_B).is_err());
    assert!(edit.remove_reply(slide, COMMENT_A, REPLY_A).unwrap());
    edit.update(slide, COMMENT_A, |comment| {
        comment.assigned_to = None;
    })
    .unwrap();
    assert!(edit.remove(slide, COMMENT_B).unwrap());
    assert!(edit.remove_author(AUTHOR_B).unwrap());
    let committed = edit.commit().unwrap();
    assert!(committed.diagnostics().changed());
}

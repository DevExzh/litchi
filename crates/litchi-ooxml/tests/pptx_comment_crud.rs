use litchi_ooxml::pptx::comments::{
    PresentationComment, PresentationCommentAuthor, PresentationCommentConformance,
    PresentationComments, SlideCommentList,
};
use litchi_ooxml::pptx::modern_comment_authors::ModernCommentAuthor;
use litchi_ooxml::pptx::modern_comments::{ModernComment, ModernCommentReply};
use litchi_ooxml::pptx::{
    add_modern_comment, add_modern_comment_author, add_modern_comment_reply,
    add_presentation_comment, add_presentation_comment_author, find_modern_comment,
    find_modern_comment_reply, find_presentation_comment, find_presentation_comment_author,
    remove_modern_comment, remove_modern_comment_author, remove_modern_comment_reply,
    remove_presentation_comment, remove_presentation_comment_author,
    reorder_modern_comment_authors, reorder_modern_comments,
    reorder_presentation_comment_authors, reorder_presentation_comments,
    replace_modern_comment, replace_modern_comment_author, replace_modern_comment_reply,
    replace_presentation_comment, replace_presentation_comment_author,
    store_presentation_comments, update_modern_comment, update_modern_comment_author,
    update_modern_comment_reply, update_presentation_comment,
    update_presentation_comment_author,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const SLIDE: &str = "/ppt/slides/slide1.xml";
const AUTHOR_A: &str = "{CD37207E-7903-4ED4-8AE8-017538D2DF7E}";
const AUTHOR_B: &str = "{0B2043D4-0908-4C42-8A79-51EA2CC309F7}";
const COMMENT_A: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";
const COMMENT_B: &str = "{ABCDEF12-3456-4ABC-8DEF-1234567890AB}";
const REPLY_A: &str = "{E524A04C-CF22-45D7-A60D-09322EA5A80D}";

fn package() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new(SLIDE).unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.into(),
        br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#.to_vec(),
    );
    presentation.relate_to("slides/slide1.xml", rt::SLIDE);
    package.add_part(Box::new(presentation));
    package.add_part(Box::new(BlobPart::new(
        slide_name.clone(),
        ct::PML_SLIDE.into(),
        br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#.to_vec(),
    )));
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    (package, slide_name)
}

fn legacy_author(id: u32, name: &str) -> PresentationCommentAuthor {
    PresentationCommentAuthor {
        id,
        name: name.into(),
        initials: name.chars().next().unwrap().to_string(),
        last_index: 1,
        color_index: id,
    }
}

fn legacy_comment(author_id: u32, index: u32, text: &str) -> PresentationComment {
    PresentationComment {
        author_id,
        date_time: Some("2026-07-19T12:00:00Z".into()),
        index,
        x: 100,
        y: 200,
        text: text.into(),
    }
}

fn modern_author(id: &str, name: &str) -> ModernCommentAuthor {
    ModernCommentAuthor {
        id: id.into(),
        name: name.into(),
        initials: Some(name.chars().next().unwrap().to_string()),
        user_id: format!("{}@example.test", name.to_ascii_lowercase()),
        provider_id: "local-test".into(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    }
}

fn modern_comment(id: &str, author_id: &str, title: &str) -> ModernComment {
    ModernComment {
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

fn modern_reply(id: &str, author_id: &str) -> ModernCommentReply {
    ModernCommentReply {
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
    let (mut package, _) = package();
    let graph = PresentationComments {
        author_relationship_id: "rIdLegacyAuthors".into(),
        author_part_name: "/ppt/commentAuthors.xml".into(),
        authors: vec![legacy_author(1, "Ada")],
        slides: vec![SlideCommentList {
            slide_part_name: SLIDE.into(),
            relationship_id: "rIdLegacyComments".into(),
            part_name: "/ppt/comments/comment1.xml".into(),
            comments: vec![legacy_comment(1, 1, "first")],
        }],
    };
    store_presentation_comments(
        &mut package,
        &graph,
        PresentationCommentConformance::Transitional,
    )
    .unwrap();

    add_presentation_comment_author(
        &mut package,
        legacy_author(2, "Grace"),
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    add_presentation_comment(
        &mut package,
        SLIDE,
        legacy_comment(2, 1, "second"),
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    let mut changed = legacy_comment(2, 1, "updated");
    changed.x = -50;
    update_presentation_comment(
        &mut package,
        SLIDE,
        2,
        1,
        changed.clone(),
        PresentationCommentConformance::Strict,
    )
    .unwrap();
    replace_presentation_comment(
        &mut package,
        SLIDE,
        2,
        1,
        changed,
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    reorder_presentation_comments(
        &mut package,
        SLIDE,
        &[(2, 1), (1, 1)],
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    reorder_presentation_comment_authors(
        &mut package,
        &[2, 1],
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    let mut grace = find_presentation_comment_author(&package, 2).unwrap().unwrap();
    grace.name = "Grace Hopper".into();
    update_presentation_comment_author(
        &mut package,
        2,
        grace.clone(),
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    replace_presentation_comment_author(
        &mut package,
        2,
        grace,
        PresentationCommentConformance::Transitional,
    )
    .unwrap();
    assert_eq!(
        find_presentation_comment(&package, SLIDE, 2, 1)
            .unwrap()
            .unwrap()
            .text,
        "updated"
    );
    assert!(remove_presentation_comment_author(
        &mut package,
        2,
        PresentationCommentConformance::Transitional
    )
    .is_err());

    let comment_part = PackURI::new("/ppt/comments/comment1.xml").unwrap();
    let mut shared_owner = BlobPart::new(
        PackURI::new("/ppt/shared-owner.xml").unwrap(),
        "application/xml".into(),
        Vec::new(),
    );
    shared_owner.relate_to("comments/comment1.xml", "urn:test:shared");
    package.add_part(Box::new(shared_owner));
    assert!(remove_presentation_comment(
        &mut package,
        SLIDE,
        2,
        1,
        PresentationCommentConformance::Transitional
    )
    .unwrap());
    assert!(remove_presentation_comment_author(
        &mut package,
        2,
        PresentationCommentConformance::Transitional
    )
    .unwrap());
    assert!(remove_presentation_comment(
        &mut package,
        SLIDE,
        1,
        1,
        PresentationCommentConformance::Transitional
    )
    .unwrap());
    assert!(package.get_part(&comment_part).is_ok());
}

#[test]
fn modern_author_comment_and_reply_crud_is_graph_checked() {
    let (mut package, slide) = package();
    add_modern_comment_author(&mut package, modern_author(AUTHOR_A, "Ada")).unwrap();
    add_modern_comment_author(&mut package, modern_author(AUTHOR_B, "Grace")).unwrap();
    assert!(add_modern_comment(
        &mut package,
        &slide,
        modern_comment(COMMENT_A, "{11111111-1111-4111-8111-111111111111}", "bad")
    )
    .is_err());
    assert!(find_modern_comment(&package, &slide, COMMENT_A).unwrap().is_none());

    add_modern_comment(
        &mut package,
        &slide,
        modern_comment(COMMENT_A, AUTHOR_A, "first"),
    )
    .unwrap();
    add_modern_comment(
        &mut package,
        &slide,
        modern_comment(COMMENT_B, AUTHOR_B, "second"),
    )
    .unwrap();
    assert!(add_modern_comment(
        &mut package,
        &slide,
        modern_comment(COMMENT_A, AUTHOR_A, "duplicate")
    )
    .is_err());
    add_modern_comment_reply(
        &mut package,
        &slide,
        COMMENT_A,
        modern_reply(REPLY_A, AUTHOR_B),
    )
    .unwrap();
    assert!(add_modern_comment_reply(
        &mut package,
        &slide,
        COMMENT_B,
        modern_reply(REPLY_A, AUTHOR_A)
    )
    .is_err());
    update_modern_comment(&mut package, &slide, COMMENT_A, |comment| {
        comment.title = Some("updated".into());
        comment.assigned_to = Some(vec![AUTHOR_B.into()]);
    })
    .unwrap();
    update_modern_comment_reply(&mut package, &slide, COMMENT_A, REPLY_A, |reply| {
        reply.created = "2026-07-19T12:02:00Z".into();
    })
    .unwrap();
    let replacement_reply = modern_reply(REPLY_A, AUTHOR_B);
    replace_modern_comment_reply(
        &mut package,
        &slide,
        COMMENT_A,
        REPLY_A,
        replacement_reply,
    )
    .unwrap();
    let mut replacement = find_modern_comment(&package, &slide, COMMENT_B)
        .unwrap()
        .unwrap();
    replacement.title = Some("replacement".into());
    replace_modern_comment(&mut package, &slide, COMMENT_B, replacement).unwrap();
    reorder_modern_comments(
        &mut package,
        &slide,
        &[COMMENT_B.into(), COMMENT_A.into()],
    )
    .unwrap();
    reorder_modern_comment_authors(&mut package, &[AUTHOR_B.into(), AUTHOR_A.into()]).unwrap();
    update_modern_comment_author(&mut package, AUTHOR_A, |author| {
        author.name = "Ada Lovelace".into();
    })
    .unwrap();
    let mut ada = modern_author(AUTHOR_A, "Ada Lovelace");
    ada.initials = Some("AL".into());
    replace_modern_comment_author(&mut package, AUTHOR_A, ada).unwrap();
    assert_eq!(
        find_modern_comment_reply(&package, &slide, COMMENT_A, REPLY_A)
            .unwrap()
            .unwrap()
            .author_id,
        AUTHOR_B
    );
    assert!(remove_modern_comment_author(&mut package, AUTHOR_B).is_err());
    assert!(remove_modern_comment_reply(&mut package, &slide, COMMENT_A, REPLY_A).unwrap());
    update_modern_comment(&mut package, &slide, COMMENT_A, |comment| {
        comment.assigned_to = None;
    })
    .unwrap();
    assert!(remove_modern_comment(&mut package, &slide, COMMENT_B).unwrap());
    assert!(remove_modern_comment_author(&mut package, AUTHOR_B).unwrap());
}

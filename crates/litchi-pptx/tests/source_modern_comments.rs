//! Source-backed modern comments.
//!
//! The fixture keeps the comments part on one selected slide and puts a
//! deliberately malformed part on another slide.  This makes the tests
//! exercise the source owner's relationship locality as well as the detached
//! semantic projection.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the in-memory OPC fixture and assertions intentionally panic"
)]

use std::io::{self, Cursor};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceLimitKind, TargetMode};
use litchi_pptx::modern_comments::{
    MODERN_COMMENT_AUTHOR_CONTENT_TYPE, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE,
    MODERN_COMMENT_CONTENT_TYPE, MODERN_COMMENT_RELATIONSHIP_TYPE,
};
use litchi_pptx::{Error, Package, SourceBackedPresentation, SourceLimits};

const PRESENTATION: &str = "/ppt/presentation.xml";
const SLIDE_TWO: &str = "/ppt/slides/slide2.xml";
const SLIDE_THREE: &str = "/ppt/slides/slide3.xml";
const AUTHORS: &str = "/ppt/authors.xml";
const MODERN_A: &str = "/ppt/comments/modernComment_257_A.xml";
const MODERN_C: &str = "/ppt/comments/modernComment_258_C.xml";

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PC: &str = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const P188: &str = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const ADA: &str = "{A1B2C3D4-0000-4000-8000-000000000001}";
const GRACE: &str = "{A1B2C3D4-0000-4000-8000-000000000002}";
const COMMENT_ONE: &str = "{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C01}";
const COMMENT_TWO: &str = "{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C02}";
const COMMENT_THREE: &str = "{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C03}";
const REPLY_ONE: &str = "{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C11}";

fn authors_xml() -> String {
    format!(
        r#"<p188:authorLst xmlns:a="{A}" xmlns:p188="{P188}"><p188:author id="{ADA}" name="Ada Lovelace" initials="AL" userId="ada@example.com" providerId="None"/><p188:author id="{GRACE}" name="Grace Hopper" initials="GH" userId="grace@example.com" providerId="None"/></p188:authorLst>"#
    )
}

/// `a` is declared only on the list root.  Bodies therefore exercise the
/// namespace inheritance that real PowerPoint files use.
fn modern_list(comments: &str) -> String {
    format!(
        r#"<p188:cmLst xmlns:a="{A}" xmlns:mc="{MC}" xmlns:pc="{PC}" xmlns:p188="{P188}" xmlns:x="urn:source-modern-comments-test">{comments}</p188:cmLst>"#
    )
}

fn paragraph(text: &str) -> String {
    format!(r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>"#)
}

fn reply(id: &str, author: &str, body: &str) -> String {
    format!(
        concat!(
            r#"<p188:reply id="{id}" authorId="{author}" created="2026-10-01T10:00:00.000">"#,
            r#"<p188:txBody><a:bodyPr/><a:lstStyle/>{body}</p188:txBody></p188:reply>"#,
        ),
        id = id,
        author = author,
        body = body,
    )
}

fn modern_comment(id: &str, author: &str, body: Option<&str>, replies: &str) -> String {
    let text_body = body.map_or_else(String::new, |body| {
        format!(r#"<p188:txBody><a:bodyPr/><a:lstStyle/>{body}</p188:txBody>"#)
    });
    let reply_list = if replies.is_empty() {
        String::new()
    } else {
        format!(r#"<p188:replyLst>{replies}</p188:replyLst>"#)
    };
    format!(
        concat!(
            r#"<p188:cm id="{id}" authorId="{author}" created="2026-10-01T09:30:00.000">"#,
            r#"<pc:sldMkLst><pc:docMk/><pc:sldMk cId="1" sldId="257"/></pc:sldMkLst>"#,
            r#"<p188:pos x="10" y="20"/>{reply_list}{text_body}</p188:cm>"#,
        ),
        id = id,
        author = author,
        reply_list = reply_list,
        text_body = text_body,
    )
}

fn add_part(opc: &mut OpcPackage, name: &str, content_type: &str, xml: impl Into<Vec<u8>>) {
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new(name).unwrap(),
        content_type.into(),
        xml.into(),
    )))
    .unwrap();
}

fn relate(opc: &mut OpcPackage, source: &str, kind: &str, target: &str, id: &str) {
    opc.get_part_mut(&PackURI::new(source).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(kind.into(), target.into(), id.into(), TargetMode::Internal)
        .unwrap();
}

fn deck() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for title in ["Overview", "Comments", "Malformed elsewhere"] {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .set_title(title);
    }
    let package = builder.build().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();

    add_part(
        &mut opc,
        AUTHORS,
        MODERN_COMMENT_AUTHOR_CONTENT_TYPE,
        authors_xml(),
    );
    relate(
        &mut opc,
        PRESENTATION,
        MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE,
        "authors.xml",
        "rIdModernAuthors",
    );

    let comments = modern_list(&format!(
        "{}{}{}",
        modern_comment(
            COMMENT_ONE,
            ADA,
            Some(&format!(
                "{}{}",
                paragraph("Check the Q3 numbers"),
                paragraph("and the chart")
            )),
            &reply(REPLY_ONE, GRACE, &paragraph("Agreed &amp; done")),
        ),
        // Omitted txBody and an empty txBody are both intentional.  The
        // projection must expose an empty semantic string for either form.
        modern_comment(COMMENT_TWO, GRACE, None, ""),
        modern_comment(COMMENT_THREE, ADA, Some(r#"<a:p/>"#), "",),
    ));
    add_part(&mut opc, MODERN_A, MODERN_COMMENT_CONTENT_TYPE, comments);
    relate(
        &mut opc,
        SLIDE_TWO,
        MODERN_COMMENT_RELATIONSHIP_TYPE,
        "../comments/modernComment_257_A.xml",
        "rIdModernComments",
    );

    add_part(
        &mut opc,
        MODERN_C,
        MODERN_COMMENT_CONTENT_TYPE,
        modern_list(&modern_comment(
            "{6C7C3B3A-0D1F-4C3B-9D5B-2F0E4A1B7C04}",
            ADA,
            Some(&paragraph("Unrelated slide")),
            "",
        )),
    );
    relate(
        &mut opc,
        SLIDE_THREE,
        MODERN_COMMENT_RELATIONSHIP_TYPE,
        "../comments/modernComment_258_C.xml",
        "rIdOtherComments",
    );

    PackageWriter::to_bytes(&opc).unwrap()
}

fn rewrite(bytes: &[u8], edit: impl FnOnce(&mut OpcPackage)) -> Vec<u8> {
    let mut opc = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    edit(&mut opc);
    PackageWriter::to_bytes(&opc).unwrap()
}

fn edit_blob(opc: &mut OpcPackage, part_name: &str, edit: impl FnOnce(String) -> String) {
    let uri = PackURI::new(part_name).unwrap();
    let part = opc.get_part_mut(&uri).unwrap();
    let xml = String::from_utf8(part.blob().to_vec()).unwrap();
    part.set_blob(edit(xml).into_bytes());
}

fn rebind_reply_list_drawingml(xml: String) -> String {
    let start = xml.find("<p188:replyLst>").unwrap();
    let relative_end = xml[start..].find("</p188:replyLst>").unwrap();
    let end = start + relative_end + "</p188:replyLst>".len();
    let mut reply_list = xml[start..end].to_owned();
    reply_list = reply_list.replace(
        "<p188:replyLst>",
        &format!(r#"<p188:replyLst xmlns:d="{A}">"#),
    );
    reply_list = reply_list.replace("<a:", "<d:").replace("</a:", "</d:");
    format!("{}{}{}", &xml[..start], reply_list, &xml[end..])
}

fn open(bytes: &[u8]) -> SourceBackedPresentation {
    SourceBackedPresentation::from_reader(Cursor::new(bytes)).unwrap()
}

fn modern_error(bytes: &[u8]) -> Error {
    match SourceBackedPresentation::from_reader(Cursor::new(bytes))
        .and_then(|source| source.slide(1).unwrap().modern_comments().map(drop))
    {
        Ok(()) => panic!("the malformed modern-comment graph was accepted"),
        Err(error) => error,
    }
}

fn assert_modern_refused(bytes: &[u8], label: &str) {
    let error = modern_error(bytes);
    assert!(
        matches!(
            error,
            Error::Relationship(_) | Error::Invalid(_) | Error::ContentType { .. } | Error::Opc(_)
        ),
        "{label}: unexpected error category: {error:?}"
    );
}

fn assert_modern_error_redacted(bytes: &[u8], authored_token: &str, label: &str) {
    let error = modern_error(bytes);
    assert!(
        !error.to_string().contains(authored_token),
        "{label}: Display leaked authored token: {error}"
    );
    assert!(
        !format!("{error:?}").contains(authored_token),
        "{label}: Debug leaked authored token: {error:?}"
    );
}

#[test]
fn source_modern_comments_detach_text_and_inherit_drawingml_namespace() {
    let list = {
        let source = open(&deck());
        assert!(
            source
                .slide(0)
                .unwrap()
                .modern_comments()
                .unwrap()
                .is_none()
        );
        source.slide(1).unwrap().modern_comments().unwrap().unwrap()
    };

    assert_eq!(list.entries().len(), 3);
    let first = &list.entries()[0];
    assert_eq!(first.text(), Some("Check the Q3 numbers\nand the chart"));
    assert_eq!(first.replies()[0].text(), Some("Agreed & done"));
    assert_eq!(first.author_name(), Some("Ada Lovelace"));
    assert_eq!(first.replies()[0].author_name(), Some("Grace Hopper"));
    let raw = litchi_pptx::presentation::modern_comments::raw::data(&list);
    assert!(!raw.is_empty());
    assert_eq!(
        litchi_pptx::presentation::modern_comments::raw::comment_id(first),
        COMMENT_ONE
    );
    assert_eq!(
        litchi_pptx::presentation::modern_comments::raw::reply_id(&first.replies()[0]),
        REPLY_ONE
    );

    // The source owner is gone, while the semantic values remain available.
    // `None` preserves an omitted body and `Some("")` preserves an empty one.
    assert_eq!(list.entries()[1].text(), None);
    assert_eq!(list.entries()[2].text(), Some(""));

    // Ordinary diagnostics contain shape information only.  Authored text,
    // names and package identity belong behind explicit raw accessors.
    let debug = format!("{list:?} {first:?}");
    for authored in [
        "Check the Q3 numbers",
        "Agreed & done",
        AUTHORS,
        MODERN_A,
        "rIdModernComments",
        "Ada Lovelace",
    ] {
        assert!(
            !debug.contains(authored),
            "debug leaked {authored}: {debug}"
        );
    }
}

#[test]
fn eager_modern_comments_match_the_selected_source_projection() {
    let bytes = deck();
    let eager_package = Package::from_reader(Cursor::new(bytes.clone())).unwrap();
    let eager = eager_package
        .presentation()
        .unwrap()
        .slide(1)
        .unwrap()
        .unwrap()
        .modern_comments()
        .unwrap()
        .unwrap();
    let source_package = open(&bytes);
    let source = source_package
        .slide(1)
        .unwrap()
        .modern_comments()
        .unwrap()
        .unwrap();

    assert_eq!(eager.entries().len(), source.entries().len());
    assert_eq!(
        litchi_pptx::presentation::modern_comments::raw::data(&eager),
        litchi_pptx::presentation::modern_comments::raw::data(&source)
    );
    assert_eq!(
        litchi_pptx::presentation::modern_comments::raw::author_data(&eager),
        litchi_pptx::presentation::modern_comments::raw::author_data(&source)
    );
    for (eager, source) in eager.entries().iter().zip(source.entries()) {
        assert_eq!(eager.text(), source.text());
        assert_eq!(eager.author_name(), source.author_name());
        assert_eq!(eager.created(), source.created());
        assert_eq!(eager.start_date(), source.start_date());
        assert_eq!(eager.due_date(), source.due_date());
        assert_eq!(eager.status(), source.status());
        assert_eq!(eager.progress(), source.progress());
        assert_eq!(eager.title(), source.title());
        assert_eq!(eager.location(), source.location());
        assert_eq!(eager.replies().len(), source.replies().len());
        for (eager, source) in eager.replies().iter().zip(source.replies()) {
            assert_eq!(eager.text(), source.text());
            assert_eq!(eager.author_name(), source.author_name());
            assert_eq!(eager.created(), source.created());
            assert_eq!(eager.status(), source.status());
        }
    }
}

#[test]
fn eager_modern_comments_read_only_the_selected_graph() {
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_C, |xml| {
            xml.replace("</p188:cmLst>", "<p188:bad/></p188:cmLst>")
        });
    });
    let package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    assert!(
        presentation
            .slide(1)
            .unwrap()
            .unwrap()
            .modern_comments()
            .unwrap()
            .is_some()
    );
    assert!(
        presentation
            .slide(2)
            .unwrap()
            .unwrap()
            .modern_comments()
            .is_err()
    );
}

#[test]
fn source_modern_comments_honor_reply_list_drawingml_namespace_scope() {
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_A, rebind_reply_list_drawingml);
    });
    let source = open(&bytes);
    let list = source.slide(1).unwrap().modern_comments().unwrap().unwrap();
    assert_eq!(list.entries()[0].replies()[0].text(), Some("Agreed & done"));
}

#[test]
fn source_modern_comments_refuse_external_duplicate_and_unresolved_graph_edges() {
    let cases: Vec<(&str, Box<dyn FnOnce(&mut OpcPackage)>)> = vec![
        (
            "missing comments part",
            Box::new(|opc: &mut OpcPackage| {
                assert!(opc.remove_part(&PackURI::new(MODERN_A).unwrap()));
            }),
        ),
        (
            "wrong comments content type",
            Box::new(|opc: &mut OpcPackage| {
                opc.get_part_mut(&PackURI::new(MODERN_A).unwrap())
                    .unwrap()
                    .set_content_type(ct::PML_COMMENTS.into())
                    .unwrap();
            }),
        ),
        (
            "comments part has outbound relationship",
            Box::new(|opc: &mut OpcPackage| {
                opc.get_part_mut(&PackURI::new(MODERN_A).unwrap())
                    .unwrap()
                    .rels_mut()
                    .try_add_relationship(
                        rt::IMAGE.into(),
                        "../slides/slide1.xml".into(),
                        "rIdImage".into(),
                        TargetMode::Internal,
                    )
                    .unwrap();
            }),
        ),
        (
            "external slide relationship",
            Box::new(|opc: &mut OpcPackage| {
                let slide = PackURI::new(SLIDE_TWO).unwrap();
                let rels = opc.get_part_mut(&slide).unwrap().rels_mut();
                rels.remove("rIdModernComments").unwrap();
                rels.try_add_relationship(
                    MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                    "https://example.invalid/comments.xml".into(),
                    "rIdModernComments".into(),
                    TargetMode::External,
                )
                .unwrap();
            }),
        ),
        (
            "duplicate slide relationship",
            Box::new(|opc: &mut OpcPackage| {
                let slide = PackURI::new(SLIDE_TWO).unwrap();
                opc.get_part_mut(&slide)
                    .unwrap()
                    .rels_mut()
                    .try_add_relationship(
                        MODERN_COMMENT_RELATIONSHIP_TYPE.into(),
                        "../comments/modernComment_257_A.xml".into(),
                        "rIdModernComments2".into(),
                        TargetMode::Internal,
                    )
                    .unwrap();
            }),
        ),
        (
            "missing author part",
            Box::new(|opc: &mut OpcPackage| {
                assert!(opc.remove_part(&PackURI::new(AUTHORS).unwrap()));
            }),
        ),
        (
            "wrong author content type",
            Box::new(|opc: &mut OpcPackage| {
                opc.get_part_mut(&PackURI::new(AUTHORS).unwrap())
                    .unwrap()
                    .set_content_type(ct::PML_COMMENTS.into())
                    .unwrap();
            }),
        ),
        (
            "external author relationship",
            Box::new(|opc: &mut OpcPackage| {
                let presentation = PackURI::new(PRESENTATION).unwrap();
                let rels = opc.get_part_mut(&presentation).unwrap().rels_mut();
                rels.remove("rIdModernAuthors").unwrap();
                rels.try_add_relationship(
                    MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                    "https://example.invalid/authors.xml".into(),
                    "rIdModernAuthors".into(),
                    TargetMode::External,
                )
                .unwrap();
            }),
        ),
        (
            "duplicate author relationship",
            Box::new(|opc: &mut OpcPackage| {
                opc.get_part_mut(&PackURI::new(PRESENTATION).unwrap())
                    .unwrap()
                    .rels_mut()
                    .try_add_relationship(
                        MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.into(),
                        "authors.xml".into(),
                        "rIdModernAuthors2".into(),
                        TargetMode::Internal,
                    )
                    .unwrap();
            }),
        ),
        (
            "author part has outbound relationship",
            Box::new(|opc: &mut OpcPackage| {
                opc.get_part_mut(&PackURI::new(AUTHORS).unwrap())
                    .unwrap()
                    .rels_mut()
                    .try_add_relationship(
                        rt::IMAGE.into(),
                        "../slides/slide1.xml".into(),
                        "rIdImage".into(),
                        TargetMode::Internal,
                    )
                    .unwrap();
            }),
        ),
        (
            "duplicate author ID",
            Box::new(|opc: &mut OpcPackage| {
                edit_blob(opc, AUTHORS, |xml| xml.replace(GRACE, ADA));
            }),
        ),
        (
            "duplicate comment ID",
            Box::new(|opc: &mut OpcPackage| {
                edit_blob(opc, MODERN_A, |xml| {
                    xml.replacen(COMMENT_THREE, COMMENT_ONE, 1)
                });
            }),
        ),
        (
            "unresolved comment author",
            Box::new(|opc: &mut OpcPackage| {
                edit_blob(opc, MODERN_A, |xml| {
                    xml.replacen(
                        &format!("authorId=\"{ADA}\""),
                        "authorId=\"{A1B2C3D4-0000-4000-8000-000000000099}\"",
                        1,
                    )
                });
            }),
        ),
        (
            "unresolved reply author",
            Box::new(|opc: &mut OpcPackage| {
                edit_blob(opc, MODERN_A, |xml| {
                    xml.replace(
                        &format!(r#"<p188:reply id="{REPLY_ONE}" authorId="{GRACE}""#),
                        &format!(
                            r#"<p188:reply id="{REPLY_ONE}" authorId="{{A1B2C3D4-0000-4000-8000-000000000099}}""#
                        ),
                    )
                });
            }),
        ),
        (
            "unresolved assignment author",
            Box::new(|opc: &mut OpcPackage| {
                edit_blob(opc, MODERN_A, |xml| {
                    xml.replace(
                        &format!(r#"<p188:cm id="{COMMENT_ONE}" authorId="{ADA}""#),
                        &format!(
                            r#"<p188:cm id="{COMMENT_ONE}" authorId="{ADA}" assignedTo="{{A1B2C3D4-0000-4000-8000-000000000099}}""#
                        ),
                    )
                });
            }),
        ),
    ];

    for (label, edit) in cases {
        let bytes = rewrite(&deck(), edit);
        assert_modern_refused(&bytes, label);
    }
}

#[test]
fn source_modern_comments_redact_authored_guid_and_token_diagnostics() {
    let malformed_comment_guid = "authored-comment-guid-secret";
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_A, |xml| {
            xml.replace(COMMENT_ONE, malformed_comment_guid)
        });
    });
    assert_modern_error_redacted(&bytes, malformed_comment_guid, "malformed comment GUID");

    let malformed_status = "authored-status-token-secret";
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_A, |xml| {
            xml.replace(
                &format!(r#"<p188:cm id="{COMMENT_ONE}" authorId="{ADA}""#),
                &format!(
                    r#"<p188:cm id="{COMMENT_ONE}" authorId="{ADA}" status="{malformed_status}""#
                ),
            )
        });
    });
    assert_modern_error_redacted(&bytes, malformed_status, "malformed status token");

    let malformed_author_guid = "authored-author-guid-secret";
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, AUTHORS, |xml| xml.replace(ADA, malformed_author_guid));
    });
    assert_modern_error_redacted(&bytes, malformed_author_guid, "malformed author GUID");
}

#[test]
fn source_modern_comments_read_selected_slide_without_decoding_unrelated_malformed_payload() {
    let bytes = rewrite(&deck(), |opc| {
        opc.get_part_mut(&PackURI::new(MODERN_C).unwrap())
            .unwrap()
            .set_blob(modern_list("<p188:bad/>").into_bytes());
    });
    let source = open(&bytes);
    let list = source.slide(1).unwrap().modern_comments().unwrap().unwrap();
    assert_eq!(
        list.entries()[0].text(),
        Some("Check the Q3 numbers\nand the chart")
    );
    assert!(source.slide(2).unwrap().modern_comments().is_err());
}

#[test]
fn source_modern_comments_honor_caller_part_limits_before_decode() {
    let padding = "x".repeat(64 * 1024);
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_A, |xml| {
            xml.replace("</p188:cmLst>", &format!("<!--{padding}--></p188:cmLst>"))
        });
    });
    let opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let presentation_bytes = opc
        .get_part(&PackURI::new(PRESENTATION).unwrap())
        .unwrap()
        .blob()
        .len();
    let comments_bytes = opc
        .get_part(&PackURI::new(MODERN_A).unwrap())
        .unwrap()
        .blob()
        .len();
    assert!(comments_bytes > presentation_bytes);
    let limits = SourceLimits::builder()
        .max_part_bytes((comments_bytes - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    let error = match SourceBackedPresentation::from_reader_with_options(
        Cursor::new(bytes),
        litchi_pptx::OpenOptions::default().source_limits(limits),
    ) {
        Ok(source) => source.slide(1).unwrap().modern_comments().unwrap_err(),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(litchi_opc::OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            ..
        })
    ));
}

#[test]
fn source_modern_comments_project_the_mce_fallback_body() {
    let old_body = format!(
        r#"<p188:txBody><a:bodyPr/><a:lstStyle/>{}{}</p188:txBody>"#,
        paragraph("Check the Q3 numbers"),
        paragraph("and the chart")
    );
    let new_body = concat!(
        r#"<p188:txBody><a:bodyPr/><a:lstStyle/>"#,
        r#"<mc:AlternateContent><mc:Choice Requires="x"><x:discard/></mc:Choice>"#,
        r#"<mc:Fallback><a:p><a:r><a:t>MCE fallback</a:t></a:r></a:p></mc:Fallback>"#,
        r#"</mc:AlternateContent></p188:txBody>"#,
    );
    let bytes = rewrite(&deck(), |opc| {
        edit_blob(opc, MODERN_A, |xml| xml.replace(&old_body, new_body));
    });
    let source = open(&bytes);
    let list = source.slide(1).unwrap().modern_comments().unwrap().unwrap();
    assert_eq!(list.entries()[0].text(), Some("MCE fallback"));
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(61, self.revision.load(Ordering::SeqCst)))
    }
}

#[test]
fn source_modern_comments_fence_new_reads_and_keep_detached_values_readable() {
    let source_bytes = deck();
    let source = Arc::new(VersionedSource::new(source_bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let list = presentation
        .slide(1)
        .unwrap()
        .modern_comments()
        .unwrap()
        .unwrap();

    source.changed();
    assert!(matches!(
        presentation.slide(1).unwrap().modern_comments(),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
    ));
    assert_eq!(
        list.entries()[0].text(),
        Some("Check the Q3 numbers\nand the chart")
    );
    assert_eq!(list.entries()[0].replies()[0].text(), Some("Agreed & done"));
}

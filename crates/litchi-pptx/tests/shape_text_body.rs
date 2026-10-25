//! Focused PresentationML shape text-body ownership and namespace tests.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::shape::{Scene, raw};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn shape(body: &str, extra: &str, shape_extra: &str, tree_extra: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="{PML}" {extra}><p:cSld><p:spTree {tree_extra}><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp {shape_extra}><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>{body}</p:sp>
        </p:spTree></p:cSld></p:sld>"#
    )
}

const BODY: &str = r#"<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" lvl="1"/><a:r><a:rPr b="1" sz="2400"/><a:t>hello</a:t></a:r><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>1</a:t></a:fld><a:r><a:rPr><a:hlinkClick r:id="rId1" action="ppaction://hlinksldjump"/></a:rPr><a:t>link</a:t></a:r></a:p></p:txBody>"#;

#[test]
fn body_properties_and_runs_are_read_in_each_ancestor_namespace_scope() {
    let variants = [
        (
            "root",
            format!(r#"xmlns:a="{DML}" xmlns:r="{REL}""#),
            String::new(),
            String::new(),
        ),
        (
            "tree",
            format!(r#"xmlns:r="{REL}""#),
            String::new(),
            format!(r#"xmlns:a="{DML}""#),
        ),
        (
            "shape",
            format!(r#"xmlns:r="{REL}""#),
            format!(r#"xmlns:a="{DML}""#),
            String::new(),
        ),
        (
            "body",
            format!(r#"xmlns:a="{DML}" xmlns:r="{REL}""#),
            String::new(),
            String::new(),
        ),
    ];

    for (label, extra, shape_extra, tree_extra) in variants {
        let body = if label == "body" {
            BODY.replacen("<p:txBody>", &format!(r#"<p:txBody xmlns:a="{DML}">"#), 1)
        } else {
            BODY.to_owned()
        };
        let xml = shape(&body, &extra, &shape_extra, &tree_extra);
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT)
            .unwrap_or_else(|error| panic!("{label}: {error}"));
        let selected = scene.get("Text").unwrap().unwrap();
        let body = selected
            .text_body()
            .unwrap_or_else(|error| panic!("{label}: {error}"))
            .unwrap_or_else(|| panic!("{label}: missing body"));
        assert_eq!(body.text(), "hello1link", "{label}");
        assert_eq!(
            body.properties.vertical_anchor,
            litchi_drawingml::text::Anchor::Center
        );
        assert_eq!(body.paragraphs.len(), 1);
        assert_eq!(
            body.paragraphs[0].properties.alignment,
            Some(litchi_drawingml::text::Alignment::Center)
        );
        assert_eq!(body.paragraphs[0].properties.level, Some(1));
        assert_eq!(body.paragraphs[0].runs[0].properties.bold, Some(true));
        assert_eq!(
            body.paragraphs[0].runs[0]
                .properties
                .size
                .map(|value| value.get()),
            Some(2400)
        );
        assert!(matches!(
            body.paragraphs[0].runs[1].kind,
            litchi_drawingml::text::body::RunKind::Field(_)
        ));
        assert!(body.paragraphs[0].runs[2].properties.hyperlink.is_some());
    }
}

#[test]
fn group_namespace_is_retained_and_bom_offsets_cover_the_processed_owner() {
    let xml = format!(
        r#"<p:sld xmlns:p="{PML}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
        <p:grpSp xmlns:d="{DML}"><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                <p:txBody><d:bodyPr/><d:p><d:r><d:t>grouped</d:t></d:r></d:p></p:txBody>
            </p:sp>
        </p:grpSp></p:spTree></p:cSld></p:sld>"#
    );
    let mut marked = vec![0xEF, 0xBB, 0xBF];
    marked.extend_from_slice(xml.as_bytes());
    let scene = Scene::read(&marked, litchi_pptx::shape::Limits::DEFAULT).expect("BOM owner");
    let selected = scene.get("Text").unwrap().unwrap();
    assert_eq!(selected.text_body().unwrap().unwrap().text(), "grouped");
    let span = raw::text_body_span(selected).expect("raw text-body span");
    let start = usize::try_from(span.start()).unwrap();
    let end = usize::try_from(span.end().unwrap()).unwrap();
    assert!(start >= 3);
    assert!(scene.xml()[start..end].starts_with(b"<p:txBody>"));
    assert_eq!(&scene.xml()[start..end], &marked[start..end]);
}

#[test]
fn mid_tree_namespace_rebindings_are_used_for_lazy_bodies() {
    let xml = format!(
        r#"<p:sld xmlns:p="{PML}"><p:cSld xmlns:d="{DML}"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="MidTree"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                <p:txBody><d:bodyPr/><d:p><d:r><d:t>mid-tree</d:t></d:r></d:p></p:txBody>
            </p:sp>
        </p:spTree></p:cSld></p:sld>"#
    );
    let scene =
        Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).expect("mid-tree owner");
    let selected = scene.get("MidTree").unwrap().unwrap();
    assert_eq!(selected.text_body().unwrap().unwrap().text(), "mid-tree");
}

#[test]
fn owned_lazy_body_keeps_the_caller_limits() {
    let xml = shape(BODY, &format!(r#"xmlns:a="{DML}" xmlns:r="{REL}""#), "", "");
    let limits = litchi_pptx::shape::Limits::new(xml.len(), xml.len(), 64, 128, 1, 64)
        .expect("finite limits");
    let scene = litchi_pptx::shape::owned::Scene::read(xml.into_bytes(), limits)
        .expect("small bounded owner");
    assert_eq!(scene.limits(), limits);
    let selected = scene.shape("Text").unwrap().unwrap();
    assert_eq!(
        selected.view().text_body().unwrap().unwrap().text(),
        "hello1link"
    );
}

#[test]
fn duplicate_direct_bodies_are_refused_but_nested_foreign_body_is_not_owned() {
    let duplicate = format!(
        "{}{}",
        BODY, r#"<p:txBody><a:bodyPr/><a:p><a:r><a:t>second</a:t></a:r></a:p></p:txBody>"#
    );
    let duplicate = shape(
        &duplicate,
        &format!(r#"xmlns:a="{DML}" xmlns:r="{REL}""#),
        "",
        "",
    );
    let scene = Scene::read(duplicate.as_bytes(), litchi_pptx::shape::Limits::DEFAULT)
        .expect("duplicate owner remains indexed");
    let selected = scene.get("Text").unwrap().unwrap();
    assert!(selected.text_body().is_err());
    assert!(raw::text_body_span(selected).is_some());

    let nested = shape(
        r#"<p:spPr><p:extLst><p:txBody><a:p><a:r><a:t>nested</a:t></a:r></a:p></p:txBody></p:extLst></p:spPr>"#,
        &format!(r#"xmlns:a="{DML}""#),
        "",
        "",
    );
    let nested_scene =
        Scene::read(nested.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).expect("nested owner");
    let selected = nested_scene.get("Text").unwrap().unwrap();
    assert!(raw::text_body_span(selected).is_none());
    assert_eq!(selected.text_body().unwrap(), None);
}

#[test]
fn markup_compatibility_selection_and_strict_property_errors_are_preserved() {
    let mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
    let xml = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:mc="{mc}" xmlns:p14="{p14}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <mc:AlternateContent><mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="2" name="Choice"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:pPr algn="middle"/><a:r><a:t>chosen</a:t></a:r></a:p></p:txBody></p:sp></mc:Choice>
            <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="3" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent>
        </p:spTree></p:cSld></p:sld>"#
    );
    let scene =
        Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).expect("MCE owner");
    assert!(scene.is_rewritten());
    let selected = scene.get("Choice").unwrap().unwrap();
    assert!(scene.get("Fallback").unwrap().is_none());
    assert!(matches!(
        selected.text_body(),
        Err(litchi_pptx::Error::Drawing(_))
    ));
}

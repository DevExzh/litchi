use super::{Anchor, Form, scan};
use crate::package::story::StoryDialect;

const T_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const S_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const T_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const S_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const T_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const T_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WPI: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WPC: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WPG: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn run(xml: &str, dialect: StoryDialect) -> crate::Result<Vec<Anchor>> {
    scan(xml.as_bytes(), dialect, 256, 64, 32)
}

fn ids(anchors: &[Anchor]) -> Vec<(&str, Form)> {
    anchors
        .iter()
        .map(|anchor| (anchor.relationship_id.as_str(), anchor.form))
        .collect()
}

#[test]
fn base_anchor_requires_the_story_dialect_but_allows_aliases_and_default_word_namespace() {
    let strict = format!(
        r#"<document xmlns="{S_WORD}" xmlns:z="{S_REL}"><body><p><r><contentPart z:id="sInk"/></r></p></body></document>"#
    );
    assert_eq!(
        ids(&run(&strict, StoryDialect::Strict).expect("strict base anchor")),
        vec![("sInk", Form::Base)]
    );
    assert!(
        run(&strict, StoryDialect::Transitional)
            .expect("foreign strict content is ignored")
            .is_empty()
    );

    let transitional = format!(
        r#"<q:document xmlns:q="{T_WORD}" xmlns:u="{T_REL}"><q:body><q:p><q:r><q:contentPart u:id="tInk"/></q:r></q:p></q:body></q:document>"#
    );
    assert_eq!(
        ids(&run(&transitional, StoryDialect::Transitional).expect("transitional base anchor")),
        vec![("tInk", Form::Base)]
    );
}

#[test]
fn base_anchor_does_not_use_local_name_or_unbound_relationship_guesses() {
    let foreign = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:x="urn:foreign" xmlns:u="{T_REL}"><d:body><d:p><d:r><x:contentPart x:id="ignored"/><d:contentPart id="missing"/></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&foreign, StoryDialect::Transitional).is_err());

    let unbound = format!(
        r#"<d:document xmlns:d="{T_WORD}"><d:body><d:p><d:r><d:contentPart id="missing"/></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&unbound, StoryDialect::Transitional).is_err());

    let wrong_relationship = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:s="{S_REL}"><d:body><d:p><d:r><d:contentPart s:id="wrongDialect"/></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&wrong_relationship, StoryDialect::Transitional).is_err());
}

#[test]
fn base_content_part_must_be_directly_hosted_by_run() {
    let xml = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:body><d:p><d:contentPart r:id="bad"/></d:p></d:body></d:document>"#
    );
    assert!(run(&xml, StoryDialect::Transitional).is_err());
}

#[test]
fn drawing_ink_anchor_requires_exact_uri_and_legal_canvas_or_group_ancestry() {
    let graphic_data = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:w14="{W14}" xmlns:wpi="{WPI}" xmlns:mc="{MC}" xmlns:r="{T_REL}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpi"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="drawingInk"/></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&graphic_data, StoryDialect::Transitional).expect("drawing ink anchor")),
        vec![("drawingInk", Form::Drawing)]
    );

    let canvas = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:wpc="{WPC}" xmlns:w14="{W14}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpc"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPC}"><wpc:wpc><w14:contentPart r:id="canvasInk"/></wpc:wpc></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&canvas, StoryDialect::Transitional).expect("canvas content part")),
        vec![("canvasInk", Form::GenericDrawing)]
    );

    let group = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:wpg="{WPG}" xmlns:w14="{W14}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpg"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPG}"><wpg:wgp><wpg:grpSp><w14:contentPart r:id="groupInk"/></wpg:grpSp></wpg:wgp></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&group, StoryDialect::Transitional).expect("group content part")),
        vec![("groupInk", Form::GenericDrawing)]
    );

    let wrong_uri = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:w14="{W14}" xmlns:r="{T_REL}"><d:body><d:p><d:r><d:drawing><wp:inline><a:graphic><a:graphicData uri="urn:other"><w14:contentPart r:id="bad"/></a:graphicData></a:graphic></wp:inline></d:drawing></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&wrong_uri, StoryDialect::Transitional).is_err());

    let canvas_group = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:wpc="{WPC}" xmlns:wpg="{WPG}" xmlns:w14="{W14}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpc"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPC}"><wpc:wpc><wpg:wgp><w14:contentPart r:id="canvasGroupInk"/></wpg:wgp></wpc:wpc></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&canvas_group, StoryDialect::Transitional).expect("canvas group content part")),
        vec![("canvasGroupInk", Form::GenericDrawing)]
    );
}

#[test]
fn drawing_ink_requires_the_wordprocessing_drawing_host_path() {
    let direct = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:w14="{W14}" xmlns:r="{T_REL}"><d:body><d:p><d:r><a:graphicData uri="{WPI}"><w14:contentPart r:id="bad"/></a:graphicData></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&direct, StoryDialect::Transitional).is_err());
}

#[test]
fn drawing_ink_accepts_the_anchor_wordprocessing_drawing_host() {
    let xml = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:w14="{W14}" xmlns:wpi="{WPI}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpi"><d:drawing><wp:anchor><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="anchoredInk"/></a:graphicData></a:graphic></wp:anchor></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&xml, StoryDialect::Transitional).expect("anchored drawing Ink")),
        vec![("anchoredInk", Form::Drawing)]
    );
}

#[test]
fn generic_drawing_mce_checks_requires_fallback_and_choice_count() {
    let mismatched_requires = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:wpc="{WPC}" xmlns:w14="{W14}" xmlns:wpi="{WPI}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpi"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPC}"><wpc:wpc><w14:contentPart r:id="wrongRequires"/></wpc:wpc></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&mismatched_requires, StoryDialect::Transitional).is_err());

    let valid = mismatched_requires.replace("Requires=\"wpi\"", "Requires=\"wpc\"");
    assert!(run(&valid, StoryDialect::Transitional).is_ok());
    let without_mce = valid
        .replace("<mc:AlternateContent><mc:Choice Requires=\"wpc\">", "")
        .replace(
            "</mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent>",
            "",
        );
    assert!(run(&without_mce, StoryDialect::Transitional).is_err());
    let missing_fallback = valid.replace("<mc:Fallback><d:pict/></mc:Fallback>", "");
    assert!(run(&missing_fallback, StoryDialect::Transitional).is_err());

    let invalid_fallback = mismatched_requires
        .replace("Requires=\"wpi\"", "Requires=\"wpc\"")
        .replace(
            "<mc:Fallback><d:pict/></mc:Fallback>",
            "<mc:Fallback><d:p/></mc:Fallback>",
        );
    assert!(run(&invalid_fallback, StoryDialect::Transitional).is_err());

    let duplicate_choice = mismatched_requires
        .replace("Requires=\"wpi\"", "Requires=\"wpc\"")
        .replace(
            "</mc:Choice><mc:Fallback>",
            "</mc:Choice><mc:Choice Requires=\"wpc\"><d:p/></mc:Choice><mc:Fallback>",
        );
    assert!(run(&duplicate_choice, StoryDialect::Transitional).is_err());
}

#[test]
fn generic_canvas_mce_allows_multiple_content_parts_in_one_choice() {
    let xml = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:wpc="{WPC}" xmlns:w14="{W14}" xmlns:r="{T_REL}" xmlns:mc="{MC}"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpc"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPC}"><wpc:wpc><w14:contentPart r:id="canvasOne"/><w14:contentPart r:id="canvasTwo"/></wpc:wpc></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(
            &xml.replace("wp:inline", "wp:anchor"),
            StoryDialect::Transitional
        )
        .expect("two canvas content parts")),
        vec![
            ("canvasOne", Form::GenericDrawing),
            ("canvasTwo", Form::GenericDrawing)
        ]
    );
}

#[test]
fn drawing_relationship_namespace_is_transitional_extension_namespace() {
    let strict_relationship = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:w14="{W14}" xmlns:r="{S_REL}"><d:body><d:p><d:r><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="wrong"/></a:graphicData></a:graphic></wp:inline></d:drawing></d:r></d:p></d:body></d:document>"#
    );
    assert!(run(&strict_relationship, StoryDialect::Transitional).is_err());
}

#[test]
fn active_mce_choice_is_inventoried_and_fallback_is_not_projected() {
    let xml = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:a="{T_DML}" xmlns:wp="{T_WP}" xmlns:w14="{W14}" xmlns:wpi="{WPI}" xmlns:r="{T_REL}" xmlns:mc="{MC}" mc:Ignorable="w14"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="wpi"><d:drawing><wp:inline><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="choice"/></a:graphicData></a:graphic></wp:inline></d:drawing></mc:Choice><mc:Fallback><d:pict/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&xml, StoryDialect::Transitional).expect("active Ink Choice")),
        vec![("choice", Form::Drawing)]
    );
}

#[test]
fn mce_choice_budget_is_independent_of_pending_ink_candidates() {
    let xml = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:r="{T_REL}" xmlns:mc="{MC}" xmlns:wpi="{WPI}" xmlns:u="urn:unknown"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="u"><d:p/></mc:Choice><mc:Choice Requires="wpi"><d:contentPart r:id="choice"/></mc:Choice><mc:Choice Requires="u"><d:p/></mc:Choice><mc:Fallback><d:p/></mc:Fallback></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert_eq!(
        ids(&run(&xml, StoryDialect::Transitional).expect("all bounded choices are scanned")),
        vec![("choice", Form::Base)]
    );
}

#[test]
fn unknown_mce_choice_is_dropped_before_anchor_semantics() {
    let unknown = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:r="{T_REL}" xmlns:mc="{MC}" xmlns:u="urn:unknown"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="u"><d:contentPart r:id="inactive"/></mc:Choice><mc:Fallback/></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert!(
        run(&unknown, StoryDialect::Transitional)
            .expect("unknown Choice is inactive")
            .is_empty()
    );

    let malformed = format!(
        r#"<d:document xmlns:d="{T_WORD}" xmlns:mc="{MC}" xmlns:u="urn:unknown"><d:body><d:p><d:r><mc:AlternateContent><mc:Choice Requires="u"><d:contentPart/></mc:Choice><mc:Fallback/></mc:AlternateContent></d:r></d:p></d:body></d:document>"#
    );
    assert!(
        run(&malformed, StoryDialect::Transitional)
            .expect("inactive malformed candidate is dropped")
            .is_empty()
    );
}

#[test]
fn active_base_content_part_has_empty_ct_rel_content() {
    let space = ' ';
    let whitespace = format!(
        r#"<d:r xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:contentPart r:id="ok">{space}
        </d:contentPart></d:r>"#
    );
    assert_eq!(
        ids(&run(&whitespace, StoryDialect::Transitional).expect("whitespace is legal")),
        vec![("ok", Form::Base)]
    );

    let text = format!(
        r#"<d:r xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:contentPart r:id="bad">ink</d:contentPart></d:r>"#
    );
    assert!(run(&text, StoryDialect::Transitional).is_err());

    let element = format!(
        r#"<d:r xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:contentPart r:id="bad"><d:custom/></d:contentPart></d:r>"#
    );
    assert!(run(&element, StoryDialect::Transitional).is_err());
}

#[test]
fn malformed_envelopes_and_unknown_prefixes_are_rejected() {
    let cases = [
        b"".as_slice(),
        b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document><w:document/>".as_slice(),
        b"<document xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><body/></document>trailing".as_slice(),
        b"<!DOCTYPE document><document/>".as_slice(),
        b"<?pi?><document/>".as_slice(),
        b" \n<?xml version=\"1.0\"?><document/>".as_slice(),
        b"<u:document><u:body/></u:document>".as_slice(),
    ];
    for xml in cases {
        assert!(scan(xml, StoryDialect::Transitional, 256, 64, 32).is_err());
    }
}

#[test]
fn limits_are_applied_without_rejecting_zero_anchor_budget_up_front() {
    let empty = format!(r#"<d:r xmlns:d="{T_WORD}"/>"#);
    assert!(
        scan(empty.as_bytes(), StoryDialect::Transitional, 1, 1, 0)
            .expect("empty zero-anchor scan")
            .is_empty()
    );

    let one =
        format!(r#"<d:r xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:contentPart r:id="one"/></d:r>"#);
    assert!(scan(one.as_bytes(), StoryDialect::Transitional, 2, 2, 0).is_err());
    assert!(scan(one.as_bytes(), StoryDialect::Transitional, 1, 2, 32).is_err());
    assert!(scan(one.as_bytes(), StoryDialect::Transitional, 2, 1, 32).is_err());

    let two = format!(
        r#"<d:r xmlns:d="{T_WORD}" xmlns:r="{T_REL}"><d:contentPart r:id="one"/><d:contentPart r:id="two"/></d:r>"#
    );
    assert!(scan(two.as_bytes(), StoryDialect::Transitional, 8, 2, 1).is_err());
}

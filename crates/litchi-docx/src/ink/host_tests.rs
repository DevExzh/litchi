use std::ops::Range;

use super::{Host, capture, reference_values};
use crate::ink::Limits;
use crate::package::story::StoryDialect;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WPI: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WPC: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WPG: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn capture_story(xml: &str) -> crate::Result<Vec<Host>> {
    capture(
        xml.as_bytes(),
        StoryDialect::Transitional,
        Limits::default(),
    )
}

fn span(xml: &str, needle: &str) -> Range<usize> {
    let start = xml.find(needle).expect("fixture span");
    start..start + needle.len()
}

fn drawing(placement: &str, id: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpi">{}</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>"#,
        drawing_object(placement, id)
    )
}

fn drawing_object(placement: &str, id: &str) -> String {
    format!(
        r#"<w:drawing><wp:{placement}><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="{id}"/></a:graphicData></a:graphic></wp:{placement}></w:drawing>"#
    )
}

fn group_object(id: &str, metadata: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><a:graphic><a:graphicData uri="{WPG}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>{metadata}</wpg:grpSpPr><w14:contentPart r:id="{id}"/></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn group_fragment(object: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpg">{object}</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>"#
    )
}

const GROUP_TRANSFORM: &str = r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm>"#;

fn story(body: &str, extra_namespaces: &str) -> String {
    format!(
        r#"<w:document xmlns:w="{WORD}" xmlns:r="{REL}" xmlns:a="{DML}" xmlns:wp="{WP}" xmlns:w14="{W14}" xmlns:mc="{MC}" xmlns:wpi="{WPI}" xmlns:wpc="{WPC}" xmlns:wpg="{WPG}" xmlns:v="urn:schemas-microsoft-com:vml" {extra_namespaces}><w:body>{body}</w:body></w:document>"#
    )
}

#[test]
fn direct_base_anchor_has_exact_source_and_removal_ranges() {
    let xml = story(
        r#"<w:p><w:r><w:t>before</w:t><w:contentPart r:id="base"/><w:t>after</w:t></w:r></w:p>"#,
        "",
    );
    let hosts = capture_story(&xml).expect("base host");
    assert_eq!(hosts.len(), 1);
    let host = &hosts[0];
    let anchor = span(&xml, r#"<w:contentPart r:id="base"/>"#);
    assert_eq!(host.anchor.relationship_id, "base");
    assert!(host.removable);
    assert_eq!(host.anchor_span, anchor);
    assert_eq!(host.removal_span, Some(anchor.clone()));
    assert_eq!(host.removal_start_tag, Some(anchor));
}

#[test]
fn drawing_inline_and_anchor_remove_the_whole_alternate_content_host() {
    for placement in ["inline", "anchor"] {
        let fragment = drawing(placement, "ink");
        let xml = story(
            &format!(r#"<w:p><w:r><w:t>left</w:t>{fragment}<w:t>right</w:t></w:r></w:p>"#),
            "",
        );
        let hosts = capture_story(&xml).expect("drawing host");
        assert_eq!(hosts.len(), 1, "{placement}");
        let host = &hosts[0];
        let whole = span(&xml, &fragment);
        let anchor = span(&xml, r#"<w14:contentPart r:id="ink"/>"#);
        assert!(host.removable, "{placement}");
        assert_eq!(host.anchor_span, anchor, "{placement}");
        assert_eq!(host.removal_span, Some(whole.clone()), "{placement}");
        assert_eq!(
            host.removal_start_tag,
            Some(span(&xml, "<mc:AlternateContent>")),
            "{placement}"
        );
    }
}

#[test]
fn canonical_group_metadata_is_removable() {
    let object = group_object("groupInk", GROUP_TRANSFORM);
    let fragment = group_fragment(&object);
    let xml = story(&format!(r#"<w:p><w:r>{fragment}</w:r></w:p>"#), "");
    let hosts = capture_story(&xml).expect("canonical group host");
    assert_eq!(hosts.len(), 1);
    assert!(hosts[0].removable);
}

#[test]
fn unsupported_group_metadata_refuses_only_that_host() {
    let object = group_object(
        "groupInk",
        r#"<a:xfrm rot="0" flipH="1"><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm>"#,
    );
    let fragment = group_fragment(&object);
    let xml = story(
        &format!(r#"<w:p><w:r><w:contentPart r:id="base"/>{fragment}</w:r></w:p>"#),
        "",
    );
    let hosts = capture_story(&xml).expect("unsupported group metadata remains inventory-only");
    assert_eq!(hosts.len(), 2);
    assert!(hosts[0].removable);
    assert!(!hosts[1].removable);
    assert!(hosts[1].removal_span.is_none());
    assert!(hosts[1].removal_start_tag.is_none());
}

#[test]
fn unsupported_group_does_not_poison_a_later_drawing_sibling() {
    let object = group_object(
        "groupInk",
        r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="not-an-integer" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm>"#,
    );
    let group = group_fragment(&object);
    let drawing = drawing("inline", "drawingInk");
    let xml = story(&format!(r#"<w:p><w:r>{group}{drawing}</w:r></w:p>"#), "");
    let hosts = capture_story(&xml).expect("unsupported group remains inventory-only");
    assert_eq!(hosts.len(), 2);
    assert!(!hosts[0].removable);
    assert!(hosts[0].removal_span.is_none());
    assert!(hosts[0].removal_start_tag.is_none());
    assert!(hosts[1].removable);
    assert!(hosts[1].removal_span.is_some());
    assert!(hosts[1].removal_start_tag.is_some());
}

#[test]
fn group_metadata_order_or_extra_object_refuses_removal() {
    for (label, metadata) in [
        (
            "reordered transform",
            r#"<a:xfrm><a:ext cx="127000" cy="127000"/><a:off x="0" y="0"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm>"#,
        ),
        (
            "extra object",
            r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm><wpg:wsp/>"#,
        ),
    ] {
        let object = group_object("groupInk", metadata);
        let fragment = group_fragment(&object);
        let xml = story(&format!(r#"<w:p><w:r>{fragment}</w:r></w:p>"#), "");
        let hosts = capture_story(&xml).expect("malformed group remains inventory-only");
        assert_eq!(hosts.len(), 1, "{label}");
        assert!(!hosts[0].removable, "{label}");
        assert!(hosts[0].removal_span.is_none(), "{label}");
        assert!(hosts[0].removal_start_tag.is_none(), "{label}");
    }
}

#[test]
fn missing_empty_or_duplicate_group_transform_refuses_only_its_host() {
    let missing = group_object("groupInk", "");
    for (label, object) in [
        ("missing transform", missing.clone()),
        (
            "empty properties",
            missing.replace("<wpg:grpSpPr></wpg:grpSpPr>", "<wpg:grpSpPr/>"),
        ),
        ("empty transform", group_object("groupInk", "<a:xfrm/>")),
        (
            "duplicate transform",
            group_object("groupInk", &format!("{GROUP_TRANSFORM}{GROUP_TRANSFORM}")),
        ),
    ] {
        let fragment = group_fragment(&object);
        let sibling = drawing("inline", "drawingInk");
        let xml = story(&format!("<w:p><w:r>{fragment}{sibling}</w:r></w:p>"), "");
        let hosts = capture_story(&xml).expect("unsupported group remains inventory-only");
        assert_eq!(hosts.len(), 2, "{label}");
        assert!(!hosts[0].removable, "{label}");
        assert!(hosts[0].removal_span.is_none(), "{label}");
        assert!(hosts[1].removable, "{label}");
    }
}

#[test]
fn same_run_siblings_are_outside_the_drawing_removal_span() {
    let fragment = drawing("inline", "ink");
    let xml = story(
        &format!(r#"<w:p><w:r><w:t>left</w:t>{fragment}<w:t>right</w:t></w:r></w:p>"#),
        "",
    );
    let host = &capture_story(&xml).expect("drawing siblings")[0];
    let removal = host.removal_span.clone().expect("eligible removal");
    assert_eq!(&xml.as_bytes()[removal], fragment.as_bytes());
    assert!(xml.contains("<w:t>left</w:t>"));
    assert!(xml.contains("<w:t>right</w:t>"));
}

#[test]
fn base_mce_anchor_is_inventory_only_until_alternative_semantics_are_modeled() {
    let xml = story(
        r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="unknown"><w:contentPart r:id="inactive"/></mc:Choice><mc:Fallback><w:contentPart r:id="active"/></mc:Fallback></mc:AlternateContent></w:r></w:p>"#,
        r#"xmlns:unknown="urn:future""#,
    );
    let hosts = capture_story(&xml).expect("active fallback base host");
    assert_eq!(hosts.len(), 1);
    assert_eq!(hosts[0].anchor.relationship_id, "active");
    assert!(!hosts[0].removable);
    assert!(hosts[0].removal_span.is_none());
    assert!(hosts[0].removal_start_tag.is_none());
}

#[test]
fn nonempty_base_content_part_is_rejected_before_capture() {
    let xml = story(
        r#"<w:p><w:r><w:contentPart r:id="bad"><w:t>text</w:t></w:contentPart></w:r></w:p>"#,
        "",
    );
    assert!(capture_story(&xml).is_err());
}

#[test]
fn multiple_active_drawing_objects_in_one_choice_are_not_removable() {
    let first = drawing_object("inline", "one");
    let second = drawing_object("anchor", "two");
    let fragment = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpi">{first}{second}</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>"#
    );
    let xml = story(&format!(r#"<w:p><w:r>{fragment}</w:r></w:p>"#), "");
    let hosts = capture_story(&xml).expect("drawing siblings are inventoried");
    assert_eq!(hosts.len(), 2);
    assert!(hosts.iter().all(|host| !host.removable));
    assert!(hosts.iter().all(|host| host.removal_start_tag.is_none()));
    assert!(hosts.iter().all(|host| host.removal_span.is_none()));
}

#[test]
fn extra_choice_content_is_not_removable() {
    let object = drawing_object("inline", "ink");
    let fragment = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpi">{object}<w:t>unselected</w:t></mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>"#
    );
    let xml = story(&format!(r#"<w:p><w:r>{fragment}</w:r></w:p>"#), "");
    let hosts = capture_story(&xml).expect("choice sibling host");
    assert_eq!(hosts.len(), 1);
    assert!(!hosts[0].removable);
    assert!(hosts[0].removal_start_tag.is_none());
    assert!(hosts[0].removal_span.is_none());
}

#[test]
fn structural_wrapper_text_before_or_after_an_anchor_is_not_discarded() {
    for insertion in ["unselected", "<![CDATA[unselected]]>", "&#65;"] {
        for (needle, replacement) in [
            ("<wp:inline>".to_owned(), format!("<wp:inline>{insertion}")),
            (
                "</wp:inline>".to_owned(),
                format!("{insertion}</wp:inline>"),
            ),
        ] {
            let fragment = drawing("inline", "ink").replace(&needle, &replacement);
            let xml = story(&format!("<w:p><w:r>{fragment}</w:r></w:p>"), "");
            let hosts = capture_story(&xml).expect("readable unmodeled text");
            assert_eq!(hosts.len(), 1);
            assert!(!hosts[0].removable);
            assert!(hosts[0].removal_span.is_none());
        }
    }
}

#[test]
fn multiple_vml_fallback_shapes_are_not_removable() {
    let fragment = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpi"><w:drawing><wp:inline><a:graphic><a:graphicData uri="{WPI}"><w14:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="selected"/><v:shape id="unselected"/></w:pict></mc:Fallback></mc:AlternateContent>"#
    );
    let xml = story(&format!(r#"<w:p><w:r>{fragment}</w:r></w:p>"#), "");
    let hosts = capture_story(&xml).expect("fallback shape host");
    assert_eq!(hosts.len(), 1);
    assert!(!hosts[0].removable);
    assert!(hosts[0].removal_start_tag.is_none());
    assert!(hosts[0].removal_span.is_none());
}

#[test]
fn multiple_canvas_slots_are_classified_but_refuse_removal() {
    let xml = story(
        &format!(
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpc"><w:drawing><wp:inline><a:graphic><a:graphicData uri="{WPC}"><wpc:wpc><w14:contentPart r:id="one"/><w14:contentPart r:id="two"/></wpc:wpc></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="fallback"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"#
        ),
        "",
    );
    let hosts = capture_story(&xml).expect("canvas slots are inventoried");
    assert_eq!(hosts.len(), 2);
    assert!(hosts.iter().all(|host| !host.removable));
    assert!(hosts.iter().all(|host| host.removal_start_tag.is_none()));
    assert!(hosts.iter().all(|host| host.removal_span.is_none()));
}

#[test]
fn nested_or_foreign_fallback_objects_are_not_discarded() {
    for extra in [
        "<v:group><v:shape id=\"nested\"/></v:group>",
        "<v:rect/>",
        "<v:textbox/>",
        "<x:object xmlns:x=\"urn:producer\"/>",
    ] {
        for closing in ["</w:pict>", "</v:shape>"] {
            let fragment = drawing("inline", "ink").replace(closing, &format!("{extra}{closing}"));
            let xml = story(&format!("<w:p><w:r>{fragment}</w:r></w:p>"), "");
            let hosts = capture_story(&xml).expect("readable fallback extension");
            assert_eq!(hosts.len(), 1);
            assert!(!hosts[0].removable);
            assert!(hosts[0].removal_span.is_none());
        }
    }
}

#[test]
fn reference_values_include_inactive_branches_and_unknown_attribute_tokens() {
    let xml = format!(
        r##"<w:document xmlns:w="{WORD}" xmlns:r="{REL}" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r r:id="run" r:embed="embedded" r:link="linked" o:relid="vml" custom="#future other"/></w:p></w:body></w:document>"##
    );
    let values = reference_values(xml.as_bytes(), Limits::default()).expect("reference census");
    for value in ["run", "embedded", "linked", "vml", "future", "other"] {
        assert!(values.iter().any(|candidate| candidate == value), "{value}");
    }
}

/// Outcome of `reference_values` for one input, as frozen text.
fn reference_values_envelope_outcome(xml: &str) -> String {
    match reference_values(xml.as_bytes(), Limits::default()) {
        Ok(values) => format!("Ok({})", values.len()),
        Err(error) => format!("{error:?}"),
    }
}

#[test]
fn reference_values_xml_envelope_matrix() {
    let rows: [(&str, &str, &str); 25] = [
        (
            "empty input",
            "",
            r#"Invalid("DOCX story has an unterminated XML root")"#,
        ),
        (
            "whitespace only",
            " \n",
            r#"Invalid("DOCX story has an unterminated XML root")"#,
        ),
        (
            "declaration only",
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"Invalid("DOCX story has an unterminated XML root")"#,
        ),
        (
            "empty-element root",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
            "Ok(1)",
        ),
        (
            "start and end root",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            "Ok(1)",
        ),
        (
            "root then whitespace",
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/></w:body></w:document>\n ",
            "Ok(1)",
        ),
        (
            "root then comment",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><!--c-->"#,
            "Ok(1)",
        ),
        (
            "root then processing instruction",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><?pi x?>"#,
            r#"Invalid("DOCX story rejects processing instructions")"#,
        ),
        (
            "root then text",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>text"#,
            r#"Invalid("DOCX story has non-whitespace text outside its root")"#,
        ),
        (
            "root then a second identical root",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has more than one XML root")"#,
        ),
        (
            "root then a foreign start element",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><foreign></foreign>"#,
            r#"Invalid("DOCX story has more than one XML root")"#,
        ),
        (
            "root then a foreign empty element",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><foreign/>"#,
            r#"Invalid("DOCX story has more than one XML root")"#,
        ),
        (
            "declaration after the root",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document><?xml version="1.0" encoding="UTF-8"?>"#,
            r#"Invalid("DOCX story has an XML declaration outside its prolog")"#,
        ),
        (
            "two declarations",
            r#"<?xml version="1.0" encoding="UTF-8"?><?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has an XML declaration outside its prolog")"#,
        ),
        (
            "comment before the declaration",
            r#"<!--c--><?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has an XML declaration outside its prolog")"#,
        ),
        (
            "text before the root",
            r#"text<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has non-whitespace text outside its root")"#,
        ),
        (
            "unterminated root",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body>"#,
            r#"Invalid("DOCX story has an unterminated XML root")"#,
        ),
        (
            "stray end tag",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document></w:document>"#,
            r#"Xml("ill-formed document: close tag `</w:document>` does not match any open tag")"#,
        ),
        (
            "document type declaration",
            r#"<!DOCTYPE x><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story rejects DTDs")"#,
        ),
        (
            "top-level character data section",
            r#"<![CDATA[x]]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has CDATA outside its root")"#,
        ),
        (
            "declaration then root",
            r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            "Ok(1)",
        ),
        (
            "whitespace before the declaration",
            r#" <?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            r#"Invalid("DOCX story has an XML declaration outside its prolog")"#,
        ),
        (
            "comment then root",
            r#"<!--c--><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            "Ok(1)",
        ),
        (
            "declaration then comment then root",
            r#"<?xml version="1.0" encoding="UTF-8"?><!--c--><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            "Ok(1)",
        ),
        (
            "root then a nested stray declaration",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><?xml version="1.0" encoding="UTF-8"?></w:document>"#,
            r#"Invalid("DOCX story has an XML declaration outside its prolog")"#,
        ),
    ];
    for (name, xml, expected) in rows {
        assert_eq!(reference_values_envelope_outcome(xml), expected, "{name}");
    }
}

/// Every sequence of at most `max_len` symbols, shortest first.
fn sequences(symbols: &[char], max_len: usize) -> Vec<String> {
    let mut all = vec![String::new()];
    let mut start = 0;
    for _ in 0..max_len {
        let end = all.len();
        for index in start..end {
            for &symbol in symbols {
                let mut next = all[index].clone();
                next.push(symbol);
                all.push(next);
            }
        }
        start = end;
    }
    all
}

/// One character per captured host (`R` removable, `-` not), or the error.
fn group_host_outcome(group_children: &str) -> String {
    let object = format!(
        r#"<w:drawing><wp:inline><a:graphic><a:graphicData uri="{WPG}"><wpg:wgp>{group_children}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    );
    let fragment = group_fragment(&object);
    let xml = story(&format!("<w:p><w:r>{fragment}</w:r></w:p>"), "");
    match capture_story(&xml) {
        Ok(hosts) => hosts
            .iter()
            .map(|host| if host.removable { 'R' } else { '-' })
            .collect(),
        Err(error) => format!("error: {error}"),
    }
}

/// How many sequences produced each outcome, as `outcome:count`.
fn outcome_counts(outcomes: &[(String, String)]) -> Vec<String> {
    let mut counts = std::collections::BTreeMap::<String, usize>::new();
    for (_, outcome) in outcomes {
        *counts.entry(outcome.clone()).or_default() += 1;
    }
    counts
        .into_iter()
        .map(|(outcome, count)| format!("{outcome}:{count}"))
        .collect()
}

fn removable_sequences(outcomes: &[(String, String)]) -> Vec<&str> {
    outcomes
        .iter()
        .filter(|(_, outcome)| outcome.contains('R'))
        .map(|(sequence, _)| sequence.as_str())
        .collect::<Vec<_>>()
}

#[test]
fn group_child_order_matrix_accepts_only_the_canonical_sequence() {
    let outcomes: Vec<_> = sequences(&['N', 'P', 'C', 'X'], 4)
        .into_iter()
        .map(|sequence| {
            let children: String = sequence
                .chars()
                .map(|symbol| match symbol {
                    'N' => "<wpg:cNvGrpSpPr/>".to_owned(),
                    'P' => format!("<wpg:grpSpPr>{GROUP_TRANSFORM}</wpg:grpSpPr>"),
                    'C' => r#"<w14:contentPart r:id="groupInk"/>"#.to_owned(),
                    _ => "<wpg:wsp/>".to_owned(),
                })
                .collect();
            let outcome = group_host_outcome(&children);
            (sequence, outcome)
        })
        .collect();
    assert_eq!(removable_sequences(&outcomes), ["NPC"]);
    assert_eq!(
        outcome_counts(&outcomes),
        [":121", "-:141", "--:64", "---:13", "----:1", "R:1"]
    );
}

#[test]
fn group_properties_child_matrix_accepts_only_one_transform() {
    let outcomes: Vec<_> = sequences(&['T', 'X'], 3)
        .into_iter()
        .map(|sequence| {
            let children: String = sequence
                .chars()
                .map(|symbol| match symbol {
                    'T' => GROUP_TRANSFORM,
                    _ => "<a:ln/>",
                })
                .collect();
            let outcome = group_host_outcome(&format!(
                r#"<wpg:cNvGrpSpPr/><wpg:grpSpPr>{children}</wpg:grpSpPr><w14:contentPart r:id="groupInk"/>"#
            ));
            (sequence, outcome)
        })
        .collect();
    assert_eq!(removable_sequences(&outcomes), ["T"]);
    assert_eq!(outcome_counts(&outcomes), ["-:14", "R:1"]);
}

#[test]
fn group_transform_child_order_matrix_accepts_only_the_canonical_sequence() {
    let outcomes: Vec<_> = sequences(&['O', 'E', 'H', 'X', 'Z'], 5)
        .into_iter()
        .map(|sequence| {
            let children: String = sequence
                .chars()
                .map(|symbol| match symbol {
                    'O' => r#"<a:off x="0" y="0"/>"#,
                    'E' => r#"<a:ext cx="127000" cy="127000"/>"#,
                    'H' => r#"<a:chOff x="0" y="0"/>"#,
                    'X' => r#"<a:chExt cx="127000" cy="127000"/>"#,
                    _ => "<a:other/>",
                })
                .collect();
            let outcome = group_host_outcome(&format!(
                r#"<wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>{children}</a:xfrm></wpg:grpSpPr><w14:contentPart r:id="groupInk"/>"#
            ));
            (sequence, outcome)
        })
        .collect();
    assert_eq!(removable_sequences(&outcomes), ["OEHX"]);
    assert_eq!(outcome_counts(&outcomes), ["-:3905", "R:1"]);
}

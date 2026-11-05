//! `Paragraph::addressable_runs` (the `OfficeCLI` `r[N]` set) and
//! `OpaqueInline::kind`.

use crate::paragraph::{Inline, InlineKind, Paragraph};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn paragraph(inner: &str) -> Paragraph {
    Paragraph::new(
        format!(
            concat!(
                r#"<w:p xmlns:w="{}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" "#,
                r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">{}</w:p>"#
            ),
            W, inner
        )
        .into_bytes(),
    )
}

fn texts(paragraph: &Paragraph) -> Vec<String> {
    paragraph
        .addressable_runs()
        .unwrap()
        .iter()
        .map(|run| run.text().unwrap())
        .collect()
}

fn run(text: &str) -> String {
    format!("<w:r><w:t>{text}</w:t></w:r>")
}

#[test]
fn excluded_runs_follow_officecli() {
    let paragraph = paragraph(&[
        run("plain"),
        format!(r#"<w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt>"#, run("control")),
        format!(r#"<w:fldSimple w:instr="PAGE">{}</w:fldSimple>"#, run("field")),
        r#"<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>"#
            .to_owned(),
        format!(
            r#"<mc:AlternateContent><mc:Choice Requires="w14">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
            run("choice"),
            run("fallback")
        ),
        format!(
            r#"<w:r><w:t>host</w:t><w:pict><w:txbxContent><w:p>{}</w:p></w:txbxContent></w:pict></w:r>"#,
            run("boxed")
        ),
        run("tail"),
    ]
    .concat());
    assert_eq!(texts(&paragraph), ["plain", "hostboxed", "tail"]);
    // The outermost-run set is unchanged.
    assert_eq!(paragraph.runs().unwrap().len(), 8);
}

#[test]
fn wrapped_runs_are_counted() {
    let paragraph = paragraph(&[
        format!(r#"<w:hyperlink w:anchor="a">{}</w:hyperlink>"#, run("link")),
        format!(r#"<w:ins w:id="1" w:author="A">{}</w:ins>"#, run("inserted")),
        r#"<w:del w:id="2" w:author="A"><w:r><w:delText>deleted</w:delText></w:r></w:del>"#.to_owned(),
        format!(r#"<w:moveTo w:id="3" w:author="A">{}</w:moveTo>"#, run("moved")),
        format!(r#"<w:smartTag w:uri="u" w:element="e">{}</w:smartTag>"#, run("tag")),
        format!(r#"<w:customXml w:element="e">{}</w:customXml>"#, run("custom")),
        format!(r#"<w:dir w:val="rtl">{}</w:dir>"#, run("dir")),
        format!(r#"<w:bdo w:val="ltr">{}</w:bdo>"#, run("bdo")),
        format!(r#"<m:oMath><w:ins w:id="4" w:author="A">{}</w:ins></m:oMath>"#, run("math")),
        format!(
            r#"<w:hyperlink w:anchor="b"><w:sdt><w:sdtContent>{}</w:sdtContent></w:sdt>{}</w:hyperlink>"#,
            run("linked control"),
            run("after")
        ),
    ]
    .concat());
    // `Run::text` leaves out deleted text, so the deleted run reads empty.
    assert_eq!(
        texts(&paragraph),
        [
            "link", "inserted", "", "moved", "tag", "custom", "dir", "bdo", "math", "after"
        ]
    );
}

#[test]
fn phonetic_guide_runs_follow_their_host_in_document_order() {
    let paragraph = paragraph(concat!(
        r#"<w:r><w:t>a</w:t></w:r>"#,
        r#"<w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>kan</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢</w:t></w:r></w:rubyBase></w:ruby></w:r>"#,
        r#"<w:r><w:t>b</w:t></w:r>"#
    ));
    let runs = paragraph.addressable_runs().unwrap();
    assert_eq!(runs.len(), 5);
    let texts: Vec<String> = runs.iter().map(|run| run.text().unwrap()).collect();
    assert_eq!(texts[0], "a");
    assert_eq!(texts[2], "kan");
    assert_eq!(texts[3], "漢");
    assert_eq!(texts[4], "b");
    // The guide's runs are runs of their own, with their own properties.
    assert!(runs[2].xml_bytes().starts_with(b"<w:r>"));
    assert_eq!(paragraph.runs().unwrap().len(), 3);
}

#[test]
fn a_comment_reference_must_be_a_direct_child() {
    let paragraph = paragraph(concat!(
        r#"<w:r><mc:AlternateContent><mc:Choice><w:commentReference w:id="1"/></mc:Choice></mc:AlternateContent><w:t>kept</w:t></w:r>"#,
        r#"<w:r><w:t>also</w:t><w:commentReference w:id="2"/></w:r>"#,
        r#"<w:r xmlns:x="urn:x"><x:commentReference/><w:t>foreign</w:t></w:r>"#
    ));
    assert_eq!(texts(&paragraph), ["kept", "foreign"]);
}

#[test]
fn foreign_wrappers_do_not_exclude_but_alternatives_do_in_any_namespace() {
    let paragraph = paragraph(
        &[
            format!(r#"<x:sdt xmlns:x="urn:x">{}</x:sdt>"#, run("foreign sdt")),
            format!(
                r#"<x:AlternateContent xmlns:x="urn:x">{}</x:AlternateContent>"#,
                run("x alt")
            ),
            format!(r#"<w:Fallback>{}</w:Fallback>"#, run("w fallback")),
            r#"<x:r xmlns:x="urn:x"><x:t>not a run</x:t></x:r>"#.to_owned(),
        ]
        .concat(),
    );
    assert_eq!(texts(&paragraph), ["foreign sdt"]);
}

#[test]
fn addressable_runs_match_runs_without_excluded_markup() {
    let paragraph = paragraph(&[run("a"), "<w:r/>".to_owned(), run("b")].concat());
    let addressable = paragraph.addressable_runs().unwrap();
    let outermost = paragraph.runs().unwrap();
    assert_eq!(addressable.len(), 3);
    for (left, right) in addressable.iter().zip(outermost.iter()) {
        assert_eq!(left.xml_bytes(), right.xml_bytes());
    }
}

#[test]
fn addressable_runs_resolve_strict_and_cut_fragments() {
    let strict = Paragraph::new(
        br#"<s:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:r><s:t>a</s:t></s:r><s:fldSimple><s:r><s:t>b</s:t></s:r></s:fldSimple></s:p>"#
            .to_vec(),
    );
    assert_eq!(texts(&strict), ["a"]);
    let cut = Paragraph::new(
        br#"<w:p><w:r><w:t>a</w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>b</w:t></w:r></w:sdtContent></w:sdt></w:p>"#
            .to_vec(),
    );
    assert_eq!(texts(&cut), ["a"]);
}

#[test]
fn addressable_runs_refuse_malformed_paragraphs() {
    let not_a_paragraph =
        Paragraph::new(format!(r#"<w:r xmlns:w="{W}"><w:t>x</w:t></w:r>"#).into_bytes());
    assert!(not_a_paragraph.addressable_runs().is_err());
    assert!(
        Paragraph::new(format!(r#"<w:p xmlns:w="{W}"><w:r>"#).into_bytes())
            .addressable_runs()
            .is_err()
    );
    assert!(
        Paragraph::new(format!(r#"<w:p xmlns:w="{W}"/><w:p xmlns:w="{W}"/>"#).into_bytes())
            .addressable_runs()
            .is_err()
    );
    let empty = Paragraph::new(format!(r#"<w:p xmlns:w="{W}"/>"#).into_bytes());
    assert!(empty.addressable_runs().unwrap().is_empty());
}

#[test]
fn opaque_inlines_report_their_kind() {
    let paragraph = paragraph(concat!(
        r#"<w:pPr/><w:hyperlink w:anchor="a"/><w:fldSimple w:instr="PAGE"/><w:subDoc/><w:customXml w:element="e"/>"#,
        r#"<w:smartTag w:uri="u" w:element="e"/><w:sdt/><w:dir w:val="rtl"/><w:bdo w:val="ltr"/><w:proofErr w:type="spellStart"/>"#,
        r#"<w:permStart w:id="1"/><w:permEnd w:id="1"/><w:bookmarkStart w:name="no-id"/><w:bookmarkEnd/>"#,
        r#"<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/>"#,
        r#"<w:moveFromRangeStart w:id="1" w:name="m"/><w:moveFromRangeEnd w:id="1"/>"#,
        r#"<w:moveToRangeStart w:id="2" w:name="m"/><w:moveToRangeEnd w:id="2"/>"#,
        r#"<w:customXmlInsRangeStart w:id="1"/><w:customXmlInsRangeEnd w:id="1"/>"#,
        r#"<w:customXmlDelRangeStart w:id="1"/><w:customXmlDelRangeEnd w:id="1"/>"#,
        r#"<w:customXmlMoveFromRangeStart w:id="1"/><w:customXmlMoveFromRangeEnd w:id="1"/>"#,
        r#"<w:customXmlMoveToRangeStart w:id="1"/><w:customXmlMoveToRangeEnd w:id="1"/>"#,
        r#"<w:ins/><w:del/><w:moveFrom/><w:moveTo/><m:oMath/><m:oMathPara/><mc:AlternateContent/>"#,
        r#"<w:futureElement/><x:sdt xmlns:x="urn:x"/><w:bookmarkStart w:id="5" w:name="typed"/><w:r/>"#
    ));
    let inlines = paragraph.inlines().unwrap();
    let kinds: Vec<InlineKind> = inlines
        .iter()
        .filter_map(|inline| match inline {
            Inline::Unknown(opaque) => Some(opaque.kind()),
            _ => None,
        })
        .collect();
    assert_eq!(
        kinds,
        [
            InlineKind::Hyperlink,
            InlineKind::SimpleField,
            InlineKind::SubDocument,
            InlineKind::CustomXml,
            InlineKind::SmartTag,
            InlineKind::ContentControl,
            InlineKind::BidiEmbedding,
            InlineKind::BidiOverride,
            InlineKind::ProofError,
            InlineKind::PermissionStart,
            InlineKind::PermissionEnd,
            InlineKind::BookmarkStart,
            InlineKind::BookmarkEnd,
            InlineKind::CommentRangeStart,
            InlineKind::CommentRangeEnd,
            InlineKind::MoveFromRangeStart,
            InlineKind::MoveFromRangeEnd,
            InlineKind::MoveToRangeStart,
            InlineKind::MoveToRangeEnd,
            InlineKind::CustomXmlInsertionRangeStart,
            InlineKind::CustomXmlInsertionRangeEnd,
            InlineKind::CustomXmlDeletionRangeStart,
            InlineKind::CustomXmlDeletionRangeEnd,
            InlineKind::CustomXmlMoveFromRangeStart,
            InlineKind::CustomXmlMoveFromRangeEnd,
            InlineKind::CustomXmlMoveToRangeStart,
            InlineKind::CustomXmlMoveToRangeEnd,
            InlineKind::Insertion,
            InlineKind::Deletion,
            InlineKind::MoveFrom,
            InlineKind::MoveTo,
            InlineKind::Math,
            InlineKind::MathParagraph,
            InlineKind::AlternateContent,
            InlineKind::Other,
            InlineKind::Other,
        ]
    );
    for kind in &kinds[..kinds.len() - 2] {
        assert!(kind.local_name().is_some(), "{kind:?}");
    }
    assert_eq!(InlineKind::Other.local_name(), None);
    assert!(matches!(inlines.last(), Some(Inline::Run(_))));
    assert!(
        inlines
            .iter()
            .any(|inline| matches!(inline, Inline::BookmarkStart(bookmark) if bookmark.id() == 5))
    );
}

#[test]
fn math_and_alternatives_are_classified_by_namespace_or_conventional_prefix() {
    let paragraph = Paragraph::new(
        format!(
            concat!(
                r#"<w:p xmlns:w="{}" xmlns:o="http://schemas.openxmlformats.org/officeDocument/2006/math" "#,
                r#"xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x">"#,
                r#"<o:oMath/><c:AlternateContent/><x:oMath/><x:AlternateContent/></w:p>"#
            ),
            W
        )
        .into_bytes(),
    );
    let kinds: Vec<InlineKind> = paragraph
        .inlines()
        .unwrap()
        .iter()
        .filter_map(|inline| match inline {
            Inline::Unknown(opaque) => Some(opaque.kind()),
            _ => None,
        })
        .collect();
    assert_eq!(
        kinds,
        [
            InlineKind::Math,
            InlineKind::AlternateContent,
            InlineKind::Other,
            InlineKind::Other
        ]
    );

    // In a fragment cut from a part the conventional prefixes are undeclared.
    let cut = Paragraph::new(
        br#"<w:p><m:oMathPara/><mc:AlternateContent/><w:hyperlink/><q:hyperlink/></w:p>"#.to_vec(),
    );
    let kinds: Vec<InlineKind> = cut
        .inlines()
        .unwrap()
        .iter()
        .filter_map(|inline| match inline {
            Inline::Unknown(opaque) => Some(opaque.kind()),
            _ => None,
        })
        .collect();
    assert_eq!(
        kinds,
        [
            InlineKind::MathParagraph,
            InlineKind::AlternateContent,
            InlineKind::Hyperlink,
            InlineKind::Other
        ]
    );
}

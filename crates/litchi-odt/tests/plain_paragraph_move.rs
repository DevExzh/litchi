use litchi_odf_common::package::raw_identical_members;
use litchi_odt::{
    Document,
    core::{PackageWriter, Profile},
    protection::Policy,
    transaction::{DurablePatch, Position},
};

mod support;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn content(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:scripts/><o:automatic-styles/><o:body><o:text>{body}</o:text></o:body></o:document-content>"#
    )
}

fn snapshot(body: &str) -> litchi_odt::transaction::Snapshot {
    let xml = content(body);
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", xml.as_bytes()),
            ("meta.xml", b"<meta>opaque producer bytes</meta>"),
            ("Pictures/opaque.bin", &[0x5a; 512]),
        ],
    ))
    .unwrap()
    .snapshot()
    .unwrap()
}

#[test]
fn moves_exact_plain_fragments_and_raw_preserves_every_other_member() {
    let source =
        snapshot("\n  <t:p>alpha&amp;one</t:p>\n\t<t:p><![CDATA[beta<two>]]></t:p>\n  <t:p/>\n");
    let source_content = source.document().unwrap().get_file("content.xml").unwrap();
    let first = b"<t:p>alpha&amp;one</t:p>";
    let second = b"<t:p><![CDATA[beta<two>]]></t:p>";
    let third = b"<t:p/>";

    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let commit = edit.commit().unwrap();
    let target_content = commit
        .snapshot()
        .document()
        .unwrap()
        .get_file("content.xml")
        .unwrap();
    assert_eq!(target_content.len(), source_content.len());
    let first_at = |haystack: &[u8], needle: &[u8]| {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
            .unwrap()
    };
    assert!(first_at(&target_content, second) < first_at(&target_content, third));
    assert!(first_at(&target_content, third) < first_at(&target_content, first));
    for fragment in [first.as_slice(), second.as_slice(), third.as_slice()] {
        assert_eq!(
            source_content
                .windows(fragment.len())
                .filter(|window| *window == fragment)
                .count(),
            1
        );
        assert_eq!(
            target_content
                .windows(fragment.len())
                .filter(|window| *window == fragment)
                .count(),
            1
        );
    }

    let identical = raw_identical_members(source.as_bytes(), commit.snapshot().as_bytes()).unwrap();
    assert!(!identical.contains("content.xml"));
    for path in [
        "mimetype",
        "meta.xml",
        "Pictures/opaque.bin",
        "META-INF/manifest.xml",
    ] {
        assert!(identical.contains(path), "{path}");
    }
    let texts = commit
        .snapshot()
        .document()
        .unwrap()
        .paragraphs()
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(texts, ["beta<two>", "", "alpha&one"]);
}

#[test]
fn patch_is_exact_deterministic_reversible_and_stale_checked() {
    let source = snapshot("<t:p>zero</t:p><t:p>one</t:p><t:p>two</t:p>");
    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(2), Position::new(0))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.patch().apply(&source).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
    assert!(commit.patch().apply(commit.snapshot()).is_err());

    let durable = commit.patch().durable().unwrap();
    let json = durable.to_deterministic_json().unwrap();
    assert_eq!(json, durable.to_deterministic_json().unwrap());
    let decoded = DurablePatch::from_deterministic_json(&json).unwrap();
    assert_eq!(
        decoded.apply(&source).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        decoded
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
    assert!(decoded.apply(commit.snapshot()).is_err());
}

#[test]
fn equal_positions_are_an_exact_noop_without_staging() {
    let source = snapshot("<t:p>zero</t:p><t:p>one</t:p>");
    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(1), Position::new(1))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().as_bytes(), source.as_bytes());
    assert!(commit.results().is_empty());
}

#[test]
fn paragraph_move_conflicts_with_intervening_content_edits() {
    let source = snapshot("<t:p>zero</t:p><t:p>one</t:p><t:p>two</t:p>");
    let mut movement = source.edit();
    movement
        .move_plain_paragraph(Position::new(0), Position::new(2))
        .unwrap();
    let mut replacement = source.edit();
    replacement
        .replace_paragraph(Position::new(1), "changed")
        .unwrap();
    let mut joined = source.joined_edit();
    joined.join("move", movement).unwrap();
    assert!(joined.join("replace-middle", replacement).is_err());
}

#[test]
fn refuses_structural_rich_tracked_unknown_and_mce_bodies() {
    let cases = [
        "<t:p>rich<t:span>span</t:span></t:p><t:p>tail</t:p>",
        "<t:section><t:p>nested</t:p></t:section><t:p>tail</t:p>",
        "<t:list><t:list-item><t:p>nested</t:p></t:list-item></t:list><t:p>tail</t:p>",
        "<table:table xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"/><t:p>one</t:p><t:p>two</t:p>",
        "<t:tracked-changes/><t:p>one</t:p><t:p>two</t:p>",
        "<unknown xmlns=\"urn:unknown\"/><t:p>one</t:p><t:p>two</t:p>",
        "<t:p t:style-name=\"P1\">styled</t:p><t:p>tail</t:p>",
        "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"/><t:p>one</t:p><t:p>two</t:p>",
        "<t:p xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\">one</t:p><t:p>two</t:p>",
    ];
    for body in cases {
        let source = snapshot(body);
        let mut edit = source.edit();
        edit.move_plain_paragraph(Position::new(0), Position::new(1))
            .unwrap();
        assert!(edit.commit().is_err(), "{body}");
    }
}

#[test]
fn refuses_duplicate_or_stray_empty_body_ownership() {
    for xml in [
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:body/><o:body><o:text><t:p>one</t:p><t:p>two</t:p></o:text></o:body></o:document-content>"#
        ),
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:text/><o:body><o:text><t:p>one</t:p><t:p>two</t:p></o:text></o:body></o:document-content>"#
        ),
    ] {
        let document = match Document::from_bytes(support::package(
            MIMETYPE,
            &[("content.xml", xml.as_bytes())],
        )) {
            Ok(document) => document,
            Err(litchi_core::Error::InvalidFormat(_)) => continue,
            Err(error) => panic!("open produced a non-format error: {error:?}"),
        };
        let source = document.snapshot().unwrap();
        assert_move_refused(&source);
    }
}

#[test]
fn refuses_scripts_signatures_and_protection() {
    let scripted_xml = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><o:scripts><o:script s:language="x"><![CDATA[payload]]></o:script></o:scripts><o:body><o:text><t:p>one</t:p><t:p>two</t:p></o:text></o:body></o:document-content>"#
    );
    let scripted = Document::from_bytes(support::package(
        MIMETYPE,
        &[("content.xml", scripted_xml.as_bytes())],
    ))
    .unwrap()
    .snapshot()
    .unwrap();
    assert_move_refused(&scripted);

    let signed_xml = content("<t:p>one</t:p><t:p>two</t:p>");
    let signed = Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", signed_xml.as_bytes()),
            ("META-INF/documentsignatures.xml", b"opaque signature"),
        ],
    ))
    .unwrap()
    .snapshot()
    .unwrap();
    let mut signed_noop = signed.edit();
    signed_noop
        .move_plain_paragraph(Position::new(0), Position::new(0))
        .unwrap();
    assert_eq!(
        signed_noop.commit().unwrap().snapshot().as_bytes(),
        signed.as_bytes()
    );
    assert_move_refused(&signed);

    let mut encrypted_writer = PackageWriter::new();
    encrypted_writer.set_mimetype(MIMETYPE).unwrap();
    encrypted_writer
        .set_encryption("secret", Profile::compatible())
        .unwrap();
    encrypted_writer
        .add_file(
            "content.xml",
            content("<t:p>one</t:p><t:p>two</t:p>").as_bytes(),
        )
        .unwrap();
    let encrypted =
        Document::from_bytes_with_password(encrypted_writer.finish_to_bytes().unwrap(), "secret")
            .unwrap()
            .snapshot()
            .unwrap();
    assert_move_refused(&encrypted);

    let source = snapshot("<t:p>one</t:p><t:p>two</t:p>");
    let mut protection = source.edit();
    protection
        .set_protection(&Policy::default().with_read_only(Some(true)))
        .unwrap();
    let protected = protection.commit().unwrap().into_snapshot();
    assert_move_refused(&protected);
}

#[test]
fn checks_position_and_paragraph_bounds() {
    let source = snapshot("<t:p>one</t:p><t:p>two</t:p>");
    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(2), Position::new(0))
        .unwrap();
    assert!(edit.commit().is_err());

    let mut over_limit = source.edit();
    assert!(
        over_limit
            .move_plain_paragraph(Position::new(4_096), Position::new(0))
            .is_err()
    );

    let body = "<t:p>x</t:p>".repeat(4_097);
    let source = snapshot(&body);
    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    assert!(edit.commit().is_err());
}

fn assert_move_refused(source: &litchi_odt::transaction::Snapshot) {
    let mut edit = source.edit();
    edit.move_plain_paragraph(Position::new(0), Position::new(1))
        .unwrap();
    assert!(edit.commit().is_err());
}

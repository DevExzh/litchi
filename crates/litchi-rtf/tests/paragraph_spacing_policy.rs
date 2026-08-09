#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{ParagraphSpacingPolicy, RtfDocument, RtfWriter, StyleBlock};
fn block<'a>(d: &'a RtfDocument<'a>, s: &str) -> &'a StyleBlock<'a> {
    d.blocks().iter().find(|b| b.text.contains(s)).unwrap()
}

#[test]
fn parses_inherits_resets_and_keeps_destinations_inert() {
    let d=RtfDocument::parse(concat!(r#"{\rtf1\ansi\pard\sb120\sa240\lisb25\lisa50\sbauto1\saauto1\nosnaplinegrid\contextualspace Outer\par "#,r#"{\sbauto0\saauto0\lisb0\lisa0 Inner\par }Tail\par "#,r#"{\pard Reset\par }{\*\unknown\sbauto0\saauto0\lisb1\lisa1 Ignored}Visible\par}"#)).unwrap();
    let outer = block(&d, "Outer").paragraph.spacing_policy;
    assert_eq!(
        outer,
        ParagraphSpacingPolicy {
            automatic_before: true,
            automatic_after: true,
            list_before: Some(25),
            list_after: Some(50),
            snap_to_line_grid: false,
            contextual_spacing: true
        }
    );
    let inner = block(&d, "Inner").paragraph.spacing_policy;
    assert!(!inner.automatic_before && !inner.automatic_after);
    assert_eq!(inner.list_before, Some(0));
    assert_eq!(block(&d, "Tail").paragraph.spacing_policy, outer);
    assert_eq!(
        block(&d, "Reset").paragraph.spacing_policy,
        ParagraphSpacingPolicy::default()
    );
    assert_eq!(block(&d, "Visible").paragraph.spacing_policy, outer);
}

#[test]
fn stylesheet_and_deterministic_writer_round_trip() {
    let d=RtfDocument::parse(r"{\rtf1{\stylesheet{\s8\sb100\sa200\lisb25\lisa50\sbauto1\saauto1\nosnaplinegrid\contextualspace Spaced;}}\pard\lisb25\lisa50\sbauto1\saauto1\nosnaplinegrid\contextualspace Body\par}").unwrap();
    let expected = d
        .stylesheet()
        .get(8)
        .unwrap()
        .paragraph
        .unwrap()
        .spacing_policy;
    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&d).unwrap();
    let text = String::from_utf8(first.clone()).unwrap();
    assert!(text.contains(r"\lisb25\lisa50\sbauto1\saauto1\nosnaplinegrid\contextualspace"));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        reparsed
            .stylesheet()
            .get(8)
            .unwrap()
            .paragraph
            .unwrap()
            .spacing_policy,
        expected
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_real_libreoffice_stylesheet_fixture() {
    let bytes = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf107480.rtf"
    );
    let marker = br"{\stylesheet";
    let start = bytes
        .windows(marker.len())
        .position(|w| w == marker)
        .unwrap();
    let mut depth = 0;
    let mut end = None;
    for (i, b) in bytes[start..].iter().enumerate() {
        match b {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(start + i + 1);
                    break;
                }
            },
            _ => {},
        }
    }
    let mut isolated = br"{\rtf1\ansi".to_vec();
    isolated.extend_from_slice(&bytes[start..end.unwrap()]);
    isolated.push(b'}');
    let d = RtfDocument::parse_bytes(&isolated).unwrap();
    let p = d
        .stylesheet()
        .get(3)
        .unwrap()
        .paragraph
        .unwrap()
        .spacing_policy;
    assert!(p.automatic_before && p.automatic_after);
    let list_fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo77996.rtf"
    );
    let sequence = br"\lisb0 \sa100 \lisa0";
    assert!(
        list_fixture
            .windows(sequence.len())
            .any(|window| window == sequence)
    );
}

#[test]
fn rejects_missing_out_of_range_and_selector_parameters() {
    for s in [
        r"{\rtf1\sbauto X}",
        r"{\rtf1\sbauto2 X}",
        r"{\rtf1\saauto-1 X}",
        r"{\rtf1\lisb X}",
        r"{\rtf1\lisb-1 X}",
        r"{\rtf1\lisa1000001 X}",
        r"{\rtf1\nosnaplinegrid0 X}",
        r"{\rtf1\contextualspace1 X}",
    ] {
        assert!(RtfDocument::parse(s).is_err(), "accepted {s}");
    }
}

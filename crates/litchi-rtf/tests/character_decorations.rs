#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{CharacterBorderStyle, RtfDocument, RtfWriter};

#[test]
fn parses_libreoffice_character_border_and_shading_fixtures() {
    let border_source =
        std::fs::read_to_string(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/hidden-para-separator.rtf",
        ))
        .unwrap();
    let border_document = RtfDocument::parse(&border_source).unwrap();
    assert!(border_document.runs().iter().any(|run| {
        run.text().contains('C')
            && run.formatting.character_border.is_some_and(|border| {
                border.style == CharacterBorderStyle::Single && border.width == 10
            })
    }));

    let shading_source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/165717.rtf"),
    )
    .unwrap();
    let shading_document = RtfDocument::parse(&shading_source).unwrap();
    assert!(shading_document.stylesheet().styles().iter().any(|style| {
        style
            .formatting
            .character_shading
            .is_some_and(|shading| shading.background_color == 8)
    }));
}

#[test]
fn inherits_resets_writes_and_keeps_ignored_destinations_inert() {
    let source = concat!(
        r#"{\rtf1\ansi\chbrdr\brdrs\brdrw10\brdrcf2\brsp3\brdrsh\brdrframe"#,
        r#"\chshdng2500\chcfpat3\chcbpat4 Outer"#,
        r#"{\chbrdr\brdrdb\brdrw20\brdrcf5\brsp6 Inner}"#,
        r#"Tail{\plain Plain}"#,
        r#"{\*\unknown\chbrdr\brdrs\brdrw75\chshdng10000 ignored}Visible}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let blocks = document.blocks();
    let outer = blocks
        .iter()
        .find(|block| block.text.contains("Outer"))
        .unwrap();
    let inner = blocks
        .iter()
        .find(|block| block.text.contains("Inner"))
        .unwrap();
    let tail = blocks
        .iter()
        .find(|block| block.text.contains("Tail"))
        .unwrap();
    let plain = blocks
        .iter()
        .find(|block| block.text.contains("Plain"))
        .unwrap();
    let visible = blocks
        .iter()
        .find(|block| block.text.contains("Visible"))
        .unwrap();

    let outer_border = outer.formatting.character_border.unwrap();
    assert_eq!(outer_border.style, CharacterBorderStyle::Single);
    assert_eq!(outer_border.width, 10);
    assert!(outer_border.shadow && outer_border.frame);
    assert_eq!(outer.formatting.character_shading.unwrap().amount, 2500);
    assert_eq!(
        inner.formatting.character_border.unwrap().style,
        CharacterBorderStyle::Double
    );
    assert_eq!(inner.formatting.character_border.unwrap().width, 20);
    assert_eq!(tail.formatting.character_border, Some(outer_border));
    assert!(plain.formatting.character_border.is_none());
    assert!(plain.formatting.character_shading.is_none());
    assert_eq!(visible.formatting.character_border, Some(outer_border));

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let written = String::from_utf8(output).unwrap();
    assert!(written.contains(
        r"\chbrdr\brdrs\brdrw10\brdrcf2\brsp3\brdrsh\brdrframe\chshdng2500\chcfpat3\chcbpat4"
    ));
    let reparsed = RtfDocument::parse(&written).unwrap();
    assert!(reparsed.blocks().iter().any(|block| {
        block.text.contains("Inner")
            && block.formatting.character_border.is_some_and(|border| {
                border.style == CharacterBorderStyle::Double && border.width == 20
            })
    }));
}

#[test]
fn rejects_malformed_character_decorations() {
    for source in [
        r"{\rtf1\chbrdr1\brdrs X}",
        r"{\rtf1\chbrdr\brdrs\brdrs X}",
        r"{\rtf1\chbrdr\brdrw X}",
        r"{\rtf1\chbrdr\brdrw-1 X}",
        r"{\rtf1\chbrdr\brdrw76 X}",
        r"{\rtf1\chbrdr\brdrcf-1 X}",
        r"{\rtf1\chbrdr\brdrcf65536 X}",
        r"{\rtf1\chshdng X}",
        r"{\rtf1\chshdng-1 X}",
        r"{\rtf1\chshdng10001 X}",
        r"{\rtf1\chcfpat65536 X}",
        r"{\rtf1\chcbpat-1 X}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    assert!(RtfDocument::parse(r"{\rtf1{\*\unknown\chbrdr1\chshdng10001 ignored}Visible}").is_ok());
}

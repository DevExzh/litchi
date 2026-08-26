#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    AssociatedCharacterBaseline, AssociatedCharacterFormatting, CharacterBaseline,
    CharacterExpansion, CharacterPositioning, RtfDocument, RtfWriter,
};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a litchi_rtf::StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap()
}

#[test]
fn parses_inherits_resets_and_keeps_destinations_inert() {
    let source = r"{\rtf1\ansi\super\expnd4\charscalex80\kerning16 Outer{\dn3\expndtw20 Inner}{Tail}\nosupersub Normal{\up2 Raised}{\plain Plain}{\*\unknown\up999999\expndtw999999 ignored}Visible}";
    let document = RtfDocument::parse(source).unwrap();
    let outer = block(&document, "Outer");
    assert_eq!(
        outer.formatting.character_positioning.baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        outer.formatting.character_positioning.expansion,
        CharacterExpansion::QuarterPoints(4)
    );
    assert_eq!(
        outer
            .formatting
            .character_positioning
            .horizontal_scale_percent,
        80
    );
    assert_eq!(
        outer.formatting.character_positioning.kerning_half_points,
        16
    );
    assert_eq!(
        block(&document, "Inner")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::LoweredHalfPoints(3)
    );
    assert_eq!(
        block(&document, "Inner")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::Twips(20)
    );
    assert_eq!(
        block(&document, "Tail")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        block(&document, "Normal")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Normal
    );
    assert_eq!(
        block(&document, "Raised")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::RaisedHalfPoints(2)
    );
    assert_eq!(
        block(&document, "Plain").formatting.character_positioning,
        CharacterPositioning::default()
    );
    assert_eq!(
        block(&document, "Visible")
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Normal
    );
}

#[test]
fn wire_positioning_boundaries_and_explicit_expansion_resets_round_trip() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\expnd-31680\charscalex1\kerning0 Quarter}{\expndtw31680\charscalex600\kerning32767 Twips}{\expnd4 Active\expnd0 Reset\expndtw0 TwipReset}{\charscalex100 ScaleReset}}",
    )
    .unwrap();
    assert_eq!(
        block(&document, "Quarter").formatting.character_positioning,
        CharacterPositioning {
            baseline: CharacterBaseline::Normal,
            expansion: CharacterExpansion::QuarterPoints(-31680),
            horizontal_scale_percent: 1,
            kerning_half_points: 0,
        }
    );
    assert_eq!(block(&document, "Quarter").formatting.char_spacing, -31680);
    assert_eq!(block(&document, "Quarter").formatting.char_scale, 1);
    assert_eq!(block(&document, "Quarter").formatting.kerning, 0);
    assert_eq!(
        block(&document, "Twips").formatting.character_positioning,
        CharacterPositioning {
            baseline: CharacterBaseline::Normal,
            expansion: CharacterExpansion::Twips(31680),
            horizontal_scale_percent: 600,
            kerning_half_points: 32767,
        }
    );
    assert_eq!(block(&document, "Twips").formatting.char_spacing, 31680);
    assert_eq!(block(&document, "Twips").formatting.char_scale, 600);
    assert_eq!(block(&document, "Twips").formatting.kerning, 32767);
    assert_eq!(
        block(&document, "Reset")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::None
    );
    assert_eq!(block(&document, "Reset").formatting.char_spacing, 0);
    assert_eq!(
        block(&document, "TwipReset")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::None
    );
    assert_eq!(block(&document, "TwipReset").formatting.char_spacing, 0);
    assert_eq!(
        block(&document, "ScaleReset")
            .formatting
            .character_positioning
            .horizontal_scale_percent,
        100
    );
    assert_eq!(block(&document, "ScaleReset").formatting.char_scale, 100);

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reopened = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        block(&reopened, "Quarter").formatting.character_positioning,
        block(&document, "Quarter").formatting.character_positioning
    );
    assert_eq!(
        block(&reopened, "Twips").formatting.character_positioning,
        block(&document, "Twips").formatting.character_positioning
    );
    assert_eq!(
        block(&reopened, "Reset")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::None
    );
    assert_eq!(block(&reopened, "Reset").formatting.char_spacing, 0);
    assert_eq!(
        block(&reopened, "TwipReset")
            .formatting
            .character_positioning
            .expansion,
        CharacterExpansion::None
    );
    assert_eq!(block(&reopened, "TwipReset").formatting.char_spacing, 0);
    assert_eq!(
        block(&reopened, "ScaleReset")
            .formatting
            .character_positioning
            .horizontal_scale_percent,
        100
    );
    assert_eq!(block(&reopened, "ScaleReset").formatting.char_scale, 100);
}

#[test]
fn writer_is_deterministic_and_preserves_units() {
    let document =
        RtfDocument::parse(r"{\rtf1 A{\up2\expndtw-15\charscalex75\kerning8 B}{\sub\expnd3 C}}")
            .unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first_text = String::from_utf8(first).unwrap();
    assert!(first_text.contains("\\up2"));
    assert!(first_text.contains("\\expndtw-15"));
    let reparsed = RtfDocument::parse(&first_text).unwrap();
    assert_eq!(
        block(&reparsed, "B").formatting.character_positioning,
        block(&document, "B").formatting.character_positioning
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first_text, String::from_utf8(second).unwrap());
}

#[test]
fn parses_libreoffice_superscript_fixture() {
    let source = std::fs::read_to_string(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf87034.rtf"),
    )
    .unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    assert!(
        document
            .runs()
            .iter()
            .any(|run| run.formatting.character_positioning.baseline
                == CharacterBaseline::Superscript)
    );
}

#[test]
fn rejects_out_of_range_parameters() {
    for source in [
        r"{\rtf1\up-1 X}",
        r"{\rtf1\up31681 X}",
        r"{\rtf1\dn-1 X}",
        r"{\rtf1\expnd31681 X}",
        r"{\rtf1\expndtw-31681 X}",
        r"{\rtf1\charscalex0 X}",
        r"{\rtf1\charscalex601 X}",
        r"{\rtf1\kerning-1 X}",
        r"{\rtf1\kerning32768 X}",
        r"{\rtf1\nosupersub1 X}",
        r"{\rtf1\ansi{\header\nosupersub1 H}Body}",
        r"{\rtf1 A{\footnote\nosupersub1 N}B}",
        r"{\rtf1{\stylesheet{\s0\nosupersub1 Normal;}}Body}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(RtfDocument::parse(r"{\rtf1 A{\*\unknown\nosupersub1 hidden}B}").is_ok());
    assert!(
        RtfDocument::parse(
            r"{\rtf1{\stylesheet{\s0 Normal;}{\*\futurestyle\nosupersub1 ignored;}}Body}"
        )
        .is_ok()
    );
}

#[test]
fn typed_zero_baseline_offsets_are_invalid_but_wire_zero_is_normal() {
    for baseline in [
        CharacterBaseline::RaisedHalfPoints(0),
        CharacterBaseline::LoweredHalfPoints(0),
    ] {
        let positioning = CharacterPositioning {
            baseline,
            ..CharacterPositioning::default()
        };
        assert!(positioning.validate().is_err());
    }
    for baseline in [
        AssociatedCharacterBaseline::RaisedHalfPoints(0),
        AssociatedCharacterBaseline::LoweredHalfPoints(0),
    ] {
        let mut formatting = AssociatedCharacterFormatting::default();
        assert!(formatting.set_baseline(Some(baseline)).is_err());
        formatting.baseline = Some(baseline);
        assert!(formatting.validate().is_err());
    }
    for source in [r"{\rtf1\up0 X}", r"{\rtf1\dn0 X}"] {
        let document = RtfDocument::parse(source).unwrap();
        assert_eq!(
            block(&document, "X")
                .formatting
                .character_positioning
                .baseline,
            CharacterBaseline::Normal
        );
    }
    for source in [r"{\rtf1\aup0 X}", r"{\rtf1\adn0 X}"] {
        let document = RtfDocument::parse(source).unwrap();
        assert_eq!(block(&document, "X").formatting.associated.baseline, None);
    }
}

#[test]
fn non_body_stories_preserve_one_baseline_and_refuse_mixed_states() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\header {\super\aup2 Head}}A{\footnote {\sub\adn2 Note}}B}",
    )
    .unwrap();
    assert_eq!(
        document.sections()[0].headers_footers[0].paragraphs[0]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        document.notes()[0]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Subscript
    );
    assert_eq!(
        document.sections()[0].headers_footers[0].paragraphs[0]
            .formatting
            .associated
            .baseline,
        Some(AssociatedCharacterBaseline::RaisedHalfPoints(2))
    );
    assert_eq!(
        document.notes()[0].formatting.associated.baseline,
        Some(AssociatedCharacterBaseline::LoweredHalfPoints(2))
    );
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reopened = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reopened.sections()[0].headers_footers[0].paragraphs[0]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        reopened.notes()[0]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Subscript
    );
    assert_eq!(
        reopened.sections()[0].headers_footers[0].paragraphs[0]
            .formatting
            .associated
            .baseline,
        Some(AssociatedCharacterBaseline::RaisedHalfPoints(2))
    );
    assert_eq!(
        reopened.notes()[0].formatting.associated.baseline,
        Some(AssociatedCharacterBaseline::LoweredHalfPoints(2))
    );

    for source in [
        r"{\rtf1\ansi{\header\super Head\nosupersub Tail}Body}",
        r"{\rtf1 A{\footnote\sub One\par\nosupersub Two}B}",
        r"{\rtf1 A{\endnote\up2 One\dn2 Two}B}",
        r"{\rtf1\ansi{\header\aup2 One\aup4 Two}Body}",
        r"{\rtf1 A{\footnote\adn2 One\par\adn4 Two}B}",
        r"{\rtf1\ansi{\header\super One\line\sub Two}Body}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let representable =
        RtfDocument::parse(r"{\rtf1\ansi{\header\super One\par\nosupersub Two}Body}").unwrap();
    assert_eq!(
        representable.sections()[0].headers_footers[0].paragraphs[0]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Superscript
    );
    assert_eq!(
        representable.sections()[0].headers_footers[0].paragraphs[1]
            .formatting
            .character_positioning
            .baseline,
        CharacterBaseline::Normal
    );
}

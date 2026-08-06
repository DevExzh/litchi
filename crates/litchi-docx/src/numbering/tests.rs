//! Public API and malformed-input coverage for the WordprocessingML numbering owner.

use super::*;

use crate::{Error, Result};
use litchi_opc::PackURI;
use std::mem::size_of;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W12: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn parse(xml: &[u8]) -> Result<Collection> {
    parse_numbering(xml)
}

fn numbering_xml(definitions: &str) -> String {
    format!(
        r#"<w:numbering xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12">{definitions}</w:numbering>"#
    )
}

fn assert_rejected_without_panicking(xml: &[u8]) {
    let parsed = std::panic::catch_unwind(|| parse(xml));
    match parsed {
        Ok(Err(error)) => assert!(
            matches!(error, Error::Invalid(_) | Error::Xml(_)),
            "unexpected numbering error: {error}"
        ),
        Ok(Ok(_)) => panic!("malformed numbering XML was accepted"),
        Err(_) => panic!("numbering parser panicked on malformed XML"),
    }
}

#[test]
fn parses_complete_level_and_override() {
    let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="List"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/><w:lvlRestart w:val="0"/><w:isLgl/><w:pStyle w:val="ListParagraph"/><w:lvlPicBulletId w:val="7"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>"#).unwrap();
    let level = &value.abstract_nums()[0].levels()[0];
    assert_eq!(level.start, 3);
    assert_eq!(level.level_text.as_deref(), Some("%1."));
    assert_eq!(level.restart, Restart::Never);
    assert!(level.legal);
    assert_eq!(value.abstract_nums()[0].num_type(), Some(MultiLevel::Multi));
    assert_eq!(value.nums()[0].overrides()[0].start_override, Some(5));
}

#[test]
fn parses_word_2012_restart_policy_on_abstract_definitions() {
    let value = parse(
        numbering_xml(
            r#"<w:abstractNum w:abstractNumId="1" w12:restartNumberingAfterBreak="on"/><w:abstractNum w:abstractNumId="2" w12:restartNumberingAfterBreak="0"/><w:abstractNum w:abstractNumId="3"/>"#,
        )
        .as_bytes(),
    )
    .unwrap();

    assert_eq!(
        value
            .get_abstract_num(1)
            .unwrap()
            .restart_numbering_after_break(),
        Some(true)
    );
    assert_eq!(
        value
            .get_abstract_num(2)
            .unwrap()
            .restart_numbering_after_break(),
        Some(false)
    );
    assert_eq!(
        value
            .get_abstract_num(3)
            .unwrap()
            .restart_numbering_after_break(),
        None
    );

    for (lexical, expected) in [
        ("1", true),
        ("true", true),
        ("on", true),
        ("0", false),
        ("false", false),
        ("off", false),
    ] {
        let xml = numbering_xml(&format!(
            r#"<w:abstractNum w:abstractNumId="1" w12:restartNumberingAfterBreak="{lexical}"/>"#
        ));
        assert_eq!(
            parse(xml.as_bytes())
                .unwrap()
                .get_abstract_num(1)
                .unwrap()
                .restart_numbering_after_break(),
            Some(expected),
            "lexical ST_OnOff value {lexical}"
        );
    }
}

#[test]
fn numbering_snapshot_edits_only_the_extension_seam_and_round_trips() {
    let source_xml = numbering_xml(
        r#"<w:abstractNum w:abstractNumId="1" w:foo="keep" w12:restartNumberingAfterBreak='0'><x:future xmlns:x="urn:future" x:value="keep"/></w:abstractNum>"#,
    );
    let source = Snapshot::from_xml(source_xml.as_bytes().to_vec()).unwrap();
    assert_eq!(source.xml_bytes(), source_xml.as_bytes());
    assert_eq!(
        source.restart_numbering_after_break(1).unwrap(),
        Some(false)
    );

    let mut no_op = source.edit();
    no_op
        .set_restart_numbering_after_break(1, Some(false))
        .unwrap();
    let no_op_commit = no_op.commit().unwrap();
    assert_eq!(no_op_commit.snapshot().xml_bytes(), source.xml_bytes());

    let mut edit = source.edit();
    edit.set_restart_numbering_after_break(1, Some(true))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.xml_bytes(), source_xml.as_bytes());
    assert_eq!(
        commit.snapshot().restart_numbering_after_break(1).unwrap(),
        Some(true)
    );
    let changed = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(changed.contains(r#"w:foo="keep""#));
    assert!(changed.contains(r#"w12:restartNumberingAfterBreak='true'"#));
    assert!(changed.contains(r#"x:future xmlns:x="urn:future" x:value="keep""#));
    assert_eq!(
        commit.patch().before_restart_numbering_after_break(1),
        Some(Some(false))
    );
    assert_eq!(
        commit.patch().after_restart_numbering_after_break(1),
        Some(Some(true))
    );

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), source_xml.as_bytes());
    assert_eq!(
        restored.restart_numbering_after_break(1).unwrap(),
        Some(false)
    );
}

#[test]
fn numbering_snapshot_can_add_and_remove_a_prefixed_extension_without_reformatting() {
    let source_xml = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><x:future xmlns:x="urn:future"/></w:abstractNum></w:numbering>"#
    );
    let source = Snapshot::from_xml(source_xml.as_bytes().to_vec()).unwrap();
    let mut edit = source.edit();
    edit.set_restart_numbering_after_break(1, Some(false))
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
    assert!(
        changed.contains(r#"xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml""#)
    );
    assert!(
        changed
            .contains(r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#)
    );
    assert!(changed.contains(r#"mc:Ignorable="w15""#));
    assert!(changed.contains(r#"w15:restartNumberingAfterBreak="false""#));
    assert!(changed.contains(r#"<x:future xmlns:x="urn:future"/>"#));
    assert_eq!(
        commit.snapshot().restart_numbering_after_break(1).unwrap(),
        Some(false)
    );

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), source_xml.as_bytes());
    assert_eq!(restored.restart_numbering_after_break(1).unwrap(), None);
}

#[test]
fn numbering_snapshot_rejects_invalid_extension_scope_and_stale_edits_atomically() {
    let invalid_value = format!(
        r#"<w:numbering xmlns:w="{W}" xmlns:w12="{W12}" xmlns:mc="{MC}" mc:Ignorable="w12"><w:abstractNum w:abstractNumId="1" w12:restartNumberingAfterBreak="sometimes"/></w:numbering>"#
    );
    assert!(Snapshot::from_xml(invalid_value.into_bytes()).is_err());

    let missing_ignorable = format!(
        r#"<w:numbering xmlns:w="{W}" xmlns:w12="{W12}"><w:abstractNum w:abstractNumId="1" w12:restartNumberingAfterBreak="true"/></w:numbering>"#
    );
    assert!(Snapshot::from_xml(missing_ignorable.into_bytes()).is_err());

    let source =
        Snapshot::from_xml(numbering_xml(r#"<w:abstractNum w:abstractNumId="1"/>"#).into_bytes())
            .unwrap();
    assert!(
        source
            .edit()
            .set_restart_numbering_after_break(99, Some(true))
            .is_err()
    );

    let mut edit = source.edit();
    edit.set_restart_numbering_after_break(1, Some(true))
        .unwrap();
    let commit = edit.commit().unwrap();
    let alternate = Snapshot::from_xml(
        numbering_xml(r#"<w:abstractNum w:abstractNumId="1"/><x:future xmlns:x="urn:other"/>"#)
            .into_bytes(),
    )
    .unwrap();
    assert!(commit.patch().apply(&alternate).is_err());
    assert_eq!(source.restart_numbering_after_break(1).unwrap(), None);
}

#[test]
fn parses_vml_picture_bullet_definition() {
    let value = parse(br##"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="3"><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" xmlns:o="urn:schemas-microsoft-com:office:office"/><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:12pt;height:12pt" o:bullet="t"><v:imagedata r:id="rId4" o:title="bullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlPicBulletId w:val="3"/></w:lvl></w:abstractNum></w:numbering>"##).unwrap();
    assert_eq!(value.picture_bullets().len(), 1);
    let bullet = value.get_picture_bullet(3).expect("picture bullet 3");
    assert_eq!(bullet.id(), 3);
    assert_eq!(bullet.image_relationship_id(), Some("rId4"));
    assert!(value.get_picture_bullet(4).is_none());
    assert_eq!(
        value.abstract_nums()[0].levels()[0].picture_bullet_id,
        Some(3)
    );
}

#[test]
fn parses_drawingml_picture_bullet_definition() {
    let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="1"><w:pict><a:blip r:embed="rId9"/></w:pict></w:numPicBullet></w:numbering>"#).unwrap();
    let bullet = value.get_picture_bullet(1).expect("picture bullet 1");
    assert_eq!(bullet.image_relationship_id(), Some("rId9"));
}

#[test]
fn parses_picture_bullet_without_image() {
    let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet><w:numPicBullet w:numPicBulletId="2"/></w:numbering>"#).unwrap();
    assert_eq!(value.picture_bullets().len(), 2);
    assert_eq!(
        value.get_picture_bullet(0).unwrap().image_relationship_id(),
        None
    );
    assert_eq!(
        value.get_picture_bullet(2).unwrap().image_relationship_id(),
        None
    );
}

#[test]
fn rejects_duplicate_picture_bullet_ids() {
    let duplicate = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numPicBullet w:numPicBulletId="1"/><w:numPicBullet w:numPicBulletId="1"/></w:numbering>"#;
    assert!(parse(duplicate).is_err());
}

#[test]
fn rejects_malformed_numbering_order_without_panicking() {
    for xml in [
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:num w:numId="2"/></w:abstractNum></w:numbering>"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:num w:numId="2"/></w:lvl></w:abstractNum></w:numbering>"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="2"><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="1"/></w:lvlOverride><w:abstractNumId w:val="1"/></w:num></w:numbering>"#[..],
    ] {
        assert_rejected_without_panicking(xml);
    }
}

#[test]
fn rejects_truncated_numbering_states_without_panicking() {
    for xml in [
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numPicBullet w:numPicBulletId="1">"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="2"><w:lvlOverride w:ilvl="0">"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1">"#[..],
        &br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="2">"#[..],
    ] {
        assert_rejected_without_panicking(xml);
    }
}

#[test]
fn parses_libreoffice_picture_bullet_fixture() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let relative = "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx";
    let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
        .unwrap_or_else(|error| panic!("failed to open {relative}: {error}"));
    let uri = PackURI::new("/word/numbering.xml").expect("valid numbering URI");
    let bytes = package
        .blob_for(&uri)
        .unwrap_or_else(|error| panic!("failed to load numbering part: {error}"));
    let numbering = parse_numbering(bytes.as_ref()).unwrap();
    let bullet = numbering
        .get_picture_bullet(0)
        .expect("fixture defines picture bullet 0");
    // LibreOffice stripped the image payload from this fixture; only the
    // definition shell and the level linkage remain.
    assert_eq!(bullet.image_relationship_id(), None);
    let level = numbering
        .abstract_nums()
        .iter()
        .flat_map(|abstract_num| abstract_num.levels())
        .find(|level| level.picture_bullet_id.is_some())
        .expect("fixture level references a picture bullet");
    assert_eq!(level.picture_bullet_id, Some(0));
    assert!(
        numbering
            .get_picture_bullet(level.picture_bullet_id.unwrap())
            .is_some()
    );
}

#[test]
fn rejects_duplicate_levels_and_bad_level_indices() {
    let duplicate = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#;
    assert!(parse(duplicate).is_err());
    let bad = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="9"/></w:abstractNum></w:numbering>"#;
    assert!(parse(bad).is_err());
}

#[test]
fn strict_and_transitional_number_format_domains_are_exhaustive() {
    const TOKENS: [&str; 63] = [
        "decimal",
        "upperRoman",
        "lowerRoman",
        "upperLetter",
        "lowerLetter",
        "ordinal",
        "cardinalText",
        "ordinalText",
        "hex",
        "chicago",
        "ideographDigital",
        "japaneseCounting",
        "aiueo",
        "iroha",
        "decimalFullWidth",
        "decimalHalfWidth",
        "japaneseLegal",
        "japaneseDigitalTenThousand",
        "decimalEnclosedCircle",
        "decimalFullWidth2",
        "aiueoFullWidth",
        "irohaFullWidth",
        "decimalZero",
        "bullet",
        "ganada",
        "chosung",
        "decimalEnclosedFullstop",
        "decimalEnclosedParen",
        "decimalEnclosedCircleChinese",
        "ideographEnclosedCircle",
        "ideographTraditional",
        "ideographZodiac",
        "ideographZodiacTraditional",
        "taiwaneseCounting",
        "ideographLegalTraditional",
        "taiwaneseCountingThousand",
        "taiwaneseDigital",
        "chineseCounting",
        "chineseLegalSimplified",
        "chineseCountingThousand",
        "koreanDigital",
        "koreanCounting",
        "koreanLegal",
        "koreanDigital2",
        "vietnameseCounting",
        "russianLower",
        "russianUpper",
        "none",
        "numberInDash",
        "hebrew1",
        "hebrew2",
        "arabicAlpha",
        "arabicAbjad",
        "hindiVowels",
        "hindiConsonants",
        "hindiNumbers",
        "hindiCounting",
        "thaiLetters",
        "thaiNumbers",
        "thaiCounting",
        "bahtText",
        "dollarText",
        "custom",
    ];
    let mut values = std::collections::HashSet::new();
    for raw in TOKENS {
        let parsed = raw.parse::<Format>().unwrap();
        assert!(values.insert(parsed), "duplicate enum mapping for {raw}");
        assert_eq!(parsed.as_str(), raw);
        assert_eq!(parsed.to_string(), raw);

        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let xml = format!(
                r#"<w:numbering xmlns:w="{namespace}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="{raw}"/></w:lvl></w:abstractNum></w:numbering>"#
            );
            assert_eq!(
                parse(xml.as_bytes()).unwrap().abstract_nums()[0].levels()[0].format,
                parsed
            );
        }
    }
    assert_eq!(values.len(), 63);
    assert_eq!(Format::Custom as u8, 62);
    assert_eq!(size_of::<Format>(), 1);
    assert!("vendorNumbering".parse::<Format>().is_err());
    assert!("Decimal".parse::<Format>().is_err());

    for namespace in [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let xml = format!(
            r#"<w:numbering xmlns:w="{namespace}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="vendorNumbering"/></w:lvl></w:abstractNum></w:numbering>"#
        );
        assert!(parse(xml.as_bytes()).is_err());
    }
}

#[test]
fn multi_level_type_is_a_closed_compact_domain() {
    for (raw, expected) in [
        ("singleLevel", MultiLevel::Single),
        ("multilevel", MultiLevel::Multi),
        ("hybridMultilevel", MultiLevel::Hybrid),
    ] {
        assert_eq!(raw.parse(), Ok(expected));
        assert_eq!(expected.as_str(), raw);
        assert_eq!(expected.to_string(), raw);

        for namespace in [
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ] {
            let xml = format!(
                r#"<w:numbering xmlns:w="{namespace}"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="{raw}"/><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#
            );
            assert_eq!(
                parse(xml.as_bytes()).unwrap().abstract_nums()[0].num_type(),
                Some(expected)
            );
        }
    }
    assert!("multi".parse::<MultiLevel>().is_err());
    assert!("Multilevel".parse::<MultiLevel>().is_err());
    assert_eq!(size_of::<MultiLevel>(), 1);

    for namespace in [
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://purl.oclc.org/ooxml/wordprocessingml/main",
    ] {
        let xml = format!(
            r#"<w:numbering xmlns:w="{namespace}"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multi"/><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#
        );
        assert!(parse(xml.as_bytes()).is_err());
    }
}

#[test]
fn parses_poi_and_libreoffice_numbering_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for relative in [
        "test-data/poi/test-data/document/Numbering.docx",
        "test-data/poi/test-data/document/NumberingWOverrides.docx",
        "test-data/poi/test-data/document/ComplexNumberedLists.docx",
        "test-data/poi/test-data/document/NumberingWithOutOfOrderId.docx",
        "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/listWithLgl.docx",
        "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/decimal-numbering-no-leveltext.docx",
        "test-data/libreoffice-core/sw/qa/extras/ooxmlimport/data/numbering-circle.docx",
        "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
        "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx",
    ] {
        let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
            .unwrap_or_else(|error| panic!("failed to open {relative}: {error}"));
        let uri = PackURI::new("/word/numbering.xml").expect("valid numbering URI");
        let bytes = package
            .blob_for(&uri)
            .unwrap_or_else(|error| panic!("failed to load numbering part in {relative}: {error}"));
        let numbering = parse_numbering(bytes.as_ref())
            .unwrap_or_else(|error| panic!("failed to parse numbering in {relative}: {error}"));
        assert!(
            numbering.abstract_num_count() != 0,
            "fixture has no definitions: {relative}"
        );
    }
}

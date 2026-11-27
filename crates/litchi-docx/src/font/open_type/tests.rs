use super::{Ligatures, Limits, NumForm, NumSpacing, OnOff, OpenType};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn sample() -> Vec<u8> {
    format!(
        r#"<w:r xmlns:w="{WORD}" xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14">
          <w:rPr>
            <w14:ligatures w14:val="standardContextual"/>
            <x:future xmlns:x="urn:future"><x:data>retain me</x:data></x:future>
            <w14:numForm w14:val="oldStyle"/>
            <w14:numSpacing w14:val="tabular"/>
            <w14:stylisticSets>
              <w14:styleSet w14:id="1"/>
              <w14:styleSet w14:id="4" w14:val="0"/>
            </w14:stylisticSets>
            <w14:cntxtAlts/>
          </w:rPr>
          <w:t>Office</w:t>
        </w:r>"#
    )
    .into_bytes()
}

#[test]
fn parses_typed_open_type_family_and_authored_defaults() {
    let value = OpenType::parse(&sample(), &Limits::DEFAULT).expect("OpenType sample should parse");
    assert_eq!(value.ligatures, Some(Ligatures::StandardContextual));
    assert_eq!(value.num_form, Some(NumForm::OldStyle));
    assert_eq!(value.num_spacing, Some(NumSpacing::Tabular));
    assert!(value.stylistic_sets_present());
    assert_eq!(value.stylistic_sets.len(), 2);
    assert_eq!(value.stylistic_sets[0].id.get(), 1);
    assert_eq!(value.stylistic_sets[0].enabled, None);
    assert_eq!(value.stylistic_sets[1].enabled, Some(false));
    assert_eq!(value.cntxt_alts, Some(OnOff::default_on()));
    assert_eq!(value.cntxt_alts.map(OnOff::effective), Some(true));
}

#[test]
fn absent_and_explicit_on_off_forms_remain_distinct() {
    let absent = OpenType::parse(
        format!(r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"/>"#).as_bytes(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(absent.cntxt_alts, None);
    assert!(!absent.stylistic_sets_present());

    let authored = OpenType::parse(
        format!(
            r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"><w14:cntxtAlts w14:val="1"/><w14:stylisticSets/></w:rPr>"#
        )
        .as_bytes(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(authored.cntxt_alts, Some(OnOff::on()));
    assert_eq!(authored.cntxt_alts.unwrap().authored(), Some(true));
    assert!(authored.stylistic_sets_present());
    assert!(authored.stylistic_sets.is_empty());

    let explicit_false = OpenType::parse(
        format!(
            r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"><w14:cntxtAlts w14:val="false"/></w:rPr>"#
        )
        .as_bytes(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(explicit_false.cntxt_alts, Some(OnOff::off()));
}

#[test]
fn run_facade_reads_and_edits_open_type_features() {
    let mut run = crate::paragraph::Run::new(sample());
    assert_eq!(run.open_type().unwrap().num_form, Some(NumForm::OldStyle));
    let mut next = run.open_type().unwrap();
    next.set_stylistic_sets(None).unwrap();
    next.cntxt_alts = Some(OnOff::off());
    run.set_open_type(next).unwrap();
    let value = run.open_type().unwrap();
    assert!(!value.stylistic_sets_present());
    assert_eq!(value.cntxt_alts, Some(OnOff::off()));
    assert!(run.text().unwrap().contains("Office"));
}

/// Outcome of `codec::parse` for one input, as frozen text.
fn open_type_run_envelope_outcome(xml: &str) -> String {
    match super::codec::parse(xml.as_bytes()) {
        Ok(_) => "Ok".to_owned(),
        Err(error) => format!("{error:?}"),
    }
}

#[test]
fn open_type_run_xml_envelope_matrix() {
    let rows: [(&str, &str, &str); 21] = [
        (
            "empty input",
            "",
            r#"InvalidFormat("Word OpenType XML is not one complete element")"#,
        ),
        (
            "whitespace only",
            " \n",
            r#"InvalidFormat("Word OpenType XML is not one complete element")"#,
        ),
        (
            "declaration only",
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "empty-element root",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
            "Ok",
        ),
        (
            "start and end root",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            "Ok",
        ),
        (
            "root then whitespace",
            "<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:r>\n ",
            "Ok",
        ),
        (
            "root then comment",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><!--c-->"#,
            "Ok",
        ),
        (
            "root then processing instruction",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><?pi x?>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "root then text",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>text"#,
            r#"InvalidFormat("Word OpenType XML has text outside its root")"#,
        ),
        (
            "root then a second identical root",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML must have one w:r or w:rPr root")"#,
        ),
        (
            "root then a foreign start element",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><foreign></foreign>"#,
            r#"InvalidFormat("Word OpenType XML must have one w:r or w:rPr root")"#,
        ),
        (
            "root then a foreign empty element",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><foreign/>"#,
            r#"InvalidFormat("Word OpenType XML must have one w:r or w:rPr root")"#,
        ),
        (
            "declaration after the root",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r><?xml version="1.0" encoding="UTF-8"?>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "two declarations",
            r#"<?xml version="1.0" encoding="UTF-8"?><?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "comment before the declaration",
            r#"<!--c--><?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "text before the root",
            r#"text<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML has text outside its root")"#,
        ),
        (
            "unterminated root",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"InvalidFormat("Word OpenType XML is not one complete element")"#,
        ),
        (
            "stray end tag",
            r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r></w:r>"#,
            r#"Xml("ill-formed document: close tag `</w:r>` does not match any open tag")"#,
        ),
        (
            "document type declaration",
            r#"<!DOCTYPE x><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
        (
            "top-level character data section",
            r#"<![CDATA[x]]><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML has unexpected character data")"#,
        ),
        (
            "declaration then root",
            r#"<?xml version="1.0" encoding="UTF-8"?><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>"#,
            r#"InvalidFormat("Word OpenType XML cannot contain declarations or processing instructions")"#,
        ),
    ];
    for (name, xml, expected) in rows {
        assert_eq!(open_type_run_envelope_outcome(xml), expected, "{name}");
    }
}

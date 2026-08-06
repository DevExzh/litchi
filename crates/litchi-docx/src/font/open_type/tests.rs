use super::{Ligatures, NumForm, NumSpacing, OnOff, OpenType, Snapshot, StyleSet, StyleSetId};

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
    let value = OpenType::parse(&sample()).expect("OpenType sample should parse");
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
fn snapshot_noop_and_edit_preserve_unknown_xml() {
    let source = sample();
    let snapshot = Snapshot::from_xml(source.clone()).expect("snapshot");
    let noop = snapshot.edit().commit().expect("no-op commit");
    assert_eq!(noop.snapshot().xml_bytes(), source.as_slice());

    let mut edit = snapshot.edit();
    edit.set_ligatures(Some(Ligatures::All));
    let changed = edit.commit().expect("OpenType edit");
    let output = changed.snapshot().xml_bytes();
    assert!(
        output
            .windows(
                b"<x:future xmlns:x=\"urn:future\"><x:data>retain me</x:data></x:future>".len()
            )
            .any(|window| {
                window == b"<x:future xmlns:x=\"urn:future\"><x:data>retain me</x:data></x:future>"
            })
    );
    assert_eq!(
        OpenType::parse(output).unwrap().ligatures,
        Some(Ligatures::All)
    );

    let restored = changed
        .patch()
        .inverse()
        .apply(changed.snapshot())
        .expect("inverse patch");
    assert_eq!(
        OpenType::parse(restored.xml_bytes()).unwrap(),
        OpenType::parse(&source).unwrap()
    );
    assert!(String::from_utf8_lossy(restored.xml_bytes()).contains("retain me"));
}

#[test]
fn absent_and_explicit_on_off_forms_remain_distinct() {
    let absent =
        OpenType::parse(format!(r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"/>"#).as_bytes())
            .unwrap();
    assert_eq!(absent.cntxt_alts, None);
    assert!(!absent.stylistic_sets_present());

    let authored = OpenType::parse(
        format!(
            r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"><w14:cntxtAlts w14:val="1"/><w14:stylisticSets/></w:rPr>"#
        )
        .as_bytes(),
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
    )
    .unwrap();
    assert_eq!(explicit_false.cntxt_alts, Some(OnOff::off()));
}

#[test]
fn transactions_validate_domains_and_handle_empty_roots() {
    let id = StyleSetId::new(3).unwrap();
    let mut value = OpenType::default();
    value
        .set_style_set(StyleSet::new(id).with_enabled(Some(true)))
        .unwrap();
    assert!(value.stylistic_sets_present());

    let xml = format!(r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"/>"#);
    let snapshot = Snapshot::from_xml(xml.as_bytes().to_vec()).unwrap();
    let mut edit = snapshot.edit();
    edit.set_num_form(Some(NumForm::Lining));
    edit.set_cntxt_alts(Some(OnOff::off()));
    edit.set_style_set(StyleSet::new(id)).unwrap();
    let committed = edit.commit().unwrap();
    assert_eq!(
        OpenType::parse(committed.snapshot().xml_bytes())
            .unwrap()
            .num_form,
        Some(NumForm::Lining)
    );
    let output = String::from_utf8_lossy(committed.snapshot().xml_bytes());
    assert!(output.contains("mc:Ignorable=\"w14\""));
    assert!(output.contains("w14:styleSet w14:id=\"3\""));

    let invalid = format!(
        r#"<w:rPr xmlns:w="{WORD}" xmlns:w14="{W14}"><w14:stylisticSets><w14:styleSet w14:id="21"/></w14:stylisticSets></w:rPr>"#
    );
    assert!(OpenType::parse(invalid.as_bytes()).is_err());
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

#[cfg(test)]
mod chpx_position_hresi_effect_writer_tests {
    pub(super) use super::super::super::super::support::*;
    use crate::parts::chp::{CharacterPosition, HresiOperand, HyphenationMode, TextEffect};
    use std::io::Cursor;

    #[test]
    fn emits_canonical_typed_sprms_and_round_trips_package() {
        let position = CharacterPosition::new(-3168).unwrap();
        let hyphenation =
            HresiOperand::with_character(HyphenationMode::DeleteAndChange, b'Z').unwrap();
        let formatting = CharacterFormatting {
            position: Some(position),
            hyphenation: Some(hyphenation),
            text_effect: Some(TextEffect::Shimmer),
            ..CharacterFormatting::default()
        };
        let mut fonts = FontTableBuilder::new();
        let grpprl = build_chpx_grpprl(&formatting, &mut fonts);
        let mut expected = Vec::new();
        expected.extend_from_slice(&SPRM_C_HPS_POS.to_le_bytes());
        expected.extend_from_slice(&(-3168i16).to_le_bytes());
        expected.extend_from_slice(&SPRM_C_HRESI.to_le_bytes());
        expected.extend_from_slice(&[6, b'Z']);
        expected.extend_from_slice(&SPRM_C_SFXT_TEXT.to_le_bytes());
        expected.push(6);
        assert_eq!(grpprl, expected);

        let properties = crate::parts::chp::CharacterProperties::from_sprm(&grpprl).unwrap();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);

        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![("effects".to_string(), formatting)],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let mut package = crate::Package::from_reader(Cursor::new(output.into_inner())).unwrap();
        let document = package.document().unwrap();
        let paragraphs = document.paragraphs().unwrap();
        let runs = paragraphs[0].runs().unwrap();
        let properties = runs[0].properties();
        assert_eq!(properties.position, position);
        assert_eq!(properties.hyphenation, hyphenation);
        assert_eq!(properties.text_effect, TextEffect::Shimmer);
    }
}

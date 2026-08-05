use super::codec::*;
use super::model::{TextCFException, TextPFException};
use crate::consts::RecordType;
use crate::records::record::Record;
use crate::slide_show_settings::ColorIndexKind;
use crate::text_run::{
    ParagraphAlignment, ParagraphFontAlignment, ParagraphTabAlignment, ParagraphTextDirection,
};

fn cf_record(data: &[u8]) -> Record {
    Record {
        record_type: RecordType::TxCFStyleAtom,
        record_type_raw: TEXT_CF_EXCEPTION_TYPE,
        version: 0,
        instance: 0,
        data_length: data.len() as u32,
        data: data.to_vec(),
        children: Vec::new(),
    }
}

fn pf_record(data: &[u8]) -> Record {
    Record {
        record_type: RecordType::TxPFStyleAtom,
        record_type_raw: TEXT_PF_EXCEPTION_TYPE,
        version: 0,
        instance: 0,
        data_length: data.len() as u32,
        data: data.to_vec(),
        children: Vec::new(),
    }
}

/// Mask: bold | italic | emboss | fHasStyle(3) | typeface | ansiTypeface
/// | size | color | position.
fn sample_cf_payload() -> Vec<u8> {
    let masks = CF_MASK_BOLD
        | CF_MASK_ITALIC
        | CF_MASK_EMBOSS
        | (3 << 10)
        | CF_MASK_TYPEFACE
        | CF_MASK_ANSI_TYPEFACE
        | CF_MASK_SIZE
        | CF_MASK_COLOR
        | CF_MASK_POSITION;
    let mut data = Vec::new();
    data.extend_from_slice(&masks.to_le_bytes());
    // fontStyle: bold + italic set, pp9rt = 3.
    data.extend_from_slice(&0x0C03u16.to_le_bytes());
    data.extend_from_slice(&7u16.to_le_bytes()); // fontRef
    data.extend_from_slice(&11u16.to_le_bytes()); // ansiFontRef
    data.extend_from_slice(&2400i16.to_le_bytes()); // fontSize
    data.extend_from_slice(&[0x12, 0x34, 0x56, 0xFE]); // sRGB color
    data.extend_from_slice(&(-30i16).to_le_bytes()); // position
    data
}

#[test]
fn parses_cf_exception_and_round_trips() {
    let payload = sample_cf_payload();
    let parsed = TextCFException::parse_record(&cf_record(&payload)).unwrap();
    assert_eq!(
        parsed.masks(),
        CF_MASK_BOLD
            | CF_MASK_ITALIC
            | CF_MASK_EMBOSS
            | (3 << 10)
            | CF_MASK_TYPEFACE
            | CF_MASK_ANSI_TYPEFACE
            | CF_MASK_SIZE
            | CF_MASK_COLOR
            | CF_MASK_POSITION
    );
    let style = parsed.font_style().unwrap();
    assert!(style.bold());
    assert!(style.italic());
    assert!(!style.underline());
    assert!(!style.emboss());
    assert_eq!(style.pp9_run_group(), 3);
    assert_eq!(parsed.font_ref(), Some(7));
    assert_eq!(parsed.old_east_asian_font_ref(), None);
    assert_eq!(parsed.ansi_font_ref(), Some(11));
    assert_eq!(parsed.symbol_font_ref(), None);
    assert_eq!(parsed.font_size(), Some(2400));
    let color = parsed.color().unwrap();
    assert_eq!(color.red, 0x12);
    assert_eq!(color.green, 0x34);
    assert_eq!(color.blue, 0x56);
    assert_eq!(color.kind, ColorIndexKind::Srgb);
    assert_eq!(parsed.position(), Some(-30));

    assert_eq!(parsed.to_bytes()[8..], payload[..]);
}

#[test]
fn parses_empty_cf_exception_and_round_trips() {
    let payload = 0u32.to_le_bytes().to_vec();
    let parsed = TextCFException::parse_record(&cf_record(&payload)).unwrap();
    assert_eq!(parsed.masks(), 0);
    assert_eq!(parsed.font_style(), None);
    assert_eq!(parsed.to_bytes()[8..], payload[..]);
}

#[test]
fn cf_style_presence_follows_fhas_style_only() {
    // Only fHasStyle = 2 set: fontStyle exists, pp9rt read from CFStyle.
    let mut payload = Vec::new();
    payload.extend_from_slice(&(2u32 << 10).to_le_bytes());
    payload.extend_from_slice(&0x1400u16.to_le_bytes()); // pp9rt = 5
    let parsed = TextCFException::parse_record(&cf_record(&payload)).unwrap();
    let style = parsed.font_style().unwrap();
    assert!(!style.bold());
    assert_eq!(style.pp9_run_group(), 5);
    assert_eq!(parsed.to_bytes()[8..], payload[..]);
}

#[test]
fn rejects_malformed_cf_exception() {
    // Wrong record type.
    assert!(TextCFException::parse_record(&pf_record(&[])).is_err());
    // Truncated masks.
    assert!(TextCFException::parse_record(&cf_record(&[1, 0])).is_err());
    // Forbidden pp10ext mask bit.
    assert!(TextCFException::parse_record(&cf_record(&0x0010_0000u32.to_le_bytes())).is_err());
    // Forbidden reserved mask bits.
    assert!(TextCFException::parse_record(&cf_record(&0x8000_0000u32.to_le_bytes())).is_err());
    // Truncated fontStyle.
    assert!(TextCFException::parse_record(&cf_record(&CF_MASK_BOLD.to_le_bytes())).is_err());
    // fontSize below the minimum.
    let mut payload = CF_MASK_SIZE.to_le_bytes().to_vec();
    payload.extend_from_slice(&0i16.to_le_bytes());
    assert!(TextCFException::parse_record(&cf_record(&payload)).is_err());
    // fontSize above the maximum.
    let mut payload = CF_MASK_SIZE.to_le_bytes().to_vec();
    payload.extend_from_slice(&4001i16.to_le_bytes());
    assert!(TextCFException::parse_record(&cf_record(&payload)).is_err());
    // position above the maximum.
    let mut payload = CF_MASK_POSITION.to_le_bytes().to_vec();
    payload.extend_from_slice(&101i16.to_le_bytes());
    assert!(TextCFException::parse_record(&cf_record(&payload)).is_err());
    // Invalid color index byte.
    let mut payload = CF_MASK_COLOR.to_le_bytes().to_vec();
    payload.extend_from_slice(&[0, 0, 0, 0x08]);
    assert!(TextCFException::parse_record(&cf_record(&payload)).is_err());
    // Trailing bytes after a complete structure.
    let mut payload = 0u32.to_le_bytes().to_vec();
    payload.push(0);
    assert!(TextCFException::parse_record(&cf_record(&payload)).is_err());
    // Nonzero instance.
    let mut record = cf_record(&0u32.to_le_bytes());
    record.instance = 1;
    assert!(TextCFException::parse_record(&record).is_err());
}

/// Mask: hasBullet | bulletHasFont | bulletChar | bulletFont | bulletSize
/// | align | lineSpacing | leftMargin | defaultTabSize | tabStops
/// | fontAlign | charWrap | wordWrap | textDirection.
fn sample_pf_payload() -> Vec<u8> {
    let masks = PF_MASK_HAS_BULLET
        | PF_MASK_BULLET_HAS_FONT
        | PF_MASK_BULLET_CHAR
        | PF_MASK_BULLET_FONT
        | PF_MASK_BULLET_SIZE
        | PF_MASK_ALIGN
        | PF_MASK_LINE_SPACING
        | PF_MASK_LEFT_MARGIN
        | PF_MASK_DEFAULT_TAB_SIZE
        | PF_MASK_TAB_STOPS
        | PF_MASK_FONT_ALIGN
        | PF_MASK_CHAR_WRAP
        | PF_MASK_WORD_WRAP
        | PF_MASK_TEXT_DIRECTION;
    let mut data = Vec::new();
    data.extend_from_slice(&0u16.to_le_bytes()); // reserved
    data.extend_from_slice(&masks.to_le_bytes());
    data.extend_from_slice(&0x0003u16.to_le_bytes()); // bulletFlags
    data.extend_from_slice(&0x2022u16.to_le_bytes()); // bulletChar U+2022
    data.extend_from_slice(&2u16.to_le_bytes()); // bulletFontRef
    data.extend_from_slice(&(-1200i16).to_le_bytes()); // bulletSize, points
    data.extend_from_slice(&1u16.to_le_bytes()); // align center
    data.extend_from_slice(&150i16.to_le_bytes()); // lineSpacing, percent
    data.extend_from_slice(&288i16.to_le_bytes()); // leftMargin
    data.extend_from_slice(&720i16.to_le_bytes()); // defaultTabSize
    data.extend_from_slice(&2u16.to_le_bytes()); // two tab stops
    data.extend_from_slice(&100i16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes()); // left tab
    data.extend_from_slice(&(-40i16).to_le_bytes());
    data.extend_from_slice(&3u16.to_le_bytes()); // decimal tab
    data.extend_from_slice(&3u16.to_le_bytes()); // fontAlign upholdFixed
    data.extend_from_slice(&0x0003u16.to_le_bytes()); // wrapFlags
    data.extend_from_slice(&1u16.to_le_bytes()); // right-to-left
    data
}

#[test]
fn parses_pf_exception_and_round_trips() {
    let payload = sample_pf_payload();
    let parsed = TextPFException::parse_record(&pf_record(&payload)).unwrap();

    let flags = parsed.bullet_flags().unwrap();
    assert!(flags.has_bullet());
    assert!(flags.bullet_has_font());
    assert!(!flags.bullet_has_color());
    assert!(!flags.bullet_has_size());
    assert_eq!(parsed.bullet_char(), Some(0x2022));
    assert_eq!(parsed.bullet_font_ref(), Some(2));
    assert_eq!(parsed.bullet_size(), Some(-1200));
    assert_eq!(parsed.bullet_color(), None);
    assert_eq!(parsed.text_alignment(), Some(ParagraphAlignment::Center));
    assert_eq!(parsed.line_spacing(), Some(150));
    assert_eq!(parsed.space_before(), None);
    assert_eq!(parsed.left_margin(), Some(288));
    assert_eq!(parsed.indent(), None);
    assert_eq!(parsed.default_tab_size(), Some(720));
    let stops = parsed.tab_stops().unwrap();
    assert_eq!(stops.len(), 2);
    assert_eq!(stops[0].position, 100);
    assert_eq!(stops[0].alignment, ParagraphTabAlignment::Left);
    assert_eq!(stops[1].position, -40);
    assert_eq!(stops[1].alignment, ParagraphTabAlignment::Decimal);
    assert_eq!(
        parsed.font_align(),
        Some(ParagraphFontAlignment::UpholdFixed)
    );
    let wrap = parsed.wrap_flags().unwrap();
    assert!(wrap.char_wrap());
    assert!(wrap.word_wrap());
    assert!(!wrap.overflow());
    assert_eq!(
        parsed.text_direction(),
        Some(ParagraphTextDirection::RightToLeft)
    );

    assert_eq!(parsed.to_bytes()[8..], payload[..]);
}

#[test]
fn parses_empty_pf_exception_and_round_trips() {
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&0u32.to_le_bytes());
    let parsed = TextPFException::parse_record(&pf_record(&payload)).unwrap();
    assert_eq!(parsed.masks(), 0);
    assert_eq!(parsed.bullet_flags(), None);
    assert_eq!(parsed.tab_stops(), None);
    assert_eq!(parsed.to_bytes()[8..], payload[..]);
}

#[test]
fn rejects_malformed_pf_exception() {
    // Wrong record type.
    assert!(TextPFException::parse_record(&cf_record(&[])).is_err());
    // Truncated reserved field.
    assert!(TextPFException::parse_record(&pf_record(&[0])).is_err());
    // Nonzero reserved field.
    assert!(TextPFException::parse_record(&pf_record(&[1, 0])).is_err());
    // Forbidden bulletBlip mask bit.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&0x0080_0000u32.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // BulletFlags reserved bits.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_HAS_BULLET.to_le_bytes());
    payload.extend_from_slice(&0x0010u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // NUL bullet character.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_BULLET_CHAR.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Invalid alignment value.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_ALIGN.to_le_bytes());
    payload.extend_from_slice(&7u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // lineSpacing above the maximum percentage.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_LINE_SPACING.to_le_bytes());
    payload.extend_from_slice(&13201i16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Truncated tab stops.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_TAB_STOPS.to_le_bytes());
    payload.extend_from_slice(&2u16.to_le_bytes());
    payload.extend_from_slice(&[0; 4]);
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Invalid tab-stop type.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_TAB_STOPS.to_le_bytes());
    payload.extend_from_slice(&1u16.to_le_bytes());
    payload.extend_from_slice(&0i16.to_le_bytes());
    payload.extend_from_slice(&4u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // PFWrapFlags reserved bits.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_CHAR_WRAP.to_le_bytes());
    payload.extend_from_slice(&0x0008u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Invalid text direction.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&PF_MASK_TEXT_DIRECTION.to_le_bytes());
    payload.extend_from_slice(&2u16.to_le_bytes());
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Trailing bytes after a complete structure.
    let mut payload = 0u16.to_le_bytes().to_vec();
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.push(0);
    assert!(TextPFException::parse_record(&pf_record(&payload)).is_err());
    // Nonzero version.
    let mut record = pf_record(&[0; 6]);
    record.version = 0xF;
    assert!(TextPFException::parse_record(&record).is_err());
}

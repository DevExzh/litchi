use super::codec::{AUTO_CAPTION_FIB_INDEX, CAPI_SIZE, CAPTION_FIB_INDEX};
use super::{
    AutoEntry, AutoTable, Definition, Format, Heading, Info, LabelTable, Location, Numbering,
    Separator, Tables,
};
use crate::parts::fib::FileInformationBlock;

fn info() -> Info {
    Info::new(Location::Below, None, false, Format::Arabic)
}

fn chapter_info() -> Info {
    Info::new(
        Location::Above,
        Some(Numbering::new(Heading::Level2, Separator::EnDash)),
        true,
        Format::Arabic,
    )
}

fn definition(label: &str, value: Info) -> Definition {
    Definition::try_new(label.to_owned(), value).expect("caption definition is valid")
}

fn fib_with_pointers(pairs: &[(usize, usize, usize)]) -> FileInformationBlock {
    let mut data = vec![0u8; 154 + 117 * 8];
    data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    data[10..12].copy_from_slice(&1u16.to_le_bytes());
    data[152..154].copy_from_slice(&117u16.to_le_bytes());
    data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes());
    for &(index, offset, length) in pairs {
        let pointer = 154 + index * 8;
        data[pointer..pointer + 4].copy_from_slice(&(offset as u32).to_le_bytes());
        data[pointer + 4..pointer + 8].copy_from_slice(&(length as u32).to_le_bytes());
    }
    FileInformationBlock::parse(&data).expect("fixture FIB parses")
}

#[test]
fn capi_round_trips_and_canonicalizes_ignored_bits() {
    let value = chapter_info();
    let bytes = value.to_bytes();
    assert_eq!(bytes.len(), Info::SIZE);
    assert_eq!(u16::from_le_bytes([bytes[0], bytes[1]]), 0x8015);
    assert_eq!(u16::from_le_bytes([bytes[2], bytes[3]]), 0);
    assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x2013);
    assert_eq!(Info::from_bytes(&bytes).expect("CAPI parses"), value);

    let mut noncanonical = info().to_bytes();
    noncanonical[0] |= 0x78;
    noncanonical[1] = 0x7F;
    noncanonical[4..6].copy_from_slice(&0xFFFFu16.to_le_bytes());
    let parsed = Info::from_bytes(&noncanonical).expect("ignored CAPI fields parse");
    assert_eq!(parsed, info());
    assert_eq!(parsed.to_bytes(), info().to_bytes());
}

#[test]
fn capi_rejects_invalid_domains() {
    let mut bytes = info().to_bytes();
    bytes[0] = 0x2;
    assert!(Info::from_bytes(&bytes).is_err());

    bytes = [0x4, 0, 0, 0, 0x2E, 0];
    assert!(Info::from_bytes(&bytes).is_err());
    bytes = [0x4 | 0xA << 3, 0, 0, 0, 0x2E, 0];
    assert!(Info::from_bytes(&bytes).is_err());
    bytes = [0x4 | 0x1 << 3, 0, 0, 0, 0x3B, 0];
    assert!(Info::from_bytes(&bytes).is_err());
    bytes = [0x4 | 0x1 << 3, 0, 0, 0x40, 0x2E, 0];
    assert!(Info::from_bytes(&bytes).is_err());
    bytes = [0x4 | 0x1 << 3, 0, 0x40, 0, 0x2E, 0];
    assert!(Info::from_bytes(&bytes).is_err());
    assert!(Info::from_bytes(&bytes[..CAPI_SIZE - 1]).is_err());
}

#[test]
fn label_table_round_trips_unicode_and_chapter_options() {
    let value = LabelTable::try_new(vec![
        definition("Equation", info()),
        definition("Figure", chapter_info()),
        definition(
            "Table",
            Info::new(Location::Below, None, true, Format::UpperRoman),
        ),
    ])
    .expect("label table is valid");
    let bytes = value.to_bytes().expect("label table serializes");
    let parsed = LabelTable::parse_bytes(&bytes).expect("label table parses");
    assert_eq!(parsed, value);
    assert_eq!(parsed.definitions()[1].label(), "Figure");
    assert_eq!(parsed.definitions()[2].info().omit_label(), true);
    assert_eq!(parsed.to_bytes().expect("parsed table serializes"), bytes);

    let unicode =
        LabelTable::try_new(vec![definition("表🙂", info())]).expect("unicode label is valid");
    assert_eq!(
        LabelTable::parse_bytes(&unicode.to_bytes().unwrap()).unwrap(),
        unicode
    );
}

#[test]
fn label_table_rejects_malformed_header_payload_and_label() {
    assert!(LabelTable::parse_bytes(&[]).is_err());

    let valid = LabelTable::try_new(vec![definition("Figure", info())])
        .unwrap()
        .to_bytes()
        .unwrap();
    let mut bytes = valid.clone();
    bytes[0] = 0;
    assert!(LabelTable::parse_bytes(&bytes).is_err());

    let mut bytes = valid.clone();
    bytes[4] = 2;
    assert!(LabelTable::parse_bytes(&bytes).is_err());

    let mut bytes = valid.clone();
    bytes[2] = 2;
    assert!(LabelTable::parse_bytes(&bytes).is_err());

    let mut bytes = valid.clone();
    bytes.truncate(bytes.len() - 2);
    assert!(LabelTable::parse_bytes(&bytes).is_err());

    let mut bytes = valid;
    bytes.push(0);
    assert!(LabelTable::parse_bytes(&bytes).is_err());

    let mut long = vec![0xFF, 0xFF, 1, 0, CAPI_SIZE as u8, 0, 41, 0];
    long.extend(std::iter::repeat(b'a').take(82));
    long.extend_from_slice(&info().to_bytes());
    assert!(LabelTable::parse_bytes(&long).is_err());

    let long = "a".repeat(41);
    assert!(Definition::try_new(long, info()).is_err());
}

#[test]
fn auto_table_round_trips_and_rejects_invalid_strings() {
    let value = AutoTable::try_new(vec![
        AutoEntry::try_new("Excel.Sheet.8".to_owned(), 1).unwrap(),
        AutoEntry::try_new("Equation.3".to_owned(), 0).unwrap(),
    ])
    .expect("AutoCaption table is valid");
    let bytes = value.to_bytes().expect("AutoCaption serializes");
    let parsed = AutoTable::parse_bytes(&bytes).expect("AutoCaption parses");
    assert_eq!(parsed, value);
    assert_eq!(parsed.entries()[0].prog_id(), "Excel.Sheet.8");
    assert_eq!(parsed.to_bytes().expect("parsed table serializes"), bytes);

    assert!(AutoTable::parse_bytes(&[]).is_err());
    let mut wrong_extra = bytes.clone();
    wrong_extra[4] = 6;
    assert!(AutoTable::parse_bytes(&wrong_extra).is_err());
    let mut truncated = bytes.clone();
    truncated.truncate(truncated.len() - 1);
    assert!(AutoTable::parse_bytes(&truncated).is_err());
    let mut trailing = bytes;
    trailing.extend_from_slice(&[0, 0]);
    assert!(AutoTable::parse_bytes(&trailing).is_err());

    let long = "a".repeat(usize::from(u16::MAX) + 1);
    assert!(AutoEntry::try_new(long, 0).is_err());
}

#[test]
fn semantic_fib_facade_validates_cross_table_indexes() {
    let labels = LabelTable::try_new(vec![
        definition("Equation", info()),
        definition("Figure", chapter_info()),
    ])
    .unwrap()
    .to_bytes()
    .unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Equation.3".to_owned(), 1).unwrap(),
    ])
    .unwrap()
    .to_bytes()
    .unwrap();
    let mut table_stream = vec![0u8; 11];
    table_stream.extend_from_slice(&labels);
    table_stream.extend_from_slice(&auto);
    let fib = fib_with_pointers(&[
        (CAPTION_FIB_INDEX, 11, labels.len()),
        (AUTO_CAPTION_FIB_INDEX, 11 + labels.len(), auto.len()),
    ]);
    let tables = Tables::parse(&fib, &table_stream).expect("caption metadata parses");
    assert_eq!(tables.labels().unwrap().len(), 2);
    assert_eq!(tables.auto().unwrap().entries()[0].caption_index(), 1);
    assert_eq!(tables.auto_captions(), tables.auto());
}

#[test]
fn semantic_facade_rejects_dangling_or_out_of_range_fib_tables() {
    let labels = LabelTable::try_new(vec![definition("Equation", info())])
        .unwrap()
        .to_bytes()
        .unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Equation.3".to_owned(), 1).unwrap(),
    ])
    .unwrap()
    .to_bytes()
    .unwrap();
    let mut table_stream = labels.clone();
    table_stream.extend_from_slice(&auto);
    let fib = fib_with_pointers(&[
        (CAPTION_FIB_INDEX, 0, labels.len()),
        (AUTO_CAPTION_FIB_INDEX, labels.len(), auto.len()),
    ]);
    assert!(Tables::parse(&fib, &table_stream).is_err());

    let fib = fib_with_pointers(&[(CAPTION_FIB_INDEX, 0, labels.len() + 1)]);
    assert!(Tables::parse(&fib, &labels).is_err());

    let auto_only = fib_with_pointers(&[(AUTO_CAPTION_FIB_INDEX, 0, auto.len())]);
    assert!(Tables::parse(&auto_only, &auto).is_err());

    let empty_auto = AutoTable::default().to_bytes().unwrap();
    let fib = fib_with_pointers(&[(AUTO_CAPTION_FIB_INDEX, 0, empty_auto.len())]);
    let parsed = Tables::parse(&fib, &empty_auto).unwrap();
    assert!(parsed.labels().is_none());
    assert!(parsed.auto().unwrap().is_empty());
}

#[test]
fn ignores_caption_pointers_in_non_template_documents() {
    let labels = LabelTable::try_new(vec![definition("Figure", info())])
        .unwrap()
        .to_bytes()
        .unwrap();
    let template_fib = fib_with_pointers(&[(CAPTION_FIB_INDEX, 0, labels.len())]);
    let mut ordinary_raw = template_fib.raw_data().to_vec();
    ordinary_raw[10..12].copy_from_slice(&0u16.to_le_bytes());
    let ordinary_fib = FileInformationBlock::parse(&ordinary_raw).unwrap();
    assert_eq!(
        Tables::parse(&ordinary_fib, &[]).unwrap(),
        Tables::default()
    );
}

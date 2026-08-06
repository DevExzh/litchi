use super::*;
use crate::parts::fib::FileInformationBlock;
use crate::writer::fib::FibBuilder;
use litchi_ole_common::toolbar::{Flags, Header, Restrictions, Type, WString};

fn sample() -> CommandBars<'static> {
    let header = Header::new(
        0,
        Restrictions::new(Type::Basic).expect("basic restrictions"),
        0,
        Flags::default(),
        WString::new("Toolbar").expect("toolbar name"),
    )
    .expect("toolbar header");
    let toolbar = Toolbar::new(
        XString::new("Custom").expect("custom toolbar name"),
        header,
        [0xA5; 100],
        0,
        0,
        0xCAFE,
    )
    .expect("custom toolbar");
    let custom_toolbar = Customization::toolbar(toolbar).expect("toolbar customization");
    let delta = ToolbarDelta::new(
        Operation::Insert,
        true,
        0,
        2,
        0x1111_1111,
        0x2222_2222,
        0x3333_3333,
        0x0001,
        4,
    );
    let custom_delta = Customization::deltas(7, vec![delta]).expect("delta customization");
    let wrapper = ToolbarWrapper::new(vec![custom_toolbar, custom_delta]).expect("wrapper");

    CommandBars::new(vec![
        Entry::MacroCommands(
            MacroCommands::new(vec![MacroCommand::new(
                4,
                5,
                0x56,
                0,
                0xFFFF,
                0x1111_1111,
                0,
                0x2222_2222,
                0x3333_3333,
            )])
            .expect("macro commands"),
        ),
        Entry::AllocatedCommands(
            AllocatedCommands::new(vec![AllocatedCommand::new(3, 0x123, 0x05)])
                .expect("allocated commands"),
        ),
        Entry::KeyMaps(
            KeyMaps::new(
                KeyMapKind::Regular,
                vec![KeyMap::new(0x10, 0x00FF, Action::Command, 0x44, 0, 0)],
            )
            .expect("key maps"),
        ),
        Entry::Toolbar(wrapper),
    ])
    .expect("command bars")
}

#[test]
fn round_trip_preserves_typed_records_and_opaque_bytes() {
    let value = sample();
    let bytes = to_bytes(&value).expect("serialize");
    let parsed = parse_bytes(&bytes).expect("parse");
    assert_eq!(parsed, value);
    assert_eq!(to_bytes(&parsed).expect("serialize parsed"), bytes);
    assert_eq!(parsed.entries().len(), 4);
    assert_eq!(parsed.version(), 0xFF);
    assert_eq!(parsed.terminator(), 0x40);
}

#[test]
fn parses_optional_fib_pointer_and_writes_a_new_table_range() {
    let value = sample();
    let payload = to_bytes(&value).expect("serialize");
    let offset = 9usize;
    let mut table_stream = vec![0xAA; offset];
    table_stream.extend_from_slice(&payload);
    let fib = fib_with_pointer(offset, payload.len());
    assert_eq!(
        parse(&fib, &table_stream).expect("parse pointer"),
        Some(value)
    );

    let mut builder = FibBuilder::new();
    let mut generated_table = vec![0xCC; 4];
    write(&mut builder, &mut generated_table, &sample()).expect("append command bars");
    let generated_fib = FileInformationBlock::parse(&builder.generate().expect("generate FIB"))
        .expect("parse generated FIB");
    assert_eq!(
        generated_fib.get_table_pointer(FIB_INDEX_CMDS),
        Some((4, payload.len() as u32))
    );
    assert_eq!(&generated_table[4..], payload.as_slice());
}

#[test]
fn absent_or_empty_fib_pointer_is_optional() {
    let fib = fib_with_pointer(0, 0);
    assert_eq!(parse(&fib, &[]).expect("empty pointer"), None);
}

#[test]
fn rejects_bad_counts_reserved_words_and_relationships() {
    let valid = to_bytes(&sample()).expect("serialize");

    let mut missing_terminator = valid.clone();
    missing_terminator.pop();
    assert!(parse_bytes(&missing_terminator).is_err());

    let mut invalid_macro_count = valid.clone();
    invalid_macro_count[2..6].copy_from_slice(&(-1i32).to_le_bytes());
    assert!(parse_bytes(&invalid_macro_count).is_err());

    let mut invalid_macro_reserved = valid.clone();
    invalid_macro_reserved[6] = 0;
    assert!(parse_bytes(&invalid_macro_reserved).is_err());

    let mut invalid_toolbar_index = valid;
    let marker = [0x12, 0x12, 0, 0, 0x07, 0x06, 0, 0x0C, 0, 0x12];
    let wrapper_offset = invalid_toolbar_index
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("wrapper marker");
    let ctb_index = wrapper_offset + 14 + 2 + 4 + 8 + 4 + 2 + 2;
    invalid_toolbar_index[ctb_index..ctb_index + 4].copy_from_slice(&99i32.to_le_bytes());
    assert!(parse_bytes(&invalid_toolbar_index).is_err());
}

#[test]
fn rejects_unknown_and_variable_records_without_guessing_boundaries() {
    assert!(parse_bytes(&[0xFF, 0x10, 0x40]).is_err());
    assert!(parse_bytes(&[0xFF, 0x7F, 0x40]).is_err());
}

#[test]
fn rejects_truncated_variable_strings_and_table_ranges() {
    assert!(XString::parse(&[1, 0, 0x41]).is_err());
    let payload = to_bytes(&sample()).expect("serialize");
    let fib = fib_with_pointer(4, payload.len() + 1);
    assert!(parse(&fib, &[0; 4]).is_err());
}

fn fib_with_pointer(offset: usize, length: usize) -> FileInformationBlock {
    let pointer_offset = 154 + FIB_INDEX_CMDS * 8;
    let mut data = vec![0; pointer_offset + 8];
    data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    data[152..154].copy_from_slice(&((FIB_INDEX_CMDS + 1) as u16).to_le_bytes());
    data[pointer_offset..pointer_offset + 4].copy_from_slice(&(offset as u32).to_le_bytes());
    data[pointer_offset + 4..pointer_offset + 8].copy_from_slice(&(length as u32).to_le_bytes());
    FileInformationBlock::parse(&data).expect("test FIB")
}

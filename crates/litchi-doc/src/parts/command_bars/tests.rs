use super::*;
use crate::parts::fib::FileInformationBlock;
use crate::writer::fib::FibBuilder;
use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::toolbar::{
    ControlFlags, ControlHeader, ControlType, Data, Flags, GeneralFlags, GeneralInfo, Header,
    Restrictions, SpecificFlags, Type, WString,
};

fn control(
    control_type: ControlType,
    control_id: u16,
    command: Option<u32>,
    specific: Vec<u8>,
) -> Control<'static> {
    let header = ControlHeader::new(
        control_type,
        control_id,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("control header");
    let general =
        GeneralInfo::new(GeneralFlags::default(), None, None, None, None).expect("general info");
    let data = if matches!(control_type, ControlType::ActiveX) {
        None
    } else {
        Some(Data::new(general, specific).expect("control data"))
    };
    Control::new(
        header,
        command.map(|raw| CommandId::new(raw).expect("command id")),
        data,
    )
    .expect("control")
}

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
    let wrapper = ToolbarWrapper::new(vec![custom_toolbar, custom_delta])
        .expect("wrapper")
        .with_delta_controls(vec![control(
            ControlType::Label,
            0x1234,
            Some(0x0000_1231),
            Vec::new(),
        )])
        .expect("wrapper controls");

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
        Entry::CommandStrings(
            CommandStrings::new(
                (0..6)
                    .map(|index| {
                        CommandString::new(
                            XString::new(&format!("command-{index}")).expect("command string"),
                            index as u16,
                        )
                        .expect("command string entry")
                    })
                    .collect(),
            )
            .expect("command strings"),
        ),
        Entry::MacroNames(
            MacroNames::new(vec![
                MacroName::new(4, XString::new("AutoMacro").expect("macro name"))
                    .expect("macro name entry"),
            ])
            .expect("macro names"),
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
    assert_eq!(parsed.entries().len(), 6);
    assert_eq!(parsed.version(), 0xFF);
    assert_eq!(parsed.terminator(), 0x40);
    let wrapper = match &parsed.entries()[5] {
        Entry::Toolbar(wrapper) => wrapper,
        _ => panic!("toolbar wrapper entry"),
    };
    assert_eq!(wrapper.delta_controls().len(), 1);
    assert_eq!(
        wrapper.delta_controls()[0].control_type(),
        ControlType::Label
    );
    assert_eq!(
        wrapper.delta_controls()[0].command().unwrap().raw(),
        0x0000_1231
    );
    let strings = match &parsed.entries()[3] {
        Entry::CommandStrings(strings) => strings,
        _ => panic!("command string table"),
    };
    assert_eq!(strings.strings()[3].text().text(), "command-3");
    assert_eq!(strings.strings()[3].references(), 3);
    let names = match &parsed.entries()[4] {
        Entry::MacroNames(names) => names,
        _ => panic!("macro name table"),
    };
    assert_eq!(names.names()[0].index(), 4);
    assert_eq!(names.names()[0].name().text(), "AutoMacro");
}

#[test]
fn command_bar_snapshots_are_owned_and_failure_atomic() {
    let source = sample();
    let snapshot = source.snapshot().expect("owned snapshot");
    let stale = snapshot.edit();
    let mut transaction = snapshot.edit();
    transaction.clear();
    let commit = transaction.commit().expect("semantic commit");

    assert!(!commit.is_noop());
    assert!(!commit.snapshot().is_present());
    assert!(commit.apply(&snapshot).is_ok());
    let mut changed_stale = stale;
    changed_stale.clear();
    assert!(commit.apply(changed_stale.snapshot()).is_err());
    assert_eq!(commit.inverse().apply(commit.snapshot()).unwrap(), snapshot);
    assert_eq!(snapshot.command_bars().unwrap(), &source.into_owned());
}

#[test]
fn package_editor_publishes_command_bars_and_preserves_unrelated_bytes() {
    let original_table = b"opaque table prefix\0";
    let mut editor = Editor::open(base_document(original_table)).expect("DOC package opens");
    assert!(!editor.command_bars().is_present());

    let value = sample();
    let committed = editor.set(value.clone()).expect("command bars publish");
    let bytes = committed
        .snapshot()
        .finish()
        .expect("package snapshot renders");
    let reopened = Editor::open(bytes.clone()).expect("published package reopens");
    assert_eq!(reopened.command_bars().command_bars(), Some(&value));

    let mut ole = OleFile::open(std::io::Cursor::new(bytes.clone())).expect("CFB opens");
    let table = ole.open_stream(&["0Table"]).expect("table stream exists");
    assert!(table.starts_with(original_table));
    assert!(committed.package_patch().inverse().apply(&bytes).is_ok());

    let mut cleared = Editor::open(bytes).expect("published package opens for clear");
    let cleared = cleared.clear().expect("command bars clear");
    let cleared_bytes = cleared.snapshot().finish().expect("clear renders");
    let mut ole = OleFile::open(std::io::Cursor::new(cleared_bytes)).expect("cleared CFB opens");
    let word = ole
        .open_stream(&["WordDocument"])
        .expect("WordDocument stream exists");
    let pointer = 154 + FIB_INDEX_CMDS * 8;
    assert_eq!(&word[pointer..pointer + 8], &[0; 8]);

    let mut stale_editor = Editor::open(base_document(b"stale")).expect("stale package opens");
    let stale = stale_editor.edit();
    stale_editor.set(sample()).expect("publish stale package");
    assert!(stale_editor.apply(stale).is_err());
}

#[test]
fn round_trip_parses_ctb_controls_and_shared_general_metadata() {
    let header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0).with_save_ui_strings(true),
        0,
        None,
    )
    .expect("button header");
    let general = GeneralInfo::new(
        GeneralFlags::default().with_save_text(true),
        Some(WString::new("Bold").expect("custom text")),
        None,
        None,
        None,
    )
    .expect("general info");
    let button = Control::new(
        header,
        None,
        Some(Data::new(general, vec![0]).expect("button data")),
    )
    .expect("button");
    let toolbar = Toolbar::new(
        XString::new("Custom").expect("custom toolbar name"),
        Header::new(
            1,
            Restrictions::new(Type::Basic).expect("basic restrictions"),
            0,
            Flags::default(),
            WString::new("Toolbar").expect("toolbar name"),
        )
        .expect("toolbar header"),
        [0xA5; 100],
        0,
        0,
        0xCAFE,
    )
    .expect("custom toolbar")
    .with_controls(vec![button])
    .expect("toolbar controls");
    let value = CommandBars::new(vec![Entry::Toolbar(
        ToolbarWrapper::new(vec![
            Customization::toolbar(toolbar).expect("customization"),
        ])
        .expect("wrapper"),
    )])
    .expect("command bars");

    let bytes = to_bytes(&value).expect("serialize");
    let parsed = parse_bytes(&bytes).expect("parse");
    assert_eq!(to_bytes(&parsed).expect("serialize parsed"), bytes);
    let toolbar = match &parsed.entries()[0] {
        Entry::Toolbar(wrapper) => match &wrapper.customizations()[0].data() {
            CustomizationData::Toolbar(toolbar) => toolbar,
            _ => panic!("custom toolbar"),
        },
        _ => panic!("toolbar wrapper"),
    };
    assert_eq!(toolbar.controls().len(), 1);
    let data = toolbar.controls()[0].data().expect("TBCData");
    assert_eq!(
        data.general().custom_text().expect("custom text").text(),
        "Bold"
    );
    assert_eq!(data.specific(), &[0]);
}

#[test]
fn round_trip_preserves_menu_and_combo_specific_boundaries() {
    let mut menu_specific = 1u32.to_le_bytes().to_vec();
    menu_specific.extend_from_slice(&WString::new("Menu").expect("menu name").to_bytes());
    let mut combo_specific = 1i16.to_le_bytes().to_vec();
    combo_specific.extend_from_slice(&WString::new("One").expect("combo item").to_bytes());
    combo_specific.extend_from_slice(&(-1i16).to_le_bytes());
    combo_specific.extend_from_slice(&(-1i16).to_le_bytes());
    combo_specific.extend_from_slice(&0i16.to_le_bytes());
    combo_specific.extend_from_slice(&(-1i16).to_le_bytes());
    combo_specific.extend_from_slice(&WString::new("").expect("edit text").to_bytes());

    let value = command_bars_with_controls(vec![
        control(ControlType::Popup, 1, None, menu_specific.clone()),
        control(ControlType::DropDown, 1, None, combo_specific.clone()),
    ]);
    let bytes = to_bytes(&value).expect("serialize");
    let parsed = parse_bytes(&bytes).expect("parse");
    assert_eq!(to_bytes(&parsed).expect("serialize parsed"), bytes);
    let toolbar = match &parsed.entries()[0] {
        Entry::Toolbar(wrapper) => match wrapper.customizations()[0].data() {
            CustomizationData::Toolbar(toolbar) => toolbar,
            _ => panic!("custom toolbar"),
        },
        _ => panic!("toolbar wrapper"),
    };
    assert_eq!(
        toolbar.controls()[0].data().expect("menu data").specific(),
        menu_specific
    );
    assert_eq!(
        toolbar.controls()[1].data().expect("combo data").specific(),
        combo_specific
    );
}

#[test]
fn rejects_zero_length_combo_item_array_at_the_control_boundary() {
    let value = command_bars_with_controls(vec![control(
        ControlType::DropDown,
        1,
        None,
        0i16.to_le_bytes().to_vec(),
    )]);
    let bytes = to_bytes(&value).expect("serialize");
    assert!(parse_bytes(&bytes).is_err());
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
fn rejects_unknown_records_without_guessing_boundaries() {
    assert!(parse_bytes(&[0xFF, 0x10, 0x40]).is_err());
    assert!(parse_bytes(&[0xFF, 0x7F, 0x40]).is_err());
}

#[test]
fn rejects_invalid_command_string_and_macro_name_boundaries() {
    let value = sample();
    let mut bytes = to_bytes(&value).expect("serialize");
    let strings_tag = bytes
        .windows(5)
        .position(|window| window == [0x10, 0xFF, 0xFF, 6, 0])
        .expect("TcgSttbf tag");
    bytes[strings_tag + 3..strings_tag + 5].copy_from_slice(&3u16.to_le_bytes());
    assert!(parse_bytes(&bytes).is_err());

    let mut bytes = to_bytes(&value).expect("serialize");
    let names_tag = bytes
        .windows(5)
        .position(|window| window == [0x11, 1, 0, 4, 0])
        .expect("MacroNames tag");
    let name_terminator = names_tag + 1 + 2 + 2 + 2 + "AutoMacro".encode_utf16().count() * 2;
    bytes[name_terminator..name_terminator + 2].copy_from_slice(&[1, 0]);
    assert!(parse_bytes(&bytes).is_err());
}

#[test]
fn rejects_truncated_variable_strings_and_table_ranges() {
    assert!(XString::parse(&[1, 0, 0x41]).is_err());
    let payload = to_bytes(&sample()).expect("serialize");
    let fib = fib_with_pointer(4, payload.len() + 1);
    assert!(parse(&fib, &[0; 4]).is_err());
}

#[test]
fn rejects_truncated_tbc_specific_data_without_guessing_a_boundary() {
    let value = sample_with_button();
    let mut bytes = to_bytes(&value).expect("serialize");
    bytes.pop();
    assert!(parse_bytes(&bytes).is_err());
}

#[test]
fn rejects_invalid_tbc_command_type() {
    let value = sample();
    let mut bytes = to_bytes(&value).expect("serialize");
    let control = control(ControlType::Label, 0x1234, Some(0x0000_1231), Vec::new());
    let encoded = to_control_bytes(&control).expect("control bytes");
    let offset = bytes
        .windows(encoded.len())
        .position(|window| window == encoded.as_slice())
        .expect("encoded control");
    let command_offset = offset + 11;
    bytes[command_offset..command_offset + 4].copy_from_slice(&0x0000_1232u32.to_le_bytes());
    assert!(parse_bytes(&bytes).is_err());
}

fn sample_with_button() -> CommandBars<'static> {
    let header = Header::new(
        1,
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
    .expect("custom toolbar")
    .with_controls(vec![control(ControlType::Button, 1, None, vec![0])])
    .expect("toolbar controls");
    CommandBars::new(vec![Entry::Toolbar(
        ToolbarWrapper::new(vec![
            Customization::toolbar(toolbar).expect("customization"),
        ])
        .expect("wrapper"),
    )])
    .expect("command bars")
}

fn command_bars_with_controls(controls: Vec<Control<'static>>) -> CommandBars<'static> {
    let header = Header::new(
        i16::try_from(controls.len()).expect("control count"),
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
    .expect("custom toolbar")
    .with_controls(controls)
    .expect("toolbar controls");
    CommandBars::new(vec![Entry::Toolbar(
        ToolbarWrapper::new(vec![
            Customization::toolbar(toolbar).expect("customization"),
        ])
        .expect("wrapper"),
    )])
    .expect("command bars")
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

fn base_document(table: &[u8]) -> Vec<u8> {
    let pointer_count = 117usize;
    let word_len = 154 + pointer_count * 8;
    let mut word = vec![0; word_len];
    word[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    word[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());

    let mut writer = OleWriter::new();
    writer
        .create_stream(&["WordDocument"], &word)
        .expect("WordDocument stream");
    writer
        .create_stream(&["0Table"], table)
        .expect("0Table stream");
    let mut output = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB write");
    output.into_inner()
}

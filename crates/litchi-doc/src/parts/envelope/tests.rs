use super::codec::{self, FIB_INDEX};
use super::model::{
    Attachment, Envelope, FollowUpStatus, Importance, MSO_ENVELOPE_CLSID, Message, PropertyValue,
    RecipientCollection, RecipientProperties, RecipientProperty, SecurityFlags, Sensitivity, Text,
    Version,
};
use crate::parts::fib::FileInformationBlock;

fn unicode(value: &str) -> Text {
    Text::unicode(value.encode_utf16().collect())
}

fn empty_collection() -> RecipientCollection {
    RecipientCollection::default()
}

fn sample() -> Envelope {
    let recipient = RecipientProperties {
        properties: vec![
            RecipientProperty {
                property_id: 0x3001,
                value: PropertyValue::Long(7),
            },
            RecipientProperty {
                property_id: 0x3002,
                value: PropertyValue::Null(0),
            },
            RecipientProperty {
                property_id: 0x3003,
                value: PropertyValue::Boolean(true),
            },
            RecipientProperty {
                property_id: 0x3004,
                value: PropertyValue::SystemTime { high: 1, low: 2 },
            },
            RecipientProperty {
                property_id: 0x3005,
                value: PropertyValue::Error(0x8004_0005),
            },
            RecipientProperty {
                property_id: 0x3006,
                value: PropertyValue::String8(b"recipient".to_vec().into_boxed_slice()),
            },
            RecipientProperty {
                property_id: 0x3007,
                value: PropertyValue::Unicode("Unicode".encode_utf16().collect()),
            },
            RecipientProperty {
                property_id: 0x3008,
                value: PropertyValue::Binary(vec![1, 2, 3].into_boxed_slice()),
            },
            RecipientProperty {
                property_id: 0x3009,
                value: PropertyValue::MultiString8(vec![
                    b"one".to_vec().into_boxed_slice(),
                    b"two".to_vec().into_boxed_slice(),
                ]),
            },
            RecipientProperty {
                property_id: 0x300A,
                value: PropertyValue::MultiBinary(vec![vec![4, 5].into_boxed_slice()]),
            },
        ],
    };
    let message = Message {
        version: Version::Office8,
        last_sent_time: 0,
        flag_status: FollowUpStatus::Flagged,
        reply_time: super::validation::MAX_MINUTE_TIME,
        request: unicode("reply"),
        sent_representing_entry_id: vec![1, 2, 3].into_boxed_slice(),
        sent_representing_name: unicode("sender"),
        internet_account_stamp: unicode("stamp"),
        internet_account_name: unicode("account"),
        expiry_time: super::validation::MAX_MINUTE_TIME,
        deferred_delivery_time: 0,
        delete_after_submit: false,
        security: SecurityFlags {
            signed: true,
            encrypted: false,
        },
        delivery_report: true,
        read_receipt: false,
        categories: unicode("category"),
        sensitivity: Sensitivity::Private,
        importance: Importance::High,
        subject: unicode("subject"),
        voting_options: b"yes;no".to_vec().into_boxed_slice(),
        reply_recipients: RecipientCollection {
            recipients: vec![recipient],
        },
        contact_link_recipients: Some(empty_collection()),
        recipients: empty_collection(),
        attachments: vec![Attachment {
            method: 1,
            name: "a.txt".encode_utf16().collect(),
            data: vec![0xDE, 0xAD].into_boxed_slice(),
        }],
        intro_text: Some("intro".encode_utf16().collect()),
        tail: vec![0xFA, 0xCE].into_boxed_slice(),
    };
    Envelope::from_message(message).expect("valid envelope fixture")
}

fn fib_with_pointer(offset: u32, length: u32) -> FileInformationBlock {
    let pointer_offset = 154 + FIB_INDEX * 8;
    let mut data = vec![0; pointer_offset + 8];
    data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    data[152..154].copy_from_slice(&((FIB_INDEX + 1) as u16).to_le_bytes());
    data[pointer_offset..pointer_offset + 4].copy_from_slice(&offset.to_le_bytes());
    data[pointer_offset + 4..pointer_offset + 8].copy_from_slice(&length.to_le_bytes());
    FileInformationBlock::parse(&data).expect("valid fixture FIB")
}

#[test]
fn known_envelope_round_trips_all_typed_properties() {
    let expected = sample();
    let bytes = expected.to_bytes().expect("encode envelope");
    assert_eq!(
        Envelope::parse_bytes(&bytes).expect("decode envelope"),
        expected
    );
}

#[test]
fn office6_uses_ansi_and_omits_unicode_only_fields() {
    let mut message = Message::default();
    message.version = Version::Office6;
    message.request = Text::ansi(b"reply".to_vec());
    message.sent_representing_name = Text::ansi(b"sender".to_vec());
    message.internet_account_stamp = Text::ansi(b"stamp".to_vec());
    message.internet_account_name = Text::ansi(b"account".to_vec());
    message.categories = Text::ansi(b"category".to_vec());
    message.subject = Text::ansi(b"subject".to_vec());
    message.contact_link_recipients = None;
    message.intro_text = None;
    let expected = Envelope::from_message(message).expect("valid Office 6 envelope");
    let bytes = expected.to_bytes().expect("encode Office 6 envelope");
    assert_eq!(
        Envelope::parse_bytes(&bytes).expect("decode Office 6 envelope"),
        expected
    );
}

#[test]
fn unknown_clsid_round_trips_as_bounded_opaque_payload() {
    let expected = Envelope::opaque([7; 16], vec![1, 2, 3]).expect("valid opaque envelope");
    let bytes = expected.to_bytes().expect("encode opaque envelope");
    assert_eq!(
        Envelope::parse_bytes(&bytes).expect("decode opaque envelope"),
        expected
    );
    assert!(expected.message().is_none());
}

#[test]
fn fib_pointer_reads_envelope_from_table_stream() {
    let bytes = sample().to_bytes().expect("encode envelope");
    let offset = 5u32;
    let mut table_stream = vec![0xCC; offset as usize];
    table_stream.extend_from_slice(&bytes);
    let fib = fib_with_pointer(offset, bytes.len() as u32);
    assert_eq!(
        codec::parse_fib(&fib, &table_stream).expect("read FIB range"),
        Some(sample())
    );
}

#[test]
fn rejects_malformed_version_flags_encoding_and_bounds() {
    let mut version = sample().to_bytes().expect("encode envelope");
    version[16..20].copy_from_slice(&7u32.to_le_bytes());
    assert!(Envelope::parse_bytes(&version).is_err());

    let mut security = sample().to_bytes().expect("encode envelope");
    let security_offset =
        16 + 4 * 4 + 2 + 5 * 2 + 4 + 3 + 2 + 6 * 2 + 2 + 5 * 2 + 2 + 7 * 2 + 4 * 2 + 4;
    security[security_offset..security_offset + 4].copy_from_slice(&4u32.to_le_bytes());
    assert!(Envelope::parse_bytes(&security).is_err());

    let mut tail = sample().to_bytes().expect("encode envelope");
    tail.extend_from_slice(&[0xAA, 0xBB]);
    let parsed = Envelope::parse_bytes(&tail).expect("supported body tail is opaque");
    assert_eq!(
        parsed.message().unwrap().tail.as_ref(),
        &[0xFA, 0xCE, 0xAA, 0xBB]
    );
    assert!(Envelope::parse_bytes(&[0; 15]).is_err());

    let mut message = sample().message().expect("sample must be typed").clone();
    message.contact_link_recipients = None;
    assert!(Envelope::from_message(message).is_err());
}

#[test]
fn rejects_unpaired_utf16_on_model_and_parse() {
    let mut message = sample().message().expect("sample must be typed").clone();
    message.subject = Text::unicode(vec![0xD800]);
    assert!(Envelope::from_message(message).is_err());

    let mut bytes = sample().to_bytes().expect("encode envelope");
    let subject = "subject".encode_utf16().collect::<Vec<_>>();
    let needle = subject
        .iter()
        .flat_map(|unit| unit.to_le_bytes())
        .collect::<Vec<_>>();
    let offset = bytes
        .windows(needle.len())
        .position(|window| window == needle.as_slice())
        .expect("subject bytes in fixture");
    bytes[offset..offset + 2].copy_from_slice(&0xD800u16.to_le_bytes());
    assert!(Envelope::parse_bytes(&bytes).is_err());
}

#[test]
fn rejects_known_clsid_with_opaque_payload() {
    assert!(Envelope::opaque(MSO_ENVELOPE_CLSID, vec![1, 2, 3]).is_err());
}

#[test]
fn office8_accepts_the_full_u16_character_length_without_confusing_bytes_and_chars() {
    let mut message = Message::default();
    message.subject = Text::unicode(vec![u16::from(b'A'); u16::MAX as usize]);
    let envelope = Envelope::from_message(message).expect("maximum Office 8 string is valid");
    let bytes = envelope
        .to_bytes()
        .expect("maximum Office 8 string encodes");
    assert_eq!(Envelope::parse_bytes(&bytes).unwrap(), envelope);
}

#[test]
fn model_rejects_recipient_and_attachment_bounds_before_encoding() {
    let mut message = sample().message().unwrap().clone();
    message.reply_recipients.recipients = vec![RecipientProperties::default(); 65_536 + 1];
    assert!(Envelope::from_message(message).is_err());

    let mut message = sample().message().unwrap().clone();
    message.attachments = vec![Attachment {
        method: 1,
        name: vec![u16::from(b'x'); u8::MAX as usize + 1].into_boxed_slice(),
        data: Box::default(),
    }];
    assert!(Envelope::from_message(message).is_err());
}

#[test]
fn semantic_patch_is_clone_first_reversible_and_source_checked() {
    let source = sample().snapshot().unwrap();
    let mut transaction = source.edit();
    transaction
        .update_message(|message| message.subject = unicode("edited"))
        .unwrap();
    let committed = transaction.commit().unwrap();
    assert_eq!(
        committed.snapshot().message().unwrap().subject,
        unicode("edited")
    );
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(committed.snapshot())
            .unwrap(),
        source
    );
    assert!(committed.patch().apply(&super::Snapshot::empty()).is_err());
}

#[test]
fn package_editor_appends_only_the_new_range_and_publishes_reversible_patches() {
    let prefix = b"opaque table bytes";
    let original = write_doc(prefix, None);
    let mut editor = super::Editor::open(original.clone()).unwrap();
    let committed = editor.set(sample()).unwrap();
    let written = committed.snapshot().finish().unwrap();
    assert_eq!(
        committed.package_patch().inverse().apply(&written).unwrap(),
        original
    );

    let mut ole = litchi_cfb::OleFile::open(std::io::Cursor::new(written.clone())).unwrap();
    let table = ole.open_stream(&["0Table"]).unwrap();
    assert!(table.starts_with(prefix));
    drop(table);
    drop(ole);

    let mut reopened = super::Editor::open(written.clone()).unwrap();
    assert_eq!(reopened.value(), Some(&sample()));
    let edited = reopened
        .update_message(|message| message.subject = unicode("edited"))
        .unwrap();
    assert_eq!(
        edited.snapshot().message().unwrap().tail.as_ref(),
        &[0xFA, 0xCE]
    );

    let cleared = super::Editor::open(edited.snapshot().finish().unwrap())
        .unwrap()
        .clear()
        .unwrap();
    let finished = cleared.snapshot().finish().unwrap();
    let mut ole = litchi_cfb::OleFile::open(std::io::Cursor::new(finished)).unwrap();
    let word = ole.open_stream(&["WordDocument"]).unwrap();
    assert_eq!(&word[154 + FIB_INDEX * 8..154 + FIB_INDEX * 8 + 8], &[0; 8]);
    let table = ole.open_stream(&["0Table"]).unwrap();
    assert!(table.starts_with(prefix));
}

fn write_doc(prefix: &[u8], payload: Option<&[u8]>) -> Vec<u8> {
    let mut table_stream = prefix.to_vec();
    let pointer = 154 + FIB_INDEX * 8;
    let mut word = vec![0u8; pointer + 8];
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[152..154].copy_from_slice(&((FIB_INDEX + 1) as u16).to_le_bytes());
    if let Some(payload) = payload {
        let offset = table_stream.len();
        table_stream.extend_from_slice(payload);
        word[pointer..pointer + 4].copy_from_slice(&u32::try_from(offset).unwrap().to_le_bytes());
        word[pointer + 4..pointer + 8]
            .copy_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    }

    let mut writer = litchi_cfb::OleWriter::new();
    writer.create_stream(&["WordDocument"], &word).unwrap();
    writer.create_stream(&["0Table"], &table_stream).unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

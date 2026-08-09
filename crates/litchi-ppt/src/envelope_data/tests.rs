#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::MAX_MINUTE_TIME;
use super::model::{
    EnvelopeData, EnvelopePayload, MSO_ENVELOPE_CLSID, MsoAttachment, MsoEnvelope, MsoEnvelopeText,
    MsoEnvelopeVersion, MsoFollowUpStatus, MsoImportance, MsoPropertyValue, MsoRecipientCollection,
    MsoRecipientProperties, MsoRecipientProperty, MsoSecurityFlags, MsoSensitivity,
};

fn empty_collection() -> MsoRecipientCollection {
    MsoRecipientCollection::default()
}

fn sample() -> EnvelopeData {
    let unicode = |value: &str| MsoEnvelopeText::Unicode(value.encode_utf16().collect());
    EnvelopeData {
        clsid: MSO_ENVELOPE_CLSID,
        payload: EnvelopePayload::Mso(MsoEnvelope {
            version: MsoEnvelopeVersion::Office8,
            last_sent_time: 0,
            flag_status: MsoFollowUpStatus::Flagged,
            reply_time: MAX_MINUTE_TIME,
            request: unicode("reply"),
            sent_representing_entry_id: vec![1, 2, 3],
            sent_representing_name: unicode("sender"),
            internet_account_stamp: unicode("stamp"),
            internet_account_name: unicode("account"),
            expiry_time: MAX_MINUTE_TIME,
            deferred_delivery_time: 0,
            delete_after_submit: false,
            security: MsoSecurityFlags {
                signed: true,
                encrypted: false,
            },
            delivery_report: true,
            read_receipt: false,
            categories: unicode("category"),
            sensitivity: MsoSensitivity::Private,
            importance: MsoImportance::High,
            subject: unicode("subject"),
            voting_options: b"yes;no".to_vec(),
            reply_recipients: MsoRecipientCollection {
                recipients: vec![MsoRecipientProperties {
                    properties: vec![
                        MsoRecipientProperty {
                            property_id: 0x3001,
                            value: MsoPropertyValue::Unicode("Recipient".encode_utf16().collect()),
                        },
                        MsoRecipientProperty {
                            property_id: 0x0c15,
                            value: MsoPropertyValue::Boolean(true),
                        },
                    ],
                }],
            },
            contact_link_recipients: Some(empty_collection()),
            recipients: empty_collection(),
            attachments: vec![MsoAttachment {
                method: 1,
                name: "a.txt".encode_utf16().collect(),
                data: vec![0xde, 0xad],
            }],
            intro_text: Some("intro".encode_utf16().collect()),
        }),
    }
}

#[test]
fn known_envelope_round_trips() {
    let expected = sample();
    let record = expected.to_record().unwrap();
    assert_eq!(EnvelopeData::parse(&record).unwrap(), expected);
}

#[test]
fn unknown_clsid_is_bounded_opaque_data() {
    let expected = EnvelopeData {
        clsid: [7; 16],
        payload: EnvelopePayload::Opaque(vec![1, 2, 3]),
    };
    assert_eq!(
        EnvelopeData::parse(&expected.to_record().unwrap()).unwrap(),
        expected
    );
}

#[test]
fn rejects_reserved_flags_and_version_mismatches() {
    let mut record = sample().to_record().unwrap();
    let security_offset =
        16 + 4 * 4 + 2 + 5 * 2 + 4 + 3 + 2 + 6 * 2 + 2 + 5 * 2 + 2 + 7 * 2 + 4 * 2 + 4;
    record.data[security_offset..security_offset + 4].copy_from_slice(&4u32.to_le_bytes());
    assert!(EnvelopeData::parse(&record).is_err());

    let mut value = sample();
    let EnvelopePayload::Mso(envelope) = &mut value.payload else {
        unreachable!();
    };
    envelope.contact_link_recipients = None;
    assert!(value.to_record().is_err());
}

#[test]
fn rejects_unpaired_utf16_on_parse_and_write() {
    let mut value = sample();
    let EnvelopePayload::Mso(envelope) = &mut value.payload else {
        unreachable!();
    };
    envelope.subject = MsoEnvelopeText::Unicode(vec![0xd800]);
    assert!(value.to_record().is_err());

    let mut record = sample().to_record().unwrap();
    let subject = "subject".encode_utf16().collect::<Vec<_>>();
    let bytes = subject
        .iter()
        .flat_map(|unit| unit.to_le_bytes())
        .collect::<Vec<_>>();
    let offset = record
        .data
        .windows(bytes.len())
        .position(|window| window == bytes)
        .unwrap();
    record.data[offset..offset + 2].copy_from_slice(&0xd800u16.to_le_bytes());
    assert!(EnvelopeData::parse(&record).is_err());
}

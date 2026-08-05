use super::{Address, CurrentRecipient, Slip, Text};
use crate::records::PptRecord;

fn text(value: &str) -> Text {
    Text::from_ansi_bytes(value.as_bytes().to_vec()).unwrap()
}

#[test]
fn routing_slip_round_trips_typed_state_and_undefined_bytes() {
    let mut originator = Address::new(text("origin"));
    originator.trailing_undefined = 0xa1;
    let mut first = Address::new(text("first"));
    first.trailing_undefined = 0xb2;
    let mut second = Address::new(text("second"));
    second.trailing_undefined = 0xc3;
    let mut slip = Slip::new(
        originator,
        vec![first, second],
        text("subject"),
        text("message"),
    )
    .unwrap();
    slip.current_recipient = CurrentRecipient::Recipient(2);
    slip.one_after_another = true;
    slip.return_when_done = true;
    slip.track_status = true;
    slip.document_routed = true;
    slip.cycle_completed = true;
    slip.unused1 = 0x1122_3344;
    slip.unused2 = 0x5566_7788;
    slip.trailing_undefined = vec![0xaa; 11];

    let record = slip.to_record().unwrap();
    assert_eq!(Slip::parse(&record).unwrap(), slip);
    assert_eq!(record.record_type_raw, 0x0406);
    assert_eq!(record.children, Vec::<PptRecord>::new());
}

#[test]
fn routing_slip_preserves_recipient_order() {
    let slip = Slip::new(
        Address::new(text("origin")),
        vec![Address::new(text("first")), Address::new(text("second"))],
        text("subject"),
        text("message"),
    )
    .unwrap();

    let parsed = Slip::parse(&slip.to_record().unwrap()).unwrap();
    let names = parsed
        .recipients
        .iter()
        .map(|address| address.text.to_string_lossy())
        .collect::<Vec<_>>();
    assert_eq!(names, ["first", "second"]);
}

#[test]
fn routing_slip_rejects_reserved_flags_and_bad_recipients() {
    let slip = Slip::new(Address::new(text("origin")), vec![], text(""), text("")).unwrap();
    let mut record = slip.to_record().unwrap();
    record.data[16..20].copy_from_slice(&0x08u32.to_le_bytes());
    assert!(Slip::parse(&record).is_err());

    let mut record = slip.to_record().unwrap();
    record.data[12..16].copy_from_slice(&2u32.to_le_bytes());
    assert!(Slip::parse(&record).is_err());
}

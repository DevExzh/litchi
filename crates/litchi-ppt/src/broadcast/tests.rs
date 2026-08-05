use super::codec::{BROADCAST_INFO_RECORD_TYPE, C_STRING_RECORD_TYPE, validate_system_time};
use super::model::{Broadcast, BroadcastProperties};
use crate::records::Record;
use crate::slide_sync::SystemTime;

fn time(hour: u16) -> SystemTime {
    SystemTime {
        year: 2026,
        month: 7,
        day_of_week: 0,
        day: 19,
        hour,
        minute: 30,
        second: 15,
        millisecond: 125,
    }
}

fn broadcast() -> Broadcast {
    Broadcast {
        title: Some("Quarterly update".into()),
        description: Some("Roadmap and results".into()),
        speaker: Some("Ada".into()),
        contact: Some("Grace".into()),
        remote_server_name: Some("CAMERA01".into()),
        email_address: Some("feedback@example.test".into()),
        email_name: Some("Feedback".into()),
        chat_url: Some("http://chat.example.test/room?id=7".into()),
        archive_directory: Some("C:\\Archive".into()),
        netshow_files_base_directory: Some("\\\\server\\share".into()),
        netshow_files_directory: Some("\\\\server\\share\\netshow".into()),
        netshow_server_name: Some("NETSHOW01".into()),
        ppt_files_base_directory: "\\\\server\\share".into(),
        ppt_files_directory: "\\\\server\\share\\ppt".into(),
        ppt_files_base_url: "http://slides.example.test/base".into(),
        user_name: "scheduler".into(),
        broadcast_date_time: "2026-07-19T09-30".into(),
        presentation_name: "quarterly.ppt".into(),
        asd_file_name: "\\\\server\\share\\stream.asd".into(),
        entry_id: Some("calendar-item-42".into()),
        properties: BroadcastProperties {
            send_audio: true,
            send_video: true,
            camera_remote: true,
            use_netshow: true,
            use_other_server: false,
            can_email: true,
            can_chat: true,
            archive: true,
            speaker_notes: true,
            quarter_screen: false,
            show_tools: true,
            record_only: false,
            start_time: time(9),
            end_time: time(10),
        },
    }
}

#[test]
fn round_trips_all_broadcast_atoms_and_flags() {
    let expected = broadcast();
    let record = expected.to_record().unwrap();
    let children = Record::parse_sequence_strict(&record.data, "test").unwrap();
    assert_eq!(children.len(), 21);
    assert_eq!(children[0].instance, 1);
    assert_eq!(children[19].instance, 20);
    assert_eq!(children[20].record_type_raw, BROADCAST_INFO_RECORD_TYPE);
    assert_eq!(Broadcast::parse(&record).unwrap(), expected);
}

#[test]
fn rejects_dependency_order_reserved_and_lexical_failures() {
    let mut value = broadcast();
    value.remote_server_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.netshow_server_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.email_name = None;
    assert!(value.validate().is_err());
    value = broadcast();
    value.chat_url = Some("https://wrong-scheme.example".into());
    assert!(value.validate().is_err());

    let valid = broadcast().to_record().unwrap();
    let children = Record::parse_sequence_strict(&valid.data, "test").unwrap();
    let mut data =
        super::codec::record_bytes(0, 2, C_STRING_RECORD_TYPE, &children[1].data).unwrap();
    data.extend_from_slice(
        &super::codec::record_bytes(0, 1, C_STRING_RECORD_TYPE, &children[0].data).unwrap(),
    );
    for child in &children[2..] {
        data.extend_from_slice(
            &super::codec::record_bytes(
                child.version,
                child.instance,
                child.record_type_raw,
                &child.data,
            )
            .unwrap(),
        );
    }
    let mut wrong_order = valid.clone();
    wrong_order.data = data;
    wrong_order.data_length = wrong_order.data.len() as u32;
    assert!(Broadcast::parse(&wrong_order).is_err());

    let mut reserved = valid;
    let atom_start = reserved.data.len() - 34;
    reserved.data[atom_start + 1] |= 0x10;
    assert!(Broadcast::parse(&reserved).is_err());
}

#[test]
fn validates_system_time_and_exact_strict_string_bounds() {
    let mut invalid = time(0);
    invalid.year = 2025;
    invalid.month = 2;
    invalid.day = 29;
    assert!(validate_system_time(invalid).is_err());
    invalid.year = 2024;
    assert!(validate_system_time(invalid).is_ok());
    let mut value = broadcast();
    value.archive_directory = Some("x".repeat(255));
    assert!(value.validate().is_err());
    value = broadcast();
    value.ppt_files_base_directory = "not-unc".into();
    assert!(value.validate().is_err());
    value = broadcast();
    value.user_name = "bad/name".into();
    assert!(value.validate().is_err());
}

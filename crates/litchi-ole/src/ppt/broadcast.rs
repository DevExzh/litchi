//! Inert PowerPoint 9 presentation-broadcast metadata from MS-PPT 2.4.17.

use crate::consts::PptRecordType;
use chrono::NaiveDate;

use super::package::{PptError, Result};
use super::records::PptRecord;
use super::slide_sync::PowerPointSystemTime;

const BROADCAST_CONTAINER_RECORD_TYPE: u16 = 0x177e;
const BROADCAST_INFO_RECORD_TYPE: u16 = 0x177f;
const C_STRING_RECORD_TYPE: u16 = 0x0fba;
const MAX_ENTRY_ID_BYTES: usize = 1 << 20;
const MAX_CONTAINER_BYTES: usize = 2 << 20;
const BROADCAST_FLAG_MASK: u16 = 0x0fff;

/// Fixed `BroadcastDocInfoAtom` values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointBroadcastProperties {
    pub send_audio: bool,
    pub send_video: bool,
    pub camera_remote: bool,
    pub use_netshow: bool,
    pub use_other_server: bool,
    pub can_email: bool,
    pub can_chat: bool,
    pub archive: bool,
    pub speaker_notes: bool,
    pub quarter_screen: bool,
    pub show_tools: bool,
    pub record_only: bool,
    pub start_time: PowerPointSystemTime,
    pub end_time: PowerPointSystemTime,
}

impl PowerPointBroadcastProperties {
    fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type_raw != BROADCAST_INFO_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 34
        {
            return Err(PptError::Corrupted(
                "BroadcastDocInfoAtom has an invalid record header or size".to_string(),
            ));
        }
        let flags = u16::from_le_bytes([record.data[0], record.data[1]]);
        if flags & !BROADCAST_FLAG_MASK != 0 {
            return Err(PptError::Corrupted(
                "BroadcastDocInfoAtom has nonzero reserved flags".to_string(),
            ));
        }
        Ok(Self {
            send_audio: flags & 0x001 != 0,
            send_video: flags & 0x002 != 0,
            camera_remote: flags & 0x004 != 0,
            use_netshow: flags & 0x008 != 0,
            use_other_server: flags & 0x010 != 0,
            can_email: flags & 0x020 != 0,
            can_chat: flags & 0x040 != 0,
            archive: flags & 0x080 != 0,
            speaker_notes: flags & 0x100 != 0,
            quarter_screen: flags & 0x200 != 0,
            show_tools: flags & 0x400 != 0,
            record_only: flags & 0x800 != 0,
            start_time: parse_system_time(&record.data[2..18])?,
            end_time: parse_system_time(&record.data[18..34])?,
        })
    }

    fn to_record_bytes(self) -> Result<Vec<u8>> {
        validate_system_time(self.start_time)?;
        validate_system_time(self.end_time)?;
        let flags = u16::from(self.send_audio)
            | u16::from(self.send_video) << 1
            | u16::from(self.camera_remote) << 2
            | u16::from(self.use_netshow) << 3
            | u16::from(self.use_other_server) << 4
            | u16::from(self.can_email) << 5
            | u16::from(self.can_chat) << 6
            | u16::from(self.archive) << 7
            | u16::from(self.speaker_notes) << 8
            | u16::from(self.quarter_screen) << 9
            | u16::from(self.show_tools) << 10
            | u16::from(self.record_only) << 11;
        let mut data = Vec::with_capacity(34);
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&system_time_bytes(self.start_time));
        data.extend_from_slice(&system_time_bytes(self.end_time));
        record_bytes(0, 0, BROADCAST_INFO_RECORD_TYPE, &data)
    }
}

/// One fully typed `BroadcastDocInfo9Container`.
///
/// Paths, server names, URLs, calendar identifiers, and all capability flags
/// are inert metadata. Parsing or writing this value never starts a broadcast,
/// connects to a server, sends mail, opens a URL, or reads an ASD file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointBroadcast {
    pub title: Option<String>,
    pub description: Option<String>,
    pub speaker: Option<String>,
    pub contact: Option<String>,
    pub remote_server_name: Option<String>,
    pub email_address: Option<String>,
    pub email_name: Option<String>,
    pub chat_url: Option<String>,
    pub archive_directory: Option<String>,
    pub netshow_files_base_directory: Option<String>,
    pub netshow_files_directory: Option<String>,
    pub netshow_server_name: Option<String>,
    pub ppt_files_base_directory: String,
    pub ppt_files_directory: String,
    pub ppt_files_base_url: String,
    pub user_name: String,
    pub broadcast_date_time: String,
    pub presentation_name: String,
    pub asd_file_name: String,
    pub entry_id: Option<String>,
    pub properties: PowerPointBroadcastProperties,
}

/// All PowerPoint 9 broadcast descriptions in document order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointBroadcasts {
    pub broadcasts: Vec<PowerPointBroadcast>,
}

impl PowerPointBroadcasts {
    pub(crate) fn parse_document(document: &PptRecord) -> Result<Self> {
        let mut broadcasts = Vec::new();
        for record in document.versioned_binary_tag_records(9)? {
            if record.record_type_raw == BROADCAST_CONTAINER_RECORD_TYPE {
                broadcasts.push(PowerPointBroadcast::parse(&record)?);
            }
        }
        Ok(Self { broadcasts })
    }
}

impl PowerPointBroadcast {
    /// Parse one strict broadcast container and all specification-defined children.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type_raw != BROADCAST_CONTAINER_RECORD_TYPE
            || record.version != 0x0f
            || record.instance != 0
            || record.data.len() > MAX_CONTAINER_BYTES
        {
            return Err(PptError::Corrupted(
                "BroadcastDocInfo9Container has an invalid header or exceeds the resource cap"
                    .to_string(),
            ));
        }
        let children =
            PptRecord::parse_sequence_strict(&record.data, "BroadcastDocInfo9Container")?;
        let mut strings: [Option<String>; 20] = std::array::from_fn(|_| None);
        let mut last_instance = 0u16;
        let mut properties = None;
        for child in children {
            if child.record_type_raw == C_STRING_RECORD_TYPE {
                if properties.is_some()
                    || child.instance == 0
                    || child.instance > 20
                    || child.instance <= last_instance
                {
                    return Err(PptError::Corrupted(
                        "Broadcast container has duplicate or out-of-order string atoms"
                            .to_string(),
                    ));
                }
                let descriptor = string_descriptor(child.instance).unwrap();
                strings[usize::from(child.instance - 1)] = Some(parse_string(&child, descriptor)?);
                last_instance = child.instance;
            } else if child.record_type_raw == BROADCAST_INFO_RECORD_TYPE && properties.is_none() {
                properties = Some(PowerPointBroadcastProperties::parse(&child)?);
            } else {
                return Err(PptError::Corrupted(
                    "Broadcast container has an unexpected or duplicate child".to_string(),
                ));
            }
        }
        let properties = properties.ok_or_else(|| {
            PptError::Corrupted("Broadcast container is missing BroadcastDocInfoAtom".to_string())
        })?;
        for required in 13..=19 {
            if strings[required - 1].is_none() {
                return Err(PptError::Corrupted(format!(
                    "Broadcast container is missing required CString instance {required}"
                )));
            }
        }
        let value = Self {
            title: strings[0].take(),
            description: strings[1].take(),
            speaker: strings[2].take(),
            contact: strings[3].take(),
            remote_server_name: strings[4].take(),
            email_address: strings[5].take(),
            email_name: strings[6].take(),
            chat_url: strings[7].take(),
            archive_directory: strings[8].take(),
            netshow_files_base_directory: strings[9].take(),
            netshow_files_directory: strings[10].take(),
            netshow_server_name: strings[11].take(),
            ppt_files_base_directory: strings[12].take().unwrap(),
            ppt_files_directory: strings[13].take().unwrap(),
            ppt_files_base_url: strings[14].take().unwrap(),
            user_name: strings[15].take().unwrap(),
            broadcast_date_time: strings[16].take().unwrap(),
            presentation_name: strings[17].take().unwrap(),
            asd_file_name: strings[18].take().unwrap(),
            entry_id: strings[19].take(),
            properties,
        };
        value.validate()?;
        Ok(value)
    }

    /// Validate lexical, size, date, and cross-field requirements.
    pub fn validate(&self) -> Result<()> {
        for (instance, value) in self.string_values() {
            if let Some(value) = value {
                validate_string(value, string_descriptor(instance).unwrap())?;
            }
        }
        validate_system_time(self.properties.start_time)?;
        validate_system_time(self.properties.end_time)?;
        if self.properties.camera_remote && self.remote_server_name.is_none() {
            return Err(PptError::Corrupted(
                "Remote-camera broadcast is missing BCRexServerNameAtom".to_string(),
            ));
        }
        if self.properties.use_netshow
            && (self.netshow_files_directory.is_none() || self.netshow_server_name.is_none())
        {
            return Err(PptError::Corrupted(
                "NetShow broadcast is missing its files directory or server name".to_string(),
            ));
        }
        if self.properties.can_email && self.email_name.is_none() {
            return Err(PptError::Corrupted(
                "Email-enabled broadcast is missing BCEmailNameAtom".to_string(),
            ));
        }
        Ok(())
    }

    /// Encode a canonical broadcast container in specification child order.
    pub fn to_record(&self) -> Result<PptRecord> {
        self.validate()?;
        let mut data = Vec::new();
        for (instance, value) in self.string_values() {
            if let Some(value) = value {
                let encoded = encode_utf16(value);
                data.extend_from_slice(&record_bytes(0, instance, C_STRING_RECORD_TYPE, &encoded)?);
            }
        }
        data.extend_from_slice(&self.properties.to_record_bytes()?);
        if data.len() > MAX_CONTAINER_BYTES {
            return Err(PptError::Corrupted(
                "Broadcast container exceeds the resource cap".to_string(),
            ));
        }
        let data_length = u32::try_from(data.len())
            .map_err(|_| PptError::Corrupted("Broadcast container length overflow".to_string()))?;
        Ok(PptRecord {
            record_type: PptRecordType::from(BROADCAST_CONTAINER_RECORD_TYPE),
            record_type_raw: BROADCAST_CONTAINER_RECORD_TYPE,
            version: 0x0f,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }

    fn string_values(&self) -> [(u16, Option<&str>); 20] {
        [
            (1, self.title.as_deref()),
            (2, self.description.as_deref()),
            (3, self.speaker.as_deref()),
            (4, self.contact.as_deref()),
            (5, self.remote_server_name.as_deref()),
            (6, self.email_address.as_deref()),
            (7, self.email_name.as_deref()),
            (8, self.chat_url.as_deref()),
            (9, self.archive_directory.as_deref()),
            (10, self.netshow_files_base_directory.as_deref()),
            (11, self.netshow_files_directory.as_deref()),
            (12, self.netshow_server_name.as_deref()),
            (13, Some(&self.ppt_files_base_directory)),
            (14, Some(&self.ppt_files_directory)),
            (15, Some(&self.ppt_files_base_url)),
            (16, Some(&self.user_name)),
            (17, Some(&self.broadcast_date_time)),
            (18, Some(&self.presentation_name)),
            (19, Some(&self.asd_file_name)),
            (20, self.entry_id.as_deref()),
        ]
    }
}

#[derive(Clone, Copy)]
struct StringDescriptor {
    instance: u16,
    max_bytes: usize,
    kind: StringKind,
    name: &'static str,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum StringKind {
    Unicode,
    Printable,
    Machine,
    HttpUrl,
    UncOrLocal,
    Unc,
    UncOrHttp,
    FileFragment,
}

fn string_descriptor(instance: u16) -> Option<StringDescriptor> {
    let (max_bytes, kind, name) = match instance {
        1 => (510, StringKind::Unicode, "BCTitleAtom"),
        2 => (2040, StringKind::Unicode, "BCDescriptionAtom"),
        3 => (510, StringKind::Printable, "BCSpeakerAtom"),
        4 => (510, StringKind::Printable, "BCContactAtom"),
        5 => (510, StringKind::Machine, "BCRexServerNameAtom"),
        6 => (510, StringKind::Printable, "BCEmailAddressAtom"),
        7 => (510, StringKind::Printable, "BCEmailNameAtom"),
        8 => (4164, StringKind::HttpUrl, "BCChatUrlAtom"),
        9 => (508, StringKind::UncOrLocal, "BCArchiveDirAtom"),
        10 => (508, StringKind::Unc, "BCNetShowFilesBaseDirAtom"),
        11 => (492, StringKind::Unc, "BCNetShowFilesDirAtom"),
        12 => (510, StringKind::Machine, "BCNetShowServerNameAtom"),
        13 => (508, StringKind::Unc, "BCPptFilesBaseDirAtom"),
        14 => (492, StringKind::Unc, "BCPptFilesDirAtom"),
        15 => (4116, StringKind::UncOrHttp, "BCPptFilesBaseUrlAtom"),
        16 => (508, StringKind::FileFragment, "BCUserNameAtom"),
        17 => (518, StringKind::FileFragment, "BCBroadcastDateTimeAtom"),
        18 => (508, StringKind::FileFragment, "BCPresentationNameAtom"),
        19 => (518, StringKind::Unc, "BCAsdFileNameAtom"),
        20 => (MAX_ENTRY_ID_BYTES, StringKind::Unicode, "BCEntryIDAtom"),
        _ => return None,
    };
    Some(StringDescriptor {
        instance,
        max_bytes,
        kind,
        name,
    })
}

fn parse_string(record: &PptRecord, descriptor: StringDescriptor) -> Result<String> {
    if record.version != 0
        || record.instance != descriptor.instance
        || record.data.len() > descriptor.max_bytes
        || record.data.len() % 2 != 0
    {
        return Err(PptError::Corrupted(format!(
            "{} has an invalid record header or size",
            descriptor.name
        )));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if descriptor.kind == StringKind::Unicode && unit == 0 {
            break;
        }
        units.push(unit);
    }
    let value = String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted(format!("{} contains invalid UTF-16", descriptor.name)))?;
    validate_string(&value, descriptor)?;
    Ok(value)
}

fn validate_string(value: &str, descriptor: StringDescriptor) -> Result<()> {
    let byte_len = value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| PptError::Corrupted(format!("{} length overflow", descriptor.name)))?;
    if byte_len > descriptor.max_bytes {
        return Err(PptError::Corrupted(format!(
            "{} exceeds its MS-PPT byte limit",
            descriptor.name
        )));
    }
    if value.encode_utf16().any(|unit| unit == 0) {
        return Err(PptError::Corrupted(format!(
            "{} contains an embedded null",
            descriptor.name
        )));
    }
    match descriptor.kind {
        StringKind::Unicode => {},
        StringKind::Printable => {
            if value
                .encode_utf16()
                .any(|unit| matches!(unit, 0x0000..=0x001f | 0x007f..=0x009f))
            {
                return Err(PptError::Corrupted(format!(
                    "{} contains a forbidden control character",
                    descriptor.name
                )));
            }
        },
        StringKind::Machine => validate_machine_name(value, descriptor.name)?,
        StringKind::HttpUrl => validate_http_url(value, descriptor.name)?,
        StringKind::UncOrLocal => validate_unc_or_local_path(value, descriptor.name)?,
        StringKind::Unc => validate_unc_path(value, descriptor.name)?,
        StringKind::UncOrHttp => {
            if is_http_url(value) {
                validate_http_url(value, descriptor.name)?;
            } else {
                validate_unc_path(value, descriptor.name)?;
            }
        },
        StringKind::FileFragment => validate_file_fragment(value, descriptor.name)?,
    }
    Ok(())
}

fn validate_machine_name(value: &str, name: &str) -> Result<()> {
    if value.is_empty()
        || value.starts_with(['.', '-'])
        || value.ends_with(['.', '-'])
        || value.chars().any(|ch| {
            ch.is_control()
                || ch.is_whitespace()
                || matches!(ch, '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|')
        })
    {
        return invalid_lexical(name, "computer name");
    }
    Ok(())
}

fn validate_file_fragment(value: &str, name: &str) -> Result<()> {
    if value.is_empty()
        || matches!(value, "." | "..")
        || value.chars().any(|ch| {
            ch.is_control() || matches!(ch, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*')
        })
    {
        return invalid_lexical(name, "file or directory name fragment");
    }
    Ok(())
}

fn validate_unc_path(value: &str, name: &str) -> Result<()> {
    if !value.starts_with("\\\\") {
        return invalid_lexical(name, "UNC path");
    }
    let mut components = value[2..].split('\\');
    if components.next().is_none_or(str::is_empty) || components.next().is_none_or(str::is_empty) {
        return invalid_lexical(name, "UNC path");
    }
    if value.chars().any(char::is_control) {
        return invalid_lexical(name, "UNC path");
    }
    Ok(())
}

fn validate_unc_or_local_path(value: &str, name: &str) -> Result<()> {
    if value.starts_with("\\\\") {
        return validate_unc_path(value, name);
    }
    if value.is_empty() || value.chars().any(char::is_control) {
        return invalid_lexical(name, "UNC or local path");
    }
    Ok(())
}

fn is_http_url(value: &str) -> bool {
    value
        .get(..7)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("http://"))
}

fn validate_http_url(value: &str, name: &str) -> Result<()> {
    if !is_http_url(value) {
        return invalid_lexical(name, "HTTP URI");
    }
    let remainder = &value[7..];
    let authority_end = remainder.find(['/', '?', '#']).unwrap_or(remainder.len());
    if authority_end == 0 || !value.is_ascii() {
        return invalid_lexical(name, "HTTP URI");
    }
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            if index + 2 >= bytes.len()
                || !bytes[index + 1].is_ascii_hexdigit()
                || !bytes[index + 2].is_ascii_hexdigit()
            {
                return invalid_lexical(name, "HTTP URI");
            }
            index += 3;
            continue;
        }
        if !(byte.is_ascii_alphanumeric() || b"-._~:/?#[]@!$&'()*+,;=".contains(&byte)) {
            return invalid_lexical(name, "HTTP URI");
        }
        index += 1;
    }
    Ok(())
}

fn invalid_lexical<T>(name: &str, kind: &str) -> Result<T> {
    Err(PptError::Corrupted(format!("{name} is not a valid {kind}")))
}

fn parse_system_time(data: &[u8]) -> Result<PowerPointSystemTime> {
    if data.len() != 16 {
        return Err(PptError::Corrupted(
            "PowerPoint SYSTEMTIME must contain 16 bytes".to_string(),
        ));
    }
    let field = |index: usize| u16::from_le_bytes([data[index], data[index + 1]]);
    let value = PowerPointSystemTime {
        year: field(0),
        month: field(2),
        day_of_week: field(4),
        day: field(6),
        hour: field(8),
        minute: field(10),
        second: field(12),
        millisecond: field(14),
    };
    validate_system_time(value)?;
    Ok(value)
}

fn validate_system_time(value: PowerPointSystemTime) -> Result<()> {
    if !(1601..=30_827).contains(&value.year)
        || value.day_of_week > 6
        || value.hour > 23
        || value.minute > 59
        || value.second > 59
        || value.millisecond > 999
        || NaiveDate::from_ymd_opt(value.year.into(), value.month.into(), value.day.into())
            .is_none()
    {
        return Err(PptError::Corrupted(
            "PowerPoint broadcast contains an invalid SYSTEMTIME".to_string(),
        ));
    }
    Ok(())
}

fn system_time_bytes(value: PowerPointSystemTime) -> [u8; 16] {
    let fields = [
        value.year,
        value.month,
        value.day_of_week,
        value.day,
        value.hour,
        value.minute,
        value.second,
        value.millisecond,
    ];
    let mut bytes = [0u8; 16];
    for (field, output) in fields.into_iter().zip(bytes.chunks_exact_mut(2)) {
        output.copy_from_slice(&field.to_le_bytes());
    }
    bytes
}

fn encode_utf16(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let data_length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("Broadcast record length overflow".to_string()))?;
    let mut bytes = Vec::with_capacity(8 + data.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&data_length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn time(hour: u16) -> PowerPointSystemTime {
        PowerPointSystemTime {
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

    fn broadcast() -> PowerPointBroadcast {
        PowerPointBroadcast {
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
            properties: PowerPointBroadcastProperties {
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
        let children = PptRecord::parse_sequence_strict(&record.data, "test").unwrap();
        assert_eq!(children.len(), 21);
        assert_eq!(children[0].instance, 1);
        assert_eq!(children[19].instance, 20);
        assert_eq!(children[20].record_type_raw, BROADCAST_INFO_RECORD_TYPE);
        assert_eq!(PowerPointBroadcast::parse(&record).unwrap(), expected);
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
        let children = PptRecord::parse_sequence_strict(&valid.data, "test").unwrap();
        let mut data = record_bytes(0, 2, C_STRING_RECORD_TYPE, &children[1].data).unwrap();
        data.extend_from_slice(
            &record_bytes(0, 1, C_STRING_RECORD_TYPE, &children[0].data).unwrap(),
        );
        for child in &children[2..] {
            data.extend_from_slice(
                &record_bytes(
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
        assert!(PowerPointBroadcast::parse(&wrong_order).is_err());

        let mut reserved = valid;
        let atom_start = reserved.data.len() - 34;
        reserved.data[atom_start + 1] |= 0x10;
        assert!(PowerPointBroadcast::parse(&reserved).is_err());
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
}

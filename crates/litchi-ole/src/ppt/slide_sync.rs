//! Inert PowerPoint 12 slide-library synchronization metadata.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use chrono::NaiveDate;

/// A validated Windows `SYSTEMTIME` used by PowerPoint synchronization records.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointSystemTime {
    /// Gregorian year.
    pub year: u16,
    /// Month (`1` through `12`).
    pub month: u16,
    /// `0` is Sunday and `6` is Saturday.
    pub day_of_week: u16,
    /// Day of the month.
    pub day: u16,
    /// Hour (`0` through `23`).
    pub hour: u16,
    /// Minute (`0` through `59`).
    pub minute: u16,
    /// Second (`0` through `59`).
    pub second: u16,
    /// Millisecond (`0` through `999`).
    pub millisecond: u16,
}

/// Read-only metadata connecting a presentation slide to a slide-library item.
///
/// Parsing this structure never accesses the URL or performs synchronization.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSlideSyncInfo {
    /// Unique server-side slide identifier.
    pub server_slide_id: String,
    /// HTTP URL of the source slide library.
    pub slide_library_url: String,
    /// Time the server-side slide was last modified.
    pub server_modified: PowerPointSystemTime,
    /// Time the slide was inserted into this presentation.
    pub client_inserted: PowerPointSystemTime,
}

impl PowerPointSlideSyncInfo {
    /// Parse the optional synchronization container directly below `root`.
    pub fn parse(root: &PptRecord) -> Result<Option<Self>> {
        let containers = root.find_children(PptRecordType::RoundTripSlideSyncInfo12);
        if containers.len() > 1 {
            return Err(PptError::Corrupted(
                "Slide contains duplicate RoundTripSlideSyncInfo12 containers".to_string(),
            ));
        }
        let Some(container) = containers.first() else {
            return Ok(None);
        };
        if container.version != 0x0f || container.instance != 0 {
            return Err(PptError::Corrupted(
                "RoundTripSlideSyncInfo12 has an invalid record header".to_string(),
            ));
        }
        let children =
            PptRecord::parse_sequence_strict(&container.data, "RoundTripSlideSyncInfo12")?;
        if children.len() != 3 {
            return Err(PptError::Corrupted(
                "RoundTripSlideSyncInfo12 must contain exactly three records".to_string(),
            ));
        }

        let server_slide_id = parse_unicode_string(&children[0], 0, "ServerIdAtom")?;
        let slide_library_url = parse_unicode_string(&children[1], 1, "SlideLibUrlAtom")?;
        let parsed_url = url::Url::parse(&slide_library_url).map_err(|_| {
            PptError::Corrupted("SlideLibUrlAtom does not contain a valid HTTP URI".to_string())
        })?;
        if parsed_url.scheme() != "http" {
            return Err(PptError::Corrupted(
                "SlideLibUrlAtom does not use the HTTP scheme".to_string(),
            ));
        }

        let atom = &children[2];
        if atom.record_type != PptRecordType::RoundTripSlideSyncInfoAtom12
            || atom.version != 0
            || atom.instance != 0
            || atom.data.len() != 32
        {
            return Err(PptError::Corrupted(
                "SlideSyncInfoAtom12 has an invalid record header or size".to_string(),
            ));
        }

        Ok(Some(Self {
            server_slide_id,
            slide_library_url,
            server_modified: parse_system_time(&atom.data[..16], "server modified time")?,
            client_inserted: parse_system_time(&atom.data[16..], "client inserted time")?,
        }))
    }
}

fn parse_unicode_string(record: &PptRecord, instance: u16, name: &str) -> Result<String> {
    if record.record_type != PptRecordType::CString
        || record.version != 0
        || record.instance != instance
        || record.data.len() & 1 != 0
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        if matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f) {
            return Err(PptError::Corrupted(format!(
                "{name} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted(format!("{name} contains invalid UTF-16")))
}

fn parse_system_time(data: &[u8], name: &str) -> Result<PowerPointSystemTime> {
    let field = |index: usize| u16::from_le_bytes([data[index], data[index + 1]]);
    let time = PowerPointSystemTime {
        year: field(0),
        month: field(2),
        day_of_week: field(4),
        day: field(6),
        hour: field(8),
        minute: field(10),
        second: field(12),
        millisecond: field(14),
    };
    if !(1601..=30_827).contains(&time.year)
        || time.day_of_week > 6
        || time.hour > 23
        || time.minute > 59
        || time.second > 59
        || time.millisecond > 999
        || NaiveDate::from_ymd_opt(time.year.into(), time.month.into(), time.day.into()).is_none()
    {
        return Err(PptError::Corrupted(format!(
            "Slide synchronization {name} is not a valid SYSTEMTIME"
        )));
    }
    Ok(time)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&(data.len() as u32).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn utf16(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn system_time(fields: [u16; 8]) -> Vec<u8> {
        fields.into_iter().flat_map(u16::to_le_bytes).collect()
    }

    fn valid_children() -> Vec<u8> {
        let server = record_bytes(0, 0, 4026, &utf16("server-slide-42"));
        let url = record_bytes(0, 1, 4026, &utf16("http://example.com/slides?id=42"));
        let mut times = system_time([2024, 2, 4, 29, 23, 59, 58, 999]);
        times.extend_from_slice(&system_time([1601, 1, 1, 1, 0, 0, 0, 0]));
        let atom = record_bytes(0, 0, 0x3715, &times);
        [server, url, atom].concat()
    }

    fn container(version: u16, instance: u16, children: &[u8]) -> PptRecord {
        let bytes = record_bytes(version, instance, 0x3714, children);
        PptRecord::parse(&bytes, 0).unwrap().0
    }

    fn root(children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            version: 0x0f,
            instance: 0,
            record_type: PptRecordType::Slide,
            record_type_raw: PptRecordType::Slide.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_slide_library_synchronization_metadata() {
        assert_eq!(
            PowerPointSlideSyncInfo::parse(&root(Vec::new())).unwrap(),
            None
        );
        let parsed =
            PowerPointSlideSyncInfo::parse(&root(vec![container(0x0f, 0, &valid_children())]))
                .unwrap()
                .unwrap();

        assert_eq!(parsed.server_slide_id, "server-slide-42");
        assert_eq!(parsed.slide_library_url, "http://example.com/slides?id=42");
        assert_eq!(parsed.server_modified.year, 2024);
        assert_eq!(parsed.server_modified.day, 29);
        assert_eq!(parsed.server_modified.millisecond, 999);
        assert_eq!(parsed.client_inserted.year, 1601);
    }

    #[test]
    fn rejects_malformed_slide_sync_container_structure_and_strings() {
        let valid = valid_children();
        let server = record_bytes(0, 0, 4026, &utf16("id"));
        let url = record_bytes(0, 1, 4026, &utf16("http://example.com"));
        let atom = valid[valid.len() - 40..].to_vec();
        let reordered = [url.clone(), server.clone(), atom.clone()].concat();
        let extra = [valid.clone(), record_bytes(0, 0, 9999, &[])].concat();
        for document in [
            root(vec![container(0, 0, &valid)]),
            root(vec![container(0x0f, 1, &valid)]),
            root(vec![container(0x0f, 0, &valid), container(0x0f, 0, &valid)]),
            root(vec![container(0x0f, 0, &valid[..valid.len() - 40])]),
            root(vec![container(0x0f, 0, &reordered)]),
            root(vec![container(0x0f, 0, &extra)]),
        ] {
            assert!(PowerPointSlideSyncInfo::parse(&document).is_err());
        }

        for (server, url) in [
            (record_bytes(1, 0, 4026, &utf16("id")), utf16("http://a")),
            (record_bytes(0, 1, 4026, &utf16("id")), utf16("http://a")),
            (record_bytes(0, 0, 4026, b"i"), utf16("http://a")),
            (record_bytes(0, 0, 4026, &[1, 0]), utf16("http://a")),
            (record_bytes(0, 0, 4026, &[0x00, 0xd8]), utf16("http://a")),
            (record_bytes(0, 0, 4026, &utf16("id")), utf16("https://a")),
            (record_bytes(0, 0, 4026, &utf16("id")), utf16("not a URI")),
        ] {
            let url = record_bytes(0, 1, 4026, &url);
            let children = [server, url, atom.clone()].concat();
            assert!(
                PowerPointSlideSyncInfo::parse(&root(vec![container(0x0f, 0, &children)])).is_err()
            );
        }

        for invalid_url in [
            record_bytes(1, 1, 4026, &utf16("http://example.com")),
            record_bytes(0, 0, 4026, &utf16("http://example.com")),
            record_bytes(0, 1, 4026, b"x"),
            record_bytes(0, 1, 4026, &[0x00, 0xd8]),
        ] {
            let children = [server.clone(), invalid_url, atom.clone()].concat();
            assert!(
                PowerPointSlideSyncInfo::parse(&root(vec![container(0x0f, 0, &children)])).is_err()
            );
        }
    }

    #[test]
    fn rejects_invalid_slide_sync_atom_and_system_times() {
        let server = record_bytes(0, 0, 4026, &utf16("id"));
        let url = record_bytes(0, 1, 4026, &utf16("http://example.com"));

        for (version, instance, payload) in [
            (1, 0, vec![0; 32]),
            (0, 1, vec![0; 32]),
            (0, 0, vec![0; 31]),
            (0, 0, vec![0; 33]),
        ] {
            let atom = record_bytes(version, instance, 0x3715, &payload);
            let children = [server.clone(), url.clone(), atom].concat();
            assert!(
                PowerPointSlideSyncInfo::parse(&root(vec![container(0x0f, 0, &children)])).is_err()
            );
        }

        for invalid in [
            [1600, 1, 0, 1, 0, 0, 0, 0],
            [30_828, 1, 0, 1, 0, 0, 0, 0],
            [2024, 0, 0, 1, 0, 0, 0, 0],
            [2023, 2, 0, 29, 0, 0, 0, 0],
            [2024, 1, 7, 1, 0, 0, 0, 0],
            [2024, 1, 0, 1, 24, 0, 0, 0],
            [2024, 1, 0, 1, 0, 60, 0, 0],
            [2024, 1, 0, 1, 0, 0, 60, 0],
            [2024, 1, 0, 1, 0, 0, 0, 1000],
        ] {
            let mut times = system_time(invalid);
            times.extend_from_slice(&system_time([2024, 1, 1, 1, 0, 0, 0, 0]));
            let atom = record_bytes(0, 0, 0x3715, &times);
            let children = [server.clone(), url.clone(), atom].concat();
            assert!(
                PowerPointSlideSyncInfo::parse(&root(vec![container(0x0f, 0, &children)])).is_err()
            );
        }
    }
}

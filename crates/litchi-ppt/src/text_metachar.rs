//! Header/footer metacharacter atoms (MS-PPT 2.9.47-2.9.52).
//!
//! Metacharacter atoms mark placeholder positions — slide number, header,
//! footer, and date/time — inside master and header/footer text bodies.
//! Everything here is inert: placeholders are never substituted, formatted,
//! or laid out.

use super::header_footer::PowerPointDateTimeFormatId;
use super::records::record::PptRecord;
use crate::consts::PptRecordType;
use crate::package::{PptError, Result};

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

/// Largest valid `DateTimeMCAtom` format identifier (MS-PPT 2.9.50);
/// `HeadersFootersAtom` permits 13 but metacharacters do not.
const MAX_METACHAR_FORMAT_ID: u8 = 12;
/// Byte length of a position-only metachar atom.
const POSITION_ATOM_LEN: usize = 4;
/// Byte length of a `DateTimeMCAtom` payload.
const DATE_TIME_ATOM_LEN: usize = 8;
/// Byte length of an `RTFDateTimeMCAtom` payload.
const RTF_DATE_TIME_ATOM_LEN: usize = 0x84;
/// Byte length of the `char2` RTF format string.
const RTF_FORMAT_LEN: usize = 128;

/// The kind of placeholder a metacharacter marks (MS-PPT 2.9.47-2.9.52).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PowerPointMetacharKind {
    /// `SlideNumberMCAtom`: a slide-number placeholder.
    SlideNumber,
    /// `HeaderMCAtom`: a header placeholder.
    Header,
    /// `FooterMCAtom`: a footer placeholder.
    Footer,
    /// `GenericDateMCAtom`: a generic date placeholder without a format.
    GenericDate,
    /// `DateTimeMCAtom`: a date/time placeholder with a format identifier.
    DateTime,
    /// `RTFDateTimeMCAtom`: a date/time placeholder with an RTF format string.
    RtfDateTime,
}

/// One metacharacter placeholder in a text body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointTextMetachar {
    /// Position of the metacharacter in the corresponding text.
    position: u32,
    kind: PowerPointMetacharKind,
    /// `DateTimeMCAtom` format identifier, when present.
    datetime_format: Option<PowerPointDateTimeFormatId>,
    /// `RTFDateTimeMCAtom` format string, when present.
    rtf_format: Option<String>,
}

impl PowerPointTextMetachar {
    /// Position of the metacharacter in the corresponding text.
    pub const fn position(&self) -> u32 {
        self.position
    }

    /// The placeholder kind.
    pub const fn kind(&self) -> PowerPointMetacharKind {
        self.kind
    }

    /// `DateTimeMCAtom` format identifier, when this is a date/time atom.
    pub const fn datetime_format(&self) -> Option<PowerPointDateTimeFormatId> {
        self.datetime_format
    }

    /// `RTFDateTimeMCAtom` format string, when this is an RTF date/time atom.
    pub fn rtf_format(&self) -> Option<&str> {
        self.rtf_format.as_deref()
    }
}

fn read_position(data: &[u8], record_type: PptRecordType) -> Result<u32> {
    if data.len() < POSITION_ATOM_LEN {
        return Err(corrupted(format!(
            "{record_type:?} is truncated below its position field"
        )));
    }
    Ok(u32::from_le_bytes(
        data[..4].try_into().expect("length checked"),
    ))
}

/// Parse the metacharacter atoms of one text body, in record order.
pub(crate) fn metachars_from_records<'a>(
    records: impl IntoIterator<Item = &'a PptRecord>,
) -> Result<Vec<PowerPointTextMetachar>> {
    let mut result = Vec::new();
    for record in records {
        if record.version != 0 {
            if matches!(
                record.record_type,
                PptRecordType::SlideNumberMCAtom
                    | PptRecordType::GenericDateMCAtom
                    | PptRecordType::HeaderMCAtom
                    | PptRecordType::FooterMCAtom
                    | PptRecordType::DateTimeMCAtom
                    | PptRecordType::RtfDateTimeMCAtom
            ) {
                return Err(corrupted(format!(
                    "{:?} has a nonzero record version",
                    record.record_type
                )));
            }
            continue;
        }
        let metachar = match record.record_type {
            PptRecordType::SlideNumberMCAtom
            | PptRecordType::GenericDateMCAtom
            | PptRecordType::HeaderMCAtom
            | PptRecordType::FooterMCAtom => {
                if record.data.len() != POSITION_ATOM_LEN {
                    return Err(corrupted(format!(
                        "{:?} must contain exactly a position field",
                        record.record_type
                    )));
                }
                let kind = match record.record_type {
                    PptRecordType::SlideNumberMCAtom => PowerPointMetacharKind::SlideNumber,
                    PptRecordType::GenericDateMCAtom => PowerPointMetacharKind::GenericDate,
                    PptRecordType::HeaderMCAtom => PowerPointMetacharKind::Header,
                    _ => PowerPointMetacharKind::Footer,
                };
                PowerPointTextMetachar {
                    position: read_position(&record.data, record.record_type)?,
                    kind,
                    datetime_format: None,
                    rtf_format: None,
                }
            },
            PptRecordType::DateTimeMCAtom => {
                if record.data.len() != DATE_TIME_ATOM_LEN {
                    return Err(corrupted("DateTimeMCAtom must contain 8 bytes"));
                }
                let index = record.data[4];
                if index > MAX_METACHAR_FORMAT_ID {
                    return Err(corrupted("DateTimeMCAtom format ID is outside 0..=12"));
                }
                PowerPointTextMetachar {
                    position: read_position(&record.data, record.record_type)?,
                    kind: PowerPointMetacharKind::DateTime,
                    datetime_format: Some(PowerPointDateTimeFormatId::new(index)?),
                    rtf_format: None,
                }
            },
            PptRecordType::RtfDateTimeMCAtom => {
                if record.data.len() != RTF_DATE_TIME_ATOM_LEN {
                    return Err(corrupted("RTFDateTimeMCAtom must contain 0x84 bytes"));
                }
                let format_bytes = &record.data[4..4 + RTF_FORMAT_LEN];
                let end = format_bytes
                    .iter()
                    .position(|byte| *byte == 0)
                    .unwrap_or(RTF_FORMAT_LEN);
                PowerPointTextMetachar {
                    position: read_position(&record.data, record.record_type)?,
                    kind: PowerPointMetacharKind::RtfDateTime,
                    datetime_format: None,
                    rtf_format: Some(
                        format_bytes[..end]
                            .iter()
                            .map(|byte| char::from(*byte))
                            .collect(),
                    ),
                }
            },
            _ => continue,
        };
        result.push(metachar);
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::consts::PptRecordType;

    fn atom(record_type: PptRecordType, data: &[u8]) -> PptRecord {
        PptRecord {
            version: 0,
            instance: 0,
            record_type,
            record_type_raw: record_type as u16,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    #[test]
    fn parses_position_only_metachars() {
        let records = [
            atom(PptRecordType::SlideNumberMCAtom, &7u32.to_le_bytes()),
            atom(PptRecordType::FooterMCAtom, &3u32.to_le_bytes()),
        ];
        let metachars = metachars_from_records(records.iter()).unwrap();
        assert_eq!(metachars.len(), 2);
        assert_eq!(metachars[0].kind(), PowerPointMetacharKind::SlideNumber);
        assert_eq!(metachars[0].position(), 7);
        assert_eq!(metachars[1].kind(), PowerPointMetacharKind::Footer);
        assert_eq!(metachars[1].position(), 3);
    }

    #[test]
    fn parses_datetime_and_rtf_metachars() {
        let mut date_time = Vec::new();
        date_time.extend_from_slice(&11u32.to_le_bytes());
        date_time.push(5);
        date_time.extend_from_slice(&[0; 3]);
        let mut rtf = Vec::new();
        rtf.extend_from_slice(&2u32.to_le_bytes());
        let mut format = [0u8; RTF_FORMAT_LEN];
        format[..8].copy_from_slice(b"MM/dd/yy");
        rtf.extend_from_slice(&format);
        let records = [
            atom(PptRecordType::DateTimeMCAtom, &date_time),
            atom(PptRecordType::RtfDateTimeMCAtom, &rtf),
        ];
        let metachars = metachars_from_records(records.iter()).unwrap();
        assert_eq!(metachars[0].kind(), PowerPointMetacharKind::DateTime);
        assert_eq!(metachars[0].datetime_format().unwrap().get(), 5);
        assert_eq!(metachars[1].kind(), PowerPointMetacharKind::RtfDateTime);
        assert_eq!(metachars[1].rtf_format(), Some("MM/dd/yy"));
    }

    #[test]
    fn rejects_malformed_metachars() {
        // Truncated position atom.
        let short = [atom(PptRecordType::FooterMCAtom, &[1, 2])];
        assert!(metachars_from_records(short.iter()).is_err());
        // Oversized position atom.
        let long = [atom(PptRecordType::HeaderMCAtom, &[0; 8])];
        assert!(metachars_from_records(long.iter()).is_err());
        // Format ID above 12.
        let mut bad_format = Vec::new();
        bad_format.extend_from_slice(&0u32.to_le_bytes());
        bad_format.push(13);
        bad_format.extend_from_slice(&[0; 3]);
        let bad = [atom(PptRecordType::DateTimeMCAtom, &bad_format)];
        assert!(metachars_from_records(bad.iter()).is_err());
        // Truncated RTF atom.
        let short_rtf = [atom(PptRecordType::RtfDateTimeMCAtom, &[0; 20])];
        assert!(metachars_from_records(short_rtf.iter()).is_err());
    }
}

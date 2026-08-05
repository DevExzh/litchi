//! Bounded binary codecs for PowerPoint header/footer records.

use super::model::{
    PowerPointDateTimeFormatId, PowerPointHeaderFooter, PowerPointHeaderFooterOptions,
    PowerPointHeaderFooterScope,
};
use crate::consts::PptRecordType;
use crate::package::{PptError, Result};
use crate::records::PptRecord;

pub(crate) const HEADERS_FOOTERS_RECORD_TYPE: u16 = 0x0FD9;
pub(crate) const HEADERS_FOOTERS_ATOM_RECORD_TYPE: u16 = 0x0FDA;
pub(crate) const CSTRING_RECORD_TYPE: u16 = 0x0FBA;
pub(crate) const CONTAINER_VERSION: u16 = 0x000F;
pub(crate) const ATOM_VERSION: u16 = 0;
pub(crate) const PRESENTATION_SLIDES_INSTANCE: u16 = 3;
pub(crate) const NOTES_AND_HANDOUTS_INSTANCE: u16 = 4;
pub(crate) const LOCAL_INSTANCE: u16 = 0;
pub(crate) const USER_DATE_INSTANCE: u16 = 0;
pub(crate) const HEADER_INSTANCE: u16 = 1;
pub(crate) const FOOTER_INSTANCE: u16 = 2;
pub(crate) const HEADERS_FOOTERS_ATOM_LENGTH: usize = 4;
pub(crate) const USER_DATE_MAX_BYTES: usize = 510;
pub(crate) const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_AGGREGATE_TEXT_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_HEADER_FOOTER_ENTRIES: usize = 65_536;
pub(crate) const MAX_SCANNED_RECORDS: usize = 1_000_000;
const KNOWN_FLAG_MASK: u16 = 0x003F;

impl PowerPointHeaderFooterScope {
    pub(crate) fn record_instance(self) -> u16 {
        match self {
            Self::PresentationSlides => PRESENTATION_SLIDES_INSTANCE,
            Self::NotesAndHandouts => NOTES_AND_HANDOUTS_INSTANCE,
            Self::Local { .. } => LOCAL_INSTANCE,
        }
    }

    pub(crate) fn permits_header_atom(self) -> bool {
        matches!(self, Self::NotesAndHandouts)
    }
}

impl PowerPointHeaderFooterOptions {
    fn from_atom(record: &PptRecord) -> Result<Self> {
        validate_record_header(
            record,
            PptRecordType::HeadersFootersAtom,
            HEADERS_FOOTERS_ATOM_RECORD_TYPE,
            ATOM_VERSION,
            0,
        )?;
        if record.data_length as usize != HEADERS_FOOTERS_ATOM_LENGTH
            || record.data.len() != HEADERS_FOOTERS_ATOM_LENGTH
            || !record.children.is_empty()
        {
            return Err(corrupted(
                "HeadersFootersAtom must have exactly four data bytes",
            ));
        }
        let format_id = i16::from_le_bytes([record.data[0], record.data[1]]);
        if !(0..=i16::from(PowerPointDateTimeFormatId::MAX)).contains(&format_id) {
            return Err(corrupted(
                "header/footer datetime format ID is outside 0..=13",
            ));
        }
        let mask = u16::from_le_bytes([record.data[2], record.data[3]]);
        if mask & !KNOWN_FLAG_MASK != 0 {
            return Err(corrupted(
                "HeadersFootersAtom has nonzero reserved flag bits",
            ));
        }
        Ok(Self {
            datetime_format: PowerPointDateTimeFormatId::new(format_id as u8)?,
            show_date: mask & 0x0001 != 0,
            use_current_datetime: mask & 0x0002 != 0,
            use_user_date: mask & 0x0004 != 0,
            show_slide_number: mask & 0x0008 != 0,
            show_header: mask & 0x0010 != 0,
            show_footer: mask & 0x0020 != 0,
        })
    }

    fn mask(self) -> u16 {
        u16::from(self.show_date)
            | (u16::from(self.use_current_datetime) << 1)
            | (u16::from(self.use_user_date) << 2)
            | (u16::from(self.show_slide_number) << 3)
            | (u16::from(self.show_header) << 4)
            | (u16::from(self.show_footer) << 5)
    }
}

impl PowerPointHeaderFooter {
    /// Strictly parse one already-materialized `RT_HeadersFooters` record.
    ///
    /// The supplied scope is checked against the record instance. Direct-parent
    /// placement is validated by [`PowerPointHeaderFooters`] when parsing a
    /// complete presentation.
    pub fn parse_record(record: &PptRecord, scope: PowerPointHeaderFooterScope) -> Result<Self> {
        let mut aggregate = 0usize;
        Self::parse_record_bounded(record, scope, &mut aggregate)
    }

    pub(crate) fn parse_record_bounded(
        record: &PptRecord,
        scope: PowerPointHeaderFooterScope,
        aggregate: &mut usize,
    ) -> Result<Self> {
        validate_record_header(
            record,
            PptRecordType::HeadersFooters,
            HEADERS_FOOTERS_RECORD_TYPE,
            CONTAINER_VERSION,
            scope.record_instance(),
        )?;
        if record.data_length as usize != record.data.len() {
            return Err(corrupted("HeadersFooters container payload is truncated"));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "HeadersFooters")?;
        let Some(atom) = children.first() else {
            return Err(corrupted(
                "HeadersFooters container is missing HeadersFootersAtom",
            ));
        };
        let options = PowerPointHeaderFooterOptions::from_atom(atom)?;

        let mut user_date = None;
        let mut header = None;
        let mut footer = None;
        let mut previous_instance = None;
        for child in &children[1..] {
            if child.record_type != PptRecordType::CString
                || child.record_type_raw != CSTRING_RECORD_TYPE
                || child.version != ATOM_VERSION
            {
                return Err(corrupted(
                    "HeadersFooters contains an unexpected child record",
                ));
            }
            if child.data_length as usize != child.data.len() || child.data.len() % 2 != 0 {
                return Err(corrupted(
                    "header/footer CString has an invalid byte length",
                ));
            }
            if child.data.len() > MAX_TEXT_BYTES {
                return Err(corrupted(
                    "header/footer CString exceeds the resource limit",
                ));
            }
            if previous_instance.is_some_and(|previous| child.instance <= previous) {
                return Err(corrupted(
                    "header/footer CString children are duplicated or out of order",
                ));
            }
            previous_instance = Some(child.instance);
            *aggregate = aggregate
                .checked_add(child.data.len())
                .ok_or_else(|| corrupted("header/footer aggregate size overflow"))?;
            if *aggregate > MAX_AGGREGATE_TEXT_BYTES {
                return Err(corrupted(
                    "header/footer strings exceed the aggregate resource limit",
                ));
            }
            let value = decode_printable_unicode(&child.data)?;
            match child.instance {
                USER_DATE_INSTANCE => {
                    if child.data.len() > USER_DATE_MAX_BYTES {
                        return Err(corrupted("UserDateAtom exceeds 510 bytes"));
                    }
                    user_date = Some(value);
                },
                HEADER_INSTANCE if scope.permits_header_atom() => header = Some(value),
                HEADER_INSTANCE => {
                    return Err(corrupted(
                        "HeaderAtom is not permitted in this header/footer scope",
                    ));
                },
                FOOTER_INSTANCE => footer = Some(value),
                _ => return Err(corrupted("header/footer CString has an invalid instance")),
            }
        }

        Ok(Self {
            scope,
            options,
            user_date,
            header,
            footer,
            placeholder_display: None,
        })
    }

    /// Serialize this metadata as one canonical `RT_HeadersFooters` record.
    ///
    /// Serialization is record-local and deterministic. It does not evaluate a
    /// date or modify an OLE persistence directory.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate_for_write()?;
        let mut body = Vec::new();

        let atom_data = [
            self.options.datetime_format.get(),
            0,
            self.options.mask() as u8,
            (self.options.mask() >> 8) as u8,
        ];
        append_record(
            &mut body,
            ATOM_VERSION,
            0,
            HEADERS_FOOTERS_ATOM_RECORD_TYPE,
            &atom_data,
        )?;
        if let Some(value) = &self.user_date {
            append_cstring(&mut body, USER_DATE_INSTANCE, value)?;
        }
        if let Some(value) = &self.header {
            append_cstring(&mut body, HEADER_INSTANCE, value)?;
        }
        if let Some(value) = &self.footer {
            append_cstring(&mut body, FOOTER_INSTANCE, value)?;
        }

        let mut output = Vec::with_capacity(body.len().saturating_add(8));
        append_record(
            &mut output,
            CONTAINER_VERSION,
            self.scope.record_instance(),
            HEADERS_FOOTERS_RECORD_TYPE,
            &body,
        )?;
        Ok(output)
    }

    fn validate_for_write(&self) -> Result<()> {
        if self.header.is_some() && !self.scope.permits_header_atom() {
            return Err(corrupted(
                "HeaderAtom is not permitted in this header/footer scope",
            ));
        }
        let mut aggregate = 0usize;
        for (kind, value) in [
            (USER_DATE_INSTANCE, self.user_date.as_deref()),
            (HEADER_INSTANCE, self.header.as_deref()),
            (FOOTER_INSTANCE, self.footer.as_deref()),
        ] {
            let Some(value) = value else { continue };
            let bytes = validated_encoded_len(value)?;
            if kind == USER_DATE_INSTANCE && bytes > USER_DATE_MAX_BYTES {
                return Err(corrupted("UserDateAtom exceeds 510 bytes"));
            }
            aggregate = aggregate
                .checked_add(bytes)
                .ok_or_else(|| corrupted("header/footer aggregate size overflow"))?;
        }
        if aggregate > MAX_AGGREGATE_TEXT_BYTES {
            return Err(corrupted(
                "header/footer strings exceed the aggregate resource limit",
            ));
        }
        Ok(())
    }
}

pub(crate) fn validate_record_header(
    record: &PptRecord,
    expected_type: PptRecordType,
    expected_raw_type: u16,
    expected_version: u16,
    expected_instance: u16,
) -> Result<()> {
    if record.record_type != expected_type
        || record.record_type_raw != expected_raw_type
        || record.version != expected_version
        || record.instance != expected_instance
    {
        return Err(corrupted(
            "header/footer record header does not match [MS-PPT]",
        ));
    }
    Ok(())
}

pub(crate) fn decode_printable_unicode(data: &[u8]) -> Result<String> {
    let mut units = Vec::with_capacity(data.len() / 2);
    let mut terminated = false;
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if terminated {
            if unit != 0 {
                return Err(corrupted(
                    "PrintableUnicodeString has data after its terminator",
                ));
            }
            continue;
        }
        if unit == 0 {
            terminated = true;
            continue;
        }
        if is_forbidden_printable_unit(unit) {
            return Err(corrupted(
                "PrintableUnicodeString contains a forbidden control character",
            ));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| corrupted("PrintableUnicodeString contains invalid UTF-16"))
}

fn is_forbidden_printable_unit(unit: u16) -> bool {
    matches!(unit, 0x0000..=0x001F | 0x007F..=0x009F)
}

pub(crate) fn validated_encoded_len(value: &str) -> Result<usize> {
    let mut units = 0usize;
    for unit in value.encode_utf16() {
        if is_forbidden_printable_unit(unit) {
            return Err(corrupted(
                "PrintableUnicodeString contains a forbidden control character",
            ));
        }
        units = units
            .checked_add(1)
            .ok_or_else(|| corrupted("header/footer string length overflow"))?;
    }
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| corrupted("header/footer string length overflow"))?;
    if bytes > MAX_TEXT_BYTES {
        return Err(corrupted(
            "header/footer CString exceeds the resource limit",
        ));
    }
    Ok(bytes)
}

fn append_cstring(output: &mut Vec<u8>, instance: u16, value: &str) -> Result<()> {
    let encoded_len = validated_encoded_len(value)?;
    let mut data = Vec::with_capacity(encoded_len);
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    append_record(output, ATOM_VERSION, instance, CSTRING_RECORD_TYPE, &data)
}

fn append_record(
    output: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<()> {
    if version > 0x000F || instance > 0x0FFF {
        return Err(corrupted("PowerPoint record header field overflow"));
    }
    let length = u32::try_from(data.len())
        .map_err(|_| corrupted("PowerPoint record payload exceeds u32"))?;
    output
        .try_reserve(8usize.saturating_add(data.len()))
        .map_err(|_| corrupted("unable to reserve header/footer record memory"))?;
    let version_instance = version | (instance << 4);
    output.extend_from_slice(&version_instance.to_le_bytes());
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(data);
    Ok(())
}

pub(crate) fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

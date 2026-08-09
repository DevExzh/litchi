//! Inert `PowerPoint` 9 Web-publishing metadata from MS-PPT 2.4.18.

use crate::consts::RecordType;

use super::package::{Error, Result};
use super::records::Record;

const HTML_DOC_INFO_RECORD_TYPE: u16 = 0x177b;
const HTML_PUBLISH_INFO_RECORD_TYPE: u16 = 0x177c;
const HTML_PUBLISH_CONTAINER_RECORD_TYPE: u16 = 0x177d;
const C_STRING_RECORD_TYPE: u16 = 0x0fba;
const MAX_FILE_NAME_BYTES: usize = 510;
const MAX_NAMED_SHOW_BYTES: usize = 62;

/// Registered code-page identifier used for Web output.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CodePage(u32);

impl CodePage {
    #[must_use]
    pub const fn new(id: u32) -> Self {
        Self(id)
    }

    #[must_use]
    pub const fn id(self) -> u32 {
        self.0
    }
}

/// Text/background colors used in generated Web frames.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum WebFrameColors {
    Browser = 0,
    PresentationText = 1,
    PresentationAccent = 2,
    WhiteTextOnBlack = 3,
    BlackTextOnWhite = 4,
}

/// Target monitor resolution for generated Web pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum WebScreenSize {
    Pixels544x376 = 0,
    Pixels640x480 = 1,
    Pixels720x512 = 2,
    Pixels800x600 = 3,
    Pixels1024x768 = 4,
    Pixels1152x882 = 5,
    Pixels1152x900 = 6,
    Pixels1280x1024 = 7,
    Pixels1600x1200 = 8,
    Pixels1800x1440 = 9,
    Pixels1920x1200 = 10,
}

/// Browser technology target for Web publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum WebOutput {
    Html3 = 1,
    Html4 = 2,
    Dual = 4,
}

/// Strictly typed `HTMLDocInfo9Atom` settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "fields mirror the independent `HTMLDocInfo9Atom` flag bits one-to-one"
)]
pub struct HtmlDocumentSettings {
    pub encoding: CodePage,
    pub frame_colors: WebFrameColors,
    pub screen_size: WebScreenSize,
    pub output: WebOutput,
    pub show_frame: bool,
    pub resize_graphics: bool,
    pub organize_in_folder: bool,
    pub use_long_file_names: bool,
    pub rely_on_vml: bool,
    pub allow_png: bool,
    pub show_slide_animation: bool,
}

impl HtmlDocumentSettings {
    /// Parse one strict `HTMLDocInfo9Atom`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or size is invalid, an enum field
    /// is out of range, or a reserved flag bit is set.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type_raw != HTML_DOC_INFO_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 16
        {
            return Err(Error::Corrupted(
                "HTMLDocInfo9Atom has an invalid record header or size".to_string(),
            ));
        }
        let encoding = CodePage::new(u32::from_le_bytes([
            record.data[4],
            record.data[5],
            record.data[6],
            record.data[7],
        ]));
        let frame_colors = match u16::from_le_bytes([record.data[8], record.data[9]]) {
            0 => WebFrameColors::Browser,
            1 => WebFrameColors::PresentationText,
            2 => WebFrameColors::PresentationAccent,
            3 => WebFrameColors::WhiteTextOnBlack,
            4 => WebFrameColors::BlackTextOnWhite,
            _ => {
                return Err(Error::Corrupted(
                    "HTMLDocInfo9Atom has an invalid frame-color mode".to_string(),
                ));
            },
        };
        let screen_size = match record.data[10] {
            0 => WebScreenSize::Pixels544x376,
            1 => WebScreenSize::Pixels640x480,
            2 => WebScreenSize::Pixels720x512,
            3 => WebScreenSize::Pixels800x600,
            4 => WebScreenSize::Pixels1024x768,
            5 => WebScreenSize::Pixels1152x882,
            6 => WebScreenSize::Pixels1152x900,
            7 => WebScreenSize::Pixels1280x1024,
            8 => WebScreenSize::Pixels1600x1200,
            9 => WebScreenSize::Pixels1800x1440,
            10 => WebScreenSize::Pixels1920x1200,
            _ => {
                return Err(Error::Corrupted(
                    "HTMLDocInfo9Atom has an invalid screen size".to_string(),
                ));
            },
        };
        let output = match record.data[12] {
            1 => WebOutput::Html3,
            2 => WebOutput::Html4,
            4 => WebOutput::Dual,
            _ => {
                return Err(Error::Corrupted(
                    "HTMLDocInfo9Atom has an invalid Web output target".to_string(),
                ));
            },
        };
        let flags = record.data[13];
        if flags & 0x80 != 0 {
            return Err(Error::Corrupted(
                "HTMLDocInfo9Atom has a nonzero reserved flag".to_string(),
            ));
        }
        Ok(Self {
            encoding,
            frame_colors,
            screen_size,
            output,
            show_frame: flags & 0x01 != 0,
            resize_graphics: flags & 0x02 != 0,
            organize_in_folder: flags & 0x04 != 0,
            use_long_file_names: flags & 0x08 != 0,
            rely_on_vml: flags & 0x10 != 0,
            allow_png: flags & 0x20 != 0,
            show_slide_animation: flags & 0x40 != 0,
        })
    }

    /// Discover the single Web-document atom in the PPT9 document tag.
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == HTML_DOC_INFO_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT9 document tag contains multiple HTMLDocInfo9Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Encode a canonical atom with undefined fields zeroed.
    #[must_use]
    pub fn to_record(self) -> Record {
        let mut data = vec![0; 16];
        data[4..8].copy_from_slice(&self.encoding.id().to_le_bytes());
        data[8..10].copy_from_slice(&(self.frame_colors as u16).to_le_bytes());
        data[10] = self.screen_size as u8;
        data[12] = self.output as u8;
        data[13] = u8::from(self.show_frame)
            | u8::from(self.resize_graphics) << 1
            | u8::from(self.organize_in_folder) << 2
            | u8::from(self.use_long_file_names) << 3
            | u8::from(self.rely_on_vml) << 4
            | u8::from(self.allow_png) << 5
            | u8::from(self.show_slide_animation) << 6;
        Record {
            record_type: RecordType::from(HTML_DOC_INFO_RECORD_TYPE),
            record_type_raw: HTML_DOC_INFO_RECORD_TYPE,
            version: 0,
            instance: 0,
            data_length: 16,
            data,
            children: Vec::new(),
        }
    }
}

/// Strictly typed `HTMLPublishInfo9Container` metadata.
///
/// The range fields are always preserved because the binary atom always stores
/// them, even when [`Self::use_slide_range`] is false. `load_in_browser` is
/// exposed as inert intent only; this library never launches a browser.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`HtmlPublishSettings` is the established public API name; renaming it would break downstream crates"
)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "fields mirror the independent `HTMLPublishInfoAtom` flag bits one-to-one"
)]
pub struct HtmlPublishSettings {
    pub file_name: String,
    pub named_show: Option<String>,
    pub start_slide: u32,
    pub end_slide: u32,
    pub output: WebOutput,
    pub use_slide_range: bool,
    pub use_named_show: bool,
    pub load_in_browser: bool,
    pub show_speaker_notes: bool,
}

impl HtmlPublishSettings {
    /// Validate an in-memory publication description without performing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn validate(&self) -> Result<()> {
        validate_printable_unicode(
            &self.file_name,
            MAX_FILE_NAME_BYTES,
            "HTML publication file name",
        )?;
        if let Some(named_show) = &self.named_show {
            validate_printable_unicode(
                named_show,
                MAX_NAMED_SHOW_BYTES,
                "HTML publication named show",
            )?;
        }
        if self.start_slide > i32::MAX as u32 || self.end_slide > i32::MAX as u32 {
            return Err(Error::Corrupted(
                "HTML publication slide index exceeds the signed 32-bit range".to_string(),
            ));
        }
        if self.use_named_show && self.named_show.is_none() {
            return Err(Error::Corrupted(
                "HTML publication selects a named show but has no NamedShowAtom".to_string(),
            ));
        }
        Ok(())
    }

    /// Parse one exact `HTMLPublishInfo9Container`.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, child grammar, or any atom
    /// field is malformed or out of range.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type_raw != HTML_PUBLISH_CONTAINER_RECORD_TYPE
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "HTMLPublishInfo9Container has an invalid record header".to_string(),
            ));
        }
        let children = Record::parse_sequence_strict(&record.data, "HTMLPublishInfo9Container")?;
        if !(2..=3).contains(&children.len()) {
            return Err(Error::Corrupted(
                "HTMLPublishInfo9Container must contain a file name, optional named show, and publish atom"
                    .to_string(),
            ));
        }

        let file_name =
            parse_printable_unicode(&children[0], 0, MAX_FILE_NAME_BYTES, "FileNameAtom")?;
        let (named_show, info) = if children.len() == 3 {
            (
                Some(parse_printable_unicode(
                    &children[1],
                    1,
                    MAX_NAMED_SHOW_BYTES,
                    "NamedShowAtom",
                )?),
                &children[2],
            )
        } else {
            (None, &children[1])
        };
        if info.record_type_raw != HTML_PUBLISH_INFO_RECORD_TYPE
            || info.version != 0
            || info.instance != 0
            || info.data.len() != 12
        {
            return Err(Error::Corrupted(
                "HTMLPublishInfoAtom has an invalid record header or size".to_string(),
            ));
        }

        let start_slide_raw =
            i32::from_le_bytes([info.data[0], info.data[1], info.data[2], info.data[3]]);
        let end_slide_raw =
            i32::from_le_bytes([info.data[4], info.data[5], info.data[6], info.data[7]]);
        let Ok(start_slide) = u32::try_from(start_slide_raw) else {
            return Err(Error::Corrupted(
                "HTMLPublishInfoAtom contains a negative slide index".to_string(),
            ));
        };
        let Ok(end_slide) = u32::try_from(end_slide_raw) else {
            return Err(Error::Corrupted(
                "HTMLPublishInfoAtom contains a negative slide index".to_string(),
            ));
        };
        let output = parse_web_output(info.data[8], "HTMLPublishInfoAtom")?;
        let flags = info.data[9];
        if flags & 0xf0 != 0 {
            return Err(Error::Corrupted(
                "HTMLPublishInfoAtom has nonzero reserved flags".to_string(),
            ));
        }
        let value = Self {
            file_name,
            named_show,
            start_slide,
            end_slide,
            output,
            use_slide_range: flags & 0x01 != 0,
            use_named_show: flags & 0x02 != 0,
            load_in_browser: flags & 0x04 != 0,
            show_speaker_notes: flags & 0x08 != 0,
        };
        value.validate()?;
        Ok(value)
    }

    /// Discover and cross-validate the single PPT9 publication container.
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == HTML_PUBLISH_CONTAINER_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT9 document tag contains multiple HTMLPublishInfo9Container records".to_string(),
            ));
        }
        let value = Self::parse(record)?;
        if let Some(name) = &value.named_show {
            let named_shows =
                super::named_shows::NamedShows::parse(document)?.ok_or_else(|| {
                    Error::Corrupted(
                        "HTML publication names a show but the document has no named shows"
                            .to_string(),
                    )
                })?;
            if !named_shows.shows.iter().any(|show| show.name == *name) {
                return Err(Error::Corrupted(
                    "HTML publication NamedShowAtom does not match a document named show"
                        .to_string(),
                ));
            }
        }
        Ok(Some(value))
    }

    /// Encode a canonical `HTMLPublishInfo9Container` with undefined bytes zeroed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record(&self) -> Result<Record> {
        self.validate()?;
        let file_name = encode_printable_unicode(&self.file_name);
        let mut data = record_bytes(0, 0, C_STRING_RECORD_TYPE, &file_name)?;
        if let Some(named_show) = &self.named_show {
            let encoded_show = encode_printable_unicode(named_show);
            data.extend_from_slice(&record_bytes(0, 1, C_STRING_RECORD_TYPE, &encoded_show)?);
        }

        let start_slide = i32::try_from(self.start_slide).map_err(|_err| {
            Error::Corrupted(
                "HTML publication slide index exceeds the signed 32-bit range".to_string(),
            )
        })?;
        let end_slide = i32::try_from(self.end_slide).map_err(|_err| {
            Error::Corrupted(
                "HTML publication slide index exceeds the signed 32-bit range".to_string(),
            )
        })?;
        let mut info = [0u8; 12];
        info[0..4].copy_from_slice(&start_slide.to_le_bytes());
        info[4..8].copy_from_slice(&end_slide.to_le_bytes());
        info[8] = self.output as u8;
        info[9] = u8::from(self.use_slide_range)
            | u8::from(self.use_named_show) << 1
            | u8::from(self.load_in_browser) << 2
            | u8::from(self.show_speaker_notes) << 3;
        data.extend_from_slice(&record_bytes(0, 0, HTML_PUBLISH_INFO_RECORD_TYPE, &info)?);
        let data_length = u32::try_from(data.len()).map_err(|_err| {
            Error::Corrupted("HTML publication container length overflow".to_string())
        })?;
        Ok(Record {
            record_type: RecordType::from(HTML_PUBLISH_CONTAINER_RECORD_TYPE),
            record_type_raw: HTML_PUBLISH_CONTAINER_RECORD_TYPE,
            version: 0x0f,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

fn parse_web_output(value: u8, record_name: &str) -> Result<WebOutput> {
    match value {
        1 => Ok(WebOutput::Html3),
        2 => Ok(WebOutput::Html4),
        4 => Ok(WebOutput::Dual),
        _ => Err(Error::Corrupted(format!(
            "{record_name} has an invalid Web output target"
        ))),
    }
}

fn parse_printable_unicode(
    record: &Record,
    instance: u16,
    max_bytes: usize,
    record_name: &str,
) -> Result<String> {
    if record.record_type_raw != C_STRING_RECORD_TYPE
        || record.version != 0
        || record.instance != instance
        || record.data.len() > max_bytes
        || !record.data.len().is_multiple_of(2)
    {
        return Err(Error::Corrupted(format!(
            "{record_name} has an invalid record header or size"
        )));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if is_forbidden_printable_unit(unit) {
            return Err(Error::Corrupted(format!(
                "{record_name} contains a forbidden control character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_err| Error::Corrupted(format!("{record_name} contains invalid UTF-16")))
}

fn validate_printable_unicode(value: &str, max_bytes: usize, field_name: &str) -> Result<()> {
    let mut bytes = 0usize;
    for unit in value.encode_utf16() {
        if is_forbidden_printable_unit(unit) {
            return Err(Error::Corrupted(format!(
                "{field_name} contains a forbidden control character"
            )));
        }
        bytes = bytes
            .checked_add(2)
            .ok_or_else(|| Error::Corrupted(format!("{field_name} length overflow")))?;
        if bytes > max_bytes {
            return Err(Error::Corrupted(format!(
                "{field_name} exceeds its MS-PPT byte limit"
            )));
        }
    }
    Ok(())
}

fn encode_printable_unicode(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

const fn is_forbidden_printable_unit(unit: u16) -> bool {
    matches!(unit, 0x0000..=0x001f | 0x007f..=0x009f)
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let data_length = u32::try_from(data.len()).map_err(|_err| {
        Error::Corrupted("PowerPoint Web-publishing record length overflow".to_string())
    })?;
    let mut bytes = Vec::with_capacity(8 + data.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&data_length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn settings() -> HtmlDocumentSettings {
        HtmlDocumentSettings {
            encoding: CodePage::new(65001),
            frame_colors: WebFrameColors::PresentationAccent,
            screen_size: WebScreenSize::Pixels1920x1200,
            output: WebOutput::Dual,
            show_frame: true,
            resize_graphics: true,
            organize_in_folder: true,
            use_long_file_names: true,
            rely_on_vml: false,
            allow_png: true,
            show_slide_animation: true,
        }
    }

    #[test]
    fn round_trips_typed_web_document_settings() {
        let expected = settings();
        let record = expected.to_record();
        assert_eq!(record.data[0..4], [0; 4]);
        assert_eq!(record.data[11], 0);
        assert_eq!(record.data[14..16], [0; 2]);
        assert_eq!(HtmlDocumentSettings::parse(&record).unwrap(), expected);
    }

    #[test]
    fn ignores_undefined_bytes_but_rejects_enum_and_reserved_values() {
        let mut record = settings().to_record();
        record.data[0..4].copy_from_slice(&u32::MAX.to_le_bytes());
        record.data[11] = 0xff;
        record.data[14..16].copy_from_slice(&u16::MAX.to_le_bytes());
        assert!(HtmlDocumentSettings::parse(&record).is_ok());

        record.data[13] = 0x80;
        assert!(HtmlDocumentSettings::parse(&record).is_err());
        record.data[13] = 0;
        record.data[12] = 3;
        assert!(HtmlDocumentSettings::parse(&record).is_err());
    }

    fn publication() -> HtmlPublishSettings {
        HtmlPublishSettings {
            file_name: "https://example.test/slides.html".to_string(),
            named_show: Some("Executive".to_string()),
            start_slide: 1,
            end_slide: 12,
            output: WebOutput::Dual,
            use_slide_range: true,
            use_named_show: true,
            load_in_browser: true,
            show_speaker_notes: false,
        }
    }

    #[test]
    fn round_trips_exact_html_publication_container() {
        let expected = publication();
        let record = expected.to_record().unwrap();
        let children = Record::parse_sequence_strict(&record.data, "test").unwrap();
        assert_eq!(children.len(), 3);
        assert_eq!(children[0].instance, 0);
        assert_eq!(children[1].instance, 1);
        assert_eq!(children[2].record_type_raw, HTML_PUBLISH_INFO_RECORD_TYPE);
        assert_eq!(children[2].data[10..12], [0, 0]);
        assert_eq!(HtmlPublishSettings::parse(&record).unwrap(), expected);
    }

    #[test]
    fn preserves_optional_named_show_and_undefined_atom_bytes() {
        let mut expected = publication();
        expected.named_show = None;
        expected.use_named_show = false;
        let mut record = expected.to_record().unwrap();
        let info_offset = record.data.len() - 12;
        record.data[info_offset + 10..info_offset + 12].copy_from_slice(&[0xaa, 0x55]);
        let parsed = HtmlPublishSettings::parse(&record).unwrap();
        assert_eq!(parsed, expected);
        let canonical = parsed.to_record().unwrap();
        assert_eq!(canonical.data[canonical.data.len() - 2..], [0, 0]);
    }

    #[test]
    fn rejects_malformed_publish_headers_values_order_and_dependencies() {
        let mut value = publication();
        value.named_show = None;
        assert!(value.validate().is_err());
        value.use_named_show = false;
        value.start_slide = i32::MAX as u32 + 1;
        assert!(value.validate().is_err());
        value.start_slide = 0;
        value.file_name = "x".repeat(256);
        assert!(value.validate().is_err());

        let valid = publication().to_record().unwrap();
        let children = Record::parse_sequence_strict(&valid.data, "test").unwrap();
        let mut wrong_order = record_bytes(
            children[1].version,
            children[1].instance,
            children[1].record_type_raw,
            &children[1].data,
        )
        .unwrap();
        wrong_order.extend_from_slice(
            &record_bytes(
                children[0].version,
                children[0].instance,
                children[0].record_type_raw,
                &children[0].data,
            )
            .unwrap(),
        );
        wrong_order.extend_from_slice(
            &record_bytes(
                children[2].version,
                children[2].instance,
                children[2].record_type_raw,
                &children[2].data,
            )
            .unwrap(),
        );
        let mut record = valid.clone();
        record.data = wrong_order;
        record.data_length = u32::try_from(record.data.len()).unwrap();
        assert!(HtmlPublishSettings::parse(&record).is_err());

        let mut negative = valid.clone();
        let atom_start = negative.data.len() - 12;
        negative.data[atom_start..atom_start + 4].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(HtmlPublishSettings::parse(&negative).is_err());

        let mut reserved = valid;
        let flags_offset = reserved.data.len() - 3;
        reserved.data[flags_offset] |= 0x10;
        assert!(HtmlPublishSettings::parse(&reserved).is_err());
    }
}

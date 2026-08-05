//! Strict, inert PowerPoint document-level slide-show settings.

use super::named_shows::{NamedShow, NamedShows};
use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;
use std::ops::RangeInclusive;

const PAYLOAD_LEN: usize = 80;
const NAMED_SHOW_BYTES: usize = 64;
const KNOWN_FLAGS: u16 = 0x01ff;

/// The meaning of the index byte in an MS-PPT `ColorIndexStruct`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u8)]
pub enum ColorIndexKind {
    Background = 0x00,
    Text = 0x01,
    Shadow = 0x02,
    TitleText = 0x03,
    Fill = 0x04,
    Accent1 = 0x05,
    Accent2 = 0x06,
    Accent3 = 0x07,
    Srgb = 0xfe,
    Undefined = 0xff,
}

impl ColorIndexKind {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(Self::Background),
            0x01 => Ok(Self::Text),
            0x02 => Ok(Self::Shadow),
            0x03 => Ok(Self::TitleText),
            0x04 => Ok(Self::Fill),
            0x05 => Ok(Self::Accent1),
            0x06 => Ok(Self::Accent2),
            0x07 => Ok(Self::Accent3),
            0xfe => Ok(Self::Srgb),
            0xff => Ok(Self::Undefined),
            _ => corrupted(format!("invalid ColorIndexStruct index {value:#04x}")),
        }
    }
}

/// A four-byte MS-PPT `ColorIndexStruct`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ColorIndex {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
    pub kind: ColorIndexKind,
}

impl ColorIndex {
    /// Parse one four-byte `ColorIndexStruct` (MS-PPT 2.12.2).
    pub(crate) fn parse_bytes(data: &[u8]) -> Result<Self> {
        let [red, green, blue, index]: [u8; 4] = data
            .try_into()
            .map_err(|_| Error::Corrupted("ColorIndexStruct is truncated".to_string()))?;
        Ok(Self {
            red,
            green,
            blue,
            kind: ColorIndexKind::parse(index)?,
        })
    }

    /// Serialize as a four-byte `ColorIndexStruct`.
    pub(crate) fn to_bytes(self) -> [u8; 4] {
        [self.red, self.green, self.blue, self.kind as u8]
    }
}

/// The nine defined `SlideShowDocInfoAtom` flags.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SlideShowFlags {
    pub auto_advance: bool,
    pub will_skip_builds: bool,
    pub use_slide_range: bool,
    pub use_named_show: bool,
    pub browse_mode: bool,
    pub kiosk_mode: bool,
    pub will_skip_narration: bool,
    pub loop_continuously: bool,
    pub hide_scroll_bar: bool,
}

impl SlideShowFlags {
    fn parse(bits: u16) -> Result<Self> {
        if bits & !KNOWN_FLAGS != 0 {
            return corrupted("SlideShowDocInfoAtom has nonzero reserved flag bits");
        }
        let flags = Self {
            auto_advance: bits & (1 << 0) != 0,
            will_skip_builds: bits & (1 << 1) != 0,
            use_slide_range: bits & (1 << 2) != 0,
            use_named_show: bits & (1 << 3) != 0,
            browse_mode: bits & (1 << 4) != 0,
            kiosk_mode: bits & (1 << 5) != 0,
            will_skip_narration: bits & (1 << 6) != 0,
            loop_continuously: bits & (1 << 7) != 0,
            hide_scroll_bar: bits & (1 << 8) != 0,
        };
        flags.validate()?;
        Ok(flags)
    }

    fn bits(self) -> Result<u16> {
        self.validate()?;
        Ok(self.auto_advance as u16
            | (self.will_skip_builds as u16) << 1
            | (self.use_slide_range as u16) << 2
            | (self.use_named_show as u16) << 3
            | (self.browse_mode as u16) << 4
            | (self.kiosk_mode as u16) << 5
            | (self.will_skip_narration as u16) << 6
            | (self.loop_continuously as u16) << 7
            | (self.hide_scroll_bar as u16) << 8)
    }

    fn validate(self) -> Result<()> {
        if self.browse_mode && self.kiosk_mode {
            return corrupted("slide show cannot use browse mode and kiosk mode together");
        }
        Ok(())
    }
}

/// A validated document-level `SlideShowDocInfoAtom`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlideShowSettings {
    pub pen_color: ColorIndex,
    pub restart_time_millis: i32,
    pub start_slide: u16,
    pub end_slide: u16,
    pub named_show: String,
    pub flags: SlideShowFlags,
    /// Undefined bytes preserved for a lossless record roundtrip.
    pub unused: [u8; 2],
}

impl SlideShowSettings {
    /// Parse the unique direct settings atom below `document`.
    pub fn parse(document: &Record) -> Result<Option<Self>> {
        let records = document
            .children
            .iter()
            .filter(|record| record.record_type_raw == RecordType::SlideShowDocInfoAtom.as_u16())
            .collect::<Vec<_>>();
        if records.len() > 1 {
            return corrupted("DocumentContainer contains duplicate SlideShowDocInfoAtom records");
        }
        let Some(record) = records.first() else {
            return Ok(None);
        };
        if record.version != 1
            || record.instance != 0
            || record.data.len() != PAYLOAD_LEN
            || record.data_length != PAYLOAD_LEN as u32
        {
            return corrupted("SlideShowDocInfoAtom has an invalid header or size");
        }
        let data = record.data.as_slice();
        let pen_color = ColorIndex {
            red: data[0],
            green: data[1],
            blue: data[2],
            kind: ColorIndexKind::parse(data[3])?,
        };
        let restart_time_millis = i32::from_le_bytes(data[4..8].try_into().expect("fixed slice"));
        let start_slide = parse_nonnegative_i16(&data[8..10], "startSlide")?;
        let end_slide = parse_nonnegative_i16(&data[10..12], "endSlide")?;
        let named_show = parse_char2(&data[12..12 + NAMED_SHOW_BYTES])?;
        let flags = SlideShowFlags::parse(u16::from_le_bytes([data[76], data[77]]))?;
        let settings = Self {
            pen_color,
            restart_time_millis,
            start_slide,
            end_slide,
            named_show,
            flags,
            unused: [data[78], data[79]],
        };
        settings.validate()?;
        Ok(Some(settings))
    }

    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical SlideShowDocInfoAtom did not consume its bytes");
        }
        Ok(record)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.extend_from_slice(&[
            self.pen_color.red,
            self.pen_color.green,
            self.pen_color.blue,
            self.pen_color.kind as u8,
        ]);
        payload.extend_from_slice(&self.restart_time_millis.to_le_bytes());
        payload.extend_from_slice(&(self.start_slide as i16).to_le_bytes());
        payload.extend_from_slice(&(self.end_slide as i16).to_le_bytes());
        payload.extend_from_slice(&encode_char2(&self.named_show)?);
        payload.extend_from_slice(&self.flags.bits()?.to_le_bytes());
        payload.extend_from_slice(&self.unused);
        debug_assert_eq!(payload.len(), PAYLOAD_LEN);

        let mut bytes = Vec::with_capacity(PAYLOAD_LEN + 8);
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&RecordType::SlideShowDocInfoAtom.as_u16().to_le_bytes());
        bytes.extend_from_slice(&(PAYLOAD_LEN as u32).to_le_bytes());
        bytes.extend_from_slice(&payload);
        Ok(bytes)
    }

    /// The selected one-based slide range, if range mode is active.
    pub fn selected_slide_range(&self) -> Option<RangeInclusive<u16>> {
        self.flags
            .use_slide_range
            .then_some(self.start_slide..=self.end_slide)
    }

    /// Resolve the active named show, ignoring it when range mode takes precedence.
    pub fn selected_named_show<'a>(&self, named_shows: &'a NamedShows) -> Option<&'a NamedShow> {
        (!self.flags.use_slide_range && self.flags.use_named_show)
            .then(|| {
                named_shows
                    .shows
                    .iter()
                    .find(|show| show.name == self.named_show)
            })
            .flatten()
    }

    /// Validate the named-show reference in the document-wide context.
    pub fn validate_named_show(&self, named_shows: &NamedShows) -> Result<()> {
        if !self.flags.use_slide_range
            && self.flags.use_named_show
            && self.selected_named_show(named_shows).is_none()
        {
            return corrupted("SlideShowDocInfoAtom references an absent named show");
        }
        Ok(())
    }

    fn validate(&self) -> Result<()> {
        self.flags.validate()?;
        if self.start_slide > i16::MAX as u16 || self.end_slide > i16::MAX as u16 {
            return corrupted("slide-show range endpoint exceeds signed 16-bit range");
        }
        if self.flags.use_slide_range && (self.start_slide == 0 || self.end_slide == 0) {
            return corrupted("active slide-show range endpoints must be nonzero");
        }
        encode_char2(&self.named_show)?;
        Ok(())
    }
}

fn parse_nonnegative_i16(data: &[u8], field: &str) -> Result<u16> {
    let value = i16::from_le_bytes(data.try_into().expect("two-byte slice"));
    u16::try_from(value)
        .map_err(|_| Error::Corrupted(format!("SlideShowDocInfoAtom {field} is negative")))
}

fn parse_char2(data: &[u8]) -> Result<String> {
    let mut units = Vec::with_capacity(data.len() / 2);
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| Error::Corrupted("namedShow contains invalid UTF-16".to_string()))
}

fn encode_char2(value: &str) -> Result<[u8; NAMED_SHOW_BYTES]> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > NAMED_SHOW_BYTES / 2 {
        return corrupted("namedShow exceeds 32 UTF-16 code units");
    }
    if units.contains(&0) {
        return corrupted("namedShow contains an embedded null");
    }
    let mut data = [0; NAMED_SHOW_BYTES];
    for (slot, unit) in data.chunks_exact_mut(2).zip(units) {
        slot.copy_from_slice(&unit.to_le_bytes());
    }
    Ok(data)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root(children: Vec<Record>) -> Record {
        Record {
            version: 0x0f,
            instance: 0,
            record_type: RecordType::Document,
            record_type_raw: RecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    fn settings() -> SlideShowSettings {
        SlideShowSettings {
            pen_color: ColorIndex {
                red: 12,
                green: 34,
                blue: 56,
                kind: ColorIndexKind::Srgb,
            },
            restart_time_millis: 30_000,
            start_slide: 2,
            end_slide: 7,
            named_show: "Executive".into(),
            flags: SlideShowFlags {
                auto_advance: true,
                use_slide_range: true,
                kiosk_mode: true,
                loop_continuously: true,
                ..SlideShowFlags::default()
            },
            unused: [0xaa, 0x55],
        }
    }

    #[test]
    fn protocol_shaped_settings_roundtrip_and_preserve_unused_bytes() {
        let expected = settings();
        let parsed = SlideShowSettings::parse(&root(vec![expected.to_record().unwrap()]))
            .unwrap()
            .unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(parsed.selected_slide_range(), Some(2..=7));
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn resolves_named_show_only_when_range_mode_is_inactive() {
        let shows = NamedShows {
            shows: vec![NamedShow {
                name: "Executive".into(),
                slide_ids: Some(vec![0x100]),
            }],
        };
        let mut value = settings();
        value.flags.use_slide_range = false;
        value.flags.use_named_show = true;
        assert_eq!(value.selected_named_show(&shows).unwrap().name, "Executive");
        value.validate_named_show(&shows).unwrap();
        value.named_show = "Missing".into();
        assert!(value.validate_named_show(&shows).is_err());
        value.flags.use_slide_range = true;
        assert!(value.selected_named_show(&shows).is_none());
        value.validate_named_show(&shows).unwrap();
    }

    #[test]
    fn rejects_invalid_headers_color_flags_ranges_and_names() {
        let valid = settings().to_record().unwrap();
        assert!(SlideShowSettings::parse(&root(vec![valid.clone(), valid])).is_err());
        let mut bytes = settings().to_record_bytes().unwrap();
        for (offset, replacement) in [(0, 0u8), (8 + 3, 0x08), (8 + 77, 0x02)] {
            let mut hostile = bytes.clone();
            hostile[offset] = replacement;
            let record = Record::parse(&hostile, 0).unwrap().0;
            assert!(SlideShowSettings::parse(&root(vec![record])).is_err());
        }
        bytes[8 + 8..8 + 10].copy_from_slice(&(-1i16).to_le_bytes());
        let record = Record::parse(&bytes, 0).unwrap().0;
        assert!(SlideShowSettings::parse(&root(vec![record])).is_err());

        let mut value = settings();
        value.flags.browse_mode = true;
        assert!(value.to_record_bytes().is_err());
        value.flags.browse_mode = false;
        value.start_slide = 0;
        assert!(value.to_record_bytes().is_err());
        value.start_slide = 1;
        value.named_show = "x".repeat(33);
        assert!(value.to_record_bytes().is_err());
    }
}

//! Slide-list text bodies and the outline references that place them in
//! shapes (MS-PPT 2.4.14.3, 2.9.41, 2.9.78; L-PPT-R1, finding F-10).
//!
//! `PowerPoint` stores the text of a slide's title and body placeholders in
//! the presentation's `SlideListWithTextContainer`, not in the placeholder
//! shapes. Each text body there starts at a `TextHeaderAtom` that follows the
//! slide's `SlidePersistAtom`, and the placeholder's `ClientTextbox` holds
//! only an `OutlineTextRefAtom` whose index selects one of those bodies: the
//! zero-based position of its `TextHeaderAtom` among the ones that follow the
//! slide's `SlidePersistAtom`. Apache POI (`HSLFTextParagraph`) resolves the
//! reference by the same ordinal; LibreOffice (`svdfppt.cxx`) parses the
//! selected body as if it were the shape's own `ClientTextbox`.
//!
//! An [`OutlineTextBody`] keeps one body's records framed exactly as a
//! `ClientTextbox` payload, so a shape that references it is read with the
//! same decoders as a shape that holds its text inline.

use crate::consts::RecordType;
use crate::package::{Error, Result, SourceLimits};
use crate::records::Record;
use crate::text_interaction::TextType;

/// The `TextHeaderAtom` payload length (MS-PPT 2.9.41).
const TEXT_HEADER_ATOM_SIZE: usize = 4;

/// One text body of a presentation slide's slide-list text (MS-PPT
/// 2.4.14.3): a `TextHeaderAtom` and the records that follow it, up to the
/// next `TextHeaderAtom` or `SlidePersistAtom`.
///
/// A shape whose `ClientTextbox` holds an `OutlineTextRefAtom` (MS-PPT
/// 2.9.78) shows the body that the reference's index selects; the slide
/// reader resolves the reference, so [`crate::Slide::shapes`] returns such
/// a placeholder with this body's text. [`crate::Slide::outline_text_bodies`]
/// lists a slide's bodies in slide-list order, which is the order an
/// [`crate::OutlineTextRef`] index counts.
///
/// # Examples
///
/// ```no_run
/// use litchi_ppt::Package;
///
/// let mut package = Package::open("deck.ppt")?;
/// let presentation = package.presentation()?;
/// for slide in presentation.slides()? {
///     for (index, body) in slide.outline_text_bodies().iter().enumerate() {
///         println!("{index}: {:?} {:?}", body.text_type(), body.text()?);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct OutlineTextBody {
    /// The body's `TextTypeEnum` (`TextHeaderAtom.textType`).
    text_type: TextType,
    /// `TextHeaderAtom.recInstance`: the body's item in the slide's text
    /// placeholder list.
    text_header_instance: u16,
    /// The body's records, from its `TextHeaderAtom` on, each with its
    /// eight-byte record header.
    records: Box<[u8]>,
    /// Source policy retained by this source-backed projection.
    source_limits: SourceLimits,
}

impl PartialEq for OutlineTextBody {
    fn eq(&self, other: &Self) -> bool {
        self.text_type == other.text_type
            && self.text_header_instance == other.text_header_instance
            && self.records == other.records
    }
}

impl Eq for OutlineTextBody {}

impl OutlineTextBody {
    /// Splits the records that follow one slide's `SlidePersistAtom` into
    /// text bodies, in slide-list order.
    ///
    /// Records before the first `TextHeaderAtom` belong to no body and are
    /// skipped, as POI skips them. Each body's records are framed as a
    /// `ClientTextbox` payload.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Corrupted`] when a `TextHeaderAtom` is malformed or
    /// a record was truncated by the containing stream, and
    /// [`Error::AllocationFailed`] when a body cannot be allocated.
    pub(crate) fn collect(records: &[&Record], source_limits: SourceLimits) -> Result<Vec<Self>> {
        let headers = records
            .iter()
            .filter(|record| record.record_type == RecordType::TextHeaderAtom)
            .count();
        let mut bodies = Vec::new();
        bodies
            .try_reserve_exact(headers)
            .map_err(|_err| Error::AllocationFailed("PPT slide-list text bodies"))?;
        let mut rest = records;
        while let Some(start) = rest
            .iter()
            .position(|record| record.record_type == RecordType::TextHeaderAtom)
        {
            let header = rest[start];
            let body = &rest[start..];
            let end = body[1..]
                .iter()
                .position(|record| record.record_type == RecordType::TextHeaderAtom)
                .map_or(body.len(), |next| next + 1);
            bodies.push(Self {
                text_type: text_header_type(header)?,
                text_header_instance: header.instance,
                records: framed(&body[..end])?,
                source_limits,
            });
            rest = &body[end..];
        }
        Ok(bodies)
    }

    /// The body's text type (`TextHeaderAtom.textType`).
    #[must_use]
    pub const fn text_type(&self) -> TextType {
        self.text_type
    }

    /// The `TextHeaderAtom` record instance, which MS-PPT defines as the
    /// body's item in the slide's text placeholder list.
    #[must_use]
    pub const fn text_header_instance(&self) -> u16 {
        self.text_header_instance
    }

    /// The body's records, starting at its `TextHeaderAtom`, framed as an
    /// `OfficeArtClientTextbox` payload (MS-PPT 2.9.76).
    #[must_use]
    pub fn records(&self) -> &[u8] {
        &self.records
    }

    /// The body's text, decoded exactly as the text of a placeholder that
    /// holds the same records inline: every character atom's text, trimmed
    /// and joined with `\n`, with paragraph breaks kept as `\r`. Empty when
    /// the body has no characters.
    ///
    /// # Errors
    ///
    /// Returns an error when the body's records cannot be traversed.
    pub fn text(&self) -> Result<String> {
        Ok(crate::odraw::text_from_textbox_payload_bounded(
            &self.records,
            self.source_limits
                .max_events()
                .try_into()
                .unwrap_or(u32::MAX),
            self.source_limits.max_depth(),
        )?
        .unwrap_or_default())
    }
}

/// Validates a slide-list `TextHeaderAtom` the way the text-interaction
/// reader validates it, and returns its text type.
fn text_header_type(record: &Record) -> Result<TextType> {
    let data = record.data.as_slice();
    if record.version != 0
        || record.record_type_raw != RecordType::TextHeaderAtom as u16
        || usize::try_from(record.data_length).ok() != Some(TEXT_HEADER_ATOM_SIZE)
        || data.len() != TEXT_HEADER_ATOM_SIZE
    {
        return Err(Error::Corrupted(
            "TextHeaderAtom has an invalid header or length".to_string(),
        ));
    }
    TextType::parse(u32::from_le_bytes([data[0], data[1], data[2], data[3]]))
}

/// Frames `records` as one payload, each with its eight-byte header,
/// allocated once at its exact length.
fn framed(records: &[&Record]) -> Result<Box<[u8]>> {
    let overflow = || Error::Corrupted("slide-list text body size overflow".to_string());
    let mut length = 0usize;
    for record in records {
        if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
            return Err(Error::Corrupted(
                "slide-list text record is truncated".to_string(),
            ));
        }
        length = length
            .checked_add(8)
            .and_then(|length| length.checked_add(record.data.len()))
            .ok_or_else(overflow)?;
    }
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(length)
        .map_err(|_err| Error::AllocationFailed("PPT slide-list text body"))?;
    for record in records {
        let version_instance = (record.version & 0x000f) | (record.instance << 4);
        payload.extend_from_slice(&version_instance.to_le_bytes());
        payload.extend_from_slice(&record.record_type_raw.to_le_bytes());
        payload.extend_from_slice(&record.data_length.to_le_bytes());
        payload.extend_from_slice(record.data.as_slice());
    }
    Ok(payload.into_boxed_slice())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn parse(bytes: &[u8]) -> Vec<Record> {
        let mut records = Vec::new();
        let mut offset = 0;
        while offset < bytes.len() {
            let (record, consumed) =
                Record::parse(bytes, offset, &SourceLimits::default()).unwrap();
            records.push(record);
            offset += consumed;
        }
        records
    }

    #[test]
    fn splits_bodies_at_text_headers_and_keeps_their_framing() {
        let title_header = record_bytes(0, 0, 3999, &0u32.to_le_bytes());
        let title_text = record_bytes(0, 0, 4008, b"Title");
        let body_header = record_bytes(0, 1, 3999, &1u32.to_le_bytes());
        let body_text = record_bytes(0, 0, 4000, &[b'A', 0, b'\r', 0, b'B', 0]);
        let special = record_bytes(0, 0, 4010, &0u32.to_le_bytes());
        // A record before the first TextHeaderAtom belongs to no body.
        let stray = record_bytes(0, 0, 4008, b"stray");
        let bytes = [
            stray,
            title_header.clone(),
            title_text.clone(),
            body_header.clone(),
            body_text.clone(),
            special.clone(),
        ]
        .concat();
        let records = parse(&bytes);
        let borrowed = records.iter().collect::<Vec<_>>();

        let bodies = OutlineTextBody::collect(&borrowed, SourceLimits::default()).unwrap();
        assert_eq!(bodies.len(), 2);
        assert_eq!(bodies[0].text_type(), TextType::Title);
        assert_eq!(bodies[0].text_header_instance(), 0);
        assert_eq!(bodies[0].records(), [title_header, title_text].concat());
        assert_eq!(bodies[0].text().unwrap(), "Title");
        assert_eq!(bodies[1].text_type(), TextType::Body);
        assert_eq!(bodies[1].text_header_instance(), 1);
        assert_eq!(
            bodies[1].records(),
            [body_header, body_text, special].concat()
        );
        assert_eq!(bodies[1].text().unwrap(), "A\rB");
    }

    #[test]
    fn a_body_without_characters_has_empty_text() {
        let bytes = [
            record_bytes(0, 0, 3999, &1u32.to_le_bytes()),
            record_bytes(0, 0, 4010, &0u32.to_le_bytes()),
        ]
        .concat();
        let records = parse(&bytes);
        let bodies =
            OutlineTextBody::collect(&records.iter().collect::<Vec<_>>(), SourceLimits::default())
                .unwrap();
        assert_eq!(bodies.len(), 1);
        assert_eq!(bodies[0].text().unwrap(), "");
    }

    #[test]
    fn body_text_honors_the_retained_source_depth_and_event_limits() {
        let header = record_bytes(0, 0, 3999, &1u32.to_le_bytes());
        let text = record_bytes(0, 0, 4000, &[b'A', 0]);
        let nested = record_bytes(0, 0, 1000, &text);
        let outer = record_bytes(0, 0, 1000, &nested);
        let records = parse(&[header, outer].concat());
        let depth_limited = SourceLimits::builder()
            .max_depth(1)
            .unwrap()
            .build()
            .unwrap();
        let bodies =
            OutlineTextBody::collect(&records.iter().collect::<Vec<_>>(), depth_limited).unwrap();
        assert!(bodies[0].text().is_err());

        let header = record_bytes(0, 0, 3999, &1u32.to_le_bytes());
        let first = record_bytes(0, 0, 4000, &[b'A', 0]);
        let second = record_bytes(0, 0, 4000, &[b'B', 0]);
        let records = parse(&[header, first, second].concat());
        let event_limited = SourceLimits::builder()
            .max_events(1)
            .unwrap()
            .build()
            .unwrap();
        let bodies =
            OutlineTextBody::collect(&records.iter().collect::<Vec<_>>(), event_limited).unwrap();
        assert!(bodies[0].text().is_err());
    }

    #[test]
    fn rejects_malformed_text_headers() {
        for bytes in [
            // Undefined TextTypeEnum value.
            record_bytes(0, 0, 3999, &3u32.to_le_bytes()),
            // Wrong payload length.
            record_bytes(0, 0, 3999, &[0; 2]),
            // Nonzero record version.
            record_bytes(1, 0, 3999, &0u32.to_le_bytes()),
        ] {
            let records = parse(&bytes);
            assert!(matches!(
                OutlineTextBody::collect(
                    &records.iter().collect::<Vec<_>>(),
                    SourceLimits::default(),
                ),
                Err(Error::Corrupted(_))
            ));
        }
    }

    #[test]
    fn rejects_records_truncated_by_their_container() {
        let mut header = Record::parse(
            &record_bytes(0, 0, 3999, &0u32.to_le_bytes()),
            0,
            &SourceLimits::default(),
        )
        .unwrap()
        .0;
        let text = Record::parse(
            &record_bytes(0, 0, 4008, b"abc"),
            0,
            &SourceLimits::default(),
        )
        .unwrap()
        .0;
        let mut truncated = text.clone();
        truncated.data_length = 9;
        assert!(OutlineTextBody::collect(&[&header, &text], SourceLimits::default()).is_ok());
        assert!(matches!(
            OutlineTextBody::collect(&[&header, &truncated], SourceLimits::default()),
            Err(Error::Corrupted(_))
        ));
        header.data_length = 8;
        assert!(OutlineTextBody::collect(&[&header], SourceLimits::default()).is_err());
    }
}

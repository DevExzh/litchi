//! Bounded, inert slide-library synchronization metadata for PowerPoint slides.
//!
//! Implements the ECMA-376 Part 4 section 4.7 Slide Synchronization Data part
//! (`CT_SlideSyncProperties`, root element `p:sldSyncPr`) that links a slide to
//! a slide-library item on a server. Loading and storing this part never
//! contacts the server, opens the library, or performs any synchronization.

use crate::common::mce::process_ooxml;
use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use litchi_core::xml::escape_xml;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

/// Content type of a Slide Synchronization Data part (ECMA-376 Part 1).
pub const SLIDE_SYNC_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideSyncData+xml";

/// Relationship type from a Slide part to its Slide Synchronization Data part
/// (ECMA-376 Part 1).
pub const SLIDE_SYNC_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncData";

const MAX_PART_XML_BYTES: usize = 1024 * 1024;
const MAX_TOTAL_XML_BYTES: usize = 64 * 1024 * 1024;
/// A slide library can hold at most one synchronized copy per slide; slides
/// are themselves bounded well below this figure by the package readers.
const MAX_SYNC_PARTS: usize = 4096;
const MAX_XML_NODES: usize = 4096;
const MAX_XML_DEPTH: usize = 32;
const MAX_SERVER_SLIDE_ID_BYTES: usize = 4096;
/// Digits retained after the decimal point of an `xsd:dateTime` second.
const MAX_FRACTION_DIGITS: usize = 16;
/// `xsd:dateTime` years are written as at least four digits; PowerPoint never
/// emits years outside the four-digit range.
const MIN_YEAR: u32 = 1;
const MAX_YEAR: u32 = 9999;
/// XML Schema 1.0 section 3.2.7.3: a time-zone offset must not exceed 14:00.
const MAX_TIMEZONE_HOURS: u8 = 14;

/// UTC offset recorded by an `xsd:dateTime` slide-synchronization timestamp.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlideSyncOffset {
    /// No time-zone designator was recorded.
    Unspecified,
    /// The `Z` designator (UTC).
    Utc,
    /// An explicit `+hh:mm` or `-hh:mm` offset from UTC.
    Offset {
        /// True when the offset is behind UTC (`-hh:mm`).
        negative: bool,
        /// Whole-hour component (`0..=14`).
        hours: u8,
        /// Minute component (`0..=59`, and `0` when `hours == 14`).
        minutes: u8,
    },
}

/// A validated `xsd:dateTime` timestamp from a Slide Synchronization Data
/// part. Leap-second values (`ss = 60`, permitted by the lexical space of
/// `xsd:dateTime`) are rejected because PowerPoint never records them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideSyncDateTime {
    /// Gregorian year (`1..=9999`).
    pub year: u32,
    /// Month (`1..=12`).
    pub month: u8,
    /// Day of the month, valid for the recorded year and month.
    pub day: u8,
    /// Hour (`0..=23`).
    pub hour: u8,
    /// Minute (`0..=59`).
    pub minute: u8,
    /// Second (`0..=59`).
    pub second: u8,
    /// Fractional-second digits (`1..=16` characters), when recorded.
    pub fraction_digits: Option<String>,
    /// Recorded UTC offset.
    pub offset: SlideSyncOffset,
}

impl SlideSyncDateTime {
    /// Parse and validate an `xsd:dateTime` lexical form.
    pub fn parse(value: &str) -> Result<Self> {
        let digits = |text: &str, label: &str| -> Result<u32> {
            if text.is_empty() || !text.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid(format!(
                    "slide-synchronization {label} is not numeric"
                )));
            }
            text.parse::<u32>()
                .map_err(|_| invalid(format!("slide-synchronization {label} is out of range")))
        };
        let (date, rest) = value.split_once('T').ok_or_else(|| {
            invalid("slide-synchronization timestamp is missing the 'T' separator")
        })?;
        let mut date_parts = date.split('-');
        let year_text = date_parts.next().unwrap_or_default();
        if year_text.len() != 4 {
            return Err(invalid("slide-synchronization year is not four digits"));
        }
        let year = digits(year_text, "year")?;
        if !(MIN_YEAR..=MAX_YEAR).contains(&year) {
            return Err(invalid("slide-synchronization year is out of range"));
        }
        let month = digits(date_parts.next().unwrap_or_default(), "month")?;
        let day = digits(date_parts.next().unwrap_or_default(), "day")?;
        if date_parts.next().is_some() {
            return Err(invalid("slide-synchronization date has trailing fields"));
        }
        let month = u8::try_from(month)
            .map_err(|_| invalid("slide-synchronization month is out of range"))?;
        let day =
            u8::try_from(day).map_err(|_| invalid("slide-synchronization day is out of range"))?;
        if chrono::NaiveDate::from_ymd_opt(year as i32, month.into(), day.into()).is_none() {
            return Err(invalid("slide-synchronization date is not a calendar day"));
        }

        let (time, offset) = match rest.find(['Z', '+', '-']) {
            Some(index) => {
                let (time, zone) = rest.split_at(index);
                let offset = match zone.as_bytes()[0] {
                    b'Z' => {
                        if zone.len() != 1 {
                            return Err(invalid(
                                "slide-synchronization UTC designator has trailing data",
                            ));
                        }
                        SlideSyncOffset::Utc
                    },
                    marker => {
                        let mut zone_parts = zone[1..].split(':');
                        let hours =
                            digits(zone_parts.next().unwrap_or_default(), "time-zone hours")?;
                        let minutes =
                            digits(zone_parts.next().unwrap_or_default(), "time-zone minutes")?;
                        if zone_parts.next().is_some() {
                            return Err(invalid(
                                "slide-synchronization time zone has trailing fields",
                            ));
                        }
                        let hours = u8::try_from(hours).map_err(|_| {
                            invalid("slide-synchronization time-zone hours are out of range")
                        })?;
                        let minutes = u8::try_from(minutes).map_err(|_| {
                            invalid("slide-synchronization time-zone minutes are out of range")
                        })?;
                        if hours > MAX_TIMEZONE_HOURS
                            || (hours == MAX_TIMEZONE_HOURS && minutes != 0)
                            || minutes > 59
                        {
                            return Err(invalid("slide-synchronization time zone exceeds 14:00"));
                        }
                        SlideSyncOffset::Offset {
                            negative: marker == b'-',
                            hours,
                            minutes,
                        }
                    },
                };
                (time, offset)
            },
            None => (rest, SlideSyncOffset::Unspecified),
        };

        let (hms, fraction_digits) = match time.split_once('.') {
            Some((hms, fraction)) => {
                if fraction.is_empty()
                    || fraction.len() > MAX_FRACTION_DIGITS
                    || !fraction.bytes().all(|byte| byte.is_ascii_digit())
                {
                    return Err(invalid(
                        "slide-synchronization fractional seconds are invalid",
                    ));
                }
                (hms, Some(fraction.to_string()))
            },
            None => (time, None),
        };
        let mut time_parts = hms.split(':');
        let hour = digits(time_parts.next().unwrap_or_default(), "hour")?;
        let minute = digits(time_parts.next().unwrap_or_default(), "minute")?;
        let second = digits(time_parts.next().unwrap_or_default(), "second")?;
        if time_parts.next().is_some() {
            return Err(invalid("slide-synchronization time has trailing fields"));
        }
        let hour = u8::try_from(hour)
            .map_err(|_| invalid("slide-synchronization hour is out of range"))?;
        let minute = u8::try_from(minute)
            .map_err(|_| invalid("slide-synchronization minute is out of range"))?;
        let second = u8::try_from(second)
            .map_err(|_| invalid("slide-synchronization second is out of range"))?;
        if hour > 23 || minute > 59 || second > 59 {
            return Err(invalid("slide-synchronization time of day is out of range"));
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
            fraction_digits,
            offset,
        })
    }

    /// Serialize in the canonical PowerPoint `xsd:dateTime` lexical form.
    pub fn to_lexical(&self) -> String {
        let mut out = format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            self.year, self.month, self.day, self.hour, self.minute, self.second
        );
        if let Some(ref fraction) = self.fraction_digits {
            out.push('.');
            out.push_str(fraction);
        }
        match self.offset {
            SlideSyncOffset::Unspecified => {},
            SlideSyncOffset::Utc => out.push('Z'),
            SlideSyncOffset::Offset {
                negative,
                hours,
                minutes,
            } => {
                out.push(if negative { '-' } else { '+' });
                out.push_str(&format!("{hours:02}:{minutes:02}"));
            },
        }
        out
    }
}

/// Typed `CT_SlideSyncProperties` (ECMA-376 Part 4 section 4.7.1): the
/// synchronization metadata linking one slide to a slide-library item.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideSyncProperties {
    /// Server-side slide file identifier (`serverSldId`).
    pub server_slide_id: String,
    /// Last modification time of the server-side slide (`serverSldModifiedTime`).
    pub server_modified_time: SlideSyncDateTime,
    /// Time the slide was inserted into this presentation (`clientInsertedTime`).
    pub client_inserted_time: SlideSyncDateTime,
    /// Whether the part carried a `p:extLst` extension list. Extension
    /// payloads are opaque vendor data and are not retained.
    pub has_extension_list: bool,
}

impl SlideSyncProperties {
    /// Parse the XML payload of a Slide Synchronization Data part.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_PART_XML_BYTES {
            return Err(limit("slide-synchronization XML bytes"));
        }
        let xml = process_ooxml(xml)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut properties: Option<Self> = None;
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            let is_empty_element = matches!(event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    nodes = increment(nodes, MAX_XML_NODES, "slide-synchronization XML nodes")?;
                    depth = increment(depth, MAX_XML_DEPTH, "slide-synchronization XML depth")?;
                    let is_sync_root =
                        is_presentationml_name(&namespace, element.name(), b"sldSyncPr");
                    let is_extension_list =
                        is_presentationml_name(&namespace, element.name(), b"extLst");
                    match depth {
                        1 => {
                            if properties.is_some() || closed_root {
                                return Err(invalid(
                                    "slide-synchronization part has multiple root elements",
                                ));
                            }
                            if !is_sync_root {
                                return Err(invalid(
                                    "slide-synchronization part must have a sldSyncPr root",
                                ));
                            }
                            properties = Some(Self::from_root(&element, decoder)?);
                        },
                        2 if is_extension_list => {
                            let entry = properties.as_mut().ok_or_else(|| {
                                invalid("missing slide-synchronization root element")
                            })?;
                            if entry.has_extension_list {
                                return Err(invalid(
                                    "duplicate slide-synchronization extension list",
                                ));
                            }
                            entry.has_extension_list = true;
                        },
                        2 => {
                            return Err(invalid(
                                "unexpected child of the slide-synchronization root",
                            ));
                        },
                        _ if is_sync_root => {
                            return Err(invalid("nested slide-synchronization root element"));
                        },
                        _ => {},
                    }
                    if is_empty_element {
                        depth -= 1;
                        if depth == 0 {
                            closed_root = true;
                        }
                    }
                },
                Event::Text(text) if depth > 0 => {
                    let content = text
                        .xml_content(quick_xml::XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    if depth == 1 && !content.trim().is_empty() {
                        return Err(invalid(
                            "slide-synchronization root cannot carry text content",
                        ));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid slide-synchronization XML nesting"))?;
                    if depth == 0 {
                        closed_root = true;
                    }
                },
                Event::Eof if depth != 0 => {
                    return Err(invalid("unterminated slide-synchronization XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        properties.ok_or_else(|| invalid("slide-synchronization part is empty"))
    }

    /// Serialize as a complete Slide Synchronization Data part.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        if self.server_slide_id.len() > MAX_SERVER_SLIDE_ID_BYTES {
            return Err(limit("slide-synchronization server slide ID bytes"));
        }
        let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        out.extend_from_slice(
            br#"<p:sldSyncPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" serverSldId=""#,
        );
        out.extend_from_slice(escape_xml(&self.server_slide_id).as_bytes());
        out.extend_from_slice(br#"" serverSldModifiedTime=""#);
        out.extend_from_slice(self.server_modified_time.to_lexical().as_bytes());
        out.extend_from_slice(br#"" clientInsertedTime=""#);
        out.extend_from_slice(self.client_inserted_time.to_lexical().as_bytes());
        if self.has_extension_list {
            out.extend_from_slice(br#""><p:extLst/></p:sldSyncPr>"#);
        } else {
            out.extend_from_slice(br#""/>"#);
        }
        if out.len() > MAX_PART_XML_BYTES {
            return Err(limit("serialized slide-synchronization XML bytes"));
        }
        Self::parse(&out)?;
        Ok(out)
    }

    fn from_root(element: &BytesStart<'_>, decoder: Decoder) -> Result<Self> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
            if attribute.key.prefix().is_some() {
                continue;
            }
            match attribute.key.local_name().as_ref() {
                b"serverSldId" | b"serverSldModifiedTime" | b"clientInsertedTime" => {},
                _ => {
                    return Err(invalid(format!(
                        "unexpected slide-synchronization attribute '{}'",
                        String::from_utf8_lossy(attribute.key.local_name().as_ref())
                    )));
                },
            }
        }
        let server_slide_id = unqualified_attribute_value(element, b"serverSldId", decoder)?
            .ok_or_else(|| invalid("missing slide-synchronization serverSldId attribute"))?;
        if server_slide_id.len() > MAX_SERVER_SLIDE_ID_BYTES {
            return Err(limit("slide-synchronization server slide ID bytes"));
        }
        let server_modified_time = SlideSyncDateTime::parse(
            &unqualified_attribute_value(element, b"serverSldModifiedTime", decoder)?.ok_or_else(
                || invalid("missing slide-synchronization serverSldModifiedTime attribute"),
            )?,
        )?;
        let client_inserted_time = SlideSyncDateTime::parse(
            &unqualified_attribute_value(element, b"clientInsertedTime", decoder)?.ok_or_else(
                || invalid("missing slide-synchronization clientInsertedTime attribute"),
            )?,
        )?;
        Ok(Self {
            server_slide_id,
            server_modified_time,
            client_inserted_time,
            has_extension_list: false,
        })
    }
}

/// A Slide Synchronization Data part bound to the slide that references it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideSyncPropertiesPart {
    /// Relationship ID the slide uses to reference the part.
    pub relationship_id: String,
    /// Part name of the source slide (for example `/ppt/slides/slide1.xml`).
    pub slide_part_name: String,
    /// Part name of the synchronization data part.
    pub part_name: String,
    /// Parsed synchronization metadata.
    pub properties: SlideSyncProperties,
}

/// Load every Slide Synchronization Data part in a package, validating that
/// each one is the target of exactly one implicit relationship from a Slide
/// part (ECMA-376 Part 1) and that slides synchronize at most one item.
pub fn load_slide_sync_properties(package: &OpcPackage) -> Result<Vec<SlideSyncPropertiesPart>> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == SLIDE_SYNC_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "slide-synchronization relationship cannot originate at the package root",
        ));
    }

    let mut loaded = Vec::new();
    let mut total_xml_bytes = 0usize;
    for source in package.iter_parts() {
        let relationships: Vec<_> = source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == SLIDE_SYNC_RELATIONSHIP_TYPE)
            .collect();
        if relationships.is_empty() {
            continue;
        }
        if source.content_type() != ct::PML_SLIDE {
            return Err(invalid(
                "slide-synchronization relationship must originate at a slide part",
            ));
        }
        if relationships.len() > 1 {
            return Err(invalid(
                "slide part has multiple slide-synchronization relationships",
            ));
        }
        let relationship = relationships[0];
        if relationship.is_external() {
            return Err(invalid(
                "slide-synchronization relationship cannot be external",
            ));
        }
        if loaded.len() >= MAX_SYNC_PARTS {
            return Err(limit("slide-synchronization part count"));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != SLIDE_SYNC_CONTENT_TYPE {
            return Err(OoxmlError::InvalidContentType {
                expected: SLIDE_SYNC_CONTENT_TYPE.into(),
                got: part.content_type().into(),
            });
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "slide-synchronization part cannot have outbound relationships",
            ));
        }
        total_xml_bytes = total_xml_bytes
            .checked_add(part.blob().len())
            .ok_or_else(|| limit("total slide-synchronization XML bytes"))?;
        if total_xml_bytes > MAX_TOTAL_XML_BYTES {
            return Err(limit("total slide-synchronization XML bytes"));
        }
        loaded.push(SlideSyncPropertiesPart {
            relationship_id: relationship.r_id().to_string(),
            slide_part_name: source.partname().to_string(),
            part_name: target.to_string(),
            properties: SlideSyncProperties::parse(part.blob())?,
        });
    }

    for part in package.iter_parts() {
        if part.content_type() != SLIDE_SYNC_CONTENT_TYPE {
            continue;
        }
        let references = loaded
            .iter()
            .filter(|entry| entry.part_name.as_str() == part.partname().as_str())
            .count();
        match references {
            1 => {},
            0 => {
                return Err(invalid(
                    "package contains an orphan slide-synchronization part",
                ));
            },
            _ => {
                return Err(invalid(
                    "slide-synchronization part is referenced by multiple slides",
                ));
            },
        }
    }
    Ok(loaded)
}

/// Attach a Slide Synchronization Data part to a slide, failing without side
/// effects when the slide already synchronizes a slide-library item.
pub fn store_slide_sync_properties(
    package: &mut OpcPackage,
    value: &SlideSyncPropertiesPart,
) -> Result<()> {
    validate_ncname(&value.relationship_id)?;
    let slide_name = PackURI::new(&value.slide_part_name).map_err(OoxmlError::InvalidUri)?;
    let part_name = PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    let slide = package.get_part(&slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_SLIDE.into(),
            got: slide.content_type().into(),
        });
    }
    if slide
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == SLIDE_SYNC_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "slide part already has a slide-synchronization relationship",
        ));
    }
    if slide.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(
            "slide-synchronization relationship ID already exists",
        ));
    }
    if package
        .iter_parts()
        .any(|part| part.partname() == &part_name)
    {
        return Err(invalid(format!(
            "part '{part_name}' already exists in the package"
        )));
    }

    let xml = value.properties.to_xml()?;
    let target = part_name.relative_ref(slide_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name,
        SLIDE_SYNC_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&slide_name)?
        .rels_mut()
        .add_relationship(
            SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

fn increment(value: usize, max: usize, label: &str) -> Result<usize> {
    let next = value.checked_add(1).ok_or_else(|| limit(label))?;
    if next > max {
        return Err(limit(label));
    }
    Ok(next)
}

fn validate_ncname(value: &str) -> Result<()> {
    let mut characters = value.chars();
    let valid = characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        });
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "slide-synchronization relationship ID is not an XML NCName",
        ))
    }
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("exceeded maximum {what}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const P_STRICT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const MAIN_CT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    const SLIDE_PATH: &str = "/ppt/slides/slide1.xml";
    const SYNC_PATH: &str = "/ppt/slideSyncData/slideSyncData1.xml";

    fn sync_xml() -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldSyncPr xmlns:p="{P}" serverSldId="server-slide-42" serverSldModifiedTime="2009-04-14T16:20:41.230" clientInsertedTime="2009-04-15T08:00:00Z"><p:extLst><p:ext uri="{{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}}"/></p:extLst></p:sldSyncPr>"#
        )
    }

    fn properties() -> SlideSyncProperties {
        SlideSyncProperties::parse(sync_xml().as_bytes()).unwrap()
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            MAIN_CT.into(),
            Vec::new(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(SLIDE_PATH).unwrap(),
            ct::PML_SLIDE.into(),
            Vec::new(),
        )));
        package
    }

    fn add_sync_part(package: &mut OpcPackage, relationship_id: &str) {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SYNC_PATH).unwrap(),
                SLIDE_SYNC_CONTENT_TYPE.into(),
                sync_xml().into_bytes(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(SLIDE_PATH).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
                "../slideSyncData/slideSyncData1.xml".into(),
                relationship_id.into(),
                false,
            );
    }

    #[test]
    fn parses_slide_sync_properties_and_strict_aliases() {
        let parsed = properties();
        assert_eq!(parsed.server_slide_id, "server-slide-42");
        assert_eq!(parsed.server_modified_time.year, 2009);
        assert_eq!(parsed.server_modified_time.month, 4);
        assert_eq!(parsed.server_modified_time.day, 14);
        assert_eq!(parsed.server_modified_time.hour, 16);
        assert_eq!(
            parsed.server_modified_time.fraction_digits.as_deref(),
            Some("230")
        );
        assert_eq!(
            parsed.server_modified_time.offset,
            SlideSyncOffset::Unspecified
        );
        assert_eq!(parsed.client_inserted_time.offset, SlideSyncOffset::Utc);
        assert!(parsed.has_extension_list);

        let strict = format!(
            r#"<x:sldSyncPr xmlns:x="{P_STRICT}" xmlns:f="urn:foreign" f:serverSldId="spoof" serverSldId="strict-id" serverSldModifiedTime="2024-02-29T23:59:58+05:30" clientInsertedTime="1601-01-01T00:00:00-14:00"/>"#
        );
        let parsed = SlideSyncProperties::parse(strict.as_bytes()).unwrap();
        assert_eq!(parsed.server_slide_id, "strict-id");
        assert_eq!(parsed.server_modified_time.day, 29);
        assert_eq!(
            parsed.server_modified_time.offset,
            SlideSyncOffset::Offset {
                negative: false,
                hours: 5,
                minutes: 30,
            }
        );
        assert_eq!(
            parsed.client_inserted_time.offset,
            SlideSyncOffset::Offset {
                negative: true,
                hours: 14,
                minutes: 0,
            }
        );
        assert!(!parsed.has_extension_list);
    }

    #[test]
    fn slide_sync_properties_round_trip() {
        for xml in [sync_xml(), {
            let parsed = SlideSyncProperties {
                has_extension_list: false,
                ..properties()
            };
            String::from_utf8(parsed.to_xml().unwrap()).unwrap()
        }] {
            let expected = SlideSyncProperties::parse(xml.as_bytes()).unwrap();
            let serialized = expected.to_xml().unwrap();
            let reparsed = SlideSyncProperties::parse(&serialized).unwrap();
            assert_eq!(reparsed, expected);
        }
        let lexical = properties().server_modified_time.to_lexical();
        assert_eq!(lexical, "2009-04-14T16:20:41.230");
    }

    #[test]
    fn rejects_malformed_slide_sync_documents() {
        let wrap = |body: &str| format!(r#"<p:sldSyncPr xmlns:p="{P}" {body}/>"#);
        let cases = [
            wrap(""),
            wrap(r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41""#),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41" vendor="x""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-13-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2023-02-29T16:20:41" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T24:20:41" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:60" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41Zx" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41+15:00" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41+14:01" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41. x" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41." clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="2009-04-14 16:20:41" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            wrap(
                r#"serverSldId="a" serverSldModifiedTime="009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41""#,
            ),
            format!(r#"<p:wrong xmlns:p="{P}"/>"#),
            r#"<f:sldSyncPr xmlns:f="urn:foreign"/>"#.to_string(),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41">text</p:sldSyncPr>"#
            ),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41"><p:sldSyncPr/></p:sldSyncPr>"#
            ),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41"><p:extLst/><p:extLst/></p:sldSyncPr>"#
            ),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41"><p:extLst><p:sldSyncPr/></p:extLst></p:sldSyncPr>"#
            ),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41"/>"#
            ) + &sync_xml(),
            format!(
                r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a" serverSldModifiedTime="2009-04-14T16:20:41" clientInsertedTime="2009-04-14T16:20:41" serverSldId="b"/>"#
            ),
            format!(r#"<p:sldSyncPr xmlns:p="{P}" serverSldId="a""#),
            r#"<?xml version="1.0"?>"#.to_string(),
        ];
        for xml in cases {
            assert!(
                SlideSyncProperties::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn package_round_trip_keeps_sync_metadata_inert() {
        let expected = SlideSyncPropertiesPart {
            relationship_id: "rIdSync".into(),
            slide_part_name: SLIDE_PATH.into(),
            part_name: SYNC_PATH.into(),
            properties: properties(),
        };
        let mut package = package();
        store_slide_sync_properties(&mut package, &expected).unwrap();
        let loaded = load_slide_sync_properties(&package).unwrap();
        assert_eq!(loaded, std::slice::from_ref(&expected));
        assert!(
            store_slide_sync_properties(&mut package, &expected).is_err(),
            "storing a second synchronization part for one slide must fail"
        );
    }

    #[test]
    fn rejects_hostile_slide_sync_package_graphs() {
        // Relationship from the package root.
        let mut root_rel = package();
        root_rel.rels_mut().add_relationship(
            SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
            "ppt/slideSyncData/slideSyncData1.xml".into(),
            "rIdSync".into(),
            false,
        );
        assert!(load_slide_sync_properties(&root_rel).is_err());

        // Relationship from a non-slide part.
        let mut wrong_source = package();
        wrong_source
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
                "../slideSyncData/slideSyncData1.xml".into(),
                "rIdSync".into(),
                false,
            );
        assert!(load_slide_sync_properties(&wrong_source).is_err());

        // Two synchronization relationships from one slide.
        let mut doubled = package();
        add_sync_part(&mut doubled, "rIdSync");
        doubled
            .get_part_mut(&PackURI::new(SLIDE_PATH).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
                "../slideSyncData/slideSyncData1.xml".into(),
                "rIdSync2".into(),
                false,
            );
        assert!(load_slide_sync_properties(&doubled).is_err());

        // External relationship.
        let mut external = package();
        external
            .get_part_mut(&PackURI::new(SLIDE_PATH).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
                "http://example.invalid/sync".into(),
                "rIdSync".into(),
                true,
            );
        assert!(load_slide_sync_properties(&external).is_err());

        // Target with the wrong content type.
        let mut wrong_ct = package();
        wrong_ct
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SYNC_PATH).unwrap(),
                ct::PML_SLIDE.into(),
                Vec::new(),
            )))
            .unwrap();
        wrong_ct
            .get_part_mut(&PackURI::new(SLIDE_PATH).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                SLIDE_SYNC_RELATIONSHIP_TYPE.into(),
                "../slideSyncData/slideSyncData1.xml".into(),
                "rIdSync".into(),
                false,
            );
        assert!(load_slide_sync_properties(&wrong_ct).is_err());

        // Orphan synchronization part.
        let mut orphan = package();
        orphan
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SYNC_PATH).unwrap(),
                SLIDE_SYNC_CONTENT_TYPE.into(),
                sync_xml().into_bytes(),
            )))
            .unwrap();
        assert!(load_slide_sync_properties(&orphan).is_err());

        // Synchronization part with outbound relationships.
        let mut outbound = package();
        add_sync_part(&mut outbound, "rIdSync");
        outbound
            .get_part_mut(&PackURI::new(SYNC_PATH).unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".into(),
                "../media/image1.png".into(),
                "rIdImage".into(),
                false,
            );
        assert!(load_slide_sync_properties(&outbound).is_err());
    }

    #[test]
    fn rejects_invalid_store_requests_without_side_effects() {
        let mut package = package();
        let base = SlideSyncPropertiesPart {
            relationship_id: "rIdSync".into(),
            slide_part_name: SLIDE_PATH.into(),
            part_name: SYNC_PATH.into(),
            properties: properties(),
        };
        for broken in [
            SlideSyncPropertiesPart {
                relationship_id: "1bad".into(),
                ..base.clone()
            },
            SlideSyncPropertiesPart {
                slide_part_name: "/ppt/presentation.xml".into(),
                ..base.clone()
            },
            SlideSyncPropertiesPart {
                slide_part_name: "/ppt/slides/missing.xml".into(),
                ..base.clone()
            },
            SlideSyncPropertiesPart {
                part_name: SLIDE_PATH.into(),
                ..base.clone()
            },
        ] {
            assert!(store_slide_sync_properties(&mut package, &broken).is_err());
        }
        assert!(load_slide_sync_properties(&package).unwrap().is_empty());
    }
}

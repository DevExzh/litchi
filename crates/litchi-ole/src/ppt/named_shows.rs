//! Strict, inert PowerPoint named/custom slide-show metadata.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use std::collections::HashSet;

const MAX_NAMED_SHOWS: usize = 4_096;
const MAX_NAME_UNITS: usize = 32_768;
const MAX_SLIDES_PER_SHOW: usize = 1_048_576;

/// One named show in source order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointNamedShow {
    pub name: String,
    /// Ordered slide references. `None` preserves an absent slides atom.
    pub slide_ids: Option<Vec<u32>>,
}

impl PowerPointNamedShow {
    /// Return ordered references that exist in the presentation.
    ///
    /// Source references are retained for lossless rewriting even though
    /// MS-PPT requires consumers to ignore references to absent slides.
    pub fn resolved_slide_ids<'a>(
        &'a self,
        presentation_slide_ids: &'a HashSet<u32>,
    ) -> impl Iterator<Item = u32> + 'a {
        self.slide_ids
            .iter()
            .flatten()
            .copied()
            .filter(|id| *id != 0 && presentation_slide_ids.contains(id))
    }
}

/// The optional `NamedShowsContainer` directly below a presentation document.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PowerPointNamedShows {
    pub shows: Vec<PowerPointNamedShow>,
}

impl PowerPointNamedShows {
    pub fn parse(document: &PptRecord) -> Result<Option<Self>> {
        let containers = document
            .children
            .iter()
            .filter(|record| record.record_type_raw == PptRecordType::NamedShows.as_u16())
            .collect::<Vec<_>>();
        if containers.len() > 1 {
            return corrupted("DocumentContainer contains duplicate NamedShowsContainer records");
        }
        let Some(container) = containers.first() else {
            return Ok(None);
        };
        require_header(
            container,
            0x0f,
            0,
            PptRecordType::NamedShows,
            "NamedShowsContainer",
        )?;
        if usize::try_from(container.data_length).ok() != Some(container.data.len()) {
            return corrupted("NamedShowsContainer has a truncated payload");
        }
        let children = PptRecord::parse_sequence_strict(&container.data, "NamedShowsContainer")?;
        if children.len() > MAX_NAMED_SHOWS {
            return corrupted(format!(
                "NamedShowsContainer exceeds {MAX_NAMED_SHOWS} named shows"
            ));
        }
        let shows = children
            .iter()
            .map(parse_named_show)
            .collect::<Result<Vec<_>>>()?;
        Ok(Some(Self { shows }))
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = PptRecord::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical NamedShowsContainer did not consume its record bytes");
        }
        Ok(record)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        if self.shows.len() > MAX_NAMED_SHOWS {
            return corrupted(format!(
                "NamedShowsContainer exceeds {MAX_NAMED_SHOWS} named shows"
            ));
        }
        let mut children = Vec::new();
        for show in &self.shows {
            let name = record_bytes(
                0,
                0,
                PptRecordType::CString.as_u16(),
                &encode_name(&show.name)?,
            )?;
            let mut show_children = name;
            if let Some(slide_ids) = &show.slide_ids {
                if slide_ids.len() > MAX_SLIDES_PER_SHOW {
                    return corrupted(format!(
                        "named show exceeds {MAX_SLIDES_PER_SHOW} slide references"
                    ));
                }
                let mut payload = Vec::with_capacity(slide_ids.len().saturating_mul(4));
                for &slide_id in slide_ids {
                    validate_slide_id_ref(slide_id)?;
                    payload.extend_from_slice(&slide_id.to_le_bytes());
                }
                show_children.extend_from_slice(&record_bytes(
                    0,
                    0,
                    PptRecordType::NamedShowSlides.as_u16(),
                    &payload,
                )?);
            }
            children.extend_from_slice(&record_bytes(
                0x0f,
                0,
                PptRecordType::NamedShow.as_u16(),
                &show_children,
            )?);
        }
        record_bytes(0x0f, 0, PptRecordType::NamedShows.as_u16(), &children)
    }
}

fn parse_named_show(record: &PptRecord) -> Result<PowerPointNamedShow> {
    require_header(
        record,
        0x0f,
        0,
        PptRecordType::NamedShow,
        "NamedShowContainer",
    )?;
    if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
        return corrupted("NamedShowContainer has a truncated payload");
    }
    let children = PptRecord::parse_sequence_strict(&record.data, "NamedShowContainer")?;
    if !(1..=2).contains(&children.len()) {
        return corrupted("NamedShowContainer must contain a name and at most one slide-list atom");
    }
    let name = parse_name(&children[0])?;
    let slide_ids = if let Some(slides) = children.get(1) {
        require_header(
            slides,
            0,
            0,
            PptRecordType::NamedShowSlides,
            "NamedShowSlidesAtom",
        )?;
        if slides.data.len() % 4 != 0 {
            return corrupted("NamedShowSlidesAtom length is not a multiple of four");
        }
        let count = slides.data.len() / 4;
        if count > MAX_SLIDES_PER_SHOW {
            return corrupted(format!(
                "named show exceeds {MAX_SLIDES_PER_SHOW} slide references"
            ));
        }
        let mut ids = Vec::with_capacity(count);
        for bytes in slides.data.chunks_exact(4) {
            let id = u32::from_le_bytes(bytes.try_into().expect("four-byte chunk"));
            validate_slide_id_ref(id)?;
            ids.push(id);
        }
        Some(ids)
    } else {
        None
    };
    Ok(PowerPointNamedShow { name, slide_ids })
}

fn parse_name(record: &PptRecord) -> Result<String> {
    require_header(record, 0, 0, PptRecordType::CString, "NamedShowNameAtom")?;
    if !record.data.len().is_multiple_of(2) {
        return corrupted("NamedShowNameAtom has odd UTF-16 byte length");
    }
    if record.data.len() / 2 > MAX_NAME_UNITS {
        return corrupted(format!(
            "NamedShowNameAtom exceeds {MAX_NAME_UNITS} UTF-16 code units"
        ));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        if matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f) {
            return corrupted("NamedShowNameAtom contains a non-printable character");
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted("NamedShowNameAtom contains invalid UTF-16".to_string()))
}

fn encode_name(name: &str) -> Result<Vec<u8>> {
    let units = name.encode_utf16().collect::<Vec<_>>();
    if units.len() > MAX_NAME_UNITS {
        return corrupted(format!(
            "NamedShowNameAtom exceeds {MAX_NAME_UNITS} UTF-16 code units"
        ));
    }
    if units
        .iter()
        .any(|unit| matches!(*unit, 0x0000..=0x001f | 0x007f..=0x009f))
    {
        return corrupted("NamedShowNameAtom contains a non-printable character");
    }
    Ok(units.into_iter().flat_map(u16::to_le_bytes).collect())
}

fn validate_slide_id_ref(id: u32) -> Result<()> {
    if id != 0 && !(0x0000_0100..=0x7fff_ffff).contains(&id) {
        return corrupted(format!("named show contains invalid SlideIdRef {id:#010x}"));
    }
    Ok(())
}

fn require_header(
    record: &PptRecord,
    version: u16,
    instance: u16,
    record_type: PptRecordType,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != record_type.as_u16()
    {
        return corrupted(format!("invalid {context} record header"));
    }
    Ok(())
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".to_string()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root(children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            version: 0x0f,
            instance: 0,
            record_type: PptRecordType::Document,
            record_type_raw: PptRecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    fn parsed_record(bytes: &[u8]) -> PptRecord {
        PptRecord::parse(bytes, 0).unwrap().0
    }

    #[test]
    fn protocol_shaped_named_shows_roundtrip_and_resolve() {
        let shows = PowerPointNamedShows {
            shows: vec![
                PowerPointNamedShow {
                    name: "Executive overview".into(),
                    slide_ids: Some(vec![0x100, 0x222, 0, 0x101]),
                },
                PowerPointNamedShow {
                    name: "Empty".into(),
                    slide_ids: Some(Vec::new()),
                },
                PowerPointNamedShow {
                    name: String::new(),
                    slide_ids: None,
                },
            ],
        };
        let parsed = PowerPointNamedShows::parse(&root(vec![shows.to_record().unwrap()]))
            .unwrap()
            .unwrap();
        assert_eq!(parsed, shows);
        let existing = HashSet::from([0x100, 0x101]);
        assert_eq!(
            parsed.shows[0]
                .resolved_slide_ids(&existing)
                .collect::<Vec<_>>(),
            vec![0x100, 0x101]
        );
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            shows.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn rejects_duplicate_and_malformed_container_grammar() {
        let valid = PowerPointNamedShows {
            shows: vec![PowerPointNamedShow {
                name: "Show".into(),
                slide_ids: Some(vec![0x100]),
            }],
        }
        .to_record()
        .unwrap();
        assert!(PowerPointNamedShows::parse(&root(vec![valid.clone(), valid])).is_err());

        let name_data = ('S' as u16).to_le_bytes();
        let name = record_bytes(0, 0, PptRecordType::CString.as_u16(), &name_data).unwrap();
        let slides = record_bytes(0, 0, PptRecordType::NamedShowSlides.as_u16(), &[0; 4]).unwrap();
        for children in [
            Vec::new(),
            slides.clone(),
            [slides.clone(), name.clone()].concat(),
            [name.clone(), slides.clone(), slides].concat(),
        ] {
            let show = record_bytes(0x0f, 0, PptRecordType::NamedShow.as_u16(), &children).unwrap();
            let outer = record_bytes(0x0f, 0, PptRecordType::NamedShows.as_u16(), &show).unwrap();
            assert!(PowerPointNamedShows::parse(&root(vec![parsed_record(&outer)])).is_err());
        }
    }

    #[test]
    fn rejects_hostile_names_lengths_and_slide_ids() {
        for name in ["line\nbreak", "nul\0suffix"] {
            let shows = PowerPointNamedShows {
                shows: vec![PowerPointNamedShow {
                    name: name.into(),
                    slide_ids: None,
                }],
            };
            assert!(shows.to_record_bytes().is_err());
        }
        for slide_id in [1, 0xff, 0x8000_0000, u32::MAX] {
            let shows = PowerPointNamedShows {
                shows: vec![PowerPointNamedShow {
                    name: "Show".into(),
                    slide_ids: Some(vec![slide_id]),
                }],
            };
            assert!(shows.to_record_bytes().is_err());
        }
        for child in [
            record_bytes(0, 0, PptRecordType::CString.as_u16(), b"x").unwrap(),
            record_bytes(
                0,
                0,
                PptRecordType::CString.as_u16(),
                &0xd800u16.to_le_bytes(),
            )
            .unwrap(),
            record_bytes(0, 0, PptRecordType::CString.as_u16(), &1u16.to_le_bytes()).unwrap(),
            record_bytes(0, 0, PptRecordType::NamedShowSlides.as_u16(), &[0, 1, 2]).unwrap(),
        ] {
            let show = record_bytes(0x0f, 0, PptRecordType::NamedShow.as_u16(), &child).unwrap();
            let outer = record_bytes(0x0f, 0, PptRecordType::NamedShows.as_u16(), &show).unwrap();
            assert!(PowerPointNamedShows::parse(&root(vec![parsed_record(&outer)])).is_err());
        }
    }
}

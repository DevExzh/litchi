//! Normative logical slide directory for binary PowerPoint files.

use crate::consts::PptRecordType;
use crate::ppt::current_user::CurrentUser;
use crate::ppt::package::{PptError, Result};
use crate::ppt::persist::PersistMapping;
use crate::ppt::records::PptRecord;
use std::collections::HashMap;

const USER_EDIT_ATOM: u16 = 0x0ff5;
const SLIDE_LIST_INSTANCE: u16 = 0;
const SLIDE_PERSIST_ATOM_SIZE: usize = 20;
const SLIDE_CONTAINER: u16 = 1006;

/// One logical presentation-slide reference from `SlideListWithTextContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideDirectoryEntry {
    persist_id: u32,
    slide_id: u32,
    flags: u32,
    text_placeholder_count: u32,
    list_text: String,
}

impl SlideDirectoryEntry {
    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    pub fn flags(&self) -> u32 {
        self.flags
    }

    pub fn text_placeholder_count(&self) -> u32 {
        self.text_placeholder_count
    }

    pub fn list_text(&self) -> &str {
        &self.list_text
    }
}

/// Ordered logical presentation slides from the live `DocumentContainer`.
#[derive(Debug, Clone)]
pub struct SlideDirectory {
    document_persist_id: u32,
    document_offset: usize,
    entries: Vec<SlideDirectoryEntry>,
    by_slide_id: HashMap<u32, usize>,
    by_persist_id: HashMap<u32, usize>,
}

impl SlideDirectory {
    pub(crate) fn build(
        document_data: &[u8],
        current_user_data: &[u8],
        persist_mapping: &PersistMapping,
    ) -> Result<Self> {
        let current_user = CurrentUser::parse(current_user_data)?;
        let user_edit_offset = usize::try_from(current_user.current_edit_offset())
            .map_err(|_| PptError::Corrupted("current edit offset does not fit usize".to_string()))?;
        let user_edit = read_header(document_data, user_edit_offset, "UserEditAtom")?;
        if user_edit.version != 0 || user_edit.instance != 0 || user_edit.record_type != USER_EDIT_ATOM {
            return Err(PptError::Corrupted(
                "CurrentUser does not reference a valid UserEditAtom".to_string(),
            ));
        }
        if !matches!(user_edit.data_len, 28 | 32) {
            return Err(PptError::Corrupted(format!(
                "UserEditAtom has invalid length {}",
                user_edit.data_len
            )));
        }
        let payload_offset = user_edit_offset
            .checked_add(8)
            .ok_or_else(|| PptError::Corrupted("UserEditAtom offset overflow".to_string()))?;
        let user_edit_data = checked_slice(
            document_data,
            payload_offset,
            user_edit.data_len,
            "UserEditAtom",
        )?;
        let document_persist_id = read_u32(user_edit_data, 16, "docPersistIdRef")?;
        if document_persist_id == 0 {
            return Err(PptError::Corrupted(
                "UserEditAtom has a null docPersistIdRef".to_string(),
            ));
        }
        let document_offset = persist_mapping
            .get_offset(document_persist_id)
            .ok_or_else(|| {
                PptError::Corrupted(format!(
                    "document persist ID {document_persist_id} has no directory entry"
                ))
            })?;
        let document_offset = usize::try_from(document_offset).map_err(|_| {
            PptError::Corrupted("document persist offset does not fit usize".to_string())
        })?;
        let (document, _) = PptRecord::parse(document_data, document_offset)?;
        if document.record_type != PptRecordType::Document
            || document.version != 0x0f
            || document.instance != 0
        {
            return Err(PptError::Corrupted(format!(
                "persist ID {document_persist_id} does not resolve to a DocumentContainer"
            )));
        }

        let mut slide_list = None;
        for child in document.extract_slide_list_with_texts() {
            if child.get_instance() != SLIDE_LIST_INSTANCE {
                continue;
            }
            if child.version != 0x0f {
                return Err(PptError::Corrupted(
                    "presentation SlideListWithTextContainer has invalid version".to_string(),
                ));
            }
            if slide_list.replace(child).is_some() {
                return Err(PptError::Corrupted(
                    "duplicate presentation SlideListWithTextContainer".to_string(),
                ));
            }
        }

        let mut entries: Vec<SlideDirectoryEntry> = Vec::new();
        let mut by_slide_id = HashMap::new();
        let mut by_persist_id = HashMap::new();
        if let Some(slide_list) = slide_list {
            entries.reserve(slide_list.data.len() / 28);
            for child in &slide_list.children {
                if child.record_type != PptRecordType::SlidePersistAtom {
                    if let Some(entry) = entries.last_mut()
                        && let Ok(value) = child.extract_text()
                        && !value.is_empty()
                    {
                        if !entry.list_text.is_empty() {
                            entry.list_text.push('\n');
                        }
                        entry.list_text.push_str(&value);
                    }
                    continue;
                }
                if child.version != 0
                    || child.instance != 0
                    || child.data.len() != SLIDE_PERSIST_ATOM_SIZE
                {
                    return Err(PptError::Corrupted(format!(
                        "invalid SlidePersistAtom header or length: version={}, instance={}, length={}",
                        child.version,
                        child.instance,
                        child.data.len()
                    )));
                }
                let persist_id = read_u32(&child.data, 0, "SlidePersistAtom.persistIdRef")?;
                let flags = read_u32(&child.data, 4, "SlidePersistAtom flags")?;
                let text_placeholder_count =
                    read_u32(&child.data, 8, "SlidePersistAtom.cTexts")?;
                let slide_id = read_u32(&child.data, 12, "SlidePersistAtom.slideId")?;
                if persist_id == 0 || slide_id == 0 {
                    return Err(PptError::Corrupted(
                        "SlidePersistAtom has a null persistIdRef or slideId".to_string(),
                    ));
                }
                if text_placeholder_count > 8 {
                    return Err(PptError::Corrupted(format!(
                        "SlidePersistAtom cTexts exceeds 8: {text_placeholder_count}"
                    )));
                }
                if by_slide_id.contains_key(&slide_id) {
                    return Err(PptError::Corrupted(format!(
                        "duplicate presentation slideId {slide_id}"
                    )));
                }
                if by_persist_id.contains_key(&persist_id) {
                    return Err(PptError::Corrupted(format!(
                        "duplicate presentation persistIdRef {persist_id}"
                    )));
                }
                let slide_offset = persist_mapping.get_offset(persist_id).ok_or_else(|| {
                    PptError::Corrupted(format!(
                        "slide persist ID {persist_id} has no directory entry"
                    ))
                })?;
                let slide_offset = usize::try_from(slide_offset).map_err(|_| {
                    PptError::Corrupted("slide persist offset does not fit usize".to_string())
                })?;
                let slide_header = read_header(document_data, slide_offset, "SlideContainer")?;
                if slide_header.record_type != SLIDE_CONTAINER
                    || slide_header.version != 0x0f
                    || slide_header.instance != 0
                {
                    return Err(PptError::Corrupted(format!(
                        "persist ID {persist_id} does not resolve to a SlideContainer"
                    )));
                }
                let index = entries.len();
                entries.push(SlideDirectoryEntry {
                    persist_id,
                    slide_id,
                    flags,
                    text_placeholder_count,
                    list_text: String::new(),
                });
                by_slide_id.insert(slide_id, index);
                by_persist_id.insert(persist_id, index);
            }
        }

        Ok(Self {
            document_persist_id,
            document_offset,
            entries,
            by_slide_id,
            by_persist_id,
        })
    }

    pub fn entries(&self) -> &[SlideDirectoryEntry] {
        &self.entries
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn get_by_slide_id(&self, slide_id: u32) -> Option<&SlideDirectoryEntry> {
        self.by_slide_id.get(&slide_id).map(|&index| &self.entries[index])
    }

    pub fn get_by_persist_id(&self, persist_id: u32) -> Option<&SlideDirectoryEntry> {
        self.by_persist_id
            .get(&persist_id)
            .map(|&index| &self.entries[index])
    }

    pub fn document_persist_id(&self) -> u32 {
        self.document_persist_id
    }

    pub(crate) fn document_offset(&self) -> usize {
        self.document_offset
    }

    #[cfg(test)]
    pub(crate) fn new_for_test(document_offset: usize) -> Self {
        Self {
            document_persist_id: 1,
            document_offset,
            entries: Vec::new(),
            by_slide_id: HashMap::new(),
            by_persist_id: HashMap::new(),
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct RawHeader {
    version: u16,
    instance: u16,
    record_type: u16,
    data_len: usize,
}

fn read_header(data: &[u8], offset: usize, name: &str) -> Result<RawHeader> {
    let bytes = checked_slice(data, offset, 8, &format!("{name} header"))?;
    let version_instance = u16::from_le_bytes(bytes[0..2].try_into().unwrap());
    let data_len = usize::try_from(u32::from_le_bytes(bytes[4..8].try_into().unwrap()))
        .map_err(|_| PptError::Corrupted(format!("{name} length does not fit usize")))?;
    let total = 8usize
        .checked_add(data_len)
        .ok_or_else(|| PptError::Corrupted(format!("{name} length overflow")))?;
    checked_slice(data, offset, total, name)?;
    Ok(RawHeader {
        version: version_instance & 0x000f,
        instance: version_instance >> 4,
        record_type: u16::from_le_bytes(bytes[2..4].try_into().unwrap()),
        data_len,
    })
}

fn checked_slice<'a>(data: &'a [u8], offset: usize, len: usize, name: &str) -> Result<&'a [u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| PptError::Corrupted(format!("{name} range overflow")))?;
    data.get(offset..end)
        .ok_or_else(|| PptError::Corrupted(format!("{name} is truncated")))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let bytes = checked_slice(data, offset, 4, name)?;
    Ok(u32::from_le_bytes(bytes.try_into().unwrap()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::Package;

    fn record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&(data.len() as u32).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn slide_persist(persist_id: u32, slide_id: u32, c_texts: u32) -> Vec<u8> {
        let mut data = Vec::with_capacity(20);
        data.extend_from_slice(&persist_id.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&c_texts.to_le_bytes());
        data.extend_from_slice(&slide_id.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        record(0, 0, 0x03f3, &data)
    }

    fn current_user(edit_offset: u32) -> Vec<u8> {
        let mut data = vec![0u8; 32];
        data[2..4].copy_from_slice(&0x0ff6u16.to_le_bytes());
        data[4..8].copy_from_slice(&24u32.to_le_bytes());
        data[8..12].copy_from_slice(&20u32.to_le_bytes());
        data[12..16].copy_from_slice(&0xe391_c05fu32.to_le_bytes());
        data[16..20].copy_from_slice(&edit_offset.to_le_bytes());
        data[22..24].copy_from_slice(&0x03f4u16.to_le_bytes());
        data[28..32].copy_from_slice(&8u32.to_le_bytes());
        data
    }

    fn synthetic_directory(
        entries: &[(u32, u32, u32)],
        extra_slide: bool,
    ) -> (Vec<u8>, Vec<u8>, PersistMapping) {
        let slide_list_data: Vec<u8> = entries
            .iter()
            .flat_map(|&(persist_id, slide_id, c_texts)| {
                slide_persist(persist_id, slide_id, c_texts)
            })
            .collect();
        let slide_list = record(0x0f, 0, 0x0ff0, &slide_list_data);
        let document = record(0x0f, 0, 0x03e8, &slide_list);
        let mut stream = document;
        let mut mapping = PersistMapping::new();
        mapping.add_mapping(1, 0);
        for &(persist_id, _, _) in entries {
            let offset = stream.len() as u32;
            mapping.add_mapping(persist_id, offset);
            stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        }
        if extra_slide {
            let offset = stream.len() as u32;
            mapping.add_mapping(99, offset);
            stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        }
        let edit_offset = stream.len() as u32;
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));
        (stream, current_user(edit_offset), mapping)
    }

    #[test]
    fn preserves_list_order_and_excludes_unreferenced_slides() {
        let (stream, current_user, mapping) = synthetic_directory(
            &[(11, 300, 0), (5, 100, 0), (9, 200, 0)],
            true,
        );
        let directory = SlideDirectory::build(&stream, &current_user, &mapping).unwrap();
        assert_eq!(
            directory.entries().iter().map(|entry| entry.persist_id()).collect::<Vec<_>>(),
            [11, 5, 9]
        );
        assert!(directory.get_by_persist_id(99).is_none());
    }

    #[test]
    fn excludes_rt_slide_referenced_only_by_master_list() {
        let master_list = record(0x0f, 1, 0x0ff0, &slide_persist(8, 0x8000_0001, 0));
        let slide_list = record(0x0f, 0, 0x0ff0, &slide_persist(3, 256, 0));
        let mut document_data = master_list;
        document_data.extend_from_slice(&slide_list);
        let mut stream = record(0x0f, 0, 0x03e8, &document_data);
        let master_offset = stream.len() as u32;
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        let slide_offset = stream.len() as u32;
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        let edit_offset = stream.len() as u32;
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));

        let mut mapping = PersistMapping::new();
        mapping.add_mapping(1, 0);
        mapping.add_mapping(8, master_offset);
        mapping.add_mapping(3, slide_offset);
        let directory =
            SlideDirectory::build(&stream, &current_user(edit_offset), &mapping).unwrap();
        assert_eq!(directory.entries().len(), 1);
        assert_eq!(directory.entries()[0].persist_id(), 3);
        assert!(directory.get_by_persist_id(8).is_none());
    }

    #[test]
    fn rejects_duplicate_malformed_and_wrong_type_entries() {
        let (stream, current_user, mapping) =
            synthetic_directory(&[(3, 256, 0), (4, 256, 0)], false);
        assert!(SlideDirectory::build(&stream, &current_user, &mapping).is_err());

        let (stream, current_user, mapping) = synthetic_directory(&[(3, 256, 9)], false);
        assert!(SlideDirectory::build(&stream, &current_user, &mapping).is_err());

        let (stream, current_user, mut mapping) = synthetic_directory(&[(3, 256, 0)], false);
        mapping.add_mapping(3, 0);
        assert!(SlideDirectory::build(&stream, &current_user, &mapping).is_err());
    }

    #[test]
    fn current_user_selects_live_document_in_multi_edit_stream() {
        let (mut stream, _, mut mapping) = synthetic_directory(&[(3, 256, 0)], false);
        let live_document_offset = stream.len() as u32;
        let live_list = record(0x0f, 0, 0x0ff0, &slide_persist(4, 900, 0));
        stream.extend_from_slice(&record(0x0f, 0, 0x03e8, &live_list));
        mapping.add_mapping(2, live_document_offset);
        let slide_offset = stream.len() as u32;
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        mapping.add_mapping(4, slide_offset);
        let edit_offset = stream.len() as u32;
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&2u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));

        let directory =
            SlideDirectory::build(&stream, &current_user(edit_offset), &mapping).unwrap();
        assert_eq!(directory.entries()[0].slide_id(), 900);
        assert_eq!(directory.document_persist_id(), 2);
    }

    #[test]
    fn poi_reordered_and_basic_fixtures_match_normative_ids() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/poi/test-data/slideshow");
        let mut package = Package::open(root.join("incorrect_slide_order.ppt")).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        assert_eq!(
            slides.iter().map(|slide| slide.slide_id()).collect::<Vec<_>>(),
            [256, 258, 257]
        );
        assert_eq!(
            slides.iter().map(|slide| slide.persist_id()).collect::<Vec<_>>(),
            [3, 5, 4]
        );
        for (slide, title) in slides.iter().zip(["Slide 1", "Slide 2", "Slide 3"]) {
            let text = slide.text().unwrap();
            assert!(text.contains(title), "expected {title:?}, got {text:?}");
        }

        let mut package = Package::open(root.join("basic_test_ppt_file.ppt")).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        assert_eq!(
            slides.iter().map(|slide| slide.slide_id()).collect::<Vec<_>>(),
            [256, 257]
        );
        assert_eq!(
            slides.iter().map(|slide| slide.persist_id()).collect::<Vec<_>>(),
            [4, 6]
        );
        assert_eq!(presentation.slide_count(), 2);
        assert_eq!(presentation.extract_text_fast().unwrap().len(), 2);
    }
}

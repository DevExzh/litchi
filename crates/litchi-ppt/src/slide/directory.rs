//! Normative logical slide directory for binary `PowerPoint` files.

use crate::consts::RecordType;
use crate::current_user::CurrentUser;
use crate::package::{Error, RecordLimits, Result};
use crate::persist::PersistMapping;
use crate::records::Record;
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
    outline_text_interactions: Vec<crate::TextBodyInteractions>,
    outline_text_refs: Vec<crate::OutlineTextRef>,
}

impl SlideDirectoryEntry {
    #[must_use]
    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    #[must_use]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    #[must_use]
    pub fn flags(&self) -> u32 {
        self.flags
    }

    #[must_use]
    pub fn text_placeholder_count(&self) -> u32 {
        self.text_placeholder_count
    }

    #[must_use]
    pub fn list_text(&self) -> &str {
        &self.list_text
    }

    #[must_use]
    pub fn outline_text_interactions(&self) -> &[crate::TextBodyInteractions] {
        &self.outline_text_interactions
    }

    /// Validated outline text references of this slide's shapes.
    #[must_use]
    pub fn outline_text_refs(&self) -> &[crate::OutlineTextRef] {
        &self.outline_text_refs
    }
}

/// Ordered logical presentation slides from the live `DocumentContainer`.
#[allow(
    clippy::module_name_repetitions,
    reason = "`SlideDirectory` is the established public API name re-exported as `slide::SlideDirectory`; renaming it would break downstream crates"
)]
#[derive(Debug, Clone)]
pub struct SlideDirectory {
    document_persist_id: u32,
    document_offset: usize,
    entries: Vec<SlideDirectoryEntry>,
    by_slide_id: HashMap<u32, usize>,
    by_persist_id: HashMap<u32, usize>,
}

impl SlideDirectory {
    #[cfg(test)]
    pub(crate) fn build(
        document_data: &[u8],
        current_user_data: &[u8],
        persist_mapping: &PersistMapping,
    ) -> Result<Self> {
        Self::build_with_limits(
            document_data,
            current_user_data,
            persist_mapping,
            RecordLimits::default(),
        )
    }

    pub(crate) fn build_with_limits(
        document_data: &[u8],
        current_user_data: &[u8],
        persist_mapping: &PersistMapping,
        limits: RecordLimits,
    ) -> Result<Self> {
        if document_data.len() > limits.max_input_bytes {
            return Err(Error::ResourceLimit(format!(
                "PowerPoint Document stream size {} exceeds limit {}",
                document_data.len(),
                limits.max_input_bytes
            )));
        }
        let current_user = CurrentUser::parse_with_limits(current_user_data, limits)?;
        let user_edit_offset =
            usize::try_from(current_user.current_edit_offset()).map_err(|_err| {
                Error::Corrupted("current edit offset does not fit usize".to_string())
            })?;
        let user_edit = read_header(document_data, user_edit_offset, "UserEditAtom")?;
        if user_edit.version != 0
            || user_edit.instance != 0
            || user_edit.record_type != USER_EDIT_ATOM
        {
            return Err(Error::Corrupted(
                "CurrentUser does not reference a valid UserEditAtom".to_string(),
            ));
        }
        if !matches!(user_edit.data_len, 28 | 32) {
            return Err(Error::Corrupted(format!(
                "UserEditAtom has invalid length {}",
                user_edit.data_len
            )));
        }
        let payload_offset = user_edit_offset
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("UserEditAtom offset overflow".to_string()))?;
        let user_edit_data = checked_slice(
            document_data,
            payload_offset,
            user_edit.data_len,
            "UserEditAtom",
        )?;
        let document_persist_id = read_u32(user_edit_data, 16, "docPersistIdRef")?;
        if document_persist_id == 0 {
            return Err(Error::Corrupted(
                "UserEditAtom has a null docPersistIdRef".to_string(),
            ));
        }
        let document_offset = persist_mapping
            .get_offset(document_persist_id)
            .ok_or_else(|| {
                Error::Corrupted(format!(
                    "document persist ID {document_persist_id} has no directory entry"
                ))
            })?;
        let document_offset_usize = usize::try_from(document_offset).map_err(|_err| {
            Error::Corrupted("document persist offset does not fit usize".to_string())
        })?;
        let (document, _) =
            Record::parse_with_limits(document_data, document_offset_usize, limits)?;
        if document.record_type != RecordType::Document
            || document.version != 0x0f
            || document.instance != 0
        {
            return Err(Error::Corrupted(format!(
                "persist ID {document_persist_id} does not resolve to a DocumentContainer"
            )));
        }

        let mut slide_list = None;
        for child in document.extract_slide_list_with_texts() {
            if child.get_instance() != SLIDE_LIST_INSTANCE {
                continue;
            }
            if child.version != 0x0f {
                return Err(Error::Corrupted(
                    "presentation SlideListWithTextContainer has invalid version".to_string(),
                ));
            }
            if slide_list.replace(child).is_some() {
                return Err(Error::Corrupted(
                    "duplicate presentation SlideListWithTextContainer".to_string(),
                ));
            }
        }

        let mut entries: Vec<SlideDirectoryEntry> = Vec::new();
        let mut by_slide_id = HashMap::new();
        let mut by_persist_id = HashMap::new();
        if let Some(slide_list_container) = slide_list {
            entries
                .try_reserve(slide_list_container.children.len())
                .map_err(|_err| Error::AllocationFailed("PPT slide directory"))?;
            by_slide_id
                .try_reserve(slide_list_container.children.len())
                .map_err(|_err| Error::AllocationFailed("PPT slide ID index"))?;
            by_persist_id
                .try_reserve(slide_list_container.children.len())
                .map_err(|_err| Error::AllocationFailed("PPT slide persist index"))?;
            for child in &slide_list_container.children {
                if child.record_type != RecordType::SlidePersistAtom {
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
                    return Err(Error::Corrupted(format!(
                        "invalid SlidePersistAtom header or length: version={}, instance={}, length={}",
                        child.version,
                        child.instance,
                        child.data.len()
                    )));
                }
                let persist_id = read_u32(&child.data, 0, "SlidePersistAtom.persistIdRef")?;
                let flags = read_u32(&child.data, 4, "SlidePersistAtom flags")?;
                let text_placeholder_count = read_u32(&child.data, 8, "SlidePersistAtom.cTexts")?;
                let slide_id = read_u32(&child.data, 12, "SlidePersistAtom.slideId")?;
                if persist_id == 0 || slide_id == 0 {
                    return Err(Error::Corrupted(
                        "SlidePersistAtom has a null persistIdRef or slideId".to_string(),
                    ));
                }
                if text_placeholder_count > 8 {
                    return Err(Error::Corrupted(format!(
                        "SlidePersistAtom cTexts exceeds 8: {text_placeholder_count}"
                    )));
                }
                if by_slide_id.contains_key(&slide_id) {
                    return Err(Error::Corrupted(format!(
                        "duplicate presentation slideId {slide_id}"
                    )));
                }
                if by_persist_id.contains_key(&persist_id) {
                    return Err(Error::Corrupted(format!(
                        "duplicate presentation persistIdRef {persist_id}"
                    )));
                }
                let slide_offset = persist_mapping.get_offset(persist_id).ok_or_else(|| {
                    Error::Corrupted(format!(
                        "slide persist ID {persist_id} has no directory entry"
                    ))
                })?;
                let slide_offset_usize = usize::try_from(slide_offset).map_err(|_err| {
                    Error::Corrupted("slide persist offset does not fit usize".to_string())
                })?;
                let slide_header =
                    read_header(document_data, slide_offset_usize, "SlideContainer")?;
                if slide_header.record_type != SLIDE_CONTAINER
                    || slide_header.version != 0x0f
                    || slide_header.instance != 0
                {
                    return Err(Error::Corrupted(format!(
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
                    outline_text_interactions: Vec::new(),
                    outline_text_refs: Vec::new(),
                });
                by_slide_id.insert(slide_id, index);
                by_persist_id.insert(persist_id, index);
            }
            for set in slide_list_container.group_into_slide_atoms_sets() {
                let slide_id =
                    read_u32(&set.slide_persist_atom.data, 12, "SlidePersistAtom.slideId")?;
                let Some(&index) = by_slide_id.get(&slide_id) else {
                    return Err(Error::Corrupted(format!(
                        "text records reference unknown slideId {slide_id}"
                    )));
                };
                entries[index].outline_text_interactions = set.text_interactions()?;
                entries[index].outline_text_refs = set.outline_text_refs()?;
            }
        }

        Ok(Self {
            document_persist_id,
            document_offset: document_offset_usize,
            entries,
            by_slide_id,
            by_persist_id,
        })
    }

    #[must_use]
    pub fn entries(&self) -> &[SlideDirectoryEntry] {
        &self.entries
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    #[must_use]
    pub fn get_by_slide_id(&self, slide_id: u32) -> Option<&SlideDirectoryEntry> {
        self.by_slide_id
            .get(&slide_id)
            .map(|&index| &self.entries[index])
    }

    #[must_use]
    pub fn get_by_persist_id(&self, persist_id: u32) -> Option<&SlideDirectoryEntry> {
        self.by_persist_id
            .get(&persist_id)
            .map(|&index| &self.entries[index])
    }

    #[must_use]
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
    let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
    let data_len = usize::try_from(u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]))
        .map_err(|_err| Error::Corrupted(format!("{name} length does not fit usize")))?;
    let total = 8usize
        .checked_add(data_len)
        .ok_or_else(|| Error::Corrupted(format!("{name} length overflow")))?;
    checked_slice(data, offset, total, name)?;
    Ok(RawHeader {
        version: version_instance & 0x000f,
        instance: version_instance >> 4,
        record_type: u16::from_le_bytes([bytes[2], bytes[3]]),
        data_len,
    })
}

fn checked_slice<'a>(data: &'a [u8], offset: usize, len: usize, name: &str) -> Result<&'a [u8]> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| Error::Corrupted(format!("{name} range overflow")))?;
    data.get(offset..end)
        .ok_or_else(|| Error::Corrupted(format!("{name} is truncated")))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let bytes = checked_slice(data, offset, 4, name)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::Package;

    fn record(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
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
            let offset = u32::try_from(stream.len()).unwrap();
            mapping.add_mapping(persist_id, offset);
            stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        }
        if extra_slide {
            let offset = u32::try_from(stream.len()).unwrap();
            mapping.add_mapping(99, offset);
            stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        }
        let edit_offset = u32::try_from(stream.len()).unwrap();
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));
        (stream, current_user(edit_offset), mapping)
    }

    #[test]
    fn preserves_list_order_and_excludes_unreferenced_slides() {
        let (stream, current_user, mapping) =
            synthetic_directory(&[(11, 300, 0), (5, 100, 0), (9, 200, 0)], true);
        let directory = SlideDirectory::build(&stream, &current_user, &mapping).unwrap();
        assert_eq!(
            directory
                .entries()
                .iter()
                .map(SlideDirectoryEntry::persist_id)
                .collect::<Vec<_>>(),
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
        let master_offset = u32::try_from(stream.len()).unwrap();
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        let slide_offset = u32::try_from(stream.len()).unwrap();
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        let edit_offset = u32::try_from(stream.len()).unwrap();
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

        let (dup_stream, dup_user, dup_mapping) = synthetic_directory(&[(3, 256, 9)], false);
        assert!(SlideDirectory::build(&dup_stream, &dup_user, &dup_mapping).is_err());

        let (zero_stream, zero_user, mut zero_mapping) = synthetic_directory(&[(3, 256, 0)], false);
        zero_mapping.add_mapping(3, 0);
        assert!(SlideDirectory::build(&zero_stream, &zero_user, &zero_mapping).is_err());
    }

    #[test]
    fn current_user_selects_live_document_in_multi_edit_stream() {
        let (mut stream, _, mut mapping) = synthetic_directory(&[(3, 256, 0)], false);
        let live_document_offset = u32::try_from(stream.len()).unwrap();
        let live_list = record(0x0f, 0, 0x0ff0, &slide_persist(4, 900, 0));
        stream.extend_from_slice(&record(0x0f, 0, 0x03e8, &live_list));
        mapping.add_mapping(2, live_document_offset);
        let slide_offset = u32::try_from(stream.len()).unwrap();
        stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        mapping.add_mapping(4, slide_offset);
        let edit_offset = u32::try_from(stream.len()).unwrap();
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&2u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));

        let directory =
            SlideDirectory::build(&stream, &current_user(edit_offset), &mapping).unwrap();
        assert_eq!(directory.entries()[0].slide_id(), 900);
        assert_eq!(directory.document_persist_id(), 2);
    }

    #[test]
    fn carries_slide_list_text_range_interactions_into_directory_entries() {
        let mut atom = [0u8; 16];
        atom[12] = 0xff;
        let interaction = record(0x0f, 0, 4082, &record(0, 0, 4083, &atom));
        let anchor = record(0, 0, 4063, &[1, 0, 0, 0, 3, 0, 0, 0]);
        let mut slide_list_data = slide_persist(3, 256, 1);
        slide_list_data.extend_from_slice(&record(0, 0, 3999, &1u32.to_le_bytes()));
        slide_list_data.extend_from_slice(&record(0, 0, 4008, b"ABC"));
        slide_list_data.extend_from_slice(&interaction);
        slide_list_data.extend_from_slice(&anchor);
        let slide_list = record(0x0f, 0, 4080, &slide_list_data);
        let mut stream = record(0x0f, 0, 1000, &slide_list);
        let slide_offset = u32::try_from(stream.len()).unwrap();
        stream.extend_from_slice(&record(0x0f, 0, 1006, &[]));
        let edit_offset = u32::try_from(stream.len()).unwrap();
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 4085, &edit));
        let mut mapping = PersistMapping::new();
        mapping.add_mapping(1, 0);
        mapping.add_mapping(3, slide_offset);

        let directory =
            SlideDirectory::build(&stream, &current_user(edit_offset), &mapping).unwrap();
        let bodies = directory.entries()[0].outline_text_interactions();
        assert_eq!(bodies.len(), 1);
        assert_eq!(bodies[0].text, "ABC");
        assert_eq!(bodies[0].interactions[0].range.begin(), 1);
        assert_eq!(bodies[0].interactions[0].range.end(), 3);
    }

    #[test]
    fn poi_reordered_and_basic_fixtures_match_normative_ids() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/slideshow");
        let mut package = Package::open(root.join("incorrect_slide_order.ppt")).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        assert_eq!(
            slides
                .iter()
                .map(super::super::types::Slide::slide_id)
                .collect::<Vec<_>>(),
            [256, 258, 257]
        );
        assert_eq!(
            slides
                .iter()
                .map(super::super::types::Slide::persist_id)
                .collect::<Vec<_>>(),
            [3, 5, 4]
        );
        for (slide, title) in slides.iter().zip(["Slide 1", "Slide 2", "Slide 3"]) {
            let text = slide.text().unwrap();
            assert!(text.contains(title), "expected {title:?}, got {text:?}");
        }

        let mut basic_package = Package::open(root.join("basic_test_ppt_file.ppt")).unwrap();
        let basic_presentation = basic_package.presentation().unwrap();
        let basic_slides = basic_presentation.slides().unwrap();
        assert_eq!(
            basic_slides
                .iter()
                .map(super::super::types::Slide::slide_id)
                .collect::<Vec<_>>(),
            [256, 257]
        );
        assert_eq!(
            basic_slides
                .iter()
                .map(super::super::types::Slide::persist_id)
                .collect::<Vec<_>>(),
            [4, 6]
        );
        assert_eq!(basic_presentation.slide_count(), 2);
        assert_eq!(basic_presentation.extract_text_fast().unwrap().len(), 2);
    }
}

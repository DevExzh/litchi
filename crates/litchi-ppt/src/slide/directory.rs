//! Normative logical slide directory for binary `PowerPoint` files.

use crate::consts::RecordType;
use crate::current_user::CurrentUser;
use crate::leniency::{DefectLocation, SlideListDefect, ToleranceReport};
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
            &mut ToleranceReport::default(),
        )
    }

    /// Build the directory, repairing non-structural slide-list defects
    /// into `report` (findings F5.4 and F5.5; see [`SlideListDefect`]).
    pub(crate) fn build_with_limits(
        document_data: &[u8],
        current_user_data: &[u8],
        persist_mapping: &PersistMapping,
        limits: RecordLimits,
        report: &mut ToleranceReport,
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
            // The entry index each SlidePersistAtom produced, in order;
            // `None` for a dropped entry, whose text records are skipped.
            let mut atom_entries: Vec<Option<usize>> = Vec::new();
            atom_entries
                .try_reserve(slide_list_container.children.len())
                .map_err(|_err| Error::AllocationFailed("PPT slide directory"))?;
            let mut current = None;
            for child in &slide_list_container.children {
                if child.record_type != RecordType::SlidePersistAtom {
                    if let Some(entry) = current.and_then(|index: usize| entries.get_mut(index))
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
                current = directory_entry(
                    child,
                    document_data,
                    persist_mapping,
                    &by_slide_id,
                    &by_persist_id,
                    report,
                )?
                .map(|entry| {
                    let index = entries.len();
                    by_slide_id.insert(entry.slide_id, index);
                    by_persist_id.insert(entry.persist_id, index);
                    entries.push(entry);
                    index
                });
                atom_entries.push(current);
            }
            for (set, index) in slide_list_container
                .group_into_slide_atoms_sets()
                .into_iter()
                .zip(atom_entries)
            {
                let Some(index) = index else {
                    continue;
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

/// Decode one presentation `SlidePersistAtom` into a directory entry,
/// repairing non-structural defects into `report`.
///
/// Returns `Ok(None)` for an entry that is dropped (findings F5.4 and F5.5);
/// a slide record whose own framing overruns the document stream stays an
/// error.
fn directory_entry(
    atom: &Record,
    document_data: &[u8],
    persist_mapping: &PersistMapping,
    by_slide_id: &HashMap<u32, usize>,
    by_persist_id: &HashMap<u32, usize>,
    report: &mut ToleranceReport,
) -> Result<Option<SlideDirectoryEntry>> {
    let length = u32::try_from(atom.data.len()).unwrap_or(u32::MAX);
    // POI and LibreOffice read the fields by offset without validating the
    // header; persistIdRef, flags, cTexts, and slideId need 16 bytes.
    if atom.data.len() < 16 {
        report.record(SlideListDefect::TruncatedSlidePersistAtom, length);
        return Ok(None);
    }
    let persist_id = read_u32(&atom.data, 0, "SlidePersistAtom.persistIdRef")?;
    let flags = read_u32(&atom.data, 4, "SlidePersistAtom flags")?;
    let text_placeholder_count = read_u32(&atom.data, 8, "SlidePersistAtom.cTexts")?;
    let slide_id = read_u32(&atom.data, 12, "SlidePersistAtom.slideId")?;
    let location = DefectLocation {
        persist_id: (persist_id != 0).then_some(persist_id),
        shape_id: None,
    };
    if atom.version != 0 || atom.instance != 0 || atom.data.len() != SLIDE_PERSIST_ATOM_SIZE {
        report.record_at(SlideListDefect::SlidePersistAtomHeader, length, location);
    }
    if persist_id == 0 || slide_id == 0 {
        report.record_at(SlideListDefect::NullSlideReference, slide_id, location);
        return Ok(None);
    }
    if by_slide_id.contains_key(&slide_id) {
        report.record_at(SlideListDefect::DuplicateSlideId, slide_id, location);
        return Ok(None);
    }
    if by_persist_id.contains_key(&persist_id) {
        report.record_at(
            SlideListDefect::DuplicateSlidePersistRef,
            slide_id,
            location,
        );
        return Ok(None);
    }
    if !resolves_to_slide(document_data, persist_mapping, persist_id)? {
        report.record_at(SlideListDefect::StaleSlidePersistRef, slide_id, location);
        return Ok(None);
    }
    let text_placeholder_count = if text_placeholder_count > 8 {
        report.record_at(
            SlideListDefect::TextPlaceholderCount,
            text_placeholder_count,
            location,
        );
        8
    } else {
        text_placeholder_count
    };
    Ok(Some(SlideDirectoryEntry {
        persist_id,
        slide_id,
        flags,
        text_placeholder_count,
        list_text: String::new(),
        outline_text_interactions: Vec::new(),
        outline_text_refs: Vec::new(),
    }))
}

/// Whether `persist_id` maps to a `SlideContainer` header inside the
/// document stream.
///
/// A persist ID the live directory does not map, an offset whose record
/// header lies outside the stream, or a header of another record type is a
/// stale reference LibreOffice skips (finding F5.5). A `SlideContainer`
/// header whose length overruns the stream is broken framing and stays an
/// error.
fn resolves_to_slide(
    document_data: &[u8],
    persist_mapping: &PersistMapping,
    persist_id: u32,
) -> Result<bool> {
    let Some(offset) = persist_mapping
        .get_offset(persist_id)
        .and_then(|offset| usize::try_from(offset).ok())
    else {
        return Ok(false);
    };
    let Some(header) = offset
        .checked_add(8)
        .and_then(|end| document_data.get(offset..end))
    else {
        return Ok(false);
    };
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let record_type = u16::from_le_bytes([header[2], header[3]]);
    if record_type != SLIDE_CONTAINER || version_instance != 0x000f {
        return Ok(false);
    }
    read_header(document_data, offset, "SlideContainer")?;
    Ok(true)
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

    fn build_reporting(
        stream: &[u8],
        current_user: &[u8],
        mapping: &PersistMapping,
    ) -> (SlideDirectory, Vec<(crate::Defect, u32, Option<u32>)>) {
        let mut report = ToleranceReport::default();
        let directory = SlideDirectory::build_with_limits(
            stream,
            current_user,
            mapping,
            RecordLimits::default(),
            &mut report,
        )
        .unwrap();
        let defects = report
            .defects()
            .iter()
            .map(|defect| (defect.defect, defect.context, defect.location.persist_id))
            .collect();
        (directory, defects)
    }

    fn persist_ids(directory: &SlideDirectory) -> Vec<u32> {
        directory
            .entries()
            .iter()
            .map(SlideDirectoryEntry::persist_id)
            .collect()
    }

    /// One bad slide-list entry is dropped (or repaired) and recorded; the
    /// rest of the directory still builds (F5.4).
    #[test]
    fn repairs_duplicate_and_malformed_entries_per_entry() {
        use crate::Defect;

        let (stream, current_user, mapping) =
            synthetic_directory(&[(3, 256, 0), (4, 256, 0), (5, 257, 0)], false);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [3, 5]);
        assert_eq!(
            defects,
            [(
                Defect::SlideList(SlideListDefect::DuplicateSlideId),
                256,
                Some(4)
            )]
        );

        let (stream, current_user, mapping) =
            synthetic_directory(&[(3, 256, 0), (3, 257, 0)], false);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [3]);
        assert_eq!(
            defects,
            [(
                Defect::SlideList(SlideListDefect::DuplicateSlidePersistRef),
                257,
                Some(3)
            )]
        );

        let (stream, current_user, mapping) = synthetic_directory(&[(3, 256, 9)], false);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(directory.entries()[0].text_placeholder_count(), 8);
        assert_eq!(
            defects,
            [(
                Defect::SlideList(SlideListDefect::TextPlaceholderCount),
                9,
                Some(3)
            )]
        );

        let (stream, current_user, mapping) = synthetic_directory(&[(3, 0, 0), (4, 300, 0)], false);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [4]);
        assert_eq!(
            defects,
            [(
                Defect::SlideList(SlideListDefect::NullSlideReference),
                0,
                Some(3)
            )]
        );

        // The conformant control repairs nothing.
        let (stream, current_user, mapping) =
            synthetic_directory(&[(3, 256, 8), (4, 257, 0)], false);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [3, 4]);
        assert!(defects.is_empty());
    }

    /// A nonconforming SlidePersistAtom header is read by field offset when
    /// the fields are present and dropped when they are not (F5.4).
    #[test]
    fn reads_nonconforming_slide_persist_atoms_by_offset() {
        use crate::Defect;

        let slide_list_data = [
            // recInstance 1, 24-byte payload: read by offset.
            record(0, 1, 0x03f3, &{
                let mut data = slide_persist(3, 256, 0)[8..].to_vec();
                data.extend_from_slice(&[0; 4]);
                data
            }),
            // 12-byte payload: no slideId, dropped.
            record(0, 0, 0x03f3, &slide_persist(4, 257, 0)[8..20]),
            slide_persist(5, 258, 0),
        ]
        .concat();
        let slide_list = record(0x0f, 0, 0x0ff0, &slide_list_data);
        let mut stream = record(0x0f, 0, 0x03e8, &slide_list);
        let mut mapping = PersistMapping::new();
        mapping.add_mapping(1, 0);
        for persist_id in [3, 4, 5] {
            mapping.add_mapping(persist_id, u32::try_from(stream.len()).unwrap());
            stream.extend_from_slice(&record(0x0f, 0, 0x03ee, &[]));
        }
        let edit_offset = u32::try_from(stream.len()).unwrap();
        let mut edit = vec![0u8; 28];
        edit[16..20].copy_from_slice(&1u32.to_le_bytes());
        stream.extend_from_slice(&record(0, 0, 0x0ff5, &edit));

        let (directory, defects) = build_reporting(&stream, &current_user(edit_offset), &mapping);
        assert_eq!(persist_ids(&directory), [3, 5]);
        assert_eq!(
            directory
                .entries()
                .iter()
                .map(SlideDirectoryEntry::slide_id)
                .collect::<Vec<_>>(),
            [256, 258]
        );
        assert_eq!(
            defects,
            [
                (
                    Defect::SlideList(SlideListDefect::SlidePersistAtomHeader),
                    24,
                    Some(3)
                ),
                (
                    Defect::SlideList(SlideListDefect::TruncatedSlidePersistAtom),
                    12,
                    None
                ),
            ]
        );
    }

    /// A slide reference the live persist directory does not map, that maps
    /// outside the stream, or that maps to another record is skipped and
    /// recorded, as LibreOffice skips it (F5.5).
    #[test]
    fn skips_stale_slide_references() {
        use crate::Defect;

        let stale = |persist_id| {
            (
                Defect::SlideList(SlideListDefect::StaleSlidePersistRef),
                256,
                Some(persist_id),
            )
        };

        // Mapped to the DocumentContainer at offset 0.
        let (stream, current_user, mut mapping) =
            synthetic_directory(&[(3, 256, 0), (4, 257, 0)], false);
        mapping.add_mapping(3, 0);
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [4]);
        assert_eq!(defects, [stale(3)]);

        // Not mapped at all.
        let (stream, current_user, mapping) =
            synthetic_directory(&[(3, 256, 0), (4, 257, 0)], false);
        let mut unmapped = PersistMapping::new();
        for (&persist_id, &offset) in mapping.iter() {
            if persist_id != 3 {
                unmapped.add_mapping(persist_id, offset);
            }
        }
        let (directory, defects) = build_reporting(&stream, &current_user, &unmapped);
        assert_eq!(persist_ids(&directory), [4]);
        assert_eq!(defects, [stale(3)]);

        // Mapped past the end of the stream.
        let (stream, current_user, mut mapping) =
            synthetic_directory(&[(3, 256, 0), (4, 257, 0)], false);
        mapping.add_mapping(3, u32::try_from(stream.len()).unwrap());
        let (directory, defects) = build_reporting(&stream, &current_user, &mapping);
        assert_eq!(persist_ids(&directory), [4]);
        assert_eq!(defects, [stale(3)]);
    }

    /// A `SlideContainer` header whose length overruns the document stream
    /// is broken framing, not a stale reference.
    #[test]
    fn rejects_slide_containers_that_overrun_the_stream() {
        let (mut stream, current_user, mapping) = synthetic_directory(&[(3, 256, 0)], false);
        let slide_offset = mapping.get_offset(3).unwrap() as usize;
        stream[slide_offset + 4..slide_offset + 8].copy_from_slice(&0x00FF_FFFFu32.to_le_bytes());
        assert!(SlideDirectory::build(&stream, &current_user, &mapping).is_err());
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

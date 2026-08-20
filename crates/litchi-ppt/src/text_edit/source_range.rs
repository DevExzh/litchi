//! Positional, selector-only reads for the source-backed PPT text owner.
//!
//! This module deliberately stops at the semantic boundary. It resolves the
//! live UserEdit/PersistDirectory mapping, the slide directory, macro owners,
//! and one selected Slide through `SharedOleFile::read_stream_range`; the
//! common CFB publisher still owns full-artifact identity and candidate
//! validation.

use super::{
    Error, Refusal, Result, SourceResolved, Target, inspect_live_macro_records,
    inspect_source_text_atom_parts, map_ole_error, native_shape_id,
};
use crate::consts::RecordType;
use crate::current_user::CurrentUser;
use crate::package::{Error as PackageError, RecordLimits};
use crate::records::{Record, RecordParseSession};
use litchi_cfb::SharedOleFile;
use std::collections::{HashMap, HashSet};

const HEADER_LEN: usize = 8;
const USER_EDIT_ATOM: u16 = 4085;
const PERSIST_FULL: u16 = 6001;
const PERSIST_INCREMENTAL: u16 = 6002;
const SLIDE_LIST_INSTANCE: u16 = 0;
const SLIDE_PERSIST_ATOM_SIZE: usize = 20;
const SLIDE_CONTAINER: u16 = 1006;
const USER_EDIT_CHAIN_LIMIT: usize = 4_096;

#[derive(Debug, Clone, Copy)]
struct Header {
    bytes: [u8; HEADER_LEN],
    version: u16,
    instance: u16,
    record_type: u16,
    data_len: usize,
    total_len: usize,
}

struct StreamReader<'a> {
    shared: &'a SharedOleFile,
    refs: Vec<&'a str>,
    length: u64,
}

impl<'a> StreamReader<'a> {
    fn new(shared: &'a SharedOleFile, path: &'a [String], length: u64) -> Self {
        Self {
            shared,
            refs: path.iter().map(String::as_str).collect(),
            length,
        }
    }

    fn read_into(&self, offset: u64, output: &mut [u8], label: &str) -> Result<()> {
        let end = offset
            .checked_add(u64::try_from(output.len()).map_err(|_error| {
                PackageError::ResourceLimit(format!("{label} range length does not fit u64"))
            })?)
            .ok_or_else(|| PackageError::Corrupted(format!("{label} range overflows")))?;
        if end > self.length {
            return Err(PackageError::Corrupted(format!(
                "{label} range {offset}..{end} exceeds stream length {}",
                self.length
            ))
            .into());
        }
        self.shared
            .read_stream_range(&self.refs, offset, output)
            .map_err(map_ole_error)
    }

    fn read_vec(&self, offset: u64, length: usize, label: &str) -> Result<Vec<u8>> {
        let mut output = Vec::new();
        output
            .try_reserve_exact(length)
            .map_err(|_error| PackageError::AllocationFailed("PPT source-backed range"))?;
        output.resize(length, 0);
        self.read_into(offset, &mut output, label)?;
        Ok(output)
    }

    fn header(&self, offset: u64, label: &str) -> Result<Header> {
        let mut bytes = [0_u8; HEADER_LEN];
        self.read_into(offset, &mut bytes, &format!("{label} header"))?;
        let version_instance = u16::from_le_bytes([bytes[0], bytes[1]]);
        let data_len =
            usize::try_from(u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]])).map_err(
                |_error| PackageError::Corrupted(format!("{label} length does not fit usize")),
            )?;
        let total_len = HEADER_LEN
            .checked_add(data_len)
            .ok_or_else(|| PackageError::Corrupted(format!("{label} length overflows")))?;
        let end = offset
            .checked_add(u64::try_from(total_len).map_err(|_error| {
                PackageError::ResourceLimit(format!("{label} length does not fit u64"))
            })?)
            .ok_or_else(|| PackageError::Corrupted(format!("{label} range overflows")))?;
        if end > self.length {
            return Err(PackageError::Corrupted(format!(
                "{label} extends beyond its containing stream"
            ))
            .into());
        }
        Ok(Header {
            bytes,
            version: version_instance & 0x000F,
            instance: version_instance >> 4,
            record_type: u16::from_le_bytes([bytes[2], bytes[3]]),
            data_len,
            total_len,
        })
    }

    fn record(
        &self,
        offset: u64,
        header: Header,
        label: &str,
        limits: RecordLimits,
    ) -> Result<Vec<u8>> {
        if header.data_len > limits.max_record_payload_bytes {
            return Err(PackageError::ResourceLimit(format!(
                "PPT record payload size {} exceeds limit {}",
                header.data_len, limits.max_record_payload_bytes
            ))
            .into());
        }
        if header.total_len > limits.max_record_bytes {
            return Err(PackageError::ResourceLimit(format!(
                "PPT record size {} exceeds limit {}",
                header.total_len, limits.max_record_bytes
            ))
            .into());
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(header.total_len)
            .map_err(|_error| PackageError::AllocationFailed("PPT source-backed record"))?;
        output.extend_from_slice(&header.bytes);
        output.resize(header.total_len, 0);
        if header.data_len != 0 {
            self.read_into(
                offset.checked_add(HEADER_LEN as u64).ok_or_else(|| {
                    PackageError::Corrupted(format!("{label} payload offset overflows"))
                })?,
                &mut output[HEADER_LEN..],
                &format!("{label} payload"),
            )?;
        }
        Ok(output)
    }
}

#[derive(Debug)]
struct LiveMapping {
    mappings: HashMap<u32, u32>,
    document_persist_id: u32,
    document_offset: u64,
}

#[derive(Debug)]
struct SlideEntry {
    persist_id: u32,
    offset: u64,
}

#[derive(Debug)]
struct SlideIndex {
    entries: Vec<SlideEntry>,
}

/// Resolve one source-backed semantic target without materializing either
/// complete metadata stream.
pub(super) fn resolve_source_target(
    shared: &SharedOleFile,
    document_path: &[String],
    current_user_path: &[String],
    target: Target,
    limits: RecordLimits,
) -> Result<SourceResolved> {
    let document_refs: Vec<_> = document_path.iter().map(String::as_str).collect();
    let current_user_refs: Vec<_> = current_user_path.iter().map(String::as_str).collect();
    let document_length = shared.stream_len(&document_refs).map_err(map_ole_error)?;
    let current_user_length = shared
        .stream_len(&current_user_refs)
        .map_err(map_ole_error)?;
    let aggregate_length = document_length
        .checked_add(current_user_length)
        .ok_or_else(|| PackageError::ResourceLimit("PPT metadata stream sizes overflow".into()))?;
    let input_limit = u64::try_from(limits.max_input_bytes)
        .map_err(|_error| PackageError::ResourceLimit("PPT input limit does not fit u64".into()))?;
    let aggregate_limit = u64::try_from(limits.max_aggregate_input_bytes).map_err(|_error| {
        PackageError::ResourceLimit("PPT aggregate input limit does not fit u64".into())
    })?;
    if document_length > input_limit {
        return Err(PackageError::ResourceLimit(format!(
            "PowerPoint Document stream size {document_length} exceeds limit {}",
            limits.max_input_bytes
        ))
        .into());
    }
    if current_user_length > input_limit {
        return Err(PackageError::ResourceLimit(format!(
            "CurrentUser stream size {current_user_length} exceeds limit {}",
            limits.max_input_bytes
        ))
        .into());
    }
    if aggregate_length > aggregate_limit {
        return Err(PackageError::ResourceLimit(format!(
            "PPT Document and Current User streams total {aggregate_length} bytes, exceeding limit {}",
            limits.max_aggregate_input_bytes
        ))
        .into());
    }

    let document = StreamReader::new(shared, document_path, document_length);
    let current_user_reader = StreamReader::new(shared, current_user_path, current_user_length);
    let current_user = read_current_user(&current_user_reader, limits)?;
    if current_user.is_encrypted() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }

    let mut session = RecordParseSession::new(limits, 0)?;
    let mapping = read_live_mapping(
        &document,
        current_user.current_edit_offset(),
        limits,
        &mut session,
    )?;
    let index = build_slide_index(&document, &mapping, limits, &mut session)?;
    let slide_position = target.slide.get();
    let entry =
        index
            .entries
            .get(slide_position)
            .ok_or(Error::Refused(Refusal::SlideNotFound {
                position: target.slide,
            }))?;

    let slide_header = document.header(entry.offset, "SlideContainer")?;
    let slide_bytes = document.record(entry.offset, slide_header, "SlideContainer", limits)?;
    let (slide, consumed) = session.parse_strict_record(&slide_bytes, 0)?;
    if consumed != slide_bytes.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    if slide.record_type != RecordType::Slide {
        return Err(
            PackageError::Corrupted("selected persist record is not a Slide".into()).into(),
        );
    }
    let ppdrawing = slide
        .find_child(RecordType::PPDrawing)
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let shapes = crate::slide::types::Slide::parse_shape_enums(&ppdrawing.data)?;
    let shape = shapes
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let shape_id = native_shape_id(shape);
    let slide_offset = usize::try_from(entry.offset).map_err(|_error| {
        PackageError::ResourceLimit("PPT slide offset does not fit usize".into())
    })?;
    let atom = inspect_source_text_atom_parts(&slide_bytes, slide_offset, &slide, shape_id)?;

    shared.source_version().map_err(map_ole_error)?;
    Ok(SourceResolved {
        target,
        document_path: document_path.to_vec(),
        current_user_path: current_user_path.to_vec(),
        slide_persist_id: entry.persist_id,
        atom_offset: atom.offset,
        kind: atom.kind,
        payload: atom.payload,
        text: atom.text,
    })
}

fn read_current_user(reader: &StreamReader<'_>, limits: RecordLimits) -> Result<CurrentUser> {
    if reader.length < 28 {
        return Err(PackageError::Corrupted("CurrentUser stream too short".into()).into());
    }
    let mut header = [0_u8; 28];
    reader.read_into(0, &mut header, "CurrentUser")?;
    let stream_len = usize::try_from(reader.length).map_err(|_error| {
        PackageError::ResourceLimit("CurrentUser length does not fit usize".into())
    })?;
    let prefix_len = CurrentUser::source_prefix_len(&header, stream_len)?;
    let prefix = if prefix_len == header.len() {
        return CurrentUser::parse_with_limits(&header, limits).map_err(Into::into);
    } else {
        reader.read_vec(0, prefix_len, "CurrentUser prefix")?
    };
    CurrentUser::parse_with_limits(&prefix, limits).map_err(Into::into)
}

fn read_live_mapping(
    reader: &StreamReader<'_>,
    mut edit_offset: u32,
    limits: RecordLimits,
    session: &mut RecordParseSession,
) -> Result<LiveMapping> {
    let mut mappings = HashMap::new();
    let mut seen = HashSet::new();
    let mut expanded_entries = 0_usize;
    let mut document_persist_id = 0_u32;
    let mut first_edit = true;

    while edit_offset != 0 {
        if !seen.insert(edit_offset) || seen.len() > USER_EDIT_CHAIN_LIMIT {
            return Err(
                PackageError::Corrupted("cyclic or excessive UserEdit chain".into()).into(),
            );
        }
        let edit_offset_u64 = u64::from(edit_offset);
        let header = reader.header(edit_offset_u64, "UserEditAtom")?;
        if header.record_type != USER_EDIT_ATOM
            || header.version != 0
            || header.instance != 0
            || !matches!(header.data_len, 28 | 32)
        {
            let message = if first_edit {
                "CurrentUser does not reference a valid UserEditAtom"
            } else {
                "invalid historical UserEditAtom header"
            };
            return Err(PackageError::Corrupted(message.into()).into());
        }
        let record = reader.record(edit_offset_u64, header, "UserEditAtom", limits)?;
        session.account_materialized_record(0, header.data_len)?;
        if first_edit {
            document_persist_id = read_u32(&record, 24, "docPersistIdRef")?;
            if document_persist_id == 0 {
                return Err(PackageError::Corrupted(
                    "UserEditAtom has a null docPersistIdRef".into(),
                )
                .into());
            }
        }
        let directory_offset = read_u32(&record, 20, "persistDirectoryOffset")?;
        let directory_offset_u64 = u64::from(directory_offset);
        if directory_offset_u64 >= edit_offset_u64 {
            return Err(PackageError::Corrupted(
                "PersistDirectoryAtom does not precede its UserEditAtom".into(),
            )
            .into());
        }
        let directory_header = reader.header(directory_offset_u64, "PersistDirectoryAtom")?;
        if !matches!(
            directory_header.record_type,
            PERSIST_FULL | PERSIST_INCREMENTAL
        ) || directory_header.version != 0
            || directory_header.instance != 0
            || !directory_header.data_len.is_multiple_of(4)
        {
            return Err(PackageError::Corrupted(
                "invalid PersistDirectoryAtom header or shape".into(),
            )
            .into());
        }
        let directory_end = directory_offset_u64
            .checked_add(u64::try_from(directory_header.total_len).map_err(|_error| {
                PackageError::ResourceLimit("PersistDirectoryAtom length does not fit u64".into())
            })?)
            .ok_or_else(|| {
                PackageError::Corrupted("PersistDirectoryAtom range overflows".into())
            })?;
        if directory_end > edit_offset_u64 {
            return Err(PackageError::Corrupted(
                "PersistDirectoryAtom range overlaps its UserEditAtom".into(),
            )
            .into());
        }
        let directory = reader.record(
            directory_offset_u64,
            directory_header,
            "PersistDirectoryAtom",
            limits,
        )?;
        session.account_materialized_record(0, directory_header.data_len)?;
        merge_directory(
            &mut mappings,
            &directory,
            directory_offset_u64,
            limits,
            &mut expanded_entries,
        )?;
        let previous_edit = read_u32(&record, 16, "previousUserEdit")?;
        if previous_edit != 0 && u64::from(previous_edit) > edit_offset_u64 {
            return Err(PackageError::Corrupted(
                "UserEdit chain does not point to an earlier record".into(),
            )
            .into());
        }
        edit_offset = previous_edit;
        first_edit = false;
    }

    if document_persist_id == 0 {
        return Err(PackageError::Corrupted("missing Document persist ID".into()).into());
    }
    let document_offset = u64::from(*mappings.get(&document_persist_id).ok_or_else(|| {
        PackageError::Corrupted("live Document persist mapping is missing".into())
    })?);
    Ok(LiveMapping {
        mappings,
        document_persist_id,
        document_offset,
    })
}

fn merge_directory(
    mapping: &mut HashMap<u32, u32>,
    directory: &[u8],
    directory_offset: u64,
    limits: RecordLimits,
    expanded_entries: &mut usize,
) -> Result<()> {
    let payload_len = directory
        .len()
        .checked_sub(HEADER_LEN)
        .ok_or_else(|| PackageError::Corrupted("PersistDirectoryAtom is truncated".into()))?;
    if !payload_len.is_multiple_of(4) {
        return Err(PackageError::Corrupted(
            "PersistDirectoryAtom payload is not aligned to 4 bytes".into(),
        )
        .into());
    }
    let mut directory_ids = HashSet::new();
    directory_ids
        .try_reserve((payload_len / 8).min(limits.max_records))
        .map_err(|_error| PackageError::AllocationFailed("PPT persist directory IDs"))?;
    let mut offset = HEADER_LEN;
    while offset < directory.len() {
        let info = read_u32(directory, offset, "PersistDirectory run")?;
        offset = offset.checked_add(4).ok_or_else(|| {
            PackageError::Corrupted("PersistDirectory run offset overflows".into())
        })?;
        let base = info & 0x000F_FFFF;
        let count = info >> 20;
        if count == 0 {
            return Err(PackageError::Corrupted("zero persist run".into()).into());
        }
        let count = usize::try_from(count).map_err(|_error| {
            PackageError::ResourceLimit("PersistDirectory run length does not fit usize".into())
        })?;
        let expanded = expanded_entries.checked_add(count).ok_or_else(|| {
            PackageError::ResourceLimit("PPT persist mapping count overflows".into())
        })?;
        if expanded > limits.max_records {
            return Err(PackageError::ResourceLimit(format!(
                "PPT persist mapping count {expanded} exceeds limit {}",
                limits.max_records
            ))
            .into());
        }
        mapping
            .try_reserve(count)
            .map_err(|_error| PackageError::AllocationFailed("PPT persist mapping"))?;
        directory_ids
            .try_reserve(count)
            .map_err(|_error| PackageError::AllocationFailed("PPT persist directory IDs"))?;
        for index in 0..count {
            let value = read_u32(directory, offset, "PersistDirectory offset")?;
            offset = offset.checked_add(4).ok_or_else(|| {
                PackageError::Corrupted("PersistDirectory offset overflows".into())
            })?;
            let persist_id = base
                .checked_add(u32::try_from(index).map_err(|_error| {
                    PackageError::ResourceLimit(
                        "PersistDirectory persist ID index does not fit u32".into(),
                    )
                })?)
                .ok_or_else(|| {
                    PackageError::Corrupted("PersistDirectory persist ID overflows".into())
                })?;
            if !directory_ids.insert(persist_id) {
                return Err(PackageError::Corrupted(format!(
                    "duplicate persist identifier {persist_id} in PersistDirectoryAtom"
                ))
                .into());
            }
            if u64::from(value) >= directory_offset {
                return Err(PackageError::Corrupted(format!(
                    "PersistDirectory offset {value} does not precede its directory"
                ))
                .into());
            }
            mapping.entry(persist_id).or_insert(value);
        }
        *expanded_entries = expanded;
    }
    Ok(())
}

fn build_slide_index(
    reader: &StreamReader<'_>,
    mapping: &LiveMapping,
    limits: RecordLimits,
    session: &mut RecordParseSession,
) -> Result<SlideIndex> {
    let document_header = reader.header(mapping.document_offset, "DocumentContainer")?;
    if document_header.record_type != RecordType::Document as u16
        || document_header.version != 0x0F
        || document_header.instance != 0
    {
        return Err(PackageError::Corrupted(format!(
            "persist ID {} does not resolve to a DocumentContainer",
            mapping.document_persist_id
        ))
        .into());
    }
    session.account_existing_header(0, document_header.data_len)?;
    let document_end = mapping
        .document_offset
        .checked_add(u64::try_from(document_header.total_len).map_err(|_error| {
            PackageError::ResourceLimit("DocumentContainer length does not fit u64".into())
        })?)
        .ok_or_else(|| PackageError::Corrupted("DocumentContainer range overflows".into()))?;
    let payload_start = mapping
        .document_offset
        .checked_add(HEADER_LEN as u64)
        .ok_or_else(|| PackageError::Corrupted("DocumentContainer payload overflows".into()))?;

    let mut slide_list = None;
    let mut owners = super::LiveMacroOwners::default();
    walk_document_children(
        reader,
        payload_start,
        document_end,
        1,
        limits,
        session,
        &mut owners,
        &mut slide_list,
    )?;

    let mut entries = Vec::new();
    let mut slide_ids = HashSet::new();
    let mut persist_ids = HashSet::new();
    if let Some(slide_list) = slide_list {
        if slide_list.version != 0x0F {
            return Err(PackageError::Corrupted(
                "presentation SlideListWithTextContainer has invalid version".into(),
            )
            .into());
        }
        entries
            .try_reserve(slide_list.children.len())
            .map_err(|_error| PackageError::AllocationFailed("PPT slide directory"))?;
        slide_ids
            .try_reserve(slide_list.children.len())
            .map_err(|_error| PackageError::AllocationFailed("PPT slide ID index"))?;
        persist_ids
            .try_reserve(slide_list.children.len())
            .map_err(|_error| PackageError::AllocationFailed("PPT slide persist index"))?;
        for child in &slide_list.children {
            if child.record_type != RecordType::SlidePersistAtom {
                // The owned resolver accumulates optional list text here. The
                // source-backed text owner does not expose that projection,
                // but still executes the same tolerant extraction path so a
                // malformed optional text record cannot alter ownership.
                let _ = child.extract_text();
                continue;
            }
            if child.version != 0
                || child.instance != 0
                || child.data.len() != SLIDE_PERSIST_ATOM_SIZE
            {
                return Err(PackageError::Corrupted(format!(
                    "invalid SlidePersistAtom header or length: version={}, instance={}, length={}",
                    child.version,
                    child.instance,
                    child.data.len()
                ))
                .into());
            }
            let persist_id = read_u32(&child.data, 0, "SlidePersistAtom.persistIdRef")?;
            let slide_id = read_u32(&child.data, 12, "SlidePersistAtom.slideId")?;
            let text_placeholder_count = read_u32(&child.data, 8, "SlidePersistAtom.cTexts")?;
            if persist_id == 0 || slide_id == 0 {
                return Err(PackageError::Corrupted(
                    "SlidePersistAtom has a null persistIdRef or slideId".into(),
                )
                .into());
            }
            if text_placeholder_count > 8 {
                return Err(PackageError::Corrupted(format!(
                    "SlidePersistAtom cTexts exceeds 8: {text_placeholder_count}"
                ))
                .into());
            }
            if !slide_ids.insert(slide_id) {
                return Err(PackageError::Corrupted(format!(
                    "duplicate presentation slideId {slide_id}"
                ))
                .into());
            }
            if !persist_ids.insert(persist_id) {
                return Err(PackageError::Corrupted(format!(
                    "duplicate presentation persistIdRef {persist_id}"
                ))
                .into());
            }
            let slide_offset = mapping.mappings.get(&persist_id).copied().ok_or_else(|| {
                PackageError::Corrupted(format!(
                    "slide persist ID {persist_id} has no directory entry"
                ))
            })?;
            let slide_offset = u64::from(slide_offset);
            let slide_header = reader.header(slide_offset, "SlideContainer")?;
            if slide_header.record_type != SLIDE_CONTAINER
                || slide_header.version != 0x0F
                || slide_header.instance != 0
            {
                return Err(PackageError::Corrupted(format!(
                    "persist ID {persist_id} does not resolve to a SlideContainer"
                ))
                .into());
            }
            entries.push(SlideEntry {
                persist_id,
                offset: slide_offset,
            });
        }
        for set in slide_list.group_into_slide_atoms_sets() {
            let slide_id = read_u32(&set.slide_persist_atom.data, 12, "SlidePersistAtom.slideId")?;
            if !slide_ids.contains(&slide_id) {
                return Err(PackageError::Corrupted(format!(
                    "text records reference unknown slideId {slide_id}"
                ))
                .into());
            }
            let _ = set.text_interactions()?;
            let _ = set.outline_text_refs()?;
        }
    }

    // Keep the caller-selected limits in the signature and ensure the source
    // stream check remains explicit even when no SlideList is present.
    if reader.length
        > u64::try_from(limits.max_input_bytes).map_err(|_error| {
            PackageError::ResourceLimit("PPT input limit does not fit u64".into())
        })?
    {
        return Err(PackageError::ResourceLimit(format!(
            "PowerPoint Document stream size {} exceeds limit {}",
            reader.length, limits.max_input_bytes
        ))
        .into());
    }
    Ok(SlideIndex { entries })
}

fn walk_document_children(
    reader: &StreamReader<'_>,
    start: u64,
    end: u64,
    depth: usize,
    limits: RecordLimits,
    session: &mut RecordParseSession,
    owners: &mut super::LiveMacroOwners,
    slide_list: &mut Option<Record>,
) -> Result<()> {
    let mut offset = start;
    while offset < end {
        let remaining = end
            .checked_sub(offset)
            .ok_or_else(|| PackageError::Corrupted("Document child range underflows".into()))?;
        if remaining < HEADER_LEN as u64 {
            return Err(PackageError::Corrupted(
                "DocumentContainer ends with a truncated record header".into(),
            )
            .into());
        }
        let header = reader.header(offset, "Document child")?;
        let child_end = offset
            .checked_add(u64::try_from(header.total_len).map_err(|_error| {
                PackageError::ResourceLimit("Document child length does not fit u64".into())
            })?)
            .ok_or_else(|| PackageError::Corrupted("Document child range overflows".into()))?;
        if child_end > end {
            return Err(PackageError::Corrupted(
                "Document child extends beyond DocumentContainer".into(),
            )
            .into());
        }

        let record_type = RecordType::from(header.record_type);
        let parse_whole = header.record_type == RecordType::DocInfoList as u16
            || header.record_type == RecordType::VBAInfo as u16
            || (depth == 1
                && header.record_type == RecordType::SlideListWithText as u16
                && header.instance == SLIDE_LIST_INSTANCE);
        if parse_whole {
            if depth == 1
                && header.record_type == RecordType::SlideListWithText as u16
                && header.instance == SLIDE_LIST_INSTANCE
            {
                if slide_list.is_some() {
                    return Err(PackageError::Corrupted(
                        "duplicate presentation SlideListWithTextContainer".into(),
                    )
                    .into());
                }
            }
            let bytes = reader.record(offset, header, "Document metadata record", limits)?;
            let (record, consumed) = session.parse_strict_record_at_depth(&bytes, 0, depth)?;
            if consumed != bytes.len() {
                return Err(PackageError::Corrupted(
                    "Document metadata record has trailing bytes".into(),
                )
                .into());
            }
            inspect_live_macro_records(&record, false, false, owners)?;
            if depth == 1
                && header.record_type == RecordType::SlideListWithText as u16
                && header.instance == SLIDE_LIST_INSTANCE
            {
                *slide_list = Some(record);
            }
        } else if Record::is_container_record(record_type) {
            session.account_existing_header(depth, header.data_len)?;
            let child_start = offset.checked_add(HEADER_LEN as u64).ok_or_else(|| {
                PackageError::Corrupted("Document child payload overflows".into())
            })?;
            walk_document_children(
                reader,
                child_start,
                child_end,
                depth
                    .checked_add(1)
                    .ok_or_else(|| PackageError::Corrupted("PPT record depth overflows".into()))?,
                limits,
                session,
                owners,
                slide_list,
            )?;
        } else {
            session.account_existing_header(depth, header.data_len)?;
        }
        offset = child_end;
    }
    Ok(())
}

fn read_u32(data: &[u8], offset: usize, label: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| PackageError::Corrupted(format!("{label} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn raw_record(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut record = Vec::with_capacity(HEADER_LEN + payload.len());
        record.extend_from_slice(&0_u16.to_le_bytes());
        record.extend_from_slice(&record_type.to_le_bytes());
        record.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        record.extend_from_slice(payload);
        record
    }

    #[test]
    fn persist_mapping_expansion_is_bounded_at_exact_and_one_under_limits() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&((2_u32 << 20) | 10).to_le_bytes());
        payload.extend_from_slice(&0_u32.to_le_bytes());
        payload.extend_from_slice(&4_u32.to_le_bytes());
        let directory = raw_record(PERSIST_INCREMENTAL, &payload);

        let mut exact_mapping = HashMap::new();
        let mut exact_expanded = 0;
        merge_directory(
            &mut exact_mapping,
            &directory,
            100,
            RecordLimits {
                max_records: 2,
                ..RecordLimits::default()
            },
            &mut exact_expanded,
        )
        .unwrap();
        assert_eq!(exact_mapping.len(), 2);
        assert_eq!(exact_expanded, 2);

        let mut one_under_mapping = HashMap::new();
        let mut one_under_expanded = 0;
        assert!(matches!(
            merge_directory(
                &mut one_under_mapping,
                &directory,
                100,
                RecordLimits {
                    max_records: 1,
                    ..RecordLimits::default()
                },
                &mut one_under_expanded,
            ),
            Err(Error::Package(PackageError::ResourceLimit(message)))
                if message.contains("persist mapping count")
        ));
    }

    #[test]
    fn persist_mapping_rejects_duplicate_ids_within_one_directory() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&((1_u32 << 20) | 7).to_le_bytes());
        payload.extend_from_slice(&0_u32.to_le_bytes());
        payload.extend_from_slice(&((1_u32 << 20) | 7).to_le_bytes());
        payload.extend_from_slice(&4_u32.to_le_bytes());
        let directory = raw_record(PERSIST_INCREMENTAL, &payload);
        let mut mapping = HashMap::new();
        let mut expanded = 0;

        assert!(matches!(
            merge_directory(
                &mut mapping,
                &directory,
                100,
                RecordLimits::default(),
                &mut expanded,
            ),
            Err(Error::Package(PackageError::Corrupted(message)))
                if message.contains("duplicate persist identifier")
        ));
    }

    #[test]
    fn persist_mapping_keeps_newest_wins_across_historical_directories() {
        let mut current_payload = Vec::new();
        current_payload.extend_from_slice(&((1_u32 << 20) | 7).to_le_bytes());
        current_payload.extend_from_slice(&10_u32.to_le_bytes());
        let current = raw_record(PERSIST_INCREMENTAL, &current_payload);
        let mut historical_payload = Vec::new();
        historical_payload.extend_from_slice(&((1_u32 << 20) | 7).to_le_bytes());
        historical_payload.extend_from_slice(&20_u32.to_le_bytes());
        let historical = raw_record(PERSIST_INCREMENTAL, &historical_payload);
        let mut mapping = HashMap::new();
        let mut expanded = 0;

        merge_directory(
            &mut mapping,
            &current,
            100,
            RecordLimits::default(),
            &mut expanded,
        )
        .unwrap();
        merge_directory(
            &mut mapping,
            &historical,
            100,
            RecordLimits::default(),
            &mut expanded,
        )
        .unwrap();
        assert_eq!(mapping.get(&7), Some(&10));
        assert_eq!(expanded, 2);
    }
}

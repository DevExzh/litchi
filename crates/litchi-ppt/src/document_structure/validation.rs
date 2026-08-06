//! Structural validation for the legacy PPT document owner.

use std::collections::HashSet;

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

use super::model::{CustomTableStylesPlacement, DocumentStructure, Limits, Master, Slide};

const DOCUMENT_ATOM_PAYLOAD: usize = 40;
const PERSIST_ATOM_PAYLOAD: usize = 20;

/// Validate one complete document tree and build its typed structural view.
pub(super) fn validate_document(document: &Record, limits: Limits) -> Result<DocumentStructure> {
    let limits = limits.validate()?;
    validate_budget(document, 0, &mut 0usize, limits)?;
    header(document, RecordType::Document, 0x0f, 0)?;

    let mut document_atom = None;
    let mut master_list = None;
    let mut slide_list = None;
    let mut notes_list = None;
    let mut end_document = Vec::new();
    let mut custom_styles = Vec::new();

    let mut previous_rank = None;
    for (index, child) in document.children.iter().enumerate() {
        match child.record_type {
            RecordType::DocumentAtom => {
                if document_atom.replace(index).is_some() {
                    return corrupted("DocumentContainer contains duplicate DocumentAtom records");
                }
                atom_header(child, RecordType::DocumentAtom, 1, 0, DOCUMENT_ATOM_PAYLOAD)?;
                check_order(&mut previous_rank, 0, "DocumentAtom")?;
            },
            RecordType::SlideListWithText => {
                list_header(child)?;
                let slot = match child.instance {
                    0 => &mut slide_list,
                    1 => &mut master_list,
                    2 => &mut notes_list,
                    instance => {
                        return corrupted(format!(
                            "DocumentContainer has an invalid SlideListWithText instance {instance}"
                        ));
                    },
                };
                if slot.replace(index).is_some() {
                    return corrupted(
                        "DocumentContainer contains duplicate SlideListWithText containers",
                    );
                }
                check_order(
                    &mut previous_rank,
                    match child.instance {
                        1 => 10,
                        0 => 20,
                        _ => 30,
                    },
                    "SlideListWithText",
                )?;
            },
            RecordType::EndDocument => {
                atom_header(child, RecordType::EndDocument, 0, 0, 0)?;
                end_document.push(index);
            },
            RecordType::RoundTripCustomTableStyles12Atom => {
                if child.instance != 0 || !child.children.is_empty() {
                    return corrupted(
                        "RoundTripCustomTableStyles12Atom has an invalid record header",
                    );
                }
                if usize::try_from(child.data_length).ok() != Some(child.data.len()) {
                    return corrupted(
                        "RoundTripCustomTableStyles12Atom length does not match its payload",
                    );
                }
                custom_styles.push(index);
            },
            _ => {},
        }
    }

    let end_index = match end_document.as_slice() {
        [index] => *index,
        [] => return corrupted("DocumentContainer must contain exactly one EndDocumentAtom"),
        _ => return corrupted("DocumentContainer must contain exactly one EndDocumentAtom"),
    };
    let custom_table_styles = match custom_styles.as_slice() {
        [] => {
            if end_index.checked_add(1) != Some(document.children.len()) {
                return corrupted(
                    "DocumentContainer has records after EndDocumentAtom without custom table styles",
                );
            }
            None
        },
        [style_index]
            if style_index.checked_add(1) == Some(end_index)
                && end_index.checked_add(1) == Some(document.children.len()) =>
        {
            Some(CustomTableStylesPlacement::BeforeEndDocument)
        },
        [style_index]
            if end_index.checked_add(1) == Some(*style_index)
                && style_index.checked_add(1) == Some(document.children.len()) =>
        {
            Some(CustomTableStylesPlacement::AfterEndDocument)
        },
        _ => {
            return corrupted(
                "EndDocumentAtom and optional custom table styles do not form the document tail",
            );
        },
    };

    let mut masters = Vec::new();
    let mut slides = Vec::new();
    let mut persist_ids = HashSet::new();
    if let Some(index) = master_list {
        parse_master_list(&document.children[index], &mut masters, &mut persist_ids)?;
    }
    if let Some(index) = slide_list {
        parse_slide_list(&document.children[index], &mut slides, &mut persist_ids)?;
    }

    let mut master_ids = HashSet::with_capacity(masters.len());
    for master in &masters {
        if !master_ids.insert(master.master_id()) {
            return corrupted(format!(
                "master list contains duplicate master identifier {:#010x}",
                master.master_id()
            ));
        }
    }
    let mut slide_ids = HashSet::with_capacity(slides.len());
    for slide in &slides {
        if !slide_ids.insert(slide.slide_id()) {
            return corrupted(format!(
                "slide list contains duplicate slide identifier {}",
                slide.slide_id()
            ));
        }
    }

    let document_atom_child_index = document_atom.ok_or_else(|| {
        Error::Corrupted("DocumentContainer is missing its required DocumentAtom".into())
    })?;

    Ok(DocumentStructure::new(
        end_index,
        custom_table_styles,
        document_atom_child_index,
        master_list,
        slide_list,
        notes_list,
        masters,
        slides,
    ))
}

fn parse_master_list(
    list: &Record,
    masters: &mut Vec<Master>,
    persist_ids: &mut HashSet<u32>,
) -> Result<()> {
    for (index, child) in list.children.iter().enumerate() {
        if child.record_type != RecordType::SlidePersistAtom {
            continue;
        }
        let (persist_id, flags, master_id) = parse_master_persist(child)?;
        if !persist_ids.insert(persist_id) {
            return corrupted(format!(
                "document structure reuses persist identifier {persist_id}"
            ));
        }
        masters.push(Master::new(index, persist_id, master_id, flags));
    }
    Ok(())
}

fn parse_slide_list(
    list: &Record,
    slides: &mut Vec<Slide>,
    persist_ids: &mut HashSet<u32>,
) -> Result<()> {
    let mut current: Option<(u32, usize)> = None;
    for (index, child) in list.children.iter().enumerate() {
        if child.record_type == RecordType::SlidePersistAtom {
            if let Some((text_count, seen_texts)) = current.take() {
                if seen_texts > usize::try_from(text_count).unwrap_or(usize::MAX) {
                    return corrupted("slide list contains more text headers than cTexts");
                }
            }
            let (persist_id, flags, text_count, slide_id) = parse_slide_persist(child)?;
            if !persist_ids.insert(persist_id) {
                return corrupted(format!(
                    "document structure reuses persist identifier {persist_id}"
                ));
            }
            slides.push(Slide::new(index, persist_id, slide_id, flags, text_count));
            current = Some((text_count, 0));
            continue;
        }

        if child.record_type == RecordType::TextHeaderAtom {
            if let Some((_, seen_texts)) = current.as_mut() {
                *seen_texts = seen_texts.saturating_add(1);
            } else {
                return corrupted("slide-list text records precede their SlidePersistAtom");
            }
        } else if is_slide_text_record(child.record_type) && current.is_none() {
            return corrupted("slide-list text records precede their SlidePersistAtom");
        }
    }
    if let Some((text_count, seen_texts)) = current {
        if seen_texts > usize::try_from(text_count).unwrap_or(usize::MAX) {
            return corrupted("slide list contains more text headers than cTexts");
        }
    }
    Ok(())
}

fn parse_master_persist(record: &Record) -> Result<(u32, u32, u32)> {
    atom_header(
        record,
        RecordType::SlidePersistAtom,
        0,
        0,
        PERSIST_ATOM_PAYLOAD,
    )?;
    let persist_id = read_u32(record, 0, "MasterPersistAtom.persistIdRef")?;
    let flags = read_u32(record, 4, "MasterPersistAtom.flags")?;
    let reserved = read_u32(record, 8, "MasterPersistAtom.reserved3")?;
    let master_id = read_u32(record, 12, "MasterPersistAtom.masterId")?;
    let reserved2 = read_u32(record, 16, "MasterPersistAtom.reserved4")?;
    if persist_id == 0 || master_id < 0x8000_0000 {
        return corrupted("MasterPersistAtom has an invalid persist or master identifier");
    }
    if flags & !0x0000_0004 != 0 || reserved != 0 || reserved2 != 0 {
        return corrupted("MasterPersistAtom contains nonzero reserved fields");
    }
    Ok((persist_id, flags, master_id))
}

fn parse_slide_persist(record: &Record) -> Result<(u32, u32, u32, u32)> {
    atom_header(
        record,
        RecordType::SlidePersistAtom,
        0,
        0,
        PERSIST_ATOM_PAYLOAD,
    )?;
    let persist_id = read_u32(record, 0, "SlidePersistAtom.persistIdRef")?;
    let flags = read_u32(record, 4, "SlidePersistAtom.flags")?;
    let text_count = read_u32(record, 8, "SlidePersistAtom.cTexts")?;
    let slide_id = read_u32(record, 12, "SlidePersistAtom.slideId")?;
    let reserved = read_u32(record, 16, "SlidePersistAtom.reserved3")?;
    if persist_id == 0 || slide_id == 0 || text_count > 8 || flags & !0x0000_0006 != 0 {
        return corrupted("SlidePersistAtom contains an invalid identifier, flag, or cTexts value");
    }
    if reserved != 0 {
        return corrupted("SlidePersistAtom contains a nonzero reserved field");
    }
    Ok((persist_id, flags, text_count, slide_id))
}

fn is_slide_text_record(record_type: RecordType) -> bool {
    matches!(
        record_type,
        RecordType::TextHeaderAtom
            | RecordType::TextCharsAtom
            | RecordType::TextBytesAtom
            | RecordType::StyleTextPropAtom
            | RecordType::TextBookmarkAtom
            | RecordType::TextSpecInfoAtom
            | RecordType::InteractiveInfo
            | RecordType::TextInteractiveInfoAtom
            | RecordType::SlideNumberMCAtom
            | RecordType::DateTimeMCAtom
            | RecordType::GenericDateMCAtom
            | RecordType::HeaderMCAtom
            | RecordType::FooterMCAtom
            | RecordType::RtfDateTimeMCAtom
    )
}

fn check_order(previous: &mut Option<u8>, rank: u8, name: &str) -> Result<()> {
    if previous.is_some_and(|old| rank < old) {
        return corrupted(format!("DocumentContainer {name} is out of order"));
    }
    *previous = Some(rank);
    Ok(())
}

fn header(record: &Record, kind: RecordType, version: u16, instance: u16) -> Result<()> {
    if record.record_type != kind
        || record.record_type_raw != kind.as_u16()
        || record.version != version
        || record.instance != instance
    {
        return corrupted(format!("{kind:?} has an invalid record header"));
    }
    Ok(())
}

fn list_header(record: &Record) -> Result<()> {
    header(record, RecordType::SlideListWithText, 0x0f, record.instance)?;
    if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
        return corrupted("SlideListWithText length does not match its payload");
    }
    Ok(())
}

fn atom_header(
    record: &Record,
    kind: RecordType,
    version: u16,
    instance: u16,
    payload_len: usize,
) -> Result<()> {
    header(record, kind, version, instance)?;
    if record.children.is_empty()
        && usize::try_from(record.data_length).ok() == Some(record.data.len())
        && record.data.len() == payload_len
    {
        return Ok(());
    }
    if record.data.len() != payload_len
        || usize::try_from(record.data_length).ok() != Some(payload_len)
        || !record.children.is_empty()
    {
        return corrupted(format!("{kind:?} has an invalid payload or length"));
    }
    Ok(())
}

fn read_u32(record: &Record, offset: usize, name: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| Error::Corrupted(format!("{name} offset overflow")))?;
    let bytes = record
        .data
        .get(offset..end)
        .ok_or_else(|| Error::Corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::Corrupted(format!("{name} is truncated"))
    })?))
}

fn validate_budget(record: &Record, depth: usize, count: &mut usize, limits: Limits) -> Result<()> {
    if depth > limits.max_depth {
        return Err(Error::InvalidFormat(
            "document-structure record nesting exceeds its limit".into(),
        ));
    }
    *count = count
        .checked_add(1)
        .ok_or_else(|| Error::Corrupted("document-structure record count overflow".into()))?;
    if *count > limits.max_records {
        return Err(Error::InvalidFormat(
            "document-structure record count exceeds its limit".into(),
        ));
    }
    if record.data.len() > limits.max_bytes
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return Err(Error::Corrupted(
            "document-structure record length does not match its payload".into(),
        ));
    }
    for child in &record.children {
        validate_budget(child, depth.saturating_add(1), count, limits)?;
    }
    Ok(())
}

pub(super) fn list_index(document: &Record, instance: u16) -> Result<Option<usize>> {
    let mut found = None;
    for (index, child) in document.children.iter().enumerate() {
        if child.record_type == RecordType::SlideListWithText && child.instance == instance {
            if found.replace(index).is_some() {
                return corrupted("DocumentContainer contains duplicate structural lists");
            }
        }
    }
    Ok(found)
}

pub(super) fn is_top_level_known(record: &Record) -> bool {
    matches!(
        record.record_type,
        RecordType::DocumentAtom
            | RecordType::SlideListWithText
            | RecordType::EndDocument
            | RecordType::RoundTripCustomTableStyles12Atom
    )
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

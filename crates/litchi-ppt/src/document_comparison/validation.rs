//! Structural and resource validation for document-comparison records.

use crate::package::{Error, Result};

use super::model::{DiffNode, DiffType};

pub(super) const MAX_SLIDE_LIST_ENTRIES: usize = 1_000_000;
pub(super) const MAX_REVIEWER_NAME_BYTES: usize = 104;

pub(super) fn validate_diff_children(parent: DiffType, children: &[DiffNode]) -> Result<()> {
    use DiffType as T;
    match parent {
        T::NamedShowList => require_repeated(children, &[T::NamedShow]),
        T::MasterList => require_repeated(children, &[T::MainMaster, T::Slide]),
        T::SlideList => require_repeated(children, &[T::Slide]),
        T::ShapeList => require_repeated(children, &[T::Shape]),
        T::TableList => require_repeated(children, &[T::Table]),
        T::Document => require_ordered(
            children,
            &[
                (T::HeaderFooter, Some(true)),
                (T::HeaderFooter, Some(false)),
                (T::NamedShowList, None),
                (T::MasterList, None),
                (T::SlideList, None),
            ],
        ),
        T::MainMaster => require_ordered(
            children,
            &[(T::ShapeList, None), (T::TableList, None), (T::Notes, None)],
        ),
        T::Slide => require_ordered(
            children,
            &[
                (T::ShapeList, None),
                (T::TableList, None),
                (T::SlideShow, None),
                (T::HeaderFooter, Some(true)),
                (T::Notes, None),
            ],
        ),
        T::Shape => require_ordered(
            children,
            &[
                (T::Text, None),
                (T::RecolorInfo, None),
                (T::ExternalObject, None),
                (T::InteractiveInfo, Some(true)),
                (T::InteractiveInfo, Some(false)),
            ],
        ),
        _ if children.is_empty() => Ok(()),
        _ => corrupted("leaf Diff10 record contains child records"),
    }
}

fn require_repeated(children: &[DiffNode], allowed: &[DiffType]) -> Result<()> {
    if children
        .iter()
        .all(|child| allowed.contains(&child.diff_type()))
    {
        Ok(())
    } else {
        corrupted("Diff10 list contains a child of the wrong type")
    }
}

fn require_ordered(children: &[DiffNode], grammar: &[(DiffType, Option<bool>)]) -> Result<()> {
    let mut previous = None;
    for child in children {
        let rank = grammar
            .iter()
            .position(|(diff_type, index)| {
                child.diff_type() == *diff_type
                    && index.is_none_or(|value| child.headers.index == value)
            })
            .ok_or_else(|| {
                Error::Corrupted("Diff10 container contains a child of the wrong type".to_string())
            })?;
        if previous.is_some_and(|value| rank <= value) {
            return corrupted("Diff10 children are duplicated or out of order");
        }
        previous = Some(rank);
    }
    Ok(())
}

pub(super) fn validate_reviewer_name(name: &str) -> Result<()> {
    if name.encode_utf16().count() * 2 > MAX_REVIEWER_NAME_BYTES {
        return corrupted("reviewer name exceeds 104 bytes");
    }
    if name.chars().any(|character| {
        let value = character as u32;
        value == 0 || value <= 0x1F || (0x7F..=0x9F).contains(&value)
    }) {
        return corrupted("reviewer name is not a PrintableUnicodeString");
    }
    Ok(())
}

pub(super) fn validate_count(count: usize) -> Result<()> {
    if count > MAX_SLIDE_LIST_ENTRIES {
        return corrupted("slide-list table count exceeds the MS-PPT limit");
    }
    Ok(())
}

pub(super) fn validate_atom(
    record: &crate::records::Record,
    kind: crate::consts::RecordType,
    payload_size: usize,
) -> Result<()> {
    if record.record_type != kind
        || record.version != 0
        || record.instance != 0
        || record.data.len() != payload_size
        || usize::try_from(record.data_length).ok() != Some(payload_size)
        || !record.children.is_empty()
    {
        return corrupted("PowerPoint document-comparison atom has an invalid header or length");
    }
    Ok(())
}

impl DiffNode {
    pub(super) fn validate_node(&self) -> Result<()> {
        if self.headers.index
            && !matches!(
                self.headers.diff_type,
                DiffType::HeaderFooter | DiffType::InteractiveInfo
            )
        {
            return corrupted("Diff10 fIndex is invalid for its diff type");
        }
        if let Some(flag_type) = self.flags.diff_type() {
            if flag_type != self.headers.diff_type {
                return corrupted("Diff10 flags do not match the record tag");
            }
        } else if matches!(
            self.headers.diff_type,
            DiffType::Document
                | DiffType::Slide
                | DiffType::MainMaster
                | DiffType::Shape
                | DiffType::Table
                | DiffType::Text
                | DiffType::Notes
        ) {
            return corrupted("Diff10 record is missing its typed flags");
        }
        validate_diff_children(self.headers.diff_type, &self.children)
    }
}

pub(super) fn corrupted<T>(message: &str) -> Result<T> {
    Err(Error::Corrupted(message.to_string()))
}

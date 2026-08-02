//! Paragraph-scoped list-style variation boundaries.

use crate::{Error, Result};

use super::super::drop_cap::ParagraphStart;
use super::storage::ListBoundaryStorage;

pub(super) fn effective_style_id(
    storage: &ListBoundaryStorage,
    paragraph: ParagraphStart,
) -> Result<u64> {
    let target = paragraph.utf16_index();
    if storage.paragraph_starts.binary_search(&target).is_err() {
        return Err(Error::InvalidFormat(format!(
            "UTF-16 index {target} is not a paragraph start"
        )));
    }
    storage
        .boundaries
        .iter()
        .take_while(|entry| entry.0 <= target)
        .last()
        .map(|entry| entry.1)
        .ok_or_else(|| {
            Error::InvalidFormat("iWork text storage has no list-style boundary at zero".to_owned())
        })
}

pub(super) fn style_isolated_to_paragraph(
    storage: &ListBoundaryStorage,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let target = paragraph.utf16_index();
    let target_boundary = storage
        .boundaries
        .binary_search_by_key(&target, |entry| entry.0)
        .is_ok();
    if !target_boundary {
        return Ok(false);
    }
    let Some(next_paragraph) = storage
        .paragraph_starts
        .iter()
        .copied()
        .find(|start| *start > target)
    else {
        return Ok(true);
    };
    Ok(storage
        .boundaries
        .binary_search_by_key(&next_paragraph, |entry| entry.0)
        .is_ok())
}

pub(super) fn paragraph_boundaries_with_style(
    storage: &ListBoundaryStorage,
    paragraph: ParagraphStart,
    replacement_style_id: u64,
) -> Result<Vec<(u32, u64)>> {
    let target = paragraph.utf16_index();
    if storage.paragraph_starts.binary_search(&target).is_err() {
        return Err(Error::InvalidFormat(format!(
            "UTF-16 index {target} is not a paragraph start"
        )));
    }
    let mut result = Vec::with_capacity(storage.boundaries.len().saturating_add(2));
    for &start in &storage.paragraph_starts {
        let style_id = if start == target {
            replacement_style_id
        } else {
            storage
                .boundaries
                .iter()
                .take_while(|entry| entry.0 <= start)
                .last()
                .map(|entry| entry.1)
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "iWork text storage has no list-style boundary at zero".to_owned(),
                    )
                })?
        };
        if result
            .last()
            .is_none_or(|previous: &(u32, u64)| previous.1 != style_id)
        {
            result.push((start, style_id));
        }
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn storage() -> ListBoundaryStorage {
        ListBoundaryStorage {
            archive_name: "Index/Test.iwa".to_owned(),
            boundaries: vec![(0, 10), (20, 11)],
            paragraph_starts: vec![0, 10, 20],
        }
    }

    #[test]
    fn changing_one_paragraph_splits_and_restores_boundaries() {
        assert_eq!(
            paragraph_boundaries_with_style(
                &storage(),
                ParagraphStart::from_utf16_index(10).unwrap(),
                99,
            )
            .unwrap(),
            [(0, 10), (10, 99), (20, 11)]
        );
    }

    #[test]
    fn only_single_paragraph_style_spans_are_isolated() {
        let storage = storage();
        assert!(!style_isolated_to_paragraph(&storage, ParagraphStart::ZERO).unwrap());
        assert!(
            !style_isolated_to_paragraph(&storage, ParagraphStart::from_utf16_index(10).unwrap())
                .unwrap()
        );
        assert!(
            style_isolated_to_paragraph(&storage, ParagraphStart::from_utf16_index(20).unwrap())
                .unwrap()
        );
    }
}

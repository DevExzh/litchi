//! Structural validation for the terminal records of a PPT `DocumentContainer`.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Position of the optional PowerPoint 12 custom table-style package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomTableStylesPlacement {
    BeforeEndDocument,
    AfterEndDocument,
}

/// Strictly validated terminal structure of a `DocumentContainer`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentStructure {
    pub end_document_child_index: usize,
    pub custom_table_styles: Option<CustomTableStylesPlacement>,
}

impl DocumentStructure {
    /// Validate the exact MS-PPT document tail.
    pub(crate) fn parse(document: &Record) -> Result<Self> {
        if document.record_type != RecordType::Document
            || document.version != 0x0f
            || document.instance != 0
        {
            return corrupted("DocumentContainer has an invalid record header");
        }

        let end_indices: Vec<_> = document
            .children
            .iter()
            .enumerate()
            .filter_map(|(index, record)| {
                (record.record_type == RecordType::EndDocument).then_some(index)
            })
            .collect();
        if end_indices.len() != 1 {
            return corrupted("DocumentContainer must contain exactly one EndDocumentAtom");
        }
        let end_index = end_indices[0];
        let end = &document.children[end_index];
        if end.version != 0 || end.instance != 0 || end.data_length != 0 || !end.data.is_empty() {
            return corrupted("EndDocumentAtom has an invalid record header or payload");
        }

        let table_style_indices: Vec<_> = document
            .children
            .iter()
            .enumerate()
            .filter_map(|(index, record)| {
                (record.record_type == RecordType::RoundTripCustomTableStyles12Atom)
                    .then_some(index)
            })
            .collect();
        if table_style_indices.len() > 1 {
            return corrupted(
                "DocumentContainer contains duplicate RoundTripCustomTableStyles12Atom records",
            );
        }

        let child_count = document.children.len();
        let custom_table_styles = match table_style_indices.first().copied() {
            None if end_index.checked_add(1) == Some(child_count) => None,
            Some(table_index)
                if table_index.checked_add(1) == Some(end_index)
                    && end_index.checked_add(1) == Some(child_count) =>
            {
                Some(CustomTableStylesPlacement::BeforeEndDocument)
            },
            Some(table_index)
                if end_index.checked_add(1) == Some(table_index)
                    && table_index.checked_add(1) == Some(child_count) =>
            {
                Some(CustomTableStylesPlacement::AfterEndDocument)
            },
            _ => {
                return corrupted(
                    "EndDocumentAtom and optional custom table styles do not form the document tail",
                );
            },
        };

        Ok(Self {
            end_document_child_index: end_index,
            custom_table_styles,
        })
    }
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(Error::Corrupted(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(record_type: RecordType, version: u16, data: Vec<u8>) -> Record {
        Record {
            record_type,
            record_type_raw: record_type.as_u16(),
            version,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        }
    }

    fn document(children: Vec<Record>) -> Record {
        let mut value = record(RecordType::Document, 0x0f, Vec::new());
        value.children = children;
        value
    }

    #[test]
    fn accepts_both_defined_custom_table_style_placements() {
        let prefix = record(RecordType::DocumentAtom, 0, Vec::new());
        let end = record(RecordType::EndDocument, 0, Vec::new());
        let styles = record(RecordType::RoundTripCustomTableStyles12Atom, 0, Vec::new());
        assert_eq!(
            DocumentStructure::parse(&document(
                vec![prefix.clone(), styles.clone(), end.clone(),]
            ))
            .unwrap()
            .custom_table_styles,
            Some(CustomTableStylesPlacement::BeforeEndDocument)
        );
        assert_eq!(
            DocumentStructure::parse(&document(vec![prefix, end, styles]))
                .unwrap()
                .custom_table_styles,
            Some(CustomTableStylesPlacement::AfterEndDocument)
        );
    }

    #[test]
    fn rejects_missing_duplicate_nonempty_and_nonterminal_end_records() {
        let end = record(RecordType::EndDocument, 0, Vec::new());
        assert!(DocumentStructure::parse(&document(Vec::new())).is_err());
        assert!(DocumentStructure::parse(&document(vec![end.clone(), end.clone()])).is_err());
        assert!(
            DocumentStructure::parse(&document(
                vec![record(RecordType::EndDocument, 0, vec![0]),]
            ))
            .is_err()
        );
        assert!(
            DocumentStructure::parse(&document(vec![
                end,
                record(RecordType::DocumentAtom, 0, Vec::new()),
            ]))
            .is_err()
        );
    }
}

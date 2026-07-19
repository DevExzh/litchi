//! Structural validation for the terminal records of a PPT `DocumentContainer`.

use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;

/// Position of the optional PowerPoint 12 custom table-style package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointCustomTableStylesPlacement {
    BeforeEndDocument,
    AfterEndDocument,
}

/// Strictly validated terminal structure of a `DocumentContainer`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointDocumentStructure {
    pub end_document_child_index: usize,
    pub custom_table_styles: Option<PowerPointCustomTableStylesPlacement>,
}

impl PowerPointDocumentStructure {
    /// Validate the exact MS-PPT document tail.
    pub(crate) fn parse(document: &PptRecord) -> Result<Self> {
        if document.record_type != PptRecordType::Document
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
                (record.record_type == PptRecordType::EndDocument).then_some(index)
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
                (record.record_type == PptRecordType::RoundTripCustomTableStyles12Atom)
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
                Some(PowerPointCustomTableStylesPlacement::BeforeEndDocument)
            },
            Some(table_index)
                if end_index.checked_add(1) == Some(table_index)
                    && table_index.checked_add(1) == Some(child_count) =>
            {
                Some(PowerPointCustomTableStylesPlacement::AfterEndDocument)
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
    Err(PptError::Corrupted(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(record_type: PptRecordType, version: u16, data: Vec<u8>) -> PptRecord {
        PptRecord {
            record_type,
            record_type_raw: record_type.as_u16(),
            version,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children: Vec::new(),
        }
    }

    fn document(children: Vec<PptRecord>) -> PptRecord {
        let mut value = record(PptRecordType::Document, 0x0f, Vec::new());
        value.children = children;
        value
    }

    #[test]
    fn accepts_both_defined_custom_table_style_placements() {
        let prefix = record(PptRecordType::DocumentAtom, 0, Vec::new());
        let end = record(PptRecordType::EndDocument, 0, Vec::new());
        let styles = record(
            PptRecordType::RoundTripCustomTableStyles12Atom,
            0,
            Vec::new(),
        );
        assert_eq!(
            PowerPointDocumentStructure::parse(&document(vec![
                prefix.clone(),
                styles.clone(),
                end.clone(),
            ]))
            .unwrap()
            .custom_table_styles,
            Some(PowerPointCustomTableStylesPlacement::BeforeEndDocument)
        );
        assert_eq!(
            PowerPointDocumentStructure::parse(&document(vec![prefix, end, styles]))
                .unwrap()
                .custom_table_styles,
            Some(PowerPointCustomTableStylesPlacement::AfterEndDocument)
        );
    }

    #[test]
    fn rejects_missing_duplicate_nonempty_and_nonterminal_end_records() {
        let end = record(PptRecordType::EndDocument, 0, Vec::new());
        assert!(PowerPointDocumentStructure::parse(&document(Vec::new())).is_err());
        assert!(
            PowerPointDocumentStructure::parse(&document(vec![end.clone(), end.clone()])).is_err()
        );
        assert!(
            PowerPointDocumentStructure::parse(&document(vec![record(
                PptRecordType::EndDocument,
                0,
                vec![0]
            ),]))
            .is_err()
        );
        assert!(
            PowerPointDocumentStructure::parse(&document(vec![
                end,
                record(PptRecordType::DocumentAtom, 0, Vec::new()),
            ]))
            .is_err()
        );
    }
}

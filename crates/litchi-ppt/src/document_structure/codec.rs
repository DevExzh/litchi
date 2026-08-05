//! MS-PPT validation for a document container's terminal records.

use super::model::{CustomTableStylesPlacement, DocumentStructure};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

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

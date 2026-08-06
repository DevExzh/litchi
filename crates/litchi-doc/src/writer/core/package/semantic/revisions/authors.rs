use crate::writer::core::{Writer, model::*};
use crate::writer::fib::FibBuilder;
use std::collections::HashMap;
impl Writer {
    pub(in crate::writer::core::package) fn build_revision_writer_data(
        &self,
    ) -> Result<Option<RevisionWriterData>, WriteError> {
        let mut authors = vec!["Unknown".to_string()];
        let mut indexes = HashMap::from([("Unknown".to_string(), 0u16)]);
        let mut has_revisions = false;
        let mut index_author = |author: &str| -> Result<(), WriteError> {
            has_revisions = true;
            if !indexes.contains_key(author) {
                if authors.len() >= 0x8000 {
                    return Err(WriteError::InvalidData(
                        "DOC revision author table exceeds the signed author-index range"
                            .to_string(),
                    ));
                }
                let index = authors.len() as u16;
                authors.push(author.to_string());
                indexes.insert(author.to_string(), index);
            }
            Ok(())
        };
        if let Some(revision) = &self.section_formatting_revision {
            index_author(&revision.author)?;
        }
        for style in &self.styles {
            if let Some(revision) = &style.revision {
                index_author(&revision.author)?;
            }
        }
        let table_paragraphs = self.tables.iter().flat_map(|table| {
            table
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.paragraphs.iter())
        });
        for paragraph in self.paragraphs.iter().chain(table_paragraphs) {
            let mut formatting = Some(&paragraph.formatting);
            while let Some(current) = formatting {
                if let Some(revision) = &current.formatting_revision {
                    index_author(&revision.author)?;
                }
                if let Some(revision) = &current.numbering_revision {
                    index_author(&revision.author)?;
                }
                formatting = current.preserved_properties_for_revision.as_deref();
            }
            for run in &paragraph.runs {
                let mut formatting = Some(&run.formatting);
                while let Some(current) = formatting {
                    if let Some(revision) = &current.insertion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.deletion_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.formatting_revision {
                        index_author(&revision.author)?;
                    }
                    if let Some(revision) = &current.display_field_revision {
                        index_author(&revision.author)?;
                    }
                    formatting = current.preserved_properties_for_revision.as_deref();
                }
            }
        }
        if !has_revisions {
            return Ok(None);
        }

        let mut table = Vec::new();
        table.extend_from_slice(&0xFFFFu16.to_le_bytes());
        table.extend_from_slice(&(authors.len() as u16).to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        for author in authors {
            let units = author.encode_utf16().collect::<Vec<_>>();
            let length = u16::try_from(units.len()).map_err(|_| {
                WriteError::InvalidData(
                    "DOC revision author exceeds the STTB string-length limit".to_string(),
                )
            })?;
            table.extend_from_slice(&length.to_le_bytes());
            table.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        Ok(Some(RevisionWriterData { indexes, table }))
    }
    pub(in crate::writer::core::package) fn append_revision_author_table(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        revisions: &RevisionWriterData,
    ) {
        let offset = table_stream.len() as u32;
        fib.set_sttbf_rmark(offset, revisions.table.len() as u32);
        table_stream.extend_from_slice(&revisions.table);
    }
}

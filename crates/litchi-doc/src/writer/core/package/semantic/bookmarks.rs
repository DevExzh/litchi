use crate::writer::bookmarks::BookmarkEntry;
use crate::writer::core::{Writer, model::*};
use crate::writer::fib::FibBuilder;
use std::collections::HashMap;
impl Writer {
    pub(in crate::writer::core::package) fn build_bookmark_tables(
        entries: &[BookmarkEntry],
        document_end: u32,
    ) -> Result<Option<BookmarkTableData>, WriteError> {
        if entries.is_empty() {
            return Ok(None);
        }
        if entries.len() > 0x3FFB {
            return Err(WriteError::InvalidData(
                "DOC standard bookmark table exceeds 0x3FFB entries".to_string(),
            ));
        }
        let mut unique = std::collections::HashSet::with_capacity(entries.len());
        let mut records = Vec::with_capacity(entries.len());
        for (index, entry) in entries.iter().enumerate() {
            let units = entry.name.encode_utf16().collect::<Vec<_>>();
            if units.is_empty() || units.len() >= 40 || !unique.insert(entry.name.clone()) {
                return Err(WriteError::InvalidData(
                    "DOC bookmark names must be unique and contain 1 through 39 UTF-16 code units"
                        .to_string(),
                ));
            }
            if entry.start > entry.end || entry.end > document_end {
                return Err(WriteError::InvalidData(
                    "DOC bookmark range must be ordered and inside the document parts".to_string(),
                ));
            }
            let mut bkc = u16::from(entry.is_native) << 14;
            if let Some((first, limit)) = entry.column_range {
                if first >= limit || first > 0x7F || limit > 0x3F {
                    return Err(WriteError::InvalidData(
                        "DOC bookmark column range exceeds BKC limits".to_string(),
                    ));
                }
                bkc |= 0x8000 | u16::from(first) | (u16::from(limit) << 8);
            }
            records.push((index, entry, units, bkc));
        }

        let sentinel = document_end.checked_add(1).ok_or_else(|| {
            WriteError::InvalidData("DOC bookmark sentinel CP overflows".to_string())
        })?;
        let mut start_order = records.iter().collect::<Vec<_>>();
        start_order.sort_by_key(|record| (record.1.start, record.0));
        let mut end_order = records.iter().collect::<Vec<_>>();
        end_order.sort_by_key(|record| (record.1.end, record.0));
        let end_indexes = end_order
            .iter()
            .enumerate()
            .map(|(end_index, record)| (record.0, end_index as u16))
            .collect::<HashMap<_, _>>();

        let mut names = Vec::new();
        names.extend_from_slice(&0xFFFFu16.to_le_bytes());
        names.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        names.extend_from_slice(&0u16.to_le_bytes());
        for record in &start_order {
            names.extend_from_slice(&(record.2.len() as u16).to_le_bytes());
            names.extend(record.2.iter().copied().flat_map(u16::to_le_bytes));
        }

        let mut starts = Vec::with_capacity((entries.len() + 1) * 4 + entries.len() * 4);
        for record in &start_order {
            starts.extend_from_slice(&record.1.start.to_le_bytes());
        }
        starts.extend_from_slice(&sentinel.to_le_bytes());
        for record in &start_order {
            starts.extend_from_slice(&end_indexes[&record.0].to_le_bytes());
            starts.extend_from_slice(&record.3.to_le_bytes());
        }

        let mut ends = Vec::with_capacity((entries.len() + 1) * 4);
        for record in &end_order {
            ends.extend_from_slice(&record.1.end.to_le_bytes());
        }
        ends.extend_from_slice(&sentinel.to_le_bytes());
        Ok(Some(BookmarkTableData {
            names,
            starts,
            ends,
        }))
    }
    pub(in crate::writer::core::package) fn append_bookmark_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        bookmarks: &BookmarkTableData,
    ) {
        let mut offset = table_stream.len() as u32;
        fib.set_sttbf_bkmk(offset, bookmarks.names.len() as u32);
        table_stream.extend_from_slice(&bookmarks.names);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkf(offset, bookmarks.starts.len() as u32);
        table_stream.extend_from_slice(&bookmarks.starts);
        offset = table_stream.len() as u32;
        fib.set_plcf_bkl(offset, bookmarks.ends.len() as u32);
        table_stream.extend_from_slice(&bookmarks.ends);
    }
}

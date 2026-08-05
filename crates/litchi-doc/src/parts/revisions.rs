//! Revision-mark author table parsing for Word 97+ documents.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// Parsed `SttbfRMark` author names.
#[derive(Debug, Clone, Default)]
pub struct RevisionAuthorTable {
    authors: Vec<String>,
}

impl RevisionAuthorTable {
    /// Parse the optional revision-mark author table at FIB index 51.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let Some((offset, length)) = fib.get_table_pointer(51).filter(|(_, length)| *length != 0)
        else {
            return Ok(Self::default());
        };
        let start = usize::try_from(offset)
            .map_err(|_| PackageError::Corrupted("SttbfRMark offset is too large".to_string()))?;
        let length = usize::try_from(length)
            .map_err(|_| PackageError::Corrupted("SttbfRMark length is too large".to_string()))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| PackageError::Corrupted("SttbfRMark range overflows".to_string()))?;
        let data = table_stream.get(start..end).ok_or_else(|| {
            PackageError::Corrupted("SttbfRMark extends beyond the table stream".to_string())
        })?;
        if data.len() < 6
            || read_u16(data, 0, "SttbfRMark fExtend")? != 0xFFFF
            || read_u16(data, 4, "SttbfRMark cbExtra")? != 0
        {
            return Err(PackageError::Corrupted(
                "SttbfRMark has an invalid header".to_string(),
            ));
        }
        let count = usize::from(read_u16(data, 2, "SttbfRMark count")?);
        if count == 0 {
            return Err(PackageError::Corrupted(
                "SttbfRMark must begin with the Unknown author".to_string(),
            ));
        }
        let mut offset = 6usize;
        let mut authors = Vec::with_capacity(count);
        for _ in 0..count {
            let char_count = usize::from(read_u16(data, offset, "revision author length")?);
            offset = offset.checked_add(2).ok_or_else(|| {
                PackageError::Corrupted("revision author offset overflows".to_string())
            })?;
            let byte_count = char_count.checked_mul(2).ok_or_else(|| {
                PackageError::Corrupted("revision author length overflows".to_string())
            })?;
            let author_end = offset.checked_add(byte_count).ok_or_else(|| {
                PackageError::Corrupted("revision author range overflows".to_string())
            })?;
            let author_data = data.get(offset..author_end).ok_or_else(|| {
                PackageError::Corrupted("revision author is truncated".to_string())
            })?;
            let units = author_data
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>();
            authors.push(String::from_utf16(&units).map_err(|_| {
                PackageError::Corrupted("revision author contains invalid UTF-16".to_string())
            })?);
            offset = author_end;
        }
        if offset != data.len() || authors.first().map(String::as_str) != Some("Unknown") {
            return Err(PackageError::Corrupted(
                "SttbfRMark has trailing bytes or does not begin with Unknown".to_string(),
            ));
        }
        Ok(Self { authors })
    }

    /// Author names by the indexes stored in revision SPRMs.
    pub fn authors(&self) -> &[String] {
        &self.authors
    }

    /// Resolve one revision-author index.
    pub fn get(&self, index: u16) -> Option<&str> {
        self.authors.get(usize::from(index)).map(String::as_str)
    }

    #[cfg(test)]
    pub(crate) fn from_authors(authors: &[&str]) -> Self {
        Self {
            authors: authors.iter().map(|author| (*author).to_string()).collect(),
        }
    }
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| PackageError::Corrupted(format!("invalid {field}: {error}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn set_pointer(fib: &mut [u8], offset: u32, length: u32) {
        fib[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib[152..154].copy_from_slice(&93u16.to_le_bytes());
        let start = 154 + 51 * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn table(authors: &[&str]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0xFFFFu16.to_le_bytes());
        data.extend_from_slice(&(authors.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        for author in authors {
            let units = author.encode_utf16().collect::<Vec<_>>();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            data.extend(units.into_iter().flat_map(u16::to_le_bytes));
        }
        data
    }

    #[test]
    fn parses_unicode_revision_authors() {
        let authors = table(&["Unknown", "张三 😀"]);
        let mut fib_data = vec![0; 154 + 93 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        set_pointer(&mut fib_data, 0, authors.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let parsed = RevisionAuthorTable::parse(&fib, &authors).unwrap();
        assert_eq!(parsed.authors(), ["Unknown", "张三 😀"]);
        assert_eq!(parsed.get(1), Some("张三 😀"));
        assert_eq!(parsed.get(2), None);
    }

    #[test]
    fn rejects_invalid_revision_author_tables() {
        let mut fib_data = vec![0; 154 + 93 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        let wrong_first = table(&["Someone"]);
        set_pointer(&mut fib_data, 0, wrong_first.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(RevisionAuthorTable::parse(&fib, &wrong_first).is_err());

        let mut trailing = table(&["Unknown"]);
        trailing.push(0);
        set_pointer(&mut fib_data, 0, trailing.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(RevisionAuthorTable::parse(&fib, &trailing).is_err());
    }
}

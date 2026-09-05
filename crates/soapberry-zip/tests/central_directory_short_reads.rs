#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "Small generated ZIP fixtures and explicit bounded reader splits."
)]

use soapberry_zip::{ErrorKind, ReaderAt, ZipArchive, ZipLocator, office::StreamingArchiveWriter};
use std::io;

#[derive(Debug)]
struct Capped<'a> {
    bytes: &'a [u8],
    cap: usize,
}

impl ReaderAt for Capped<'_> {
    fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        let remaining = self.bytes.get(start..).unwrap_or_default();
        let count = remaining.len().min(output.len()).min(self.cap);
        output[..count].copy_from_slice(&remaining[..count]);
        Ok(count)
    }
}

fn fixture() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("a", b"").unwrap();
    writer.write_stored("b", b"").unwrap();
    writer.finish_to_bytes().unwrap()
}

fn names(
    bytes: &[u8],
    cap: usize,
    scratch_size: usize,
) -> Result<Vec<Vec<u8>>, soapberry_zip::Error> {
    let reader = Capped { bytes, cap };
    let mut locate_scratch = [0_u8; 512];
    let archive = ZipLocator::new()
        .locate_in_reader(reader, &mut locate_scratch, bytes.len() as u64)
        .map_err(|(_, error)| error)?;
    let mut scratch = vec![0_u8; scratch_size];
    let mut entries = archive.entries(&mut scratch);
    let mut result = Vec::new();
    while let Some(entry) = entries.next_entry()? {
        result.push(entry.file_path().as_ref().to_vec());
    }
    Ok(result)
}

#[test]
fn fixed_header_refill_counts_bytes_already_buffered() {
    let bytes = fixture();
    let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
    assert_eq!(bytes.len() - 22 - archive.directory_offset() as usize, 94);
    // The first read returns 80 bytes: record a consumes 47, leaving 33
    // buffered bytes and only 14 unread bytes in the declared directory.
    assert_eq!(
        names(&bytes, 80, 512).unwrap(),
        [b"a".to_vec(), b"b".to_vec()]
    );
}

#[test]
fn exactly_buffered_fixed_header_needs_only_its_variable_field() {
    let bytes = fixture();
    // After record a, all 46 fixed bytes of record b are already buffered.
    // Only its one-byte name remains in the underlying source.
    assert_eq!(
        names(&bytes, 93, 512).unwrap(),
        [b"a".to_vec(), b"b".to_vec()]
    );
}

#[test]
fn directory_records_are_independent_of_read_and_scratch_boundaries() {
    let bytes = fixture();
    for cap in 1..=128 {
        for scratch in [46, 47, 64, 92, 93, 94, 128, 512] {
            assert_eq!(
                names(&bytes, cap, scratch)
                    .unwrap_or_else(|error| panic!("cap={cap} scratch={scratch}: {error}")),
                [b"a".to_vec(), b"b".to_vec()]
            );
        }
    }
}

#[test]
fn partial_final_fixed_header_is_not_silently_ignored() {
    let bytes = fixture();
    let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
    let central_start = archive.directory_offset() as usize;
    let mut end = bytes[bytes.len() - 22..].to_vec();
    // Declare one full record plus 21 bytes of a truncated second record.
    // Count one allows the locator's minimum-size check; the entry iterator
    // must still reject malformed bytes inside the declared directory span.
    end[8..10].copy_from_slice(&1_u16.to_le_bytes());
    end[10..12].copy_from_slice(&1_u16.to_le_bytes());
    end[12..16].copy_from_slice(&68_u32.to_le_bytes());
    let mut truncated = bytes[..central_start + 68].to_vec();
    truncated.extend_from_slice(&end);
    let error = names(&truncated, 128, 512).expect_err("partial record must fail closed");
    assert!(matches!(error.kind(), ErrorKind::Eof));
}

#[test]
fn exact_final_fixed_header_is_parsed_at_directory_end() {
    let bytes = fixture();
    let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
    let start = archive.directory_offset() as usize;
    let second = start + 47;
    let mut end = bytes[bytes.len() - 22..].to_vec();
    end[12..16].copy_from_slice(&93_u32.to_le_bytes());
    let mut empty_name = bytes[..second + 46].to_vec();
    empty_name[second + 28..second + 30].copy_from_slice(&0_u16.to_le_bytes());
    empty_name.extend_from_slice(&end);
    // The central iterator exposes the fixed record even when it has no
    // variable fields. Higher-level path validation remains separate.
    assert_eq!(
        names(&empty_name, 128, 512).unwrap(),
        [b"a".to_vec(), vec![]]
    );
}

#[test]
fn zip32_and_zip64_extra_fields_survive_short_reads() {
    let fixtures: [&[u8]; 3] = [
        include_bytes!(
            "../../../docs/performance/results/change-0416/corpus/zip32-signed-store-deflate.zip"
        ),
        include_bytes!(
            "../../../docs/performance/results/change-0416/corpus/zip64-central-local-signed.zip"
        ),
        include_bytes!(
            "../../../docs/performance/results/change-0416/corpus/zip64-central-local-zip32-tail-signed.zip"
        ),
    ];
    for (index, bytes) in fixtures.into_iter().enumerate() {
        let expected = names(bytes, usize::MAX, 4096).unwrap();
        assert!(!expected.is_empty());
        for cap in [1, 2, 7, 45, 46, 97, 4096] {
            for scratch in [128, 512, 4096] {
                assert_eq!(
                    names(bytes, cap, scratch).unwrap_or_else(|error| panic!(
                        "fixture={index} cap={cap} scratch={scratch}: {error}"
                    )),
                    expected
                );
            }
        }
    }
}

#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use soapberry_zip::office::{ArchiveLimits, ArchiveReader, IndexedArchive};
use soapberry_zip::{RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive};

const MAX_INPUT_BYTES: usize = 1 << 20;
const MAX_FILES: usize = 256;
const MAX_MEMBER_NAME_BYTES: u64 = 4 << 10;
const MAX_METADATA_BYTES: u64 = 64 << 10;
const MAX_ENTRY_BYTES: u64 = 1 << 20;
const MAX_PRECOMPRESSED_PROGRESS_EVENTS: usize = 8;

fn fuzz_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: MAX_FILES,
        max_member_name_bytes: MAX_MEMBER_NAME_BYTES,
        max_metadata_bytes: MAX_METADATA_BYTES,
        max_compressed_size: MAX_ENTRY_BYTES,
        max_entry_size: MAX_ENTRY_BYTES,
        max_total_size: MAX_ENTRY_BYTES,
    }
}

fn exercise_borrowed_reader(reader: &ArchiveReader<'_>) {
    let mut borrowed_probe_done = false;
    let mut file_count = 0usize;

    for name in reader.file_names() {
        file_count += 1;
        let _ = black_box(reader.metadata(name));
        let stored = reader.is_stored(name);

        // This is the public strict borrowed path.  It validates every local
        // span, including local names, sizes, CRCs, and data descriptors,
        // before publishing a source slice.  Probe one Store member so an
        // archive with many entries does not turn the O(n^2) overlap proof
        // into an avoidable fuzzing cost.
        if !borrowed_probe_done && matches!(stored, Ok(true)) {
            if let Ok(Some(payload)) = reader.read_stored_borrowed(name) {
                if let Ok(metadata) = reader.metadata(name) {
                    assert_eq!(
                        u64::try_from(payload.len()).ok(),
                        Some(metadata.uncompressed_size())
                    );
                    assert!(!metadata.is_directory());
                }
                black_box(payload);
            }
            borrowed_probe_done = true;
        }
    }

    assert_eq!(file_count, reader.len());
}

fn exercise_reader_at_metadata<R: ReaderAt>(archive: &IndexedArchive<R>) {
    black_box(archive.archive_end_offset());
    black_box(archive.preservation_entry_count());
    black_box(archive.preservation_metadata_bytes());
    black_box(archive.has_encrypted_entries());
    black_box(archive.has_data_descriptor_entries());
    black_box(archive.archive_is_zip64());
    black_box(archive.has_zip64_metadata());
    black_box(archive.all_local_spans_bounded());

    let mut file_count = 0usize;
    for name in archive.file_names() {
        file_count += 1;
        assert!(archive.contains(name));
        let Some(entry_id) = archive.entry_id(name) else {
            continue;
        };
        if let Ok(metadata) = archive.metadata_for(entry_id) {
            assert!(!metadata.is_directory());
            black_box((metadata.compressed_size(), metadata.uncompressed_size()));
        }
        let _ = black_box(archive.metadata(name));
        let _ = black_box(archive.is_stored(name));
    }
    assert_eq!(file_count, archive.len());
}

fn exercise_preservation_index<R: ReaderAt>(archive: &IndexedArchive<R>, limits: ArchiveLimits) {
    let mut scratch = [0u8; RECOMMENDED_BUFFER_SIZE];
    let Ok(index) = archive.preservation_index_with_limits(&mut scratch, limits) else {
        return;
    };

    // Preservation indexing is metadata-only, but it runs the strict
    // ReaderAt local-header and descriptor validation for every member.
    assert_eq!(index.entries().len(), archive.preservation_entry_count());
    assert_eq!(index.archive_end_offset(), archive.archive_end_offset());
    for entry in index.entries() {
        let local_span = entry.local_span();
        let central_record = entry.central_record();
        assert!(local_span.start < local_span.end);
        assert!(central_record.start < central_record.end);
        let _ = black_box(entry.id());
        let _ = black_box(entry.compression_method());
        let _ = black_box(entry.compressed_size());
        let _ = black_box(entry.uncompressed_size());
        let _ = black_box(entry.raw_name_bytes());
    }
}

fn exercise_precompressed<R: ReaderAt>(archive: &IndexedArchive<R>, borrowed: &ArchiveReader<'_>) {
    for name in archive.file_names() {
        let Some(entry_id) = archive.entry_id(name) else {
            continue;
        };
        let Ok(metadata) = archive.metadata_for(entry_id) else {
            continue;
        };
        if metadata.uncompressed_size() > MAX_ENTRY_BYTES {
            continue;
        }

        // The borrowed reader supplies the already-decoded logical bytes.
        // Its admission limits keep this allocation within MAX_ENTRY_BYTES;
        // the precompressed API then captures and verifies only the matching
        // bounded source member.
        let Ok(decoded) = borrowed.read(name) else {
            continue;
        };
        let Ok(decoded_size) = u64::try_from(decoded.len()) else {
            continue;
        };
        assert_eq!(decoded_size, metadata.uncompressed_size());

        let mut progress_events = 0usize;
        let result =
            archive.read_entry_precompressed_with_progress(entry_id, &decoded, |progress| {
                progress_events = progress_events.saturating_add(1);
                let _ = black_box(progress);
                if progress_events > MAX_PRECOMPRESSED_PROGRESS_EVENTS {
                    Err(())
                } else {
                    Ok(())
                }
            });
        if let Ok(token) = result {
            assert_eq!(token.compressed_size(), metadata.compressed_size());
            assert_eq!(token.uncompressed_size(), decoded_size);
            assert_eq!(token.crc32(), soapberry_zip::crc32(&decoded));
            let _ = black_box(token.compression_method());
        }

        // One bounded member is enough to cover both successful small-member
        // tokens and callback cancellation for larger admitted members.
        break;
    }
}

fn exercise_bounded_paths(data: &[u8]) {
    let limits = fuzz_limits();
    let borrowed = ArchiveReader::new_with_limits(data, limits);
    let reader_at = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits);

    if let Ok(reader) = &borrowed {
        exercise_borrowed_reader(reader);
    }
    if let Ok(archive) = &reader_at {
        exercise_reader_at_metadata(archive);
        exercise_preservation_index(archive, limits);
    }

    // Both bounded indexes own the same public metadata contract.  Comparing
    // successful results keeps this target sensitive to drift between the
    // contiguous borrowed source and the positional ReaderAt source.
    if let (Ok(reader), Ok(archive)) = (&borrowed, &reader_at) {
        exercise_precompressed(archive, reader);
        for name in reader.file_names() {
            assert!(archive.contains(name));
            if let (Ok(borrowed_metadata), Ok(reader_at_metadata)) =
                (reader.metadata(name), archive.metadata(name))
            {
                assert_eq!(borrowed_metadata, reader_at_metadata);
            }
        }
    }
}

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }

    // Parse the ZIP via the slice-based entrypoint. This exercises EOCD
    // discovery (incl. ZIP64 locator + ZIP64 EOCD) and central-directory
    // header parsing.
    if let Ok(archive) = ZipArchive::from_slice(data) {
        // Touch the EOCD-derived metadata.
        let _ = archive.entries_hint();
        let _ = archive.eocd_offset();
        let _ = archive.directory_offset();
        let _ = archive.end_offset();

        // Iterate central-directory metadata and retain the existing path
        // probes. Local headers and data descriptors are exercised by the
        // bounded borrowed and preservation paths below.
        for entry_result in archive.entries() {
            let Ok(entry) = entry_result else { break };
            let path = entry.file_path();
            let _ = path.as_ref();
            let _ = path.try_normalize();
            let _ = entry.is_dir();
            let _ = entry.compression_method();
            let _ = entry.compressed_size_hint();
            let _ = entry.uncompressed_size_hint();
            let _ = entry.crc32();
        }
    }

    exercise_bounded_paths(data);
});

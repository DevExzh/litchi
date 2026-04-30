#![no_main]

use libfuzzer_sys::fuzz_target;
use soapberry_zip::ZipArchive;

fuzz_target!(|data: &[u8]| {
    // Parse the ZIP via the slice-based entrypoint. This exercises EOCD
    // discovery (incl. ZIP64 locator + ZIP64 EOCD) and central-directory
    // header parsing.
    if let Ok(archive) = ZipArchive::from_slice(data) {
        // Touch the EOCD-derived metadata.
        let _ = archive.entries_hint();
        let _ = archive.eocd_offset();
        let _ = archive.directory_offset();
        let _ = archive.end_offset();

        // Iterate central-directory entries — this is where the bulk of the
        // parser code lives (local file headers, extra fields, ZIP64 extras,
        // file path decoding, mode bits, timestamps, data descriptors).
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
});

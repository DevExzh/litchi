//! Physical ZIP64 Deflate output observation, separate from the ABBA probe.
use soapberry_zip::office::{
    ArchiveLimits, IndexedArchive, StreamingArchiveLimits, StreamingArchiveWriter,
};
use soapberry_zip::{CompressionMethod, FileReader};
use std::fs::File;
use std::io::{self, Write};
use std::time::Instant;

struct Digest {
    bytes: u64,
    crc: crc32fast::Hasher,
}
impl Write for Digest {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes += bytes.len() as u64;
        self.crc.update(bytes);
        Ok(bytes.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    let path = &args[1];
    let size: u64 = args[2].parse()?;
    let output_limit = size
        .checked_add(size / 8)
        .and_then(|n| n.checked_add(65536))
        .ok_or("size overflow")?;
    let mut block = [0u8; 65536];
    let mut state = 0x0315_0415_a987_1234u64;
    for byte in &mut block {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        *byte = state as u8;
    }
    let limits = StreamingArchiveLimits::new(4, 64, 4096)
        .with_byte_limits(size, size, output_limit)
        .with_compressed_size_limit(output_limit - 4096);
    let file = File::create(path)?;
    let mut input_crc = crc32fast::Hasher::new();
    let start = Instant::now();
    let writer = StreamingArchiveWriter::with_writer_and_limits(file, limits);
    let mut entry = writer.start_entry("payload.bin", CompressionMethod::Deflate)?;
    let mut remaining = size;
    while remaining > 0 {
        let n = remaining.min(block.len() as u64) as usize;
        entry.write_all(&block[..n])?;
        input_crc.update(&block[..n]);
        remaining -= n as u64;
    }
    let file = entry.finish()?.finish()?;
    let elapsed_write_ns = start.elapsed().as_nanos();
    let output_bytes = file.metadata()?.len();
    drop(file);
    let expected_crc = input_crc.finalize();

    let start = Instant::now();
    let limits = ArchiveLimits {
        max_entry_size: size,
        max_total_size: size,
        max_compressed_size: output_limit,
        ..ArchiveLimits::default()
    };
    let archive = IndexedArchive::from_reader_with_limits(
        FileReader::from(File::open(path)?),
        output_bytes,
        limits,
    )?;
    let metadata = archive.metadata("payload.bin")?;
    let mut decoded = Digest {
        bytes: 0,
        crc: crc32fast::Hasher::new(),
    };
    assert_eq!(archive.read_to("payload.bin", &mut decoded)?, size);
    assert_eq!(decoded.bytes, size);
    assert_eq!(decoded.crc.finalize(), expected_crc);
    assert!(metadata.compressed_size() >= u64::from(u32::MAX));
    assert!(archive.archive_is_zip64());
    assert!(archive.has_data_descriptor_entries());
    let elapsed_read_ns = start.elapsed().as_nanos();
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "input_bytes": size, "compressed_bytes": metadata.compressed_size(), "output_bytes": output_bytes,
            "input_buffer_bytes": block.len(), "input_crc32": expected_crc,
            "archive_zip64": true, "descriptor": true, "verified": true,
            "elapsed_write_ns": elapsed_write_ns, "elapsed_read_ns": elapsed_read_ns,
            "payload": "repeated 65536-byte xorshift block; seed 0x03150415a9871234",
            "measurement_scope": "write includes input CRC observer; process RSS includes separate subsequent bounded readback"
        }))?
    );
    Ok(())
}

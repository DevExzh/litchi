//! Retryable sink interruptions must not poison the bounded owned-entry adapter.
use std::io::{self, Write};

use soapberry_zip::CompressionMethod;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

#[derive(Default)]
struct ShortInterruptedSink {
    bytes: Vec<u8>,
    interrupt_next: bool,
}

impl Write for ShortInterruptedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.interrupt_next {
            self.interrupt_next = false;
            return Err(io::ErrorKind::Interrupted.into());
        }
        self.interrupt_next = true;
        let count = bytes.len().min(37);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn interrupted_owned_payload_writes_resume_without_poisoning_or_double_counting() {
    let mut state = 0x9e37_79b9_u32;
    let payload: Vec<u8> = (0..128 * 1024 + 19)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            (state >> 24) as u8
        })
        .collect();
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let writer = StreamingArchiveWriter::with_writer(ShortInterruptedSink::default());
        let mut entry = writer.start_entry("payload", method).unwrap();
        let mut accepted = 0;
        let mut interruptions = 0;
        while accepted < payload.len() {
            match entry.write(&payload[accepted..]) {
                Ok(count) => {
                    assert!(count > 0);
                    accepted += count;
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {
                    interruptions += 1;
                    assert!(!entry.is_poisoned());
                    assert_eq!(entry.uncompressed_bytes(), accepted as u64);
                },
                Err(error) => panic!("unexpected sink failure: {error}"),
            }
        }
        assert!(interruptions > 0);
        assert_eq!(entry.uncompressed_bytes(), payload.len() as u64);
        let sink = entry.finish().unwrap().finish().unwrap();
        let archive = ArchiveReader::new(&sink.bytes).unwrap();
        assert_eq!(archive.read("payload").unwrap(), payload);
    }
}

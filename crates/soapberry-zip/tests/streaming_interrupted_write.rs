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
        let writer = StreamingArchiveWriter::with_writer(
            ShortInterruptedSink::default(),
            soapberry_zip::office::StreamingArchiveLimits::default(),
        );
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
        let archive =
            ArchiveReader::new(&sink.bytes, soapberry_zip::office::ArchiveLimits::default())
                .unwrap();
        assert_eq!(archive.read("payload").unwrap(), payload);
    }
}

#[derive(Default)]
struct InterruptedFlushSink {
    bytes: Vec<u8>,
    interrupted_flushes: usize,
}

impl Write for InterruptedFlushSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.interrupted_flushes == 0 {
            self.interrupted_flushes = 1;
            return Err(io::ErrorKind::Interrupted.into());
        }
        Ok(())
    }
}

#[test]
fn interrupted_owned_entry_flush_is_retryable_without_poisoning() {
    let payload: Vec<u8> = (0..40_000_u32).map(|index| (index % 251) as u8).collect();
    for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
        let writer = StreamingArchiveWriter::with_writer(
            InterruptedFlushSink::default(),
            soapberry_zip::office::StreamingArchiveLimits::default(),
        );
        let mut entry = writer.start_entry("payload", method).unwrap();
        let (head, tail) = payload.split_at(payload.len() / 2);
        entry.write_all(head).unwrap();

        let error = entry.flush().unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::Interrupted);
        assert!(!entry.is_poisoned());
        assert!(!entry.progress().is_poisoned());
        assert_eq!(entry.uncompressed_bytes(), head.len() as u64);

        entry.flush().unwrap();
        entry.write_all(tail).unwrap();
        assert_eq!(entry.uncompressed_bytes(), payload.len() as u64);
        let sink = entry.finish().unwrap().finish().unwrap();
        assert_eq!(sink.interrupted_flushes, 1);
        let archive =
            ArchiveReader::new(&sink.bytes, soapberry_zip::office::ArchiveLimits::default())
                .unwrap();
        assert_eq!(archive.read("payload").unwrap(), payload);
    }
}

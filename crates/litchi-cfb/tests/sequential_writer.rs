#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused writer tests use panic-on-failure assertions"
)]

use litchi_cfb::writer::{
    OleWriter, SequentialOleWriter, SequentialWriteError, SequentialWriteProgress,
    SequentialWriterLimits, SequentialWriterOptions,
};
use litchi_cfb::{OleError, OleFile};
use litchi_core::CancellationSource;
use std::io::{self, Cursor, Write};

fn publish(writer: SequentialOleWriter<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    writer.write_to(&mut output).unwrap();
    output
}

#[test]
fn mixed_minifat_fat_nested_storage_clsids_reopen() {
    let mut writer = SequentialOleWriter::new();
    writer.set_root_clsid([0xA1; 16]);
    writer.create_storage(&["Nested"]).unwrap();
    writer.set_storage_clsid(&["Nested"], [0xB2; 16]).unwrap();
    writer
        .add_stream(&["Empty"], 0, Cursor::new(Vec::<u8>::new()))
        .unwrap();
    writer
        .add_stream(&["Nested", "Small"], 4_095, Cursor::new(vec![0x11; 4_095]))
        .unwrap();
    writer
        .add_stream(&["Large"], 4_096, Cursor::new(vec![0x22; 4_096]))
        .unwrap();

    let bytes = publish(writer);
    assert_eq!(bytes.len() % 512, 0);
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    assert_eq!(
        file.root_entry().unwrap().clsid,
        "A1A1A1A1-A1A1-A1A1-A1A1-A1A1A1A1A1A1"
    );
    assert_eq!(
        file.list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Nested")
            .unwrap()
            .clsid,
        "B2B2B2B2-B2B2-B2B2-B2B2-B2B2B2B2B2B2"
    );
    assert_eq!(file.open_stream(&["Empty"]).unwrap(), Vec::<u8>::new());
    assert_eq!(
        file.open_stream(&["Nested", "Small"]).unwrap(),
        vec![0x11; 4_095]
    );
    assert_eq!(file.open_stream(&["Large"]).unwrap(), vec![0x22; 4_096]);
    assert_eq!(file.stream_len(&["Nested", "Small"]).unwrap(), 4_095);
}

#[test]
fn sector_size_4096_and_current_writer_wire_parity() {
    let mut sequential = SequentialOleWriter::with_sector_size(4096).unwrap();
    sequential
        .add_stream(&["Small"], 7, Cursor::new(b"payload".to_vec()))
        .unwrap();
    sequential
        .add_stream(&["Large"], 4_096, Cursor::new(vec![0x5A; 4_096]))
        .unwrap();
    let bytes = publish(sequential);
    assert_eq!(bytes.len() % 4096, 0);
    let mut parsed = OleFile::open(Cursor::new(bytes)).unwrap();
    assert_eq!(parsed.open_stream(&["Small"]).unwrap(), b"payload");

    let mut legacy = OleWriter::with_sector_size(512).unwrap();
    legacy.create_stream(&["Small"], b"payload").unwrap();
    legacy
        .create_stream_owned(&["Large"], vec![0x5A; 4_096])
        .unwrap();
    let mut legacy_bytes = Cursor::new(Vec::new());
    legacy.write_to(&mut legacy_bytes).unwrap();

    let mut sequential_512 = SequentialOleWriter::new();
    sequential_512
        .add_stream(&["Small"], 7, Cursor::new(b"payload".to_vec()))
        .unwrap();
    sequential_512
        .add_stream(&["Large"], 4_096, Cursor::new(vec![0x5A; 4_096]))
        .unwrap();
    assert_eq!(publish(sequential_512), legacy_bytes.into_inner());
}

#[test]
fn source_length_failures_are_typed_and_leave_prefix_evidence() {
    let mut short = SequentialOleWriter::new();
    short
        .add_stream(&["Short"], 5, Cursor::new(b"four".to_vec()))
        .unwrap();
    let mut sink = Vec::new();
    let error = short.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::SourceLength {
            expected: 5,
            observed: 4,
            progress: SequentialWriteProgress::Prefix { .. },
            ..
        }
    ));
    assert!(!sink.is_empty());

    let mut extra = SequentialOleWriter::new();
    extra
        .add_stream(&["Extra"], 4, Cursor::new(b"five!".to_vec()))
        .unwrap();
    let error = extra.write_to(&mut Vec::new()).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::SourceLength {
            expected: 4,
            observed: 5,
            progress: SequentialWriteProgress::Prefix { .. },
            ..
        }
    ));
}

#[test]
fn sink_zero_and_flush_failures_report_progress() {
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Data"], 4_096, Cursor::new(vec![0x33; 4_096]))
        .unwrap();
    let error = writer.write_to(&mut ZeroSink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::WriteZero {
            progress: SequentialWriteProgress::Untouched
        }
    ));

    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Data"], 1, Cursor::new(vec![7]))
        .unwrap();
    let mut sink = FlushFailSink::default();
    let error = writer.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::Flush {
            progress: SequentialWriteProgress::CompleteUnflushed { .. },
            ..
        }
    ));

    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Data"], 1, Cursor::new(vec![7]))
        .unwrap();
    let mut sink = FailSink {
        bytes: Vec::new(),
        limit: 600,
    };
    let error = writer.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::Sink {
            progress: SequentialWriteProgress::Prefix { accepted, .. },
            ..
        } if accepted >= 512
    ));
}

#[test]
fn short_nonseek_sink_matches_reference_bytes() {
    let mut reference_writer = SequentialOleWriter::new();
    reference_writer
        .add_stream(&["Data"], 4_095, Cursor::new(vec![0x44; 4_095]))
        .unwrap();
    let reference = publish(reference_writer);

    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Data"], 4_095, Cursor::new(vec![0x44; 4_095]))
        .unwrap();
    let mut sink = ShortSink {
        bytes: Vec::new(),
        maximum: 7,
    };
    let report = writer.write_to(&mut sink).unwrap();
    assert_eq!(sink.bytes, reference);
    assert_eq!(
        report.output_bytes(),
        u64::try_from(reference.len()).unwrap()
    );
}

#[test]
fn cancellation_after_header_reports_an_exact_prefix() {
    let (source, token) = CancellationSource::pair();
    let mut options = SequentialWriterOptions::default().with_cancellation(token);
    options.publication_buffer_bytes = 512;
    let mut writer = SequentialOleWriter::with_options(options).unwrap();
    writer
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let mut sink = CancelAfterFirstWrite {
        bytes: Vec::new(),
        source,
        cancelled: false,
    };
    let error = writer.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::Cancelled {
            progress: SequentialWriteProgress::Prefix { accepted: 512, .. }
        }
    ));
}

#[test]
fn interrupted_retries_recheck_cancellation_before_retrying() {
    let (source, token) = CancellationSource::pair();
    let options = SequentialWriterOptions::default().with_cancellation(token);
    let mut writer = SequentialOleWriter::with_options(options).unwrap();
    writer
        .add_stream(
            &["Data"],
            1,
            InterruptingReader {
                source,
                interrupted: false,
            },
        )
        .unwrap();
    let mut output = Vec::new();
    let error = writer.write_to(&mut output).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::Cancelled {
            progress: SequentialWriteProgress::Prefix { accepted: 512, .. }
        }
    ));

    let (source, token) = CancellationSource::pair();
    let options = SequentialWriterOptions::default().with_cancellation(token);
    let mut writer = SequentialOleWriter::with_options(options).unwrap();
    writer
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let mut sink = InterruptingSink {
        source,
        interrupted: false,
        calls: 0,
    };
    let error = writer.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::Cancelled {
            progress: SequentialWriteProgress::Untouched
        }
    ));
    assert_eq!(sink.calls, 1);
}

#[test]
fn output_limit_accepts_exact_and_rejects_one_under_without_sink_calls() {
    let mut baseline = SequentialOleWriter::new();
    baseline
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let bytes = publish(baseline);
    let exact = u64::try_from(bytes.len()).unwrap();

    let options = SequentialWriterOptions::default().with_limits(SequentialWriterLimits {
        max_output_bytes: exact,
        ..SequentialWriterLimits::default()
    });
    let mut accepted = SequentialOleWriter::with_options(options).unwrap();
    accepted
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    assert_eq!(publish(accepted).len(), bytes.len());

    let options = SequentialWriterOptions::default().with_limits(SequentialWriterLimits {
        max_output_bytes: exact - 1,
        ..SequentialWriterLimits::default()
    });
    let mut refused = SequentialOleWriter::with_options(options).unwrap();
    refused
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let mut sink = CountingSink::default();
    let error = refused.write_to(&mut sink).unwrap_err();
    assert!(matches!(error, SequentialWriteError::LimitExceeded { .. }));
    assert_eq!(sink.calls, 0);
}

#[test]
fn difat_boundary_streams_without_retaining_source_payload() {
    let sectors = 109_u64 * 128;
    let length = sectors * 512;
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(
            &["Large"],
            length,
            RepeatReader {
                remaining: length,
                byte: 0x7A,
            },
        )
        .unwrap();
    let bytes = publish(writer);
    let fat_count = u32::from_le_bytes(bytes[0x2C..0x30].try_into().unwrap());
    let difat_count = u32::from_le_bytes(bytes[0x48..0x4C].try_into().unwrap());
    assert!(fat_count > 109);
    assert!(difat_count >= 1);
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    assert_eq!(file.stream_len(&["Large"]).unwrap(), length);
    assert_eq!(
        file.open_stream(&["Large"]).unwrap().len(),
        usize::try_from(length).unwrap()
    );
}

#[test]
fn exact_fat_boundary_has_no_difat_and_one_over_has_difat() {
    let exact_length = 13_842_u64 * 512;
    let mut exact = SequentialOleWriter::new();
    exact
        .add_stream(
            &["Large"],
            exact_length,
            RepeatReader {
                remaining: exact_length,
                byte: 0x21,
            },
        )
        .unwrap();
    let exact_bytes = publish(exact);
    assert_eq!(
        u32::from_le_bytes(exact_bytes[0x2C..0x30].try_into().unwrap()),
        109
    );
    assert_eq!(
        u32::from_le_bytes(exact_bytes[0x48..0x4C].try_into().unwrap()),
        0
    );

    let over_length = 13_843_u64 * 512;
    let mut over = SequentialOleWriter::new();
    over.add_stream(
        &["Large"],
        over_length,
        RepeatReader {
            remaining: over_length,
            byte: 0x22,
        },
    )
    .unwrap();
    let over_bytes = publish(over);
    assert_eq!(
        u32::from_le_bytes(over_bytes[0x2C..0x30].try_into().unwrap()),
        110
    );
    assert_eq!(
        u32::from_le_bytes(over_bytes[0x48..0x4C].try_into().unwrap()),
        1
    );
}

#[test]
fn planning_limits_do_not_call_sink() {
    let limits = SequentialWriterLimits::new(1, 16, 8, 128, 8, 1024, 512);
    let options = SequentialWriterOptions::default().with_limits(limits);
    let mut writer = SequentialOleWriter::with_options(options).unwrap();
    writer
        .add_stream(&["TooLarge"], 9, Cursor::new(vec![0; 9]))
        .unwrap_err();

    let mut sink = CountingSink::default();
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let metadata_limits = SequentialWriterLimits {
        max_metadata_bytes: 1,
        ..SequentialWriterLimits::default()
    };
    let options = SequentialWriterOptions::default().with_limits(metadata_limits);
    let mut limited = SequentialOleWriter::with_options(options).unwrap();
    limited
        .add_stream(&["Data"], 1, Cursor::new(vec![1]))
        .unwrap();
    let error = limited.write_to(&mut sink).unwrap_err();
    assert!(matches!(
        error,
        SequentialWriteError::LimitExceeded {
            resource: "metadata bytes",
            ..
        }
    ));
    assert_eq!(sink.calls, 0);
    drop(writer);
}

#[test]
fn atomic_save_reopens_candidate_before_replace() {
    let path = std::env::temp_dir().join(format!(
        "litchi-cfb-sequential-{}-{}.ole",
        std::process::id(),
        std::thread::current().name().unwrap_or("test")
    ));
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Saved"], 5, Cursor::new(b"saved".to_vec()))
        .unwrap();
    let report = writer.save(&path).unwrap();
    assert_eq!(report.payload_bytes(), 5);
    let file = std::fs::File::open(&path).unwrap();
    let mut parsed = OleFile::open(file).unwrap();
    assert_eq!(parsed.open_stream(&["Saved"]).unwrap(), b"saved");
    std::fs::remove_file(path).unwrap();
}

#[test]
fn atomic_save_overwrites_existing_destination() {
    let path = std::env::temp_dir().join(format!(
        "litchi-cfb-sequential-overwrite-{}-{}.ole",
        std::process::id(),
        std::thread::current().name().unwrap_or("test")
    ));
    std::fs::write(&path, b"old destination").unwrap();
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&["Saved"], 5, Cursor::new(b"new!!".to_vec()))
        .unwrap();
    writer.save(&path).unwrap();

    let mut parsed = OleFile::open(std::fs::File::open(&path).unwrap()).unwrap();
    assert_eq!(parsed.open_stream(&["Saved"]).unwrap(), b"new!!");
    std::fs::remove_file(path).unwrap();
}

#[derive(Default)]
struct CountingSink {
    calls: usize,
}

impl Write for CountingSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        self.calls += 1;
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroSink;

impl Write for ZeroSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Default)]
struct FlushFailSink {
    bytes: Vec<u8>,
}

impl Write for FlushFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Err(io::Error::other("flush failed"))
    }
}

struct ShortSink {
    bytes: Vec<u8>,
    maximum: usize,
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.maximum);
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FailSink {
    bytes: Vec<u8>,
    limit: usize,
}

impl Write for FailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.limit {
            return Err(io::Error::other("sink failed"));
        }
        let count = bytes.len().min(self.limit - self.bytes.len());
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct CancelAfterFirstWrite {
    bytes: Vec<u8>,
    source: CancellationSource,
    cancelled: bool,
}

struct InterruptingReader {
    source: CancellationSource,
    interrupted: bool,
}

impl io::Read for InterruptingReader {
    fn read(&mut self, _bytes: &mut [u8]) -> io::Result<usize> {
        if self.interrupted {
            panic!("sequential writer retried an interrupted cancelled source");
        }
        self.interrupted = true;
        self.source.cancel();
        Err(io::Error::from(io::ErrorKind::Interrupted))
    }
}

struct InterruptingSink {
    source: CancellationSource,
    interrupted: bool,
    calls: usize,
}

impl Write for InterruptingSink {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        self.calls += 1;
        if self.interrupted {
            panic!("sequential writer retried an interrupted cancelled sink");
        }
        self.interrupted = true;
        self.source.cancel();
        Err(io::Error::from(io::ErrorKind::Interrupted))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for CancelAfterFirstWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.cancelled {
            self.cancelled = true;
            self.source.cancel();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct RepeatReader {
    remaining: u64,
    byte: u8,
}

impl io::Read for RepeatReader {
    fn read(&mut self, bytes: &mut [u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Ok(0);
        }
        let count = usize::try_from(self.remaining)
            .unwrap_or(bytes.len())
            .min(bytes.len());
        bytes[..count].fill(self.byte);
        self.remaining -= u64::try_from(count).unwrap();
        Ok(count)
    }
}

#[allow(dead_code)]
fn _ole_error_is_not_silenced(error: OleError) -> String {
    error.to_string()
}

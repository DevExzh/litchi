//! What a `StreamingArchiveWriter` reports before and after the failure that
//! poisons it: `is_poisoned`, `last_limit`, `progress`, the error of every
//! later call, and the bytes the sink holds at each step.
use std::cell::RefCell;
use std::io::{self, Cursor, Write};
use std::rc::Rc;

use soapberry_zip::CompressionMethod;
use soapberry_zip::office::{
    ArchiveLimits, ArchiveReader, StreamingArchiveFailure, StreamingArchiveLimits,
    StreamingArchiveWriter,
};

/// A sink that accepts `budget` bytes in total, splitting the write that
/// crosses the budget, and then fails every write. `fail_flush` makes every
/// flush fail as well.
#[derive(Clone)]
struct BudgetSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    budget: usize,
    fail_flush: bool,
}

impl BudgetSink {
    fn new(budget: usize) -> Self {
        Self {
            bytes: Rc::default(),
            budget,
            fail_flush: false,
        }
    }

    fn failing_flush() -> Self {
        Self {
            bytes: Rc::default(),
            budget: usize::MAX,
            fail_flush: true,
        }
    }

    fn len(&self) -> usize {
        self.bytes.borrow().len()
    }
}

impl Write for BudgetSink {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let mut bytes = self.bytes.borrow_mut();
        let room = self.budget - bytes.len();
        if room == 0 {
            return Err(io::Error::new(
                io::ErrorKind::StorageFull,
                "sink budget spent",
            ));
        }
        let count = buffer.len().min(room);
        bytes.extend_from_slice(&buffer[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.fail_flush {
            return Err(io::Error::new(
                io::ErrorKind::ConnectionReset,
                "sink flush refused",
            ));
        }
        Ok(())
    }
}

/// The writer's observable state on one line.
fn state<W: Write>(writer: &StreamingArchiveWriter<W>) -> String {
    format!(
        "poisoned={} limit={:?} output={} entries={} progress={:?}",
        writer.is_poisoned(),
        writer.last_limit(),
        writer.output_bytes(),
        writer.entry_count(),
        writer.progress(),
    )
}

/// A failure's error text, progress and limit on one line.
fn failure(failure: &StreamingArchiveFailure) -> String {
    format!(
        "error={:?} progress={:?} limit={:?}",
        failure.error().to_string(),
        failure.progress(),
        failure.limit(),
    )
}

fn small_limits() -> StreamingArchiveLimits {
    StreamingArchiveLimits::new(4, 16, 64).with_byte_limits(16, 32, 4096)
}

#[test]
fn a_usable_writer_reports_no_limit_through_refusals_and_finish() {
    let mut writer = StreamingArchiveWriter::new(StreamingArchiveLimits::default());
    assert_eq!(
        state(&writer),
        "poisoned=false limit=None output=0 entries=0 \
         progress=StreamingArchiveProgress { output_bytes: 0, poisoned: false }"
    );
    writer.write_stored("one", b"payload").unwrap();
    assert_eq!(
        state(&writer),
        "poisoned=false limit=None output=40 entries=1 \
         progress=StreamingArchiveProgress { output_bytes: 40, poisoned: false }"
    );

    // A refused member is not a failure of the writer.
    let refused = writer.write_stored("one", b"again").unwrap_err();
    assert_eq!(
        refused.to_string(),
        "Invalid input: duplicate normalized member name: one"
    );
    assert_eq!(
        state(&writer),
        "poisoned=false limit=None output=40 entries=1 \
         progress=StreamingArchiveProgress { output_bytes: 40, poisoned: false }"
    );

    // An entry hands back a usable writer.
    let mut entry = writer
        .start_entry("two", CompressionMethod::Deflate)
        .unwrap();
    entry.write_all(b"second payload").unwrap();
    let (writer, progress) = entry.finish_with_progress().unwrap();
    assert_eq!(
        format!("{progress:?}"),
        "StreamingArchiveProgress { output_bytes: 105, poisoned: false }"
    );
    assert_eq!(
        state(&writer),
        "poisoned=false limit=None output=105 entries=2 \
         progress=StreamingArchiveProgress { output_bytes: 105, poisoned: false }"
    );

    let (sink, progress) = writer.finish_with_progress().unwrap();
    assert_eq!(
        format!("{progress:?}"),
        "StreamingArchiveProgress { output_bytes: 225, poisoned: false }"
    );
    let bytes = sink.into_inner();
    assert_eq!(bytes.len(), 225);
    let reader = ArchiveReader::new(&bytes, ArchiveLimits::default()).unwrap();
    assert_eq!(reader.read("one").unwrap(), b"payload");
    assert_eq!(reader.read("two").unwrap(), b"second payload");
}

/// A writer poisoned by its sink failing inside the first member.
fn sink_poisoned(method: &str) -> (StreamingArchiveWriter<BudgetSink>, BudgetSink, String) {
    let sink = BudgetSink::new(36);
    let mut writer =
        StreamingArchiveWriter::with_writer(sink.clone(), StreamingArchiveLimits::default());
    let payload = b"payload that does not fit the sink";
    let error = match method {
        "stored" => writer.write_stored("a", payload),
        "deflated" => writer.write_deflated("a", payload),
        "deflated sized" => writer.write_deflated_sized("a", payload),
        "stored stream" => writer.write_stored_stream("a", Cursor::new(payload)),
        "deflated stream" => writer.write_deflated_stream("a", Cursor::new(payload)),
        other => panic!("unknown method {other}"),
    }
    .unwrap_err();
    (writer, sink, error.to_string())
}

#[test]
fn a_sink_failure_poisons_every_member_route_without_a_limit() {
    for method in [
        "stored",
        "deflated",
        "deflated sized",
        "stored stream",
        "deflated stream",
    ] {
        let (writer, sink, first) = sink_poisoned(method);
        assert_eq!(first, "sink budget spent", "{method}");
        assert_eq!(
            state(&writer),
            "poisoned=true limit=None output=36 entries=0 \
             progress=StreamingArchiveProgress { output_bytes: 36, poisoned: true }",
            "{method}"
        );
        assert_eq!(sink.len(), 36, "{method}");
    }
}

#[test]
fn a_poisoned_writer_refuses_every_later_call_and_writes_nothing_more() {
    let poisoned =
        "Invalid input: streaming archive writer is poisoned after 36 accepted output bytes";
    let after = "poisoned=true limit=None output=36 entries=0 \
                 progress=StreamingArchiveProgress { output_bytes: 36, poisoned: true }";
    let refusal = "error=\"Invalid input: streaming archive writer is poisoned after 36 accepted \
                   output bytes\" \
                   progress=StreamingArchiveProgress { output_bytes: 36, poisoned: true } \
                   limit=None";

    let (mut writer, sink, _) = sink_poisoned("stored stream");
    for attempt in 0..2 {
        for error in [
            writer.write_stored("b", b"x").unwrap_err(),
            writer.write_deflated("b", b"x").unwrap_err(),
            writer.write_deflated_sized("b", b"x").unwrap_err(),
            writer
                .write_stored_stream("b", Cursor::new(b"x"))
                .unwrap_err(),
            writer
                .write_deflated_stream("b", Cursor::new(b"x"))
                .unwrap_err(),
        ] {
            assert_eq!(error.to_string(), poisoned, "attempt {attempt}");
        }
        assert_eq!(state(&writer), after, "attempt {attempt}");
        assert_eq!(sink.len(), 36, "attempt {attempt}");
    }

    let refused = writer.finish_with_progress().err().unwrap();
    assert_eq!(failure(&refused), refusal);
    assert_eq!(sink.len(), 36);

    let (writer, sink, _) = sink_poisoned("stored stream");
    let refused = writer
        .start_entry("b", CompressionMethod::Store)
        .err()
        .unwrap();
    assert_eq!(failure(&refused), refusal);
    assert_eq!(sink.len(), 36);

    let (writer, sink, _) = sink_poisoned("stored stream");
    assert_eq!(writer.finish().err().unwrap().to_string(), poisoned);
    assert_eq!(sink.len(), 36);
}

#[test]
fn a_compressed_limit_poisons_with_that_limit() {
    let limits = small_limits().with_compressed_size_limit(3);
    let mut writer = StreamingArchiveWriter::new(limits);
    let error = writer
        .write_stored_stream("over", Cursor::new(b"abcd"))
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "ZIP compressed member size limit exceeded: declared 4, maximum 3"
    );
    let limit = "Some(StreamingLimitExceeded { resource: CompressedBytes, actual: 4, maximum: 3 })";
    let poisoned =
        "Invalid input: streaming archive writer is poisoned after 34 accepted output bytes";
    let progress = "StreamingArchiveProgress { output_bytes: 34, poisoned: true }";
    let after = format!("poisoned=true limit={limit} output=34 entries=0 progress={progress}");
    assert_eq!(state(&writer), after);

    let error = writer.write_stored("next", b"x").unwrap_err();
    assert_eq!(error.to_string(), poisoned);
    assert_eq!(state(&writer), after);

    // Finishing a poisoned writer repeats the limit that poisoned it.
    let refused = writer.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        format!("error={poisoned:?} progress={progress} limit={limit}")
    );

    // A poisoned writer that is asked for an entry reports no limit.
    let mut writer = StreamingArchiveWriter::new(limits);
    writer
        .write_stored_stream("over", Cursor::new(b"abcd"))
        .unwrap_err();
    let refused = writer
        .start_entry("next", CompressionMethod::Store)
        .err()
        .unwrap();
    assert_eq!(
        failure(&refused),
        format!("error={poisoned:?} progress={progress} limit=None")
    );
}

#[test]
fn an_output_limit_poisons_with_that_limit() {
    let limits = StreamingArchiveLimits::new(4, 16, 64).with_byte_limits(4096, 4096, 64);
    let mut writer = StreamingArchiveWriter::new(limits);
    let payload = [b'x'; 50];
    let error = writer
        .write_stored_stream("big", Cursor::new(payload))
        .unwrap_err();
    assert_eq!(
        error.to_string(),
        "streaming ZIP output bytes limit exceeded: attempted 83, maximum 64"
    );
    let limit = "Some(StreamingLimitExceeded { resource: OutputBytes, actual: 83, maximum: 64 })";
    let poisoned =
        "Invalid input: streaming archive writer is poisoned after 33 accepted output bytes";
    let progress = "StreamingArchiveProgress { output_bytes: 33, poisoned: true }";
    let after = format!("poisoned=true limit={limit} output=33 entries=0 progress={progress}");
    assert_eq!(state(&writer), after);

    let error = writer.write_stored("next", b"x").unwrap_err();
    assert_eq!(error.to_string(), poisoned);
    assert_eq!(state(&writer), after);

    let refused = writer.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        format!("error={poisoned:?} progress={progress} limit={limit}")
    );
}

#[test]
fn a_failed_finish_reports_poison_and_only_its_own_limit() {
    // The central directory does not fit the sink.
    let sink = BudgetSink::new(60);
    let mut writer =
        StreamingArchiveWriter::with_writer(sink.clone(), StreamingArchiveLimits::default());
    writer.write_stored("one", b"payload").unwrap();
    assert_eq!(
        state(&writer),
        "poisoned=false limit=None output=40 entries=1 \
         progress=StreamingArchiveProgress { output_bytes: 40, poisoned: false }"
    );
    let refused = writer.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        "error=\"sink budget spent\" \
         progress=StreamingArchiveProgress { output_bytes: 60, poisoned: true } limit=None"
    );
    assert_eq!(sink.len(), 60);

    // The final flush fails after every byte was accepted.
    let sink = BudgetSink::failing_flush();
    let mut writer =
        StreamingArchiveWriter::with_writer(sink.clone(), StreamingArchiveLimits::default());
    writer.write_stored("one", b"payload").unwrap();
    let refused = writer.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        "error=\"sink flush refused\" \
         progress=StreamingArchiveProgress { output_bytes: 111, poisoned: true } limit=None"
    );
    assert_eq!(sink.len(), 111);
}

#[test]
fn an_entry_that_cannot_start_or_finish_reports_poison_without_a_writer() {
    // The local header does not fit the sink.
    let sink = BudgetSink::new(10);
    let writer =
        StreamingArchiveWriter::with_writer(sink.clone(), StreamingArchiveLimits::default());
    let refused = writer
        .start_entry("a", CompressionMethod::Store)
        .err()
        .unwrap();
    assert_eq!(
        failure(&refused),
        "error=\"sink budget spent\" \
         progress=StreamingArchiveProgress { output_bytes: 10, poisoned: true } limit=None"
    );
    assert_eq!(sink.len(), 10);

    // The payload does not fit the sink.
    let sink = BudgetSink::new(36);
    let writer =
        StreamingArchiveWriter::with_writer(sink.clone(), StreamingArchiveLimits::default());
    let mut entry = writer.start_entry("a", CompressionMethod::Store).unwrap();
    let error = entry.write_all(b"payload").unwrap_err();
    assert_eq!(
        format!("{:?} {error}", error.kind()),
        "StorageFull sink budget spent"
    );
    let second = entry.write(b"payload").unwrap_err();
    assert_eq!(
        format!("{:?} {second}", second.kind()),
        "InvalidInput streaming ZIP entry writer is poisoned"
    );
    let refused = entry.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        "error=\"sink budget spent\" \
         progress=StreamingArchiveProgress { output_bytes: 36, poisoned: true } limit=None"
    );
    assert_eq!(sink.len(), 36);

    // The entry's compressed bytes exceed their limit when it finishes.
    let limits = small_limits().with_compressed_size_limit(3);
    let writer = StreamingArchiveWriter::new(limits);
    let mut entry = writer.start_entry("a", CompressionMethod::Store).unwrap();
    let outcome = entry.write(b"abcd").map_err(|error| error.to_string());
    assert_eq!(
        format!("{outcome:?}"),
        "Err(\"owned ZIP entry compressed limit exceeded: attempted 4, maximum 3\")"
    );
    let refused = entry.finish_with_progress().err().unwrap();
    assert_eq!(
        failure(&refused),
        "error=\"ZIP compressed member size limit exceeded: declared 4, maximum 3\" \
         progress=StreamingArchiveProgress { output_bytes: 31, poisoned: true } \
         limit=Some(StreamingLimitExceeded { resource: CompressedBytes, actual: 4, maximum: 3 })"
    );
}

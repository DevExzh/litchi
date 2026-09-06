use super::{GeneratedXmlEnvelope, GeneratedXmlLimits, PackageWriter, PackageWriterLimits};
use litchi_core::Error;
use soapberry_zip::office::ArchiveReader;
use std::cell::Cell;
use std::error::Error as StdError;
use std::io::{self, Read, Write};
use std::rc::Rc;

fn envelope() -> GeneratedXmlEnvelope {
    GeneratedXmlEnvelope::try_new(b"<root>", b"</root>").unwrap()
}

fn error_chain_contains(error: &dyn StdError, needle: &str) -> bool {
    let mut current = Some(error);
    while let Some(error) = current {
        if error.to_string().contains(needle) {
            return true;
        }
        current = error.source();
    }
    false
}

#[test]
fn generated_xml_is_read_back_and_has_one_manifest_record() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    let mut emitted = false;
    let report = writer
        .add_generated_xml(
            "content.xml",
            "text/xml",
            envelope(),
            GeneratedXmlLimits::default(),
            128,
            |output| {
                if emitted {
                    return Ok(false);
                }
                emitted = true;
                output
                    .write_all(b"<row id=\"1\">x</row>")
                    .map_err(Error::Io)?;
                Ok(true)
            },
        )
        .unwrap();
    assert_eq!(report.fragments(), 1);

    let bytes = writer.finish_to_bytes().unwrap();
    let archive = ArchiveReader::new(&bytes).unwrap();
    assert_eq!(
        archive.read("content.xml").unwrap(),
        b"<root><row id=\"1\">x</row></root>"
    );
    let manifest = archive.read("META-INF/manifest.xml").unwrap();
    let manifest = std::str::from_utf8(&manifest).unwrap();
    assert_eq!(
        manifest
            .matches("manifest:full-path=\"content.xml\"")
            .count(),
        1
    );
    assert_eq!(
        manifest.matches("manifest:media-type=\"text/xml\"").count(),
        1
    );
}

#[test]
fn first_producer_failure_preserves_source_and_finish_is_refused() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    let error = writer
        .add_generated_xml(
            "content.xml",
            "text/xml",
            envelope(),
            GeneratedXmlLimits::default(),
            128,
            |_output| Err(Error::InvalidFormat("first producer failure".to_string())),
        )
        .unwrap_err();
    assert!(error_chain_contains(&error, "first producer failure"));
    assert!(error.written().is_some_and(|written| written > 0));
    assert!(writer.finish_to_writer().is_err());
}

#[test]
fn later_producer_failure_preserves_source_and_finish_is_refused() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    let mut calls = 0usize;
    let error = writer
        .add_generated_xml(
            "content.xml",
            "text/xml",
            envelope(),
            GeneratedXmlLimits::default(),
            128,
            |output| {
                let call = calls;
                calls += 1;
                if call == 0 {
                    output.write_all(b"<row/>").map_err(Error::Io)?;
                    return Ok(true);
                }
                Err(Error::InvalidFormat("later producer failure".to_string()))
            },
        )
        .unwrap_err();
    assert!(calls >= 2);
    assert!(error_chain_contains(&error, "later producer failure"));
    assert!(error.written().is_some_and(|written| written > 0));
    assert!(writer.finish_to_writer().is_err());
}

#[test]
fn rejected_archive_budget_does_not_pull_the_first_fragment() {
    let limits = PackageWriterLimits {
        max_entries: 1,
        ..PackageWriterLimits::default()
    };
    let mut writer = PackageWriter::new_with_limits(limits);
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    let reads = Rc::new(Cell::new(0usize));
    let reads_for_callback = Rc::clone(&reads);
    let error = writer
        .add_generated_xml(
            "content.xml",
            "text/xml",
            envelope(),
            GeneratedXmlLimits::default(),
            128,
            move |_output| {
                reads_for_callback.set(reads_for_callback.get() + 1);
                Ok(false)
            },
        )
        .unwrap_err();
    assert_eq!(reads.get(), 0);
    assert!(error.limit().is_some());
}

struct CountingReader {
    reads: Rc<Cell<usize>>,
}

impl Read for CountingReader {
    fn read(&mut self, _output: &mut [u8]) -> io::Result<usize> {
        self.reads.set(self.reads.get() + 1);
        Ok(0)
    }
}

#[test]
fn opaque_reader_xml_refusal_still_precedes_source_read() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    let reads = Rc::new(Cell::new(0usize));
    let error = writer
        .add_file_reader_with_media_type(
            "content.xml",
            CountingReader {
                reads: Rc::clone(&reads),
            },
            "text/xml",
        )
        .unwrap_err();
    assert_eq!(reads.get(), 0);
    assert!(error.to_string().contains("rejects XML member"));
}

#[derive(Clone, Debug)]
struct ShortSinkControl {
    accepted: Rc<Cell<usize>>,
    fail_after: Rc<Cell<Option<usize>>>,
}

#[derive(Debug)]
struct ShortSink {
    control: ShortSinkControl,
}

impl Write for ShortSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        let accepted = self.control.accepted.get();
        let amount = self
            .control
            .fail_after
            .get()
            .map_or(input.len().min(7), |limit| {
                limit.saturating_sub(accepted).min(input.len()).min(7)
            });
        if amount == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test short sink failure",
            ));
        }
        self.control.accepted.set(accepted + amount);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn short_sink_failure_reports_accepted_progress() {
    let control = ShortSinkControl {
        accepted: Rc::new(Cell::new(0)),
        fail_after: Rc::new(Cell::new(None)),
    };
    let mut writer = PackageWriter::with_writer(ShortSink {
        control: control.clone(),
    });
    writer
        .set_mimetype_streaming("application/vnd.oasis.opendocument.text")
        .unwrap();
    control
        .fail_after
        .set(Some(control.accepted.get().saturating_add(128)));
    let mut emitted = 0usize;
    let result = writer.add_generated_xml(
        "content.xml",
        "text/xml",
        envelope(),
        GeneratedXmlLimits::default(),
        128,
        |output| {
            if emitted != 0 {
                return Ok(false);
            }
            emitted += 1;
            output.write_all(b"<row>payload</row>").map_err(Error::Io)?;
            Ok(true)
        },
    );
    let error = match result {
        Err(error) => error,
        Ok(_) => writer.finish_to_writer().unwrap_err(),
    };
    assert!(emitted > 0);
    assert!(error.written().is_some_and(|written| written > 0));
}

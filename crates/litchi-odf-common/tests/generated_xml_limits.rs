#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "These are fixed XML-limit boundary tests; unexpected results are failures."
)]

//! Typed XML audit boundaries through the public generated-member writer.

use litchi_core::Error;
use litchi_odf_common::core::{
    GeneratedXmlEnvelope, GeneratedXmlLimitResource as Resource, GeneratedXmlLimits,
    GeneratedXmlReport, PackageWriter, PackageWriterError,
};
use soapberry_zip::office::ArchiveReader;
use std::cell::{Cell, RefCell};
use std::io::{self, Write};
use std::rc::Rc;

const MIME: &str = "application/vnd.oasis.opendocument.text";

#[derive(Debug, Clone)]
struct SharedSink(Rc<RefCell<Vec<u8>>>);

impl Write for SharedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.0.borrow_mut().extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Return either the completed archive/report or the typed limit failure and
/// the exact bytes accepted by the caller-owned sink at that point.
fn publish(
    prefix: &[u8],
    suffix: &[u8],
    limits: GeneratedXmlLimits,
    max_fragment_bytes: usize,
    fragments: Vec<&[u8]>,
) -> Result<(Vec<u8>, GeneratedXmlReport), (PackageWriterError, u64, usize)> {
    let accepted = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(SharedSink(Rc::clone(&accepted)));
    writer
        .set_mimetype_streaming(MIME)
        .expect("fixed MIME preflight");
    let envelope = GeneratedXmlEnvelope::try_new(prefix, suffix).expect("fixed shell");
    let mime_accepted = accepted.borrow().len();
    let calls = Rc::new(Cell::new(0));
    let producer_calls = Rc::clone(&calls);
    let mut index = 0usize;
    let report = match writer.add_generated_xml(
        "content.xml",
        "text/xml",
        envelope,
        limits,
        max_fragment_bytes,
        move |output| {
            producer_calls.set(producer_calls.get() + 1);
            if index == fragments.len() {
                return Ok(false);
            }
            output.write_all(fragments[index]).map_err(Error::Io)?;
            index += 1;
            Ok(true)
        },
    ) {
        Ok(report) => report,
        Err(error) => {
            let written = accepted.borrow().len() as u64;
            if calls.get() == 0 {
                assert_eq!(
                    written, mime_accepted as u64,
                    "preflight admitted no content header"
                );
            }
            return Err((error, written, calls.get()));
        },
    };
    let sink = match writer.finish_to_writer() {
        Ok(sink) => sink,
        Err(error) => {
            let written = accepted.borrow().len() as u64;
            if calls.get() == 0 {
                assert_eq!(
                    written, mime_accepted as u64,
                    "preflight admitted no content header"
                );
            }
            return Err((error, written, calls.get()));
        },
    };
    let bytes = sink.0.borrow().clone();
    Ok((bytes, report))
}

fn narrowed(resource: Resource, maximum: usize) -> GeneratedXmlLimits {
    GeneratedXmlLimits::builder()
        .limit(resource, maximum)
        .unwrap()
        .build()
}

fn report_for(bytes: &[u8]) -> xml_minifier::audit::Report {
    xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default()).unwrap()
}

fn shell_bytes(prefix: &[u8], suffix: &[u8]) -> Vec<u8> {
    let mut bytes = prefix.to_vec();
    bytes.extend_from_slice(suffix);
    bytes
}

fn assert_limit_failure(
    result: Result<(Vec<u8>, GeneratedXmlReport), (PackageWriterError, u64, usize)>,
    resource: Resource,
    actual: usize,
    maximum: usize,
    expected_callbacks: Option<usize>,
) {
    let (error, accepted, calls) = result.expect_err("one-byte-under XML limit must fail");
    if let Some(expected) = expected_callbacks {
        assert_eq!(calls, expected);
    }
    let limit = error
        .xml_limit()
        .expect("generated XML limit attribution must survive ZIP/read adapters");
    assert_eq!(limit.resource(), resource);
    assert_eq!(limit.actual(), actual);
    assert_eq!(limit.maximum(), maximum);
    // The writer must report the same accepted prefix that the sink observed;
    // this assertion deliberately allows either zero or nonzero prefixes for
    // callers that reuse the helper at a different admission boundary.
    assert_eq!(error.written(), Some(accepted));
}

fn assert_shell_exact_and_under(
    prefix: &'static [u8],
    suffix: &'static [u8],
    resource: Resource,
    exact: usize,
) {
    assert!(exact > 0, "all shell dimensions in this test are nonzero");
    let successful = publish(prefix, suffix, narrowed(resource, exact), 1, Vec::new())
        .expect("inclusive shell limit");
    let archive = ArchiveReader::new(&successful.0).unwrap();
    assert_eq!(
        archive.read("content.xml").unwrap(),
        shell_bytes(prefix, suffix)
    );
    assert_eq!(successful.1.fragments(), 0);

    let under = publish(prefix, suffix, narrowed(resource, exact - 1), 1, Vec::new());
    assert_limit_failure(under, resource, exact, exact - 1, Some(0));
}

#[test]
fn shell_limits_are_inclusive_and_one_under_is_typed() {
    let bytes_shell = b"<root></root>";
    let bytes_report = report_for(bytes_shell);
    assert_shell_exact_and_under(b"<root>", b"</root>", Resource::Bytes, bytes_report.bytes());

    let depth_shell = b"<root><inner></inner></root>";
    let depth_report = report_for(depth_shell);
    assert_shell_exact_and_under(
        b"<root><inner>",
        b"</inner></root>",
        Resource::Depth,
        depth_report.max_depth(),
    );

    let events_report = report_for(bytes_shell);
    assert_shell_exact_and_under(
        b"<root>",
        b"</root>",
        Resource::Events,
        events_report.events(),
    );

    let attributes_shell = b"<root id=\"x\"></root>";
    let attributes_report = report_for(attributes_shell);
    assert_shell_exact_and_under(
        b"<root id=\"x\">",
        b"</root>",
        Resource::Attributes,
        attributes_report.attributes(),
    );

    let token_shell = b"<r></r>";
    let token_report = report_for(token_shell);
    // The longest lexical token is the four-byte `</r>` end event.  Keeping
    // this shell deliberately minimal makes the token boundary unambiguous.
    assert_shell_exact_and_under(b"<r>", b"</r>", Resource::TokenBytes, 4);
    assert_eq!(token_report.bytes(), token_shell.len());
}

#[test]
fn composed_fragment_aggregate_limits_are_inclusive_and_one_under() {
    const PREFIX: &[u8] = b"<root><body>";
    const SUFFIX: &[u8] = b"</body></root>";
    const FRAGMENT_A: &[u8] = b"<row a=\"x\"><cell b=\"y\">one</cell></row>";
    const FRAGMENT_B: &[u8] = b"<row c=\"z\"><cell>two</cell></row>";
    let fragments = vec![FRAGMENT_A, FRAGMENT_B];
    let fragment_capacity = FRAGMENT_A.len().max(FRAGMENT_B.len());
    let control = publish(
        PREFIX,
        SUFFIX,
        GeneratedXmlLimits::default(),
        fragment_capacity,
        fragments.clone(),
    )
    .expect("unlimited composed control");
    let report = control.1;
    let dimensions = [
        (Resource::Bytes, report.bytes()),
        (Resource::Depth, report.max_depth()),
        (Resource::Events, report.events()),
        (Resource::Attributes, report.attributes()),
        (Resource::TextBytes, report.text_bytes()),
    ];

    for (resource, exact) in dimensions {
        assert!(exact > 0);
        let successful = publish(
            PREFIX,
            SUFFIX,
            narrowed(resource, exact),
            fragment_capacity,
            fragments.clone(),
        )
        .expect("inclusive composed aggregate limit");
        assert_eq!(successful.1, report);
        let archive = ArchiveReader::new(&successful.0).unwrap();
        let expected = [PREFIX, FRAGMENT_A, FRAGMENT_B, SUFFIX].concat();
        assert_eq!(archive.read("content.xml").unwrap(), expected);

        let under = publish(
            PREFIX,
            SUFFIX,
            narrowed(resource, exact - 1),
            fragment_capacity,
            fragments.clone(),
        );
        assert_limit_failure(under, resource, exact, exact - 1, None);
    }
}

#[test]
fn token_limit_is_checked_per_fragment_even_when_aggregate_limits_allow_it() {
    const PREFIX: &[u8] = b"<root>";
    const SUFFIX: &[u8] = b"</root>";
    let fragment_storage = format!("<row a=\"{}\"/>", "x".repeat(31));
    let fragment = fragment_storage.as_bytes();
    let fragment_len = fragment.len();

    let successful = publish(
        PREFIX,
        SUFFIX,
        narrowed(Resource::TokenBytes, fragment_len),
        fragment_len,
        vec![fragment],
    )
    .expect("inclusive fragment token limit");
    let archive = ArchiveReader::new(&successful.0).unwrap();
    let expected = [PREFIX, fragment, SUFFIX].concat();
    assert_eq!(archive.read("content.xml").unwrap(), expected);

    let under = publish(
        PREFIX,
        SUFFIX,
        narrowed(Resource::TokenBytes, fragment_len - 1),
        fragment_len,
        vec![fragment],
    );
    assert_limit_failure(
        under,
        Resource::TokenBytes,
        fragment_len,
        fragment_len - 1,
        Some(1),
    );
}

#[test]
fn xml_limit_inspection_bounds_a_cyclic_caller_error_chain() {
    #[derive(Debug)]
    struct Cycle;
    impl std::fmt::Display for Cycle {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            f.write_str("cycle")
        }
    }
    impl std::error::Error for Cycle {
        fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
            Some(self)
        }
    }
    let error = PackageWriterError::Core(Error::Io(io::Error::other(Cycle)));
    assert!(error.xml_limit().is_none());
}

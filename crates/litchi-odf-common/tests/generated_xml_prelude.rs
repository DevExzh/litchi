#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "These are fixed generated-XML seam boundary tests."
)]

//! Public integration coverage for the opt-in generated-XML fixed-prelude seam.

use litchi_core::Error;
use litchi_odf_common::core::{
    GeneratedXmlEnvelope, GeneratedXmlLimitResource as Resource, GeneratedXmlLimits,
    GeneratedXmlReport, PackageWriter, PackageWriterError, PackageWriterLimits,
};
use soapberry_zip::office::ArchiveReader;
use std::cell::{Cell, RefCell};
use std::error::Error as StdError;
use std::io::{self, Write};
use std::rc::Rc;

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const MEDIA_TYPE: &str = "text/xml";

// This deliberately has both an Empty fixed child and a nested balanced
// child.  The fixed child reaches depth four; the final body/presentation
// insertion path reaches depth three and remains open at the boundary.
const PREFIX: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><root><scripts/><styles><style id="dp1"><leaf/></style></styles><body><presentation>"#;
const SUFFIX: &[u8] = b"</presentation></body></root>";
const FRAGMENT_ONE: &[u8] = b"<page id=\"1\"/>";
const FRAGMENT_TWO: &[u8] = b"<page id=\"2\"><title>two</title></page>";

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

fn report_for(bytes: &[u8]) -> xml_minifier::audit::Report {
    xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default()).unwrap()
}

fn expected_bytes(fragments: &[&[u8]]) -> Vec<u8> {
    let mut expected = PREFIX.to_vec();
    for fragment in fragments {
        expected.extend_from_slice(fragment);
    }
    expected.extend_from_slice(SUFFIX);
    expected
}

fn narrowed(resource: Resource, maximum: usize) -> GeneratedXmlLimits {
    GeneratedXmlLimits::builder()
        .limit(resource, maximum)
        .unwrap()
        .build()
}

fn assert_report_matches_audit(report: GeneratedXmlReport, expected: &[u8]) {
    let audited = report_for(expected);
    assert_eq!(report.bytes(), audited.bytes());
    assert_eq!(report.attributes(), audited.attributes());
    assert_eq!(report.events(), audited.events());
    assert_eq!(report.max_depth(), audited.max_depth());
    assert_eq!(report.text_bytes(), audited.text_bytes());
}

/// Return either the completed archive/report or the error, accepted sink
/// bytes, and callback count.  The envelope uses the new public constructor;
/// the strict constructor is tested separately below.
fn publish(
    limits: GeneratedXmlLimits,
    max_fragment_bytes: usize,
    fragments: Vec<&[u8]>,
) -> Result<(Vec<u8>, GeneratedXmlReport), (PackageWriterError, u64, usize)> {
    let accepted = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(SharedSink(Rc::clone(&accepted)));
    writer
        .set_mimetype_streaming(MIME)
        .expect("fixed MIME preflight");
    let envelope =
        GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).expect("fixed balanced prelude");
    let calls = Rc::new(Cell::new(0usize));
    let calls_for_callback = Rc::clone(&calls);
    let mut index = 0usize;
    let report = match writer.add_generated_xml(
        "content.xml",
        MEDIA_TYPE,
        envelope,
        limits,
        max_fragment_bytes,
        move |output| {
            calls_for_callback.set(calls_for_callback.get() + 1);
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
            return Err((error, written, calls.get()));
        },
    };
    let sink = match writer.finish_to_writer() {
        Ok(sink) => sink,
        Err(error) => {
            let written = accepted.borrow().len() as u64;
            return Err((error, written, calls.get()));
        },
    };
    Ok((sink.0.borrow().clone(), report))
}

fn assert_limit_failure(
    result: Result<(Vec<u8>, GeneratedXmlReport), (PackageWriterError, u64, usize)>,
    resource: Resource,
    actual: usize,
    maximum: usize,
    expected_calls: Option<usize>,
) {
    let (error, accepted, calls) = result.expect_err("one-under limit must fail");
    if let Some(expected) = expected_calls {
        assert_eq!(calls, expected);
    }
    let limit = error
        .xml_limit()
        .expect("typed generated XML limit must survive writer adapters");
    assert_eq!(limit.resource(), resource);
    assert_eq!(limit.actual(), actual);
    assert_eq!(limit.maximum(), maximum);
    assert_eq!(error.written(), Some(accepted));
}

#[test]
fn prelude_zero_fragments_publishes_exact_fixed_shell_and_full_audit() {
    let expected = expected_bytes(&[]);
    let (archive_bytes, report) = publish(GeneratedXmlLimits::default(), 1, Vec::new())
        .expect("zero-fragment prelude publication");
    let archive = ArchiveReader::new(&archive_bytes).unwrap();
    assert_eq!(archive.read("content.xml").unwrap(), expected);
    assert_eq!(report.fragments(), 0);
    assert_report_matches_audit(report, &expected);
}

#[test]
fn prelude_two_fragments_preserve_order_and_every_audit_counter() {
    let fragments = vec![FRAGMENT_ONE, FRAGMENT_TWO];
    let expected = expected_bytes(&fragments);
    let (archive_bytes, report) = publish(
        GeneratedXmlLimits::default(),
        FRAGMENT_ONE.len().max(FRAGMENT_TWO.len()),
        fragments,
    )
    .expect("two-fragment prelude publication");
    let archive = ArchiveReader::new(&archive_bytes).unwrap();
    assert_eq!(archive.read("content.xml").unwrap(), expected);
    assert_eq!(report.fragments(), 2);
    assert_report_matches_audit(report, &expected);
}

#[test]
fn fixed_balanced_child_and_insertion_path_both_contribute_to_depth() {
    let shell = report_for(&[PREFIX, SUFFIX].concat());
    assert_eq!(shell.max_depth(), 4, "the fixed leaf reaches depth four");

    let fragments = vec![b"<page><title><run/></title></page>".as_slice()];
    let expected = expected_bytes(&fragments);
    let (_, report) = publish(GeneratedXmlLimits::default(), fragments[0].len(), fragments)
        .expect("depth fixture");
    assert_report_matches_audit(report, &expected);
    assert!(report.max_depth() > shell.max_depth());
    // The open root/body/presentation path has depth three and the fragment
    // has depth three, so the composed path reaches six.
    assert_eq!(report.max_depth(), 6);
}

#[test]
fn empty_fixed_child_is_opt_in_and_strict_constructor_stays_strict() {
    assert!(GeneratedXmlEnvelope::try_new(PREFIX, SUFFIX).is_err());
    assert!(GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).is_ok());
}

fn assert_rejected(prefix: &[u8], suffix: &[u8]) {
    let error = GeneratedXmlEnvelope::try_new_with_prelude(prefix, suffix)
        .expect_err("malformed fixed shell must be refused");
    assert!(
        matches!(error, Error::InvalidFormat(_)),
        "unexpectedly accepted prelude: {:?} + {:?}",
        String::from_utf8_lossy(prefix),
        String::from_utf8_lossy(suffix),
    );
}

#[test]
fn fixed_children_under_open_ancestors_precede_the_final_start_sequence() {
    for prefix in [
        b"<root><fixed></fixed><slot>".as_slice(),
        b"<root><slot><fixed/><inner>".as_slice(),
        b"<root><slot><fixed></fixed><inner>".as_slice(),
    ] {
        let suffix = if prefix.ends_with(b"<inner>") {
            b"</inner></slot></root>".as_slice()
        } else {
            b"</slot></root>".as_slice()
        };
        assert!(GeneratedXmlEnvelope::try_new_with_prelude(prefix, suffix).is_ok());
    }
    assert_rejected(b"<root><slot><fixed/>", b"</slot></root>");
    assert_rejected(b"<root><fixed></fixed>", b"</root>");
}

#[test]
fn fixed_shell_depth_can_dominate_a_shallow_fragment() {
    let prefix = b"<root><fixed><a><b><c/></b></a></fixed><slot>";
    let suffix = b"</slot></root>";
    let mut writer = PackageWriter::new();
    writer.set_mimetype_streaming(MIME).unwrap();
    let mut emitted = false;
    let report = writer
        .add_generated_xml(
            "content.xml",
            MEDIA_TYPE,
            GeneratedXmlEnvelope::try_new_with_prelude(prefix, suffix).unwrap(),
            GeneratedXmlLimits::default(),
            7,
            |output| {
                if emitted {
                    return Ok(false);
                }
                output.write_all(b"<page/>").map_err(Error::Io)?;
                emitted = true;
                Ok(true)
            },
        )
        .unwrap();
    assert_eq!(report.max_depth(), 5);
    let expected = [prefix.as_slice(), b"<page/>", suffix.as_slice()].concat();
    assert_report_matches_audit(report, &expected);
    let archive = writer.finish_to_writer().unwrap();
    assert_eq!(
        ArchiveReader::new(archive.get_ref())
            .unwrap()
            .read("content.xml")
            .unwrap(),
        expected
    );
}

#[test]
fn prelude_rejects_forbidden_fixed_events_and_xml_space() {
    let cases: &[(&[u8], &[u8])] = &[
        (b"<root>text<slot>", b"</slot></root>"),
        (b"<root><![CDATA[text]]><slot>", b"</slot></root>"),
        (b"<root>&amp;<slot>", b"</slot></root>"),
        (b"<root><!-- comment --><slot>", b"</slot></root>"),
        (b"<root><?processing?><slot>", b"</slot></root>"),
        (b"<root><fixed>text</fixed><slot>", b"</slot></root>"),
        (
            b"<root><fixed><![CDATA[text]]></fixed><slot>",
            b"</slot></root>",
        ),
        (b"<root><fixed>&amp;</fixed><slot>", b"</slot></root>"),
        (
            b"<root><fixed><!-- comment --></fixed><slot>",
            b"</slot></root>",
        ),
        (
            b"<root><fixed><?processing?></fixed><slot>",
            b"</slot></root>",
        ),
        (b"<!DOCTYPE root><root><slot>", b"</slot></root>"),
        (
            b"<?xml version=\"1.0\"?><?xml version=\"1.0\"?><root><slot>",
            b"</slot></root>",
        ),
        (
            b"<root><fixed xml:space=\"preserve\"/><slot>",
            b"</slot></root>",
        ),
        (
            b"<root><fixed xml:space=\"preserve\"><leaf/></fixed><slot>",
            b"</slot></root>",
        ),
        (b"<root><slot xml:space=\"default\">", b"</slot></root>"),
    ];
    for (prefix, suffix) in cases {
        assert_rejected(prefix, suffix);
    }
}

#[test]
fn prelude_rejects_suffix_events_roots_and_qname_mismatches() {
    let cases: &[(&[u8], &[u8])] = &[
        (b"<root><slot>", b"<tail/></slot></root>"),
        (b"<root><slot>", b"text</slot></root>"),
        (b"<root><slot>", b"&amp;</slot></root>"),
        (b"<root><slot>", b"<!-- comment --></slot></root>"),
        (b"<root><slot>", b"</other></root>"),
        (b"<root><slot><inner>", b"</slot></inner></root>"),
        (b"<root><slot>", b"</slot></other>"),
        (b"<root><fixed/></root><other><slot>", b"</slot></other>"),
        (b"<fixed/><root><slot>", b"</slot></root>"),
        (b"<root><fixed/></root>", b""),
        (
            b"<root xmlns:a=\"urn:x\" xmlns:b=\"urn:x\"><a:fixed></b:fixed><slot>",
            b"</slot></root>",
        ),
    ];
    for (prefix, suffix) in cases {
        assert_rejected(prefix, suffix);
    }
}

#[test]
fn prelude_rejects_boundaries_inside_xml_events_and_utf8() {
    let cases: &[(&[u8], &[u8])] = &[
        (b"<root><sl", b"ot></root>"),
        (b"<root><fixed/", b"><slot></root>"),
        (b"<root><fixed value=\"\xC3", b"\xA9\"/><slot></root>"),
        (b"<root><slot", b"></root>"),
        (b"<root><slot></ro", b"ot>"),
    ];
    for (prefix, suffix) in cases {
        assert_rejected(prefix, suffix);
    }
}

fn assert_shell_limit(resource: Resource, exact: usize) {
    assert!(exact > 0);
    let exact_result = publish(narrowed(resource, exact), 1, Vec::new())
        .expect("inclusive fixed-prelude shell limit");
    let archive = ArchiveReader::new(&exact_result.0).unwrap();
    assert_eq!(archive.read("content.xml").unwrap(), expected_bytes(&[]));
    let under = publish(narrowed(resource, exact - 1), 1, Vec::new());
    assert_limit_failure(under, resource, exact, exact - 1, Some(0));
}

#[test]
fn fixed_prelude_shell_limits_are_inclusive_and_one_under_before_callback() {
    let shell = report_for(&expected_bytes(&[]));
    assert_shell_limit(Resource::Bytes, shell.bytes());
    assert_shell_limit(Resource::Depth, shell.max_depth());
    assert_shell_limit(Resource::Events, shell.events());
    assert_shell_limit(Resource::Attributes, shell.attributes());
}

#[test]
fn composed_prelude_limits_are_inclusive_and_one_under() {
    let fragments = vec![FRAGMENT_ONE, FRAGMENT_TWO];
    let expected = expected_bytes(&fragments);
    let control = publish(
        GeneratedXmlLimits::default(),
        FRAGMENT_ONE.len().max(FRAGMENT_TWO.len()),
        fragments.clone(),
    )
    .expect("unlimited prelude aggregate control");
    let report = control.1;
    assert_report_matches_audit(report, &expected);
    let dimensions = [
        (Resource::Bytes, report.bytes()),
        (Resource::Depth, report.max_depth()),
        (Resource::Events, report.events()),
        (Resource::Attributes, report.attributes()),
        (Resource::TextBytes, report.text_bytes()),
    ];
    for (resource, exact) in dimensions {
        assert!(exact > 0);
        let success = publish(
            narrowed(resource, exact),
            FRAGMENT_ONE.len().max(FRAGMENT_TWO.len()),
            fragments.clone(),
        )
        .expect("inclusive composed prelude limit");
        assert_eq!(success.1, report);
        let archive = ArchiveReader::new(&success.0).unwrap();
        assert_eq!(archive.read("content.xml").unwrap(), expected);

        let under = publish(
            narrowed(resource, exact - 1),
            FRAGMENT_ONE.len().max(FRAGMENT_TWO.len()),
            fragments.clone(),
        );
        assert_limit_failure(under, resource, exact, exact - 1, None);
    }
}

#[test]
fn archive_admission_refusal_does_not_pull_a_prelude_fragment() {
    let limits = PackageWriterLimits {
        max_entries: 1,
        ..PackageWriterLimits::default()
    };
    let accepted = Rc::new(RefCell::new(Vec::new()));
    let mut writer =
        PackageWriter::with_writer_and_limits(SharedSink(Rc::clone(&accepted)), limits);
    writer.set_mimetype_streaming(MIME).unwrap();
    let mime_bytes = accepted.borrow().len();
    let calls = Rc::new(Cell::new(0usize));
    let calls_for_callback = Rc::clone(&calls);
    let envelope = GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).unwrap();
    let error = writer
        .add_generated_xml(
            "content.xml",
            MEDIA_TYPE,
            envelope,
            GeneratedXmlLimits::default(),
            128,
            move |_output| {
                calls_for_callback.set(calls_for_callback.get() + 1);
                Ok(false)
            },
        )
        .unwrap_err();
    assert_eq!(calls.get(), 0);
    assert_eq!(accepted.borrow().len(), mime_bytes);
    assert!(error.limit().is_some());
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
fn first_and_later_prelude_producer_errors_keep_source_and_poison_writer() {
    let mut first_writer = PackageWriter::new();
    first_writer.set_mimetype_streaming(MIME).unwrap();
    let first_error = first_writer
        .add_generated_xml(
            "content.xml",
            MEDIA_TYPE,
            GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).unwrap(),
            GeneratedXmlLimits::default(),
            128,
            |_output| Err(Error::InvalidFormat("prelude first producer".to_string())),
        )
        .unwrap_err();
    assert!(error_chain_contains(&first_error, "prelude first producer"));
    assert!(first_error.written().is_some_and(|written| written > 0));
    assert!(first_writer.finish_to_writer().is_err());

    let mut later_writer = PackageWriter::new();
    later_writer.set_mimetype_streaming(MIME).unwrap();
    let mut calls = 0usize;
    let later_error = later_writer
        .add_generated_xml(
            "content.xml",
            MEDIA_TYPE,
            GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).unwrap(),
            GeneratedXmlLimits::default(),
            128,
            |output| {
                let call = calls;
                calls += 1;
                if call == 0 {
                    output.write_all(FRAGMENT_ONE).map_err(Error::Io)?;
                    return Ok(true);
                }
                Err(Error::InvalidFormat("prelude later producer".to_string()))
            },
        )
        .unwrap_err();
    assert!(calls >= 2);
    assert!(error_chain_contains(&later_error, "prelude later producer"));
    assert!(later_error.written().is_some_and(|written| written > 0));
    assert!(later_writer.finish_to_writer().is_err());
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
                "prelude short sink failure",
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
fn short_sink_crossing_fixed_prelude_reports_exact_accepted_progress() {
    let control = ShortSinkControl {
        accepted: Rc::new(Cell::new(0)),
        fail_after: Rc::new(Cell::new(None)),
    };
    let mut writer = PackageWriter::with_writer(ShortSink {
        control: control.clone(),
    });
    writer.set_mimetype_streaming(MIME).unwrap();
    control
        .fail_after
        .set(Some(control.accepted.get().saturating_add(128)));
    let mut emitted = false;
    let result = writer.add_generated_xml(
        "content.xml",
        MEDIA_TYPE,
        GeneratedXmlEnvelope::try_new_with_prelude(PREFIX, SUFFIX).unwrap(),
        GeneratedXmlLimits::default(),
        128,
        |output| {
            if emitted {
                return Ok(false);
            }
            emitted = true;
            output.write_all(FRAGMENT_ONE).map_err(Error::Io)?;
            Ok(true)
        },
    );
    let error = match result {
        Err(error) => error,
        Ok(_) => writer.finish_to_writer().unwrap_err(),
    };
    assert!(emitted);
    assert_eq!(
        error.written(),
        Some(control.accepted.get() as u64),
        "ZIP progress must equal the short sink's accepted prefix"
    );
}

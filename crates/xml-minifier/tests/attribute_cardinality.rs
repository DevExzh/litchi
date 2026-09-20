#![allow(
    clippy::unwrap_used,
    reason = "focused auditor tests report the unexpected result directly"
)]

use std::io::{self, BufRead, Read};

use xml_minifier::audit::{self, Error, Kind, Limits, Report, Resource, StreamError};

static ONE_BYTE: &[usize] = &[1];
static TWO_BYTES: &[usize] = &[2];
static THREE_BYTES: &[usize] = &[3];
static FIVE_BYTES: &[usize] = &[5];
static UNEVEN_BYTES: &[usize] = &[2, 1, 3];
static PREFIX_BYTES: &[usize] = &[7, 1, 4, 2];

fn chunk_patterns() -> [&'static [usize]; 6] {
    [
        ONE_BYTE,
        TWO_BYTES,
        THREE_BYTES,
        FIVE_BYTES,
        UNEVEN_BYTES,
        PREFIX_BYTES,
    ]
}

fn limits_for(input: &[u8]) -> Limits {
    Limits::new(
        input.len(),
        32,
        256,
        128,
        input.len().max(1),
        input.len().max(1),
    )
    .unwrap()
}

fn limits_with_attributes(input: &[u8], maximum: usize) -> Limits {
    Limits::new(
        input.len(),
        32,
        256,
        maximum,
        input.len().max(1),
        input.len().max(1),
    )
    .unwrap()
}

fn authored_slice(input: &[u8], limits: Limits) -> Result<Report, Error> {
    audit::verify_authored(input, limits)
}

fn source_slice(input: &[u8], limits: Limits) -> Result<Report, Error> {
    audit::verify_source(input, limits)
}

fn compact_slice(input: &[u8], limits: Limits) -> Result<Report, Error> {
    audit::verify(input, limits)
}

fn authored_stream<R: BufRead>(reader: R, limits: Limits) -> Result<Report, Error> {
    flatten_stream(audit::verify_authored_reader(reader, limits))
}

fn compact_stream<R: BufRead>(reader: R, limits: Limits) -> Result<Report, Error> {
    flatten_stream(audit::verify_reader(reader, limits))
}

fn flatten_stream(result: Result<Report, StreamError>) -> Result<Report, Error> {
    match result {
        Ok(report) => Ok(report),
        Err(StreamError::Audit(error)) => Err(error),
        Err(StreamError::Input(error)) => panic!("unexpected input error: {error}"),
        Err(_) => panic!("unexpected non-exhaustive stream error"),
    }
}

fn assert_authored_stream_matches(input: &[u8], limits: Limits) {
    let expected = authored_slice(input, limits);
    for chunks in chunk_patterns() {
        let actual = authored_stream(Chunked::new(input, chunks), limits);
        assert_equivalent(&expected, &actual, chunks, true);
    }
}

fn assert_compact_stream_matches(input: &[u8], limits: Limits) {
    let expected = compact_slice(input, limits);
    for chunks in chunk_patterns() {
        let actual = compact_stream(Chunked::new(input, chunks), limits);
        assert_equivalent(&expected, &actual, chunks, true);
    }
}

fn assert_equivalent(
    expected: &Result<Report, Error>,
    actual: &Result<Report, Error>,
    chunks: &[usize],
    compare_offset: bool,
) {
    match (expected, actual) {
        (Ok(expected), Ok(actual)) => assert_eq!(expected, actual, "chunks {chunks:?}"),
        (Err(expected), Err(actual)) => {
            assert_error_equivalent(expected, actual, chunks, compare_offset)
        },
        (Ok(expected), Err(actual)) => {
            panic!("slice success {expected:?} became stream error {actual:?} for {chunks:?}")
        },
        (Err(expected), Ok(actual)) => {
            panic!("slice error {expected:?} became stream success {actual:?} for {chunks:?}")
        },
    }
}

fn assert_error_equivalent(
    expected: &Error,
    actual: &Error,
    chunks: &[usize],
    compare_offset: bool,
) {
    match (expected, actual) {
        (
            Error::Limit {
                resource: expected_resource,
                limit: expected_limit,
                actual: expected_actual,
                offset: expected_offset,
            },
            Error::Limit {
                resource: actual_resource,
                limit: actual_limit,
                actual: actual_actual,
                offset: actual_offset,
            },
        ) => {
            assert_eq!(expected_resource, actual_resource, "chunks {chunks:?}");
            assert_eq!(expected_limit, actual_limit, "chunks {chunks:?}");
            assert_eq!(expected_actual, actual_actual, "chunks {chunks:?}");
            if compare_offset {
                assert_eq!(expected_offset, actual_offset, "chunks {chunks:?}");
            }
        },
        (
            Error::Encoding {
                valid_up_to: expected,
            },
            Error::Encoding {
                valid_up_to: actual,
            },
        ) => assert_eq!(expected, actual, "chunks {chunks:?}"),
        (Error::NotCompact(expected), Error::NotCompact(actual)) => {
            assert_eq!(expected, actual, "chunks {chunks:?}")
        },
        (Error::Doctype { offset: expected }, Error::Doctype { offset: actual }) => {
            assert_eq!(expected, actual, "chunks {chunks:?}");
        },
        (
            Error::Malformed {
                offset: expected, ..
            },
            Error::Malformed { offset: actual, .. },
        ) => assert_eq!(expected, actual, "chunks {chunks:?}"),
        (Error::Allocation, Error::Allocation) => {},
        (expected, actual) => panic!(
            "error category differs for chunks {chunks:?}: expected {expected:?}, actual {actual:?}"
        ),
    }
}

fn assert_limit(error: &Error, resource: Resource, limit: usize, actual: usize) {
    match error {
        Error::Limit {
            resource: found_resource,
            limit: found_limit,
            actual: found_actual,
            ..
        } => {
            assert_eq!(found_resource, &resource);
            assert_eq!(*found_limit, limit);
            assert_eq!(*found_actual, actual);
        },
        other => panic!("expected {resource:?} limit, got {other:?}"),
    }
}

fn assert_malformed(error: &Error, detail: &str) -> usize {
    match error {
        Error::Malformed {
            offset,
            detail: found,
        } => {
            assert!(
                found.contains(detail),
                "expected malformed detail containing {detail:?}, got {found:?}"
            );
            *offset
        },
        other => panic!("expected malformed XML, got {other:?}"),
    }
}

fn malformed_offset(error: &Error) -> usize {
    match error {
        Error::Malformed { offset, .. } => *offset,
        other => panic!("expected malformed XML, got {other:?}"),
    }
}

fn assert_not_compact(error: &Error, kind: Kind) -> usize {
    match error {
        Error::NotCompact(violation) => {
            assert_eq!(violation.kind(), kind);
            violation.offset()
        },
        other => panic!("expected {kind:?} compactness error, got {other:?}"),
    }
}

#[test]
fn attribute_cardinality_reports_zero_one_two_many_and_aggregate_values() {
    let fixtures = [
        (b"<root/>".as_slice(), 0),
        (br#"<root a="1"/>"#.as_slice(), 1),
        (br#"<root a="1" b="2"/>"#.as_slice(), 2),
        (
            br#"<root a="1" b="2" c="3" d="4" e="5" f="6"/>"#.as_slice(),
            6,
        ),
        (
            br#"<root a="1"><child b="2" c="3"/><leaf d="4"/></root>"#.as_slice(),
            4,
        ),
        (br#"<p:root xmlns:p="urn:test" p:value="v"/>"#.as_slice(), 2),
    ];

    for (xml, attributes) in fixtures {
        let limits = limits_for(xml);
        let source = source_slice(xml, limits).expect("source XML must be structurally valid");
        let authored = authored_slice(xml, limits).expect("authored XML must be compact");
        let compact = compact_slice(xml, limits).expect("default XML must be compact");
        assert_eq!(source.attributes(), attributes, "source fixture {xml:?}");
        assert_eq!(
            authored.attributes(),
            attributes,
            "authored fixture {xml:?}"
        );
        assert_eq!(compact.attributes(), attributes, "default fixture {xml:?}");
        assert_eq!(source, authored, "source/authored fixture {xml:?}");
        assert_eq!(authored, compact, "authored/default fixture {xml:?}");
        assert_authored_stream_matches(xml, limits);
        assert_compact_stream_matches(xml, limits);
    }
}

#[test]
fn duplicate_attributes_are_refused_before_and_after_many_attribute_scans() {
    let cases = [
        (br#"<root a="1" a="2" b="3"/>"#.as_slice(), 0),
        (br#"<root a="1" b="2" c="3" d="4" a="5"/>"#.as_slice(), 0),
        (
            br#"<root><child a="1" b="2" c="3" a="4"/></root>"#.as_slice(),
            6,
        ),
        (b"\xef\xbb\xbf<root a=\"1\" b=\"2\" a=\"3\"/>".as_slice(), 3),
    ];

    for (xml, expected_offset) in cases {
        let limits = limits_for(xml);
        let source = source_slice(xml, limits).unwrap_err();
        let authored = authored_slice(xml, limits).unwrap_err();
        let compact = compact_slice(xml, limits).unwrap_err();
        assert_eq!(
            assert_malformed(&source, "duplicated attribute"),
            expected_offset
        );
        assert_eq!(
            assert_malformed(&authored, "duplicated attribute"),
            expected_offset
        );
        assert_eq!(
            assert_malformed(&compact, "duplicated attribute"),
            expected_offset
        );
        assert_authored_stream_matches(xml, limits);
        assert_compact_stream_matches(xml, limits);
    }
}

#[test]
fn one_malformed_attribute_remains_structural_under_every_policy_and_chunking() {
    let cases = [
        (b"<root a/>".as_slice(), "attribute"),
        (b"<root a=1/>".as_slice(), "quoted"),
        (b"<root a=\"unterminated/>".as_slice(), "unterminated"),
        (b"<root a=\"1\"b=\"2\"/>".as_slice(), "separator"),
    ];

    for (xml, detail) in cases {
        let limits = limits_for(xml);
        let source = source_slice(xml, limits).unwrap_err();
        let authored = authored_slice(xml, limits).unwrap_err();
        let compact = compact_slice(xml, limits).unwrap_err();
        let source_offset = malformed_offset(&source);
        if detail == "attribute" {
            // The compact contracts classify a bare name as an attribute
            // separation defect before quick-xml can report its missing `=`;
            // the source contract must retain the structural refusal.
            assert_not_compact(&authored, Kind::AttributeSeparation);
            assert_not_compact(&compact, Kind::AttributeSeparation);
        } else if detail == "unterminated" {
            // Both parsers refuse the unclosed value, but their public
            // EOF offsets differ. Preserve each existing route's offset.
            assert_eq!(source_offset, 23);
            assert_eq!(malformed_offset(&authored), 23);
            assert_eq!(malformed_offset(&compact), 23);
            assert_malformed(&source, "attribute value not closed");
            assert_malformed(&authored, "attribute value not closed");
            assert_malformed(&compact, "attribute value not closed");
            assert_stream_malformed_at(xml, limits, true, 0, "attribute value not closed");
            assert_stream_malformed_at(xml, limits, false, 0, "attribute value not closed");
            continue;
        } else {
            assert_eq!(malformed_offset(&authored), source_offset);
            assert_eq!(malformed_offset(&compact), source_offset);
        }
        assert_authored_stream_matches(xml, limits);
        assert_compact_stream_matches(xml, limits);
    }
}

fn assert_stream_malformed_at(
    input: &[u8],
    limits: Limits,
    authored: bool,
    expected_offset: usize,
    detail: &str,
) {
    for chunks in chunk_patterns() {
        let actual = if authored {
            authored_stream(Chunked::new(input, chunks), limits)
        } else {
            compact_stream(Chunked::new(input, chunks), limits)
        }
        .unwrap_err();
        assert_eq!(
            malformed_offset(&actual),
            expected_offset,
            "chunks {chunks:?}"
        );
        assert_malformed(&actual, detail);
    }
}

#[test]
fn xml_space_inheritance_and_escaped_values_are_checked_after_cardinality() {
    let preserved = b"<root xml:space=\"preserve\">\n<child a=\"1\">\n</child></root>";
    let limits = limits_for(preserved);
    let source = source_slice(preserved, limits).unwrap();
    let authored = authored_slice(preserved, limits).unwrap();
    assert_eq!(source.attributes(), 2);
    assert_eq!(authored, source);
    assert_authored_stream_matches(preserved, limits);

    let escaped = br#"<root xml:space="pre&#115;erve">
<child/>
</root>"#;
    let limits = limits_for(escaped);
    let report = authored_slice(escaped, limits).unwrap();
    assert_eq!(report.attributes(), 1);
    assert_authored_stream_matches(escaped, limits);

    let reset = b"<root xml:space=\"preserve\"><child xml:space=\"default\">\n</child></root>";
    let limits = limits_for(reset);
    let source = source_slice(reset, limits).unwrap();
    assert_eq!(source.attributes(), 2);
    let authored = authored_slice(reset, limits).unwrap_err();
    let expected_offset = assert_not_compact(&authored, Kind::FormattingWhitespace);
    for chunks in chunk_patterns() {
        let actual = authored_stream(Chunked::new(reset, chunks), limits).unwrap_err();
        let actual_offset = assert_not_compact(&actual, Kind::FormattingWhitespace);
        assert_eq!(actual_offset, expected_offset, "chunks {chunks:?}");
    }

    let invalid = br#"<root xml:space="maybe"/>"#;
    let limits = limits_for(invalid);
    let source = source_slice(invalid, limits).unwrap_err();
    let authored = authored_slice(invalid, limits).unwrap_err();
    let source_offset = assert_malformed(&source, "xml:space");
    assert_eq!(assert_malformed(&authored, "xml:space"), source_offset);
    assert_authored_stream_matches(invalid, limits);

    let invalid_inherited = br#"<root xml:space="preserve"><child xml:space="bad">
</child></root>"#;
    let limits = limits_for(invalid_inherited);
    let error = authored_slice(invalid_inherited, limits).unwrap_err();
    let offset = assert_malformed(&error, "xml:space");
    let child_offset = invalid_inherited
        .windows(b"<child".len())
        .position(|window| window == b"<child")
        .unwrap();
    assert_eq!(offset, child_offset,);
    assert_authored_stream_matches(invalid_inherited, limits);

    let escaped_invalid = br#"<root xml:space="pre&#120;erve"/>"#;
    let limits = limits_for(escaped_invalid);
    let error = source_slice(escaped_invalid, limits).unwrap_err();
    assert_malformed(&error, "xml:space");
    assert_authored_stream_matches(escaped_invalid, limits);
}

#[test]
fn malformed_duplicate_precedes_later_xml_space_value_error() {
    let xml = br#"<root xml:space="preserve" xml:space="keep"/>"#;
    let limits = limits_for(xml);
    for result in [
        source_slice(xml, limits),
        authored_slice(xml, limits),
        compact_slice(xml, limits),
    ] {
        let error = result.unwrap_err();
        assert_malformed(&error, "duplicated attribute");
    }
    assert_authored_stream_matches(xml, limits);
    assert_compact_stream_matches(xml, limits);

    let invalid = br#"<root xml:space="keep" xml:space="preserve"/>"#;
    let error = source_slice(invalid, limits_for(invalid)).unwrap_err();
    assert_malformed(&error, "xml:space must");
}

#[test]
fn source_policy_accepts_producer_spacing_but_authored_policy_keeps_compactness_errors() {
    let xml = b"\xef\xbb\xbf<?xml version=\"1.0\"?>\r\n<root a = \"1\"  b=\"2\" >\n  <child c=\"3\" />\n</root >";
    let limits = limits_for(xml);
    let source = source_slice(xml, limits).expect("source policy must accept producer spacing");
    assert_eq!(source.attributes(), 3);
    assert_eq!(source.bytes(), xml.len());

    let default_error = compact_slice(xml, limits).unwrap_err();
    assert_not_compact(&default_error, Kind::FormattingWhitespace);
    let authored_error = authored_slice(xml, limits).unwrap_err();
    assert_not_compact(&authored_error, Kind::FormattingWhitespace);
    assert_compact_stream_matches(xml, limits);
    assert_authored_stream_matches(xml, limits);

    let compact_bom = b"\xef\xbb\xbf<root a=\"1\"/>";
    let limits = limits_for(compact_bom);
    for report in [
        source_slice(compact_bom, limits).unwrap(),
        authored_slice(compact_bom, limits).unwrap(),
        compact_slice(compact_bom, limits).unwrap(),
    ] {
        assert_eq!(report.attributes(), 1);
        assert_eq!(report.bytes(), compact_bom.len());
    }
    assert_authored_stream_matches(compact_bom, limits);
    assert_compact_stream_matches(compact_bom, limits);
}

#[test]
fn bom_offsets_and_encoding_errors_remain_physical_source_offsets() {
    let duplicate = b"\xef\xbb\xbf<root a=\"1\" a=\"2\"/>";
    let limits = limits_for(duplicate);
    let source = source_slice(duplicate, limits).unwrap_err();
    assert_eq!(assert_malformed(&source, "duplicated attribute"), 3);
    assert_authored_stream_matches(duplicate, limits);
    assert_compact_stream_matches(duplicate, limits);

    let doctype = b"\xef\xbb\xbf<!DOCTYPE root><root/>";
    let limits = limits_for(doctype);
    let source = source_slice(doctype, limits).unwrap_err();
    match source {
        Error::Doctype { offset } => assert_eq!(offset, 3),
        other => panic!("expected BOM-adjusted doctype offset, got {other:?}"),
    }
    assert_authored_stream_matches(doctype, limits);
    assert_compact_stream_matches(doctype, limits);

    let invalid = b"\xef\xbb\xbf<root>\xff</root>";
    let limits = limits_for(invalid);
    let source = source_slice(invalid, limits).unwrap_err();
    match source {
        Error::Encoding { valid_up_to } => {
            assert_eq!(valid_up_to, b"\xef\xbb\xbf<root>".len())
        },
        other => panic!("expected BOM-adjusted encoding offset, got {other:?}"),
    }
    assert_authored_stream_matches(invalid, limits);
    assert_compact_stream_matches(invalid, limits);
}

#[test]
fn attribute_and_document_resource_boundaries_are_inclusive_and_chunk_stable() {
    let xml = b"<root a=\"1\"><child b=\"2\">text</child></root>";
    let base = limits_for(xml);
    let report = authored_slice(xml, base).unwrap();
    assert_eq!(report.attributes(), 2);
    assert_eq!(report.max_depth(), 2);
    assert_eq!(report.text_bytes(), 4);

    // The exact attribute boundary must accept both attributes, while one
    // fewer fails at the child start tag rather than being silently ignored.
    let exact_attributes = limits_with_attributes(xml, report.attributes());
    assert_eq!(authored_slice(xml, exact_attributes).unwrap(), report);
    assert_authored_stream_matches(xml, exact_attributes);
    let short_attributes = limits_with_attributes(xml, report.attributes() - 1);
    let expected = authored_slice(xml, short_attributes).unwrap_err();
    assert_limit(
        &expected,
        Resource::Attributes,
        report.attributes() - 1,
        report.attributes(),
    );
    assert_authored_stream_matches(xml, short_attributes);

    let boundaries = [
        (
            Resource::Depth,
            base.narrow(Resource::Depth, report.max_depth()),
            base.narrow(Resource::Depth, report.max_depth() - 1),
        ),
        (
            Resource::Events,
            base.narrow(Resource::Events, report.events()),
            base.narrow(Resource::Events, report.events() - 1),
        ),
        (
            Resource::TextBytes,
            base.narrow(Resource::TextBytes, report.text_bytes()),
            base.narrow(Resource::TextBytes, report.text_bytes() - 1),
        ),
    ];
    for (resource, exact, short) in boundaries {
        assert_eq!(authored_slice(xml, exact).unwrap(), report);
        assert_authored_stream_matches(xml, exact);
        let expected = authored_slice(xml, short).unwrap_err();
        let (limit, actual) = match resource {
            Resource::Depth => (report.max_depth() - 1, report.max_depth()),
            Resource::Events => (report.events() - 1, report.events()),
            Resource::TextBytes => (report.text_bytes() - 1, report.text_bytes()),
            _ => unreachable!(),
        };
        assert_limit(&expected, resource, limit, actual);
        assert_authored_stream_matches(xml, short);
    }

    let largest_start_tag = b"<child b=\"2\">".len();
    let exact_token = base.narrow(Resource::TokenBytes, largest_start_tag);
    assert_eq!(authored_slice(xml, exact_token).unwrap(), report);
    assert_authored_stream_matches(xml, exact_token);
    let short_token = base.narrow(Resource::TokenBytes, largest_start_tag - 1);
    let expected = authored_slice(xml, short_token).unwrap_err();
    assert_limit(
        &expected,
        Resource::TokenBytes,
        largest_start_tag - 1,
        largest_start_tag,
    );
    assert_authored_stream_matches(xml, short_token);

    let exact_bytes = base.narrow(Resource::Bytes, xml.len());
    assert_eq!(authored_slice(xml, exact_bytes).unwrap(), report);
    let short_bytes = base.narrow(Resource::Bytes, xml.len() - 1);
    let expected = authored_slice(xml, short_bytes).unwrap_err();
    assert_limit(&expected, Resource::Bytes, xml.len() - 1, xml.len());
    // The slice auditor knows the whole length before parsing. The stream
    // auditor reports the same refusal at the first guarded window that sees
    // the limit, so compare the typed payload while keeping that documented
    // source-order offset distinction explicit.
    for chunks in chunk_patterns() {
        let actual = authored_stream(Chunked::new(xml, chunks), short_bytes).unwrap_err();
        assert_error_equivalent(&expected, &actual, chunks, false);
    }
}

struct Chunked<'a> {
    input: &'a [u8],
    offset: usize,
    chunks: &'a [usize],
    chunk_index: usize,
    exposed_end: Option<usize>,
}

impl<'a> Chunked<'a> {
    fn new(input: &'a [u8], chunks: &'a [usize]) -> Self {
        assert!(!chunks.is_empty());
        assert!(chunks.iter().all(|size| *size != 0));
        Self {
            input,
            offset: 0,
            chunks,
            chunk_index: 0,
            exposed_end: None,
        }
    }
}

impl Read for Chunked<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || self.offset == self.input.len() {
            return Ok(0);
        }
        let size = self.chunks[self.chunk_index % self.chunks.len()];
        let amount = output
            .len()
            .min(size)
            .min(self.input.len().saturating_sub(self.offset));
        output[..amount].copy_from_slice(&self.input[self.offset..self.offset + amount]);
        self.offset += amount;
        Ok(amount)
    }
}

impl BufRead for Chunked<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.offset == self.input.len() {
            return Ok(&[]);
        }
        let end = *self.exposed_end.get_or_insert_with(|| {
            let size = self.chunks[self.chunk_index % self.chunks.len()];
            (self.offset + size).min(self.input.len())
        });
        Ok(&self.input[self.offset..end])
    }

    fn consume(&mut self, amount: usize) {
        let end = self.exposed_end.unwrap_or(self.offset);
        assert!(amount <= end.saturating_sub(self.offset));
        self.offset += amount;
        if self.offset == end {
            self.exposed_end = None;
            self.chunk_index += 1;
        }
    }
}

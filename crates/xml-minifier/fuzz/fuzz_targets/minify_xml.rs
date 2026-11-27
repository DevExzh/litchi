#![no_main]

use std::io::{self, BufRead, Read};

use libfuzzer_sys::fuzz_target;
use xml_minifier::audit::{self, Error, Limits, Report, StreamError};

const MAX_INPUT: usize = 64 * 1024;
const MAX_TOKEN: usize = 4 * 1024;
const MAX_DEPTH: usize = 32;
const MAX_ATTRIBUTES: usize = 1_000;
const MAX_EVENTS: usize = 128 * 1024;
const MAX_TEXT: usize = 64 * 1024;

// The runtime auditor is the fuzz target's data-dependent surface. Keep one
// fixed invocation of the existing producer macro as a compile-time smoke
// check; procedural macros cannot consume runtime fuzz bytes.
const MINIFIER_MACRO_SMOKE: &str =
    xml_minifier::minified_xml_str!(r#"<root><empty></empty></root>"#);

fuzz_target!(|data: &[u8]| {
    // Keep every resource finite even when a caller raises libFuzzer's input
    // size. Oversized inputs are skipped so the parser always sees the full
    // source under test. The first 256 bytes also control arbitrary source
    // chunking without changing the XML payload being compared.
    if data.len() > MAX_INPUT {
        return;
    }
    let input = data;
    let controls = &data[..data.len().min(256)];
    let limits = fuzz_limits();

    assert_eq!(MINIFIER_MACRO_SMOKE, "<root><empty/></root>");
    compare_policy(input, controls, limits, false);
    compare_policy(input, controls, limits, true);
});

fn fuzz_limits() -> Limits {
    Limits::new(
        MAX_INPUT,
        MAX_DEPTH,
        MAX_EVENTS,
        MAX_ATTRIBUTES,
        MAX_TOKEN,
        MAX_TEXT,
    )
    .expect("fuzz profile must stay below immutable XML ceilings")
}

fn compare_policy(input: &[u8], controls: &[u8], limits: Limits, authored: bool) {
    let expected = if authored {
        audit::verify_authored(input, limits)
    } else {
        audit::verify(input, limits)
    };
    let actual = stream_result(input, controls, limits, authored);
    let input_is_utf8 = std::str::from_utf8(input).is_ok();

    match (&expected, &actual) {
        (Ok(expected), Ok(actual)) => assert_eq!(expected, actual),
        (
            Err(Error::Encoding {
                valid_up_to: expected,
            }),
            Err(Error::Encoding {
                valid_up_to: actual,
            }),
        ) => assert_eq!(expected, actual),
        (Err(Error::Encoding { .. }), Err(_actual)) if !input_is_utf8 => {
            // `verify` checks UTF-8 for the complete slice first. The reader
            // reports the first source-ordered failure, so an earlier malformed
            // token, compactness defect, or resource window may precede a later
            // invalid byte. This is the documented streaming precedence split.
        },
        (Err(expected), Err(actual)) => {
            if is_allowed_precedence_difference(expected, actual) {
                return;
            }
            assert_error_class(expected, actual);
        },
        (Ok(expected), Err(actual)) => {
            panic!("slice accepted input but reader rejected it: {expected:?} vs {actual:?}");
        },
        (Err(expected), Ok(actual)) => {
            panic!("slice rejected input but reader accepted it: {expected:?} vs {actual:?}");
        },
    }
}

fn is_allowed_precedence_difference(expected: &Error, actual: &Error) -> bool {
    // The slice parser can read a complete unterminated token before it reports
    // malformed XML. The guarded reader must stop at max_token + 1 instead,
    // so a truncated hostile token may be `Malformed` for the slice and a
    // typed `TokenBytes` limit for the reader.
    matches!(
        (expected, actual),
        (
            Error::Malformed { .. },
            Error::Limit {
                resource: xml_minifier::audit::Resource::TokenBytes,
                ..
            }
        )
    )
}

fn stream_result(
    input: &[u8],
    controls: &[u8],
    limits: Limits,
    authored: bool,
) -> Result<Report, Error> {
    let reader = Chunked::new(input, controls);
    let result = if authored {
        audit::verify_authored_reader(reader, limits)
    } else {
        audit::verify_reader(reader, limits)
    };
    match result {
        Ok(report) => Ok(report),
        Err(StreamError::Audit(error)) => Err(error),
        Err(StreamError::Input(error)) => {
            panic!("infallible chunk source returned an input error: {error}");
        },
        Err(_non_exhaustive) => panic!("unknown streaming error variant"),
    }
}

fn assert_error_class(expected: &Error, actual: &Error) {
    match (expected, actual) {
        (
            Error::Limit {
                resource: expected_resource,
                limit: expected_limit,
                ..
            },
            Error::Limit {
                resource: actual_resource,
                limit: actual_limit,
                ..
            },
        ) => {
            // Token lookahead and source-order byte accounting can change the
            // observed value/offset. The governed resource and configured limit
            // remain stable and are the useful differential assertion here.
            assert_eq!(expected_resource, actual_resource);
            assert_eq!(expected_limit, actual_limit);
        },
        (
            Error::Encoding {
                valid_up_to: expected,
            },
            Error::Encoding {
                valid_up_to: actual,
            },
        ) => assert_eq!(expected, actual),
        (expected, actual) => assert_eq!(error_category(expected), error_category(actual)),
    }
}

fn error_category(error: &Error) -> &'static str {
    match error {
        Error::Limit { .. } => "limit",
        Error::Encoding { .. } => "encoding",
        Error::Malformed { .. } => "malformed",
        Error::NotCompact(_) => "not-compact",
        Error::Doctype { .. } => "doctype",
        Error::Allocation => "allocation",
        _ => "unknown",
    }
}

/// A deterministic `BufRead` whose visible chunk size is selected by the
/// fuzz input. It never allocates and never returns an I/O error, so any
/// `StreamError::Input` result indicates a reader integration failure.
struct Chunked<'a> {
    input: &'a [u8],
    controls: &'a [u8],
    offset: usize,
    exposed_end: Option<usize>,
    chunk_index: usize,
}

impl<'a> Chunked<'a> {
    fn new(input: &'a [u8], controls: &'a [u8]) -> Self {
        Self {
            input,
            controls,
            offset: 0,
            exposed_end: None,
            chunk_index: 0,
        }
    }

    fn chunk_size(&self) -> usize {
        self.controls
            .get(self.chunk_index % self.controls.len().max(1))
            .map_or(1, |byte| 1 + usize::from(*byte) % 97)
    }
}

impl Read for Chunked<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let amount = available.len().min(output.len());
        output[..amount].copy_from_slice(&available[..amount]);
        self.consume(amount);
        Ok(amount)
    }
}

impl BufRead for Chunked<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.offset == self.input.len() {
            return Ok(&[]);
        }
        if self.exposed_end.is_none() {
            let end = self
                .offset
                .saturating_add(self.chunk_size())
                .min(self.input.len());
            self.exposed_end = Some(end);
        }
        let end = self.exposed_end.expect("chunk boundary initialized");
        Ok(&self.input[self.offset..end])
    }

    fn consume(&mut self, amount: usize) {
        let end = self.exposed_end.unwrap_or(self.offset);
        assert!(amount <= end.saturating_sub(self.offset));
        self.offset += amount;
        if self.offset == end {
            self.exposed_end = None;
            self.chunk_index = self.chunk_index.wrapping_add(1);
        }
    }
}

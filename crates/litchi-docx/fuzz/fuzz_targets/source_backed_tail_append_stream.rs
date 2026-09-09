#![no_main]

//! Bounded public-boundary fuzz coverage for replayable source-backed DOCX
//! paragraph streams.
//!
//! The input is either a DOCX archive or a small recipe followed by a DOCX
//! archive.  Keeping the recipe before the archive lets a corpus seed retain
//! a valid package while mutations explore event ordering, XML characters,
//! chunk boundaries, and replay changes.  The archive itself remains caller
//! input; this target intentionally has no ZIP-construction dependency.

use std::fmt;
use std::io::{self, Cursor, Write};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use libfuzzer_sys::fuzz_target;
use litchi_docx::ReadLimits;
use litchi_docx::source_backed::{
    Package, TailAppendLimits, TailAppendStreamCursor, TailAppendStreamEvent,
    TailAppendStreamLimits, TailAppendStreamSource,
};

const MAX_INPUT_BYTES: usize = 64 * 1024;
const ZIP_LOCAL_HEADER: &[u8] = b"PK\x03\x04";

#[derive(Clone, Debug)]
enum Event {
    Start,
    Text(String),
    End,
}

#[derive(Clone, Copy, Debug)]
struct CursorError;

impl fmt::Display for CursorError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("fuzz cursor failed")
    }
}

impl std::error::Error for CursorError {}

#[derive(Debug)]
struct EventCursor<'source> {
    events: &'source [Event],
    position: usize,
}

impl TailAppendStreamCursor for EventCursor<'_> {
    type Error = CursorError;

    fn next<'event>(
        &'event mut self,
    ) -> Result<Option<TailAppendStreamEvent<'event>>, Self::Error> {
        let event = self.events.get(self.position);
        self.position = self.position.saturating_add(1);
        Ok(event.map(|event| match event {
            Event::Start => TailAppendStreamEvent::ParagraphStart,
            Event::Text(text) => TailAppendStreamEvent::TextChunk(text.as_str()),
            Event::End => TailAppendStreamEvent::ParagraphEnd,
        }))
    }
}

#[derive(Debug)]
struct EventSource {
    events: Vec<Event>,
    alternate: Option<Vec<Event>>,
    opens: Arc<AtomicUsize>,
}

impl EventSource {
    fn stable(events: Vec<Event>) -> Self {
        Self {
            events,
            alternate: None,
            opens: Arc::new(AtomicUsize::new(0)),
        }
    }

    fn changing(events: Vec<Event>, alternate: Vec<Event>) -> Self {
        Self {
            events,
            alternate: Some(alternate),
            opens: Arc::new(AtomicUsize::new(0)),
        }
    }
}

impl TailAppendStreamSource for EventSource {
    type Error = CursorError;
    type Cursor<'source> = EventCursor<'source>;

    fn open<'source>(&'source self) -> Result<Self::Cursor<'source>, Self::Error> {
        let opening = self.opens.fetch_add(1, Ordering::AcqRel);
        let events = match (opening, self.alternate.as_deref()) {
            (opening, Some(alternate)) if opening > 0 => alternate,
            _ => self.events.as_slice(),
        };
        Ok(EventCursor {
            events,
            position: 0,
        })
    }
}

#[derive(Debug)]
struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PrefixFailSink {
    fn after(bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining: bytes,
        }
    }
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "fuzz sink stopped",
            ));
        }
        let accepted = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn read_limits() -> ReadLimits {
    ReadLimits::builder()
        // The fuzz input is capped at 64 KiB, while a successful candidate
        // archive also contains the generated stream and ZIP directory.
        .max_input_bytes(128 * 1024)
        .unwrap()
        .max_archive_members(32)
        .unwrap()
        .max_relationship_parts(32)
        .unwrap()
        .max_parts(32)
        .unwrap()
        .max_archive_entry_bytes(64 * 1024)
        .unwrap()
        .max_archive_total_bytes(128 * 1024)
        .unwrap()
        .max_archive_compressed_bytes(128 * 1024)
        .unwrap()
        .max_archive_metadata_bytes(64 * 1024)
        .unwrap()
        .max_part_bytes(64 * 1024)
        .unwrap()
        .max_total_part_bytes(128 * 1024)
        .unwrap()
        .max_content_types_bytes(16 * 1024)
        .unwrap()
        .max_relationship_xml_bytes(16 * 1024)
        .unwrap()
        .max_total_relationship_xml_bytes(64 * 1024)
        .unwrap()
        .max_xml_events(32 * 1024)
        .unwrap()
        .build()
        .unwrap()
}

fn stream_limits() -> TailAppendStreamLimits {
    let mut source = TailAppendLimits::default();
    source.max_source_xml_bytes = MAX_INPUT_BYTES as u64;
    source.max_text_bytes = 16 * 1024;
    source.max_fragment_bytes = 32 * 1024;
    source.max_candidate_xml_bytes = 64 * 1024;
    source.max_events = 32 * 1024;
    source.max_depth = 8;
    source.max_paragraphs = 512;
    source.max_settings_xml_bytes = 16 * 1024;
    source.max_workspace_bytes = 1024 * 1024;
    source.max_output_bytes = 128 * 1024;
    source.max_token_bytes = 512;

    let mut limits = TailAppendStreamLimits::new(source);
    limits.max_authored_paragraphs = 32;
    limits.max_authored_events = 256;
    // The encoder requires a complete paragraph envelope plus one borrowed
    // chunk to fit its replay window.  Keep the chunk ceiling comfortably
    // below the 1 KiB window so positive seeds reach preparation.
    limits.max_authored_chunk_bytes = 128;
    limits.max_authored_text_bytes = 8 * 1024;
    limits.max_authored_xml_bytes = 32 * 1024;
    limits.max_replay_bytes = 32 * 1024;
    limits.max_replay_window_bytes = 1024;
    limits.max_patch_bytes = 16 * 1024;
    limits
}

fn split_recipe(data: &[u8]) -> (&[u8], &[u8]) {
    data.windows(ZIP_LOCAL_HEADER.len())
        .position(|window| window == ZIP_LOCAL_HEADER)
        .map_or((&[][..], data), |offset| data.split_at(offset))
}

fn safe_text(bytes: &[u8]) -> String {
    let mut text = String::new();
    for byte in bytes.iter().copied().take(24) {
        text.push(match byte % 10 {
            0 => 'a',
            1 => ' ',
            2 => '<',
            3 => '&',
            4 => 'é',
            5 => '🙂',
            6 => '\'',
            7 => '"',
            // CR/LF/TAB are intentionally excluded by the plain-text XML
            // policy; use a valid non-ASCII character for this valid route.
            8 => '界',
            _ => 'z',
        });
    }
    if text.is_empty() {
        text.push_str("fuzz<&> café");
    }
    text
}

fn valid_events(recipe: &[u8]) -> Vec<Event> {
    let body = recipe.get(1..).unwrap_or_default();
    let paragraphs = 1 + body.first().copied().unwrap_or(0) as usize % 4;
    let mut events = Vec::new();
    let mut cursor = 1usize;
    for paragraph in 0..paragraphs {
        events.push(Event::Start);
        let chunks = 1 + body.get(cursor).copied().unwrap_or(paragraph as u8) as usize % 4;
        cursor = cursor.saturating_add(1);
        for chunk in 0..chunks {
            let begin = cursor.min(body.len());
            let width = 1 + body.get(cursor).copied().unwrap_or(chunk as u8) as usize % 8;
            let end = begin.saturating_add(width).min(body.len());
            let text = if begin < end {
                safe_text(&body[begin..end])
            } else if chunk == 0 {
                safe_text(&[])
            } else {
                String::new()
            };
            cursor = end.saturating_add(1);
            events.push(Event::Text(text));
        }
        events.push(Event::End);
    }
    events
}

fn recipe_events(recipe: &[u8]) -> (Vec<Event>, Option<Vec<Event>>, bool) {
    let mode = recipe.first().copied().unwrap_or(0) % 8;
    let valid = valid_events(recipe);
    match mode {
        0 => (valid, None, true),
        1 => (vec![Event::Text("outside".to_owned())], None, false),
        2 => (vec![Event::Start, Event::Start, Event::End], None, false),
        3 => (
            vec![Event::Start, Event::Text("truncated".to_owned())],
            None,
            false,
        ),
        4 => (vec![Event::Start, Event::End, Event::End], None, false),
        5 => (
            vec![Event::Start, Event::Text("bad\u{1}".to_owned()), Event::End],
            None,
            false,
        ),
        6 => {
            let mut alternate = valid.clone();
            if let Some(Event::Text(text)) = alternate
                .iter_mut()
                .find(|event| matches!(event, Event::Text(text) if !text.is_empty()))
            {
                text.push('!');
            }
            (valid, Some(alternate), false)
        },
        _ => (Vec::new(), None, false),
    }
}

fn exercise(archive: &[u8], recipe: &[u8]) -> bool {
    let Ok(package) = Package::from_reader_with_limits(Cursor::new(archive), read_limits()) else {
        return false;
    };
    let (events, alternate, expected_success) = recipe_events(recipe);
    let limits = stream_limits();
    let source = match alternate {
        Some(alternate) => EventSource::changing(events, alternate),
        None => EventSource::stable(events),
    };
    let prepared = package
        .tail_append_plain_paragraphs(source, limits)
        .prepare();
    if !expected_success {
        assert!(prepared.is_err(), "malformed authored stream was accepted");
        return true;
    }
    let Ok(plan) = prepared else {
        return false;
    };
    let before_count = package
        .document_snapshot()
        .expect("prepared source must have a semantic document snapshot")
        .paragraph_count();
    let authored_count = usize::try_from(plan.authored_proof().paragraph_count)
        .expect("bounded authored paragraph count fits usize");

    let mut output = Vec::new();
    let Ok(publication) = plan.write_to_stream(&mut output) else {
        return false;
    };
    assert!(
        !output.is_empty(),
        "successful publication produced no archive"
    );
    let current = Package::from_reader_with_limits(Cursor::new(&output), read_limits())
        .expect("successful publication must reopen");
    let after_count = current
        .document_snapshot()
        .expect("published candidate must have a semantic document snapshot")
        .paragraph_count();
    assert_eq!(after_count, before_count + authored_count);

    let mut restored = Vec::new();
    publication
        .write_inverse_to_stream(&current, &mut restored)
        .expect("immediate inverse must authorize the current candidate");
    assert_eq!(
        restored, archive,
        "immediate inverse must restore exact bytes"
    );

    let partial_package = Package::from_reader_with_limits(Cursor::new(archive), read_limits())
        .expect("the source package was accepted");
    let Ok(partial_plan) = partial_package
        .tail_append_plain_paragraphs(EventSource::stable(valid_events(recipe)), limits)
        .prepare()
    else {
        return false;
    };
    let mut sink = PrefixFailSink::after(48);
    let result = partial_plan.write_to_stream(&mut sink);
    assert!(
        result.is_err(),
        "short sink unexpectedly accepted full archive"
    );
    assert!(sink.bytes.len() <= 48);
    assert!(
        !sink.bytes.is_empty(),
        "short sink lost its accepted prefix"
    );
    true
}

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }
    let (recipe, archive) = split_recipe(data);
    let (_, _, expected_success) = recipe_events(recipe);
    let completed = exercise(archive, recipe);
    if std::env::var_os("LITCHI_DOCX_STREAM_REQUIRE_SUCCESS").is_some()
        && expected_success
        && !completed
    {
        panic!("fixed positive DOCX stream seed did not complete publication and inverse");
    }
});

//! Bounded encoder and deterministic replay reader for authored DOCX events.
//!
//! The public event, proof, provider, and error vocabulary lives in the
//! parent [`super`] module.  This module owns only the format encoder and its
//! deterministic source adapter.  Every pass keeps one caller cursor alive,
//! emits one bounded pending window, and releases each borrowed text event
//! before asking the cursor for the next event.

use std::fmt;
use std::io::{self, Read};
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::{CancellationToken, ExecutionContext, Reservation, Resource};
use sha2::{Digest as _, Sha256};

use super::{
    AuthoredPassProof, AuthoredReplayError, AuthoredReplayHandle, AuthoredReplayReader,
    AuthoredReplayReference, AuthoredStreamProof, ParagraphCursor, ParagraphStreamLimits,
    PlainParagraphEvent, ReplayableParagraphSource,
};
use crate::streaming::{append_character, escaped_character_len, is_plain_text_character};

const EVENT_START: u8 = 0;
const EVENT_TEXT: u8 = 1;
const EVENT_END: u8 = 2;
const XML_NAMESPACE_HEAD: &[u8] = b"<w:p xmlns:w=\"";
const XML_NAMESPACE_TAIL: &[u8] = b"\"><w:r><w:t xml:space=\"preserve\">";
const XML_PARAGRAPH_SUFFIX: &[u8] = b"</w:t></w:r></w:p>";
const MAX_NAMESPACE_PREFIX_BYTES: usize =
    XML_NAMESPACE_HEAD.len() + super::WORDPROCESSINGML_NAMESPACE.len() + XML_NAMESPACE_TAIL.len();

/// Explicit accounting for caller-owned cursor/provider state.
///
/// The encoder cannot infer allocations made by a provider.  Callers pass
/// the provider's declared live memory and object terms, which are reserved
/// for the duration of each pass when a package execution context is present.
/// An unmanaged compatibility package has no hierarchical budget to charge;
/// its finite scalar stream limits still apply.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(crate) struct CursorAccounting {
    /// Provider bytes live while a cursor is open.
    pub memory_bytes: u64,
    /// Provider objects live while a cursor is open.
    pub object_count: u64,
}

struct CursorReservations {
    memory: Option<Reservation>,
    objects: Option<Reservation>,
}

fn reserve_cursor(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    accounting: CursorAccounting,
) -> Result<CursorReservations, AuthoredReplayError> {
    check_execution(context, cancellation)?;
    let memory = match (context, accounting.memory_bytes) {
        (Some(context), amount) if amount != 0 => Some(
            context
                .reserve(Resource::Memory, amount)
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?,
        ),
        _ => None,
    };
    let objects = match (context, accounting.object_count) {
        (Some(context), amount) if amount != 0 => {
            match context.reserve(Resource::Objects, amount) {
                Ok(reservation) => Some(reservation),
                Err(error) => {
                    drop(memory);
                    return Err(AuthoredReplayError::Provider(Box::new(error)));
                },
            }
        },
        _ => None,
    };
    Ok(CursorReservations { memory, objects })
}

fn check_execution(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<(), AuthoredReplayError> {
    if let Some(context) = context {
        context
            .check()
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
    }
    if let Some(token) = cancellation {
        token
            .check()
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
    }
    Ok(())
}

fn reserve_memory(
    context: Option<&ExecutionContext>,
    amount: u64,
) -> Result<Option<Reservation>, AuthoredReplayError> {
    context
        .map(|context| {
            context
                .reserve(Resource::Memory, amount)
                .map(Some)
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
        })
        .unwrap_or(Ok(None))
}

fn reserve_objects(
    context: Option<&ExecutionContext>,
    amount: u64,
) -> Result<Option<Reservation>, AuthoredReplayError> {
    context
        .map(|context| {
            context
                .reserve(Resource::Objects, amount)
                .map(Some)
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))
        })
        .unwrap_or(Ok(None))
}

fn consume_work(
    context: Option<&ExecutionContext>,
    amount: u64,
) -> Result<(), AuthoredReplayError> {
    if let Some(context) = context {
        context
            .consume(Resource::Work, amount)
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
    }
    Ok(())
}

/// Error used by the private bounded pending sink.
#[derive(Debug)]
struct PendingSinkError;

impl fmt::Display for PendingSinkError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("authored replay pending window was too small")
    }
}

impl std::error::Error for PendingSinkError {}

/// Sink which receives complete fixed-window chunks from the encoder.
pub(crate) trait ChunkSink {
    /// Consume one chunk before returning.
    fn accept(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError>;
}

/// Sink used for the first deterministic proof pass.
#[derive(Debug, Default, Clone, Copy)]
pub(crate) struct DiscardSink;

impl ChunkSink for DiscardSink {
    fn accept(&mut self, _chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        Ok(())
    }
}

/// A callback-backed fixed-window encoder pass.
pub(crate) struct StreamingParagraphEncoder<S> {
    context: Option<ExecutionContext>,
    cancellation: Option<CancellationToken>,
    limits: ParagraphStreamLimits,
    strict_namespace: bool,
    sink: S,
    window: Vec<u8>,
    _window_reservation: Option<Reservation>,
    _cursor_memory_reservation: Option<Reservation>,
    _cursor_object_reservation: Option<Reservation>,
    _encoder_object_reservation: Option<Reservation>,
    state: EncoderState,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase {
    Ready,
    Paragraph,
}

struct EncoderState {
    strict_namespace: bool,
    limits: ParagraphStreamLimits,
    phase: Phase,
    paragraph_count: u64,
    event_count: u64,
    text_bytes: u64,
    encoded_xml_bytes: u64,
    event_hash: Sha256,
    encoded_hash: Sha256,
}

impl EncoderState {
    fn new(
        strict_namespace: bool,
        limits: ParagraphStreamLimits,
    ) -> Result<Self, AuthoredReplayError> {
        limits.validate().map_err(map_limits_error)?;
        validate_window_relation(limits)?;
        Ok(Self {
            strict_namespace,
            limits,
            phase: Phase::Ready,
            paragraph_count: 0,
            event_count: 0,
            text_bytes: 0,
            encoded_xml_bytes: 0,
            event_hash: Sha256::new(),
            encoded_hash: Sha256::new(),
        })
    }

    fn event_preflight<'event>(
        &mut self,
        event: PlainParagraphEvent<'event>,
    ) -> Result<EventBytes<'event>, AuthoredReplayError> {
        let next_events = checked_add(
            "authored events",
            self.event_count,
            1,
            self.limits.max_authored_events,
        )?;

        let result = match event {
            PlainParagraphEvent::ParagraphStart => {
                if self.phase != Phase::Ready {
                    return Err(AuthoredReplayError::Invalid("nested paragraph start event"));
                }
                let paragraph_count = checked_add(
                    "authored paragraphs",
                    self.paragraph_count,
                    1,
                    self.limits.max_authored_paragraphs,
                )?;
                let prefix_len = paragraph_prefix_len(self.strict_namespace)?;
                let next_xml = checked_add(
                    "authored XML bytes",
                    self.encoded_xml_bytes,
                    prefix_len,
                    self.limits.max_authored_xml_bytes,
                )?;
                self.phase = Phase::Paragraph;
                self.paragraph_count = paragraph_count;
                self.encoded_xml_bytes = next_xml;
                self.event_hash.update([EVENT_START]);
                EventBytes::ParagraphStart
            },
            PlainParagraphEvent::TextChunk(text) => {
                if self.phase != Phase::Paragraph {
                    return Err(AuthoredReplayError::Invalid(
                        "text event occurs outside a paragraph",
                    ));
                }
                let text_len = usize_to_u64(text.len(), "authored text bytes")?;
                if text_len > self.limits.max_authored_chunk_bytes {
                    return Err(AuthoredReplayError::Limit {
                        resource: "authored chunk bytes",
                        actual: text_len,
                        maximum: self.limits.max_authored_chunk_bytes,
                    });
                }
                let (escaped_len, character_count) = scan_text(text)?;
                let text_bytes = checked_add(
                    "authored text bytes",
                    self.text_bytes,
                    text_len,
                    self.limits.max_authored_text_bytes,
                )?;
                let encoded_xml_bytes = checked_add(
                    "authored XML bytes",
                    self.encoded_xml_bytes,
                    escaped_len,
                    self.limits.max_authored_xml_bytes,
                )?;
                let work = checked_add("authored event work", 1, character_count, u64::MAX - 1)?;
                self.text_bytes = text_bytes;
                self.encoded_xml_bytes = encoded_xml_bytes;
                self.event_hash.update([EVENT_TEXT]);
                self.event_hash.update(text_len.to_le_bytes());
                self.event_hash.update(text.as_bytes());
                EventBytes::Text { text, work }
            },
            PlainParagraphEvent::ParagraphEnd => {
                if self.phase != Phase::Paragraph {
                    return Err(AuthoredReplayError::Invalid(
                        "paragraph end event has no open paragraph",
                    ));
                }
                let suffix_len = usize_to_u64(XML_PARAGRAPH_SUFFIX.len(), "authored XML bytes")?;
                let encoded_xml_bytes = checked_add(
                    "authored XML bytes",
                    self.encoded_xml_bytes,
                    suffix_len,
                    self.limits.max_authored_xml_bytes,
                )?;
                self.phase = Phase::Ready;
                self.encoded_xml_bytes = encoded_xml_bytes;
                self.event_hash.update([EVENT_END]);
                EventBytes::ParagraphEnd
            },
        };
        self.event_count = next_events;
        Ok(result)
    }

    fn finish(self) -> Result<AuthoredStreamProof, AuthoredReplayError> {
        if self.phase != Phase::Ready {
            return Err(AuthoredReplayError::Invalid(
                "authored paragraph stream ended with an open paragraph",
            ));
        }
        if self.paragraph_count == 0 {
            return Err(AuthoredReplayError::Invalid(
                "authored paragraph stream is empty",
            ));
        }
        Ok(AuthoredStreamProof {
            strict_namespace: self.strict_namespace,
            paragraph_count: self.paragraph_count,
            event_count: self.event_count,
            text_bytes: self.text_bytes,
            encoded_xml_bytes: self.encoded_xml_bytes,
            event_sha256: self.event_hash.finalize().into(),
            encoded_sha256: self.encoded_hash.finalize().into(),
        })
    }
}

enum EventBytes<'a> {
    ParagraphStart,
    Text { text: &'a str, work: u64 },
    ParagraphEnd,
}

impl<S> StreamingParagraphEncoder<S>
where
    S: ChunkSink,
{
    /// Creates a fixed-window encoder using an optional package execution
    /// context. A supplied context is cloned only as a cheap handle to the
    /// same hierarchical budget; unmanaged compatibility packages still use
    /// the scalar stream limits and optional cancellation token.
    pub(crate) fn new(
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        limits: ParagraphStreamLimits,
        strict_namespace: bool,
        cursor_accounting: CursorAccounting,
        sink: S,
    ) -> Result<Self, AuthoredReplayError> {
        limits.validate().map_err(map_limits_error)?;
        validate_window_relation(limits)?;
        let cursor_reservations = reserve_cursor(context, cancellation, cursor_accounting)?;
        Self::new_with_cursor_reservations(
            context,
            cancellation,
            limits,
            strict_namespace,
            cursor_reservations,
            sink,
        )
    }

    fn new_with_cursor_reservations(
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        limits: ParagraphStreamLimits,
        strict_namespace: bool,
        cursor_reservations: CursorReservations,
        sink: S,
    ) -> Result<Self, AuthoredReplayError> {
        let state = EncoderState::new(strict_namespace, limits)?;
        check_execution(context, cancellation)?;
        let window = limits.max_replay_window_bytes;
        let window_reservation = reserve_memory(context, window)?;
        let encoder_object_reservation = match reserve_objects(context, 1) {
            Ok(reservation) => reservation,
            Err(error) => {
                drop(cursor_reservations);
                drop(window_reservation);
                return Err(error);
            },
        };
        let capacity = usize::try_from(window).map_err(|_| AuthoredReplayError::Limit {
            resource: "replay window bytes",
            actual: window,
            maximum: usize::MAX as u64,
        })?;
        let mut output = Vec::new();
        if output.try_reserve_exact(capacity).is_err() {
            drop(encoder_object_reservation);
            drop(cursor_reservations);
            drop(window_reservation);
            return Err(AuthoredReplayError::Store(
                "encoder window allocation failed",
            ));
        }
        if output.capacity() != capacity {
            drop(encoder_object_reservation);
            drop(cursor_reservations);
            drop(window_reservation);
            return Err(AuthoredReplayError::Store(
                "encoder window allocation exceeded its reservation",
            ));
        }
        Ok(Self {
            context: context.cloned(),
            cancellation: cancellation.cloned(),
            limits,
            strict_namespace,
            sink,
            window: output,
            _window_reservation: window_reservation,
            _cursor_memory_reservation: cursor_reservations.memory,
            _cursor_object_reservation: cursor_reservations.objects,
            _encoder_object_reservation: encoder_object_reservation,
            state,
        })
    }

    /// Accept one borrowed event.
    pub(crate) fn push_event(
        &mut self,
        event: PlainParagraphEvent<'_>,
    ) -> Result<(), AuthoredReplayError> {
        self.check()?;
        let event_bytes = self.state.event_preflight(event)?;
        match event_bytes {
            EventBytes::ParagraphStart => {
                let mut prefix = [0_u8; MAX_NAMESPACE_PREFIX_BYTES];
                let length = write_prefix(self.strict_namespace, &mut prefix)?;
                consume_work(self.context.as_ref(), 1)?;
                self.emit(&prefix[..length])?;
            },
            EventBytes::Text { text, work } => {
                consume_work(self.context.as_ref(), work)?;
                let mut scratch = [0_u8; 5];
                for character in text.chars() {
                    let needed = escaped_character_len(character);
                    let written = append_character(character, &mut scratch[..needed]);
                    if written != needed {
                        return Err(AuthoredReplayError::Invalid(
                            "text encoder scratch buffer overflow",
                        ));
                    }
                    self.emit(&scratch[..written])?;
                }
            },
            EventBytes::ParagraphEnd => {
                consume_work(self.context.as_ref(), 1)?;
                self.emit(XML_PARAGRAPH_SUFFIX)?;
            },
        }
        Ok(())
    }

    /// Encode all events from one cursor, leaving pass finalization to the
    /// caller so the cursor can be dropped before its accounting reservations.
    pub(crate) fn encode_cursor_events<C>(
        &mut self,
        cursor: &mut C,
    ) -> Result<(), AuthoredReplayError>
    where
        C: ParagraphCursor,
    {
        loop {
            self.check()?;
            let event = cursor
                .next()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
            let Some(event) = event else {
                break;
            };
            self.push_event(event)?;
        }
        Ok(())
    }

    /// Flush the current fixed output window without sealing the proof.
    pub(crate) fn flush_output(&mut self) -> Result<(), AuthoredReplayError> {
        self.flush_window()
    }

    /// Seal this pass after checking grammar and flushing the final window.
    pub(crate) fn finish(mut self) -> Result<AuthoredPassProof, AuthoredReplayError> {
        self.check()?;
        self.flush_window()?;
        Ok(AuthoredPassProof(self.state.finish()?))
    }

    fn check(&self) -> Result<(), AuthoredReplayError> {
        if let Some(context) = self.context.as_ref() {
            context
                .check()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        if let Some(token) = self.cancellation.as_ref() {
            token
                .check()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        Ok(())
    }

    fn emit(&mut self, bytes: &[u8]) -> Result<(), AuthoredReplayError> {
        let amount = usize_to_u64(bytes.len(), "authored XML bytes")?;
        // `event_preflight` already checked this addition before any bytes are
        // emitted.  Keep this assertion as a defensive fence for future
        // callers that add another event kind.
        if self.state.encoded_xml_bytes > self.limits.max_authored_xml_bytes {
            return Err(AuthoredReplayError::Limit {
                resource: "authored XML bytes",
                actual: self.state.encoded_xml_bytes,
                maximum: self.limits.max_authored_xml_bytes,
            });
        }
        let available = self.window.capacity().saturating_sub(self.window.len());
        if amount > usize_to_u64(available, "replay window bytes")? {
            self.flush_window()?;
        }
        let available = self.window.capacity().saturating_sub(self.window.len());
        if bytes.len() > available {
            return Err(AuthoredReplayError::Limit {
                resource: "replay window bytes",
                actual: usize_to_u64(bytes.len(), "replay window bytes")?,
                maximum: self.limits.max_replay_window_bytes,
            });
        }
        self.window.extend_from_slice(bytes);
        Ok(())
    }

    fn flush_window(&mut self) -> Result<(), AuthoredReplayError> {
        if self.window.is_empty() {
            return Ok(());
        }
        self.check()?;
        self.sink.accept(&self.window)?;
        self.state.encoded_hash.update(&self.window);
        self.window.clear();
        Ok(())
    }
}

/// A deterministic source sealed by an initial proof pass.
pub(crate) struct DeterministicReplayHandle<S> {
    source: Arc<S>,
    context: Option<ExecutionContext>,
    cancellation: Option<CancellationToken>,
    limits: ParagraphStreamLimits,
    strict_namespace: bool,
    cursor_accounting: CursorAccounting,
    durable_reference: Option<AuthoredReplayReference>,
    proof: AuthoredStreamProof,
}

impl<S> fmt::Debug for DeterministicReplayHandle<S>
where
    S: ReplayableParagraphSource,
{
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DeterministicReplayHandle")
            .field("proof", &self.proof)
            .finish_non_exhaustive()
    }
}

impl<S> DeterministicReplayHandle<S>
where
    S: ReplayableParagraphSource + 'static,
{
    /// Establish the first proof pass without retaining generated XML.
    pub(crate) fn seal(
        source: Arc<S>,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        limits: ParagraphStreamLimits,
        strict_namespace: bool,
        cursor_accounting: CursorAccounting,
    ) -> Result<Self, AuthoredReplayError> {
        let mut encoder = StreamingParagraphEncoder::new(
            context,
            cancellation,
            limits,
            strict_namespace,
            cursor_accounting,
            DiscardSink,
        )?;
        let durable_reference = source.durable_reference();
        if let Some(reference) = durable_reference.as_ref() {
            reference.validate_for(limits.max_patch_bytes)?;
        }
        let handle_source = Arc::clone(&source);
        if let Some(token) = cancellation {
            token
                .check()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        let mut cursor = source
            .open()
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        encoder.encode_cursor_events(&mut cursor)?;
        drop(cursor);
        let proof = encoder.finish()?.proof();
        if source.durable_reference() != durable_reference {
            return Err(AuthoredReplayError::Changed);
        }
        Ok(Self {
            source: handle_source,
            context: context.cloned(),
            cancellation: cancellation.cloned(),
            limits,
            strict_namespace,
            cursor_accounting,
            durable_reference,
            proof,
        })
    }

    /// Open the monotone one-cursor replay reader.
    pub(crate) fn open_reader(&self) -> Result<EncodingReader<'_, S>, AuthoredReplayError> {
        if let Some(context) = self.context.as_ref() {
            context
                .check()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        if let Some(token) = self.cancellation.as_ref() {
            token
                .check()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        }
        if self.source.durable_reference() != self.durable_reference {
            return Err(AuthoredReplayError::Changed);
        }
        self.limits.validate().map_err(map_limits_error)?;
        validate_window_relation(self.limits)?;
        let reader_memory_reservation = reserve_memory(
            self.context.as_ref(),
            usize_to_u64(
                size_of::<EncodingReader<'_, S>>(),
                "replay reader state bytes",
            )?,
        )?;
        let cursor_reservations = reserve_cursor(
            self.context.as_ref(),
            self.cancellation.as_ref(),
            self.cursor_accounting,
        )?;
        let cursor = self
            .source
            .open()
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?;
        EncodingReader::new(
            &self.source,
            cursor,
            self.context.clone(),
            self.cancellation.as_ref(),
            self.limits,
            self.strict_namespace,
            cursor_reservations,
            reader_memory_reservation,
            self.durable_reference.clone(),
            self.proof,
        )
    }
}

impl<S> AuthoredReplayHandle for DeterministicReplayHandle<S>
where
    S: ReplayableParagraphSource + 'static,
{
    fn proof(&self) -> AuthoredStreamProof {
        self.proof
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        Ok(Box::new(self.open_reader()?))
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.durable_reference.clone()
    }
}

/// Bounded reader that retains one source cursor and one encoded pending
/// window. Text borrows are consumed before the next cursor call, and each
/// borrowed chunk is emitted incrementally through that window without
/// buffering or splitting an owned paragraph.
pub(crate) struct EncodingReader<'source, S>
where
    S: ReplayableParagraphSource,
{
    source: &'source S,
    cursor: Option<S::Cursor<'source>>,
    encoder: Option<StreamingParagraphEncoder<PendingSink>>,
    context: Option<ExecutionContext>,
    cancellation: Option<CancellationToken>,
    durable_reference: Option<AuthoredReplayReference>,
    proof: AuthoredStreamProof,
    _pending_window_reservation: Option<Reservation>,
    _reader_object_reservation: Option<Reservation>,
    _reader_memory_reservation: Option<Reservation>,
    eof: bool,
    completed: Option<AuthoredPassProof>,
    poisoned: bool,
}

impl<'source, S> EncodingReader<'source, S>
where
    S: ReplayableParagraphSource,
{
    fn new(
        source: &'source S,
        cursor: S::Cursor<'source>,
        context: Option<ExecutionContext>,
        cancellation: Option<&CancellationToken>,
        limits: ParagraphStreamLimits,
        strict_namespace: bool,
        cursor_reservations: CursorReservations,
        reader_memory_reservation: Option<Reservation>,
        durable_reference: Option<AuthoredReplayReference>,
        proof: AuthoredStreamProof,
    ) -> Result<Self, AuthoredReplayError> {
        limits.validate().map_err(map_limits_error)?;
        validate_window_relation(limits)?;
        let window = usize::try_from(limits.max_replay_window_bytes).map_err(|_| {
            AuthoredReplayError::Limit {
                resource: "replay window bytes",
                actual: limits.max_replay_window_bytes,
                maximum: usize::MAX as u64,
            }
        })?;
        let pending_window_reservation =
            reserve_memory(context.as_ref(), limits.max_replay_window_bytes)?;
        let reader_object_reservation = match reserve_objects(context.as_ref(), 1) {
            Ok(reservation) => reservation,
            Err(error) => {
                drop(pending_window_reservation);
                return Err(error);
            },
        };
        let mut pending = Vec::new();
        if pending.try_reserve_exact(window).is_err() {
            drop(reader_object_reservation);
            drop(pending_window_reservation);
            return Err(AuthoredReplayError::Store(
                "replay window allocation failed",
            ));
        }
        if pending.capacity() != window {
            drop(reader_object_reservation);
            drop(pending_window_reservation);
            return Err(AuthoredReplayError::Store(
                "replay window allocation exceeded its reservation",
            ));
        }
        let pending_sink = PendingSink {
            bytes: pending,
            maximum: window,
            position: 0,
        };
        let encoder = match StreamingParagraphEncoder::new_with_cursor_reservations(
            context.as_ref(),
            cancellation,
            limits,
            strict_namespace,
            cursor_reservations,
            pending_sink,
        ) {
            Ok(encoder) => encoder,
            Err(error) => {
                drop(reader_object_reservation);
                drop(pending_window_reservation);
                return Err(error);
            },
        };
        Ok(Self {
            source,
            cursor: Some(cursor),
            encoder: Some(encoder),
            context,
            cancellation: cancellation.cloned(),
            durable_reference,
            proof,
            _pending_window_reservation: pending_window_reservation,
            _reader_object_reservation: reader_object_reservation,
            _reader_memory_reservation: reader_memory_reservation,
            eof: false,
            completed: None,
            poisoned: false,
        })
    }

    fn pump_event(&mut self) -> Result<(), AuthoredReplayError> {
        if self.eof {
            return Ok(());
        }
        check_execution(self.context.as_ref(), self.cancellation.as_ref())?;
        if self.source.durable_reference() != self.durable_reference {
            return Err(AuthoredReplayError::Changed);
        }
        let event = {
            let cursor = self.cursor.as_mut().ok_or(AuthoredReplayError::Invalid(
                "replay cursor was already finished",
            ))?;
            cursor
                .next()
                .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?
        };
        let Some(event) = event else {
            if self.source.durable_reference() != self.durable_reference {
                return Err(AuthoredReplayError::Changed);
            }
            drop(self.cursor.take());
            let encoder = self.encoder.take().ok_or(AuthoredReplayError::Invalid(
                "replay encoder was already finished",
            ))?;
            let actual = encoder.finish()?;
            if actual.proof() != self.proof {
                return Err(AuthoredReplayError::Changed);
            }
            self.completed = Some(actual);
            self.eof = true;
            return Ok(());
        };
        let encoder = self.encoder.as_mut().ok_or(AuthoredReplayError::Invalid(
            "replay encoder was already finished",
        ))?;
        encoder.push_event(event)?;
        encoder.flush_output()
    }

    fn pending(&mut self) -> Option<&mut PendingSink> {
        self.encoder.as_mut().map(|encoder| &mut encoder.sink)
    }

    fn pending_for_read(&mut self) -> io::Result<&mut PendingSink> {
        if self.encoder.is_none() {
            self.poisoned = true;
            return Err(io::Error::other(AuthoredReplayError::Invalid(
                "replay pending sink is unavailable before EOF",
            )));
        }
        match self.pending() {
            Some(pending) => Ok(pending),
            None => Err(io::Error::other(AuthoredReplayError::Invalid(
                "replay pending sink disappeared",
            ))),
        }
    }
}

impl<S> Read for EncodingReader<'_, S>
where
    S: ReplayableParagraphSource,
{
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        if self.poisoned {
            return Err(io::Error::other(
                "authored replay reader is poisoned after a terminal error",
            ));
        }
        loop {
            if self.eof {
                return Ok(0);
            }
            let (pending_len, position) = {
                let pending = self.pending_for_read()?;
                (pending.bytes.len(), pending.position)
            };
            if position < pending_len {
                let count = (pending_len - position).min(output.len());
                {
                    let pending = self.pending_for_read()?;
                    output[..count].copy_from_slice(&pending.bytes[position..position + count]);
                    pending.position = position + count;
                }
                return Ok(count);
            }
            {
                let pending = self.pending_for_read()?;
                pending.bytes.clear();
                pending.position = 0;
            }
            if let Err(error) = self.pump_event() {
                self.poisoned = true;
                return Err(io::Error::other(error));
            }
            // Empty text chunks are valid no-output events.  Keep pumping
            // until a byte, EOF, or a typed provider/event error appears.
        }
    }
}

impl<S> AuthoredReplayReader for EncodingReader<'_, S>
where
    S: ReplayableParagraphSource,
{
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
        if self.poisoned {
            return Err(AuthoredReplayError::Invalid(
                "replay reader failed before EOF",
            ));
        }
        if !self.eof {
            return Err(AuthoredReplayError::Invalid(
                "replay reader was finished before EOF",
            ));
        }
        let pending = self
            .encoder
            .as_ref()
            .map(|encoder| encoder.sink.bytes.len() != encoder.sink.position)
            .unwrap_or(false);
        if pending {
            return Err(AuthoredReplayError::Invalid(
                "replay reader was finished before pending bytes were consumed",
            ));
        }
        self.completed.ok_or(AuthoredReplayError::Invalid(
            "replay reader lost its terminal proof",
        ))
    }
}

struct PendingSink {
    bytes: Vec<u8>,
    maximum: usize,
    position: usize,
}

impl ChunkSink for PendingSink {
    fn accept(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        if self.position != self.bytes.len() {
            return Err(AuthoredReplayError::Store(
                "replay pending bytes were not drained before the next event",
            ));
        }
        if chunk.len() > self.maximum {
            return Err(AuthoredReplayError::Provider(Box::new(PendingSinkError)));
        }
        self.bytes.clear();
        self.position = 0;
        self.bytes.extend_from_slice(chunk);
        Ok(())
    }
}

fn validate_window_relation(limits: ParagraphStreamLimits) -> Result<(), AuthoredReplayError> {
    let envelope = checked_add(
        "authored XML bytes",
        paragraph_prefix_len(false)?,
        usize_to_u64(XML_PARAGRAPH_SUFFIX.len(), "authored XML bytes")?,
        u64::MAX - 1,
    )?;
    if envelope > limits.max_replay_window_bytes {
        return Err(AuthoredReplayError::Limit {
            resource: "replay window bytes",
            actual: envelope,
            maximum: limits.max_replay_window_bytes,
        });
    }
    Ok(())
}

fn paragraph_prefix_len(strict_namespace: bool) -> Result<u64, AuthoredReplayError> {
    let namespace = if strict_namespace {
        super::STRICT_WORDPROCESSINGML_NAMESPACE
    } else {
        super::WORDPROCESSINGML_NAMESPACE
    };
    let length = XML_NAMESPACE_HEAD
        .len()
        .checked_add(namespace.len())
        .and_then(|value| value.checked_add(XML_NAMESPACE_TAIL.len()))
        .ok_or(AuthoredReplayError::Limit {
            resource: "authored XML bytes",
            actual: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    usize_to_u64(length, "authored XML bytes")
}

fn write_prefix(
    strict_namespace: bool,
    output: &mut [u8; MAX_NAMESPACE_PREFIX_BYTES],
) -> Result<usize, AuthoredReplayError> {
    let namespace = if strict_namespace {
        super::STRICT_WORDPROCESSINGML_NAMESPACE
    } else {
        super::WORDPROCESSINGML_NAMESPACE
    };
    let length = XML_NAMESPACE_HEAD
        .len()
        .checked_add(namespace.len())
        .and_then(|value| value.checked_add(XML_NAMESPACE_TAIL.len()))
        .ok_or(AuthoredReplayError::Limit {
            resource: "authored XML bytes",
            actual: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    if length > output.len() {
        return Err(AuthoredReplayError::Store(
            "paragraph prefix scratch too small",
        ));
    }
    let mut position = 0;
    output[position..position + XML_NAMESPACE_HEAD.len()].copy_from_slice(XML_NAMESPACE_HEAD);
    position += XML_NAMESPACE_HEAD.len();
    output[position..position + namespace.len()].copy_from_slice(namespace);
    position += namespace.len();
    output[position..position + XML_NAMESPACE_TAIL.len()].copy_from_slice(XML_NAMESPACE_TAIL);
    Ok(length)
}

fn scan_text(text: &str) -> Result<(u64, u64), AuthoredReplayError> {
    let mut escaped = 0_u64;
    let mut characters = 0_u64;
    for character in text.chars() {
        if !is_plain_text_character(character) {
            return Err(AuthoredReplayError::Invalid(
                "text event contains an invalid XML 1.0 character",
            ));
        }
        escaped = checked_add(
            "authored XML bytes",
            escaped,
            usize_to_u64(escaped_character_len(character), "authored XML bytes")?,
            u64::MAX - 1,
        )?;
        characters = checked_add("authored text characters", characters, 1, u64::MAX - 1)?;
    }
    Ok((escaped, characters))
}

fn checked_add(
    resource: &'static str,
    left: u64,
    right: u64,
    maximum: u64,
) -> Result<u64, AuthoredReplayError> {
    let value = left.checked_add(right).ok_or(AuthoredReplayError::Limit {
        resource,
        actual: u64::MAX,
        maximum,
    })?;
    if value > maximum {
        return Err(AuthoredReplayError::Limit {
            resource,
            actual: value,
            maximum,
        });
    }
    Ok(value)
}

fn usize_to_u64(value: usize, resource: &'static str) -> Result<u64, AuthoredReplayError> {
    u64::try_from(value).map_err(|_| AuthoredReplayError::Limit {
        resource,
        actual: u64::MAX,
        maximum: u64::MAX - 1,
    })
}

fn map_limits_error(error: super::Error) -> AuthoredReplayError {
    match error {
        super::Error::Limit {
            resource,
            actual,
            maximum,
        } => AuthoredReplayError::Limit {
            resource,
            actual,
            maximum,
        },
        error => AuthoredReplayError::Provider(Box::new(error)),
    }
}

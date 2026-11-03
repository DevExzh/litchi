//! Storage-scoped highlight and text-comment engine (record 0374).
//!
//! A `TSWP.StorageArchive`'s highlight table (field 23, read and rewritten
//! with [`crate::ranged`] and [`Table::Highlight`]) names
//! `TSWP.HighlightArchive` objects (message type
//! [`HIGHLIGHT_MESSAGE_TYPE`]). Each highlight owns one
//! `TSD.CommentStorageArchive` (message type
//! [`COMMENT_STORAGE_MESSAGE_TYPE`]): an empty body makes it a plain
//! highlight, a nonempty one a text comment, whose direct replies are
//! further comment storages. This is the annotation graph of the
//! retired host's (commit `27b4840684`) `text::Draft` highlight and comment
//! families.
//!
//! The engine reads graphs through [`Objects`] and encodes or rewrites the
//! highlight and comment-storage payloads. The caller (a format package)
//! allocates identifiers, picks fresh text-attribute and storage UUIDs,
//! the creation date and the annotation author, inserts and removes the
//! objects and keeps archive metadata (object references, object UUIDs,
//! external references, the identifier watermark).
//!
//! [`Table::Highlight`]: crate::ranged::Table::Highlight

use litchi_iwa_common::comment::{AuthorId, Uuid};
use litchi_iwa_common::wire::{
    append_repeated_length_delimited_field, patch_length_delimited_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::comment::raw::{comment_id, reply_id};
use litchi_iwa_text::comment::{
    Metadata, TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
};
use litchi_iwa_text::date_time::Instant;
use litchi_iwa_text::highlight::TextHighlight;
use litchi_iwa_text::highlight::raw::from_object_id;
use litchi_iwa_text::position::TextRange;
use thiserror::Error as ThisError;

use crate::native::{self, Budget, NativeError, Writer};
use crate::ranged::{self, Table};
use crate::smart_field::TextAttributeUuid;

/// `TSWP.HighlightArchive`.
pub const HIGHLIGHT_MESSAGE_TYPE: u32 = 2_013;
/// `TSD.CommentStorageArchive`.
pub const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;

const HIGHLIGHT_STORAGE_FIELD: u32 = 1;
const HIGHLIGHT_UUID_FIELD: u32 = 2;
const COMMENT_TEXT_FIELD: u32 = 1;
const COMMENT_DATE_FIELD: u32 = 2;
const COMMENT_AUTHOR_FIELD: u32 = 3;
const COMMENT_REPLIES_FIELD: u32 = 4;
const COMMENT_UUID_FIELD: u32 = 5;
const DATE_SECONDS_FIELD: u32 = 1;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;

/// Why an annotation operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A payload is malformed or the graph is inconsistent.
    #[error("invalid text-annotation source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// The storage does not own the annotation or reply (or owns it more
    /// than once, or it has the other kind).
    #[error("text annotation or reply not owned by the storage")]
    NotOwned,
    /// The storage's highlight table refused the change.
    #[error("highlight table refused the change: {0}")]
    Table(ranged::Error),
    /// A configured wire resource was exhausted.
    #[error("text-annotation {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} text-annotation items or bytes")]
    Allocation {
        /// The attempted item or byte capacity.
        amount: usize,
    },
}

impl From<NativeError> for Error {
    fn from(error: NativeError) -> Self {
        match error {
            NativeError::Invalid(reason) => Self::InvalidSource { reason },
            NativeError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind,
                observed,
                maximum,
            },
            NativeError::Allocation { amount } => Self::Allocation { amount },
        }
    }
}

/// The result of an annotation operation.
pub type Result<T> = std::result::Result<T, Error>;

const fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource { reason }
}

fn common(error: litchi_iwa_common::Error) -> Error {
    native::common(error).into()
}

fn table(error: ranged::Error) -> Error {
    match error {
        ranged::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind,
            observed,
            maximum,
        },
        ranged::Error::Allocation { amount } => Error::Allocation { amount },
        other => Error::Table(other),
    }
}

/// Object payloads by identifier.
pub trait Objects {
    /// The `(message type, payload)` messages of object `identifier`, in
    /// archive order, or `None` when the package has no such object.
    fn messages(&self, identifier: u64) -> Option<Vec<(u32, &[u8])>>;
}

/// One comment storage of an annotation graph: the annotation's root or a
/// direct reply.
#[derive(Debug, Clone, PartialEq)]
pub struct Node {
    storage: u64,
    body: String,
    metadata: Metadata,
}

impl Node {
    /// The comment-storage object.
    #[must_use]
    pub const fn storage(&self) -> u64 {
        self.storage
    }

    /// The body; empty for a plain highlight.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }

    /// Creation date, author and storage UUID.
    #[must_use]
    pub const fn metadata(&self) -> Metadata {
        self.metadata
    }
}

/// One annotation a storage's highlight table names.
#[derive(Debug, Clone, PartialEq)]
pub struct Annotation {
    object: u64,
    range: TextRange,
    root: Node,
    replies: Vec<Node>,
}

impl Annotation {
    /// The highlight object.
    #[must_use]
    pub const fn object(&self) -> u64 {
        self.object
    }

    /// The covered range.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

    /// The owned comment storage.
    #[must_use]
    pub const fn root(&self) -> &Node {
        &self.root
    }

    /// The direct replies, in comment-storage order.
    #[must_use]
    pub fn replies(&self) -> &[Node] {
        &self.replies
    }

    /// Whether the annotation is a plain highlight (an empty body).
    #[must_use]
    pub fn is_highlight(&self) -> bool {
        self.root.body.is_empty()
    }

    /// The authors the graph names.
    pub fn authors(&self) -> impl Iterator<Item = u64> + '_ {
        std::iter::once(&self.root)
            .chain(&self.replies)
            .filter_map(|node| node.metadata.author().map(AuthorId::get))
    }

    /// The semantic highlight.
    ///
    /// # Errors
    ///
    /// [`Error::NotOwned`] for a text comment.
    pub fn highlight(&self) -> Result<TextHighlight> {
        if !self.is_highlight() {
            return Err(Error::NotOwned);
        }
        Ok(TextHighlight::new(
            from_object_id(self.object).map_err(|_| invalid("zero highlight identifier"))?,
            self.range,
        ))
    }

    /// The semantic text comment.
    ///
    /// # Errors
    ///
    /// [`Error::NotOwned`] for a plain highlight.
    pub fn comment(&self) -> Result<TextComment> {
        if self.is_highlight() {
            return Err(Error::NotOwned);
        }
        let replies = u32::try_from(self.replies.len())
            .map_err(|_| invalid("iWork text comment reply count exceeds u32"))?;
        Ok(TextComment::new(
            self.comment_id()?,
            self.range,
            TextCommentBody::from_boxed(self.root.body.clone().into_boxed_str())
                .map_err(|_| invalid("iWork text comment body is not representable"))?,
            self.root.metadata,
            replies,
        ))
    }

    fn comment_id(&self) -> Result<TextCommentId> {
        comment_id(self.object).map_err(|_| invalid("zero text comment identifier"))
    }

    /// The semantic replies of a text comment.
    ///
    /// # Errors
    ///
    /// [`Error::NotOwned`] for a plain highlight.
    pub fn comment_replies(&self) -> Result<Vec<TextCommentReply>> {
        if self.is_highlight() {
            return Err(Error::NotOwned);
        }
        let comment = self.comment_id()?;
        let mut replies = Vec::new();
        native::reserve(&mut replies, self.replies.len())?;
        for node in &self.replies {
            replies.push(TextCommentReply::new(
                reply_id(node.storage).map_err(|_| invalid("zero reply identifier"))?,
                comment,
                TextCommentReplyBody::from_boxed(node.body.clone().into_boxed_str())
                    .map_err(|_| invalid("iWork text comment reply body is not representable"))?,
                node.metadata,
            ));
        }
        Ok(replies)
    }

    /// The reply stored in `storage`.
    ///
    /// # Errors
    ///
    /// [`Error::NotOwned`] unless the annotation is a comment with that
    /// reply.
    pub fn reply(&self, storage: u64) -> Result<&Node> {
        if self.is_highlight() {
            return Err(Error::NotOwned);
        }
        self.replies
            .iter()
            .find(|node| node.storage == storage)
            .ok_or(Error::NotOwned)
    }
}

/// A decoded highlight payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Highlight {
    comment_storage: u64,
    uuid: TextAttributeUuid,
}

impl Highlight {
    /// The owned comment storage.
    #[must_use]
    pub const fn comment_storage(&self) -> u64 {
        self.comment_storage
    }

    /// The text-attribute UUID.
    #[must_use]
    pub const fn uuid(&self) -> &TextAttributeUuid {
        &self.uuid
    }
}

/// Decode and validate a `TSWP.HighlightArchive` payload.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the UUID or the comment storage is
/// missing or invalid; resource errors as configured.
pub fn highlight(payload: &[u8], limits: WireLimits) -> Result<Highlight> {
    let mut budget = Budget::new(limits);
    let mut storage = None::<native::Reference>;
    let mut uuid = None;
    native::fields(payload, 0, &mut budget, |field, budget| {
        match field.number {
            HIGHLIGHT_STORAGE_FIELD => storage
                .get_or_insert_with(native::Reference::default)
                .merge(field.bytes()?, 1, budget)?,
            HIGHLIGHT_UUID_FIELD => uuid = Some(field.string(budget)?),
            _ => {},
        }
        Ok(())
    })?;
    let uuid = uuid.ok_or(invalid(
        "iWork highlight is missing its text-attribute UUID",
    ))?;
    let uuid = TextAttributeUuid::new(&uuid)
        .map_err(|_| invalid("iWork highlight has an invalid text-attribute UUID"))?;
    let comment_storage = storage
        .ok_or(invalid("iWork highlight is missing its comment storage"))?
        .identifier;
    if comment_storage == 0 {
        return Err(invalid(
            "iWork highlight has a zero comment-storage identifier",
        ));
    }
    Ok(Highlight {
        comment_storage,
        uuid,
    })
}

/// Encode a new `TSWP.HighlightArchive` owning `comment_storage`. Its
/// archive metadata references `comment_storage`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a zero comment storage; resource errors as
/// configured.
pub fn encode_highlight(
    comment_storage: u64,
    uuid: &TextAttributeUuid,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    if comment_storage == 0 {
        return Err(invalid(
            "iWork highlight has a zero comment-storage identifier",
        ));
    }
    let mut writer = Writer::new(limits);
    writer.message(HIGHLIGHT_STORAGE_FIELD, |writer| {
        native::Reference {
            identifier: comment_storage,
        }
        .encode(writer)
    })?;
    writer.bytes(HIGHLIGHT_UUID_FIELD, uuid.as_str().as_bytes())?;
    Ok(writer.into_bytes())
}

/// A decoded comment storage.
#[derive(Debug, Clone, PartialEq)]
pub struct CommentStorage {
    body: String,
    metadata: Metadata,
    replies: Vec<u64>,
}

impl CommentStorage {
    /// The body.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }

    /// Creation date, author and storage UUID.
    #[must_use]
    pub const fn metadata(&self) -> Metadata {
        self.metadata
    }

    /// The direct replies.
    #[must_use]
    pub fn replies(&self) -> &[u64] {
        &self.replies
    }
}

fn decode_uuid(data: &[u8], depth: usize, budget: &mut Budget) -> Result<(u64, u64)> {
    let (mut lower, mut upper) = (None, None);
    native::fields(data, depth, budget, |field, _| {
        match field.number {
            UUID_LOWER_FIELD => lower = Some(field.uint64()?),
            UUID_UPPER_FIELD => upper = Some(field.uint64()?),
            _ => {},
        }
        Ok(())
    })?;
    match (lower, upper) {
        (Some(lower), Some(upper)) => Ok((lower, upper)),
        _ => Err(invalid("iWork comment storage UUID is incomplete")),
    }
}

/// Decode and validate a `TSD.CommentStorageArchive` payload.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the text, the storage UUID or a reference
/// is missing or invalid; resource errors as configured.
pub fn comment_storage(payload: &[u8], limits: WireLimits) -> Result<CommentStorage> {
    decode_comment_storage(payload, &mut Budget::new(limits))
}

fn decode_comment_storage(payload: &[u8], budget: &mut Budget) -> Result<CommentStorage> {
    let mut text = None;
    let mut seconds = None;
    let mut author = None::<native::Reference>;
    let mut uuid = None;
    let mut replies = Vec::new();
    let mut failure = None;
    native::fields(payload, 0, budget, |field, budget| {
        match field.number {
            COMMENT_TEXT_FIELD => text = Some(field.string(budget)?),
            COMMENT_DATE_FIELD => {
                native::fields(field.bytes()?, 1, budget, |inner, _| {
                    if inner.number == DATE_SECONDS_FIELD {
                        seconds = Some(inner.double()?);
                    }
                    Ok(())
                })?;
            },
            COMMENT_AUTHOR_FIELD => author
                .get_or_insert_with(native::Reference::default)
                .merge(field.bytes()?, 1, budget)?,
            COMMENT_REPLIES_FIELD => {
                let mut reply = native::Reference::default();
                reply.merge(field.bytes()?, 1, budget)?;
                native::reserve(&mut replies, 1)?;
                replies.push(reply.identifier);
            },
            COMMENT_UUID_FIELD => match decode_uuid(field.bytes()?, 1, budget) {
                Ok(value) => uuid = Some(value),
                Err(error) => failure = Some(error),
            },
            _ => {},
        }
        Ok(())
    })?;
    if let Some(error) = failure {
        return Err(error);
    }
    let body = text.ok_or(invalid("iWork comment storage is missing text presence"))?;
    if seconds.is_some_and(|value| !value.is_finite()) {
        return Err(invalid(
            "iWork comment storage has a non-finite creation date",
        ));
    }
    let (lower, upper) = uuid.ok_or(invalid("iWork comment storage is missing its UUID"))?;
    if lower == 0 && upper == 0 {
        return Err(invalid("iWork comment storage has a zero UUID"));
    }
    let author = author.map(|reference| reference.identifier);
    if author == Some(0) || replies.contains(&0) {
        return Err(invalid("iWork comment storage has a zero reference"));
    }
    let metadata = Metadata::new(
        seconds
            .map(Instant::from_reference_date_seconds)
            .transpose()
            .map_err(|_| invalid("iWork comment storage creation date is not representable"))?,
        author
            .map(AuthorId::from_raw)
            .transpose()
            .map_err(|_| invalid("iWork comment storage author is not representable"))?,
        Uuid::from_parts(lower, upper)
            .map_err(|_| invalid("iWork comment storage has a zero UUID"))?,
    );
    Ok(CommentStorage {
        body,
        metadata,
        replies,
    })
}

/// Encode a new `TSD.CommentStorageArchive` with `body`, a creation date,
/// an optional author and a storage UUID. Its archive metadata references
/// the author.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a non-finite date; resource errors as
/// configured.
pub fn encode_comment_storage(
    body: &str,
    creation_date: Instant,
    author: Option<AuthorId>,
    storage_uuid: Uuid,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let seconds = creation_date.reference_date_seconds();
    if !seconds.is_finite() {
        return Err(invalid("iWork comment storage creation date is not finite"));
    }
    let mut writer = Writer::new(limits);
    writer.bytes(COMMENT_TEXT_FIELD, body.as_bytes())?;
    writer.message(COMMENT_DATE_FIELD, |writer| {
        writer.double(DATE_SECONDS_FIELD, seconds)
    })?;
    if let Some(author) = author {
        writer.message(COMMENT_AUTHOR_FIELD, |writer| {
            native::Reference {
                identifier: author.get(),
            }
            .encode(writer)
        })?;
    }
    writer.message(COMMENT_UUID_FIELD, |writer| {
        writer.uint64(UUID_LOWER_FIELD, storage_uuid.lower())?;
        writer.uint64(UUID_UPPER_FIELD, storage_uuid.upper())
    })?;
    Ok(writer.into_bytes())
}

fn admit_output(payload: &[u8], limits: WireLimits) -> Result<()> {
    Ok(native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?)
}

/// A comment storage with its body replaced; every other byte is kept.
///
/// # Errors
///
/// As [`comment_storage`].
pub fn comment_storage_with_body(
    payload: &[u8],
    body: &str,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    decode_comment_storage(payload, &mut budget)?;
    budget.admit_bytes(payload.len())?;
    let data =
        patch_length_delimited_field(payload, COMMENT_TEXT_FIELD, true, Some(body.as_bytes()))
            .map_err(common)?;
    if decode_comment_storage(&data, &mut budget)?.body != body {
        return Err(invalid(
            "iWork comment storage body patch failed validation",
        ));
    }
    admit_output(&data, limits)?;
    Ok(data)
}

/// One change of a comment storage's direct replies.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReplyChange {
    /// Append a reply.
    Add(u64),
    /// Drop a reply.
    Remove(u64),
}

/// A comment storage with one reply added or removed; every other byte
/// is kept. The caller adds or drops the reply in the message's object
/// references.
///
/// # Errors
///
/// [`Error::NotOwned`] when an added reply is already present or a
/// removed one is not present exactly once; as [`comment_storage`].
pub fn comment_storage_with_reply(
    payload: &[u8],
    change: ReplyChange,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let before = decode_comment_storage(payload, &mut budget)?;
    budget.admit_bytes(payload.len())?;
    let (data, expected) = match change {
        ReplyChange::Add(reply) => {
            if reply == 0 || before.replies.contains(&reply) {
                return Err(Error::NotOwned);
            }
            let mut writer = Writer::new(limits);
            native::Reference { identifier: reply }.encode(&mut writer)?;
            let data = append_repeated_length_delimited_field(
                payload,
                COMMENT_REPLIES_FIELD,
                &writer.into_bytes(),
            )
            .map_err(common)?;
            let mut expected = before.replies.clone();
            native::reserve(&mut expected, 1)?;
            expected.push(reply);
            (data, expected)
        },
        ReplyChange::Remove(reply) => {
            if before.replies.iter().filter(|&&id| id == reply).count() != 1 {
                return Err(Error::NotOwned);
            }
            let payloads = repeated_length_delimited_payloads(payload, COMMENT_REPLIES_FIELD)
                .map_err(common)?;
            let kept = payloads
                .iter()
                .zip(&before.replies)
                .filter(|(_, id)| **id != reply)
                .map(|(raw, _)| raw.to_vec())
                .collect::<Vec<_>>();
            let data =
                rewrite_repeated_length_delimited_fields(payload, COMMENT_REPLIES_FIELD, &kept)
                    .map_err(common)?;
            let expected = before
                .replies
                .iter()
                .copied()
                .filter(|&id| id != reply)
                .collect();
            (data, expected)
        },
    };
    if decode_comment_storage(&data, &mut budget)?.replies != expected {
        return Err(invalid(
            "iWork comment storage reply patch failed validation",
        ));
    }
    admit_output(&data, limits)?;
    Ok(data)
}

fn single_message<'a>(
    objects: &'a (impl Objects + ?Sized),
    identifier: u64,
    message_type: u32,
    missing: &'static str,
) -> Result<&'a [u8]> {
    let messages = objects.messages(identifier).ok_or(invalid(missing))?;
    match messages.as_slice() {
        [(found, payload)] if *found == message_type => Ok(payload),
        _ => Err(invalid(
            "iWork text-annotation object must contain exactly one payload",
        )),
    }
}

fn node(
    objects: &(impl Objects + ?Sized),
    storage: u64,
    budget: &mut Budget,
) -> Result<(Node, Vec<u64>)> {
    let payload = single_message(
        objects,
        storage,
        COMMENT_STORAGE_MESSAGE_TYPE,
        "iWork text annotation references a missing comment storage",
    )?;
    let decoded = decode_comment_storage(payload, budget)?;
    Ok((
        Node {
            storage,
            body: decoded.body,
            metadata: decoded.metadata,
        },
        decoded.replies,
    ))
}

/// The annotation graph of highlight object `object`, or `None` when the
/// object is not a highlight.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the graph is inconsistent: a highlight
/// with other payloads, a plain highlight with replies, or a cyclic,
/// duplicate, empty or nested reply; resource errors as configured.
pub fn graph(
    objects: &(impl Objects + ?Sized),
    object: u64,
    range: TextRange,
    limits: WireLimits,
) -> Result<Option<Annotation>> {
    resolve(objects, object, range, &mut Budget::new(limits))
}

fn resolve(
    objects: &(impl Objects + ?Sized),
    object: u64,
    range: TextRange,
    budget: &mut Budget,
) -> Result<Option<Annotation>> {
    let messages = objects.messages(object).ok_or(invalid(
        "iWork text storage references a missing annotation object",
    ))?;
    let highlights = messages
        .iter()
        .filter(|(message_type, _)| *message_type == HIGHLIGHT_MESSAGE_TYPE)
        .count();
    if highlights == 0 {
        return Ok(None);
    }
    let [(_, payload)] = messages.as_slice() else {
        return Err(invalid(
            "iWork highlight object must contain exactly one payload",
        ));
    };
    let comment_storage = highlight(payload, budget.limits())?.comment_storage;
    let (root, reply_ids) = node(objects, comment_storage, budget)?;
    if root.body.is_empty() && !reply_ids.is_empty() {
        return Err(invalid("plain highlight unexpectedly owns comment replies"));
    }
    let mut replies: Vec<Node> = Vec::new();
    native::reserve(&mut replies, reply_ids.len())?;
    for reply in reply_ids {
        if reply == comment_storage || replies.iter().any(|node| node.storage == reply) {
            return Err(invalid(
                "iWork text annotation has a cyclic or duplicate reply",
            ));
        }
        let (node, nested) = node(objects, reply, budget)?;
        if node.body.is_empty() || !nested.is_empty() {
            return Err(invalid(
                "iWork text annotation reply must be nonempty and cannot own nested replies",
            ));
        }
        replies.push(node);
    }
    Ok(Some(Annotation {
        object,
        range,
        root,
        replies,
    }))
}

/// Every annotation the storage's highlight table names, ordered by text
/// position (objects that are not highlights are skipped).
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed table, a missing object, a
/// highlight named twice or an inconsistent graph; resource errors as
/// configured.
pub fn annotations(
    storage: &[u8],
    objects: &(impl Objects + ?Sized),
    limits: WireLimits,
) -> Result<Vec<Annotation>> {
    let mut budget = Budget::new(limits);
    let entries = ranged::entries(storage, Table::Highlight, limits).map_err(table)?;
    let mut found: Vec<Annotation> = Vec::new();
    for entry in entries {
        let Some(annotation) = resolve(objects, entry.object(), entry.range(), &mut budget)? else {
            continue;
        };
        if found.iter().any(|other| other.object == annotation.object) {
            return Err(invalid(
                "iWork text storage references an annotation object more than once",
            ));
        }
        native::reserve(&mut found, 1)?;
        found.push(annotation);
    }
    Ok(found)
}

/// The annotation `object` of the storage, of the requested kind.
///
/// # Errors
///
/// [`Error::NotOwned`] unless the storage names it once with that kind;
/// as [`annotations`].
pub fn annotation(
    storage: &[u8],
    objects: &(impl Objects + ?Sized),
    object: u64,
    highlight: bool,
    limits: WireLimits,
) -> Result<Annotation> {
    annotations(storage, objects, limits)?
        .into_iter()
        .find(|annotation| annotation.object == object && annotation.is_highlight() == highlight)
        .ok_or(Error::NotOwned)
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use litchi_iwa_protos::{tsd, tsp, tswp};
    use prost::Message;

    use super::*;

    const LIMITS: WireLimits = WireLimits::DEFAULT;
    const UUID: &str = "4c495443-4849-5731-442d-686c69676874";

    struct Map(BTreeMap<u64, Vec<(u32, Vec<u8>)>>);

    impl Objects for Map {
        fn messages(&self, identifier: u64) -> Option<Vec<(u32, &[u8])>> {
            self.0
                .get(&identifier)
                .map(|messages| messages.iter().map(|(t, p)| (*t, p.as_slice())).collect())
        }
    }

    fn storage_uuid() -> Uuid {
        Uuid::from_parts(7, 9).unwrap()
    }

    #[test]
    fn highlight_round_trips_through_prost() {
        let uuid = TextAttributeUuid::new(UUID).unwrap();
        let payload = encode_highlight(5, &uuid, LIMITS).unwrap();
        let expected = tswp::HighlightArchive {
            comment_storage: Some(tsp::Reference {
                identifier: 5,
                ..Default::default()
            }),
            text_attribute_uuid_string: Some(UUID.to_owned()),
        };
        assert_eq!(payload, expected.encode_to_vec());
        assert_eq!(highlight(&payload, LIMITS).unwrap().comment_storage(), 5);
        assert!(encode_highlight(0, &uuid, LIMITS).is_err());
    }

    #[test]
    fn comment_storage_matches_prost_and_rewrites() {
        let date = Instant::from_reference_date_seconds(100.5).unwrap();
        let author = AuthorId::from_raw(3).unwrap();
        let payload =
            encode_comment_storage("Hi", date, Some(author), storage_uuid(), LIMITS).unwrap();
        let expected = tsd::CommentStorageArchive {
            text: Some("Hi".to_owned()),
            creation_date: Some(tsp::Date { seconds: 100.5 }),
            author: Some(tsp::Reference {
                identifier: 3,
                ..Default::default()
            }),
            storage_uuid: Some(tsp::Uuid { lower: 7, upper: 9 }),
            ..Default::default()
        };
        assert_eq!(payload, expected.encode_to_vec());
        let renamed = comment_storage_with_body(&payload, "Bye", LIMITS).unwrap();
        assert_eq!(comment_storage(&renamed, LIMITS).unwrap().body(), "Bye");
        let replied = comment_storage_with_reply(&renamed, ReplyChange::Add(11), LIMITS).unwrap();
        assert_eq!(comment_storage(&replied, LIMITS).unwrap().replies(), [11]);
        assert_eq!(
            comment_storage_with_reply(&replied, ReplyChange::Add(11), LIMITS),
            Err(Error::NotOwned)
        );
        let removed =
            comment_storage_with_reply(&replied, ReplyChange::Remove(11), LIMITS).unwrap();
        assert_eq!(removed, renamed);
    }

    #[test]
    fn graph_validates_replies() {
        let date = Instant::from_reference_date_seconds(1.0).unwrap();
        let uuid = TextAttributeUuid::new(UUID).unwrap();
        let root = encode_comment_storage("Note", date, None, storage_uuid(), LIMITS).unwrap();
        let root = comment_storage_with_reply(&root, ReplyChange::Add(12), LIMITS).unwrap();
        let reply =
            encode_comment_storage("Reply", date, None, Uuid::from_parts(8, 9).unwrap(), LIMITS)
                .unwrap();
        let mut objects = Map(BTreeMap::from([
            (
                10,
                vec![(
                    HIGHLIGHT_MESSAGE_TYPE,
                    encode_highlight(11, &uuid, LIMITS).unwrap(),
                )],
            ),
            (11, vec![(COMMENT_STORAGE_MESSAGE_TYPE, root)]),
            (12, vec![(COMMENT_STORAGE_MESSAGE_TYPE, reply)]),
        ]));
        let range = TextRange::new(
            litchi_iwa_text::position::TextPosition::ZERO,
            litchi_iwa_text::position::TextPosition::from_utf16_code_units(2),
        )
        .unwrap();
        let found = graph(&objects, 10, range, LIMITS).unwrap().unwrap();
        assert!(!found.is_highlight());
        assert_eq!(found.comment_replies().unwrap().len(), 1);
        assert!(found.highlight().is_err());
        let empty = encode_comment_storage("", date, None, storage_uuid(), LIMITS).unwrap();
        let empty = comment_storage_with_reply(&empty, ReplyChange::Add(12), LIMITS).unwrap();
        objects
            .0
            .insert(11, vec![(COMMENT_STORAGE_MESSAGE_TYPE, empty)]);
        assert!(graph(&objects, 10, range, LIMITS).is_err());
        objects.0.insert(13, vec![(2_001, Vec::new())]);
        assert_eq!(graph(&objects, 13, range, LIMITS), Ok(None));
    }
}

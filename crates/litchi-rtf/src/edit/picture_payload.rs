//! Exact-source updates for conservative standalone hexadecimal picture payloads.

use super::{Commit, Edit, Error, Operation, Snapshot};
use crate::lexer::{ControlWord, Lexer, Token};
use crate::{ImageType, RtfError};
use bumpalo::Bump;
use serde_json::{Map, Value};
use std::borrow::Cow;
use std::ops::Range;

/// Maximum number of standalone picture payloads in one atomic batch.
pub const MAX_PICTURE_PAYLOAD_OPERATIONS: usize = 64;

/// Maximum decoded size of one payload accepted by this exact-source seam.
const MAX_EDITABLE_PICTURE_BYTES: usize = 64 * 1024;

/// One source-relative replacement in an atomic picture payload batch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PicturePayloadReplacement {
    position: usize,
    payload: Vec<u8>,
}

impl PicturePayloadReplacement {
    /// Creates a replacement for one zero-based standalone body picture.
    #[must_use]
    pub fn new(position: usize, payload: impl Into<Vec<u8>>) -> Self {
        Self {
            position,
            payload: payload.into(),
        }
    }

    /// Zero-based standalone body-picture position.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Replacement media bytes.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }
}

#[derive(Debug, Clone)]
pub(super) struct StagedPicturePayload {
    pub(super) position: usize,
    pub(super) image_type: ImageType,
    pub(super) before: Vec<u8>,
    pub(super) after: Vec<u8>,
    pub(super) before_transport: Vec<u8>,
    pub(super) after_transport: Vec<u8>,
}

impl StagedPicturePayload {
    pub(super) fn inverse(&self) -> Self {
        Self {
            position: self.position,
            image_type: self.image_type,
            before: self.after.clone(),
            after: self.before.clone(),
            before_transport: self.after_transport.clone(),
            after_transport: self.before_transport.clone(),
        }
    }
}

#[derive(Debug)]
struct LocatedPicture {
    payload_span: Range<usize>,
    payload_transport: Vec<u8>,
    image_type: ImageType,
    data: Vec<u8>,
}

impl Edit {
    /// Stages a same-length payload update for an existing standalone PNG or JPEG picture.
    ///
    /// The exact hexadecimal digit positions, whitespace, group controls, dimensions,
    /// and every unrelated source byte are retained. Selection resolves against the
    /// immutable source. Binary `binN` pictures, nested pictures, fields, shapes,
    /// objects, protected documents, unknown syntax, and other producer ambiguity
    /// are refused rather than normalized.
    ///
    /// # Errors
    /// Returns a typed selector, size, conflict, limit, or unsupported-source error.
    pub fn replace_picture_payload(
        &mut self,
        position: usize,
        payload: impl AsRef<[u8]>,
    ) -> Result<&mut Self, Error> {
        let replacement = PicturePayloadReplacement::new(position, payload.as_ref().to_vec());
        self.stage_picture_payloads(&[replacement], None)
    }

    /// Atomically stages a source-ordered batch of standalone picture payload updates.
    ///
    /// The batch is preflighted with one bounded lexical pass before any operation is
    /// appended. Positions must be strictly increasing, and no more than
    /// [`MAX_PICTURE_PAYLOAD_OPERATIONS`] picture operations may be staged.
    ///
    /// # Errors
    /// Returns a typed batch, selector, size, conflict, limit, or unsupported-source error.
    pub fn replace_picture_payloads(
        &mut self,
        replacements: &[PicturePayloadReplacement],
    ) -> Result<&mut Self, Error> {
        self.stage_picture_payloads(replacements, None)
    }

    fn stage_picture_payloads(
        &mut self,
        replacements: &[PicturePayloadReplacement],
        exact_transports: Option<&[Vec<u8>]>,
    ) -> Result<&mut Self, Error> {
        if replacements.is_empty() {
            return Err(Error::EmptyPicturePayloadBatch);
        }
        if replacements.len() > MAX_PICTURE_PAYLOAD_OPERATIONS {
            return Err(Error::OperationLimit {
                observed: replacements.len(),
                limit: MAX_PICTURE_PAYLOAD_OPERATIONS,
            });
        }
        for (previous, incoming) in replacements.iter().zip(replacements.iter().skip(1)) {
            if incoming.position <= previous.position {
                return Err(Error::PicturePayloadBatchOutOfOrder {
                    previous: previous.position,
                    incoming: incoming.position,
                });
            }
        }
        if exact_transports.is_some_and(|transports| transports.len() != replacements.len()) {
            return Err(Error::DurablePatch(
                "picture payload transport count differs from replacements".to_string(),
            ));
        }
        if self
            .operations
            .iter()
            .any(|operation| !operation.is_picture_payload())
        {
            return Err(Error::BodyDestinationConflict);
        }
        self.ensure_operation_room_for(replacements.len())?;
        let observed_picture_operations = self.operations.len().saturating_add(replacements.len());
        if observed_picture_operations > MAX_PICTURE_PAYLOAD_OPERATIONS {
            return Err(Error::OperationLimit {
                observed: observed_picture_operations,
                limit: MAX_PICTURE_PAYLOAD_OPERATIONS,
            });
        }

        let located = locate_standalone_pictures(&self.source)?;
        let mut staged = Vec::new();
        staged
            .try_reserve(replacements.len())
            .map_err(|_error| allocation_error("staged picture payloads", replacements.len()))?;
        let mut replacement_bytes = 0usize;
        for (replacement_index, replacement) in replacements.iter().enumerate() {
            let picture = located
                .get(replacement.position)
                .ok_or(Error::PictureOutOfRange {
                    position: replacement.position,
                    count: located.len(),
                })?;
            if replacement.payload.len() != picture.data.len() {
                return Err(Error::PicturePayloadSizeMismatch {
                    position: replacement.position,
                    expected: picture.data.len(),
                    observed: replacement.payload.len(),
                });
            }
            validate_supported_image(picture.image_type, &replacement.payload)?;
            let incoming = self.operations.len().saturating_add(staged.len());
            if let Some(existing) = self.operations.iter().position(|operation| {
                matches!(
                    operation,
                    Operation::PicturePayload(existing)
                        if existing.position == replacement.position
                )
            }) {
                return Err(Error::Conflict { existing, incoming });
            }
            let after_transport = if let Some(transports) = exact_transports {
                let transport = transports.get(replacement_index).ok_or_else(|| {
                    Error::DurablePatch("missing exact picture payload transport".to_string())
                })?;
                validate_exact_transport(
                    &picture.payload_transport,
                    transport,
                    &replacement.payload,
                )?;
                transport.clone()
            } else {
                render_payload(&picture.payload_transport, &replacement.payload)?
            };
            replacement_bytes = replacement_bytes.saturating_add(after_transport.len());
            staged.push(Operation::PicturePayload(StagedPicturePayload {
                position: replacement.position,
                image_type: picture.image_type,
                before: picture.data.clone(),
                after: replacement.payload.clone(),
                before_transport: picture.payload_transport.clone(),
                after_transport,
            }));
        }
        let replacement_limit = self.source.limits().max_source_bytes();
        let observed_replacement_bytes = self.replacement_bytes.saturating_add(replacement_bytes);
        if observed_replacement_bytes > replacement_limit {
            return Err(Error::InputTooLarge {
                observed: observed_replacement_bytes,
                limit: replacement_limit,
            });
        }
        self.replacement_bytes = observed_replacement_bytes;
        self.operations.append(&mut staged);
        Ok(self)
    }
}

pub(super) fn commit(edit: Edit, operation_count: usize) -> Result<Commit, Error> {
    if operation_count > MAX_PICTURE_PAYLOAD_OPERATIONS {
        return Err(Error::OperationLimit {
            observed: operation_count,
            limit: MAX_PICTURE_PAYLOAD_OPERATIONS,
        });
    }
    let changes = super::semantic_changes(&edit.operations, &[]);
    if changes.is_empty() {
        return Ok(Commit::new(
            edit.source.clone(),
            edit.source,
            false,
            operation_count,
            changes,
        ));
    }
    super::ensure_changed_publication_allowed(&edit.source)?;
    let located = locate_standalone_pictures(&edit.source)?;
    let source_bytes = edit
        .source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    let mut candidate_bytes = Vec::new();
    candidate_bytes
        .try_reserve(source_bytes.len())
        .map_err(|_error| allocation_error("picture candidate bytes", source_bytes.len()))?;
    candidate_bytes.extend_from_slice(source_bytes);

    for operation in &edit.operations {
        let Operation::PicturePayload(operation) = operation else {
            return Err(Error::BodyDestinationConflict);
        };
        let picture = located
            .get(operation.position)
            .ok_or(Error::PictureOutOfRange {
                position: operation.position,
                count: located.len(),
            })?;
        if picture.image_type != operation.image_type
            || picture.data != operation.before
            || picture.payload_transport != operation.before_transport
        {
            return Err(Error::StalePrecondition("picture payload differs"));
        }
        let target = candidate_bytes
            .get_mut(picture.payload_span.clone())
            .ok_or(Error::UnsupportedSource(
                "picture payload provenance is outside the source",
            ))?;
        if target.len() != operation.after_transport.len() {
            return Err(Error::UnsupportedSource(
                "picture payload transport changed length",
            ));
        }
        target.copy_from_slice(&operation.after_transport);
    }

    let snapshot = Snapshot::from_bytes_with_limits(&candidate_bytes, edit.source.limits())?;
    verify_candidate(&edit.source, &snapshot, &edit.operations)?;
    Ok(Commit::new(
        edit.source,
        snapshot,
        true,
        operation_count,
        changes,
    ))
}

pub(super) fn durable_operation(
    limits: litchi_core::patch::PatchLimits,
    operation: &StagedPicturePayload,
    source: &[u8],
) -> Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = std::collections::BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(source).as_hex()),
    );
    preconditions.insert(
        "payload_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&operation.before).as_hex()),
    );
    let mut value = Map::new();
    value.insert(
        "data".to_string(),
        Value::String(super::hex_encode(&operation.after)),
    );
    value.insert(
        "transport".to_string(),
        Value::String(super::hex_encode(&operation.after_transport)),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "picture-payload.replace",
        format!("body:picture:{}", operation.position),
        preconditions,
        Value::Object(value),
    )
}

pub(super) fn apply_durable_operation(
    source: &Snapshot,
    edit: &mut Edit,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<(), Error> {
    let located = locate_standalone_pictures(source)?;
    let (replacement, transport) = decode_durable_operation(&located, operation)?;
    let replacements = [replacement];
    let transports = [transport];
    edit.stage_picture_payloads(&replacements, Some(&transports))?;
    Ok(())
}

pub(super) fn apply_durable_patch(
    source: &Snapshot,
    operations: &[litchi_core::patch::PatchOperation],
    source_hash: &str,
) -> Result<Snapshot, Error> {
    if operations.len() > MAX_PICTURE_PAYLOAD_OPERATIONS {
        return Err(Error::OperationLimit {
            observed: operations.len(),
            limit: MAX_PICTURE_PAYLOAD_OPERATIONS,
        });
    }
    let located = locate_standalone_pictures(source)?;
    let mut decoded = Vec::new();
    decoded
        .try_reserve(operations.len())
        .map_err(|_error| allocation_error("durable picture operations", operations.len()))?;
    for operation in operations {
        if operation.preconditions.len() != 2 {
            return Err(Error::DurablePatch(
                "picture operation has an invalid precondition count".to_string(),
            ));
        }
        let expected_hash = operation
            .preconditions
            .get("artifact_sha256")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::DurablePatch("missing artifact digest".to_string()))?;
        if expected_hash != source_hash {
            return Err(Error::PatchConflict);
        }
        decoded.push(decode_durable_operation(&located, operation)?);
    }
    decoded.sort_unstable_by_key(|(replacement, _transport)| replacement.position);
    for pair in decoded.windows(2) {
        let previous = pair.first().ok_or_else(|| {
            Error::DurablePatch("picture operation ordering became inconsistent".to_string())
        })?;
        let incoming = pair.get(1).ok_or_else(|| {
            Error::DurablePatch("picture operation ordering became inconsistent".to_string())
        })?;
        if previous.0.position == incoming.0.position {
            return Err(Error::DurablePatch(
                "durable picture payload targets are duplicated".to_string(),
            ));
        }
    }
    let mut replacements = Vec::new();
    let mut transports = Vec::new();
    replacements
        .try_reserve(decoded.len())
        .map_err(|_error| allocation_error("durable picture replacements", decoded.len()))?;
    transports
        .try_reserve(decoded.len())
        .map_err(|_error| allocation_error("durable picture transports", decoded.len()))?;
    for (replacement, transport) in decoded {
        replacements.push(replacement);
        transports.push(transport);
    }
    let mut edit = source.edit();
    edit.stage_picture_payloads(&replacements, Some(&transports))?;
    Ok(edit.commit()?.into_snapshot())
}

fn decode_durable_operation(
    located: &[LocatedPicture],
    operation: &litchi_core::patch::PatchOperation,
) -> Result<(PicturePayloadReplacement, Vec<u8>), Error> {
    let position = parse_picture_target(&operation.target)?;
    let picture = located.get(position).ok_or(Error::PictureOutOfRange {
        position,
        count: located.len(),
    })?;
    let expected_digest = operation
        .preconditions
        .get("payload_sha256")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::DurablePatch("missing picture payload digest".to_string()))?;
    if litchi_core::patch::BlobId::of(&picture.data).as_hex() != expected_digest {
        return Err(Error::StalePrecondition("picture payload differs"));
    }
    let value = operation.value.as_object().ok_or_else(|| {
        Error::DurablePatch("picture payload value must be an object".to_string())
    })?;
    if value.len() != 2 {
        return Err(Error::DurablePatch(
            "picture payload value has unknown fields".to_string(),
        ));
    }
    let payload = value
        .get("data")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::DurablePatch("missing picture payload data".to_string()))?;
    let transport = value
        .get("transport")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::DurablePatch("missing picture payload transport".to_string()))?;
    let payload = decode_hex(payload, MAX_EDITABLE_PICTURE_BYTES)?;
    let transport = decode_hex(transport, picture.payload_transport.len())?;
    Ok((PicturePayloadReplacement::new(position, payload), transport))
}

fn locate_standalone_pictures(source: &Snapshot) -> Result<Vec<LocatedPicture>, Error> {
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware picture rewrite",
        ));
    }
    if !source_bytes.is_ascii() {
        return Err(Error::UnsupportedSource(
            "picture payload editing requires an ASCII RTF transport",
        ));
    }
    if !source.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "picture payload editing refuses unknown RTF syntax",
        ));
    }
    if !source.shapes().is_empty() || !source.objects().is_empty() || !source.fields().is_empty() {
        return Err(Error::UnsupportedSource(
            "picture payload editing refuses shapes, objects, and fields",
        ));
    }
    if !source.model().picture_compatibility_records().is_empty() {
        return Err(Error::UnsupportedSource(
            "picture payload editing refuses compatibility picture wrappers",
        ));
    }

    let input = std::str::from_utf8(source_bytes).map_err(|_error| {
        Error::UnsupportedSource("picture payload source is not an ASCII string")
    })?;
    let arena = Bump::new();
    let mut lexer = Lexer::new_with_limits(input, &arena, source.limits());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let mut located = Vec::new();
    located
        .try_reserve(source.pictures().len())
        .map_err(|_error| allocation_error("located picture payloads", source.pictures().len()))?;

    let mut depth = 0usize;
    let mut cursor = 0usize;
    while let Some(token) = tokens.get(cursor) {
        match token {
            Token::OpenBrace => {
                depth = depth.saturating_add(1);
                if depth == 2
                    && matches!(
                        tokens.get(cursor.saturating_add(1)),
                        Some(Token::Control(ControlWord::Picture))
                    )
                {
                    let (picture, close_index) = locate_picture_group(
                        source_bytes,
                        &tokens,
                        &spans,
                        cursor,
                        source.pictures().get(located.len()),
                    )?;
                    located.push(picture);
                    cursor = close_index;
                    depth = depth.saturating_sub(1);
                }
            },
            Token::CloseBrace => {
                depth = depth.checked_sub(1).ok_or(Error::UnsupportedSource(
                    "picture source has an unmatched closing group",
                ))?;
            },
            Token::Control(_) | Token::Text(_) | Token::Binary(_) => {},
        }
        cursor = cursor.saturating_add(1);
    }
    if depth != 0 {
        return Err(Error::UnsupportedSource(
            "picture source has an unterminated group",
        ));
    }
    if located.len() != source.pictures().len() {
        return Err(Error::UnsupportedSource(
            "all pictures must be standalone direct body groups",
        ));
    }
    Ok(located)
}

fn locate_picture_group(
    source: &[u8],
    tokens: &[Token<'_>],
    spans: &[Range<usize>],
    open_index: usize,
    semantic: Option<&crate::picture::Picture<'_>>,
) -> Result<(LocatedPicture, usize), Error> {
    let semantic = semantic.ok_or(Error::UnsupportedSource(
        "standalone picture has no matching semantic record",
    ))?;
    validate_supported_image(semantic.image_type, semantic.data())?;
    if semantic.data().len() > MAX_EDITABLE_PICTURE_BYTES {
        return Err(Error::InputTooLarge {
            observed: semantic.data().len(),
            limit: MAX_EDITABLE_PICTURE_BYTES,
        });
    }
    let mut cursor = open_index.saturating_add(2);
    let mut payload_start = None;
    let mut payload_end = None;
    let mut seen_controls = PictureControlSet::default();
    loop {
        let token = tokens.get(cursor).ok_or(Error::UnsupportedSource(
            "standalone picture group is unterminated",
        ))?;
        match token {
            Token::Control(control) if payload_start.is_none() => {
                seen_controls.accept(control)?;
            },
            Token::Text(text) => {
                let span = spans.get(cursor).ok_or(Error::UnsupportedSource(
                    "picture token has no source provenance",
                ))?;
                if !text
                    .as_bytes()
                    .iter()
                    .all(|byte| byte.is_ascii_whitespace() || byte.is_ascii_hexdigit())
                {
                    return Err(Error::UnsupportedSource(
                        "picture payload contains non-hexadecimal text",
                    ));
                }
                payload_start.get_or_insert(span.start);
                payload_end = Some(span.end);
            },
            Token::CloseBrace => break,
            Token::Binary(_) => {
                return Err(Error::UnsupportedSource(
                    "binary picture payloads are outside the exact hexadecimal closure",
                ));
            },
            Token::OpenBrace => {
                return Err(Error::UnsupportedSource(
                    "nested picture destinations are outside the exact payload closure",
                ));
            },
            Token::Control(_) => {
                return Err(Error::UnsupportedSource(
                    "picture controls after payload bytes are ambiguous",
                ));
            },
        }
        cursor = cursor.saturating_add(1);
    }
    let start = payload_start.ok_or(Error::UnsupportedSource(
        "standalone picture has no hexadecimal payload",
    ))?;
    let end = payload_end.ok_or(Error::UnsupportedSource(
        "standalone picture has no hexadecimal payload",
    ))?;
    let transport = source
        .get(start..end)
        .ok_or(Error::UnsupportedSource(
            "picture payload provenance is outside the source",
        ))?
        .to_vec();
    if decode_transport(&transport)? != semantic.data() {
        return Err(Error::UnsupportedSource(
            "picture payload provenance differs from parsed media bytes",
        ));
    }
    Ok((
        LocatedPicture {
            payload_span: start..end,
            payload_transport: transport,
            image_type: semantic.image_type,
            data: semantic.data().to_vec(),
        },
        cursor,
    ))
}

fn validate_supported_image(image_type: ImageType, payload: &[u8]) -> Result<(), Error> {
    let supported = match image_type {
        ImageType::Png => payload.starts_with(&[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]),
        ImageType::Jpeg => payload.starts_with(&[0xff, 0xd8]) && payload.ends_with(&[0xff, 0xd9]),
        ImageType::Emf | ImageType::Wmf | ImageType::Dib | ImageType::Pict | ImageType::Unknown => {
            return Err(Error::UnsupportedSource(
                "exact picture payload editing currently supports PNG and JPEG only",
            ));
        },
    };
    if !supported {
        return Err(Error::UnsupportedSource(
            "replacement bytes do not retain the declared picture format",
        ));
    }
    Ok(())
}

#[derive(Default)]
struct PictureControlSet {
    slots: u32,
}

impl PictureControlSet {
    fn accept(&mut self, control: &ControlWord<'_>) -> Result<(), Error> {
        let slot = match control {
            ControlWord::Emfblip
            | ControlWord::Pngblip
            | ControlWord::Jpegblip
            | ControlWord::Macpict
            | ControlWord::Pmmetafile(_)
            | ControlWord::Wmetafile(_)
            | ControlWord::Dibitmap(_)
            | ControlWord::Wbitmap(_) => 0,
            ControlWord::PictureWidth(_) => 1,
            ControlWord::PictureHeight(_) => 2,
            ControlWord::PictureGoalWidth(_) => 3,
            ControlWord::PictureGoalHeight(_) => 4,
            ControlWord::PictureScaleX(_) => 5,
            ControlWord::PictureScaleY(_) => 6,
            ControlWord::PictureScaled(_) => 7,
            ControlWord::PictureBitmap(_) => 8,
            ControlWord::PictureBitsPerPixel(_) => 9,
            ControlWord::PictureCropLeft(_) => 10,
            ControlWord::PictureCropRight(_) => 11,
            ControlWord::PictureCropTop(_) => 12,
            ControlWord::PictureCropBottom(_) => 13,
            ControlWord::WindowsBitmapBitsPerPixel(_) => 14,
            ControlWord::WindowsBitmapPlanes(_) => 15,
            ControlWord::WindowsBitmapWidthBytes(_) => 16,
            ControlWord::BlipTag(_) => 17,
            ControlWord::BlipUnitsPerInch(_) => 18,
            _ => {
                return Err(Error::UnsupportedSource(
                    "picture payload editing refuses unknown or dependent picture controls",
                ));
            },
        };
        let mask = 1u32 << slot;
        if self.slots & mask != 0 {
            return Err(Error::UnsupportedSource(
                "picture payload editing refuses duplicate or conflicting picture controls",
            ));
        }
        self.slots |= mask;
        Ok(())
    }
}

fn render_payload(template: &[u8], payload: &[u8]) -> Result<Vec<u8>, Error> {
    let expected_digits = payload.len().checked_mul(2).ok_or(Error::InputTooLarge {
        observed: usize::MAX,
        limit: MAX_EDITABLE_PICTURE_BYTES,
    })?;
    let actual_digits = template
        .iter()
        .filter(|byte| byte.is_ascii_hexdigit())
        .count();
    if actual_digits != expected_digits {
        return Err(Error::UnsupportedSource(
            "picture hexadecimal layout does not match its decoded length",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve(template.len())
        .map_err(|_error| allocation_error("rendered picture payload", template.len()))?;
    output.extend_from_slice(template);
    let mut nibbles = payload.iter().flat_map(|byte| [byte >> 4, byte & 0x0f]);
    for byte in &mut output {
        if byte.is_ascii_hexdigit() {
            let nibble = nibbles.next().ok_or(Error::UnsupportedSource(
                "picture hexadecimal layout ended unexpectedly",
            ))?;
            *byte = hex_digit(nibble, byte.is_ascii_uppercase());
        }
    }
    if nibbles.next().is_some() {
        return Err(Error::UnsupportedSource(
            "picture hexadecimal layout has too few digit positions",
        ));
    }
    Ok(output)
}

fn validate_exact_transport(current: &[u8], target: &[u8], data: &[u8]) -> Result<(), Error> {
    if current.len() != target.len()
        || !current
            .iter()
            .zip(target)
            .all(|(before, after)| before.is_ascii_hexdigit() == after.is_ascii_hexdigit())
        || current
            .iter()
            .zip(target)
            .any(|(before, after)| !before.is_ascii_hexdigit() && before != after)
    {
        return Err(Error::DurablePatch(
            "picture payload transport changed its whitespace layout".to_string(),
        ));
    }
    if decode_transport(target)? != data {
        return Err(Error::DurablePatch(
            "picture payload transport differs from its media bytes".to_string(),
        ));
    }
    Ok(())
}

fn decode_transport(input: &[u8]) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve(input.len() / 2)
        .map_err(|_error| allocation_error("decoded picture payload", input.len() / 2))?;
    let mut high = None;
    for byte in input
        .iter()
        .copied()
        .filter(|byte| !byte.is_ascii_whitespace())
    {
        let nibble = hex_nibble(byte).ok_or(Error::UnsupportedSource(
            "picture payload contains a non-hexadecimal byte",
        ))?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
        } else {
            high = Some(nibble);
        }
    }
    if high.is_some() {
        return Err(Error::UnsupportedSource(
            "picture payload contains an odd hexadecimal digit count",
        ));
    }
    Ok(output)
}

fn decode_hex(input: &str, limit: usize) -> Result<Vec<u8>, Error> {
    if !input.len().is_multiple_of(2) {
        return Err(Error::DurablePatch(
            "picture payload hexadecimal value has odd length".to_string(),
        ));
    }
    let output_len = input.len() / 2;
    if output_len > limit {
        return Err(Error::InputTooLarge {
            observed: output_len,
            limit,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|_error| allocation_error("durable picture payload", output_len))?;
    for pair in input.as_bytes().chunks_exact(2) {
        let high = pair.first().copied().and_then(hex_nibble).ok_or_else(|| {
            Error::DurablePatch("invalid picture payload hexadecimal".to_string())
        })?;
        let low = pair.get(1).copied().and_then(hex_nibble).ok_or_else(|| {
            Error::DurablePatch("invalid picture payload hexadecimal".to_string())
        })?;
        output.push((high << 4) | low);
    }
    Ok(output)
}

fn parse_picture_target(target: &str) -> Result<usize, Error> {
    target
        .strip_prefix("body:picture:")
        .and_then(|position| position.parse::<usize>().ok())
        .ok_or_else(|| Error::DurablePatch("invalid picture payload target".to_string()))
}

fn verify_candidate(
    source: &Snapshot,
    candidate: &Snapshot,
    operations: &[Operation],
) -> Result<(), Error> {
    if source.text() != candidate.text() || source.pictures().len() != candidate.pictures().len() {
        return Err(Error::UnsupportedSource(
            "picture payload edit changed unrelated document semantics",
        ));
    }
    for (position, (before, after)) in source
        .pictures()
        .iter()
        .zip(candidate.pictures())
        .enumerate()
    {
        let mut expected = before.clone();
        if let Some(operation) = operations.iter().find_map(|operation| match operation {
            Operation::PicturePayload(operation) if operation.position == position => {
                Some(operation)
            },
            _ => None,
        }) {
            expected.data = Cow::Owned(operation.after.clone());
        }
        if &expected != after {
            return Err(Error::UnsupportedSource(
                "picture metadata or an unselected payload changed during validation",
            ));
        }
    }
    Ok(())
}

const fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

const fn hex_digit(nibble: u8, uppercase: bool) -> u8 {
    match nibble {
        0..=9 => b'0' + nibble,
        _ if uppercase => b'A' + nibble - 10,
        _ => b'a' + nibble - 10,
    }
}

fn allocation_error(resource: &'static str, requested: usize) -> Error {
    Error::Rtf(RtfError::AllocationFailed {
        resource,
        requested,
    })
}

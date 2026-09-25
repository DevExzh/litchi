//! Durable semantic Ink edits.
//!
//! The wire operation retains the typed request that produced a source-bound
//! [`super::Patch`].  Applying it reconstructs that request and sends it
//! through the owning `Edit` authoring path.  The compact reverse closure is
//! used only to rebuild a private proof source for inverse replay; it is never
//! a permission to copy arbitrary package bytes into the published package.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::patch::{
    BlobBundle, BlobId, BlobLimitKind, Patch as CorePatch, PatchLimits, PatchOperation, Reversible,
    ReversibleOperation,
};
use litchi_core::{Position, patch::PatchError};
use litchi_opc::constants::content_type as ct;
use litchi_opc::phys_pkg::PhysPkgWriter;
use litchi_opc::{BlobPart, OpcPackage, OwnedRelationships, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use serde_json::Value;
use sha2::{Digest as _, Sha256};

use super::super::authoring::{
    AnchorGeometry, BaseProfile, FallbackImage, Geometry, HorizontalAlignment, HorizontalPosition,
    HorizontalRelativeFrom, ImageDimensions, ImageType, Placement, Point, Style, VerticalAlignment,
    VerticalPosition, VerticalRelativeFrom, WrapMode,
};
use super::{Destination, EditLimits, Patch, SemanticIntent, Snapshot};
use crate::ink::Limits;
use crate::package::story::{self, StoryKind};
use crate::{Error, Package, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const FORMAT_NAME: &str = "litchi-docx/ink";
const EDIT: &str = "ink.edit";
const RESTORE: &str = "ink.restore";
const NOOP: &str = "ink.noop";
const INTENT_HEADER: &[u8] = b"LIE1";
const RESTORE_HEADER: &[u8] = b"LIR1";
const DELTA_HEADER: &[u8] = b"LIK1";
const MAX_INTENTS: usize = 65_536;
const MAX_INTENT_BYTES: usize = 512 * 1024 * 1024;
const MAX_DELTA_RECORDS: usize = 65_536;
const MAX_MEMBER_NAME: usize = 4096;
const MAX_CONTENT_TYPE: usize = 4096;
const MAX_DELTA_BYTES: usize = 512 * 1024 * 1024;
const MAX_RELATIONSHIPS_BYTES: usize = 8 * 1024 * 1024;
const MAX_NON_PART_MEMBERS: usize = 1_048_576;
const MAX_NON_PART_METADATA_BYTES: usize = 128 * 1024 * 1024;
// Reserve room for the local ZIP header, central-directory entry, and bounded
// member-name framing before handing a proof package to `PhysPkgWriter`.
const PROOF_ENTRY_OVERHEAD: usize = 128;

/// Convert one exact reversible Ink commit to the common deterministic patch
/// envelope.  The forward side contains only typed semantic intent.  The
/// reverse side additionally retains a bounded source closure so inverse
/// replay can be proved from the current target without retaining a second
/// complete package.
impl Patch {
    /// Encode a source-bound typed Ink edit.
    ///
    /// A durable round trip is `Edit::commit`,
    /// `commit.patch().to_durable(patch_limits)`,
    /// `CorePatch::<Reversible>::to_deterministic_json` and
    /// `from_deterministic_json`, followed by
    /// [`Package::apply_durable_ink_patch`].  The source guard binds the
    /// logical OPC content-types member, relationship members and presence,
    /// part names/types/payloads, and the public non-part name/reason
    /// inventory.  ZIP compression bytes and other physical archive details
    /// are outside that logical OPC model.  Prepared payloads are decoded only
    /// through the canonical importer and are then replayed through `Edit`.
    pub fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> std::result::Result<CorePatch<Reversible>, PatchError> {
        let before_artifact = package_fingerprint(&self.before.package)?;
        let after_artifact = package_fingerprint(&self.after.package)?;
        if self.is_empty() {
            let forward = PatchOperation::new(
                limits,
                NOOP,
                "package",
                preconditions(&before_artifact, &after_artifact, None, "noop"),
                Value::Null,
            )?;
            let inverse = PatchOperation::new(
                limits,
                NOOP,
                "package",
                preconditions(&after_artifact, &before_artifact, None, "noop"),
                Value::Null,
            )?;
            return CorePatch::<Reversible>::new(
                limits,
                FORMAT_NAME,
                [ReversibleOperation::new(forward, inverse)],
                BlobBundle::new(limits.blobs()),
                BlobBundle::new(limits.blobs()),
            );
        }

        let blob_limits = limits.blobs();
        if blob_limits.max_blobs() == 0 {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::Count,
                observed: 1,
                limit: 0,
            });
        }
        if blob_limits.max_blob_bytes() == 0 {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::BlobBytes,
                observed: 1,
                limit: 0,
            });
        }
        if blob_limits.max_total_bytes() == 0 {
            return Err(PatchError::BlobLimit {
                kind: BlobLimitKind::TotalBytes,
                observed: 1,
                limit: 0,
            });
        }
        let blob_maximum = blob_limits
            .max_blob_bytes()
            .min(blob_limits.max_total_bytes())
            .min(MAX_DELTA_BYTES);
        let intent = encode_intents(self.intent.as_slice(), blob_maximum)?;
        let restore_overhead = RESTORE_HEADER.len().saturating_add(16);
        let delta_maximum = blob_maximum
            .checked_sub(restore_overhead.saturating_add(intent.len()))
            .ok_or(PatchError::InvalidText {
                field: "Ink durable restore blob size",
            })?;
        let reverse_delta = if self.reversed {
            encode_delta(&self.before.package, &self.after.package, delta_maximum)?
        } else {
            encode_delta(&self.after.package, &self.before.package, delta_maximum)?
        };
        let kind = intent_kind(self.intent.as_slice());
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let restore = encode_restore(&intent, &reverse_delta, blob_maximum)?;
        let intent_id = BlobId::of(&intent);
        let restore_id = BlobId::of(&restore);
        let intent = Arc::<[u8]>::from(intent);
        let restore = Arc::<[u8]>::from(restore);
        let (forward, inverse) = if self.reversed {
            forward_blobs.insert_shared(Arc::clone(&restore))?;
            reverse_blobs.insert_shared(Arc::clone(&intent))?;
            (
                PatchOperation::new(
                    limits,
                    RESTORE,
                    "package",
                    preconditions(
                        &before_artifact,
                        &after_artifact,
                        Some(("restore_sha256", &restore_id)),
                        kind,
                    ),
                    Value::Null,
                )?,
                PatchOperation::new(
                    limits,
                    EDIT,
                    "package",
                    preconditions(
                        &after_artifact,
                        &before_artifact,
                        Some(("intent_sha256", &intent_id)),
                        kind,
                    ),
                    Value::Null,
                )?,
            )
        } else {
            forward_blobs.insert_shared(Arc::clone(&intent))?;
            reverse_blobs.insert_shared(Arc::clone(&restore))?;
            (
                PatchOperation::new(
                    limits,
                    EDIT,
                    "package",
                    preconditions(
                        &before_artifact,
                        &after_artifact,
                        Some(("intent_sha256", &intent_id)),
                        kind,
                    ),
                    Value::Null,
                )?,
                PatchOperation::new(
                    limits,
                    RESTORE,
                    "package",
                    preconditions(
                        &after_artifact,
                        &before_artifact,
                        Some(("restore_sha256", &restore_id)),
                        kind,
                    ),
                    Value::Null,
                )?,
            )
        };
        CorePatch::<Reversible>::new(
            limits,
            FORMAT_NAME,
            [ReversibleOperation::new(forward, inverse)],
            forward_blobs,
            reverse_blobs,
        )
    }
}

impl Package {
    /// Apply one durable typed Ink edit atomically.
    ///
    /// Forward operations replay through [`super::Edit`].  Inverse operations
    /// first rebuild a private source from their exact closure, replay the
    /// original forward intent there, and then apply that checked in-memory
    /// inverse.  No durable blob is ever installed as an unverified package.
    pub fn apply_durable_ink_patch<Mode>(&mut self, patch: &CorePatch<Mode>) -> Result<Snapshot> {
        self.apply_durable_ink_patch_with_limits(patch, EditLimits::default())
    }

    /// Apply a durable Ink edit with explicit replay and reconstruction bounds.
    pub fn apply_durable_ink_patch_with_limits<Mode>(
        &mut self,
        patch: &CorePatch<Mode>,
        limits: EditLimits,
    ) -> Result<Snapshot> {
        let limits = limits.validate()?;
        let fingerprint_bounds = FingerprintBounds::from_edit_limits(limits);
        self.ensure_story_opc_current("apply_durable_ink_patch")?;
        if patch.format() != FORMAT_NAME {
            return Err(invalid_durable("unsupported format"));
        }
        if patch.operations().len() != 1 {
            return Err(invalid_durable("Ink durable patches must contain one edit"));
        }
        let operation = &patch.operations()[0];
        if operation.target != "package" || !operation.value.is_null() {
            return Err(invalid_durable("invalid Ink durable target or value"));
        }
        // Run the same bounded source admission that owns every semantic edit
        // before hashing the package.  In particular, this prevents a durable
        // no-op or stale-source check from walking an unbounded package before
        // the caller's EditLimits have had a chance to refuse it.
        let source_snapshot = {
            let edit = self.edit_ink_with_limits(limits)?;
            edit.snapshot().clone()
        };
        let source_artifact =
            package_fingerprint_with_bounds(self.opc_package(), fingerprint_bounds)
                .map_err(|error| invalid_durable(error.to_string()))?;
        if text_precondition(operation, "artifact_sha256")? != source_artifact {
            return Err(invalid_durable("Ink durable source is stale"));
        }
        let target_artifact = text_precondition(operation, "target_sha256")?;
        let kind = text_precondition(operation, "kind")?;
        match operation.op.as_str() {
            NOOP => {
                if operation.preconditions.len() != 3
                    || kind != "noop"
                    || target_artifact != source_artifact
                    || !patch.blobs().is_empty()
                {
                    return Err(invalid_durable("invalid Ink durable no-op"));
                }
                Ok(source_snapshot)
            },
            EDIT => {
                if operation.preconditions.len() != 4 {
                    return Err(invalid_durable("invalid Ink durable edit"));
                }
                let intent_id = text_precondition(operation, "intent_sha256")?;
                let intent_bytes = single_blob_by_hex(patch, intent_id)?;
                let intents = decode_intents(intent_bytes, limits)?;
                if intent_kind(intents.as_slice()) != kind {
                    return Err(invalid_durable("Ink durable edit kind mismatch"));
                }
                let commit = replay_intents(self.opc_package().clone(), &intents, limits)?;
                let generated = package_fingerprint_with_bounds(
                    &commit.patch().after.package,
                    fingerprint_bounds,
                )
                .map_err(|error| invalid_durable(error.to_string()))?;
                if generated != target_artifact {
                    return Err(invalid_durable("Ink durable typed replay target mismatch"));
                }
                self.apply_ink_patch(commit.patch())
            },
            RESTORE => {
                if operation.preconditions.len() != 4 {
                    return Err(invalid_durable("invalid Ink durable restore"));
                }
                let restore_id = text_precondition(operation, "restore_sha256")?;
                let restore_bytes = single_blob_by_hex(patch, restore_id)?;
                let (intent_bytes, delta_bytes) = decode_restore(restore_bytes, limits)?;
                let intents = decode_intents(intent_bytes, limits)?;
                if intent_kind(intents.as_slice()) != kind {
                    return Err(invalid_durable("Ink durable restore kind mismatch"));
                }
                let stories = story::capture(self.opc_package(), limits.inventory.stories)?;
                let story_names: BTreeSet<_> = stories
                    .stories()
                    .iter()
                    .map(|story| story.part().as_str().to_owned())
                    .collect();
                let records = decode_delta(delta_bytes, limits, &story_names)?;
                validate_delta_source(self.opc_package(), &records)?;
                let mut source_candidate = self.opc_package().clone();
                apply_delta(
                    &mut source_candidate,
                    &records,
                    limits.inventory,
                    limits.max_staged_bytes,
                )?;
                let reconstructed =
                    package_fingerprint_with_bounds(&source_candidate, fingerprint_bounds)
                        .map_err(|error| invalid_durable(error.to_string()))?;
                if reconstructed != target_artifact {
                    return Err(invalid_durable("Ink durable restore source mismatch"));
                }
                let commit = replay_intents(source_candidate, &intents, limits)?;
                let replayed_target = package_fingerprint_with_bounds(
                    &commit.patch().after.package,
                    fingerprint_bounds,
                )
                .map_err(|error| invalid_durable(error.to_string()))?;
                if replayed_target != source_artifact {
                    return Err(invalid_durable(
                        "Ink durable restore forward replay mismatch",
                    ));
                }
                self.apply_ink_patch(&commit.patch().inverse())
            },
            _ => Err(invalid_durable("unsupported Ink durable operation")),
        }
    }
}

fn replay_intents(
    source: OpcPackage,
    intents: &[SemanticIntent],
    limits: EditLimits,
) -> Result<super::Commit> {
    if intents.is_empty() || intents.len() > MAX_INTENTS {
        return Err(invalid_durable("invalid Ink durable intent count"));
    }
    let source_package = Package::from_opc_package(source)?;
    let mut edit = source_package.edit_ink_with_limits(limits)?;
    for intent in intents {
        match intent {
            SemanticIntent::Insert {
                destination,
                payload,
                style,
            } => edit.insert(*destination, payload.clone(), style.clone())?,
            SemanticIntent::Replace {
                position,
                payload,
                fallback,
            } => {
                if !edit.replace(Position::new(*position), payload.clone(), fallback.clone())? {
                    return Err(invalid_durable(
                        "Ink durable replace intent produced no change",
                    ));
                }
            },
            SemanticIntent::Remove { position } => {
                if !edit.remove(Position::new(*position))? {
                    return Err(invalid_durable(
                        "Ink durable remove intent produced no change",
                    ));
                }
            },
        }
    }
    let commit = edit.commit()?;
    if !commit.changed() {
        return Err(invalid_durable(
            "Ink durable intent produced no changed commit",
        ));
    }
    Ok(commit)
}

fn preconditions(
    artifact: &str,
    target: &str,
    blob: Option<(&str, &BlobId)>,
    kind: &str,
) -> BTreeMap<String, Value> {
    let mut values = BTreeMap::new();
    values.insert("artifact_sha256".into(), Value::String(artifact.to_owned()));
    values.insert("target_sha256".into(), Value::String(target.to_owned()));
    if let Some((key, blob)) = blob {
        values.insert(key.into(), Value::String(blob.as_hex()));
    }
    values.insert("kind".into(), Value::String(kind.to_owned()));
    values
}

fn text_precondition<'a>(
    operation: &'a PatchOperation,
    key: &str,
) -> std::result::Result<&'a str, Error> {
    operation
        .preconditions
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| invalid_durable("missing Ink durable precondition"))
}

fn single_blob_by_hex<'a, Mode>(patch: &'a CorePatch<Mode>, identifier: &str) -> Result<&'a [u8]> {
    if patch.blobs().len() != 1 {
        return Err(invalid_durable("Ink durable operation has extra blobs"));
    }
    patch
        .blobs()
        .ids()
        .find(|candidate| candidate.as_hex() == identifier)
        .and_then(|id| patch.blobs().get(id))
        .ok_or_else(|| invalid_durable("missing Ink durable semantic blob"))
}

fn invalid_durable(message: impl Into<String>) -> Error {
    Error::InvalidFormat(format!(
        "invalid DOCX Ink durable patch: {}",
        message.into()
    ))
}

fn intent_kind(intents: &[SemanticIntent]) -> &'static str {
    let mut kind = None;
    for intent in intents {
        let current = match intent {
            SemanticIntent::Insert { .. } => "insert",
            SemanticIntent::Replace { .. } => "replace",
            SemanticIntent::Remove { .. } => "remove",
        };
        if kind.is_some_and(|previous| previous != current) {
            return "mixed";
        }
        kind = Some(current);
    }
    kind.unwrap_or("noop")
}

struct Encoder {
    bytes: Vec<u8>,
    maximum: usize,
}

impl Encoder {
    fn new(header: &[u8], maximum: usize) -> std::result::Result<Self, PatchError> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(header.len())
            .map_err(|_| PatchError::Allocation)?;
        bytes.extend_from_slice(header);
        Ok(Self { bytes, maximum })
    }

    fn push(&mut self, value: u8) -> std::result::Result<(), PatchError> {
        self.append(&[value])
    }

    fn u32(&mut self, value: u32) -> std::result::Result<(), PatchError> {
        self.append(&value.to_le_bytes())
    }

    fn u64(&mut self, value: u64) -> std::result::Result<(), PatchError> {
        self.append(&value.to_le_bytes())
    }

    fn i64(&mut self, value: i64) -> std::result::Result<(), PatchError> {
        self.append(&value.to_le_bytes())
    }

    fn bytes(&mut self, value: &[u8]) -> std::result::Result<(), PatchError> {
        let length = u64::try_from(value.len()).map_err(|_| PatchError::InvalidText {
            field: "Ink durable semantic byte length",
        })?;
        self.u64(length)?;
        self.append(value)
    }

    fn append(&mut self, value: &[u8]) -> std::result::Result<(), PatchError> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or(PatchError::InvalidText {
                field: "Ink durable semantic byte length",
            })?;
        if length > self.maximum {
            return Err(PatchError::InvalidText {
                field: "Ink durable semantic byte length",
            });
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|_| PatchError::Allocation)?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn encode_intents(
    intents: &[SemanticIntent],
    maximum: usize,
) -> std::result::Result<Vec<u8>, PatchError> {
    if intents.is_empty() || intents.len() > MAX_INTENTS {
        return Err(PatchError::InvalidText {
            field: "Ink durable intent count",
        });
    }
    let mut encoder = Encoder::new(INTENT_HEADER, maximum)?;
    encoder.u64(intents.len() as u64)?;
    for intent in intents {
        match intent {
            SemanticIntent::Insert {
                destination,
                payload,
                style,
            } => {
                encoder.push(0)?;
                encoder.push(story_kind_code(destination.story.kind))?;
                encoder.u64(destination.story.position.get() as u64)?;
                encoder.u64(destination.paragraph.get() as u64)?;
                encoder.bytes(payload.as_bytes())?;
                encode_style(&mut encoder, style)?;
            },
            SemanticIntent::Replace {
                position,
                payload,
                fallback,
            } => {
                encoder.push(1)?;
                encoder.u64(*position as u64)?;
                encoder.bytes(payload.as_bytes())?;
                encode_fallback(&mut encoder, fallback.as_ref())?;
            },
            SemanticIntent::Remove { position } => {
                encoder.push(2)?;
                encoder.u64(*position as u64)?;
            },
        }
    }
    Ok(encoder.finish())
}

fn encode_style(encoder: &mut Encoder, style: &Style) -> std::result::Result<(), PatchError> {
    encoder.push(style.durable_kind())?;
    if let Some(profile) = style.base_profile() {
        encoder.push(match profile {
            BaseProfile::WordTextXml => 0,
            BaseProfile::InkContent => 1,
        })?;
        return Ok(());
    }
    let placement = style.placement().ok_or(PatchError::InvalidText {
        field: "Ink durable style placement",
    })?;
    encode_placement(encoder, placement)?;
    encode_fallback(encoder, style.fallback())
}

fn encode_placement(
    encoder: &mut Encoder,
    placement: &Placement,
) -> std::result::Result<(), PatchError> {
    match placement {
        Placement::Inline(extent) => {
            encoder.push(0)?;
            encoder.u64(extent.width_emu())?;
            encoder.u64(extent.height_emu())?;
        },
        Placement::Anchor(anchor) => {
            encoder.push(1)?;
            let extent = anchor_extent(*anchor);
            encoder.u64(extent.width_emu())?;
            encoder.u64(extent.height_emu())?;
            let point = anchor.durable_simple_position();
            encoder.i64(point.x())?;
            encoder.i64(point.y())?;
            encode_horizontal(encoder, anchor.durable_horizontal())?;
            encode_vertical(encoder, anchor.durable_vertical())?;
            if !matches!(anchor.durable_wrap(), WrapMode::None) {
                return Err(PatchError::InvalidText {
                    field: "Ink durable anchor wrap mode",
                });
            }
            encoder.push(0)?;
            encoder.u32(anchor.durable_relative_height())?;
            encoder.push(u8::from(anchor.durable_behind_document()))?;
            encoder.push(u8::from(anchor.durable_locked()))?;
            encoder.push(u8::from(anchor.durable_layout_in_cell()))?;
            encoder.push(u8::from(anchor.durable_allow_overlap()))?;
        },
    }
    Ok(())
}

fn anchor_extent(anchor: AnchorGeometry) -> Geometry {
    // `Placement::extent` is intentionally private to the authoring module;
    // reconstructing from the emitted public geometry keeps this codec stable.
    let point = anchor.durable_simple_position();
    let _ = point;
    // The extent is available through the placement itself in the caller.  An
    // anchor's private extent is exposed by this crate-private adapter.
    anchor.durable_extent()
}

fn encode_horizontal(
    encoder: &mut Encoder,
    position: HorizontalPosition,
) -> std::result::Result<(), PatchError> {
    match position {
        HorizontalPosition::Align {
            relative_from,
            alignment,
        } => {
            encoder.push(0)?;
            encoder.push(horizontal_relative_code(relative_from))?;
            encoder.push(horizontal_alignment_code(alignment))?;
        },
        HorizontalPosition::Offset {
            relative_from,
            offset,
        } => {
            encoder.push(1)?;
            encoder.push(horizontal_relative_code(relative_from))?;
            encoder.i64(i64::from(offset))?;
        },
    }
    Ok(())
}

fn encode_vertical(
    encoder: &mut Encoder,
    position: VerticalPosition,
) -> std::result::Result<(), PatchError> {
    match position {
        VerticalPosition::Align {
            relative_from,
            alignment,
        } => {
            encoder.push(0)?;
            encoder.push(vertical_relative_code(relative_from))?;
            encoder.push(vertical_alignment_code(alignment))?;
        },
        VerticalPosition::Offset {
            relative_from,
            offset,
        } => {
            encoder.push(1)?;
            encoder.push(vertical_relative_code(relative_from))?;
            encoder.i64(i64::from(offset))?;
        },
    }
    Ok(())
}

fn encode_fallback(
    encoder: &mut Encoder,
    fallback: Option<&FallbackImage>,
) -> std::result::Result<(), PatchError> {
    let Some(fallback) = fallback else {
        return encoder.push(0);
    };
    encoder.push(1)?;
    encoder.push(match fallback.media_type() {
        ImageType::Png => 0,
        ImageType::Jpeg => 1,
    })?;
    encoder.u32(fallback.dimensions().width())?;
    encoder.u32(fallback.dimensions().height())?;
    encoder.bytes(fallback.as_bytes())
}

fn encode_restore(
    intent: &[u8],
    delta: &[u8],
    maximum: usize,
) -> std::result::Result<Vec<u8>, PatchError> {
    let mut encoder = Encoder::new(RESTORE_HEADER, maximum)?;
    encoder.bytes(intent)?;
    encoder.bytes(delta)?;
    Ok(encoder.finish())
}

fn story_kind_code(kind: StoryKind) -> u8 {
    match kind {
        StoryKind::Main => 0,
        StoryKind::Header => 1,
        StoryKind::Footer => 2,
        StoryKind::Footnotes => 3,
        StoryKind::Endnotes => 4,
        StoryKind::Comments => 5,
        StoryKind::Glossary => 6,
    }
}

fn story_kind_from_code(code: u8) -> Result<StoryKind> {
    match code {
        0 => Ok(StoryKind::Main),
        1 => Ok(StoryKind::Header),
        2 => Ok(StoryKind::Footer),
        3 => Ok(StoryKind::Footnotes),
        4 => Ok(StoryKind::Endnotes),
        5 => Ok(StoryKind::Comments),
        6 => Ok(StoryKind::Glossary),
        _ => Err(invalid_durable("unknown Ink durable story kind")),
    }
}

fn horizontal_relative_code(value: HorizontalRelativeFrom) -> u8 {
    match value {
        HorizontalRelativeFrom::Margin => 0,
        HorizontalRelativeFrom::Page => 1,
        HorizontalRelativeFrom::Column => 2,
        HorizontalRelativeFrom::Character => 3,
        HorizontalRelativeFrom::LeftMargin => 4,
        HorizontalRelativeFrom::RightMargin => 5,
        HorizontalRelativeFrom::InsideMargin => 6,
        HorizontalRelativeFrom::OutsideMargin => 7,
    }
}

fn horizontal_relative_from_code(code: u8) -> Result<HorizontalRelativeFrom> {
    match code {
        0 => Ok(HorizontalRelativeFrom::Margin),
        1 => Ok(HorizontalRelativeFrom::Page),
        2 => Ok(HorizontalRelativeFrom::Column),
        3 => Ok(HorizontalRelativeFrom::Character),
        4 => Ok(HorizontalRelativeFrom::LeftMargin),
        5 => Ok(HorizontalRelativeFrom::RightMargin),
        6 => Ok(HorizontalRelativeFrom::InsideMargin),
        7 => Ok(HorizontalRelativeFrom::OutsideMargin),
        _ => Err(invalid_durable(
            "unknown Ink durable horizontal relative mode",
        )),
    }
}

fn horizontal_alignment_code(value: HorizontalAlignment) -> u8 {
    match value {
        HorizontalAlignment::Left => 0,
        HorizontalAlignment::Right => 1,
        HorizontalAlignment::Center => 2,
        HorizontalAlignment::Inside => 3,
        HorizontalAlignment::Outside => 4,
    }
}

fn horizontal_alignment_from_code(code: u8) -> Result<HorizontalAlignment> {
    match code {
        0 => Ok(HorizontalAlignment::Left),
        1 => Ok(HorizontalAlignment::Right),
        2 => Ok(HorizontalAlignment::Center),
        3 => Ok(HorizontalAlignment::Inside),
        4 => Ok(HorizontalAlignment::Outside),
        _ => Err(invalid_durable("unknown Ink durable horizontal alignment")),
    }
}

fn vertical_relative_code(value: VerticalRelativeFrom) -> u8 {
    match value {
        VerticalRelativeFrom::Margin => 0,
        VerticalRelativeFrom::Page => 1,
        VerticalRelativeFrom::Paragraph => 2,
        VerticalRelativeFrom::Line => 3,
        VerticalRelativeFrom::TopMargin => 4,
        VerticalRelativeFrom::BottomMargin => 5,
        VerticalRelativeFrom::InsideMargin => 6,
        VerticalRelativeFrom::OutsideMargin => 7,
    }
}

fn vertical_relative_from_code(code: u8) -> Result<VerticalRelativeFrom> {
    match code {
        0 => Ok(VerticalRelativeFrom::Margin),
        1 => Ok(VerticalRelativeFrom::Page),
        2 => Ok(VerticalRelativeFrom::Paragraph),
        3 => Ok(VerticalRelativeFrom::Line),
        4 => Ok(VerticalRelativeFrom::TopMargin),
        5 => Ok(VerticalRelativeFrom::BottomMargin),
        6 => Ok(VerticalRelativeFrom::InsideMargin),
        7 => Ok(VerticalRelativeFrom::OutsideMargin),
        _ => Err(invalid_durable(
            "unknown Ink durable vertical relative mode",
        )),
    }
}

fn vertical_alignment_code(value: VerticalAlignment) -> u8 {
    match value {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Bottom => 1,
        VerticalAlignment::Center => 2,
        VerticalAlignment::Inside => 3,
        VerticalAlignment::Outside => 4,
    }
}

fn vertical_alignment_from_code(code: u8) -> Result<VerticalAlignment> {
    match code {
        0 => Ok(VerticalAlignment::Top),
        1 => Ok(VerticalAlignment::Bottom),
        2 => Ok(VerticalAlignment::Center),
        3 => Ok(VerticalAlignment::Inside),
        4 => Ok(VerticalAlignment::Outside),
        _ => Err(invalid_durable("unknown Ink durable vertical alignment")),
    }
}

struct IntentDecoder<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> IntentDecoder<'a> {
    fn new(bytes: &'a [u8], header: &[u8]) -> Result<Self> {
        if !bytes.starts_with(header) {
            return Err(invalid_durable("invalid Ink durable semantic header"));
        }
        Ok(Self {
            bytes,
            offset: header.len(),
        })
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| invalid_durable("Ink durable semantic offset overflow"))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| invalid_durable("truncated Ink durable semantic blob"))?;
        self.offset = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8> {
        self.take(1).map(|bytes| bytes[0])
    }

    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.offset)
    }

    fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(
            self.take(4)?
                .try_into()
                .map_err(|_| invalid_durable("invalid Ink durable u32"))?,
        ))
    }

    fn u64(&mut self, resource: &'static str, maximum: usize) -> Result<usize> {
        let value = usize::try_from(u64::from_le_bytes(
            self.take(8)?
                .try_into()
                .map_err(|_| invalid_durable("invalid Ink durable u64"))?,
        ))
        .map_err(|_| invalid_durable("Ink durable integer exceeds this platform"))?;
        if value > maximum {
            return Err(Error::InkLimit {
                resource,
                actual: value,
                maximum,
            });
        }
        Ok(value)
    }

    fn i64(&mut self) -> Result<i64> {
        Ok(i64::from_le_bytes(
            self.take(8)?
                .try_into()
                .map_err(|_| invalid_durable("invalid Ink durable i64"))?,
        ))
    }

    fn bool(&mut self, resource: &'static str) -> Result<bool> {
        match self.u8()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(invalid_durable(format!(
                "invalid Ink durable boolean for {resource}"
            ))),
        }
    }

    fn bytes(&mut self, resource: &'static str, maximum: usize) -> Result<Vec<u8>> {
        let source = self.bytes_slice(resource, maximum)?;
        let length = source.len();
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink durable semantic bytes",
                source,
            })?;
        bytes.extend_from_slice(source);
        Ok(bytes)
    }

    fn bytes_slice(&mut self, resource: &'static str, maximum: usize) -> Result<&'a [u8]> {
        let length = self.u64(resource, maximum)?;
        self.take(length)
    }
}

fn decode_intents(bytes: &[u8], limits: EditLimits) -> Result<Vec<SemanticIntent>> {
    let maximum = MAX_INTENT_BYTES.min(limits.max_staged_bytes);
    if bytes.len() > maximum {
        return Err(Error::InkLimit {
            resource: "durable intent bytes",
            actual: bytes.len(),
            maximum,
        });
    }
    let mut decoder = IntentDecoder::new(bytes, INTENT_HEADER)?;
    let count = decoder.u64(
        "durable intent count",
        MAX_INTENTS.min(limits.max_operations),
    )?;
    if count == 0 {
        return Err(invalid_durable("empty Ink durable intent"));
    }
    const MIN_INTENT_RECORD_BYTES: usize = 1 + 8;
    if count > decoder.remaining() / MIN_INTENT_RECORD_BYTES {
        return Err(invalid_durable("truncated Ink durable intent records"));
    }
    let retained_bytes = count
        .checked_mul(size_of::<SemanticIntent>())
        .ok_or_else(|| invalid_durable("Ink durable intent allocation overflow"))?;
    if retained_bytes > limits.max_staged_bytes {
        return Err(Error::InkLimit {
            resource: "durable intent records",
            actual: retained_bytes,
            maximum: limits.max_staged_bytes,
        });
    }
    let mut intents = Vec::new();
    intents
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink durable intents",
            source,
        })?;
    for _ in 0..count {
        match decoder.u8()? {
            0 => {
                let kind = story_kind_from_code(decoder.u8()?)?;
                let story_position = decoder.u64("durable story position", usize::MAX)?;
                let paragraph = decoder.u64("durable paragraph position", usize::MAX)?;
                let payload = decode_prepared(
                    decoder.bytes_slice(
                        "durable prepared Ink bytes",
                        limits.inventory.max_payload_bytes,
                    )?,
                    limits.inventory.max_payload_bytes,
                )?;
                let style = decode_style(&mut decoder, limits.inventory.max_payload_bytes)?;
                intents.push(SemanticIntent::Insert {
                    destination: Destination::new(
                        crate::ink::Location::new(kind, Position::new(story_position)),
                        Position::new(paragraph),
                    ),
                    payload,
                    style,
                });
            },
            1 => {
                let position = decoder.u64("durable annotation position", usize::MAX)?;
                let payload = decode_prepared(
                    decoder.bytes_slice(
                        "durable prepared Ink bytes",
                        limits.inventory.max_payload_bytes,
                    )?,
                    limits.inventory.max_payload_bytes,
                )?;
                let fallback = decode_fallback(&mut decoder, limits.inventory.max_payload_bytes)?;
                intents.push(SemanticIntent::Replace {
                    position,
                    payload,
                    fallback,
                });
            },
            2 => {
                let position = decoder.u64("durable annotation position", usize::MAX)?;
                intents.push(SemanticIntent::Remove { position });
            },
            _ => return Err(invalid_durable("unknown Ink durable intent")),
        }
    }
    if decoder.offset != bytes.len() {
        return Err(invalid_durable("trailing Ink durable intent bytes"));
    }
    Ok(intents)
}

fn decode_prepared(bytes: &[u8], maximum: usize) -> Result<litchi_drawingml::ink::Prepared> {
    let authoring_limits = litchi_drawingml::ink::AuthoringLimits {
        max_source_bytes: maximum,
        max_output_bytes: maximum,
        ..Default::default()
    };
    litchi_drawingml::ink::Prepared::from_bytes_with_limits(bytes, authoring_limits)
        .map_err(|error| invalid_durable(format!("invalid canonical Ink payload: {error}")))
}

fn decode_style(decoder: &mut IntentDecoder<'_>, maximum: usize) -> Result<Style> {
    let kind = decoder.u8()?;
    if kind == 0 {
        return Ok(Style::base(match decoder.u8()? {
            0 => BaseProfile::WordTextXml,
            1 => BaseProfile::InkContent,
            _ => return Err(invalid_durable("unknown Ink durable base profile")),
        }));
    }
    let placement = decode_placement(decoder)?;
    let fallback = decode_fallback(decoder, maximum)?
        .ok_or_else(|| invalid_durable("Ink durable drawing style has no fallback"))?;
    match kind {
        1 => Ok(Style::drawing(placement, fallback)),
        2 => Ok(Style::canvas(placement, fallback)),
        3 => Ok(Style::group(placement, fallback)),
        _ => Err(invalid_durable("unknown Ink durable style")),
    }
}

fn decode_placement(decoder: &mut IntentDecoder<'_>) -> Result<Placement> {
    let kind = decoder.u8()?;
    let width = decoder.u64("durable geometry width", u64::MAX as usize)? as u64;
    let height = decoder.u64("durable geometry height", u64::MAX as usize)? as u64;
    let extent = Geometry::new(width, height)
        .map_err(|error| invalid_durable(format!("invalid Ink durable geometry: {error}")))?;
    match kind {
        0 => Ok(Placement::inline(extent)),
        1 => {
            let point = Point::new(decoder.i64()?, decoder.i64()?)
                .map_err(|error| invalid_durable(format!("invalid Ink durable point: {error}")))?;
            let horizontal = decode_horizontal(decoder)?;
            let vertical = decode_vertical(decoder)?;
            if decoder.bool("wrap mode")? {
                return Err(invalid_durable("unsupported Ink durable wrap mode"));
            }
            let anchor = AnchorGeometry::new(extent)
                .with_simple_position(point)
                .with_horizontal(horizontal)
                .with_vertical(vertical)
                .with_relative_height(decoder.u32()?)
                .with_behind_document(decoder.bool("behind document")?)
                .with_locked(decoder.bool("locked")?)
                .with_layout_in_cell(decoder.bool("layout in cell")?)
                .with_allow_overlap(decoder.bool("allow overlap")?);
            Ok(Placement::anchor(anchor))
        },
        _ => Err(invalid_durable("unknown Ink durable placement")),
    }
}

fn decode_horizontal(decoder: &mut IntentDecoder<'_>) -> Result<HorizontalPosition> {
    let kind = decoder.u8()?;
    let relative = horizontal_relative_from_code(decoder.u8()?)?;
    match kind {
        0 => Ok(HorizontalPosition::align(
            relative,
            horizontal_alignment_from_code(decoder.u8()?)?,
        )),
        1 => Ok(HorizontalPosition::offset(
            relative,
            i32::try_from(decoder.i64()?)
                .map_err(|_| invalid_durable("Ink durable horizontal offset out of range"))?,
        )),
        _ => Err(invalid_durable("unknown Ink durable horizontal mode")),
    }
}

fn decode_vertical(decoder: &mut IntentDecoder<'_>) -> Result<VerticalPosition> {
    let kind = decoder.u8()?;
    let relative = vertical_relative_from_code(decoder.u8()?)?;
    match kind {
        0 => Ok(VerticalPosition::align(
            relative,
            vertical_alignment_from_code(decoder.u8()?)?,
        )),
        1 => Ok(VerticalPosition::offset(
            relative,
            i32::try_from(decoder.i64()?)
                .map_err(|_| invalid_durable("Ink durable vertical offset out of range"))?,
        )),
        _ => Err(invalid_durable("unknown Ink durable vertical mode")),
    }
}

fn decode_fallback(
    decoder: &mut IntentDecoder<'_>,
    maximum: usize,
) -> Result<Option<FallbackImage>> {
    match decoder.u8()? {
        0 => Ok(None),
        1 => {
            let media = match decoder.u8()? {
                0 => ImageType::Png,
                1 => ImageType::Jpeg,
                _ => return Err(invalid_durable("unknown Ink durable image type")),
            };
            let dimensions =
                ImageDimensions::new(decoder.u32()?, decoder.u32()?).map_err(|error| {
                    invalid_durable(format!("invalid Ink durable dimensions: {error}"))
                })?;
            let bytes = decoder.bytes(
                "durable fallback image bytes",
                maximum.min(16 * 1024 * 1024),
            )?;
            FallbackImage::new(Arc::new(bytes), media, dimensions)
                .map(Some)
                .map_err(|error| invalid_durable(format!("invalid Ink durable fallback: {error}")))
        },
        _ => Err(invalid_durable("invalid Ink durable fallback marker")),
    }
}

fn decode_restore(bytes: &[u8], limits: EditLimits) -> Result<(&[u8], &[u8])> {
    let maximum = MAX_INTENT_BYTES
        .min(limits.max_staged_bytes)
        .checked_add(MAX_DELTA_BYTES.min(limits.max_package_bytes))
        .ok_or_else(|| invalid_durable("durable restore bound overflow"))?;
    if bytes.len() > maximum {
        return Err(Error::InkLimit {
            resource: "durable restore bytes",
            actual: bytes.len(),
            maximum,
        });
    }
    let mut decoder = IntentDecoder::new(bytes, RESTORE_HEADER)?;
    let intent = decoder.bytes_slice(
        "durable restore intent bytes",
        MAX_INTENT_BYTES.min(limits.max_staged_bytes),
    )?;
    let delta = decoder.bytes_slice(
        "durable restore closure bytes",
        MAX_DELTA_BYTES.min(limits.max_package_bytes),
    )?;
    if decoder.offset != bytes.len() {
        return Err(invalid_durable("trailing Ink durable restore bytes"));
    }
    Ok((intent, delta))
}
#[derive(Clone, Debug, PartialEq, Eq)]
struct Member {
    content_type: Option<String>,
    payload: Vec<u8>,
    relationships: Option<(bool, Vec<u8>)>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Record {
    kind: u8,
    name: String,
    before: Option<Member>,
    after: Option<Member>,
}

fn encode_delta(
    before: &OpcPackage,
    after: &OpcPackage,
    maximum: usize,
) -> std::result::Result<Vec<u8>, PatchError> {
    preflight_delta(before, after, maximum)?;
    let records = collect_records(before, after)?;
    if records.len() > MAX_DELTA_RECORDS {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure record count",
        });
    }
    let mut output = Vec::new();
    if maximum < DELTA_HEADER.len().saturating_add(8) {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        });
    }
    output
        .try_reserve(DELTA_HEADER.len().saturating_add(16))
        .map_err(|_| PatchError::Allocation)?;
    output.extend_from_slice(DELTA_HEADER);
    put_u64(&mut output, records.len() as u64, maximum)?;
    for record in records {
        push_bounded(&mut output, record.kind, maximum)?;
        put_text(&mut output, &record.name, maximum)?;
        put_member(&mut output, record.kind, record.before.as_ref(), maximum)?;
        put_member(&mut output, record.kind, record.after.as_ref(), maximum)?;
    }
    Ok(output)
}

fn preflight_delta(
    before: &OpcPackage,
    after: &OpcPackage,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    let before_content = before
        .source_content_types()
        .map_err(|_| PatchError::InvalidText {
            field: "Ink durable source content types",
        })?;
    let after_content = after
        .source_content_types()
        .map_err(|_| PatchError::InvalidText {
            field: "Ink durable target content types",
        })?;
    let before_root = root_relationships(before).map_err(|_| PatchError::InvalidText {
        field: "Ink durable source package relationships",
    })?;
    let after_root = root_relationships(after).map_err(|_| PatchError::InvalidText {
        field: "Ink durable target package relationships",
    })?;
    if before_root.bytes() != after_root.bytes()
        || before_root.member_present() != after_root.member_present()
    {
        return Err(PatchError::InvalidText {
            field: "Ink durable root relationship transition",
        });
    }
    let mut names = BTreeSet::new();
    for package in [before, after] {
        for part in package.iter_parts() {
            if names.len() >= MAX_DELTA_RECORDS {
                return Err(PatchError::InvalidText {
                    field: "Ink durable closure record count",
                });
            }
            names.insert(part.partname().as_str().to_owned());
        }
    }
    let mut records = 0usize;
    let mut total = DELTA_HEADER.len().saturating_add(8);
    if before_content.bytes() != after_content.bytes() {
        records = records.saturating_add(1);
        total = add_delta_size(total, 1 + framed_size("[Content_Types].xml".len()), maximum)?;
        total = add_delta_size(
            total,
            optional_xml_member_size(before_content.bytes()),
            maximum,
        )?;
        total = add_delta_size(
            total,
            optional_xml_member_size(after_content.bytes()),
            maximum,
        )?;
    }
    if before_root.bytes() != after_root.bytes()
        || before_root.member_present() != after_root.member_present()
    {
        records = records.saturating_add(1);
        total = add_delta_size(total, 1 + framed_size(1), maximum)?;
        total = add_delta_size(
            total,
            optional_xml_member_size(before_root.bytes()),
            maximum,
        )?;
        total = add_delta_size(total, optional_xml_member_size(after_root.bytes()), maximum)?;
    }
    for name in names {
        let uri = PackURI::new(name.clone()).map_err(|_| PatchError::InvalidText {
            field: "Ink durable part name",
        })?;
        if parts_equal(before, after, &uri)? {
            continue;
        }
        records = records.saturating_add(1);
        if records > MAX_DELTA_RECORDS {
            return Err(PatchError::InvalidText {
                field: "Ink durable closure record count",
            });
        }
        total = add_delta_size(total, 1 + framed_size(name.len()), maximum)?;
        total = add_delta_size(total, part_member_size(before, &uri)?, maximum)?;
        total = add_delta_size(total, part_member_size(after, &uri)?, maximum)?;
    }
    if records == 0 {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure record count",
        });
    }
    Ok(())
}

fn add_delta_size(
    current: usize,
    additional: usize,
    maximum: usize,
) -> std::result::Result<usize, PatchError> {
    let total = current
        .checked_add(additional)
        .ok_or(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        })?;
    if total > maximum {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        });
    }
    Ok(total)
}

fn framed_size(length: usize) -> usize {
    8usize.saturating_add(length)
}

fn optional_xml_member_size(bytes: &[u8]) -> usize {
    1usize.saturating_add(framed_size(bytes.len()))
}

/// Borrow a present part whose payload the closure needs.
///
/// Presence is decided first with [`OpcPackage::part_metadata`], which never
/// decodes. Under ADR 0030 `get_part` decodes the payload on first access, so
/// an error here is a present part that failed to decode, never an absence.
fn decoded_part<'package>(
    package: &'package OpcPackage,
    name: &PackURI,
) -> std::result::Result<&'package dyn Part, PatchError> {
    package.get_part(name).map_err(|_| PatchError::InvalidText {
        field: "Ink durable part payload",
    })
}

fn part_member_size(
    package: &OpcPackage,
    name: &PackURI,
) -> std::result::Result<usize, PatchError> {
    if package.part_metadata(name).is_none() {
        return Ok(1);
    }
    let part = decoded_part(package, name)?;
    let relationships =
        package
            .source_relationships(name)
            .map_err(|_| PatchError::InvalidText {
                field: "Ink durable relationship token",
            })?;
    let mut size = 1usize;
    size = size.saturating_add(framed_size(part.content_type().len()));
    size = size.saturating_add(framed_size(part.blob().len()));
    size = size.saturating_add(1);
    size = size.saturating_add(framed_size(relationships.bytes().len()));
    Ok(size)
}

fn parts_equal(
    before: &OpcPackage,
    after: &OpcPackage,
    name: &PackURI,
) -> std::result::Result<bool, PatchError> {
    match (
        before.part_metadata(name).is_some(),
        after.part_metadata(name).is_some(),
    ) {
        (false, false) => Ok(true),
        (true, true) => {
            let left = decoded_part(before, name)?;
            let right = decoded_part(after, name)?;
            let left_relationships =
                before
                    .source_relationships(name)
                    .map_err(|_| PatchError::InvalidText {
                        field: "Ink durable relationship token",
                    })?;
            let right_relationships =
                after
                    .source_relationships(name)
                    .map_err(|_| PatchError::InvalidText {
                        field: "Ink durable relationship token",
                    })?;
            Ok(left.content_type() == right.content_type()
                && left.blob() == right.blob()
                && left_relationships.member_present() == right_relationships.member_present()
                && left_relationships.bytes() == right_relationships.bytes())
        },
        _ => Ok(false),
    }
}

fn put_member(
    output: &mut Vec<u8>,
    kind: u8,
    member: Option<&Member>,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    push_bounded(output, u8::from(member.is_some()), maximum)?;
    let Some(member) = member else { return Ok(()) };
    if kind == 2 {
        put_text(
            output,
            member
                .content_type
                .as_deref()
                .ok_or(PatchError::InvalidText {
                    field: "Ink durable part content type",
                })?,
            maximum,
        )?;
        let Some((present, relationships)) = member.relationships.as_ref() else {
            return Err(PatchError::InvalidText {
                field: "Ink durable part relationships",
            });
        };
        put_bytes(output, &member.payload, maximum)?;
        push_bounded(output, u8::from(*present), maximum)?;
        put_bytes(output, relationships, maximum)?;
    } else {
        put_bytes(output, &member.payload, maximum)?;
    }
    Ok(())
}

fn put_text(
    output: &mut Vec<u8>,
    text: &str,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    if text.len() > MAX_MEMBER_NAME || !text.is_ascii() {
        return Err(PatchError::InvalidText {
            field: "Ink durable member name",
        });
    }
    put_bytes(output, text.as_bytes(), maximum)
}

fn put_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    let total = output
        .len()
        .checked_add(8)
        .and_then(|length| length.checked_add(bytes.len()))
        .ok_or(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        })?;
    if total > maximum {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        });
    }
    put_u64(
        output,
        u64::try_from(bytes.len()).map_err(|_| PatchError::InvalidText {
            field: "Ink durable byte length",
        })?,
        maximum,
    )?;
    output
        .try_reserve(bytes.len())
        .map_err(|_| PatchError::Allocation)?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn put_u64(
    output: &mut Vec<u8>,
    value: u64,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    if output.len().saturating_add(8) > maximum {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        });
    }
    output.try_reserve(8).map_err(|_| PatchError::Allocation)?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn push_bounded(
    output: &mut Vec<u8>,
    value: u8,
    maximum: usize,
) -> std::result::Result<(), PatchError> {
    if output.len() >= maximum {
        return Err(PatchError::InvalidText {
            field: "Ink durable closure bytes",
        });
    }
    output.try_reserve(1).map_err(|_| PatchError::Allocation)?;
    output.push(value);
    Ok(())
}

fn collect_records(
    before: &OpcPackage,
    after: &OpcPackage,
) -> std::result::Result<Vec<Record>, PatchError> {
    let before_content = before
        .source_content_types()
        .map_err(|_| PatchError::InvalidText {
            field: "Ink durable source content types",
        })?;
    let after_content = after
        .source_content_types()
        .map_err(|_| PatchError::InvalidText {
            field: "Ink durable target content types",
        })?;
    let before_root = root_relationships(before).map_err(|_| PatchError::InvalidText {
        field: "Ink durable source package relationships",
    })?;
    let after_root = root_relationships(after).map_err(|_| PatchError::InvalidText {
        field: "Ink durable target package relationships",
    })?;
    let mut records = Vec::new();
    if before_content.bytes() != after_content.bytes() {
        records.try_reserve(1).map_err(|_| PatchError::Allocation)?;
        records.push(Record {
            kind: 0,
            name: "[Content_Types].xml".into(),
            before: Some(Member {
                content_type: None,
                payload: clone_durable_bytes(
                    before_content.bytes(),
                    "Ink durable source content types",
                )?,
                relationships: None,
            }),
            after: Some(Member {
                content_type: None,
                payload: clone_durable_bytes(
                    after_content.bytes(),
                    "Ink durable target content types",
                )?,
                relationships: None,
            }),
        });
    }
    if before_root.bytes() != after_root.bytes()
        || before_root.member_present() != after_root.member_present()
    {
        records.try_reserve(1).map_err(|_| PatchError::Allocation)?;
        records.push(Record {
            kind: 1,
            name: "/".into(),
            before: Some(Member {
                content_type: None,
                payload: clone_durable_bytes(
                    before_root.bytes(),
                    "Ink durable source package relationships",
                )?,
                relationships: Some((before_root.member_present(), Vec::new())),
            }),
            after: Some(Member {
                content_type: None,
                payload: clone_durable_bytes(
                    after_root.bytes(),
                    "Ink durable target package relationships",
                )?,
                relationships: Some((after_root.member_present(), Vec::new())),
            }),
        });
    }

    let mut names = BTreeSet::new();
    names.extend(
        before
            .iter_parts()
            .map(|part| part.partname().as_str().to_owned()),
    );
    names.extend(
        after
            .iter_parts()
            .map(|part| part.partname().as_str().to_owned()),
    );
    for name in names {
        let uri = PackURI::new(name.clone()).map_err(|_| PatchError::InvalidText {
            field: "Ink durable part name",
        })?;
        let left = capture_part(before, &uri).map_err(|_| PatchError::InvalidText {
            field: "Ink durable source part",
        })?;
        let right = capture_part(after, &uri).map_err(|_| PatchError::InvalidText {
            field: "Ink durable target part",
        })?;
        if left == right {
            continue;
        }
        records.try_reserve(1).map_err(|_| PatchError::Allocation)?;
        records.push(Record {
            kind: 2,
            name,
            before: left,
            after: right,
        });
    }
    Ok(records)
}

fn capture_part(
    package: &OpcPackage,
    name: &PackURI,
) -> std::result::Result<Option<Member>, PatchError> {
    if package.part_metadata(name).is_none() {
        return Ok(None);
    }
    let part = decoded_part(package, name)?;
    let relationships =
        package
            .source_relationships(name)
            .map_err(|_| PatchError::InvalidText {
                field: "Ink durable relationship token",
            })?;
    Ok(Some(Member {
        content_type: Some(clone_durable_text(
            part.content_type(),
            "Ink durable part content type",
        )?),
        payload: clone_durable_bytes(part.blob(), "Ink durable part payload")?,
        relationships: Some((
            relationships.member_present(),
            clone_durable_bytes(relationships.bytes(), "Ink durable relationship token")?,
        )),
    }))
}

fn clone_durable_bytes(
    bytes: &[u8],
    _field: &'static str,
) -> std::result::Result<Vec<u8>, PatchError> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|_| PatchError::Allocation)?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn clone_durable_text(text: &str, _field: &'static str) -> std::result::Result<String, PatchError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| PatchError::Allocation)?;
    owned.push_str(text);
    Ok(owned)
}

fn root_relationships(package: &OpcPackage) -> Result<OwnedRelationships> {
    let root = PackURI::new("/").map_err(Error::Uri)?;
    Ok(package.source_relationships(&root)?)
}

#[derive(Clone, Copy)]
struct FingerprintBounds {
    max_non_part_members: usize,
    max_non_part_metadata_bytes: usize,
}

impl Default for FingerprintBounds {
    fn default() -> Self {
        Self {
            max_non_part_members: MAX_NON_PART_MEMBERS,
            max_non_part_metadata_bytes: MAX_NON_PART_METADATA_BYTES,
        }
    }
}

impl FingerprintBounds {
    fn from_edit_limits(limits: EditLimits) -> Self {
        Self {
            max_non_part_members: limits
                .inventory
                .stories
                .max_package_parts
                .min(MAX_NON_PART_MEMBERS),
            max_non_part_metadata_bytes: limits
                .inventory
                .stories
                .max_topology_bytes
                .min(MAX_NON_PART_METADATA_BYTES),
        }
    }
}

fn package_fingerprint(package: &OpcPackage) -> std::result::Result<String, PatchError> {
    package_fingerprint_with_bounds(package, FingerprintBounds::default())
}

fn package_fingerprint_with_bounds(
    package: &OpcPackage,
    bounds: FingerprintBounds,
) -> std::result::Result<String, PatchError> {
    let non_parts = package.non_part_members();
    let mut non_part_metadata_bytes = 0usize;
    if non_parts.len() > bounds.max_non_part_members {
        return Err(PatchError::InvalidText {
            field: "Ink durable non-part member count",
        });
    }
    for member in non_parts {
        non_part_metadata_bytes = non_part_metadata_bytes
            .checked_add(member.name().len())
            .and_then(|value| value.checked_add(member.reason().as_str().len()))
            .ok_or(PatchError::InvalidText {
                field: "Ink durable non-part metadata bytes",
            })?;
        if non_part_metadata_bytes > bounds.max_non_part_metadata_bytes {
            return Err(PatchError::InvalidText {
                field: "Ink durable non-part metadata bytes",
            });
        }
    }
    let content = package
        .source_content_types()
        .map_err(|_| PatchError::InvalidText {
            field: "Ink durable content types",
        })?;
    let root = root_relationships(package).map_err(|_| PatchError::InvalidText {
        field: "Ink durable package relationships",
    })?;
    let mut digest = Sha256::new();
    hash_section(&mut digest, b"content-types", content.bytes());
    hash_section(&mut digest, b"content-types-present", &[1]);
    hash_section(&mut digest, b"root-rels", root.bytes());
    hash_section(
        &mut digest,
        b"root-rels-present",
        &[u8::from(root.member_present())],
    );
    // Every payload is hashed, so every part is decoded here (ADR 0030).
    let mut parts = BTreeMap::new();
    for part in package.try_iter_parts() {
        let part = part.map_err(|_| PatchError::InvalidText {
            field: "Ink durable part payload",
        })?;
        parts.insert(part.partname().as_str().to_owned(), part);
    }
    for (name, part) in parts {
        let relationships =
            package
                .source_relationships(part.partname())
                .map_err(|_| PatchError::InvalidText {
                    field: "Ink durable relationship token",
                })?;
        hash_section(&mut digest, b"part-name", name.as_bytes());
        hash_section(
            &mut digest,
            b"part-content-type",
            part.content_type().as_bytes(),
        );
        hash_section(&mut digest, b"part-payload", part.blob());
        hash_section(&mut digest, b"part-rels", relationships.bytes());
        hash_section(
            &mut digest,
            b"part-rels-present",
            &[u8::from(relationships.member_present())],
        );
    }
    // The public OPC model deliberately does not expose non-part bytes. Bind
    // the reader's bounded classification (name and reason) so a package with
    // changed archive-junk topology cannot pass an Ink source guard. The raw
    // ZIP bytes remain outside this logical OPC fingerprint and are preserved
    // by the owning source archive when available.
    // The bounded inventory was preflighted before any package fingerprinting.
    for member in non_parts {
        hash_section(&mut digest, b"non-part-name", member.name().as_bytes());
        hash_section(
            &mut digest,
            b"non-part-reason",
            member.reason().as_str().as_bytes(),
        );
    }
    let bytes = digest.finalize();
    let mut text = String::new();
    text.try_reserve_exact(bytes.len().saturating_mul(2))
        .map_err(|_| PatchError::Allocation)?;
    for byte in bytes {
        write!(&mut text, "{byte:02x}").map_err(|_| PatchError::Allocation)?;
    }
    Ok(text)
}

fn hash_section(digest: &mut Sha256, name: &[u8], bytes: &[u8]) {
    digest.update((name.len() as u64).to_le_bytes());
    digest.update(name);
    digest.update((bytes.len() as u64).to_le_bytes());
    digest.update(bytes);
}

struct Decoder<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Decoder<'a> {
    fn new(bytes: &'a [u8]) -> Result<Self> {
        if !bytes.starts_with(DELTA_HEADER) {
            return Err(invalid_durable("invalid Ink durable closure header"));
        }
        Ok(Self {
            bytes,
            offset: DELTA_HEADER.len(),
        })
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| invalid_durable("Ink durable closure offset overflow"))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| invalid_durable("truncated Ink durable closure"))?;
        self.offset = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8> {
        self.take(1).map(|bytes| bytes[0])
    }

    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.offset)
    }

    fn u64(&mut self, resource: &'static str, maximum: usize) -> Result<usize> {
        let bytes: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_| invalid_durable("invalid Ink durable integer"))?;
        let value = usize::try_from(u64::from_le_bytes(bytes))
            .map_err(|_| invalid_durable("Ink durable integer exceeds this platform"))?;
        if value > maximum {
            return Err(Error::InkLimit {
                resource,
                actual: value,
                maximum,
            });
        }
        Ok(value)
    }

    fn bytes(&mut self, resource: &'static str, maximum: usize) -> Result<Vec<u8>> {
        let length = self.u64(resource, maximum)?;
        let source = self.take(length)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink durable closure bytes",
                source,
            })?;
        bytes.extend_from_slice(source);
        Ok(bytes)
    }

    fn text(&mut self, resource: &'static str, maximum: usize) -> Result<String> {
        let bytes = self.bytes(resource, maximum)?;
        let text = std::str::from_utf8(&bytes)
            .map_err(|_| invalid_durable("Ink durable closure text is not UTF-8"))?;
        if text.is_empty() || !text.is_ascii() {
            return Err(invalid_durable("Ink durable closure text is not canonical"));
        }
        Ok(text.to_owned())
    }
}

fn decode_delta(
    bytes: &[u8],
    limits: EditLimits,
    story_names: &BTreeSet<String>,
) -> Result<Vec<Record>> {
    let maximum = limits.max_package_bytes.min(MAX_DELTA_BYTES);
    if bytes.len() > maximum {
        return Err(Error::InkLimit {
            resource: "durable closure bytes",
            actual: bytes.len(),
            maximum,
        });
    }
    let mut decoder = Decoder::new(bytes)?;
    let count = decoder.u64("durable closure records", MAX_DELTA_RECORDS)?;
    if count == 0 {
        return Err(invalid_durable("empty Ink durable closure"));
    }
    const MIN_CLOSURE_RECORD_BYTES: usize = 1 + 8 + 1 + 1;
    if count > decoder.remaining() / MIN_CLOSURE_RECORD_BYTES {
        return Err(invalid_durable("truncated Ink durable closure records"));
    }
    let retained_bytes = count
        .checked_mul(size_of::<Record>())
        .ok_or_else(|| invalid_durable("Ink durable closure allocation overflow"))?;
    if retained_bytes > maximum {
        return Err(Error::InkLimit {
            resource: "durable closure records",
            actual: retained_bytes,
            maximum,
        });
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink durable closure records",
            source,
        })?;
    let mut seen = BTreeSet::new();
    let mut before_metadata_bytes = 0usize;
    let mut after_metadata_bytes = 0usize;
    let mut before_relationships = 0usize;
    let mut after_relationships = 0usize;
    for _ in 0..count {
        let kind = decoder.u8()?;
        if kind > 2 {
            return Err(invalid_durable("unknown Ink durable closure member kind"));
        }
        let name = decoder.text(
            "durable member name",
            MAX_MEMBER_NAME.min(limits.inventory.stories.max_topology_bytes),
        )?;
        if kind == 0 && name != "[Content_Types].xml" {
            return Err(invalid_durable("invalid content-types closure member"));
        }
        if kind == 1 && name != "/" {
            return Err(invalid_durable(
                "invalid package relationships closure member",
            ));
        }
        if kind == 2 {
            PackURI::new(name.clone()).map_err(Error::Uri)?;
        }
        let before = decode_member(&mut decoder, kind, &name, limits, story_names)?;
        let after = decode_member(&mut decoder, kind, &name, limits, story_names)?;
        if before.is_none() && after.is_none() {
            return Err(invalid_durable("empty Ink durable closure record"));
        }
        if before == after {
            return Err(invalid_durable("equal Ink durable closure record"));
        }
        if kind == 2
            && let (Some(before), Some(after)) = (&before, &after)
            && before.content_type != after.content_type
        {
            return Err(invalid_durable(
                "existing Ink durable part content type changed",
            ));
        }
        before_metadata_bytes = add_closure_metadata_size(
            before_metadata_bytes,
            name.len(),
            limits.inventory.stories.max_topology_bytes,
        )?;
        after_metadata_bytes = add_closure_metadata_size(
            after_metadata_bytes,
            name.len(),
            limits.inventory.stories.max_topology_bytes,
        )?;
        before_metadata_bytes = add_closure_metadata_bytes(
            before_metadata_bytes,
            kind,
            before.as_ref(),
            limits.inventory.stories.max_topology_bytes,
        )?;
        after_metadata_bytes = add_closure_metadata_bytes(
            after_metadata_bytes,
            kind,
            after.as_ref(),
            limits.inventory.stories.max_topology_bytes,
        )?;
        before_relationships = add_closure_relationship_count(
            before_relationships,
            before.as_ref(),
            limits.inventory,
        )?;
        after_relationships =
            add_closure_relationship_count(after_relationships, after.as_ref(), limits.inventory)?;
        if !seen.insert((kind, name.clone())) {
            return Err(invalid_durable("duplicate Ink durable closure member"));
        }
        records.push(Record {
            kind,
            name,
            before,
            after,
        });
    }
    if decoder.offset != bytes.len() {
        return Err(invalid_durable("trailing Ink durable closure bytes"));
    }
    Ok(records)
}

fn add_closure_metadata_bytes(
    current: usize,
    kind: u8,
    member: Option<&Member>,
    maximum: usize,
) -> Result<usize> {
    let Some(member) = member else {
        return Ok(current);
    };
    let additional = if kind == 2 {
        member
            .content_type
            .as_ref()
            .map_or(0, String::len)
            .checked_add(
                member
                    .relationships
                    .as_ref()
                    .map_or(0, |(_, bytes)| bytes.len()),
            )
            .ok_or_else(|| invalid_durable("Ink durable closure metadata overflow"))?
    } else {
        member.payload.len()
    };
    let total = current
        .checked_add(additional)
        .ok_or_else(|| invalid_durable("Ink durable closure metadata overflow"))?;
    if total > maximum {
        return Err(Error::InkLimit {
            resource: "durable closure metadata bytes",
            actual: total,
            maximum,
        });
    }
    Ok(total)
}

fn add_closure_metadata_size(current: usize, additional: usize, maximum: usize) -> Result<usize> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| invalid_durable("Ink durable closure metadata overflow"))?;
    if total > maximum {
        return Err(Error::InkLimit {
            resource: "durable closure metadata bytes",
            actual: total,
            maximum,
        });
    }
    Ok(total)
}

fn add_closure_relationship_count(
    current: usize,
    member: Option<&Member>,
    limits: Limits,
) -> Result<usize> {
    let Some(member) = member else {
        return Ok(current);
    };
    let Some((_, bytes)) = member.relationships.as_ref() else {
        return Ok(current);
    };
    let count = parse_relationships(bytes, limits)?.len();
    let total = current
        .checked_add(count)
        .ok_or_else(|| invalid_durable("Ink durable relationship count overflow"))?;
    if total > limits.max_relationships {
        return Err(Error::InkLimit {
            resource: "durable closure relationships",
            actual: total,
            maximum: limits.max_relationships,
        });
    }
    Ok(total)
}

fn decode_member(
    decoder: &mut Decoder<'_>,
    kind: u8,
    name: &str,
    limits: EditLimits,
    story_names: &BTreeSet<String>,
) -> Result<Option<Member>> {
    let present = decoder.u8()?;
    if present > 1 {
        return Err(invalid_durable("invalid Ink durable closure presence"));
    }
    if present == 0 {
        return Ok(None);
    }
    if kind == 2 {
        let content_type = decoder.text(
            "durable content type",
            MAX_CONTENT_TYPE.min(limits.inventory.stories.max_topology_bytes),
        )?;
        let payload_maximum = if story_names.contains(name) || is_story_content_type(&content_type)
        {
            limits.inventory.stories.max_story_bytes
        } else {
            limits.inventory.max_payload_bytes
        };
        let payload = decoder.bytes("durable part payload", payload_maximum)?;
        let relationship_present = decoder.u8()?;
        if relationship_present > 1 {
            return Err(invalid_durable("invalid Ink durable relationship presence"));
        }
        let relationships = decoder.bytes(
            "durable relationship bytes",
            MAX_RELATIONSHIPS_BYTES
                .min(limits.inventory.stories.max_topology_bytes)
                .min(limits.max_package_bytes),
        )?;
        Ok(Some(Member {
            content_type: Some(content_type),
            payload,
            relationships: Some((relationship_present == 1, relationships)),
        }))
    } else {
        Ok(Some(Member {
            content_type: None,
            payload: decoder.bytes(
                "durable package XML",
                (MAX_RELATIONSHIPS_BYTES * 16)
                    .min(limits.inventory.stories.max_topology_bytes)
                    .min(limits.max_package_bytes),
            )?,
            relationships: None,
        }))
    }
}

fn is_story_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
            | ct::WML_HEADER
            | ct::WML_FOOTER
            | ct::WML_FOOTNOTES
            | ct::WML_ENDNOTES
            | ct::WML_COMMENTS
            | ct::WML_DOCUMENT_GLOSSARY
    )
}

fn validate_delta_source(package: &OpcPackage, records: &[Record]) -> Result<()> {
    let mut seen = BTreeSet::new();
    for record in records {
        if !seen.insert((record.kind, record.name.as_str())) {
            return Err(invalid_durable("duplicate Ink durable closure member"));
        }
        match record.kind {
            0 => {
                let current = package.source_content_types()?;
                let expected = record
                    .before
                    .as_ref()
                    .ok_or_else(|| invalid_durable("missing source content-types token"))?;
                if current.bytes() != expected.payload.as_slice() {
                    return Err(invalid_durable("stale Ink content-types source"));
                }
            },
            1 => {
                let current = root_relationships(package)?;
                let expected = record
                    .before
                    .as_ref()
                    .ok_or_else(|| invalid_durable("missing source package relationships token"))?;
                if current.bytes() != expected.payload.as_slice()
                    || current.member_present()
                        != expected
                            .relationships
                            .as_ref()
                            .is_some_and(|(present, _)| *present)
                {
                    return Err(invalid_durable("stale Ink package relationship source"));
                }
            },
            2 => {
                if let (Some(before), Some(after)) = (&record.before, &record.after)
                    && before.content_type != after.content_type
                {
                    return Err(invalid_durable(
                        "existing Ink durable part content type changed",
                    ));
                }
                let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
                // Presence comes from metadata alone. A present part is
                // decoded only to compare its payload, and a decode failure
                // is reported as such, never as an absent part (ADR 0030).
                match (&record.before, package.part_metadata(&name).is_some()) {
                    (None, true) => {
                        return Err(invalid_durable("unexpected Ink durable source part"));
                    },
                    (Some(_), false) => {
                        return Err(invalid_durable("missing Ink durable source part"));
                    },
                    (None, false) => {},
                    (Some(expected), true) => {
                        let part = package.get_part(&name)?;
                        let Some((present, relationships)) = expected.relationships.as_ref() else {
                            return Err(invalid_durable("missing source part relationship token"));
                        };
                        let current = package.source_relationships(&name)?;
                        if expected.content_type.as_deref() != Some(part.content_type())
                            || expected.payload.as_slice() != part.blob()
                            || *present != current.member_present()
                            || relationships.as_slice() != current.bytes()
                        {
                            return Err(invalid_durable("stale Ink durable source part"));
                        }
                    },
                }
            },
            _ => return Err(invalid_durable("unknown Ink durable closure member kind")),
        }
    }
    Ok(())
}

fn apply_delta(
    package: &mut OpcPackage,
    records: &[Record],
    limits: Limits,
    proof_maximum: usize,
) -> Result<()> {
    if records.iter().any(|record| record.kind == 1) {
        return Err(invalid_durable(
            "root relationship transitions are not supported by Ink durable replay",
        ));
    }
    let content_record = records.iter().find(|record| record.kind == 0);
    let target_content = if let Some(record) = content_record {
        Some(
            record
                .after
                .as_ref()
                .ok_or_else(|| invalid_durable("missing target content-types token"))?,
        )
    } else {
        None
    };
    let target_content_bytes = target_content.map(|member| member.payload.as_slice());
    let source_content = package.source_content_types()?;

    let mut part_records = Vec::new();
    for record in records.iter().filter(|record| record.kind == 2) {
        part_records.push(record);
    }

    // Existing payload changes are limited to story XML.  Ink and image
    // resources are added/removed as closure members by the graph transition.
    let stories = story::capture(package, limits.stories)?;
    let story_names: BTreeSet<_> = stories
        .stories()
        .iter()
        .map(|story| story.part().as_str().to_owned())
        .collect();
    for record in &part_records {
        if let (Some(before), Some(after)) = (&record.before, &record.after)
            && before.payload != after.payload
            && (!story_names.contains(&record.name) || before.content_type != after.content_type)
        {
            return Err(invalid_durable(
                "unsupported non-story Ink part replacement",
            ));
        }
    }

    let target_content_for_proof = target_content_bytes.unwrap_or(source_content.bytes());
    let proof = exact_transition_proof_package(
        package,
        records,
        target_content_for_proof,
        &story_names,
        limits,
        proof_maximum,
    )?;
    let relationship_targets = relationship_tokens(package, records, &proof, limits)?;
    // Add targets owner by owner.  The intermediate manifest includes exactly
    // the targets in that owner batch, matching AddDelta's source-preserving
    // transaction and avoiding a manifest that names not-yet-added parts.
    let mut additions_by_owner: BTreeMap<String, Vec<&Record>> = BTreeMap::new();
    for record in &part_records {
        if record.before.is_some() || record.after.is_none() {
            continue;
        }
        let after = record.after.as_ref().expect("checked above");
        let Some((_present, bytes)) = after.relationships.as_ref() else {
            return Err(invalid_durable(
                "new Ink resource has no relationship presence token",
            ));
        };
        if !parse_relationships(bytes, limits)?.is_empty() {
            return Err(invalid_durable(
                "new Ink resource has outgoing relationships",
            ));
        }
        let owner = find_new_part_owner(package, records, &record.name, limits)?;
        additions_by_owner.entry(owner).or_default().push(record);
    }
    let mut current_content = source_content;
    let mut processed_owners = BTreeSet::new();
    for (owner_name, additions) in &additions_by_owner {
        let owner = PackURI::new(owner_name.clone()).map_err(Error::Uri)?;
        let current_owner = package.source_relationships(&owner)?;
        let replacement_owner = relationship_targets
            .get(owner_name)
            .ok_or_else(|| invalid_durable("missing Ink owner relationship target"))?;
        let mut specs: Vec<(PackURI, String)> = Vec::new();
        for record in additions {
            let part = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
            let after = record.after.as_ref().expect("checked above");
            let content_type = after
                .content_type
                .clone()
                .ok_or_else(|| invalid_durable("new Ink resource has no content type"))?;
            specs.push((part, content_type));
        }
        let mut override_specs = Vec::new();
        for (part, content_type) in &specs {
            override_specs.push((part, content_type.as_str()));
        }
        let next_content = current_content
            .with_part_overrides(&override_specs, limits.stories.max_topology_bytes)?;
        let mut parts = Vec::new();
        for record in additions {
            let after = record.after.as_ref().expect("checked above");
            let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
            let content_type = after
                .content_type
                .clone()
                .ok_or_else(|| invalid_durable("new Ink resource has no content type"))?;
            let part: Box<dyn Part + Send + Sync> = Box::new(BlobPart::new_shared(
                name,
                content_type,
                Arc::new(after.payload.clone()),
            ));
            parts.push(part);
        }
        package.try_add_parts_with_source_tokens(
            current_content.bytes(),
            &next_content,
            &current_owner,
            replacement_owner,
            parts,
        )?;
        for record in additions {
            let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
            let target_relationships = proof.source_relationships(&name)?;
            let current_relationships = package.source_relationships(&name)?;
            if current_relationships != target_relationships {
                package.try_replace_relationships(&current_relationships, &target_relationships)?;
            }
        }
        current_content = next_content;
        processed_owners.insert(owner_name.clone());
    }

    // Owners with only removals still need their exact relationship token.
    for (owner_name, replacement) in &relationship_targets {
        if processed_owners.contains(owner_name) {
            continue;
        }
        let owner = PackURI::new(owner_name.clone()).map_err(Error::Uri)?;
        let current = package.source_relationships(&owner)?;
        if current != *replacement {
            package.try_replace_relationships(&current, replacement)?;
        }
    }

    // Story XML is replaced only after graph ownership has been changed.
    // Binary closure parts are removed after all incoming owner edges change.
    for record in &part_records {
        let (Some(before), Some(after)) = (&record.before, &record.after) else {
            continue;
        };
        if before.payload == after.payload {
            continue;
        }
        let name = PackURI::new(record.name.clone()).map_err(Error::Uri)?;
        {
            let part = package.get_part(&name)?;
            if !story_names.contains(&record.name)
                || part.content_type() != after.content_type.as_deref().unwrap_or_default()
            {
                return Err(invalid_durable("unsupported Ink payload replacement"));
            }
            if part.blob() != before.payload.as_slice() {
                return Err(invalid_durable("Ink story source changed during replay"));
            }
        }
        let replacement = proof.source_xml_part(&name)?;
        package.try_replace_owned_xml_part(before.payload.as_slice(), replacement)?;
    }
    let mut removed_names = Vec::new();
    for record in &part_records {
        if record.before.is_some() && record.after.is_none() {
            removed_names.push(PackURI::new(record.name.clone()).map_err(Error::Uri)?);
        }
    }
    for name in &removed_names {
        if has_incoming(package, name)? {
            return Err(invalid_durable(
                "Ink durable removal leaves a foreign incoming edge",
            ));
        }
        if !package.remove_part(name) {
            return Err(invalid_durable("Ink durable removal target disappeared"));
        }
    }

    if let Some(target) = target_content_bytes {
        let replacement = proof.source_content_types()?;
        if replacement.bytes() != target {
            return Err(invalid_durable(
                "Ink durable proof content types changed during OPC parsing",
            ));
        }
        let current_token = package.source_content_types()?;
        package.try_replace_content_types(current_token.bytes(), &replacement)?;
        if package.source_content_types()?.bytes() != target {
            return Err(invalid_durable(
                "content-types transition is not source-preserving",
            ));
        }
    }
    Ok(())
}

/// Build a bounded OPC proof for the exact target metadata and graph closure.
/// The proof contains empty placeholders for untouched payloads, exact target
/// relationship members, and actual bytes only for changed story XML. Parsing
/// it through OPC creates source-bound content/relationship/XML tokens without
/// manufacturing trusted tokens from arbitrary durable bytes.
fn exact_transition_proof_package(
    package: &OpcPackage,
    records: &[Record],
    target: &[u8],
    story_names: &BTreeSet<String>,
    limits: Limits,
    proof_maximum: usize,
) -> Result<OpcPackage> {
    let maximum = limits.stories.max_topology_bytes;
    if proof_maximum == 0 {
        return Err(invalid_durable("Ink durable proof staging limit is zero"));
    }
    if target.len() > maximum {
        return Err(invalid_durable(
            "Ink durable target content types exceed limits",
        ));
    }
    let mut changed = BTreeMap::new();
    for record in records.iter().filter(|record| record.kind == 2) {
        if changed.insert(record.name.as_str(), record).is_some() {
            return Err(invalid_durable("duplicate Ink durable target part record"));
        }
    }
    let mut names = BTreeSet::new();
    names.extend(
        package
            .iter_parts()
            .map(|part| part.partname().as_str().to_owned()),
    );
    for record in records.iter().filter(|record| record.kind == 2) {
        if record.after.is_some() {
            names.insert(record.name.clone());
        } else {
            names.remove(&record.name);
        }
    }
    if names.len() > limits.stories.max_package_parts {
        return Err(invalid_durable(
            "Ink durable proof part count exceeds limits",
        ));
    }

    let mut total = 0;
    let mut topology = 0;
    proof_add_metadata(
        &mut total,
        &mut topology,
        target.len().saturating_add(PROOF_ENTRY_OVERHEAD),
        maximum,
        proof_maximum,
    )?;
    let root = PackURI::new("/").map_err(Error::Uri)?;
    let root_relationships = package.source_relationships(&root)?;
    if root_relationships.member_present() {
        proof_add_metadata(
            &mut total,
            &mut topology,
            root.rels_uri()
                .map_err(Error::Uri)?
                .as_str()
                .len()
                .saturating_add(PROOF_ENTRY_OVERHEAD),
            maximum,
            proof_maximum,
        )?;
        proof_add_metadata(
            &mut total,
            &mut topology,
            root_relationships.bytes().len(),
            maximum,
            proof_maximum,
        )?;
    }

    for name in &names {
        let record = changed.get(name.as_str()).copied();
        let proof_part = final_proof_part(package, record, name, story_names)?;
        proof_add_metadata(
            &mut total,
            &mut topology,
            name.len()
                .saturating_add(proof_part.content_type.len())
                .saturating_add(PROOF_ENTRY_OVERHEAD),
            maximum,
            proof_maximum,
        )?;
        if proof_part.payload.len() > limits.stories.max_story_bytes {
            return Err(invalid_durable(
                "Ink durable proof story payload exceeds limits",
            ));
        }
        total = proof_add_size(total, proof_part.payload.len(), proof_maximum)?;
        let (present, bytes) = proof_part.relationships.view();
        if present {
            let part = PackURI::new(name.clone()).map_err(Error::Uri)?;
            proof_add_metadata(
                &mut total,
                &mut topology,
                part.rels_uri()
                    .map_err(Error::Uri)?
                    .as_str()
                    .len()
                    .saturating_add(PROOF_ENTRY_OVERHEAD),
                maximum,
                proof_maximum,
            )?;
            proof_add_metadata(
                &mut total,
                &mut topology,
                bytes.len(),
                maximum,
                proof_maximum,
            )?;
        }
    }

    let content_types = PackURI::new("/[Content_Types].xml").map_err(Error::Uri)?;
    let mut writer = PhysPkgWriter::new();
    writer.write_stored(&content_types, target)?;
    if root_relationships.member_present() {
        let name = root.rels_uri().map_err(Error::Uri)?;
        writer.write_stored(&name, root_relationships.bytes())?;
    }
    for name in &names {
        let record = changed.get(name.as_str()).copied();
        let proof_part = final_proof_part(package, record, name, story_names)?;
        let part = PackURI::new(name.clone()).map_err(Error::Uri)?;
        writer.write_stored(&part, proof_part.payload)?;
    }
    for name in &names {
        let record = changed.get(name.as_str()).copied();
        let proof_part = final_proof_part(package, record, name, story_names)?;
        let (present, bytes) = proof_part.relationships.view();
        if present {
            let part = PackURI::new(name.clone()).map_err(Error::Uri)?;
            let name = part.rels_uri().map_err(Error::Uri)?;
            writer.write_stored(&name, bytes)?;
        }
    }
    let proof = writer.finish()?;
    Ok(OpcPackage::from_bytes(&proof)?)
}

enum ProofRelationships<'a> {
    Current(OwnedRelationships),
    Closure { present: bool, bytes: &'a [u8] },
}

impl ProofRelationships<'_> {
    fn view(&self) -> (bool, &[u8]) {
        match self {
            Self::Current(relationships) => (relationships.member_present(), relationships.bytes()),
            Self::Closure { present, bytes } => (*present, bytes),
        }
    }
}

struct ProofPart<'a> {
    content_type: &'a str,
    payload: &'a [u8],
    relationships: ProofRelationships<'a>,
}

fn final_proof_part<'a>(
    package: &'a OpcPackage,
    record: Option<&'a Record>,
    name: &str,
    story_names: &BTreeSet<String>,
) -> Result<ProofPart<'a>> {
    if let Some(record) = record {
        let after = record
            .after
            .as_ref()
            .ok_or_else(|| invalid_durable("missing final Ink durable part"))?;
        let content_type = after
            .content_type
            .as_deref()
            .ok_or_else(|| invalid_durable("final Ink durable part has no content type"))?;
        let (present, bytes) = after
            .relationships
            .as_ref()
            .ok_or_else(|| invalid_durable("final Ink durable part has no relationships"))?;
        let payload = if story_names.contains(name)
            && record
                .before
                .as_ref()
                .is_some_and(|before| before.payload != after.payload)
        {
            after.payload.as_slice()
        } else {
            &[]
        };
        return Ok(ProofPart {
            content_type,
            payload,
            relationships: ProofRelationships::Closure {
                present: *present,
                bytes,
            },
        });
    }
    let name = PackURI::new(name.to_owned()).map_err(Error::Uri)?;
    let part = package.get_part(&name)?;
    let relationships = package.source_relationships(&name)?;
    Ok(ProofPart {
        content_type: part.content_type(),
        payload: &[],
        relationships: ProofRelationships::Current(relationships),
    })
}

fn proof_add_size(current: usize, additional: usize, maximum: usize) -> Result<usize> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| invalid_durable("Ink durable proof size overflow"))?;
    if total > maximum {
        return Err(invalid_durable("Ink durable proof metadata exceeds limits"));
    }
    Ok(total)
}

fn proof_add_metadata(
    total: &mut usize,
    topology: &mut usize,
    additional: usize,
    topology_maximum: usize,
    proof_maximum: usize,
) -> Result<()> {
    *topology = proof_add_size(*topology, additional, topology_maximum)?;
    *total = proof_add_size(*total, additional, proof_maximum)?;
    Ok(())
}

fn has_incoming(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    for owner in package.iter_parts() {
        for relationship in owner.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn find_new_part_owner(
    package: &OpcPackage,
    records: &[Record],
    target_name: &str,
    limits: Limits,
) -> Result<String> {
    let target = PackURI::new(target_name.to_owned()).map_err(Error::Uri)?;
    let mut owners = Vec::new();
    let root = PackURI::new("/").map_err(Error::Uri)?;
    owners.push(root);
    owners.extend(package.iter_parts().map(|part| part.partname().clone()));
    for owner in owners {
        let owner_record = records
            .iter()
            .find(|record| record.kind == 2 && record.name == owner.as_str());
        let Some(record) = owner_record else { continue };
        let Some(after) = record.after.as_ref() else {
            continue;
        };
        let Some((_, bytes)) = after.relationships.as_ref() else {
            continue;
        };
        if parse_relationships(bytes, limits)?
            .iter()
            .any(|relationship| {
                !relationship.external
                    && resolve_relationship_target(&owner, &relationship.target)
                        .is_ok_and(|candidate| candidate.is_equivalent_to(&target))
            })
        {
            return Ok(owner.as_str().to_owned());
        }
    }
    // Story owners always exist in the source package.  A target may be added
    // by an owner that has no payload change; its record is still present.
    Err(invalid_durable("new Ink target has no owning relationship"))
}

fn resolve_relationship_target(
    owner: &PackURI,
    target: &str,
) -> std::result::Result<PackURI, String> {
    if target.starts_with('/') {
        PackURI::new(target.to_owned())
    } else {
        PackURI::from_rel_ref(owner.base_uri(), target)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipSpec {
    id: String,
    reltype: String,
    target: String,
    external: bool,
}

fn relationship_tokens(
    package: &OpcPackage,
    records: &[Record],
    proof: &OpcPackage,
    limits: Limits,
) -> Result<BTreeMap<String, OwnedRelationships>> {
    let mut result = BTreeMap::new();
    let mut owners = Vec::new();
    owners.push(PackURI::new("/").map_err(Error::Uri)?);
    owners.extend(package.iter_parts().map(|part| part.partname().clone()));
    for owner in owners {
        let record = records
            .iter()
            .find(|record| record.kind == 2 && record.name == owner.as_str());
        let Some(record) = record else { continue };
        let before = record
            .before
            .as_ref()
            .ok_or_else(|| invalid_durable("relationship owner cannot be a newly added part"))?;
        let Some(after) = record.after.as_ref() else {
            let Some((_, before_bytes)) = before.relationships.as_ref() else {
                continue;
            };
            if parse_relationships(before_bytes, limits)?.is_empty() {
                continue;
            }
            return Err(invalid_durable(
                "relationship owner cannot be removed with relationships",
            ));
        };
        let Some((before_present, before_bytes)) = before.relationships.as_ref() else {
            return Err(invalid_durable("missing source owner relationships"));
        };
        let Some((after_present, after_bytes)) = after.relationships.as_ref() else {
            return Err(invalid_durable("missing target owner relationships"));
        };
        let current = package.source_relationships(&owner)?;
        if current.member_present() != *before_present || current.bytes() != before_bytes {
            return Err(invalid_durable("relationship owner source is stale"));
        }
        let before_specs = parse_relationships(before_bytes, limits)?;
        let after_specs = parse_relationships(after_bytes, limits)?;
        let before_map: BTreeMap<_, _> = before_specs
            .iter()
            .map(|relationship| (relationship.id.as_str(), relationship))
            .collect();
        let after_map: BTreeMap<_, _> = after_specs
            .iter()
            .map(|relationship| (relationship.id.as_str(), relationship))
            .collect();
        for (id, source) in &before_map {
            if let Some(target) = after_map.get(id)
                && *source != *target
            {
                return Err(invalid_durable("relationship retargeting is not admitted"));
            }
        }
        let token = proof.source_relationships(&owner)?;
        if token.bytes() != after_bytes || token.member_present() != *after_present {
            return Err(invalid_durable(
                "relationship proof token does not match target",
            ));
        }
        result.insert(owner.as_str().to_owned(), token);
    }
    Ok(result)
}

fn parse_relationships(bytes: &[u8], limits: Limits) -> Result<Vec<RelationshipSpec>> {
    if bytes.len() > limits.stories.max_topology_bytes {
        return Err(Error::InkLimit {
            resource: "Ink relationship metadata bytes",
            actual: bytes.len(),
            maximum: limits.stories.max_topology_bytes,
        });
    }
    let maximum = limits
        .max_relationships
        .min(limits.stories.max_relationships_per_owner);
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut relationships = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut reltype = None;
                let mut target = None;
                let mut external = false;
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())?
                        .into_owned();
                    match attribute.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => reltype = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => {
                            external = value == "External";
                            if !external && value != "Internal" {
                                return Err(invalid_durable("invalid relationship target mode"));
                            }
                        },
                        _ => return Err(invalid_durable("unknown relationship attribute")),
                    }
                }
                if relationships.len() >= maximum {
                    return Err(invalid_durable(
                        "Ink durable relationship count exceeds limits",
                    ));
                }
                relationships
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "DOCX Ink durable relationships",
                        source,
                    })?;
                relationships.push(RelationshipSpec {
                    id: id.ok_or_else(|| invalid_durable("relationship is missing Id"))?,
                    reltype: reltype
                        .ok_or_else(|| invalid_durable("relationship is missing Type"))?,
                    target: target
                        .ok_or_else(|| invalid_durable("relationship is missing Target"))?,
                    external,
                });
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(relationships)
}

#[cfg(test)]
mod tests {
    use base64::{Engine as _, engine::general_purpose::STANDARD};
    use litchi_core::patch::{BlobLimits, Patch as CorePatch, PatchLimits, Reversible};
    use serde_json::Value;

    use super::*;
    use crate::ink::{BaseProfile, Destination, Draft, Style};

    fn limits() -> PatchLimits {
        PatchLimits::new(
            BlobLimits::new(8, 64 * 1024 * 1024, 128 * 1024 * 1024),
            32 * 1024 * 1024,
            8,
            32,
            1024 * 1024,
            64 * 1024 * 1024,
        )
    }

    fn tamper_story(package: &OpcPackage) -> OpcPackage {
        let mut tampered = package.clone();
        let name = PackURI::new("/word/document.xml").expect("document URI");
        let source = tampered.get_part(&name).expect("document").blob();
        let marker = b"<w:body";
        let offset = source
            .windows(marker.len())
            .position(|window| window == marker)
            .expect("body marker")
            + marker.len();
        let mut bytes = Vec::with_capacity(source.len() + 22);
        bytes.extend_from_slice(&source[..offset]);
        bytes.extend_from_slice(b" data-durable-forged=\"1\"");
        bytes.extend_from_slice(&source[offset..]);
        tampered
            .get_part_mut(&name)
            .expect("document mutable")
            .set_blob(bytes);
        tampered
    }

    fn make_patch() -> (Patch, CorePatch<Reversible>) {
        let mut package = Package::new().expect("new DOCX");
        package.document_mut().expect("document").add_paragraph();
        let mut bytes = std::io::Cursor::new(Vec::new());
        package
            .to_plain_stream(&mut bytes)
            .expect("materialize DOCX");
        let package =
            Package::from_reader(std::io::Cursor::new(bytes.into_inner())).expect("reopen DOCX");
        let mut edit = package.edit_ink().expect("Ink edit");
        edit.insert(
            Destination::main(Position::new(0)),
            Draft::default().finish().expect("prepared Ink"),
            Style::base(BaseProfile::InkContent),
        )
        .expect("insert");
        let commit = edit.commit().expect("commit");
        let durable = commit.patch().to_durable(limits()).expect("durable");
        (commit.patch().clone(), durable)
    }

    #[test]
    fn inverse_restore_replays_intent_before_publishing_forged_story() {
        let (patch, durable) = make_patch();
        let mut target = Package::from_opc_package(patch.after.package.clone()).expect("target");
        let forged_source = tamper_story(&patch.before.package);
        let intent_id = durable.operations()[0].preconditions["intent_sha256"]
            .as_str()
            .expect("intent hash");
        let intent = single_blob_by_hex(&durable, intent_id)
            .expect("intent blob")
            .to_vec();
        let delta = encode_delta(&patch.after.package, &forged_source, MAX_DELTA_BYTES)
            .expect("forged delta");
        let restore = encode_restore(&intent, &delta, MAX_DELTA_BYTES).expect("restore");
        let forged_source_hash = package_fingerprint(&forged_source).expect("source hash");

        let inverse = durable.inverse();
        let mut wire: Value =
            serde_json::from_slice(&inverse.to_deterministic_json().expect("inverse JSON"))
                .expect("wire value");
        let id = BlobId::of(&restore).as_hex();
        let blob = wire["forward_blobs"][0]
            .as_object_mut()
            .expect("restore blob");
        blob.insert("bytes".into(), Value::String(STANDARD.encode(&restore)));
        blob.insert("sha256".into(), Value::String(id));
        let operation = wire["operations"][0]["forward"]["preconditions"]
            .as_object_mut()
            .expect("restore preconditions");
        operation.insert("target_sha256".into(), Value::String(forged_source_hash));
        operation.insert(
            "restore_sha256".into(),
            Value::String(BlobId::of(&restore).as_hex()),
        );
        let forged = CorePatch::<Reversible>::from_deterministic_json(
            &serde_json::to_vec(&wire).expect("canonical JSON"),
            limits(),
        )
        .expect("forged patch parses");
        let before = package_fingerprint(target.opc_package()).expect("target fingerprint");
        let error = target
            .apply_durable_ink_patch(&forged)
            .expect_err("forged source must fail forward replay");
        assert!(
            error
                .to_string()
                .contains("Ink durable restore forward replay mismatch")
        );
        assert_eq!(
            package_fingerprint(target.opc_package()).expect("unchanged fingerprint"),
            before
        );
    }

    #[test]
    fn forward_replay_rejects_a_forged_unrelated_target_hash() {
        let (patch, durable) = make_patch();
        let forged_target = tamper_story(&patch.after.package);
        let forged_target_hash = package_fingerprint(&forged_target).expect("target hash");
        let mut wire: Value =
            serde_json::from_slice(&durable.to_deterministic_json().expect("durable JSON"))
                .expect("wire value");
        wire["operations"][0]["forward"]["preconditions"]["target_sha256"] =
            Value::String(forged_target_hash);
        let forged = CorePatch::<Reversible>::from_deterministic_json(
            &serde_json::to_vec(&wire).expect("canonical JSON"),
            limits(),
        )
        .expect("forged patch parses");
        let mut source = Package::from_opc_package(patch.before.package.clone()).expect("source");
        let before = package_fingerprint(source.opc_package()).expect("source fingerprint");
        assert!(source.apply_durable_ink_patch(&forged).is_err());
        assert_eq!(
            package_fingerprint(source.opc_package()).expect("unchanged"),
            before
        );
    }

    #[test]
    fn inverse_restore_rebuilds_source_content_types_from_exact_proof() {
        let (initial, _) = make_patch();
        let source = Package::from_opc_package(initial.after.package.clone()).expect("source");
        let mut edit = source.edit_ink().expect("Ink edit");
        edit.replace(
            Position::new(0),
            Draft::default().finish().expect("prepared Ink"),
            None,
        )
        .expect("replace");
        let commit = edit.commit().expect("commit");
        let durable = commit.patch().to_durable(limits()).expect("durable");
        let mut target =
            Package::from_opc_package(commit.patch().after.package.clone()).expect("target");

        target
            .apply_durable_ink_patch(&durable.inverse())
            .expect("inverse replay");
        assert_eq!(
            package_fingerprint(target.opc_package()).expect("restored fingerprint"),
            package_fingerprint(&commit.patch().before.package).expect("source fingerprint"),
        );
    }

    #[test]
    fn durable_decoders_reject_impossible_counts_before_reserving_records() {
        let mut intent = INTENT_HEADER.to_vec();
        intent.extend_from_slice(&(EditLimits::default().max_operations as u64).to_le_bytes());
        let intent_error = match decode_intents(&intent, EditLimits::default()) {
            Err(error) => error,
            Ok(_) => panic!("truncated intent count must be rejected"),
        };
        assert!(
            intent_error
                .to_string()
                .contains("truncated Ink durable intent records")
        );

        let mut closure = DELTA_HEADER.to_vec();
        closure.extend_from_slice(&(MAX_DELTA_RECORDS as u64).to_le_bytes());
        let closure_error = match decode_delta(&closure, EditLimits::default(), &BTreeSet::new()) {
            Err(error) => error,
            Ok(_) => panic!("truncated closure count must be rejected"),
        };
        assert!(
            closure_error
                .to_string()
                .contains("truncated Ink durable closure records")
        );
    }

    /// A lazily opened package with one changed XML part and, optionally, a
    /// stored member whose payload fails its CRC on first decode. Deferred
    /// open accepts it; only a read of that member's payload fails (ADR 0030).
    fn lazy_package(good: &[u8], corrupt_member: bool) -> OpcPackage {
        const MANIFEST: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#;
        const ROOT_RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        const CORRUPT_PAYLOAD: &[u8] = b"ink-durable-corrupt-payload";
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored("[Content_Types].xml", MANIFEST)
            .expect("manifest");
        writer
            .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
            .expect("root relationships");
        writer.write_stored("custom/good.xml", good).expect("good");
        if corrupt_member {
            writer
                .write_stored("custom/bad.bin", CORRUPT_PAYLOAD)
                .expect("corrupt member");
        }
        let mut bytes = writer.finish_to_bytes().expect("archive");
        if corrupt_member {
            let offset = bytes
                .windows(CORRUPT_PAYLOAD.len())
                .position(|window| window == CORRUPT_PAYLOAD)
                .expect("stored payload");
            bytes[offset] ^= 1;
        }
        OpcPackage::from_vec(bytes).expect("corruption is deferred to first decode")
    }

    fn is_undecodable_part(result: std::result::Result<impl Sized, PatchError>) -> bool {
        matches!(
            result,
            Err(PatchError::InvalidText {
                field: "Ink durable part payload"
            })
        )
    }

    // Review of the 0759 merge: the closure probes read any `get_part` error
    // as absence, so a present part whose payload failed to decode was
    // silently left out of the reverse closure.
    #[test]
    fn closure_probes_refuse_a_present_part_that_fails_to_decode() {
        let before = lazy_package(b"<a/>", true);
        let after = lazy_package(b"<b/>", false);
        let bad = PackURI::new("/custom/bad.bin").expect("part name");

        // Absence is decided from metadata alone and decodes nothing.
        assert_eq!(part_member_size(&after, &bad).expect("absent"), 1);
        assert!(capture_part(&after, &bad).expect("absent").is_none());
        assert!(!parts_equal(&before, &after, &bad).expect("presence differs"));
        assert!(!parts_equal(&after, &before, &bad).expect("presence differs"));
        assert_eq!(before.deferred_decode_counters(), Some((0, 0)));
        assert_eq!(after.deferred_decode_counters(), Some((0, 0)));

        // A present part whose payload is needed reports its decode failure.
        assert!(is_undecodable_part(part_member_size(&before, &bad)));
        assert!(is_undecodable_part(capture_part(&before, &bad)));
        assert!(is_undecodable_part(parts_equal(&before, &before, &bad)));
        assert!(is_undecodable_part(encode_delta(
            &before,
            &after,
            MAX_DELTA_BYTES
        )));
        assert!(is_undecodable_part(encode_delta(
            &after,
            &before,
            MAX_DELTA_BYTES
        )));

        // Without the corrupt member the same transition still encodes.
        let clean = lazy_package(b"<a/>", false);
        assert!(encode_delta(&clean, &after, MAX_DELTA_BYTES).is_ok());
    }

    #[test]
    fn delta_source_guard_refuses_a_present_part_that_fails_to_decode() {
        let record = |before: Option<&[u8]>| Record {
            kind: 2,
            name: "/custom/bad.bin".into(),
            before: before.map(|payload| Member {
                content_type: Some("application/octet-stream".into()),
                payload: payload.to_vec(),
                relationships: Some((false, Vec::new())),
            }),
            after: None,
        };
        fn is_refusal(result: Result<()>, reason: &str) -> bool {
            matches!(result, Err(Error::InvalidFormat(message)) if message.ends_with(reason))
        }

        // A truly absent part matches a closure that records it as absent,
        // and presence is decided without decoding anything.
        let clean = lazy_package(b"<a/>", false);
        assert!(validate_delta_source(&clean, &[record(None)]).is_ok());
        assert!(is_refusal(
            validate_delta_source(&clean, &[record(Some(b"payload"))]),
            "missing Ink durable source part"
        ));
        assert_eq!(clean.deferred_decode_counters(), Some((0, 0)));

        // A present part is never taken for an absent one, and its decode
        // failure is reported as one rather than as a missing part.
        let corrupt = lazy_package(b"<a/>", true);
        assert!(is_refusal(
            validate_delta_source(&corrupt, &[record(None)]),
            "unexpected Ink durable source part"
        ));
        assert!(matches!(
            validate_delta_source(&corrupt, &[record(Some(b"payload"))]),
            Err(Error::Opc(litchi_opc::OpcError::ZipError(_)))
        ));
    }

    #[test]
    fn durable_anchor_booleans_are_strictly_binary() {
        let mut bytes = INTENT_HEADER.to_vec();
        bytes.push(1);
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&0i64.to_le_bytes());
        bytes.extend_from_slice(&0i64.to_le_bytes());
        bytes.extend_from_slice(&[0, 0, 0]);
        bytes.extend_from_slice(&[0, 0, 0]);
        bytes.push(0);
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.push(2);
        let mut decoder = IntentDecoder::new(&bytes, INTENT_HEADER).expect("decoder");
        let error = decode_placement(&mut decoder).expect_err("non-binary anchor flag");
        assert!(error.to_string().contains("invalid Ink durable boolean"));
    }
}

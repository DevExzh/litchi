//! Focused immutable transactions for Keynote slide playback state.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::wire::patch_varint_field;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use thiserror::Error;

use super::{Package, ReadError, strict_slide_node_skipped};
use crate::{SlideSelector, SlideSelectorError};

/// An error raised while staging or committing a Keynote slide-state edit.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum EditError {
    /// The immutable source package could not provide its semantic snapshot.
    #[error(transparent)]
    Read(#[from] ReadError),
    /// A semantic slide selector was ambiguous.
    #[error(transparent)]
    Selector(#[from] SlideSelectorError),
    /// No slide has the requested exact navigator name.
    #[error("Keynote show has no slide matching the requested navigator name")]
    SlideNameNotFound,
    /// The requested semantic position is outside the base snapshot.
    #[error("Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// The missing typed source position.
        position: Position,
    },
    /// This focused transaction already contains its one supported operation.
    #[error("Keynote slide-state transaction already has a staged operation")]
    OperationAlreadyStaged,
    /// Commit was requested before an operation was staged.
    #[error("Keynote slide-state transaction has no staged operation")]
    NoStagedOperation,
    /// The physical package boundary rejected the candidate rewrite.
    #[error(transparent)]
    Archive(#[from] litchi_iwa_archive::Error),
    /// The IWA framing boundary rejected the candidate rewrite.
    #[error(transparent)]
    Iwa(#[from] litchi_iwa_core::Error),
    /// The raw protobuf wire boundary rejected the selected skip field.
    #[error(transparent)]
    Wire(#[from] litchi_iwa_common::Error),
    /// The semantic selector did not resolve to exactly one editable payload.
    #[error("invalid Keynote slide-state target: {reason}")]
    InvalidTarget {
        /// Stable description that does not expose native object identity.
        reason: &'static str,
    },
    /// Full candidate reopening did not reproduce the requested semantic value.
    #[error("Keynote slide-state candidate verification failed: {reason}")]
    Verification {
        /// Stable verification failure description.
        reason: &'static str,
    },
    /// The supplied patch was not created from this exact package artifact.
    #[error("Keynote slide-state patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct SkipIntent {
    position: Position,
    before: bool,
    after: bool,
}

/// A focused edit staged against one immutable Keynote base snapshot.
///
/// Selectors are resolved when the operation is staged, so later publication
/// always addresses the same base-snapshot position. This bounded transaction
/// intentionally accepts one skip/include operation per commit.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    intent: Option<SkipIntent>,
}

impl<'a> Edit<'a> {
    pub(super) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            intent: None,
        }
    }

    /// Stage one semantic slide skip-state replacement.
    ///
    /// # Errors
    ///
    /// Returns a typed missing/ambiguity error when `selector` does not resolve
    /// to exactly one slide in the immutable base snapshot, or when another
    /// operation has already been staged.
    pub fn set_slide_skipped<'selector>(
        &mut self,
        selector: impl Into<SlideSelector<'selector>>,
        is_skipped: bool,
    ) -> Result<&mut Self, EditError> {
        if self.intent.is_some() {
            return Err(EditError::OperationAlreadyStaged);
        }

        let semantic_selector = selector.into();
        let selected = self.source.show()?.select_slide(semantic_selector)?;
        let slide = selected.ok_or(match semantic_selector {
            SlideSelector::Name(_) => EditError::SlideNameNotFound,
            SlideSelector::Position(position) => EditError::SlidePositionNotFound { position },
        })?;
        self.intent = Some(SkipIntent {
            position: Position::new(slide.index()),
            before: slide.is_skipped(),
            after: is_skipped,
        });
        Ok(self)
    }

    /// Stage omission of one semantic slide during playback.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_slide_skipped`].
    pub fn skip_slide<'selector>(
        &mut self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<&mut Self, EditError> {
        self.set_slide_skipped(selector, true)
    }

    /// Stage inclusion of one semantic slide during playback.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::set_slide_skipped`].
    pub fn include_slide<'selector>(
        &mut self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<&mut Self, EditError> {
        self.set_slide_skipped(selector, false)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// An exact semantic no-op reuses the original source allocation and bytes.
    /// A change is published only after the complete package is reopened,
    /// validated, and semantically read back under the retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error without modifying the source when staging is empty or
    /// any wire, IWA, ZIP, limit, or readback invariant fails.
    pub fn commit(self) -> Result<Commit, EditError> {
        let intent = self.intent.ok_or(EditError::NoStagedOperation)?;
        let source_fingerprint = fingerprint(self.source.source_bytes());
        if intent.before == intent.after {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source: Arc::clone(&self.source.state.source),
                    target: Arc::clone(&self.source.state.source),
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: intent.position,
                    before: intent.before,
                    after: intent.after,
                },
                diagnostics: Diagnostics {
                    changed: false,
                    touched_components: 0,
                    full_reparse_performed: false,
                },
            });
        }

        let package = rewrite_skip_state(self.source, intent)?;
        let target_fingerprint = fingerprint(package.source_bytes());
        Ok(Commit {
            patch: Patch {
                source: Arc::clone(&self.source.state.source),
                target: Arc::clone(&package.state.source),
                source_fingerprint,
                target_fingerprint,
                position: intent.position,
                before: intent.before,
                after: intent.after,
            },
            package,
            diagnostics: Diagnostics {
                changed: true,
                touched_components: 1,
                full_reparse_performed: true,
            },
        })
    }
}

/// A reversible semantic Keynote slide-state patch.
///
/// The patch contains no native object IDs or component names. It privately
/// retains cheap immutable source and target byte handles, so [`Package::apply`]
/// authorizes publication with an exact byte comparison rather than treating
/// the compact fingerprints as collision-resistant identities.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: bool,
    after: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the base package fingerprint used for future conflict checks.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return the base-snapshot slide position affected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the semantic skip state required before this patch.
    #[must_use]
    pub const fn before(&self) -> bool {
        self.before
    }

    /// Return the semantic skip state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> bool {
        self.after
    }

    /// Return whether this patch changes the semantic skip state.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Return the reversible semantic inverse of this patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            before: self.after,
            after: self.before,
        }
    }
}

impl Package {
    /// Apply an exact-source-checked semantic patch without mutating this snapshot.
    ///
    /// The fingerprints are diagnostic prechecks only. Publication is authorized
    /// by an exact comparison with the immutable source bytes retained by `patch`.
    /// The target artifact is fully reopened, validated, and semantically read
    /// back under this package's retained resource limits.
    ///
    /// # Errors
    ///
    /// Returns [`EditError::PatchConflict`] when this is not the exact patch
    /// source, or another validation error when the retained target is invalid
    /// under this package's limits.
    pub fn apply(&self, patch: &Patch) -> Result<Commit, EditError> {
        if fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source.as_ref()
        {
            return Err(EditError::PatchConflict);
        }

        let source_slide = self
            .show()?
            .select_slide(patch.position)?
            .ok_or(EditError::PatchConflict)?;
        if source_slide.is_skipped() != patch.before {
            return Err(EditError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref() {
                return Err(EditError::PatchConflict);
            }
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: false,
                    touched_components: 0,
                    full_reparse_performed: false,
                },
            });
        }

        if fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(EditError::PatchConflict);
        }
        let candidate = Package::from_source(Arc::clone(&patch.target), self.state.limits)?;
        candidate.validate()?;
        let target_slide =
            candidate
                .show()?
                .select_slide(patch.position)?
                .ok_or(EditError::Verification {
                    reason: "patch target slide position is missing",
                })?;
        if target_slide.is_skipped() != patch.after {
            return Err(EditError::Verification {
                reason: "patch target semantic skip state is invalid",
            });
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: 1,
                full_reparse_performed: true,
            },
        })
    }
}

/// Compact evidence describing work performed by one committed transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the committed package differs semantically from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of IWA components rewritten by this commit.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The published result of one atomic Keynote slide-state transaction.
#[must_use = "a Keynote commit contains the validated immutable package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return the fully reopened package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible semantic patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn rewrite_skip_state(source: &Package, intent: SkipIntent) -> Result<Package, EditError> {
    let record =
        source
            .slide_record_at(intent.position.get())?
            .ok_or(EditError::InvalidTarget {
                reason: "resolved semantic slide has no native slide-node record",
            })?;

    let mut components = source
        .state
        .components
        .iter()
        .filter(|component| component.archive().object(record.node_identifier).is_some());
    let component = components.next().ok_or(EditError::InvalidTarget {
        reason: "slide-node component is missing",
    })?;
    if components.next().is_some() {
        return Err(EditError::InvalidTarget {
            reason: "slide-node component is ambiguous",
        });
    }
    let component_name = component.name();

    let catalog = Catalog::from_shared_bytes_with_limits(
        Arc::clone(&source.state.source),
        source.state.limits,
    )?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(EditError::InvalidTarget {
            reason: "slide-node package member is missing",
        })?;
    if entry.is_opaque() {
        return Err(EditError::InvalidTarget {
            reason: "slide-node package member uses unsupported compression",
        });
    }

    let archive_limits = source.state.limits.effective_archive_limits()?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), source.state.limits.snappy_limits()?)?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)?;
    let object = archive
        .object(record.node_identifier)
        .ok_or(EditError::InvalidTarget {
            reason: "selected slide node disappeared from its component",
        })?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == 4);
    let (message_index, message) = messages.next().ok_or(EditError::InvalidTarget {
        reason: "selected slide node has no type-4 payload",
    })?;
    if messages.next().is_some() {
        return Err(EditError::InvalidTarget {
            reason: "selected slide node has multiple type-4 payloads",
        });
    }

    let wire_limits = source.wire_limits()?;
    let physical_before = strict_slide_node_skipped(&message.data, wire_limits)?;
    if physical_before != intent.before {
        return Err(EditError::InvalidTarget {
            reason: "selected slide-node wire state disagrees with the base snapshot",
        });
    }
    let patched = patch_varint_field(&message.data, 4, true, Some(u64::from(intent.after)))?;
    if patched.len() != message.data.len() {
        return Err(EditError::InvalidTarget {
            reason: "Boolean slide-state rewrite changed payload length",
        });
    }
    if strict_slide_node_skipped(&patched, wire_limits)? != intent.after {
        return Err(EditError::Verification {
            reason: "rewritten slide-node wire value was not retained",
        });
    }

    archive
        .object_mut(record.node_identifier)
        .ok_or(EditError::InvalidTarget {
            reason: "selected slide node disappeared during staging",
        })?
        .replace_message_with_limits(
            message_index,
            RawMessage {
                type_: 4,
                data: patched,
            },
            archive_limits,
        )?;
    let rewritten_archive = archive.to_bytes_with_limits(archive_limits)?;
    let compressed = SnappyStream::compress(&rewritten_archive)?;
    let output = catalog.reassemble_to_bytes(
        &[EntryEdit::new(component_name, &compressed)],
        source.state.limits,
    )?;

    let candidate = Package::from_source(output.into(), source.state.limits)?;
    candidate.validate()?;
    let readback =
        candidate
            .show()?
            .select_slide(intent.position)?
            .ok_or(EditError::Verification {
                reason: "committed slide position is missing",
            })?;
    if readback.is_skipped() != intent.after {
        return Err(EditError::Verification {
            reason: "committed semantic skip state does not match the request",
        });
    }
    Ok(candidate)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

//! Private planning primitives for fresh, slide-owned Keynote audio.
//!
//! The public creation vocabulary deliberately hides native object identifiers
//! and archive internals.  This module keeps the identity and resource ledger
//! needed by the package owner and its metadata/playback adapters in one place.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Creation planning keeps the source admission and ledger helpers together."
)]

use std::{mem::size_of, sync::Arc};

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, EntryInsertion, ExactArtifacts};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveLimits, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_media_creation_codec as media_creation_codec,
    keynote_media_lifecycle_codec as lifecycle_codec, keynote_show_codec,
    package_metadata_codec as identity_codec, table_appearance_codec as appearance_codec,
};
use sha1::{Digest as _, Sha1};

use super::{Package, PhysicalSource};
use crate::slide::audio::creation::{
    SlideAudioCreationCommit, SlideAudioCreationDiagnostics, SlideAudioCreationError,
    SlideAudioCreationLimitKind, SlideAudioCreationPatch,
};
use crate::soundtrack::items::MAX_FILENAME_BYTES;
use crate::{MovieKind, SlideSelector};

mod data;
mod metadata;
pub(super) mod playback_build;
mod verification;

pub(super) const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
pub(super) const SLIDE_MESSAGE_TYPE: u32 = 5;
pub(super) const MOVIE_MESSAGE_TYPE: u32 = 3_007;
pub(super) const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
pub(super) const STYLESHEET_MESSAGE_TYPE: u32 = 401;
pub(super) const MEDIA_STYLE_MESSAGE_TYPE: u32 = 3_016;
pub(super) const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
pub(super) const SLIDE_BUILDS_FIELD: u32 = 2;
pub(super) const SLIDE_BUILD_CHUNKS_FIELD: u32 = 43;
pub(super) const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
pub(super) const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
pub(super) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
pub(super) const DATA_PREFIX: &str = "Data/";
pub(super) const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
pub(super) const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
pub(super) const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];

/// Native identities reserved by one fresh slide-audio transaction.
///
/// The five identifiers are the drawable, title stand-in, caption stand-in,
/// build, and build-chunk objects.  The chunk carries the build UUID in its
/// payload and therefore does not receive a separate metadata UUID addition;
/// `object_uuids` is ordered for the other four objects.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CreationIds {
    pub(super) drawable: u64,
    pub(super) title: u64,
    pub(super) caption: u64,
    pub(super) build: u64,
    pub(super) chunk: u64,
    /// Last identifier observed before allocation began.
    pub(super) expected_last_identifier: u64,
    /// Last identifier written to the metadata watermark.
    pub(super) last_identifier: u64,
    /// UUIDs for drawable, title, caption, and build, in that order.
    pub(super) object_uuids: [identity_codec::UuidBits; 4],
    /// UUID embedded in both the build and build-chunk payloads.
    pub(super) build_uuid: identity_codec::UuidBits,
    /// Deterministic seed retained by the native build archive.
    pub(super) random_number_seed: u32,
}

impl CreationIds {
    /// Return the five identifiers in the order they are appended to the
    /// selected component archive.
    #[must_use]
    pub(super) const fn object_identifiers(self) -> [u64; 5] {
        [
            self.drawable,
            self.title,
            self.caption,
            self.build,
            self.chunk,
        ]
    }

    /// Return the four identifiers that receive PackageMetadata UUID entries.
    #[must_use]
    pub(super) const fn metadata_identifiers(self) -> [u64; 4] {
        [self.drawable, self.title, self.caption, self.build]
    }

    /// Return the UUID additions in the same order as
    /// [`Self::metadata_identifiers`].
    #[must_use]
    pub(super) const fn metadata_uuids(self) -> [identity_codec::UuidBits; 4] {
        self.object_uuids
    }

    /// Validate that all newly reserved identifiers are nonzero and unique.
    pub(super) fn validate(self) -> Result<(), SlideAudioCreationError> {
        let identifiers = self.object_identifiers();
        for (index, identifier) in identifiers.iter().copied().enumerate() {
            if identifier == 0 || identifiers[..index].contains(&identifier) {
                return Err(SlideAudioCreationError::InvalidSource);
            }
        }
        if self.last_identifier < self.expected_last_identifier
            || identifiers
                .iter()
                .copied()
                .any(|identifier| identifier > self.last_identifier)
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        if self
            .object_uuids
            .iter()
            .any(|uuid| uuid == &identity_codec::UuidBits::new(0, 0))
            || self.build_uuid == identity_codec::UuidBits::new(0, 0)
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        Ok(())
    }
}

/// Source topology and bounded decode profiles selected for one transaction.
///
/// This context is intentionally private: callers select a semantic slide and
/// the package owner resolves component names, native references, and styles.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct CreationContext {
    pub(super) slide_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) slide_component: Box<str>,
    pub(super) slide_node_identifier: u64,
    pub(super) stylesheet_component: Box<str>,
    pub(super) stylesheet_identifier: u64,
    pub(super) style_identifier: u64,
    pub(super) archive_limits: ArchiveLimits,
    pub(super) wire_limits: WireLimits,
}

impl CreationContext {
    /// Reject the zero native references and empty component names that would
    /// otherwise make a later archive edit ambiguous.
    pub(super) fn validate(&self) -> Result<(), SlideAudioCreationError> {
        if self.slide_identifier == 0
            || self.slide_node_identifier == 0
            || self.stylesheet_identifier == 0
            || self.style_identifier == 0
            || self.slide_component.is_empty()
            || self.stylesheet_component.is_empty()
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        Ok(())
    }
}

/// One operation-wide finite resource ledger.
///
/// Every fallible phase debits this same ledger before allocating or invoking
/// a bounded codec.  The counters intentionally use `usize` because all
/// downstream APIs are allocation-sized values; checked conversion from the
/// package's `u64` profile happens during construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CreationBudget {
    max_input: usize,
    max_output: usize,
    max_entries: usize,
    max_objects: usize,
    max_entry_bytes: usize,
    max_total: usize,
    max_slides: usize,
    max_references: usize,
    max_media_bytes: usize,
    max_wire_fields: usize,
    max_nesting: usize,
    max_work: usize,
    max_allocations: usize,
    input: usize,
    output: usize,
    entries: usize,
    objects: usize,
    total: usize,
    slides: usize,
    references: usize,
    media_bytes: usize,
    wire_fields: usize,
    observed_nesting: usize,
    work: usize,
    allocation_bytes: usize,
    allocations: usize,
}

impl CreationBudget {
    /// Construct a ledger from the package's checked physical and semantic
    /// profiles and precharge the retained source bytes.
    pub(super) fn for_package(package: &Package) -> Result<Self, SlideAudioCreationError> {
        let physical = package.limits();
        let semantic = package.semantic_limits();
        let wire = package
            .semantic_wire_limits()
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        let max_input = to_usize(physical.max_input_bytes())?;
        let max_output = to_usize(
            physical
                .max_total_bytes()
                .checked_mul(16)
                .ok_or(SlideAudioCreationError::InvalidSource)?,
        )?;
        let max_entries = physical.max_entries();
        let max_entry_bytes = to_usize(physical.max_entry_bytes())?;
        let max_total = to_usize(physical.max_total_bytes())?;
        let max_references = semantic.max_references();
        let max_media_bytes = max_total;
        let max_wire_fields = wire.max_fields();
        let max_nesting = wire.max_nesting();
        let max_work = wire.max_rewrite_work();
        let max_allocations = max_entries
            .checked_mul(32)
            .ok_or(SlideAudioCreationError::InvalidSource)?
            .max(1);
        let mut budget = Self {
            max_input,
            max_output,
            max_entries,
            max_objects: semantic.max_objects(),
            max_entry_bytes,
            max_total,
            max_slides: semantic.max_slides(),
            max_references,
            max_media_bytes,
            max_wire_fields,
            max_nesting,
            max_work,
            max_allocations,
            input: 0,
            output: 0,
            entries: 0,
            objects: 0,
            total: 0,
            slides: 0,
            references: 0,
            media_bytes: 0,
            wire_fields: 0,
            observed_nesting: 0,
            work: 0,
            allocation_bytes: 0,
            allocations: 0,
        };
        let source = match &package.state.source {
            PhysicalSource::Package(source) if source.source_is_exact() => source,
            PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
                return Err(SlideAudioCreationError::UnsupportedSource);
            },
        };
        budget.charge_input(source.source_bytes().len())?;
        Ok(budget)
    }

    /// Charge source bytes retained or inspected by the transaction.
    pub(super) fn charge_input(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.input,
            self.max_input,
            amount,
            SlideAudioCreationLimitKind::InputBytes,
        )
    }

    /// Charge candidate output bytes.
    pub(super) fn charge_output(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.output,
            self.max_output,
            amount,
            SlideAudioCreationLimitKind::OutputBytes,
        )
    }

    /// Charge retained physical entries.
    pub(super) fn charge_entries(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.entries,
            self.max_entries,
            amount,
            SlideAudioCreationLimitKind::Entries,
        )
    }

    /// Charge semantic archive objects independently from ZIP member count.
    pub(super) fn charge_objects(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.objects,
            self.max_objects,
            amount,
            SlideAudioCreationLimitKind::Objects,
        )
    }

    /// Check one entry's bytes before decompression or copying.
    pub(super) fn charge_entry_bytes(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        if amount > self.max_entry_bytes {
            return Err(limit(
                SlideAudioCreationLimitKind::EntryBytes,
                amount,
                self.max_entry_bytes,
            ));
        }
        Ok(())
    }

    /// Charge aggregate ZIP bytes.
    pub(super) fn charge_total(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.total,
            self.max_total,
            amount,
            SlideAudioCreationLimitKind::TotalBytes,
        )
    }

    /// Charge rooted object references.
    pub(super) fn charge_references(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.references,
            self.max_references,
            amount,
            SlideAudioCreationLimitKind::References,
        )
    }

    /// Charge semantic slide traversal.
    pub(super) fn charge_slides(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.slides,
            self.max_slides,
            amount,
            SlideAudioCreationLimitKind::Slides,
        )
    }

    /// Charge caller-provided or generated audio bytes.
    pub(super) fn charge_media_bytes(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.media_bytes,
            self.max_media_bytes,
            amount,
            SlideAudioCreationLimitKind::MediaBytes,
        )
    }

    /// Charge parsed or emitted wire fields.
    pub(super) fn charge_wire_fields(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.wire_fields,
            self.max_wire_fields,
            amount,
            SlideAudioCreationLimitKind::WireFields,
        )
    }

    /// Record maximum nesting observed by a strict codec.
    pub(super) fn charge_nesting(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        self.observed_nesting = self.observed_nesting.max(amount);
        if self.observed_nesting > self.max_nesting {
            return Err(limit(
                SlideAudioCreationLimitKind::WireNesting,
                self.observed_nesting,
                self.max_nesting,
            ));
        }
        Ok(())
    }

    /// Charge aggregate codec work.
    pub(super) fn charge_work(&mut self, amount: usize) -> Result<(), SlideAudioCreationError> {
        Self::charge(
            &mut self.work,
            self.max_work,
            amount,
            SlideAudioCreationLimitKind::WireWork,
        )
    }

    /// Charge bytes and one logical allocation event before a fallible
    /// reservation or codec execution.
    pub(super) fn charge_allocations(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        let allocation_bytes = self
            .allocation_bytes
            .checked_add(amount)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if allocation_bytes > self.max_output {
            return Err(limit(
                SlideAudioCreationLimitKind::Allocations,
                allocation_bytes,
                self.max_output,
            ));
        }
        let allocations = self
            .allocations
            .checked_add(1)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if allocations > self.max_allocations {
            return Err(limit(
                SlideAudioCreationLimitKind::Allocations,
                allocations,
                self.max_allocations,
            ));
        }
        self.allocation_bytes = allocation_bytes;
        self.allocations = allocations;
        Ok(())
    }

    /// Charge an exact number of allocation events without adding bytes.
    pub(super) fn charge_allocation_events(
        &mut self,
        amount: usize,
    ) -> Result<(), SlideAudioCreationError> {
        let allocations = self
            .allocations
            .checked_add(amount)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if allocations > self.max_allocations {
            return Err(limit(
                SlideAudioCreationLimitKind::Allocations,
                allocations,
                self.max_allocations,
            ));
        }
        self.allocations = allocations;
        Ok(())
    }

    /// Return the remaining output ceiling for a bounded codec profile.
    pub(super) fn remaining_output(&self) -> usize {
        self.max_output.saturating_sub(self.output)
    }

    /// Return the remaining wire-work ceiling.
    pub(super) fn remaining_work(&self) -> usize {
        self.max_work.saturating_sub(self.work)
    }

    /// Return the remaining wire-field ceiling.
    pub(super) fn remaining_wire_fields(&self) -> usize {
        self.max_wire_fields.saturating_sub(self.wire_fields)
    }

    /// Return the remaining object-reference ceiling for bounded lifecycle
    /// codec preflights.
    pub(super) fn remaining_references(&self) -> usize {
        self.max_references.saturating_sub(self.references)
    }

    /// Return the remaining logical allocation-event ceiling for bounded
    /// media/build codec preflights.
    pub(super) fn remaining_allocations(&self) -> usize {
        self.max_allocations.saturating_sub(self.allocations)
    }

    fn charge(
        current: &mut usize,
        maximum: usize,
        amount: usize,
        kind: SlideAudioCreationLimitKind,
    ) -> Result<(), SlideAudioCreationError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if observed > maximum {
            return Err(limit(kind, observed, maximum));
        }
        *current = observed;
        Ok(())
    }
}

impl Package {
    /// Create one independently positioned, slide-owned audio control.
    ///
    /// The source remains unchanged. The returned commit contains the candidate
    /// package and an exact patch whose inverse restores the original bytes.
    /// Identical materialized audio is reused when its metadata is unambiguous.
    ///
    /// ```no_run
    /// use std::{fs, time::Duration};
    /// use litchi_iwa_common::shape::geometry::Point;
    /// use litchi_keynote::{Package, SlideSelector, slide::audio::Options};
    ///
    /// let source = Package::from_bytes(&fs::read("presentation.key")?)?;
    /// let audio = fs::read("narration.wav")?;
    /// let options = Options::new(Point { x: 120.0, y: 240.0 }, Duration::from_secs(8))?;
    /// let commit = source.add_slide_audio(SlideSelector::index(0), "narration.wav", &audio, options)?;
    /// commit.package().write_to(&mut fs::File::create("narrated.key")?)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn add_slide_audio<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        preferred_filename: &str,
        data: &[u8],
        options: crate::slide::audio::Options,
    ) -> Result<SlideAudioCreationCommit, SlideAudioCreationError> {
        validate_audio_input(self, preferred_filename, data)?;
        let mut budget = CreationBudget::for_package(self)?;
        let catalog = physical_catalog(self)?;
        budget.charge_entries(catalog.package().len())?;
        for entry in catalog.package().iter() {
            budget.charge_entry_bytes(entry.data().len())?;
            budget.charge_total(entry.data().len())?;
        }
        budget.charge_media_bytes(data.len())?;
        let context = resolve_context(self, slide.into(), &mut budget)?;
        let source_media_count = count_slide_movies(self, &context, &mut budget)?;
        let target_media_count = source_media_count
            .checked_add(1)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let ids = allocate_ids(self, &mut budget)?;
        ids.validate()?;

        // Metadata planning runs before media encoding because the data
        // identifier is a separate PackageMetadata domain and may be reused
        // by a content-identical existing entry.
        let metadata_plan = metadata::plan_and_rewrite(
            self,
            &context,
            &ids,
            preferred_filename,
            data,
            &mut budget,
        )?;
        let media_objects =
            make_audio_objects(&context, &ids, &metadata_plan, options, &mut budget)?;
        let builds = playback_build::prepare_start_audio_builds(&ids, &context, &mut budget)?;

        let data_path = metadata_plan
            .data_entry_name
            .as_deref()
            .map(normalize_data_path);
        let data_path_ref = data_path.as_deref();
        let insertion = data_path_ref.map(|path| EntryInsertion::new(path, data));
        let insertions = insertion.as_slice();
        let publication = publish_candidate(
            self,
            catalog,
            &context,
            &ids,
            media_objects,
            builds,
            &metadata_plan.compressed,
            insertions,
            &mut budget,
        )?;
        let target = publication.target;
        let event_count = publication.event_count;
        let candidate = Package::from_source_with_options(Arc::clone(&target), self.state.options)
            .map_err(|_| SlideAudioCreationError::Verification)?;
        verify_candidate(
            self,
            &candidate,
            &context,
            options,
            data,
            metadata_plan.digest,
            source_media_count,
            target_media_count,
            &mut budget,
        )?;
        verification::verify_created_graph(
            &candidate,
            &context,
            &ids,
            metadata_plan.data_identifier,
            event_count,
            &mut budget,
        )?;
        let patch = SlideAudioCreationPatch {
            artifacts: ExactArtifacts::new(catalog.shared_source(), target),
            slide_position: context.slide_position,
            movie_position: Position::new(source_media_count),
            source_media_count,
            target_media_count,
            options,
            data_digest: metadata_plan.digest,
            data_len: data.len(),
            target_contains_created_audio: true,
            created_objects: 5,
            removed_objects: 0,
            created_data: metadata_plan.created_data,
            removed_data: 0,
            touched_members: publication.touched_members,
            deleted_previews: publication.deleted_previews,
            restored_previews: 0,
        };
        let diagnostics = SlideAudioCreationDiagnostics::for_patch(&patch);
        Ok(SlideAudioCreationCommit {
            package: candidate,
            patch,
            diagnostics,
        })
    }

    /// Apply a creation patch only to its exact retained source snapshot.
    pub fn apply_slide_audio_creation(
        &self,
        patch: &SlideAudioCreationPatch,
    ) -> Result<SlideAudioCreationCommit, SlideAudioCreationError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideAudioCreationError::PatchConflict);
        }
        let mut budget = CreationBudget::for_package(self)?;
        budget.charge_output(patch.artifacts.target().len())?;
        budget.charge_allocations(patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(|_| SlideAudioCreationError::PatchConflict)?;
        candidate
            .validate()
            .map_err(|_| SlideAudioCreationError::PatchConflict)?;
        let context = resolve_context(
            &candidate,
            SlideSelector::position(patch.slide_position),
            &mut budget,
        )?;
        let source_media_count = count_slide_movies(&candidate, &context, &mut budget)?;
        if source_media_count != patch.target_media_count {
            return Err(SlideAudioCreationError::PatchConflict);
        }
        if patch.target_contains_created_audio {
            verify_created_audio_semantics(&candidate, patch, &context, &mut budget)?;
        }
        Ok(SlideAudioCreationCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideAudioCreationDiagnostics::for_patch(patch),
        })
    }
}

/// Exact publication artifacts shared by fresh audio and movie creation.
#[derive(Debug)]
pub(super) struct CreationPublication {
    pub(super) target: Arc<[u8]>,
    pub(super) event_count: usize,
    pub(super) touched_members: usize,
    pub(super) deleted_previews: usize,
}

/// Package-level projection of the metadata planner's two-asset result.
///
/// The metadata adapter remains private to the audio creation module.  Movie
/// creation only needs these staged values, so exposing this narrow projection
/// keeps the sibling transaction independent of metadata codec internals.
#[derive(Debug)]
pub(super) struct MovieMetadataPlan {
    pub(super) data_identifier: u64,
    pub(super) digest: [u8; 20],
    pub(super) compressed: Vec<u8>,
    pub(super) data_entry_name: Option<Box<str>>,
    pub(super) created_data: usize,
    pub(super) poster: MovieMetadataAsset,
}

#[derive(Debug)]
pub(super) struct MovieMetadataAsset {
    pub(super) data_identifier: u64,
    pub(super) digest: [u8; 20],
    pub(super) data_entry_name: Option<Box<str>>,
    pub(super) created_data: usize,
}

/// Adapt the shared metadata planner to the movie transaction's compact
/// package-level vocabulary.
pub(super) fn plan_movie_metadata(
    source: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    movie_filename: &str,
    movie_data: &[u8],
    poster_filename: &str,
    poster_data: &[u8],
    budget: &mut CreationBudget,
) -> Result<MovieMetadataPlan, SlideAudioCreationError> {
    let plan = metadata::plan_and_rewrite_movie(
        source,
        context,
        ids,
        movie_filename,
        movie_data,
        poster_filename,
        poster_data,
        budget,
    )?;
    let metadata::MetadataPlan {
        data_identifier,
        digest,
        compressed,
        data_entry_name,
        created_data,
        poster,
    } = plan;
    let poster = poster.ok_or(SlideAudioCreationError::InvalidSource)?;
    Ok(MovieMetadataPlan {
        data_identifier,
        digest,
        compressed,
        data_entry_name,
        created_data,
        poster: MovieMetadataAsset {
            data_identifier: poster.data_identifier,
            digest: poster.digest,
            data_entry_name: poster.data_entry_name,
            created_data: poster.created_data,
        },
    })
}

/// Adapt the movie-specific playback writer without exposing its codec enum.
pub(super) fn prepare_movie_builds(
    ids: &CreationIds,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<[ArchiveObject; 2], SlideAudioCreationError> {
    playback_build::prepare_start_movie_builds(ids, context, budget)
}

/// Verify a freshly published movie graph through the shared bounded witness.
pub(super) fn verify_movie_graph(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    position: [f32; 2],
    size: [f32; 2],
    duration_seconds: f32,
    natural_size: [f32; 2],
    expected_event_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let media = verification::CreatedMediaExpectation::movie(
        movie_data_identifier,
        poster_data_identifier,
        position,
        size,
        duration_seconds,
        natural_size,
        true,
        Some(DEFAULT_DRAWABLE_FLAGS),
        Some(0.0),
    );
    verification::verify_created_graph_with_media(
        package,
        context,
        ids,
        media,
        expected_event_count,
        budget,
    )
}

/// Read both materialized edges of one movie through one metadata witness.
pub(super) fn read_movie_content_and_poster<'a>(
    package: &'a Package,
    slide: Position,
    movie: Position,
    budget: &mut CreationBudget,
) -> Result<(&'a [u8], &'a [u8]), SlideAudioCreationError> {
    data::read_content_and_poster(package, slide, movie, budget)
}

/// Stage one media graph, update the selected slide and node cache, invalidate
/// previews, and perform one exact source reassembly.
///
/// Keeping publication in one helper is significant: audio and movie
/// creation must debit the same clone, archive, ZIP, and preview allocations,
/// and they must produce identical save-token and unknown-header behavior.
pub(super) fn publish_candidate(
    source: &Package,
    catalog: &SourceCatalog,
    context: &CreationContext,
    ids: &CreationIds,
    media_objects: [ArchiveObject; 3],
    builds: [ArchiveObject; 2],
    metadata_bytes: &[u8],
    insertions: &[EntryInsertion<'_>],
    budget: &mut CreationBudget,
) -> Result<CreationPublication, SlideAudioCreationError> {
    let source_component = source
        .state
        .source
        .components()
        .get(context.slide_component.as_ref())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let clone_bound =
        archive_clone_allocation_bound(source_component.archive(), context.archive_limits)?;
    budget.charge_allocations(clone_bound)?;
    budget.charge_allocation_events(1)?;
    let mut slide_archive = source_component.archive().clone();
    let append_bound = 5usize
        .checked_mul(size_of::<ArchiveObject>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(append_bound)?;
    budget.charge_allocation_events(1)?;
    let appended = media_objects.into_iter().chain(builds).collect::<Vec<_>>();
    slide_archive
        .append_objects_with_limits(appended, context.archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    rewrite_slide_archive(
        &mut slide_archive,
        context,
        ids,
        context.archive_limits,
        context.wire_limits,
        budget,
    )?;

    let event_count = slide_build_count(&slide_archive, context, budget)?;
    let node_edit =
        prepare_node_cache_edit(source, context, &mut slide_archive, event_count, budget)?;
    let snappy_limits = source
        .limits()
        .snappy_limits()
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let slide_bytes = serialize_component(
        &slide_archive,
        context.archive_limits,
        snappy_limits,
        budget,
    )?;
    let node_bytes = node_edit
        .as_ref()
        .map(|(_, archive)| {
            serialize_component(archive, context.archive_limits, snappy_limits, budget)
        })
        .transpose()?;

    let preview_plan = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let preview_capacity = preview_plan
        .len()
        .checked_mul(size_of::<&str>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(preview_capacity)?;
    let mut deleted_names = Vec::new();
    deleted_names
        .try_reserve_exact(preview_plan.len())
        .map_err(|_| SlideAudioCreationError::Allocation {
            amount: preview_plan.len(),
        })?;
    deleted_names.extend(preview_plan.names());

    let slide_edit = EntryEdit::new(context.slide_component.as_ref(), &slide_bytes);
    let metadata_edit = EntryEdit::new(METADATA_COMPONENT, metadata_bytes);
    let node_entry_edit = node_edit
        .as_ref()
        .zip(node_bytes.as_ref())
        .map(|((name, _), bytes)| EntryEdit::new(name.as_ref(), bytes));
    let edit_count: usize = if node_entry_edit.is_some() { 3 } else { 2 };
    let edit_capacity = edit_count
        .checked_mul(size_of::<EntryEdit<'static>>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(edit_capacity)?;
    // The insertion slice is caller-owned and already has a bounded capacity;
    // only charge the operation's view of its descriptors here.
    let insertion_capacity = insertions
        .len()
        .checked_mul(size_of::<EntryInsertion<'static>>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(insertion_capacity)?;
    let mut edits = Vec::with_capacity(edit_count);
    edits.push(slide_edit);
    edits.push(metadata_edit);
    if let Some(edit) = node_entry_edit {
        edits.push(edit);
    }

    let prepared = catalog
        .package()
        .prepare_reassembly_with_changes(insertions, &edits, &deleted_names, catalog.limits())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.charge_output(requirements.output_bytes())?;
    budget.charge_allocations(requirements.retained_bytes())?;
    budget.charge_allocations(requirements.scratch_bytes())?;
    budget.charge_allocation_events(requirements.allocations())?;
    let target: Arc<[u8]> = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?
        .into();
    let touched_members = edits
        .len()
        .checked_add(insertions.len())
        .and_then(|count| count.checked_add(deleted_names.len()))
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    Ok(CreationPublication {
        target,
        event_count,
        touched_members,
        deleted_previews: preview_plan.len(),
    })
}

fn validate_audio_input(
    package: &Package,
    filename: &str,
    data: &[u8],
) -> Result<(), SlideAudioCreationError> {
    validate_media_input(
        package,
        filename,
        data,
        litchi_iwa_common::media::Type::Audio,
    )
}

/// Validate one caller-owned media member before any package allocation.
///
/// The error vocabulary remains the audio creation vocabulary because this is
/// an internal seam; the movie owner maps the two input failures into its
/// format-specific public error without exposing this implementation detail.
pub(super) fn validate_media_input(
    package: &Package,
    filename: &str,
    data: &[u8],
    expected: litchi_iwa_common::media::Type,
) -> Result<(), SlideAudioCreationError> {
    let maximum = usize::try_from(package.limits().max_entry_bytes())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    if filename.is_empty()
        || filename.len() > MAX_FILENAME_BYTES
        || filename.len() > maximum
        || filename
            .bytes()
            .any(|byte| byte == b'/' || byte == b'\\' || byte == 0 || byte.is_ascii_control())
        || filename == "."
        || filename == ".."
    {
        return Err(SlideAudioCreationError::InvalidFilename);
    }
    let Some(dot) = filename.rfind('.') else {
        return Err(SlideAudioCreationError::InvalidFilename);
    };
    if dot == 0 || dot + 1 >= filename.len() {
        return Err(SlideAudioCreationError::InvalidFilename);
    }
    if litchi_iwa_common::media::Type::from_extension(&filename[dot + 1..]) != expected {
        return Err(SlideAudioCreationError::InvalidFilename);
    }
    if data.len() > maximum {
        return Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::EntryBytes,
            observed: u64::try_from(data.len()).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    if data.is_empty() || litchi_iwa_common::media::Type::from_bytes(data) != expected {
        return Err(SlideAudioCreationError::UnsupportedAudio);
    }
    Ok(())
}

pub(super) fn physical_catalog(
    package: &Package,
) -> Result<&SourceCatalog, SlideAudioCreationError> {
    match &package.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => Ok(catalog),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(SlideAudioCreationError::UnsupportedSource)
        },
    }
}

pub(super) fn resolve_context(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut CreationBudget,
) -> Result<CreationContext, SlideAudioCreationError> {
    budget.charge_slides(1)?;
    let slide_position = match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(|_| SlideAudioCreationError::InvalidSource)?
            .map(|_| position)
            .ok_or(SlideAudioCreationError::SlidePositionNotFound { position })?,
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideAudioCreationError::EmptySlideName);
            }
            package
                .show()
                .map_err(|_| SlideAudioCreationError::InvalidSource)?
                .select_slide(SlideSelector::name(name))
                .map_err(|_| SlideAudioCreationError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideAudioCreationError::SlideNameNotFound)?
        },
    };
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?
        .ok_or(SlideAudioCreationError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (slide_component, slide_object) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let (_, node_object) = package
        .object_with_component(record.node_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    unique_payload(node_object, SLIDE_NODE_MESSAGE_TYPE)?;
    unique_payload(slide_object, SLIDE_MESSAGE_TYPE)?;

    let show_identifier = package
        .root_show_identifier()
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let show_object = package
        .object(show_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let show_payload = unique_payload(show_object, 2)?;
    let wire_limits = package
        .semantic_wire_limits()
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let recursion_limit = u32::try_from(wire_limits.max_nesting())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let show_options = keynote_show_codec::DecodeOptions::new(
        show_payload.len(),
        package.semantic_limits().max_slides(),
        recursion_limit,
    )
    .with_max_fields(wire_limits.max_fields().min(budget.remaining_wire_fields()))
    .with_max_work_bytes(wire_limits.max_rewrite_work().min(budget.remaining_work()));
    let (show_references, show_report) =
        keynote_show_codec::decode_references_with_report(show_payload, show_options)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(show_report.fields())?;
    budget.charge_work(show_report.work_bytes())?;
    budget.charge_nesting(show_report.max_depth() as usize)?;
    budget.charge_allocation_events(show_report.allocations())?;
    let stylesheet_identifier = show_references.stylesheet_identifier();
    if stylesheet_identifier == 0 {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    let (stylesheet_component, stylesheet_object) = package
        .object_with_component(stylesheet_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let stylesheet_payload = unique_payload(stylesheet_object, STYLESHEET_MESSAGE_TYPE)?;
    let stylesheet_options = appearance_codec::DecodeOptions::new(
        stylesheet_payload
            .len()
            .min(wire_limits.max_input_bytes())
            .max(1),
        wire_limits
            .max_output_bytes()
            .min(budget.remaining_output()),
        wire_limits.max_fields().min(budget.remaining_wire_fields()),
        wire_limits.max_rewrite_work().min(budget.remaining_work()),
        recursion_limit,
        package.semantic_limits().max_objects(),
    )
    .with_max_allocations(budget.remaining_allocations());
    let (stylesheet, report) =
        appearance_codec::decode_stylesheet_with_report(stylesheet_payload, stylesheet_options)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocation_events(report.allocations())?;
    let style_identifiers = stylesheet_style_identifiers(
        stylesheet.raw(),
        wire_limits,
        stylesheet.style_count(),
        budget,
    )
    .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let style_identifier = style_identifiers
        .into_iter()
        .filter(|identifier| *identifier != 0)
        .find(|identifier| {
            package.object(*identifier).is_some_and(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == MEDIA_STYLE_MESSAGE_TYPE)
            })
        })
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let context = CreationContext {
        slide_position,
        slide_identifier: record.slide_identifier,
        slide_component: slide_component.into(),
        slide_node_identifier: record.node_identifier,
        stylesheet_component: stylesheet_component.into(),
        stylesheet_identifier,
        style_identifier,
        archive_limits: package.limits().archive_limits(),
        wire_limits,
    };
    context.validate()?;
    Ok(context)
}

fn unique_payload(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&[u8], SlideAudioCreationError> {
    let mut payload = None;
    for message in object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
    {
        if payload.replace(message.data.as_slice()).is_some() {
            return Err(SlideAudioCreationError::InvalidSource);
        }
    }
    payload.ok_or(SlideAudioCreationError::InvalidSource)
}

/// Read the stylesheet's direct style references after the strict appearance
/// codec has validated its complete registry.  The appearance snapshot keeps
/// the source borrowed by design, so this small neutral wire projection is the
/// only allocation needed to choose the existing media style object.
fn stylesheet_style_identifiers(
    payload: &[u8],
    limits: WireLimits,
    expected_count: usize,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let fields = litchi_iwa_common::wire::WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let count = fields.fields().filter(|field| field.number() == 1).count();
    if count != expected_count || count == 0 {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    budget.charge_wire_fields(fields.len())?;
    budget.charge_work(payload.len().max(1))?;
    let bytes = count
        .checked_mul(size_of::<u64>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(count)
        .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    for field in fields.fields().filter(|field| field.number() == 1) {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        if field.wire_type() != 2 {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        identifiers.push(reference_identifier(field.payload(), limits, budget)?);
    }
    Ok(identifiers)
}

fn has_zero_or_duplicate_reference(references: &[u64]) -> bool {
    references
        .iter()
        .enumerate()
        .any(|(index, identifier)| *identifier == 0 || references[..index].contains(identifier))
}

fn count_new_unique_references(existing: &[u64], candidates: &[u64]) -> usize {
    candidates
        .iter()
        .enumerate()
        .filter(|(index, identifier)| {
            !existing.contains(identifier) && !candidates[..*index].contains(identifier)
        })
        .count()
}

fn collect_lifecycle_reference_ids<'source>(
    references: impl ExactSizeIterator<Item = lifecycle_codec::Reference<'source>>,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let count = references.len();
    let bytes = count
        .checked_mul(size_of::<u64>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut identifiers = Vec::new();
    if count != 0 {
        budget.charge_allocations(bytes)?;
        identifiers
            .try_reserve_exact(count)
            .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    }
    for reference in references {
        if identifiers.len() == count {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        identifiers.push(reference.identifier());
    }
    if identifiers.len() != count {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    Ok(identifiers)
}

fn clone_u32_values(values: &[u32]) -> Result<Vec<u32>, SlideAudioCreationError> {
    let bytes = values
        .len()
        .checked_mul(size_of::<u32>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut clone = Vec::new();
    if !values.is_empty() {
        clone
            .try_reserve_exact(values.len())
            .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    }
    clone.extend_from_slice(values);
    Ok(clone)
}

fn clone_u64_values(values: &[u64]) -> Result<Vec<u64>, SlideAudioCreationError> {
    let bytes = values
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut clone = Vec::new();
    if !values.is_empty() {
        clone
            .try_reserve_exact(values.len())
            .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    }
    clone.extend_from_slice(values);
    Ok(clone)
}

pub(super) fn count_slide_movies(
    package: &Package,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<usize, SlideAudioCreationError> {
    let slide = package
        .object(context.slide_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let payload = unique_payload(slide, SLIDE_MESSAGE_TYPE)?;
    let identifiers = references_in_field(
        payload,
        SLIDE_OWNED_DRAWABLES_FIELD,
        context.wire_limits,
        budget,
    )?;
    let mut count = 0usize;
    for identifier in identifiers {
        let (component, drawable) = package
            .object_with_component(identifier)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if component != context.slide_component.as_ref() {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let matches = drawable
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .count();
        if matches > 1 {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        if matches == 1 {
            count = count
                .checked_add(1)
                .ok_or(SlideAudioCreationError::InvalidSource)?;
        }
    }
    Ok(count)
}

pub(super) fn references_in_field(
    payload: &[u8],
    number: u32,
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let view = litchi_iwa_common::wire::WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(view.len())?;
    budget.charge_work(payload.len().max(1))?;
    let count = view
        .fields()
        .filter(|field| field.number() == number)
        .count();
    budget.charge_allocations(count.saturating_mul(size_of::<u64>()))?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(count)
        .map_err(|_| SlideAudioCreationError::Allocation {
            amount: count.saturating_mul(size_of::<u64>()),
        })?;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        identifiers.push(reference_identifier(field.payload(), limits, budget)?);
    }
    Ok(identifiers)
}

pub(super) fn reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<u64, SlideAudioCreationError> {
    let view = litchi_iwa_common::wire::WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(view.len())?;
    budget.charge_work(payload.len().max(1))?;
    let mut identifier = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        if field.number() != 1 || field.wire_type() != 0 {
            continue;
        }
        if identifier.is_some() {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        if width != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != width
            || value == 0
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(SlideAudioCreationError::InvalidSource)
}

pub(super) fn allocate_ids(
    package: &Package,
    budget: &mut CreationBudget,
) -> Result<CreationIds, SlideAudioCreationError> {
    let mut maximum = 0u64;
    for component in package.state.source.components().iter() {
        budget.charge_objects(component.archive().objects.len())?;
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(SlideAudioCreationError::InvalidSource)?;
            maximum = maximum.max(identifier);
            for message_info in &object.archive_info.message_infos {
                observe_object_references(&mut maximum, &message_info.object_references, budget)?;
                budget.charge_wire_fields(message_info.field_infos.len())?;
                budget.charge_work(message_info.field_infos.len())?;
                for field_info in &message_info.field_infos {
                    observe_object_references(&mut maximum, &field_info.object_references, budget)?;
                }
            }
        }
    }
    let expected_last_identifier = metadata_last_object_identifier(package, budget)?;
    maximum = maximum.max(expected_last_identifier);
    let first = maximum
        .checked_add(1)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let last = first
        .checked_add(4)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let bytes = litchi_core::id::generate_guid_bytes();
    let build_uuid = uuid_from_guid(&bytes);
    let random_number_seed = u32::from_le_bytes(
        bytes[..4]
            .try_into()
            .map_err(|_| SlideAudioCreationError::InvalidSource)?,
    );
    let mut object_uuids = [identity_codec::UuidBits::new(0, 0); 4];
    for uuid in &mut object_uuids {
        *uuid = uuid_from_guid(&litchi_core::id::generate_guid_bytes());
    }
    Ok(CreationIds {
        drawable: first,
        title: first
            .checked_add(1)
            .ok_or(SlideAudioCreationError::InvalidSource)?,
        caption: first
            .checked_add(2)
            .ok_or(SlideAudioCreationError::InvalidSource)?,
        build: first
            .checked_add(3)
            .ok_or(SlideAudioCreationError::InvalidSource)?,
        chunk: first
            .checked_add(4)
            .ok_or(SlideAudioCreationError::InvalidSource)?,
        expected_last_identifier,
        last_identifier: last,
        object_uuids,
        build_uuid,
        random_number_seed,
    })
}

/// Include every object-reference-bearing header collection when reserving a
/// fresh native identifier range.  These references are intentionally allowed
/// to be dangling: preserving an unknown or forward-compatible header edge is
/// part of the package round trip, but allocating below it would make a new
/// object collide with that edge.
fn observe_object_references(
    maximum: &mut u64,
    references: &[u64],
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    budget.charge_references(references.len())?;
    budget.charge_work(references.len())?;
    for &identifier in references {
        if identifier != 0 {
            *maximum = (*maximum).max(identifier);
        }
    }
    Ok(())
}

/// Read the PackageMetadata object watermark through its bounded neutral
/// visitor.  Physical archive identifiers and the metadata watermark are
/// independent native identity domains; allocating from only the former can
/// make a valid package fail its subsequent metadata rewrite.
fn metadata_last_object_identifier(
    package: &Package,
    budget: &mut CreationBudget,
) -> Result<u64, SlideAudioCreationError> {
    let entry = package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            (component.name() == METADATA_COMPONENT).then(|| component.archive())
        });
    let archive = entry.ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut payload = None;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if payload.replace(message.data.as_slice()).is_some() {
                return Err(SlideAudioCreationError::InvalidSource);
            }
        }
    }
    let payload = payload.ok_or(SlideAudioCreationError::InvalidSource)?;
    let wire_limits = package
        .semantic_wire_limits()
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let options = identity_codec::RewriteOptions::new(
        payload.len().min(wire_limits.max_input_bytes()).max(1),
        wire_limits.max_output_bytes(),
        wire_limits.max_fields(),
        wire_limits.max_rewrite_work(),
        u32::try_from(wire_limits.max_nesting())
            .map_err(|_| SlideAudioCreationError::InvalidSource)?,
        package.semantic_limits().max_objects(),
        package.semantic_limits().max_references(),
        package.semantic_limits().max_objects(),
    );
    let mut visitor = MetadataWatermarkVisitor;
    let inspection =
        identity_codec::inspect_package_metadata_with_visitor(payload, options, &mut visitor)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let report = inspection.report();
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocation_events(report.allocations())?;
    Ok(inspection.last_object_identifier())
}

struct MetadataWatermarkVisitor;

impl identity_codec::PackageMetadataVisitor for MetadataWatermarkVisitor {}

fn uuid_from_guid(bytes: &[u8; 16]) -> identity_codec::UuidBits {
    let mut lower = [0u8; 8];
    lower.copy_from_slice(&bytes[..8]);
    let mut upper = [0u8; 8];
    upper.copy_from_slice(&bytes[8..]);
    identity_codec::UuidBits::new(u64::from_le_bytes(lower), u64::from_le_bytes(upper))
}

fn make_audio_objects(
    context: &CreationContext,
    ids: &CreationIds,
    metadata_plan: &metadata::MetadataPlan,
    options: crate::slide::audio::Options,
    budget: &mut CreationBudget,
) -> Result<[ArchiveObject; 3], SlideAudioCreationError> {
    let write = media_creation_codec::MediaArchiveWrite::audio(
        context.slide_identifier,
        context.style_identifier,
        ids.title,
        ids.caption,
        metadata_plan.data_identifier,
        media_creation_codec::Geometry::new(
            media_creation_codec::Point::new(options.position().x, options.position().y),
            media_creation_codec::Size::new(0.0, 0.0),
            Some(DEFAULT_DRAWABLE_FLAGS),
            Some(0.0),
        ),
        options.duration_seconds(),
    );
    let output_remaining = nonzero_residual(
        budget.remaining_output(),
        SlideAudioCreationLimitKind::OutputBytes,
    )?;
    let references_remaining = nonzero_residual(
        budget.remaining_references(),
        SlideAudioCreationLimitKind::References,
    )?;
    let fields_remaining = nonzero_residual(
        budget.remaining_wire_fields(),
        SlideAudioCreationLimitKind::WireFields,
    )?;
    let work_remaining = nonzero_residual(
        budget.remaining_work(),
        SlideAudioCreationLimitKind::WireWork,
    )?;
    let allocations_remaining = nonzero_residual(
        budget.remaining_allocations(),
        SlideAudioCreationLimitKind::Allocations,
    )?;
    let encode_options = media_creation_codec::EncodeOptions::for_write(&write)
        .with_max_output_bytes(output_remaining)
        .with_max_references(references_remaining)
        .with_max_work_bytes(work_remaining)
        .with_max_fields(fields_remaining)
        .with_max_allocations(allocations_remaining);
    let encoded = media_creation_codec::encode_media_archive_with_report(&write, encode_options)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let report = encoded.report();
    budget.charge_output(report.output_bytes())?;
    budget.charge_references(report.references())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    let movie_payload = encoded.into_bytes();
    let movie = new_archive_object(
        ids.drawable,
        MOVIE_MESSAGE_TYPE,
        movie_payload,
        &STANDARD_MESSAGE_VERSION,
        &[ids.caption, ids.title, context.style_identifier],
        &[metadata_plan.data_identifier],
        context.archive_limits,
    )?;
    let standin_payload = media_creation_codec::canonical_standin_payload().to_vec();
    let title = new_archive_object(
        ids.title,
        STANDIN_CAPTION_MESSAGE_TYPE,
        standin_payload.clone(),
        &STANDIN_CAPTION_MESSAGE_VERSION,
        &[],
        &[],
        context.archive_limits,
    )?;
    let caption = new_archive_object(
        ids.caption,
        STANDIN_CAPTION_MESSAGE_TYPE,
        standin_payload,
        &STANDIN_CAPTION_MESSAGE_VERSION,
        &[],
        &[],
        context.archive_limits,
    )?;
    Ok([movie, title, caption])
}

/// Encode the native movie drawable and its two caption stand-ins through the
/// same bounded Buffa writer used by audio creation.
///
/// The package transaction owns the identifiers and metadata plan; keeping
/// this writer here makes the media graph construction shared without making
/// the public movie vocabulary depend on native archive types.
pub(super) fn make_movie_objects(
    context: &CreationContext,
    ids: &CreationIds,
    movie_data_identifier: u64,
    poster_data_identifier: u64,
    position: (f32, f32),
    size: (f32, f32),
    natural_size: (f32, f32),
    duration_seconds: f32,
    budget: &mut CreationBudget,
) -> Result<[ArchiveObject; 3], SlideAudioCreationError> {
    if movie_data_identifier == 0 || poster_data_identifier == 0 {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    let write = media_creation_codec::MediaArchiveWrite::movie(
        context.slide_identifier,
        context.style_identifier,
        ids.title,
        ids.caption,
        movie_data_identifier,
        Some(poster_data_identifier),
        media_creation_codec::Geometry::new(
            media_creation_codec::Point::new(position.0, position.1),
            media_creation_codec::Size::new(size.0, size.1),
            Some(DEFAULT_DRAWABLE_FLAGS),
            Some(0.0),
        ),
        duration_seconds,
        media_creation_codec::Size::new(natural_size.0, natural_size.1),
        true,
    );
    let output_remaining = nonzero_residual(
        budget.remaining_output(),
        SlideAudioCreationLimitKind::OutputBytes,
    )?;
    let references_remaining = nonzero_residual(
        budget.remaining_references(),
        SlideAudioCreationLimitKind::References,
    )?;
    let fields_remaining = nonzero_residual(
        budget.remaining_wire_fields(),
        SlideAudioCreationLimitKind::WireFields,
    )?;
    let work_remaining = nonzero_residual(
        budget.remaining_work(),
        SlideAudioCreationLimitKind::WireWork,
    )?;
    let allocations_remaining = nonzero_residual(
        budget.remaining_allocations(),
        SlideAudioCreationLimitKind::Allocations,
    )?;
    let encode_options = media_creation_codec::EncodeOptions::for_write(&write)
        .with_max_output_bytes(output_remaining)
        .with_max_references(references_remaining)
        .with_max_work_bytes(work_remaining)
        .with_max_fields(fields_remaining)
        .with_max_allocations(allocations_remaining);
    let encoded = media_creation_codec::encode_media_archive_with_report(&write, encode_options)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let report = encoded.report();
    budget.charge_output(report.output_bytes())?;
    budget.charge_references(report.references())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    let movie_payload = encoded.into_bytes();
    let movie = new_archive_object(
        ids.drawable,
        MOVIE_MESSAGE_TYPE,
        movie_payload,
        &STANDARD_MESSAGE_VERSION,
        &[ids.caption, ids.title, context.style_identifier],
        &[poster_data_identifier, movie_data_identifier],
        context.archive_limits,
    )?;
    let standin_payload = media_creation_codec::canonical_standin_payload().to_vec();
    let title = new_archive_object(
        ids.title,
        STANDIN_CAPTION_MESSAGE_TYPE,
        standin_payload.clone(),
        &STANDIN_CAPTION_MESSAGE_VERSION,
        &[],
        &[],
        context.archive_limits,
    )?;
    let caption = new_archive_object(
        ids.caption,
        STANDIN_CAPTION_MESSAGE_TYPE,
        standin_payload,
        &STANDIN_CAPTION_MESSAGE_VERSION,
        &[],
        &[],
        context.archive_limits,
    )?;
    Ok([movie, title, caption])
}

pub(super) fn new_archive_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
    limits: ArchiveLimits,
) -> Result<ArchiveObject, SlideAudioCreationError> {
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
        limits,
    )
    .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(0)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    info.versions = versions.to_vec();
    info.object_references = object_references.to_vec();
    info.data_references = data_references.to_vec();
    Ok(object)
}

fn lifecycle_options(
    source_length: usize,
    limits: WireLimits,
    budget: &CreationBudget,
) -> Result<lifecycle_codec::DecodeOptions, SlideAudioCreationError> {
    let output_remaining = nonzero_residual(
        budget.remaining_output(),
        SlideAudioCreationLimitKind::OutputBytes,
    )?;
    let work_remaining = nonzero_residual(
        budget.remaining_work(),
        SlideAudioCreationLimitKind::WireWork,
    )?;
    let fields_remaining = nonzero_residual(
        budget.remaining_wire_fields(),
        SlideAudioCreationLimitKind::WireFields,
    )?;
    let references_remaining = nonzero_residual(
        budget.remaining_references(),
        SlideAudioCreationLimitKind::References,
    )?;
    let input_limit = nonzero_residual(
        limits.max_input_bytes(),
        SlideAudioCreationLimitKind::InputBytes,
    )?;
    let output_limit = nonzero_residual(
        limits.max_output_bytes(),
        SlideAudioCreationLimitKind::OutputBytes,
    )?;
    let work_limit = nonzero_residual(
        limits.max_rewrite_work(),
        SlideAudioCreationLimitKind::WireWork,
    )?;
    let field_limit =
        nonzero_residual(limits.max_fields(), SlideAudioCreationLimitKind::WireFields)?;
    let nesting_limit = nonzero_residual(
        limits.max_nesting(),
        SlideAudioCreationLimitKind::WireNesting,
    )?;
    let budget_nesting =
        nonzero_residual(budget.max_nesting, SlideAudioCreationLimitKind::WireNesting)?;
    let output = source_length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or(SlideAudioCreationError::InvalidSource)?
        .min(output_limit)
        .min(output_remaining);
    let work = source_length
        .checked_mul(256)
        .and_then(|value| value.checked_add(4096))
        .ok_or(SlideAudioCreationError::InvalidSource)?
        .min(work_limit)
        .min(work_remaining);
    let fields = field_limit.min(fields_remaining);
    let references = references_remaining;
    let depth = u32::try_from(nesting_limit.min(budget_nesting))
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    Ok(lifecycle_codec::DecodeOptions::new(
        source_length.min(input_limit).max(1),
        output,
        fields,
        work,
        references,
        depth,
    ))
}

fn nonzero_residual(
    residual: usize,
    kind: SlideAudioCreationLimitKind,
) -> Result<usize, SlideAudioCreationError> {
    if residual == 0 {
        return Err(limit(kind, 1, 0));
    }
    Ok(residual)
}

pub(super) fn rewrite_slide_archive(
    archive: &mut Archive,
    context: &CreationContext,
    ids: &CreationIds,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let object = archive
        .object_mut(context.slide_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let source = object
        .messages
        .get(message_index)
        .ok_or(SlideAudioCreationError::InvalidSource)?
        .data
        .as_slice();
    let options = lifecycle_options(source.len(), wire_limits, budget)?;
    let edit = lifecycle_codec::SlideLifecycleEdit::empty()
        .with_owned_drawables(std::slice::from_ref(&ids.drawable))
        .with_drawables_z_order(std::slice::from_ref(&ids.drawable))
        .with_builds(std::slice::from_ref(&ids.build))
        .with_build_chunks(std::slice::from_ref(&ids.chunk));
    let (payload, report) =
        lifecycle_codec::rewrite_slide_lifecycle_with_report(source, edit, options)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    replace_slide_message_with_refs(
        object,
        message_index,
        payload,
        [ids.drawable],
        [ids.drawable],
        [ids.build],
        [ids.chunk],
        archive_limits,
        wire_limits,
        budget,
    )
}

fn replace_slide_message_with_refs(
    object: &mut ArchiveObject,
    message_index: usize,
    payload: Vec<u8>,
    drawables: [u64; 1],
    z_order: [u64; 1],
    builds: [u64; 1],
    chunks: [u64; 1],
    limits: ArchiveLimits,
    wire_limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let source = object
        .messages
        .get(message_index)
        .ok_or(SlideAudioCreationError::InvalidSource)?
        .data
        .as_slice();
    let decode_options = lifecycle_options(source.len(), wire_limits, budget)?;
    let (before, before_report) =
        lifecycle_codec::decode_slide_lifecycle_with_report(source, decode_options)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(before_report.fields())?;
    budget.charge_work(before_report.work_bytes())?;
    budget.charge_nesting(before_report.max_depth() as usize)?;
    let after_options = lifecycle_options(payload.len(), wire_limits, budget)?;
    let (after, after_report) =
        lifecycle_codec::decode_slide_lifecycle_with_report(&payload, after_options)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_wire_fields(after_report.fields())?;
    budget.charge_work(after_report.work_bytes())?;
    budget.charge_nesting(after_report.max_depth() as usize)?;
    let groups = [
        (
            SLIDE_BUILDS_FIELD,
            collect_lifecycle_reference_ids(before.builds(), budget)?,
            collect_lifecycle_reference_ids(after.builds(), budget)?,
        ),
        (
            SLIDE_OWNED_DRAWABLES_FIELD,
            collect_lifecycle_reference_ids(before.owned_drawables(), budget)?,
            collect_lifecycle_reference_ids(after.owned_drawables(), budget)?,
        ),
        (
            SLIDE_DRAWABLES_Z_ORDER_FIELD,
            collect_lifecycle_reference_ids(before.drawables_z_order(), budget)?,
            collect_lifecycle_reference_ids(after.drawables_z_order(), budget)?,
        ),
        (
            SLIDE_BUILD_CHUNKS_FIELD,
            collect_lifecycle_reference_ids(before.build_chunks(), budget)?,
            collect_lifecycle_reference_ids(after.build_chunks(), budget)?,
        ),
    ];
    let expected = [
        (SLIDE_BUILDS_FIELD, builds.as_slice()),
        (SLIDE_OWNED_DRAWABLES_FIELD, drawables.as_slice()),
        (SLIDE_DRAWABLES_Z_ORDER_FIELD, z_order.as_slice()),
        (SLIDE_BUILD_CHUNKS_FIELD, chunks.as_slice()),
    ];
    for ((number, before, after), (_, appended)) in groups.iter().zip(expected) {
        if after.len() != before.len().saturating_add(appended.len())
            || after.get(before.len()..) != Some(appended)
            || has_zero_or_duplicate_reference(before)
            || has_zero_or_duplicate_reference(after)
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let _ = number;
    }
    let (metadata_capacity, metadata_events) = {
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        for (index, field) in info.field_infos.iter().enumerate() {
            if info.field_infos[..index]
                .iter()
                .any(|previous| previous.path == field.path)
            {
                return Err(SlideAudioCreationError::InvalidSource);
            }
            if has_zero_or_duplicate_reference(&field.object_references) {
                return Err(SlideAudioCreationError::InvalidSource);
            }
        }
        if has_zero_or_duplicate_reference(&info.object_references) {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let mut selected_fields = 0usize;
        let mut field_bytes = 0usize;
        for (number, before, after) in &groups {
            for field in info.field_infos.iter().filter(|field| {
                field.path.as_slice() == [*number] && !field.object_references.is_empty()
            }) {
                if field.object_references != *before {
                    return Err(SlideAudioCreationError::InvalidSource);
                }
                selected_fields = selected_fields
                    .checked_add(1)
                    .ok_or(SlideAudioCreationError::InvalidSource)?;
                field_bytes = field_bytes
                    .checked_add(
                        field
                            .path
                            .path
                            .len()
                            .checked_mul(size_of::<u32>())
                            .ok_or(SlideAudioCreationError::InvalidSource)?,
                    )
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .object_references
                                .len()
                                .checked_mul(size_of::<u64>())?,
                        )
                    })
                    .and_then(|value| value.checked_add(after.len().checked_mul(size_of::<u64>())?))
                    .ok_or(SlideAudioCreationError::InvalidSource)?;
            }
        }
        let appended_identifiers = [drawables[0], z_order[0], builds[0], chunks[0]];
        let aggregate_growth =
            count_new_unique_references(&info.object_references, &appended_identifiers);
        let capacity = info
            .object_references
            .len()
            .checked_add(aggregate_growth)
            .and_then(|count| count.checked_mul(size_of::<u64>()))
            .and_then(|value| {
                value.checked_add(info.object_references.len().checked_mul(size_of::<u64>())?)
            })
            .and_then(|value| value.checked_add(field_bytes))
            .and_then(|value| {
                value.checked_add(selected_fields.checked_mul(size_of::<(
                    usize,
                    Vec<u32>,
                    Vec<u64>,
                    Vec<u64>,
                )>())?)
            })
            .and_then(|value| {
                value.checked_add(
                    selected_fields
                        .checked_mul(size_of::<FieldObjectReferenceTransition<'static>>())?,
                )
            })
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let events = selected_fields
            .checked_mul(3)
            .and_then(|count| count.checked_add(4))
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        (capacity, events)
    };
    budget.charge_allocations(metadata_capacity)?;
    budget.charge_allocation_events(metadata_events)?;
    let (aggregate_before, aggregate_after, field_records) = {
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if groups.iter().any(|(_, before, _)| {
            before
                .iter()
                .any(|identifier| !info.object_references.contains(identifier))
        }) {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let aggregate_before = clone_u64_values(&info.object_references)?;
        let appended_identifiers = [drawables[0], z_order[0], builds[0], chunks[0]];
        let aggregate_growth =
            count_new_unique_references(&aggregate_before, &appended_identifiers);
        let aggregate_after_capacity = aggregate_before
            .len()
            .checked_add(aggregate_growth)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let aggregate_after_bytes = aggregate_after_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let mut aggregate_after = Vec::new();
        aggregate_after
            .try_reserve_exact(aggregate_after_capacity)
            .map_err(|_| SlideAudioCreationError::Allocation {
                amount: aggregate_after_bytes,
            })?;
        aggregate_after.extend_from_slice(&aggregate_before);
        for (_, _, after) in &groups {
            for identifier in after {
                if !aggregate_after.contains(identifier) {
                    if aggregate_after.len() == aggregate_after_capacity {
                        return Err(SlideAudioCreationError::InvalidSource);
                    }
                    aggregate_after.push(*identifier);
                }
            }
        }
        let selected_fields = info
            .field_infos
            .iter()
            .filter(|field| {
                [
                    SLIDE_BUILDS_FIELD,
                    SLIDE_OWNED_DRAWABLES_FIELD,
                    SLIDE_DRAWABLES_Z_ORDER_FIELD,
                    SLIDE_BUILD_CHUNKS_FIELD,
                ]
                .iter()
                .any(|number| {
                    field.path.as_slice() == [*number] && !field.object_references.is_empty()
                })
            })
            .count();
        let field_record_bytes = selected_fields
            .checked_mul(size_of::<(usize, Vec<u32>, Vec<u64>, Vec<u64>)>())
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let mut field_records = Vec::new();
        field_records
            .try_reserve_exact(selected_fields)
            .map_err(|_| SlideAudioCreationError::Allocation {
                amount: field_record_bytes,
            })?;
        for (number, before, after) in &groups {
            for (field_info_index, field) in info.field_infos.iter().enumerate() {
                if field.path.as_slice() != [*number] || field.object_references.is_empty() {
                    continue;
                }
                if field.object_references != *before {
                    return Err(SlideAudioCreationError::InvalidSource);
                }
                if field_records.len() == selected_fields {
                    return Err(SlideAudioCreationError::InvalidSource);
                }
                field_records.push((
                    field_info_index,
                    clone_u32_values(&field.path.path)?,
                    clone_u64_values(&field.object_references)?,
                    clone_u64_values(after)?,
                ));
            }
        }
        (aggregate_before, aggregate_after, field_records)
    };
    let field_transition_bytes = field_records
        .len()
        .checked_mul(size_of::<FieldObjectReferenceTransition<'static>>())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut field_transitions = Vec::new();
    field_transitions
        .try_reserve_exact(field_records.len())
        .map_err(|_| SlideAudioCreationError::Allocation {
            amount: field_transition_bytes,
        })?;
    for (field_info_index, expected_path, before, after) in &field_records {
        field_transitions.push(FieldObjectReferenceTransition {
            field_info_index: *field_info_index,
            expected_path,
            before,
            after,
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: payload,
            },
            ObjectReferenceTransition {
                aggregate_before: &aggregate_before,
                aggregate_after: &aggregate_after,
                fields: &field_transitions,
            },
            limits,
        )
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    Ok(())
}

pub(super) fn slide_build_count(
    archive: &Archive,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<usize, SlideAudioCreationError> {
    let object = archive
        .object(context.slide_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let payload = unique_payload(object, SLIDE_MESSAGE_TYPE)?;
    let identifiers = references_in_field(
        payload,
        SLIDE_BUILD_CHUNKS_FIELD,
        context.wire_limits,
        budget,
    )?;
    Ok(identifiers.len())
}

/// Upper-bound the heap retained by `Archive::clone` before invoking it.
/// Serialized bytes cover payload/header buffers; the additional structural
/// term covers cloned vectors and their element storage.
pub(super) fn archive_clone_allocation_bound(
    archive: &Archive,
    limits: ArchiveLimits,
) -> Result<usize, SlideAudioCreationError> {
    let mut amount = archive
        .encoded_len_with_limits(limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    amount = amount
        .checked_add(size_of::<Archive>())
        .and_then(|value| {
            value.checked_add(
                archive
                    .objects
                    .len()
                    .checked_mul(size_of::<ArchiveObject>())?,
            )
        })
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    for object in &archive.objects {
        // ArchiveObject keeps raw and canonical header copies private in the
        // neutral archive layer.  `encoded_len_with_limits` accounts for the
        // serialized framing, so reserve two additional header-sized buffers
        // for the retained source-preserving pair without exposing either
        // buffer or changing the archive's preservation behavior.
        let header_length = usize::try_from(object.header_length)
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        amount = amount
            .checked_add(
                header_length
                    .checked_mul(2)
                    .ok_or(SlideAudioCreationError::InvalidSource)?,
            )
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        amount = amount
            .checked_add(size_of_val(&object.archive_info))
            .and_then(|value| {
                value.checked_add(object.messages.len().checked_mul(size_of::<RawMessage>())?)
            })
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        for message in &object.messages {
            amount = amount
                .checked_add(message.data.len())
                .ok_or(SlideAudioCreationError::InvalidSource)?;
        }
        for info in &object.archive_info.message_infos {
            amount = amount
                .checked_add(size_of_val(info))
                .and_then(|value| {
                    value.checked_add(info.versions.len().checked_mul(size_of::<u32>())?)
                })
                .and_then(|value| {
                    value.checked_add(
                        info.field_infos
                            .len()
                            .checked_mul(size_of::<litchi_iwa_core::FieldInfo>())?,
                    )
                })
                .and_then(|value| {
                    value.checked_add(info.object_references.len().checked_mul(size_of::<u64>())?)
                })
                .and_then(|value| {
                    value.checked_add(info.data_references.len().checked_mul(size_of::<u64>())?)
                })
                .and_then(|value| {
                    value.checked_add(
                        info.diff_merge_version
                            .len()
                            .checked_mul(size_of::<u32>())?,
                    )
                })
                .and_then(|value| {
                    value.checked_add(info.fields_to_remove.iter().try_fold(
                        0usize,
                        |total, path| {
                            total.checked_add(
                                size_of_val(path) + path.path.len().checked_mul(size_of::<u32>())?,
                            )
                        },
                    )?)
                })
                .and_then(|value| {
                    value.checked_add(info.diff_read_version.len().checked_mul(size_of::<u32>())?)
                })
                .ok_or(SlideAudioCreationError::InvalidSource)?;
            for field in &info.field_infos {
                amount = amount
                    .checked_add(
                        field
                            .path
                            .path
                            .len()
                            .checked_mul(size_of::<u32>())
                            .ok_or(SlideAudioCreationError::InvalidSource)?,
                    )
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .object_references
                                .len()
                                .checked_mul(size_of::<u64>())?,
                        )
                    })
                    .and_then(|value| {
                        value
                            .checked_add(field.data_references.len().checked_mul(size_of::<u64>())?)
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .known_field_version
                                .len()
                                .checked_mul(size_of::<u32>())?,
                        )
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .known_field_feature_identifier
                                .as_ref()
                                .map_or(0, |value| value.len()),
                        )
                    })
                    .ok_or(SlideAudioCreationError::InvalidSource)?;
            }
        }
    }
    Ok(amount)
}

/// Return a source-owned node-cache component edit when the slide node lives
/// outside the selected slide component.  The `None` case means the selected
/// slide archive was edited in place.
pub(super) fn prepare_node_cache_edit(
    source: &Package,
    context: &CreationContext,
    slide_archive: &mut Archive,
    event_count: usize,
    budget: &mut CreationBudget,
) -> Result<Option<(Box<str>, Archive)>, SlideAudioCreationError> {
    let (component_name, node_object) = source
        .object_with_component(context.slide_node_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let message_index = unique_message_index(node_object, SLIDE_NODE_MESSAGE_TYPE)?;
    let source_payload = node_object
        .messages
        .get(message_index)
        .ok_or(SlideAudioCreationError::InvalidSource)?
        .data
        .as_slice();
    let rewritten = playback_build::rewrite_node_cache(
        source_payload,
        event_count,
        context.wire_limits,
        budget,
    )?;
    if component_name == context.slide_component.as_ref() {
        let node = slide_archive
            .object_mut(context.slide_node_identifier)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        node.replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data: rewritten,
            },
            context.archive_limits,
        )
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        return Ok(None);
    }
    let component = source
        .state
        .source
        .components()
        .get(component_name)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let clone_bound = archive_clone_allocation_bound(component.archive(), context.archive_limits)?;
    budget.charge_allocations(clone_bound)?;
    budget.charge_allocation_events(1)?;
    let mut archive = component.archive().clone();
    let node = archive
        .object_mut(context.slide_node_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    node.replace_message_preserving_header_with_limits(
        message_index,
        RawMessage {
            type_: SLIDE_NODE_MESSAGE_TYPE,
            data: rewritten,
        },
        context.archive_limits,
    )
    .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(component_name.len())?;
    Ok(Some((component_name.into(), archive)))
}

fn unique_message_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<usize, SlideAudioCreationError> {
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if index.replace(candidate).is_some() {
            return Err(SlideAudioCreationError::InvalidSource);
        }
    }
    index.ok_or(SlideAudioCreationError::InvalidSource)
}

pub(super) fn serialize_component(
    archive: &Archive,
    archive_limits: ArchiveLimits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut CreationBudget,
) -> Result<Vec<u8>, SlideAudioCreationError> {
    let expected = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(expected)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    if bytes.len() != expected {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    let compressed_bound = SnappyStream::maximum_compressed_len(bytes.len())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_allocations(compressed_bound)?;
    let compressed =
        SnappyStream::compress(&bytes).map_err(|_| SlideAudioCreationError::InvalidSource)?;
    if compressed.len() > snappy_limits.max_compressed_stream() {
        return Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::EntryBytes,
            observed: u64::try_from(compressed.len()).unwrap_or(u64::MAX),
            maximum: u64::try_from(snappy_limits.max_compressed_stream()).unwrap_or(u64::MAX),
        });
    }
    budget.charge_output(compressed.len())?;
    Ok(compressed)
}

pub(super) fn normalize_data_path(name: &str) -> Box<str> {
    if name.starts_with(DATA_PREFIX) {
        name.into()
    } else {
        let mut path = String::with_capacity(DATA_PREFIX.len() + name.len());
        path.push_str(DATA_PREFIX);
        path.push_str(name);
        path.into_boxed_str()
    }
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    context: &CreationContext,
    options: crate::slide::audio::Options,
    data: &[u8],
    digest: [u8; 20],
    source_media_count: usize,
    target_media_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    candidate
        .validate()
        .map_err(|_| SlideAudioCreationError::Verification)?;
    let source_slide = source
        .slides()
        .map_err(|_| SlideAudioCreationError::Verification)?
        .get(context.slide_position.get())
        .ok_or(SlideAudioCreationError::Verification)?;
    let candidate_slide = candidate
        .slides()
        .map_err(|_| SlideAudioCreationError::Verification)?
        .get(context.slide_position.get())
        .ok_or(SlideAudioCreationError::Verification)?;
    let source_movies = source_slide.movies();
    let candidate_movies = candidate_slide.movies();
    budget.charge_references(source_movies.len().saturating_add(candidate_movies.len()))?;
    if source_movies.len() != source_media_count
        || candidate_movies.len() != target_media_count
        || candidate_movies.get(..source_movies.len()) != Some(source_movies)
    {
        return Err(SlideAudioCreationError::Verification);
    }
    let created = candidate_movies
        .get(source_media_count)
        .copied()
        .ok_or(SlideAudioCreationError::Verification)?;
    if created.kind() != MovieKind::Audio
        || !created.position().is_some_and(|position| {
            (position.x, position.y) == (options.position().x, options.position().y)
        })
        || created.duration() != Some(options.duration())
    {
        return Err(SlideAudioCreationError::Verification);
    }
    let content = data::read_content(
        candidate,
        context.slide_position,
        Position::new(source_media_count),
        budget,
    )?;
    budget.charge_work(
        content
            .len()
            .checked_mul(2)
            .ok_or(SlideAudioCreationError::Verification)?,
    )?;
    if content != data || Sha1::digest(content).as_slice() != digest {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn verify_created_audio_semantics(
    candidate: &Package,
    patch: &SlideAudioCreationPatch,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let slide = candidate
        .slides()
        .map_err(|_| SlideAudioCreationError::PatchConflict)?
        .get(context.slide_position.get())
        .ok_or(SlideAudioCreationError::PatchConflict)?;
    let movies = slide.movies();
    let created = movies
        .get(patch.movie_position.get())
        .copied()
        .ok_or(SlideAudioCreationError::PatchConflict)?;
    if created.kind() != MovieKind::Audio
        || !created.position().is_some_and(|position| {
            (position.x, position.y) == (patch.options.position().x, patch.options.position().y)
        })
        || created.duration() != Some(patch.options.duration())
    {
        return Err(SlideAudioCreationError::PatchConflict);
    }
    budget.charge_references(movies.len())?;
    let content = data::read_content(
        candidate,
        context.slide_position,
        patch.movie_position,
        budget,
    )?;
    budget.charge_work(content.len())?;
    if content.len() != patch.data_len || Sha1::digest(content).as_slice() != patch.data_digest {
        return Err(SlideAudioCreationError::PatchConflict);
    }
    Ok(())
}

fn to_usize(value: u64) -> Result<usize, SlideAudioCreationError> {
    usize::try_from(value).map_err(|_| SlideAudioCreationError::InvalidSource)
}

fn limit(
    kind: SlideAudioCreationLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideAudioCreationError {
    SlideAudioCreationError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

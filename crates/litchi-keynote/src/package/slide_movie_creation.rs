//! Bounded, selector-first creation of one file-backed Keynote movie.
//!
//! The transaction deliberately reuses the audio creation engine's source
//! admission, graph context, identity allocator, slide reference transition,
//! node-cache update, exact reassembly, and candidate publication helpers.
//! Only the media-specific metadata plan, movie payload, and witness differ.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The transaction keeps admission, staging, and publication phases together."
)]

use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryInsertion, ExactArtifacts};
use litchi_iwa_common::media::Type as MediaType;
use sha1::{Digest as _, Sha1};

use super::Package;
use super::slide_audio_creation::{
    CreationBudget, allocate_ids, count_slide_movies, make_movie_objects, normalize_data_path,
    physical_catalog, plan_movie_metadata, prepare_movie_builds, read_movie_content_and_poster,
    resolve_context, verify_movie_graph,
};
use crate::slide::audio::creation::{SlideAudioCreationError, SlideAudioCreationLimitKind};
use crate::slide::movie::Options;
use crate::slide::movie::creation::{
    SlideMovieCreationCommit, SlideMovieCreationDiagnostics, SlideMovieCreationError,
    SlideMovieCreationLimitKind, SlideMovieCreationPatch,
};
use crate::{MovieKind, SlideSelector};

impl Package {
    /// Create one independently positioned, file-backed movie on a slide.
    ///
    /// Both movie and poster bytes are borrowed until the exact candidate has
    /// been assembled and reopened. The source package remains unchanged and
    /// the returned patch authorizes replay only for that exact source.
    pub fn add_slide_movie<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        preferred_movie_filename: &str,
        movie_data: &[u8],
        preferred_poster_filename: &str,
        poster_data: &[u8],
        options: Options,
    ) -> Result<SlideMovieCreationCommit, SlideMovieCreationError> {
        validate_input(
            self,
            preferred_movie_filename,
            movie_data,
            MediaType::Video,
            InputKind::Movie,
        )?;
        validate_input(
            self,
            preferred_poster_filename,
            poster_data,
            MediaType::Image,
            InputKind::Poster,
        )?;

        let mut budget = CreationBudget::for_package(self).map_err(map_error)?;
        let catalog = physical_catalog(self).map_err(map_error)?;
        budget
            .charge_entries(catalog.package().len())
            .map_err(map_error)?;
        for entry in catalog.package().iter() {
            budget
                .charge_entry_bytes(entry.data().len())
                .and_then(|_| budget.charge_total(entry.data().len()))
                .map_err(map_error)?;
        }
        budget
            .charge_media_bytes(movie_data.len())
            .and_then(|_| budget.charge_media_bytes(poster_data.len()))
            .map_err(map_error)?;

        let context = resolve_context(self, slide.into(), &mut budget).map_err(map_error)?;
        let source_media_count =
            count_slide_movies(self, &context, &mut budget).map_err(map_error)?;
        let target_media_count = source_media_count
            .checked_add(1)
            .ok_or(SlideMovieCreationError::InvalidSource)?;
        let ids = allocate_ids(self, &mut budget).map_err(map_error)?;
        ids.validate().map_err(map_error)?;

        let metadata_plan = plan_movie_metadata(
            self,
            &context,
            &ids,
            preferred_movie_filename,
            movie_data,
            preferred_poster_filename,
            poster_data,
            &mut budget,
        )
        .map_err(map_error)?;
        let poster_plan = &metadata_plan.poster;

        let media_objects = make_movie_objects(
            &context,
            &ids,
            metadata_plan.data_identifier,
            poster_plan.data_identifier,
            (options.position().x, options.position().y),
            (options.size().width, options.size().height),
            (options.natural_size().width, options.natural_size().height),
            options.duration_seconds(),
            &mut budget,
        )
        .map_err(map_error)?;
        let builds = prepare_movie_builds(&ids, &context, &mut budget).map_err(map_error)?;

        let movie_data_path = metadata_plan
            .data_entry_name
            .as_deref()
            .map(normalize_data_path);
        let poster_data_path = poster_plan
            .data_entry_name
            .as_deref()
            .map(normalize_data_path);
        // Keep the bounded two-asset insertion descriptor on the stack. The
        // publication helper charges the descriptor view, while this avoids
        // an uncharged heap allocation during staging.
        let mut insertion_slots = [EntryInsertion::new("", &[]); 2];
        let mut insertion_count = 0;
        if let Some(name) = movie_data_path.as_deref() {
            insertion_slots[insertion_count] = EntryInsertion::new(name, movie_data);
            insertion_count += 1;
        }
        if let Some(name) = poster_data_path.as_deref() {
            insertion_slots[insertion_count] = EntryInsertion::new(name, poster_data);
            insertion_count += 1;
        }
        let insertions = &insertion_slots[..insertion_count];
        let publication = super::slide_audio_creation::publish_candidate(
            self,
            catalog,
            &context,
            &ids,
            media_objects,
            builds,
            &metadata_plan.compressed,
            insertions,
            &mut budget,
        )
        .map_err(map_error)?;
        let target = publication.target;
        let event_count = publication.event_count;
        let candidate = Package::from_source_with_options(Arc::clone(&target), self.state.options)
            .map_err(|_| SlideMovieCreationError::Verification)?;
        verify_candidate(
            self,
            &candidate,
            &context,
            options,
            movie_data,
            poster_data,
            metadata_plan.digest,
            poster_plan.digest,
            source_media_count,
            target_media_count,
            &mut budget,
        )?;
        verify_movie_graph(
            &candidate,
            &context,
            &ids,
            metadata_plan.data_identifier,
            poster_plan.data_identifier,
            [options.position().x, options.position().y],
            [options.size().width, options.size().height],
            options.duration_seconds(),
            [options.natural_size().width, options.natural_size().height],
            event_count,
            &mut budget,
        )
        .map_err(map_error)?;

        let patch = SlideMovieCreationPatch {
            artifacts: ExactArtifacts::new(catalog.shared_source(), target),
            slide_position: context.slide_position,
            movie_position: Position::new(source_media_count),
            source_media_count,
            target_media_count,
            options,
            movie_data_digest: metadata_plan.digest,
            movie_data_len: movie_data.len(),
            poster_data_digest: poster_plan.digest,
            poster_data_len: poster_data.len(),
            target_contains_created_movie: true,
            created_objects: 5,
            removed_objects: 0,
            created_data: metadata_plan
                .created_data
                .checked_add(poster_plan.created_data)
                .ok_or(SlideMovieCreationError::InvalidSource)?,
            removed_data: 0,
            touched_members: publication.touched_members,
            deleted_previews: publication.deleted_previews,
            restored_previews: 0,
        };
        let diagnostics = SlideMovieCreationDiagnostics::for_patch(&patch);
        Ok(SlideMovieCreationCommit {
            package: candidate,
            patch,
            diagnostics,
        })
    }

    /// Apply a creation patch only to its exact retained source snapshot.
    pub fn apply_slide_movie_creation(
        &self,
        patch: &SlideMovieCreationPatch,
    ) -> Result<SlideMovieCreationCommit, SlideMovieCreationError> {
        let catalog = physical_catalog(self).map_err(map_error)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideMovieCreationError::PatchConflict);
        }
        let mut budget = CreationBudget::for_package(self).map_err(map_error)?;
        budget
            .charge_output(patch.artifacts.target().len())
            .and_then(|_| budget.charge_allocations(patch.artifacts.target().len()))
            .map_err(map_error)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(|_| SlideMovieCreationError::PatchConflict)?;
        candidate
            .validate()
            .map_err(|_| SlideMovieCreationError::PatchConflict)?;
        let context = resolve_context(
            &candidate,
            SlideSelector::position(patch.slide_position),
            &mut budget,
        )
        .map_err(map_error)?;
        let source_media_count =
            count_slide_movies(&candidate, &context, &mut budget).map_err(map_error)?;
        if source_media_count != patch.target_media_count {
            return Err(SlideMovieCreationError::PatchConflict);
        }
        if patch.target_contains_created_movie {
            verify_created_movie_semantics(&candidate, patch, &context, &mut budget)?;
        }
        Ok(SlideMovieCreationCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideMovieCreationDiagnostics::for_patch(patch),
        })
    }
}

#[derive(Debug, Clone, Copy)]
enum InputKind {
    Movie,
    Poster,
}

fn validate_input(
    package: &Package,
    filename: &str,
    bytes: &[u8],
    expected: MediaType,
    kind: InputKind,
) -> Result<(), SlideMovieCreationError> {
    super::slide_audio_creation::validate_media_input(package, filename, bytes, expected).map_err(
        |error| match error {
            SlideAudioCreationError::InvalidFilename => match kind {
                InputKind::Movie => SlideMovieCreationError::InvalidMovieFilename,
                InputKind::Poster => SlideMovieCreationError::InvalidPosterFilename,
            },
            SlideAudioCreationError::UnsupportedAudio => match kind {
                InputKind::Movie => SlideMovieCreationError::UnsupportedMovie,
                InputKind::Poster => SlideMovieCreationError::UnsupportedPoster,
            },
            other => map_error(other),
        },
    )
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    context: &super::slide_audio_creation::CreationContext,
    options: Options,
    movie_data: &[u8],
    poster_data: &[u8],
    movie_digest: [u8; 20],
    poster_digest: [u8; 20],
    source_media_count: usize,
    target_media_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideMovieCreationError> {
    candidate
        .validate()
        .map_err(|_| SlideMovieCreationError::Verification)?;
    let source_slide = source
        .slides()
        .map_err(|_| SlideMovieCreationError::Verification)?
        .get(context.slide_position.get())
        .ok_or(SlideMovieCreationError::Verification)?;
    let candidate_slide = candidate
        .slides()
        .map_err(|_| SlideMovieCreationError::Verification)?
        .get(context.slide_position.get())
        .ok_or(SlideMovieCreationError::Verification)?;
    let source_movies = source_slide.movies();
    let candidate_movies = candidate_slide.movies();
    budget
        .charge_references(source_movies.len().saturating_add(candidate_movies.len()))
        .map_err(map_error)?;
    if source_movies.len() != source_media_count
        || candidate_movies.len() != target_media_count
        || candidate_movies.get(..source_movies.len()) != Some(source_movies)
    {
        return Err(SlideMovieCreationError::Verification);
    }
    let created = candidate_movies
        .get(source_media_count)
        .ok_or(SlideMovieCreationError::Verification)?;
    if created.kind() != MovieKind::File
        || created.position().is_none_or(|position| {
            (position.x, position.y) != (options.position().x, options.position().y)
        })
        || created.size().is_none_or(|size| {
            (size.width, size.height) != (options.size().width, options.size().height)
        })
        || created.natural_size().is_none_or(|size| {
            (size.width, size.height)
                != (options.natural_size().width, options.natural_size().height)
        })
        || created.original_size().is_none_or(|size| {
            (size.width, size.height)
                != (options.natural_size().width, options.natural_size().height)
        })
        || created.duration() != Some(options.duration())
    {
        return Err(SlideMovieCreationError::Verification);
    }
    let (content, poster) = read_movie_content_and_poster(
        candidate,
        context.slide_position,
        Position::new(source_media_count),
        budget,
    )
    .map_err(map_error)?;
    budget
        .charge_work(
            movie_data
                .len()
                .checked_mul(2)
                .ok_or(SlideMovieCreationError::Verification)?,
        )
        .map_err(map_error)?;
    if content != movie_data || Sha1::digest(content).as_slice() != movie_digest {
        return Err(SlideMovieCreationError::Verification);
    }
    budget
        .charge_work(
            poster_data
                .len()
                .checked_mul(2)
                .ok_or(SlideMovieCreationError::Verification)?,
        )
        .map_err(map_error)?;
    if poster != poster_data || Sha1::digest(poster).as_slice() != poster_digest {
        return Err(SlideMovieCreationError::Verification);
    }
    Ok(())
}

fn verify_created_movie_semantics(
    candidate: &Package,
    patch: &SlideMovieCreationPatch,
    context: &super::slide_audio_creation::CreationContext,
    budget: &mut CreationBudget,
) -> Result<(), SlideMovieCreationError> {
    let slide = candidate
        .slides()
        .map_err(|_| SlideMovieCreationError::PatchConflict)?
        .get(context.slide_position.get())
        .ok_or(SlideMovieCreationError::PatchConflict)?;
    let movies = slide.movies();
    let created = movies
        .get(patch.movie_position.get())
        .ok_or(SlideMovieCreationError::PatchConflict)?;
    if created.kind() != MovieKind::File
        || created.position().is_none_or(|position| {
            (position.x, position.y) != (patch.options.position().x, patch.options.position().y)
        })
        || created.size().is_none_or(|size| {
            (size.width, size.height) != (patch.options.size().width, patch.options.size().height)
        })
        || created.natural_size().is_none_or(|size| {
            (size.width, size.height)
                != (
                    patch.options.natural_size().width,
                    patch.options.natural_size().height,
                )
        })
        || created.original_size().is_none_or(|size| {
            (size.width, size.height)
                != (
                    patch.options.natural_size().width,
                    patch.options.natural_size().height,
                )
        })
        || created.duration() != Some(patch.options.duration())
    {
        return Err(SlideMovieCreationError::PatchConflict);
    }
    budget.charge_references(movies.len()).map_err(map_error)?;
    let (content, poster) = read_movie_content_and_poster(
        candidate,
        context.slide_position,
        patch.movie_position,
        budget,
    )
    .map_err(map_error)?;
    budget
        .charge_work(
            content
                .len()
                .checked_mul(2)
                .ok_or(SlideMovieCreationError::PatchConflict)?,
        )
        .map_err(map_error)?;
    budget
        .charge_work(
            poster
                .len()
                .checked_mul(2)
                .ok_or(SlideMovieCreationError::PatchConflict)?,
        )
        .map_err(map_error)?;
    if content.len() != patch.movie_data_len
        || Sha1::digest(content).as_slice() != patch.movie_data_digest
        || poster.len() != patch.poster_data_len
        || Sha1::digest(poster).as_slice() != patch.poster_data_digest
    {
        return Err(SlideMovieCreationError::PatchConflict);
    }
    Ok(())
}

fn map_error(error: SlideAudioCreationError) -> SlideMovieCreationError {
    match error {
        SlideAudioCreationError::UnsupportedSource => SlideMovieCreationError::UnsupportedSource,
        SlideAudioCreationError::EmptySlideName => SlideMovieCreationError::EmptySlideName,
        SlideAudioCreationError::SlideNameNotFound => SlideMovieCreationError::SlideNameNotFound,
        SlideAudioCreationError::AmbiguousSelector => SlideMovieCreationError::AmbiguousSelector,
        SlideAudioCreationError::SlidePositionNotFound { position } => {
            SlideMovieCreationError::SlidePositionNotFound { position }
        },
        SlideAudioCreationError::InvalidSource
        | SlideAudioCreationError::InvalidFilename
        | SlideAudioCreationError::UnsupportedAudio => SlideMovieCreationError::InvalidSource,
        SlideAudioCreationError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideMovieCreationError::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        SlideAudioCreationError::Allocation { amount } => {
            SlideMovieCreationError::Allocation { amount }
        },
        SlideAudioCreationError::Verification => SlideMovieCreationError::Verification,
        SlideAudioCreationError::PatchConflict => SlideMovieCreationError::PatchConflict,
    }
}

fn map_limit_kind(kind: SlideAudioCreationLimitKind) -> SlideMovieCreationLimitKind {
    match kind {
        SlideAudioCreationLimitKind::InputBytes => SlideMovieCreationLimitKind::InputBytes,
        SlideAudioCreationLimitKind::OutputBytes => SlideMovieCreationLimitKind::OutputBytes,
        SlideAudioCreationLimitKind::Entries => SlideMovieCreationLimitKind::Entries,
        SlideAudioCreationLimitKind::Objects => SlideMovieCreationLimitKind::Objects,
        SlideAudioCreationLimitKind::EntryBytes => SlideMovieCreationLimitKind::EntryBytes,
        SlideAudioCreationLimitKind::TotalBytes => SlideMovieCreationLimitKind::TotalBytes,
        SlideAudioCreationLimitKind::Slides => SlideMovieCreationLimitKind::Slides,
        SlideAudioCreationLimitKind::References => SlideMovieCreationLimitKind::References,
        SlideAudioCreationLimitKind::MediaBytes => SlideMovieCreationLimitKind::MediaBytes,
        SlideAudioCreationLimitKind::WireFields => SlideMovieCreationLimitKind::WireFields,
        SlideAudioCreationLimitKind::WireNesting => SlideMovieCreationLimitKind::WireNesting,
        SlideAudioCreationLimitKind::WireWork => SlideMovieCreationLimitKind::WireWork,
        SlideAudioCreationLimitKind::Allocations => SlideMovieCreationLimitKind::Allocations,
    }
}

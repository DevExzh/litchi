//! Exact-source, selector-first title and caption transactions of ordinary
//! slide images (record 0372; a port of the migration host's
//! `set_slide_image_title`, `remove_slide_image_title`,
//! `set_slide_image_caption`, `remove_slide_image_caption` and
//! `slide_image_title_caption`).
//!
//! Image titles and captions use the native CaptionInfo/text-storage graph
//! of movie titles and captions, reached through the `TSD.ImageArchive`
//! drawable. Existing storage text, canonical graph creation and
//! fresh-stand-in removal share the movie-caption owner's bounded physical
//! lifecycle seam.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary deliberately redacts lower-layer failures."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ImageSelector, SlideSelector};

const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Which of an image's two text edges a transaction edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Text {
    /// The image title (`TSD.DrawableArchive.title`).
    Title,
    /// The image caption (`TSD.DrawableArchive.caption`).
    Caption,
}

impl Text {
    const fn movie_kind(self) -> super::slide_movie_caption::MovieTextKind {
        match self {
            Self::Title => super::slide_movie_caption::MovieTextKind::Title,
            Self::Caption => super::slide_movie_caption::MovieTextKind::Caption,
        }
    }

    const fn edge_kind(self, owner: Owner) -> super::slide_chart_caption::CaptionEdgeKind {
        use super::slide_chart_caption::CaptionEdgeKind;
        match (owner, self) {
            (Owner::Image, Self::Title) => CaptionEdgeKind::ImageTitle,
            (Owner::Image, Self::Caption) => CaptionEdgeKind::Image,
            (Owner::Shape, Self::Title) => CaptionEdgeKind::ShapeTitle,
            (Owner::Shape, Self::Caption) => CaptionEdgeKind::Shape,
        }
    }
}

/// The drawable whose title or caption a transaction edits: an ordinary
/// image, or an ordinary shape or line (record 0372), addressed by its
/// position among the slide's drawables of that native type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Owner {
    Image,
    Shape,
}

impl Owner {
    const fn captioned(self) -> super::slide_movie_caption::CaptionedDrawable {
        match self {
            Self::Image => super::slide_movie_caption::CaptionedDrawable::Image,
            Self::Shape => super::slide_movie_caption::CaptionedDrawable::Shape,
        }
    }
}

/// A finite resource governed while an image title or caption transaction
/// is prepared or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    WireBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    TextStorages,
    TextFragments,
    TextBytes,
    WireFields,
    WireNesting,
    WireWork,
    ImageTextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::ImageTextBytes => "image title or caption bytes",
        })
    }
}

/// A content-redacted failure raised by an image title or caption
/// transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the Keynote image-text source cannot be edited safely")]
    InvalidSource,
    #[error("Keynote image-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote image-text transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote image title or caption failed semantic verification")]
    Verification,
    #[error("the Keynote image-text patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("this Keynote source does not support physical image-text edits")]
    UnsupportedSource,
    #[error("the requested Keynote image-text graph operation is unsupported")]
    UnsupportedDependency,
    #[error("the Keynote image-text selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no image at position {position:?}")]
    ImagePositionNotFound { position: Position },
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } | Self::UnsupportedDependency { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::AmbiguousSelector { .. } | Self::EmptySlideName { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::ImagePositionNotFound { .. } => RefusalKind::NotFound,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One image title or caption staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    owner: Owner,
    text: Text,
    selection: super::slide_movie_caption::MovieCaptionSelection,
    after: Option<String>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide_position", &self.selection.slide_position)
            .field("image_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new(
        source: &'a Package,
        owner: Owner,
        slide_selector: SlideSelector<'_>,
        position: Position,
        text: Text,
    ) -> Result<Self, Error> {
        let selection = select_text(source, owner, slide_selector, position, text, true)?;
        let after = selection.text.as_deref().map(copy_text).transpose()?;
        Ok(Self {
            source,
            owner,
            text,
            selection,
            after,
        })
    }

    /// The selected slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// The selected image position within its slide.
    #[must_use]
    pub const fn image_position(&self) -> Position {
        self.selection.movie_position
    }

    /// The text observed when the edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// The text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage whole-text replacement, or creation of a native title or
    /// caption graph when the image has none.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for text beyond the transaction limit.
    pub fn set(mut self, text: impl AsRef<str>) -> Result<Self, Error> {
        self.after = Some(copy_text(text.as_ref())?);
        Ok(self)
    }

    /// Stage removal: native iWork keeps the prior graph for undo and
    /// attaches a fresh empty stand-in.
    ///
    /// # Errors
    ///
    /// This staging step does not fail; the `Result` keeps the edit chain
    /// uniform.
    pub fn clear(mut self) -> Result<Self, Error> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// # Errors
    ///
    /// Returns typed refusals, limits and [`Error::Verification`].
    pub fn commit(self) -> Result<Commit, Error> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let current = select_text(
            self.source,
            self.owner,
            SlideSelector::position(self.selection.slide_position),
            self.selection.movie_position,
            self.text,
            true,
        )?;
        if !current.same_identity(&self.selection) || current.text != self.selection.text {
            return Err(Error::InvalidSource);
        }
        if self.selection.text == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    owner: self.owner,
                    text: self.text,
                    selection: self.selection,
                    target_selection: current,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        self.source.validate().map_err(map_read_error)?;
        if self
            .selection
            .text
            .as_deref()
            .is_some_and(super::slide_movie_caption::contains_dependent_marker)
            || self
                .after
                .as_deref()
                .is_some_and(super::slide_movie_caption::contains_dependent_marker)
        {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self.source)
            .map_err(map_chart_error)?;
        budget
            .charge_catalog_scan(self.source)
            .map_err(map_chart_error)?;
        budget
            .charge_selection_scan(self.source, 1)
            .map_err(map_chart_error)?;
        // A new graph references the theme's stylesheet and caption
        // paragraph style; register them on the slide component first.
        let registered = if self.selection.storage_identifier.is_none() && self.after.is_some() {
            register_dependencies(self.source, &self.selection)?
        } else {
            None
        };
        let base = registered.as_ref().unwrap_or(self.source);
        let (package, touched_components, deleted_previews) = rewrite_text(
            base,
            &self.selection,
            self.text.edge_kind(self.owner),
            self.after.as_deref(),
            &mut budget,
        )?;
        budget
            .charge_selection_scan(&package, 1)
            .map_err(map_chart_error)?;
        let candidate = select_text(
            &package,
            self.owner,
            SlideSelector::position(self.selection.slide_position),
            self.selection.movie_position,
            self.text,
            true,
        )?;
        if !candidate.same_movie_identity(&self.selection) || candidate.text != self.after {
            return Err(Error::Verification);
        }
        budget
            .charge_validation_scan(&package, 1)
            .map_err(map_chart_error)?;
        verify_text_transition(
            base,
            &package,
            &self.selection,
            &candidate,
            true,
            &mut budget,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        budget
            .charge_exact_artifacts(source_bytes.len(), target.len())
            .map_err(map_chart_error)?;
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                owner: self.owner,
                text: self.text,
                selection: self.selection,
                target_selection: candidate,
                after: self.after,
                touched_components,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: Diagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// An exact-source-checked reversible image title or caption patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    owner: Owner,
    text: Text,
    selection: super::slide_movie_caption::MovieCaptionSelection,
    target_selection: super::slide_movie_caption::MovieCaptionSelection,
    after: Option<String>,
    touched_components: usize,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.selection.slide_position)
            .field("image_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The selected slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// The selected image position within its slide.
    #[must_use]
    pub const fn image_position(&self) -> Position {
        self.selection.movie_position
    }

    /// The text observed when the edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// The text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// The edited text edge.
    #[must_use]
    pub const fn text(&self) -> Text {
        self.text
    }

    /// The source package fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// The target package fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Whether source and target are identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.selection == self.target_selection && self.artifacts.is_byte_noop()
    }

    /// The exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            owner: self.owner,
            text: self.text,
            selection: self.target_selection.clone(),
            target_selection: self.selection.clone(),
            after: self.selection.text.clone(),
            touched_components: self.touched_components,
            deleted_previews: 0,
            target_requires_invalidated_previews: self.touched_components != 0
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one image title or caption commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many components the commit rewrote.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// How many root previews the commit deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was fully reparsed.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable image title or caption
/// transaction.
#[must_use = "a Keynote image-text commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The edited package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the edited package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// The reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the commit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the title or caption of one ordinary slide image (host
    /// `slide_image_title_caption`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and refusals.
    pub fn slide_image_text<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        image_selector: impl Into<ImageSelector>,
        text: Text,
    ) -> Result<Option<String>, Error> {
        Ok(select_text(
            self,
            Owner::Image,
            slide_selector.into(),
            image_selector.into().as_position(),
            text,
            true,
        )?
        .text)
    }

    /// Start an exact immutable edit of one image's title or caption (host
    /// `set_slide_image_title`, `remove_slide_image_title`,
    /// `set_slide_image_caption`, `remove_slide_image_caption`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and refusals.
    pub fn edit_slide_image_text<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        image_selector: impl Into<ImageSelector>,
        text: Text,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(
            self,
            Owner::Image,
            slide_selector.into(),
            image_selector.into().as_position(),
            text,
        )
    }

    /// Read the title or caption of one ordinary shape or line (host
    /// `slide_shape_title_caption`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and refusals; a text box is refused
    /// ([`Refusal::UnsupportedDependency`]).
    pub fn slide_shape_text<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<crate::DrawableSelector>,
        text: Text,
    ) -> Result<Option<String>, Error> {
        let slide_selector = slide_selector.into();
        let position = shape_position(self, slide_selector, drawable_selector.into())?;
        Ok(select_text(self, Owner::Shape, slide_selector, position, text, true)?.text)
    }

    /// Start an exact immutable edit of one ordinary shape's or line's
    /// title or caption (host `set_slide_shape_title`,
    /// `remove_slide_shape_title`, `set_slide_shape_caption`,
    /// `remove_slide_shape_caption`). The edit and its patch are those of
    /// image titles and captions; [`Self::apply_slide_image_text`] replays
    /// either.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and refusals.
    pub fn edit_slide_shape_text<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<crate::DrawableSelector>,
        text: Text,
    ) -> Result<Edit<'_>, Error> {
        let slide_selector = slide_selector.into();
        let position = shape_position(self, slide_selector, drawable_selector.into())?;
        Edit::new(self, Owner::Shape, slide_selector, position, text)
    }

    /// Apply an exact-source-checked image title or caption patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the
    /// patch's exact source.
    pub fn apply_slide_image_text(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let current = select_text(
            self,
            patch.owner,
            SlideSelector::position(patch.selection.slide_position),
            patch.selection.movie_position,
            patch.text,
            true,
        )?;
        if !current.same_identity(&patch.selection) || current.text != patch.selection.text {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self)
            .map_err(map_chart_error)?;
        budget.charge_catalog_scan(self).map_err(map_chart_error)?;
        budget
            .charge_selection_scan(self, 1)
            .map_err(map_chart_error)?;
        budget
            .charge_exact_artifacts(source.len(), patch.artifacts.target().len())
            .map_err(map_chart_error)?;
        budget
            .charge_candidate_reopen(patch.artifacts.target().len())
            .map_err(map_chart_error)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        budget
            .charge_selection_scan(&candidate, 1)
            .map_err(map_chart_error)?;
        let selected = select_text(
            &candidate,
            patch.owner,
            SlideSelector::position(patch.selection.slide_position),
            patch.selection.movie_position,
            patch.text,
            true,
        )?;
        if !selected.same_identity(&patch.target_selection) || selected.text != patch.after {
            return Err(Error::Verification);
        }
        budget
            .charge_validation_scan(&candidate, 1)
            .map_err(map_chart_error)?;
        let registered = if patch.selection.storage_identifier.is_none() && patch.after.is_some() {
            register_dependencies(self, &patch.selection)?
        } else {
            None
        };
        verify_text_transition(
            registered.as_ref().unwrap_or(self),
            &candidate,
            &patch.selection,
            &selected,
            patch.target_requires_invalidated_previews,
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(patch.touched_components, patch.deleted_previews),
        })
    }
}

fn select_text(
    package: &Package,
    owner: Owner,
    slide_selector: SlideSelector<'_>,
    position: Position,
    text: Text,
    mutation_guards: bool,
) -> Result<super::slide_movie_caption::MovieCaptionSelection, Error> {
    super::slide_movie_caption::select_drawable_text(
        package,
        slide_selector,
        owner.captioned(),
        position,
        mutation_guards,
        text.movie_kind(),
    )
    .map_err(map_movie_caption_error)
}

/// The position of the selected shape among the slide's ShapeInfo drawables.
fn shape_position(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    drawable: crate::DrawableSelector,
) -> Result<Position, Error> {
    super::slide_drawable::shape_info_position(package, slide_selector, drawable).map_err(|error| {
        match error {
            super::slide_drawable::Error::Refused(
                super::slide_drawable::Refusal::DrawablePositionNotFound { position },
            ) => Error::Refused(Refusal::ImagePositionNotFound { position }),
            super::slide_drawable::Error::Refused(_) => {
                Error::Refused(Refusal::UnsupportedDependency)
            },
            _ => Error::InvalidSource,
        }
    })
}

/// Register the new caption graph's foreign theme dependencies (stylesheet
/// and caption paragraph style) as external references of the image's
/// slide component, returning the registered package; `None` when every
/// dependency is already registered. The registration leaves save tokens to
/// the caption transaction that follows, which rewrites the same component.
/// The migration host wrote these edges unregistered (fixed in record 0372).
fn register_dependencies(
    source: &Package,
    selection: &super::slide_movie_caption::MovieCaptionSelection,
) -> Result<Option<Package>, Error> {
    let dependencies =
        super::slide_chart_caption::caption_dependencies(source).map_err(map_chart_error)?;
    let mut workspace =
        super::structure::workspace::Workspace::new(source).map_err(|_| Error::InvalidSource)?;
    let slide_component = workspace
        .registry
        .component_for_member(&selection.slide_component_name)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let mut added = false;
    for dependency in dependencies {
        let member = workspace
            .component_of(dependency)
            .ok_or(Error::InvalidSource)?;
        if member == selection.slide_component_name {
            continue;
        }
        let target = workspace
            .registry
            .component_for_member(&member)
            .ok_or(Error::InvalidSource)?
            .identifier;
        let present = workspace
            .registry
            .component(slide_component)
            .ok_or(Error::InvalidSource)?
            .external_references()
            .iter()
            .any(|reference| reference.component == target && reference.object == Some(dependency));
        if !present {
            workspace
                .registry
                .component_mut_untracked(slide_component)
                .ok_or(Error::InvalidSource)?
                .add_external_reference(super::structure::metadata::ExternalReference {
                    component: target,
                    object: Some(dependency),
                    is_weak: None,
                });
            added = true;
        }
    }
    if !added {
        return Ok(None);
    }
    let bytes = workspace.publish().map_err(|_| Error::InvalidSource)?;
    Ok(Some(
        source.structure_reopen(bytes).map_err(map_read_error)?,
    ))
}

fn rewrite_text(
    source: &Package,
    selection: &super::slide_movie_caption::MovieCaptionSelection,
    edge_kind: super::slide_chart_caption::CaptionEdgeKind,
    after: Option<&str>,
    budget: &mut super::slide_chart_caption::CaptionBudget,
) -> Result<(Package, usize, usize), Error> {
    let result = if selection.storage_identifier.is_some() && after.is_some() {
        let end = selection
            .text
            .as_deref()
            .ok_or(Error::InvalidSource)?
            .encode_utf16()
            .count();
        super::slide_chart_caption::rewrite_existing_caption_text_with_metadata_budget(
            source,
            selection.storage_identifier.ok_or(Error::InvalidSource)?,
            selection.slide_node_identifier,
            &selection.slide_component_name,
            end,
            after.ok_or(Error::InvalidSource)?,
            budget,
        )
    } else {
        super::slide_chart_caption::rewrite_caption_graph_operation_with_budget(
            source,
            &selection.slide_component_name,
            selection.movie_identifier,
            selection.reference_identifier,
            selection.storage_identifier,
            after,
            edge_kind,
            budget,
        )
    };
    result.map_err(map_chart_error)
}

fn verify_text_transition(
    source: &Package,
    candidate: &Package,
    before: &super::slide_movie_caption::MovieCaptionSelection,
    target: &super::slide_movie_caption::MovieCaptionSelection,
    require_invalidated_previews: bool,
    budget: &mut super::slide_chart_caption::CaptionBudget,
) -> Result<(), Error> {
    if before.storage_identifier.is_some() && target.storage_identifier.is_some() {
        super::slide_chart_caption::verify_existing_text_metadata_candidate(
            source,
            candidate,
            before.storage_identifier.ok_or(Error::InvalidSource)?,
            before.slide_node_identifier,
            require_invalidated_previews,
            budget,
        )
        .map_err(map_chart_error)
    } else {
        super::slide_chart_caption::verify_caption_graph_transition(
            source,
            candidate,
            &before.slide_component_name,
            before.reference_identifier,
            before.caption_info_identifier,
            before.storage_identifier,
            before.placement_identifier,
            before.style_identifier,
            target.reference_identifier,
            target.caption_info_identifier,
            target.storage_identifier,
            target.placement_identifier,
            target.style_identifier,
        )
        .map_err(map_chart_error)?;
        if require_invalidated_previews
            && !super::rendering_invalidation::root_previews_absent(
                candidate.state.source.package(),
            )
            .map_err(|_| Error::Verification)?
        {
            return Err(Error::Verification);
        }
        Ok(())
    }
}

fn copy_text(value: &str) -> Result<String, Error> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::ImageTextBytes,
            observed: usize_to_u64(value.len()),
            maximum: usize_to_u64(MAX_TEXT_BYTES),
        });
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_| Error::Allocation {
            amount: value.len(),
        })?;
    copy.push_str(value);
    Ok(copy)
}

fn physical_catalog(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
    }
}

fn map_movie_caption_error(error: super::slide_movie_caption::Error) -> Error {
    match error {
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::UnsupportedDependency,
        ) => Error::Refused(Refusal::UnsupportedDependency),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::AmbiguousSelector,
        ) => Error::Refused(Refusal::AmbiguousSelector),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::EmptySlideName,
        ) => Error::Refused(Refusal::EmptySlideName),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::SlideNameNotFound,
        ) => Error::Refused(Refusal::SlideNameNotFound),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::SlidePositionNotFound { position },
        ) => Error::Refused(Refusal::SlidePositionNotFound { position }),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::MoviePositionNotFound { position },
        ) => Error::Refused(Refusal::ImagePositionNotFound { position }),
        super::slide_movie_caption::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_caption_limit(kind),
            observed,
            maximum,
        },
        super::slide_movie_caption::Error::Allocation { amount } => Error::Allocation { amount },
        super::slide_movie_caption::Error::Verification => Error::Verification,
        super::slide_movie_caption::Error::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn map_chart_error(error: super::slide_chart_caption::Error) -> Error {
    match error {
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::UnsupportedDependency,
        ) => Error::Refused(Refusal::UnsupportedDependency),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::AmbiguousSelector,
        ) => Error::Refused(Refusal::AmbiguousSelector),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::EmptySlideName,
        ) => Error::Refused(Refusal::EmptySlideName),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::SlideNameNotFound,
        ) => Error::Refused(Refusal::SlideNameNotFound),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::SlidePositionNotFound { position },
        ) => Error::Refused(Refusal::SlidePositionNotFound { position }),
        super::slide_chart_caption::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_chart_limit(kind),
            observed,
            maximum,
        },
        super::slide_chart_caption::Error::Allocation { amount } => Error::Allocation { amount },
        super::slide_chart_caption::Error::Verification => Error::Verification,
        super::slide_chart_caption::Error::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::Entries,
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_caption_limit(kind: super::slide_movie_caption::LimitKind) -> LimitKind {
    use super::slide_movie_caption::LimitKind as K;
    match kind {
        K::InputBytes => LimitKind::InputBytes,
        K::OutputBytes => LimitKind::OutputBytes,
        K::WireBytes => LimitKind::WireBytes,
        K::Entries => LimitKind::Entries,
        K::EntryBytes => LimitKind::EntryBytes,
        K::TotalBytes => LimitKind::TotalBytes,
        K::Slides => LimitKind::Slides,
        K::References => LimitKind::References,
        K::TextStorages => LimitKind::TextStorages,
        K::TextFragments => LimitKind::TextFragments,
        K::TextBytes => LimitKind::TextBytes,
        K::WireFields => LimitKind::WireFields,
        K::WireNesting => LimitKind::WireNesting,
        K::WireWork => LimitKind::WireWork,
        K::CaptionBytes => LimitKind::ImageTextBytes,
    }
}

fn map_chart_limit(kind: super::slide_chart_caption::LimitKind) -> LimitKind {
    use super::slide_chart_caption::LimitKind as K;
    match kind {
        K::InputBytes => LimitKind::InputBytes,
        K::OutputBytes => LimitKind::OutputBytes,
        K::WireBytes => LimitKind::WireBytes,
        K::Entries => LimitKind::Entries,
        K::EntryBytes => LimitKind::EntryBytes,
        K::TotalBytes => LimitKind::TotalBytes,
        K::Slides => LimitKind::Slides,
        K::References => LimitKind::References,
        K::TextStorages => LimitKind::TextStorages,
        K::TextFragments => LimitKind::TextFragments,
        K::TextBytes => LimitKind::TextBytes,
        K::WireFields => LimitKind::WireFields,
        K::WireNesting => LimitKind::WireNesting,
        K::WireWork => LimitKind::WireWork,
        K::CaptionBytes => LimitKind::ImageTextBytes,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

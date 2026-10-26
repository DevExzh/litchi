//! Exact-source, selector-first Keynote slide creation: duplicate an existing
//! slide or add a fresh slide from a theme layout (record 0298).
//!
//! An [`Edit`] stages any number of operations against one immutable package.
//! Selectors are resolved against that base snapshot when an operation is
//! staged, and the transaction remembers the selected native graph privately,
//! so a later operation is never shifted by an earlier one. Destinations are
//! final positions in the slide list as it is when the operation runs
//! (`0..=` the slide count at that time). [`Edit::commit`] applies every
//! operation in order in one structural workspace, publishes one candidate
//! artifact, reopens and validates it, and verifies the created slides.

mod duplicate;
mod graph;
mod layout;
mod reassign;

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use super::structure::workspace::Workspace;
use super::structure::{StructureError, StructureLimit};
use super::{Package, PhysicalSource, ReadError};
use crate::slide::placeholder::Kind;
use crate::{SlideSelector, SlideSelectorError};

pub(super) use layout::Catalog;

/// One slide layout of the presentation theme.
///
/// A layout is identified by its position in the theme's layout list and,
/// usually, by its exact name (for example `"Title & Subtitle"`). Native
/// identifiers never cross this boundary.
///
/// ```no_run
/// use litchi_keynote::{Package, Position};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let layouts = package.slide_layouts()?;
/// let layout = layouts.first().ok_or("theme without layouts")?;
/// let mut edit = package.edit_slide_creation();
/// edit.add_slide(layout, Position::new(0))?;
/// let _commit = edit.commit()?;
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Layout {
    name: Option<String>,
    position: Position,
}

impl Layout {
    /// Return the layout's exact name, when the theme records one.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return the layout's zero-based position in the theme's layout list.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return a selector for this layout, preferring its position (which is
    /// unambiguous even when two layouts share a name).
    #[must_use]
    pub const fn selector(&self) -> LayoutSelector<'static> {
        LayoutSelector::Position(self.position)
    }
}

/// Selects one theme layout by exact name or by its position in the theme's
/// layout list.
///
/// `&str`, [`Position`], `usize` and [`&Layout`](Layout) convert into a
/// selector, so [`Edit::add_slide`] accepts any of them.
///
/// ```
/// use litchi_keynote::Position;
/// use litchi_keynote::slide::creation::LayoutSelector;
///
/// assert_eq!(LayoutSelector::from("Title Only"), LayoutSelector::Name("Title Only"));
/// assert_eq!(
///     LayoutSelector::from(2),
///     LayoutSelector::Position(Position::new(2))
/// );
/// ```
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LayoutSelector<'a> {
    /// Select the layout with this exact name.
    Name(&'a str),
    /// Select the layout at this zero-based position in the theme's list.
    Position(Position),
}

impl<'a> From<&'a str> for LayoutSelector<'a> {
    fn from(name: &'a str) -> Self {
        Self::Name(name)
    }
}

impl<'a> From<&'a String> for LayoutSelector<'a> {
    fn from(name: &'a String) -> Self {
        Self::Name(name)
    }
}

impl From<Position> for LayoutSelector<'_> {
    fn from(position: Position) -> Self {
        Self::Position(position)
    }
}

impl From<usize> for LayoutSelector<'_> {
    fn from(index: usize) -> Self {
        Self::Position(Position::new(index))
    }
}

impl From<&Layout> for LayoutSelector<'static> {
    fn from(layout: &Layout) -> Self {
        layout.selector()
    }
}

/// Why a slide or layout cannot be created by this transaction.
///
/// The reasons name content classes only; they never carry native
/// identifiers, component names or authored text.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Unsupported {
    /// The slide's objects reference content stored in other package
    /// components (for example tables, charts or comment authors).
    ExternalContent,
    /// A native message without a reference schema holds object references,
    /// so it cannot be cloned safely.
    UnmodelledContent,
    /// A package-registry record uses a shape this transaction does not edit
    /// (versioned references or ambiguous object identifiers).
    RegistryShape,
    /// The layout carries live-slide state (builds or speaker notes).
    LayoutState,
    /// The layout carries media that cannot be materialized on a live slide.
    LayoutMedia,
    /// No slide in the presentation has a speaker-notes graph to reuse for a
    /// fresh slide.
    NotesTemplate,
    /// The selected slide shares its package component with another slide,
    /// or its slide node has children.
    SharedComponent,
}

impl fmt::Display for Unsupported {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::ExternalContent => "content stored in other package components",
            Self::UnmodelledContent => "unmodelled native content with references",
            Self::RegistryShape => "an unmodelled package-registry shape",
            Self::LayoutState => "layout build or notes state",
            Self::LayoutMedia => "layout media that cannot be materialized",
            Self::NotesTemplate => "no reusable speaker-notes graph",
            Self::SharedComponent => "a shared slide component or nested slide node",
        })
    }
}

/// A finite resource governed by a slide-creation transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Slides in the presentation.
    Slides,
    /// Native objects in the package.
    Objects,
    /// The native object-identifier space.
    Identifiers,
    /// Bytes of one native payload.
    WireBytes,
    /// Fields of one native payload.
    WireFields,
    /// Nesting depth of one native payload.
    WireNesting,
    /// Rewrite work.
    Work,
    /// Bytes of one rewritten payload or member.
    OutputBytes,
    /// Bytes of one package entry, object or the package.
    EntryBytes,
    /// Rich-text storages read during verification.
    TextStorages,
    /// Rich-text fragments read during verification.
    TextFragments,
    /// Rich-text bytes read during verification.
    TextBytes,
    /// Object references read during verification.
    References,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Slides => "slides",
            Self::Objects => "objects",
            Self::Identifiers => "object identifiers",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::Work => "work",
            Self::OutputBytes => "output bytes",
            Self::EntryBytes => "entry bytes",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::References => "references",
        })
    }
}

/// A slide-creation failure that never exposes native identifiers, component
/// names or authored content.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The package cannot be edited safely (malformed or inconsistent).
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error("Keynote slide-creation {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// An allocation failed.
    #[error("could not allocate {amount} units for the Keynote slide-creation transaction")]
    Allocation {
        /// The requested amount.
        amount: usize,
    },
    /// The published candidate did not verify.
    #[error("the Keynote slide-creation candidate failed semantic verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Keynote slide-creation patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Keynote source does not support structural edits")]
    UnsupportedSource,
    /// The show's slide tree uses a topology this transaction does not edit.
    #[error("this Keynote slide topology is not supported for slide creation")]
    UnsupportedTopology,
    /// The selected slide or layout holds content that cannot be cloned.
    #[error("the Keynote slide cannot be created: it contains {reason}")]
    UnsupportedContent {
        /// The class of content that blocks the operation.
        reason: Unsupported,
    },
    /// Several slides share the selected navigator name.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// The slide selector names an empty navigator name.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide has the selected navigator name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide is at the selected position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// The theme has no layout with the selected name.
    #[error("the Keynote theme has no layout matching the requested name")]
    LayoutNameNotFound,
    /// Several layouts share the selected name.
    #[error("the Keynote layout selector is ambiguous")]
    AmbiguousLayout,
    /// The theme has no layout at the selected position.
    #[error("the Keynote theme has no layout at position {position:?}")]
    LayoutPositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// The destination is past the end of the slide list.
    #[error("slide destination {position:?} is outside 0..={count}")]
    PositionOutOfRange {
        /// The requested destination.
        position: Position,
        /// The slide count when the operation runs.
        count: usize,
    },
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::LayoutPositionNotFound { .. } => RefusalKind::InvalidValue,
            Self::LayoutNameNotFound { .. } => RefusalKind::InvalidValue,
            Self::UnsupportedSource { .. }
            | Self::UnsupportedTopology { .. }
            | Self::UnsupportedContent { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::AmbiguousLayout { .. }
            | Self::PositionOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. } | Self::SlidePositionNotFound { .. } => {
                RefusalKind::NotFound
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One staged operation, with privately remembered native identity.
#[derive(Clone, Copy)]
enum Operation {
    Duplicate {
        node: u64,
        source: Position,
        destination: usize,
    },
    Add {
        layout: usize,
        destination: usize,
    },
    Layout {
        node: u64,
        slide: Position,
        layout: usize,
    },
}

/// One bounded slide-creation edit staged against an immutable package.
///
/// Operations are applied in staging order and published as one commit.
///
/// ```no_run
/// use litchi_keynote::{Package, Position, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let mut edit = package.edit_slide_creation();
/// // Both selectors name slides of the base snapshot; the second copy
/// // lands at the end of the four-slide list the first copy produced.
/// edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?
///     .duplicate_slide(SlideSelector::name("Appendix"), Position::new(3))?;
/// let commit = edit.commit()?;
/// assert_eq!(commit.diagnostics().created_slides(), 2);
/// # Ok(())
/// # }
/// ```
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
    base_slides: Option<usize>,
    catalog: Option<Catalog>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    pub(super) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
            base_slides: None,
            catalog: None,
        }
    }

    fn base_slides(&mut self) -> Result<usize, Error> {
        if let Some(count) = self.base_slides {
            return Ok(count);
        }
        let count = self.source.show().map_err(map_read_error)?.slides().len();
        self.base_slides = Some(count);
        Ok(count)
    }

    /// Check a destination against the slide count when the next operation
    /// runs, and the semantic slide limit.
    fn destination(&mut self, destination: Position) -> Result<usize, Error> {
        let count = self
            .base_slides()?
            .checked_add(self.operations.len())
            .ok_or(Error::InvalidSource)?;
        if destination.get() > count {
            return Err(Error::Refused(Refusal::PositionOutOfRange {
                position: destination,
                count,
            }));
        }
        let maximum = self.source.semantic_limits().max_slides();
        if count >= maximum {
            return Err(Error::Limit {
                kind: LimitKind::Slides,
                observed: count as u64 + 1,
                maximum: maximum as u64,
            });
        }
        Ok(destination.get())
    }

    /// Stage a copy of one existing slide at `destination`.
    ///
    /// The selector is resolved against the base snapshot. `destination` is
    /// the final position of the copy in the slide list as it is when this
    /// operation runs, `0..=` that slide count.
    ///
    /// # Errors
    ///
    /// Returns a typed selector error when no single slide matches,
    /// [`Refusal::PositionOutOfRange`] for a destination past the end, and
    /// [`Error::Limit`] when the semantic slide limit would be
    /// exceeded. Unsupported slide content is reported by [`Self::commit`].
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.duplicate_slide(SlideSelector::index(0), Position::new(0))?;
    /// let copy = edit.commit()?.into_package();
    /// assert_eq!(
    ///     copy.slide_title(SlideSelector::index(0))?,
    ///     package.slide_title(SlideSelector::index(0))?
    /// );
    /// # Ok(())
    /// # }
    /// ```
    pub fn duplicate_slide<'s>(
        &mut self,
        selector: impl Into<SlideSelector<'s>>,
        destination: Position,
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        let show = self.source.show().map_err(map_read_error)?;
        let selected = show
            .select_slide(selector)
            .map_err(map_selector_error)?
            .ok_or(match selector {
                SlideSelector::Name(_) => Error::Refused(Refusal::SlideNameNotFound),
                SlideSelector::Position(position) => {
                    Error::Refused(Refusal::SlidePositionNotFound { position })
                },
            })?;
        let source = selected.position();
        let nodes = self
            .source
            .private_slide_node_order()
            .map_err(|_error| Error::InvalidSource)?;
        let node = *nodes.get(source.get()).ok_or(Error::InvalidSource)?;
        let destination = self.destination(destination)?;
        self.operations.push(Operation::Duplicate {
            node,
            source,
            destination,
        });
        Ok(self)
    }

    /// Stage the reassignment of one existing slide to a theme layout, as
    /// Keynote's Layout menu does.
    ///
    /// The slide keeps its text, speaker notes, builds, transition and
    /// ordinary drawables; the layout supplies the slide style, the title
    /// and body placeholder geometry, style, path and visibility, and the
    /// slide node's template identity. The slide preview is invalidated.
    /// Reassigning a slide to the layout it already uses stages nothing.
    /// The selector resolves against the base snapshot; the slide keeps its
    /// position. Layouts with layout-owned media (images, live video, media
    /// placeholders) on either side are refused by [`Self::commit`] with
    /// [`Unsupported::LayoutMedia`].
    ///
    /// # Errors
    ///
    /// Returns typed slide- and layout-selector errors.
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.set_layout(SlideSelector::index(0), "Title Only")?;
    /// let package = edit.commit()?.into_package();
    /// let layouts = package.slide_layout_list()?;
    /// assert_eq!(layouts[0].as_ref().and_then(|layout| layout.name()), Some("Title Only"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_layout<'s, 'l>(
        &mut self,
        selector: impl Into<SlideSelector<'s>>,
        layout: impl Into<LayoutSelector<'l>>,
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        let show = self.source.show().map_err(map_read_error)?;
        let selected = show
            .select_slide(selector)
            .map_err(map_selector_error)?
            .ok_or(match selector {
                SlideSelector::Name(_) => Error::Refused(Refusal::SlideNameNotFound),
                SlideSelector::Position(position) => {
                    Error::Refused(Refusal::SlidePositionNotFound { position })
                },
            })?;
        let slide = selected.position();
        if self.catalog.is_none() {
            self.catalog = Some(Catalog::read(self.source)?);
        }
        let catalog = self.catalog.as_ref().ok_or(Error::InvalidSource)?;
        let layout = catalog.select(layout.into())?;
        let nodes = self
            .source
            .private_slide_node_order()
            .map_err(|_error| Error::InvalidSource)?;
        let node = *nodes.get(slide.get()).ok_or(Error::InvalidSource)?;
        self.operations.push(Operation::Layout {
            node,
            slide,
            layout,
        });
        Ok(self)
    }

    /// Stage a fresh slide built from one theme layout at `destination`.
    ///
    /// The fresh slide's title, body and speaker notes are empty. Layouts
    /// with live-slide state or media that cannot be materialized are
    /// refused by [`Self::commit`] with [`Refusal::UnsupportedContent`].
    ///
    /// # Errors
    ///
    /// Returns a typed layout-selector error when no single layout matches,
    /// [`Refusal::PositionOutOfRange`] for a destination past the end, and
    /// [`Error::Limit`] when the semantic slide limit would be
    /// exceeded.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::placeholder::Kind;
    /// use litchi_keynote::{Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.add_slide("Title & Bullets", Position::new(1))?;
    /// let package = edit.commit()?.into_package();
    /// let mut title = package.edit_slide_text(SlideSelector::index(1), Kind::Title)?;
    /// title.set("Agenda")?;
    /// let _package = title.commit()?.into_package();
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide<'s>(
        &mut self,
        layout: impl Into<LayoutSelector<'s>>,
        destination: Position,
    ) -> Result<&mut Self, Error> {
        let selector = layout.into();
        if self.catalog.is_none() {
            self.catalog = Some(Catalog::read(self.source)?);
        }
        let catalog = self.catalog.as_ref().ok_or(Error::InvalidSource)?;
        let layout = catalog.select(selector)?;
        let destination = self.destination(destination)?;
        self.operations.push(Operation::Add {
            layout,
            destination,
        });
        Ok(self)
    }

    /// Apply every staged operation in one workspace, publish one candidate,
    /// reopen and verify it.
    ///
    /// An edit with no staged operation commits the unchanged package.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::UnsupportedContent`] for slides or layouts this
    /// transaction cannot clone, [`Error::Verification`] when the candidate
    /// does not verify, and typed source, limit and allocation errors.
    ///
    /// ```no_run
    /// use litchi_keynote::slide::creation::{Error, Refusal, Unsupported};
    /// use litchi_keynote::{Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck-with-table.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
    /// match edit.commit() {
    ///     Ok(commit) => println!("{} slides", commit.package().slides()?.len()),
    ///     Err(Error::Refused(Refusal::UnsupportedContent { reason })) => println!("refused: {reason}"),
    ///     Err(error) => return Err(error.into()),
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn commit(self) -> Result<Commit, Error> {
        let source_bytes = physical_source(self.source)?;
        if self.operations.is_empty() {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    source_slide_count: self.source.show().map_err(map_read_error)?.slides().len(),
                    created_slides: 0,
                    touched_components: 0,
                    creates: true,
                },
                diagnostics: Diagnostics {
                    changed: false,
                    created_slides: 0,
                    touched_components: 0,
                    full_reparse_performed: false,
                },
            });
        }
        let catalog = match self.catalog {
            Some(catalog) => Some(catalog),
            None if self.operations.iter().any(|operation| {
                matches!(operation, Operation::Add { .. } | Operation::Layout { .. })
            }) =>
            {
                Some(Catalog::read(self.source)?)
            },
            None => None,
        };
        commit_operations(
            self.source,
            source_bytes,
            &self.operations,
            catalog.as_ref(),
        )
    }
}

/// Exact-source-checked reversible slide-creation patch.
///
/// The patch retains the exact source and target artifacts; applying it
/// checks the exact source bytes, and its inverse removes the created slides
/// by restoring the source.
///
/// ```no_run
/// use litchi_keynote::{Package, Position, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let mut edit = package.edit_slide_creation();
/// edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?;
/// let (created, patch) = edit.commit()?.into_parts();
/// let restored = created.apply_slide_creation(&patch.inverse())?;
/// assert_eq!(restored.package().slides()?.len(), patch.source_slide_count());
/// # Ok(())
/// # }
/// ```
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    source_slide_count: usize,
    created_slides: usize,
    touched_components: usize,
    creates: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("created_slides", &self.created_slides)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            source_slide_count: self.target_slide_count(),
            created_slides: self.created_slides,
            touched_components: self.touched_components,
            creates: !self.creates,
        }
    }

    /// Return the slide count of the patch's source.
    #[must_use]
    pub const fn source_slide_count(&self) -> usize {
        self.source_slide_count
    }

    /// Return the slide count of the patch's target.
    #[must_use]
    pub const fn target_slide_count(&self) -> usize {
        if self.creates {
            self.source_slide_count + self.created_slides
        } else {
            self.source_slide_count - self.created_slides
        }
    }

    /// Return how many slides the forward transaction created.
    #[must_use]
    pub const fn created_slides(&self) -> usize {
        self.created_slides
    }

    /// Return the fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the fingerprint of the exact target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
}

/// Compact evidence for one verified slide-creation commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    created_slides: usize,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return how many slides were created (or, for an applied inverse,
    /// removed).
    #[must_use]
    pub const fn created_slides(self) -> usize {
        self.created_slides
    }

    /// Return how many package components were created or rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was fully reparsed and validated.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one slide-creation edit.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the committed package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Split the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    /// Start one slide-creation edit (duplicate slides, add slides from
    /// theme layouts).
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.duplicate_slide(SlideSelector::index(0), Position::new(1))?
    ///     .add_slide("Title & Subtitle", Position::new(2))?;
    /// let commit = edit.commit()?;
    /// assert_eq!(commit.diagnostics().created_slides(), 2);
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_slide_creation(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply a slide-creation patch only to its exact immutable source.
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, Position, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// let mut edit = package.edit_slide_creation();
    /// edit.add_slide("Blank", Position::new(0))?;
    /// let patch = edit.commit()?.patch().clone();
    /// // Replaying the patch on an independent copy of the source.
    /// let copy = Package::open("deck.key")?;
    /// let replayed = copy.apply_slide_creation(&patch)?;
    /// assert_eq!(replayed.diagnostics().created_slides(), 1);
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// exact source, and typed source errors when the target cannot be
    /// reopened.
    pub fn apply_slide_creation(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if self.show().map_err(map_read_error)?.slides().len() != patch.source_slide_count {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        if candidate.show().map_err(map_read_error)?.slides().len() != patch.target_slide_count() {
            return Err(Error::PatchConflict);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                created_slides: patch.created_slides,
                touched_components: patch.touched_components,
                full_reparse_performed: true,
            },
        })
    }

    /// List the slide layouts of the presentation theme, in theme order.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// for layout in package.slide_layouts()?.iter() {
    ///     println!("{:?}: {:?}", layout.position(), layout.name());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] when the theme's layout list is
    /// malformed.
    pub fn slide_layouts(&self) -> Result<Box<[Layout]>, Error> {
        Ok(Catalog::read(self)?.layouts())
    }

    /// List the layout each slide was made from, in show order: `None` for a
    /// slide whose template slide the theme does not list.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// for (slide, layout) in package.slide_layout_list()?.iter().enumerate() {
    ///     println!("slide {slide}: {:?}", layout.as_ref().and_then(|layout| layout.name()));
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] when the theme's layout list or a
    /// slide record is malformed.
    pub fn slide_layout_list(&self) -> Result<Box<[Option<Layout>]>, Error> {
        Catalog::read(self)?.assignments(self)
    }

    /// Return the layout the theme records as its default, if any.
    ///
    /// ```no_run
    /// use litchi_keynote::{Package, Position};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("deck.key")?;
    /// if let Some(layout) = package.default_slide_layout()? {
    ///     let mut edit = package.edit_slide_creation();
    ///     edit.add_slide(&layout, Position::new(0))?;
    ///     let _commit = edit.commit()?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] when the theme's layout list is
    /// malformed.
    pub fn default_slide_layout(&self) -> Result<Option<Layout>, Error> {
        Ok(Catalog::read(self)?.default_layout())
    }
}

/// What one operation created, for verification.
struct Created {
    node: u64,
    expectation: Expectation,
}

enum Expectation {
    /// Title, body and notes equal to the source slide's.
    Copy(Position),
    /// Empty title, body and notes.
    Empty,
}

fn commit_operations(
    source: &Package,
    source_bytes: Arc<[u8]>,
    operations: &[Operation],
    catalog: Option<&Catalog>,
) -> Result<Commit, Error> {
    source.validate().map_err(map_read_error)?;
    let source_slide_count = source.show().map_err(map_read_error)?.slides().len();
    let mut workspace = Workspace::new(source).lift()?;
    let show = graph::ShowTree::read(&workspace)?;
    let mut order = show.order().to_vec();
    if order.len() != source_slide_count {
        return Err(Error::InvalidSource);
    }
    let mut notes = None;
    let mut created = Vec::with_capacity(operations.len());
    let mut reassigned = Vec::new();
    for operation in operations {
        let (node, destination, expectation) = match *operation {
            Operation::Layout {
                node,
                slide,
                layout,
            } => {
                let record = catalog.ok_or(Error::InvalidSource)?.record(layout)?;
                if !reassign::stage(&mut workspace, node, record)?.unchanged {
                    reassigned.push((node, slide, layout));
                }
                continue;
            },
            Operation::Duplicate {
                node,
                source: position,
                destination,
            } => (
                duplicate::stage(&mut workspace, node)?,
                destination,
                Expectation::Copy(position),
            ),
            Operation::Add {
                layout,
                destination,
            } => {
                let record = catalog.ok_or(Error::InvalidSource)?.record(layout)?;
                if notes.is_none() {
                    notes = Some(layout::NoteSource::find(&workspace, show.order())?);
                }
                let note = notes.as_ref().ok_or(Error::InvalidSource)?;
                (
                    layout::stage(&mut workspace, record, note)?,
                    destination,
                    Expectation::Empty,
                )
            },
        };
        if destination > order.len() {
            return Err(Error::InvalidSource);
        }
        order.insert(destination, node);
        created.push(Created { node, expectation });
    }
    if created.is_empty() && reassigned.is_empty() {
        let source_bytes = physical_source(source)?;
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                source_slide_count,
                created_slides: 0,
                touched_components: 0,
                creates: true,
            },
            diagnostics: Diagnostics {
                changed: false,
                created_slides: 0,
                touched_components: 0,
                full_reparse_performed: false,
            },
        });
    }
    check_object_limit(source, &workspace)?;
    if !created.is_empty() {
        show.write(&mut workspace, &order)?;
    }
    let touched_components = workspace.touched_members();
    let bytes = workspace.publish().lift()?;
    let mut candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    let target_order = candidate
        .private_slide_node_order()
        .map_err(|_error| Error::Verification)?;
    if target_order.as_ref() != order.as_slice() {
        return Err(Error::Verification);
    }
    for item in &created {
        if matches!(item.expectation, Expectation::Empty) {
            let position = position_of(&order, item.node)?;
            candidate = clear_text(candidate, position)?;
        }
    }
    verify(source, &candidate, &order, &created)?;
    if !reassigned.is_empty() {
        let layouts = catalog
            .ok_or(Error::InvalidSource)?
            .assignments(&candidate)?;
        for (node, slide, layout) in &reassigned {
            let position = position_of(&order, *node)?;
            let expected = catalog.ok_or(Error::InvalidSource)?.layouts();
            if layouts.get(position.get()).cloned().flatten().as_ref() != expected.get(*layout)
                || slide_text(source, *slide)? != slide_text(&candidate, position)?
            {
                return Err(Error::Verification);
            }
        }
    }
    let target = physical_source(&candidate)?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target),
        source_slide_count,
        created_slides: created.len(),
        touched_components,
        creates: true,
    };
    Ok(Commit {
        package: candidate,
        diagnostics: Diagnostics {
            changed: !patch.is_noop(),
            created_slides: created.len(),
            touched_components,
            full_reparse_performed: true,
        },
        patch,
    })
}

fn check_object_limit(source: &Package, workspace: &Workspace<'_>) -> Result<(), Error> {
    let maximum = source.semantic_limits().max_objects();
    let observed = source
        .structure_object_count()
        .saturating_add(workspace.created_objects());
    if observed > maximum {
        return Err(Error::Limit {
            kind: LimitKind::Objects,
            observed: observed as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
}

fn position_of(order: &[u64], node: u64) -> Result<Position, Error> {
    order
        .iter()
        .position(|candidate| *candidate == node)
        .map(Position::new)
        .ok_or(Error::Verification)
}

/// Clear the title, body and notes text a fresh slide inherited from its
/// layout and notes source, chaining the focused text transactions.
fn clear_text(mut package: Package, position: Position) -> Result<Package, Error> {
    let selector = SlideSelector::position(position);
    for role in [Kind::Title, Kind::Body] {
        let text = package
            .slide_text(selector, role)
            .map_err(|_error| Error::Verification)?;
        if text.is_some_and(|text| !text.is_empty()) {
            let mut edit = package
                .edit_slide_text(selector, role)
                .map_err(|_error| Error::Verification)?;
            edit.clear().map_err(|_error| Error::Verification)?;
            package = edit
                .commit()
                .map_err(|_error| Error::Verification)?
                .into_package();
        }
    }
    let notes = package
        .slide_notes(selector)
        .map_err(|_error| Error::Verification)?;
    if notes.is_some_and(|notes| !notes.is_empty()) {
        let mut edit = package
            .edit_slide_notes(selector)
            .map_err(|_error| Error::Verification)?;
        edit.clear().map_err(|_error| Error::Verification)?;
        package = edit
            .commit()
            .map_err(|_error| Error::Verification)?
            .into_package();
    }
    Ok(package)
}

type TextTriple = (Option<String>, Option<String>, Option<String>);

fn slide_text(package: &Package, position: Position) -> Result<TextTriple, Error> {
    let selector = SlideSelector::position(position);
    let title = package
        .slide_text(selector, Kind::Title)
        .map_err(|_error| Error::Verification)?;
    let body = package
        .slide_text(selector, Kind::Body)
        .map_err(|_error| Error::Verification)?;
    let notes = package
        .slide_notes(selector)
        .map_err(|_error| Error::Verification)?;
    Ok((title, body, notes))
}

fn verify(
    source: &Package,
    candidate: &Package,
    order: &[u64],
    created: &[Created],
) -> Result<(), Error> {
    if candidate.show().map_err(map_read_error)?.slides().len() != order.len() {
        return Err(Error::Verification);
    }
    for item in created {
        let position = position_of(order, item.node)?;
        let actual = slide_text(candidate, position)?;
        match item.expectation {
            Expectation::Copy(source_position) => {
                if slide_text(source, source_position)? != actual {
                    return Err(Error::Verification);
                }
            },
            Expectation::Empty => {
                let empty = |text: &Option<String>| text.as_deref().is_none_or(str::is_empty);
                if !(empty(&actual.0) && empty(&actual.1) && empty(&actual.2)) {
                    return Err(Error::Verification);
                }
            },
        }
    }
    Ok(())
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

/// Lift kernel results into this family's error vocabulary.
trait Lift<T> {
    fn lift(self) -> Result<T, Error>;
}

impl<T> Lift<T> for Result<T, StructureError> {
    fn lift(self) -> Result<T, Error> {
        self.map_err(map_structure_error)
    }
}

fn map_structure_error(error: StructureError) -> Error {
    match error {
        StructureError::UnsupportedContent(_) => Error::Refused(Refusal::UnsupportedContent {
            reason: Unsupported::UnmodelledContent,
        }),
        StructureError::UnsupportedRegistry => Error::Refused(Refusal::UnsupportedContent {
            reason: Unsupported::RegistryShape,
        }),
        StructureError::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
        StructureError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                StructureLimit::WireBytes => LimitKind::WireBytes,
                StructureLimit::WireFields => LimitKind::WireFields,
                StructureLimit::WireNesting => LimitKind::WireNesting,
                StructureLimit::Work => LimitKind::Work,
                StructureLimit::OutputBytes => LimitKind::OutputBytes,
                StructureLimit::EntryBytes => LimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        StructureError::Allocation(amount) => Error::Allocation { amount },
        StructureError::MissingObject
        | StructureError::Registry(_)
        | StructureError::Unexpected(_)
        | StructureError::Malformed
        | StructureError::Physical => Error::InvalidSource,
    }
}

fn map_selector_error(error: SlideSelectorError) -> Error {
    match error {
        SlideSelectorError::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        SlideSelectorError::DuplicateSlideName { .. } => Error::Refused(Refusal::AmbiguousSelector),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::Limit {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Objects,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::Work,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

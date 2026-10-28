//! Exact-source, selector-first Keynote slide-chart font transactions.
//!
//! The chart-font graph is shared with Pages and Numbers through
//! [`litchi_iwa_style::chart_font`].  This module owns only the Keynote
//! selector, package publication, and semantic transaction boundary; native
//! chart identifiers and paragraph-style objects never cross that boundary.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The package boundary redacts native graph failures."
)]

use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, EntryInsertion, ExactArtifacts},
};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_style::{
    chart_font::{self as engine, Change as EngineChange},
    graph::{self, Components, Workspace},
};
use litchi_iwa_text::{character::TextPointSize, font::Font};
use thiserror::Error as ThisError;

use super::edit::{OwnedChartSelector, OwnedSlideSelector, SelectorError};
use super::slide_chart_title::{self, ChartSelection};
use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartSelector, SlideSelector};

/// A chart's effective font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartFont {
    /// The validated PostScript font identity, or the native application
    /// default when [`Font::Default`] is used.
    pub font: Font,
    /// Whether bold emphasis is enabled.
    pub bold: bool,
    /// Whether italic emphasis is enabled.
    pub italic: bool,
}

impl ChartFont {
    /// Construct a chart font from an existing validated font value.
    #[must_use]
    pub const fn new(font: Font, bold: bool, italic: bool) -> Self {
        Self { font, bold, italic }
    }

    /// Construct a named chart font after validating its PostScript name.
    ///
    /// # Errors
    ///
    /// Returns the underlying font-name validation error when `name` is
    /// empty, too long, surrounded by whitespace, or contains a control
    /// character.
    pub fn named(
        name: &str,
        bold: bool,
        italic: bool,
    ) -> Result<Self, litchi_iwa_text::font::NameError> {
        Ok(Self::new(Font::named(name)?, bold, italic))
    }
}

/// A finite resource governed by one Keynote chart-font transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Logical transaction allocations.
    Allocations,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// Bytes retained by a graph transaction.
    RetainedBytes,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Caller bytes retained by deferred selectors.
    SelectorBytes,
    /// Package-wide object identifiers.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::Allocations => "allocations",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::SelectorBytes => "selector bytes",
            Self::Identifiers => "identifiers",
        })
    }
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart-font edits")]
    UnsupportedSource,
    /// The selected chart style is not safe for this focused route.
    #[error("the selected Keynote chart style is not supported by the chart-font route")]
    UnsupportedStyle,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-font selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked semantic slide position.
        position: Position,
    },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound {
        /// Missing checked semantic chart position.
        position: Position,
    },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// Commit was requested without a staged operation.
    #[error("the Keynote chart-font transaction has no staged operation")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource | Self::UnsupportedStyle => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector | Self::EmptySlideName | Self::EmptyChartName => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound
            | Self::SlidePositionNotFound { .. }
            | Self::ChartNameNotFound
            | Self::ChartPositionNotFound { .. } => RefusalKind::NotFound,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// Failure from a Keynote chart-font read or edit.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The selected chart graph is malformed or unsupported.
    #[error("the selected Keynote chart-font source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart font {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-font transaction")]
    Allocation {
        /// Elements or bytes requested by the failed allocation.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested value.
    #[error("the edited Keynote chart font failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-font patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Compact evidence describing one chart-font commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
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

/// The fully verified result of one immutable chart-font transaction.
#[must_use = "a Keynote chart-font commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct FontCommit {
    package: Package,
    patch: FontPatch,
    diagnostics: Diagnostics,
}

impl FontCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, FontPatch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &FontPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// The fully verified result of one immutable chart-font-size transaction.
#[must_use = "a Keynote chart-font-size commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct FontSizeCommit {
    package: Package,
    patch: FontSizePatch,
    diagnostics: Diagnostics,
}

impl FontSizeCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, FontSizePatch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &FontSizePatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// Exact-source-checked reversible chart-font patch.
#[derive(Clone, PartialEq, Eq)]
pub struct FontPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    chart_identifier: u64,
    touched_components: usize,
    before: ChartFont,
    after: ChartFont,
}

impl fmt::Debug for FontPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FontPatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl FontPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Borrow the effective font required from the source package.
    #[must_use]
    pub const fn before(&self) -> &ChartFont {
        &self.before
    }

    /// Borrow the effective font produced by the candidate package.
    #[must_use]
    pub const fn after(&self) -> &ChartFont {
        &self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            slide_identifier: self.slide_identifier,
            chart_identifier: self.chart_identifier,
            touched_components: self.touched_components,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Exact-source-checked reversible chart-font-size patch.
#[derive(Clone, PartialEq)]
pub struct FontSizePatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    chart_identifier: u64,
    touched_components: usize,
    before: TextPointSize,
    after: TextPointSize,
}

impl fmt::Debug for FontSizePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FontSizePatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl FontSizePatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Return the effective point size required from the source package.
    #[must_use]
    pub const fn before(&self) -> TextPointSize {
        self.before
    }

    /// Return the effective point size produced by the candidate package.
    #[must_use]
    pub const fn after(&self) -> TextPointSize {
        self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            slide_identifier: self.slide_identifier,
            chart_identifier: self.chart_identifier,
            touched_components: self.touched_components,
            before: self.after,
            after: self.before,
        }
    }
}

/// One staged chart-font operation.
pub struct FontEdit<'a> {
    source: &'a Package,
    slide: OwnedSlideSelector,
    chart: OwnedChartSelector,
    selection: OnceLock<(ChartSelection, ChartFont)>,
    after: Option<FontOperation>,
}

impl fmt::Debug for FontEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FontEdit")
            .field("has_selection", &self.selection.get().is_some())
            .field("has_staged", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

#[derive(Clone)]
enum FontOperation {
    Set(ChartFont),
    Reset,
}

impl<'a> FontEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, Error> {
        let (slide, chart) = owned_selectors(source, slide_selector.into(), chart_selector.into())?;
        Ok(Self {
            source,
            slide,
            chart,
            selection: OnceLock::new(),
            after: None,
        })
    }

    fn resolve(&self) -> Result<&(ChartSelection, ChartFont), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let selection = slide_chart_title::select_chart(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
        )
        .map_err(map_chart_title_error)?;
        let before = read_font(self.source, selection.chart_identifier)?;
        let _ = self.selection.set((selection, before));
        self.selection.get().ok_or(Error::InvalidSource)
    }

    /// Return the selected semantic slide position.
    #[must_use = "handle the selected chart-font result"]
    pub fn slide_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.slide_position)
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use = "handle the selected chart-font result"]
    pub fn chart_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.chart_position)
    }

    /// Return the effective font observed when this edit first resolves.
    #[must_use = "handle the selected chart-font result"]
    pub fn before(&self) -> Result<&ChartFont, Error> {
        Ok(&self.resolve()?.1)
    }

    /// Return the staged effective font for a set, or the observed value for a
    /// reset (the reset's final inherited value is reported by its patch).
    #[must_use = "handle the staged chart-font result"]
    pub fn after(&self) -> Result<&ChartFont, Error> {
        match self.after.as_ref() {
            Some(FontOperation::Set(value)) => Ok(value),
            Some(FontOperation::Reset) | None => self.before(),
        }
    }

    /// Stage a complete effective chart-font value.
    #[must_use]
    pub fn set(mut self, value: ChartFont) -> Self {
        self.after = Some(FontOperation::Set(value));
        self
    }

    /// Stage removal of direct chart-font overrides.
    #[must_use]
    pub fn reset(mut self) -> Self {
        self.after = Some(FontOperation::Reset);
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<FontCommit, Error> {
        let (source_selection, before) = self.resolve()?.clone();
        let operation = self.after.clone().ok_or(Refusal::NoStagedOperation)?;
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let slide_position = source_selection.slide_position;
        let chart_position = source_selection.chart_position;

        if let FontOperation::Set(target) = &operation {
            if before == *target {
                self.source.validate().map_err(map_read_error)?;
                return Ok(FontCommit {
                    package: self.source.snapshot(),
                    patch: FontPatch {
                        artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                        slide_position,
                        chart_position,
                        slide_identifier: source_selection.slide_identifier,
                        chart_identifier: source_selection.chart_identifier,
                        touched_components: 0,
                        before: before.clone(),
                        after: before,
                    },
                    diagnostics: Diagnostics::unchanged(),
                });
            }
        }
        if !catalog.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        self.source.validate().map_err(map_read_error)?;
        let current = positional_selection(self.source, &source_selection)?;
        if !same_selection(&source_selection, &current)
            || read_font(self.source, current.chart_identifier)? != before
        {
            return Err(Error::InvalidSource);
        }
        let expected = match &operation {
            FontOperation::Set(value) => Some(value.clone()),
            FontOperation::Reset => None,
        };
        let change = match operation {
            FontOperation::Set(value) => EngineChange::Font(engine::ChartFont {
                font: value.font,
                bold: value.bold,
                italic: value.italic,
            }),
            FontOperation::Reset => EngineChange::ResetFont,
        };
        let (candidate, changed, touched_components) =
            write_graph(self.source, current.chart_identifier, &change)?;
        if !changed {
            return Ok(FontCommit {
                package: self.source.snapshot(),
                patch: FontPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position,
                    chart_position,
                    slide_identifier: source_selection.slide_identifier,
                    chart_identifier: source_selection.chart_identifier,
                    touched_components: 0,
                    before: before.clone(),
                    after: before,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate_selection = positional_selection(&candidate, &source_selection)?;
        let after = read_font(&candidate, candidate_selection.chart_identifier)?;
        if !same_selection(&source_selection, &candidate_selection)
            || expected.as_ref().is_some_and(|value| after != *value)
        {
            return Err(Error::Verification);
        }
        let target = physical_catalog(&candidate)?.shared_source();
        Ok(FontCommit {
            package: candidate,
            patch: FontPatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position,
                chart_position,
                slide_identifier: source_selection.slide_identifier,
                chart_identifier: source_selection.chart_identifier,
                touched_components,
                before,
                after,
            },
            diagnostics: Diagnostics::published(touched_components),
        })
    }
}

/// One staged chart-font-size operation.
pub struct FontSizeEdit<'a> {
    source: &'a Package,
    slide: OwnedSlideSelector,
    chart: OwnedChartSelector,
    selection: OnceLock<(ChartSelection, TextPointSize)>,
    after: Option<FontSizeOperation>,
}

impl fmt::Debug for FontSizeEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FontSizeEdit")
            .field("has_selection", &self.selection.get().is_some())
            .field("has_staged", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

#[derive(Clone, Copy)]
enum FontSizeOperation {
    Set(TextPointSize),
    Reset,
}

impl<'a> FontSizeEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, Error> {
        let (slide, chart) = owned_selectors(source, slide_selector.into(), chart_selector.into())?;
        Ok(Self {
            source,
            slide,
            chart,
            selection: OnceLock::new(),
            after: None,
        })
    }

    fn resolve(&self) -> Result<&(ChartSelection, TextPointSize), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let selection = slide_chart_title::select_chart(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
        )
        .map_err(map_chart_title_error)?;
        let before = read_size(self.source, selection.chart_identifier)?;
        let _ = self.selection.set((selection, before));
        self.selection.get().ok_or(Error::InvalidSource)
    }

    /// Return the selected semantic slide position.
    #[must_use = "handle the selected chart-font-size result"]
    pub fn slide_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.slide_position)
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use = "handle the selected chart-font-size result"]
    pub fn chart_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.chart_position)
    }

    /// Return the effective point size observed when this edit first resolves.
    #[must_use = "handle the selected chart-font-size result"]
    pub fn before(&self) -> Result<TextPointSize, Error> {
        Ok(self.resolve()?.1)
    }

    /// Return the staged point size for a set, or the observed value for a
    /// reset (the reset's final inherited value is reported by its patch).
    #[must_use = "handle the staged chart-font-size result"]
    pub fn after(&self) -> Result<TextPointSize, Error> {
        match self.after {
            Some(FontSizeOperation::Set(value)) => Ok(value),
            Some(FontSizeOperation::Reset) | None => self.before(),
        }
    }

    /// Stage a complete effective chart-font point size.
    #[must_use]
    pub fn set(mut self, value: TextPointSize) -> Self {
        self.after = Some(FontSizeOperation::Set(value));
        self
    }

    /// Stage removal of direct chart-font point-size overrides.
    #[must_use]
    pub fn reset(mut self) -> Self {
        self.after = Some(FontSizeOperation::Reset);
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<FontSizeCommit, Error> {
        let (source_selection, before) = self.resolve()?.clone();
        let operation = self.after.ok_or(Refusal::NoStagedOperation)?;
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let slide_position = source_selection.slide_position;
        let chart_position = source_selection.chart_position;

        if let FontSizeOperation::Set(target) = operation {
            if before == target {
                self.source.validate().map_err(map_read_error)?;
                return Ok(FontSizeCommit {
                    package: self.source.snapshot(),
                    patch: FontSizePatch {
                        artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                        slide_position,
                        chart_position,
                        slide_identifier: source_selection.slide_identifier,
                        chart_identifier: source_selection.chart_identifier,
                        touched_components: 0,
                        before,
                        after: before,
                    },
                    diagnostics: Diagnostics::unchanged(),
                });
            }
        }
        if !catalog.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        self.source.validate().map_err(map_read_error)?;
        let current = positional_selection(self.source, &source_selection)?;
        if !same_selection(&source_selection, &current)
            || read_size(self.source, current.chart_identifier)? != before
        {
            return Err(Error::InvalidSource);
        }
        let change = match operation {
            FontSizeOperation::Set(value) => EngineChange::Size(value),
            FontSizeOperation::Reset => EngineChange::ResetSize,
        };
        let (candidate, changed, touched_components) =
            write_graph(self.source, current.chart_identifier, &change)?;
        if !changed {
            return Ok(FontSizeCommit {
                package: self.source.snapshot(),
                patch: FontSizePatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position,
                    chart_position,
                    slide_identifier: source_selection.slide_identifier,
                    chart_identifier: source_selection.chart_identifier,
                    touched_components: 0,
                    before,
                    after: before,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate_selection = positional_selection(&candidate, &source_selection)?;
        let after = read_size(&candidate, candidate_selection.chart_identifier)?;
        if !same_selection(&source_selection, &candidate_selection)
            || (matches!(operation, FontSizeOperation::Set(_))
                && after
                    != match operation {
                        FontSizeOperation::Set(value) => value,
                        FontSizeOperation::Reset => unreachable!("reset is not a set"),
                    })
        {
            return Err(Error::Verification);
        }
        let target = physical_catalog(&candidate)?.shared_source();
        Ok(FontSizeCommit {
            package: candidate,
            patch: FontSizePatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position,
                chart_position,
                slide_identifier: source_selection.slide_identifier,
                chart_identifier: source_selection.chart_identifier,
                touched_components,
                before,
                after,
            },
            diagnostics: Diagnostics::published(touched_components),
        })
    }
}

impl Package {
    /// Read the selected slide chart's effective uniform font and face traits.
    pub fn slide_chart_font<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartFont, Error> {
        let selection = slide_chart_title::select_chart(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
        )
        .map_err(map_chart_title_error)?;
        read_font(self, selection.chart_identifier)
    }

    /// Start an exact immutable edit of one selected slide chart's font.
    pub fn edit_slide_chart_font<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<FontEdit<'_>, Error> {
        FontEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked slide-chart-font patch.
    pub fn apply_slide_chart_font(&self, patch: &FontPatch) -> Result<FontCommit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let selection = positional_selection(
            self,
            &ChartSelection {
                slide_position: patch.slide_position,
                chart_position: patch.chart_position,
                slide_identifier: patch.slide_identifier,
                chart_identifier: patch.chart_identifier,
                non_style_identifier: 0,
                non_style_message_index: 0,
                title: None,
                slide_component_name: String::new(),
                non_style_component_name: String::new(),
            },
        )?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
            || read_font(self, selection.chart_identifier)? != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(FontCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = positional_selection(&candidate, &selection)?;
        if !same_selection(&selection, &candidate_selection)
            || read_font(&candidate, candidate_selection.chart_identifier)? != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(FontCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(patch.touched_components),
        })
    }

    /// Read the selected slide chart's effective uniform point size.
    pub fn slide_chart_font_size<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<TextPointSize, Error> {
        let selection = slide_chart_title::select_chart(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
        )
        .map_err(map_chart_title_error)?;
        read_size(self, selection.chart_identifier)
    }

    /// Start an exact immutable edit of one selected slide chart's point size.
    pub fn edit_slide_chart_font_size<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<FontSizeEdit<'_>, Error> {
        FontSizeEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked slide-chart-font-size patch.
    pub fn apply_slide_chart_font_size(
        &self,
        patch: &FontSizePatch,
    ) -> Result<FontSizeCommit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let selection = positional_selection(
            self,
            &ChartSelection {
                slide_position: patch.slide_position,
                chart_position: patch.chart_position,
                slide_identifier: patch.slide_identifier,
                chart_identifier: patch.chart_identifier,
                non_style_identifier: 0,
                non_style_message_index: 0,
                title: None,
                slide_component_name: String::new(),
                non_style_component_name: String::new(),
            },
        )?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
            || read_size(self, selection.chart_identifier)? != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(FontSizeCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = positional_selection(&candidate, &selection)?;
        if !same_selection(&selection, &candidate_selection)
            || read_size(&candidate, candidate_selection.chart_identifier)? != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(FontSizeCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(patch.touched_components),
        })
    }
}

struct PackageComponents<'a>(&'a Package);

impl Components for PackageComponents<'_> {
    fn count(&self) -> usize {
        self.0.state.source.components().len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.0
            .state
            .source
            .components()
            .get_index(index)
            .map(|component| (component.name(), component.archive()))
    }
}

fn owned_selectors(
    source: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
) -> Result<(OwnedSlideSelector, OwnedChartSelector), Error> {
    let maximum_selector_bytes = source.semantic_limits().max_text_bytes();
    let mut selector_bytes = 0;
    let slide = OwnedSlideSelector::from_selector(
        slide_selector,
        maximum_selector_bytes,
        &mut selector_bytes,
    )
    .map_err(map_selector_error)?;
    let chart = OwnedChartSelector::from_selector(
        chart_selector,
        maximum_selector_bytes,
        &mut selector_bytes,
    )
    .map_err(map_selector_error)?;
    Ok((slide, chart))
}

fn positional_selection(
    package: &Package,
    expected: &ChartSelection,
) -> Result<ChartSelection, Error> {
    slide_chart_title::select_chart(
        package,
        SlideSelector::position(expected.slide_position),
        ChartSelector::index(expected.chart_position.get()),
        true,
    )
    .map_err(map_chart_title_error)
}

fn same_selection(expected: &ChartSelection, actual: &ChartSelection) -> bool {
    expected.slide_position == actual.slide_position
        && expected.chart_position == actual.chart_position
        && expected.slide_identifier == actual.slide_identifier
        && expected.chart_identifier == actual.chart_identifier
}

fn read_font(package: &Package, chart: u64) -> Result<ChartFont, Error> {
    let components = PackageComponents(package);
    let limits = graph_limits(package)?;
    let workspace = Workspace::new(&components, limits).map_err(map_graph_error)?;
    let value = engine::font(&workspace, chart).map_err(map_graph_error)?;
    Ok(ChartFont::new(value.font, value.bold, value.italic))
}

fn read_size(package: &Package, chart: u64) -> Result<TextPointSize, Error> {
    let components = PackageComponents(package);
    let limits = graph_limits(package)?;
    let workspace = Workspace::new(&components, limits).map_err(map_graph_error)?;
    engine::size(&workspace, chart).map_err(map_graph_error)
}

fn write_graph(
    package: &Package,
    chart: u64,
    change: &EngineChange,
) -> Result<(Package, bool, usize), Error> {
    let components = PackageComponents(package);
    let limits = graph_limits(package)?;
    let mut workspace = Workspace::new(&components, limits).map_err(map_graph_error)?;
    let changed = engine::set(&mut workspace, chart, change).map_err(map_graph_error)?;
    if !changed {
        return Ok((package.snapshot(), false, 0));
    }
    let changes = workspace.finish();
    let touched = changes.edits.len() + changes.insertions.len() + changes.deletions.len();
    if touched == 0 {
        return Err(Error::InvalidSource);
    }
    let candidate = publish(package, changes)?;
    Ok((candidate, true, touched))
}

fn graph_limits(package: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Refusal::UnsupportedSource.into()),
    }
}

fn publish(package: &Package, changes: graph::Changes) -> Result<Package, Error> {
    let catalog = physical_catalog(package)?;
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let mut edited = Vec::new();
    for (name, archive) in &changes.edits {
        let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
        edited.push((name.as_str(), compressed));
    }
    let edits: Vec<EntryEdit<'_>> = edited
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data.as_slice()))
        .collect();
    let names: Vec<&str> = catalog.package().iter().map(|entry| entry.name()).collect();
    let mut inserted = Vec::new();
    for insertion in &changes.insertions {
        let position = match insertion.before.as_deref() {
            Some(before) => Some(
                names
                    .iter()
                    .position(|name| *name == before)
                    .ok_or(Error::InvalidSource)?,
            ),
            None => None,
        };
        let bytes = insertion
            .archive
            .to_bytes(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
        inserted.push((insertion.name.as_str(), compressed, position));
    }
    let insertions: Vec<EntryInsertion<'_>> = inserted
        .iter()
        .map(|(name, data, position)| match position {
            Some(position) => EntryInsertion::at(*position, name, data.as_slice()),
            None => EntryInsertion::new(name, data.as_slice()),
        })
        .collect();
    let deletions: Vec<&str> = changes.deletions.iter().map(String::as_str).collect();
    let output = catalog
        .package()
        .reassemble_with_changes_to_bytes(&insertions, &edits, &deletions, catalog.limits())
        .map_err(map_archive_error)?;
    Package::from_source_with_options(output.into(), package.state.options).map_err(map_read_error)
}

fn map_selector_error(error: SelectorError) -> Error {
    match error {
        SelectorError::Limit { observed, maximum } => Error::Limit {
            kind: LimitKind::SelectorBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        SelectorError::Allocation { amount } => Error::Allocation { amount },
    }
}

fn map_chart_title_error(error: slide_chart_title::Error) -> Error {
    match error {
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::AmbiguousSelector) => {
            Refusal::AmbiguousSelector.into()
        },
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::EmptySlideName) => {
            Refusal::EmptySlideName.into()
        },
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::SlideNameNotFound) => {
            Refusal::SlideNameNotFound.into()
        },
        slide_chart_title::Error::Refused(
            crate::chart::title::Refusal::SlidePositionNotFound { position },
        ) => Refusal::SlidePositionNotFound { position }.into(),
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::ChartNameNotFound) => {
            Refusal::ChartNameNotFound.into()
        },
        slide_chart_title::Error::Refused(
            crate::chart::title::Refusal::ChartPositionNotFound { position },
        ) => Refusal::ChartPositionNotFound { position }.into(),
        slide_chart_title::Error::Refused(crate::chart::title::Refusal::EmptyChartName) => {
            Refusal::EmptyChartName.into()
        },
        slide_chart_title::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_chart_title_limit(kind),
            observed,
            maximum,
        },
        slide_chart_title::Error::Allocation { amount } => Error::Allocation { amount },
        slide_chart_title::Error::InvalidSource
        | slide_chart_title::Error::Verification
        | slide_chart_title::Error::PatchConflict => Error::InvalidSource,
    }
}

fn map_chart_title_limit(kind: slide_chart_title::LimitKind) -> LimitKind {
    use slide_chart_title::LimitKind as K;
    match kind {
        K::InputBytes => LimitKind::InputBytes,
        K::OutputBytes => LimitKind::OutputBytes,
        K::WireBytes => LimitKind::WireBytes,
        K::Entries => LimitKind::Entries,
        K::EntryBytes => LimitKind::EntryBytes,
        K::TotalBytes => LimitKind::TotalBytes,
        K::Slides => LimitKind::Slides,
        K::References => LimitKind::References,
        K::TextStorages => LimitKind::Entries,
        K::TextFragments | K::TextBytes => LimitKind::WireBytes,
        K::WireFields => LimitKind::WireFields,
        K::WireNesting => LimitKind::WireNesting,
        K::WireWork => LimitKind::WireWork,
        K::TitleBytes | K::SelectorBytes => LimitKind::SelectorBytes,
    }
}

fn map_graph_error(error: graph::Error) -> Error {
    match error {
        graph::Error::Archive(error) => map_core_error(error),
        graph::Error::Wire(error) => map_wire_error(error),
        graph::Error::Codec(_) | graph::Error::Registry(_) => Error::InvalidSource,
        graph::Error::MissingObject(_) | graph::Error::AmbiguousObject(_) => {
            Refusal::UnsupportedStyle.into()
        },
        graph::Error::Invalid(_) => Refusal::UnsupportedStyle.into(),
        graph::Error::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
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
                SemanticLimitKind::TextStorages => LimitKind::Entries,
                SemanticLimitKind::TextFragments => LimitKind::Entries,
                SemanticLimitKind::TextBytes => LimitKind::WireBytes,
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
                PayloadLimitKind::Bytes => LimitKind::WireBytes,
                PayloadLimitKind::Fields => LimitKind::WireFields,
                PayloadLimitKind::Nesting => LimitKind::WireNesting,
                PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(_)
        | ReadError::Detection(_)
        | ReadError::UnexpectedFormat { .. }
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_)
        | ReadError::Io(_) => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

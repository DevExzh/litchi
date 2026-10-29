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

use std::collections::HashSet;
use std::fmt;
use std::mem::size_of;
use std::sync::{Arc, OnceLock};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, EntryInsertion, ExactArtifacts},
};
#[cfg(test)]
use litchi_iwa_core::Limits as CoreLimits;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_core::{
    ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor,
};
use litchi_iwa_protos::chart_style_codec::Limits as ChartCodecLimits;
use litchi_iwa_style::{
    chart_font::{self as engine, Change as EngineChange},
    graph::{self, Components, Workspace},
};
pub use litchi_iwa_text::character::TextPointSize;
pub use litchi_iwa_text::font::{Font, NameError as FontNameError};
use thiserror::Error as ThisError;

use super::chart_axis_support::{
    self, AxisSupportBudget, AxisSupportError, AxisSupportSelectorError,
};
use super::edit::{OwnedChartSelector, OwnedSlideSelector, SelectorError};
use super::slide_chart_title::{self, ChartSelection};
use super::slide_table_core::{Budget as TransactionBudget, Error as TransactionError};
use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

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
    /// Font-name bytes retained by the chart graph transaction.
    TextBytes,
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
            Self::TextBytes => "text bytes",
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

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct MutationPlan {
    root_identifier: u64,
    members: Box<[String]>,
    deleted_members: Box<[String]>,
    object_ids: Box<[u64]>,
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
    plan: MutationPlan,
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
            plan: self.plan.clone(),
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
    plan: MutationPlan,
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
            plan: self.plan.clone(),
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

#[derive(Clone, PartialEq)]
enum ResetExpectation {
    Font(ChartFont),
    Size(TextPointSize),
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
        let mut budget = transaction_budget(self.source)?;
        let mut fields = FieldLedger::new(self.source)?;
        let mut text = TextLedger::new(self.source);
        admit_transaction(self.source, &mut budget)?;
        self.resolve_with_budget(&mut budget, &mut fields, &mut text)
    }

    fn resolve_with_budget(
        &self,
        budget: &mut TransactionBudget,
        fields: &mut FieldLedger,
        text: &mut TextLedger,
    ) -> Result<&(ChartSelection, ChartFont), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let selection = select_chart_transaction(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
            budget,
            fields,
        )?;
        let before = read_font_with_budget(
            self.source,
            selection.chart_identifier,
            budget,
            fields,
            text,
        )?;
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
        let mut budget = transaction_budget(self.source)?;
        let mut fields = FieldLedger::new(self.source)?;
        let mut text = TextLedger::new(self.source);
        admit_transaction(self.source, &mut budget)?;
        let (source_selection, before) = self
            .resolve_with_budget(&mut budget, &mut fields, &mut text)?
            .clone();
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
                        plan: MutationPlan::default(),
                    },
                    diagnostics: Diagnostics::unchanged(),
                });
            }
        }
        if !catalog.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        self.source.validate().map_err(map_read_error)?;
        let current =
            positional_selection(self.source, &source_selection, &mut budget, &mut fields)?;
        if !same_selection(&source_selection, &current)
            || read_font_with_budget(
                self.source,
                current.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != before
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
        let (candidate, changed, touched_components, reset_expectation, plan) = write_graph(
            self.source,
            current.chart_identifier,
            &change,
            &mut budget,
            &mut fields,
            &mut text,
        )?;
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
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate_selection =
            positional_selection(&candidate, &source_selection, &mut budget, &mut fields)?;
        let after = read_font_with_budget(
            &candidate,
            candidate_selection.chart_identifier,
            &mut budget,
            &mut fields,
            &mut text,
        )?;
        if !same_selection(&source_selection, &candidate_selection)
            || expected.as_ref().is_some_and(|value| after != *value)
            || reset_expectation.as_ref().is_some_and(
                |value| !matches!(value, ResetExpectation::Font(expected) if after == *expected),
            )
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
                plan,
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
        let mut budget = transaction_budget(self.source)?;
        let mut fields = FieldLedger::new(self.source)?;
        let mut text = TextLedger::new(self.source);
        admit_transaction(self.source, &mut budget)?;
        self.resolve_with_budget(&mut budget, &mut fields, &mut text)
    }

    fn resolve_with_budget(
        &self,
        budget: &mut TransactionBudget,
        fields: &mut FieldLedger,
        text: &mut TextLedger,
    ) -> Result<&(ChartSelection, TextPointSize), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let selection = select_chart_transaction(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
            budget,
            fields,
        )?;
        let before = read_size_with_budget(
            self.source,
            selection.chart_identifier,
            budget,
            fields,
            text,
        )?;
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
        let mut budget = transaction_budget(self.source)?;
        let mut fields = FieldLedger::new(self.source)?;
        let mut text = TextLedger::new(self.source);
        admit_transaction(self.source, &mut budget)?;
        let (source_selection, before) = self
            .resolve_with_budget(&mut budget, &mut fields, &mut text)?
            .clone();
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
                        plan: MutationPlan::default(),
                    },
                    diagnostics: Diagnostics::unchanged(),
                });
            }
        }
        if !catalog.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        self.source.validate().map_err(map_read_error)?;
        let current =
            positional_selection(self.source, &source_selection, &mut budget, &mut fields)?;
        if !same_selection(&source_selection, &current)
            || read_size_with_budget(
                self.source,
                current.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != before
        {
            return Err(Error::InvalidSource);
        }
        let change = match operation {
            FontSizeOperation::Set(value) => EngineChange::Size(value),
            FontSizeOperation::Reset => EngineChange::ResetSize,
        };
        let (candidate, changed, touched_components, reset_expectation, plan) = write_graph(
            self.source,
            current.chart_identifier,
            &change,
            &mut budget,
            &mut fields,
            &mut text,
        )?;
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
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate_selection =
            positional_selection(&candidate, &source_selection, &mut budget, &mut fields)?;
        let after = read_size_with_budget(
            &candidate,
            candidate_selection.chart_identifier,
            &mut budget,
            &mut fields,
            &mut text,
        )?;
        if !same_selection(&source_selection, &candidate_selection)
            || reset_expectation.as_ref().is_some_and(
                |value| !matches!(value, ResetExpectation::Size(expected) if after == *expected),
            )
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
                plan,
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
        let mut budget = transaction_budget(self)?;
        let mut fields = FieldLedger::new(self)?;
        let mut text = TextLedger::new(self);
        admit_transaction(self, &mut budget)?;
        let selection = select_chart_transaction(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
            &mut budget,
            &mut fields,
        )?;
        read_font_with_budget(
            self,
            selection.chart_identifier,
            &mut budget,
            &mut fields,
            &mut text,
        )
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
        let mut budget = transaction_budget(self)?;
        let mut fields = FieldLedger::new(self)?;
        let mut text = TextLedger::new(self);
        admit_transaction(self, &mut budget)?;
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
            &mut budget,
            &mut fields,
        )?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
            || read_font_with_budget(
                self,
                selection.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != patch.before
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
        preflight_candidate(self, patch.artifacts.target().len(), 0, &mut budget)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        verify_locality_members(self, &candidate, &patch.plan, &mut budget)?;
        let candidate_selection =
            positional_selection(&candidate, &selection, &mut budget, &mut fields)?;
        if !same_selection(&selection, &candidate_selection)
            || read_font_with_budget(
                &candidate,
                candidate_selection.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != patch.after
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
        let mut budget = transaction_budget(self)?;
        let mut fields = FieldLedger::new(self)?;
        let mut text = TextLedger::new(self);
        admit_transaction(self, &mut budget)?;
        let selection = select_chart_transaction(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
            &mut budget,
            &mut fields,
        )?;
        read_size_with_budget(
            self,
            selection.chart_identifier,
            &mut budget,
            &mut fields,
            &mut text,
        )
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
        let mut budget = transaction_budget(self)?;
        let mut fields = FieldLedger::new(self)?;
        let mut text = TextLedger::new(self);
        admit_transaction(self, &mut budget)?;
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
            &mut budget,
            &mut fields,
        )?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
            || read_size_with_budget(
                self,
                selection.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != patch.before
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
        preflight_candidate(self, patch.artifacts.target().len(), 0, &mut budget)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        verify_locality_members(self, &candidate, &patch.plan, &mut budget)?;
        let candidate_selection =
            positional_selection(&candidate, &selection, &mut budget, &mut fields)?;
        if !same_selection(&selection, &candidate_selection)
            || read_size_with_budget(
                &candidate,
                candidate_selection.chart_identifier,
                &mut budget,
                &mut fields,
                &mut text,
            )? != patch.after
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

fn select_chart_transaction(
    package: &Package,
    slide: SlideSelector<'_>,
    chart: ChartSelector<'_>,
    mutation_guards: bool,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<ChartSelection, Error> {
    let mut graph_budget = ChartFontSelectionBudget { budget, fields };
    slide_chart_title::select_chart_with_budget(
        package,
        slide,
        chart,
        mutation_guards,
        &mut graph_budget,
    )
    .map_err(map_axis_error)
}

struct ChartFontSelectionBudget<'a> {
    budget: &'a mut TransactionBudget,
    fields: &'a mut FieldLedger,
}

fn map_selection_budget_result(
    operation: Result<(), TransactionError>,
) -> Result<(), AxisSupportError> {
    operation
        .map_err(map_transaction_error)
        .map_err(map_selection_budget_error)
}

impl AxisSupportBudget for ChartFontSelectionBudget<'_> {
    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), AxisSupportError> {
        let passes = usize::from(mutation_guards) + 1;
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_mul(passes))
            .ok_or(AxisSupportError::InvalidSource)?;
        map_selection_budget_result(self.budget.work(amount))
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        map_selection_budget_result(self.budget.input(amount))
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError> {
        let field_capacity = payload.max(1);
        let bytes = payload
            .checked_add(
                payload
                    .checked_mul(size_of::<litchi_iwa_common::wire::WireField>())
                    .ok_or(AxisSupportError::InvalidSource)?,
            )
            .ok_or(AxisSupportError::InvalidSource)?;
        map_selection_budget_result(
            self.budget
                .work(bytes)
                .and_then(|_| self.budget.scratch(bytes))
                .and_then(|_| self.budget.allocations(field_capacity)),
        )
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError> {
        self.fields
            .charge(fields)
            .map_err(map_selection_budget_error)?;
        map_selection_budget_result(self.budget.work(fields))
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError> {
        let bytes = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(AxisSupportError::InvalidSource)?;
        map_selection_budget_result(
            self.budget
                .scratch(bytes.max(1))
                .and_then(|_| self.budget.allocations(capacity.max(1))),
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        map_selection_budget_result(self.budget.references(amount))
    }

    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), AxisSupportError> {
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_add(retained_vectors))
            .ok_or(AxisSupportError::InvalidSource)?;
        map_selection_budget_result(self.budget.work(amount))
    }

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError> {
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_mul(2))
            .ok_or(AxisSupportError::InvalidSource)?;
        map_selection_budget_result(self.budget.work(amount))
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        map_selection_budget_result(self.budget.work(amount))
    }

    fn metadata_options(
        &self,
        package: &Package,
    ) -> Result<litchi_iwa_protos::package_metadata_codec::RewriteOptions, AxisSupportError> {
        let wire = self
            .budget
            .residual(package)
            .map_err(map_transaction_error)
            .map_err(map_selection_budget_error)?;
        let recursion =
            u32::try_from(wire.max_nesting()).map_err(|_| AxisSupportError::InvalidSource)?;
        Ok(
            litchi_iwa_protos::package_metadata_codec::RewriteOptions::new(
                wire.max_input_bytes(),
                wire.max_output_bytes(),
                self.fields
                    .remaining()
                    .map_err(map_selection_budget_error)?,
                wire.max_rewrite_work(),
                recursion,
                package.state.source.components().len().max(1),
                self.budget
                    .remaining_references()
                    .map_err(map_transaction_error)
                    .map_err(map_selection_budget_error)?,
                self.budget
                    .remaining_allocations()
                    .map_err(map_transaction_error)
                    .map_err(map_selection_budget_error)?,
            ),
        )
    }

    fn charge_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    ) -> Result<(), AxisSupportError> {
        map_selection_budget_result(
            self.budget
                .input(report.input_bytes())
                .and_then(|_| self.budget.output(report.output_bytes()))
                .and_then(|_| self.budget.work(report.work_bytes()))
                .and_then(|_| self.budget.components(report.components_scanned()))
                .and_then(|_| {
                    self.budget.references(
                        report
                            .references_scanned()
                            .checked_add(report.source_references_scanned())
                            .ok_or(TransactionError::InvalidSource)?,
                    )
                })
                .and_then(|_| self.budget.allocations(report.allocations()))
                .and_then(|_| self.budget.retained(report.retained_bytes()))
                .and_then(|_| self.budget.scratch(report.scratch_bytes()))
                .and_then(|_| self.budget.nesting(report.max_depth() as usize)),
        )
        .and_then(|_| {
            self.fields
                .charge(report.fields())
                .map_err(map_selection_budget_error)
        })
    }
}

fn map_selection_budget_error(error: Error) -> AxisSupportError {
    match error {
        Error::Limit {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                LimitKind::InputBytes => chart_axis_support::AxisSupportLimitKind::InputBytes,
                LimitKind::OutputBytes => chart_axis_support::AxisSupportLimitKind::OutputBytes,
                LimitKind::WireBytes => chart_axis_support::AxisSupportLimitKind::WireBytes,
                LimitKind::Entries => chart_axis_support::AxisSupportLimitKind::Entries,
                LimitKind::EntryBytes => chart_axis_support::AxisSupportLimitKind::EntryBytes,
                LimitKind::TotalBytes => chart_axis_support::AxisSupportLimitKind::TotalBytes,
                LimitKind::Slides => chart_axis_support::AxisSupportLimitKind::Slides,
                LimitKind::References => chart_axis_support::AxisSupportLimitKind::References,
                LimitKind::WireFields => chart_axis_support::AxisSupportLimitKind::WireFields,
                LimitKind::WireNesting => chart_axis_support::AxisSupportLimitKind::WireNesting,
                LimitKind::WireWork => chart_axis_support::AxisSupportLimitKind::WireWork,
                LimitKind::TextBytes => chart_axis_support::AxisSupportLimitKind::TextBytes,
                _ => chart_axis_support::AxisSupportLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        Error::Allocation { amount } => AxisSupportError::Allocation { amount },
        _ => AxisSupportError::InvalidSource,
    }
}

fn positional_selection(
    package: &Package,
    expected: &ChartSelection,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<ChartSelection, Error> {
    select_chart_transaction(
        package,
        SlideSelector::position(expected.slide_position),
        ChartSelector::index(expected.chart_position.get()),
        true,
        budget,
        fields,
    )
}

fn same_selection(expected: &ChartSelection, actual: &ChartSelection) -> bool {
    expected.slide_position == actual.slide_position
        && expected.chart_position == actual.chart_position
        && expected.slide_identifier == actual.slide_identifier
        && expected.chart_identifier == actual.chart_identifier
}

fn read_font_with_budget(
    package: &Package,
    chart: u64,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    text: &mut TextLedger,
) -> Result<ChartFont, Error> {
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    let workspace = Workspace::new(&components, graph_limits(package)?).map_err(map_graph_error)?;
    let (value, report) = engine::font(
        &workspace,
        chart,
        engine_limits(package, budget, fields, text)?,
    )
    .map_err(map_engine_error)?;
    charge_engine_report(budget, fields, text, report)?;
    Ok(ChartFont::new(value.font, value.bold, value.italic))
}

fn read_size_with_budget(
    package: &Package,
    chart: u64,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    text: &mut TextLedger,
) -> Result<TextPointSize, Error> {
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    let workspace = Workspace::new(&components, graph_limits(package)?).map_err(map_graph_error)?;
    let (value, report) = engine::size(
        &workspace,
        chart,
        engine_limits(package, budget, fields, text)?,
    )
    .map_err(map_engine_error)?;
    charge_engine_report(budget, fields, text, report)?;
    Ok(value)
}

fn write_graph(
    package: &Package,
    chart: u64,
    change: &EngineChange,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    text: &mut TextLedger,
) -> Result<(Package, bool, usize, Option<ResetExpectation>, MutationPlan), Error> {
    admit_workspace(package, budget)?;
    let components = PackageComponents(package);
    let mut workspace =
        Workspace::new(&components, graph_limits(package)?).map_err(map_graph_error)?;
    let report = engine::set(
        &mut workspace,
        chart,
        change,
        engine_limits(package, budget, fields, text)?,
    )
    .map_err(map_engine_error)?;
    charge_engine_report(budget, fields, text, report)?;
    if !report.changed() {
        return Ok((package.snapshot(), false, 0, None, MutationPlan::default()));
    }
    let reset_expectation = match change {
        EngineChange::ResetFont => {
            let (value, report) = engine::font(
                &workspace,
                chart,
                engine_limits(package, budget, fields, text)?,
            )
            .map_err(map_engine_error)?;
            charge_engine_report(budget, fields, text, report)?;
            Some(ResetExpectation::Font(ChartFont::new(
                value.font,
                value.bold,
                value.italic,
            )))
        },
        EngineChange::ResetSize => {
            let (value, report) = engine::size(
                &workspace,
                chart,
                engine_limits(package, budget, fields, text)?,
            )
            .map_err(map_engine_error)?;
            charge_engine_report(budget, fields, text, report)?;
            Some(ResetExpectation::Size(value))
        },
        _ => None,
    };
    let changes = workspace.finish();
    let touched = changes.edits.len() + changes.insertions.len() + changes.deletions.len();
    if touched == 0 {
        return Err(Error::InvalidSource);
    }
    let plan = mutation_plan(package, chart, &changes, budget)?;
    let candidate = publish(package, changes, budget, &plan)?;
    Ok((candidate, true, touched, reset_expectation, plan))
}

fn transaction_budget(package: &Package) -> Result<TransactionBudget, Error> {
    TransactionBudget::new(package).map_err(map_transaction_error)
}

#[derive(Debug, Clone, Copy)]
struct FieldLedger {
    used: usize,
    maximum: usize,
}

impl FieldLedger {
    fn new(package: &Package) -> Result<Self, Error> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        // Selector resolution, graph mutation, reset readback, candidate
        // reopening, and inverse/locality proofs share this bounded field
        // ledger. Keep the multiplier finite and below the common wire hard
        // ceiling; each phase still charges its actual report.
        let maximum = wire
            .max_fields()
            .checked_mul(256)
            .ok_or(Error::InvalidSource)?
            .min(litchi_iwa_common::WireLimits::MAX_FIELDS);
        Ok(Self { used: 0, maximum })
    }

    fn remaining(self) -> Result<usize, Error> {
        self.maximum
            .checked_sub(self.used)
            .filter(|remaining| *remaining > 0)
            .ok_or(Error::Limit {
                kind: LimitKind::WireFields,
                observed: usize_to_u64(self.used.saturating_add(1)),
                maximum: usize_to_u64(self.maximum),
            })
    }

    fn charge(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self.used.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::WireFields,
            observed: u64::MAX,
            maximum: usize_to_u64(self.maximum),
        })?;
        if observed > self.maximum {
            return Err(Error::Limit {
                kind: LimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.maximum),
            });
        }
        self.used = observed;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct TextLedger {
    used: usize,
    maximum: usize,
}

impl TextLedger {
    fn new(package: &Package) -> Self {
        Self {
            used: 0,
            maximum: package.semantic_limits().max_text_bytes(),
        }
    }

    fn remaining(self) -> Result<usize, Error> {
        self.maximum
            .checked_sub(self.used)
            .filter(|remaining| *remaining > 0)
            .ok_or(Error::Limit {
                kind: LimitKind::TextBytes,
                observed: usize_to_u64(self.used.saturating_add(1)),
                maximum: usize_to_u64(self.maximum),
            })
    }

    fn charge(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self.used.checked_add(amount).ok_or(Error::Limit {
            kind: LimitKind::TextBytes,
            observed: u64::MAX,
            maximum: usize_to_u64(self.maximum),
        })?;
        if observed > self.maximum {
            return Err(Error::Limit {
                kind: LimitKind::TextBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.maximum),
            });
        }
        self.used = observed;
        Ok(())
    }
}

/// Admit the package inventory before selector/graph work starts.  The chart
/// engine receives only the residual profile from this same transaction
/// ledger; it never gets a source-sized default profile.
fn admit_transaction(package: &Package, budget: &mut TransactionBudget) -> Result<(), Error> {
    budget
        .preflight_semantic_scan(package)
        .map_err(map_transaction_error)
}

/// Charge the graph index envelope before constructing `Workspace`.  The
/// shared workspace uses borrowed archives but owns identifier and ambiguity
/// indexes, so those collections must be admitted at this package boundary.
fn admit_workspace(package: &Package, budget: &mut TransactionBudget) -> Result<(), Error> {
    let components = package.state.source.components();
    let component_count = components.len();
    let object_count = package.state.total_objects;
    let index_units = object_count
        .checked_mul(4)
        .and_then(|value| value.checked_add(component_count))
        .ok_or(Error::InvalidSource)?;
    let index_bytes = index_units
        .checked_mul(size_of::<(u64, usize)>() * 2)
        .ok_or(Error::InvalidSource)?;
    let source_work = component_count
        .checked_add(object_count)
        .ok_or(Error::InvalidSource)?;
    budget
        .components(component_count)
        .and_then(|_| budget.payload_objects(object_count))
        .and_then(|_| budget.work(source_work))
        .and_then(|_| budget.allocations(index_units))
        .and_then(|_| budget.retained(index_bytes))
        .and_then(|_| budget.scratch(index_bytes))
        .map_err(map_transaction_error)?;
    Ok(())
}

fn engine_limits(
    package: &Package,
    budget: &TransactionBudget,
    fields: &FieldLedger,
    text: &TextLedger,
) -> Result<engine::Limits, Error> {
    let wire = budget.residual(package).map_err(map_transaction_error)?;
    let max_nesting = u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let codec = residual_chart_codec_limits(wire, fields.remaining()?)?;
    Ok(engine::Limits::new(
        wire.max_input_bytes().min(codec.max_bytes()),
        fields
            .remaining()?
            .min(codec.max_fields())
            .min(litchi_iwa_common::WireLimits::MAX_FIELDS),
        wire.max_output_bytes().min(codec.max_bytes()),
        max_nesting,
        wire.max_rewrite_work(),
        text.remaining()?,
        budget
            .remaining_allocations()
            .map_err(map_transaction_error)?,
        budget.remaining_retained().map_err(map_transaction_error)?,
        budget.remaining_scratch().map_err(map_transaction_error)?,
    ))
}

fn residual_chart_codec_limits(
    wire: litchi_iwa_common::WireLimits,
    field_budget: usize,
) -> Result<ChartCodecLimits, Error> {
    const HARD_MAX_BYTES: usize = 256 * 1024 * 1024;
    const HARD_MAX_FIELDS: usize = 16 * 1024 * 1024;
    let bytes = wire
        .max_input_bytes()
        .min(wire.max_output_bytes())
        .clamp(1, HARD_MAX_BYTES);
    let fields = field_budget.clamp(1, HARD_MAX_FIELDS);
    ChartCodecLimits::default()
        .with_max_bytes(bytes)
        .map_err(map_codec_error)
        .and_then(|limits| limits.with_max_fields(fields).map_err(map_codec_error))
}

fn charge_engine_report(
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    text: &mut TextLedger,
    report: engine::Report,
) -> Result<(), Error> {
    budget
        .input(report.input_bytes())
        .and_then(|_| budget.output(report.output_bytes()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .and_then(|_| budget.nesting(report.max_depth() as usize))
        .map_err(map_transaction_error)?;
    fields.charge(report.fields())?;
    text.charge(report.text_bytes())
}

fn preflight_candidate(
    package: &Package,
    candidate_bytes: usize,
    inserted_components: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let objects = package
        .state
        .total_objects
        .checked_add(inserted_components)
        .ok_or(Error::InvalidSource)?;
    let messages = package
        .state
        .source
        .components()
        .iter()
        .map(|component| {
            component
                .archive()
                .objects
                .iter()
                .map(|object| object.messages.len())
                .try_fold(0usize, |sum, count| sum.checked_add(count))
                .ok_or(Error::InvalidSource)
        })
        .try_fold(0usize, |sum, count| {
            count.and_then(|count| sum.checked_add(count).ok_or(Error::InvalidSource))
        })?;
    let work = candidate_bytes
        .checked_add(objects)
        .and_then(|value| value.checked_add(messages))
        .ok_or(Error::InvalidSource)?;
    let allocations = objects
        .checked_add(messages)
        .and_then(|value| value.checked_add(inserted_components))
        .and_then(|value| value.checked_add(2))
        .ok_or(Error::InvalidSource)?;
    budget
        .input(candidate_bytes)
        .and_then(|_| budget.output(candidate_bytes))
        .and_then(|_| budget.work(work))
        .and_then(|_| budget.allocations(allocations))
        .and_then(|_| budget.retained(candidate_bytes))
        .and_then(|_| budget.scratch(candidate_bytes))
        .map_err(map_transaction_error)
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

fn admit_archive_encoding(
    archive: &Archive,
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let encoded = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let compressed =
        SnappyStream::maximum_compressed_len(encoded, snappy_limits).map_err(map_core_error)?;
    let work = encoded
        .checked_add(compressed)
        .ok_or(Error::InvalidSource)?;
    budget
        .work(work)
        .and_then(|_| budget.retained(work))
        .and_then(|_| budget.scratch(work))
        .and_then(|_| budget.allocations(2))
        .map_err(map_transaction_error)
}

fn package_object(package: &Package, identifier: u64) -> Option<&ArchiveObject> {
    package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| component.archive().object(identifier))
}

#[derive(Default)]
struct ObjectReferenceCollector {
    identifiers: Vec<u64>,
}

impl ArchiveReferenceVisitor for ObjectReferenceCollector {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object && occurrence.referenced_identifier != 0
        {
            self.identifiers.push(occurrence.referenced_identifier);
        }
        Ok(())
    }
}

struct StrictMetadataVisitor;

impl ArchiveReferenceVisitor for StrictMetadataVisitor {
    fn visit_reference(
        &mut self,
        _occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        Ok(())
    }
}

fn inspect_archive_metadata_strict(
    package: &Package,
    object: &ArchiveObject,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let header = usize::try_from(object.header_length).map_err(|_| Error::InvalidSource)?;
    budget
        .work(
            header
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|_| budget.allocations(1))
        .and_then(|_| budget.scratch(header))
        .map_err(map_transaction_error)?;
    let mut visitor = StrictMetadataVisitor;
    object
        .inspect_references_with_policy(
            &mut visitor,
            ArchiveReferencePolicy::RejectUnknownMetadata,
            graph_limits(package)?,
        )
        .map(|_| ())
        .map_err(map_core_error)
}

fn parse_locality_fields(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Vec<litchi_iwa_common::wire::WireField>, Error> {
    let span = payload.len().max(1);
    let scratch = span
        .checked_mul(size_of::<litchi_iwa_common::wire::WireField>())
        .ok_or(Error::InvalidSource)?;
    budget
        .input(payload.len())
        .and_then(|_| budget.work(span))
        .and_then(|_| budget.scratch(scratch))
        .and_then(|_| budget.allocations(1))
        .map_err(map_transaction_error)?;
    let fields =
        litchi_iwa_common::wire::parse_wire_fields(payload, limits).map_err(map_wire_error)?;
    Ok(fields)
}

fn raw_fields_except<'a>(
    fields: &[litchi_iwa_common::wire::WireField],
    payload: &'a [u8],
    excluded: u32,
    budget: &mut TransactionBudget,
) -> Result<Vec<&'a [u8]>, Error> {
    let capacity = fields.len();
    let scratch = capacity
        .checked_mul(size_of::<&'a [u8]>())
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(1)
        .and_then(|_| budget.scratch(scratch))
        .map_err(map_transaction_error)?;
    let mut raw = Vec::new();
    raw.try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: capacity })?;
    for field in fields.iter().filter(|field| field.number() != excluded) {
        raw.push(field.raw(payload).map_err(map_wire_error)?);
    }
    Ok(raw)
}

fn raw_fields_except_numbers<'a>(
    fields: &[litchi_iwa_common::wire::WireField],
    payload: &'a [u8],
    excluded: &[u32],
    budget: &mut TransactionBudget,
) -> Result<Vec<&'a [u8]>, Error> {
    let capacity = fields.len();
    let scratch = capacity
        .checked_mul(size_of::<&'a [u8]>())
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(1)
        .and_then(|_| budget.scratch(scratch))
        .map_err(map_transaction_error)?;
    let mut raw = Vec::new();
    raw.try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: capacity })?;
    for field in fields
        .iter()
        .filter(|field| !excluded.contains(&field.number()))
    {
        raw.push(field.raw(payload).map_err(map_wire_error)?);
    }
    Ok(raw)
}

fn wire_reference_id(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Option<u64>, Error> {
    let fields = parse_locality_fields(payload, limits, budget)?;
    let Some(field) = fields.iter().rev().find(|field| field.number() == 1) else {
        return Ok(None);
    };
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    let bytes = field.payload(payload).map_err(map_wire_error)?;
    let (value, width) =
        litchi_iwa_common::decode_varint_from_bytes(bytes).map_err(|_| Error::InvalidSource)?;
    (width == bytes.len())
        .then_some(Some(value))
        .ok_or(Error::InvalidSource)
}

fn payload_object_references(
    message: &RawMessage,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Vec<u64>, Error> {
    if !matches!(
        message.type_,
        CHART_MESSAGE_TYPE | PARAGRAPH_STYLE_MESSAGE_TYPE | STYLESHEET_MESSAGE_TYPE
    ) {
        return Ok(Vec::new());
    }
    let reference_capacity = message.data.len().max(1);
    let reference_scratch = reference_capacity
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget
        .allocations(1)
        .and_then(|_| budget.scratch(reference_scratch))
        .map_err(map_transaction_error)?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(reference_capacity)
        .map_err(|_| Error::Allocation {
            amount: reference_capacity,
        })?;
    if message.type_ == CHART_MESSAGE_TYPE {
        let fields = parse_locality_fields(&message.data, limits, budget)?;
        for extension in fields.iter().filter(|field| field.number() == 10_000) {
            let extension = extension.payload(&message.data).map_err(map_wire_error)?;
            let extension_fields = parse_locality_fields(extension, limits, budget)?;
            for styles in extension_fields.iter().filter(|field| field.number() == 20) {
                let reference = styles.payload(extension).map_err(map_wire_error)?;
                if let Some(identifier) = wire_reference_id(reference, limits, budget)? {
                    references.push(identifier);
                }
            }
        }
        return Ok(references);
    }
    let fields = parse_locality_fields(&message.data, limits, budget)?;
    match message.type_ {
        PARAGRAPH_STYLE_MESSAGE_TYPE => {
            for field in fields.iter().filter(|field| field.number() == 1) {
                let parent = field.payload(&message.data).map_err(map_wire_error)?;
                let parent_fields = parse_locality_fields(parent, limits, budget)?;
                for link in parent_fields
                    .iter()
                    .filter(|field| matches!(field.number(), 3 | 5))
                {
                    if link.wire_type() != 2 {
                        return Err(Error::InvalidSource);
                    }
                    if let Some(identifier) = wire_reference_id(
                        link.payload(parent).map_err(map_wire_error)?,
                        limits,
                        budget,
                    )? {
                        references.push(identifier);
                    }
                }
            }
        },
        STYLESHEET_MESSAGE_TYPE => {
            for field in fields.iter().filter(|field| field.number() == 1) {
                if let Some(identifier) = wire_reference_id(
                    field.payload(&message.data).map_err(map_wire_error)?,
                    limits,
                    budget,
                )? {
                    references.push(identifier);
                }
            }
            for field in fields.iter().filter(|field| field.number() == 5) {
                let payload = field.payload(&message.data).map_err(map_wire_error)?;
                let entries = parse_locality_fields(payload, limits, budget)?;
                for entry in entries
                    .iter()
                    .filter(|entry| entry.number() == 1 || entry.number() == 2)
                {
                    if entry.wire_type() != 2 {
                        return Err(Error::InvalidSource);
                    }
                    if let Some(identifier) = wire_reference_id(
                        entry.payload(payload).map_err(map_wire_error)?,
                        limits,
                        budget,
                    )? {
                        references.push(identifier);
                    }
                }
            }
        },
        _ => {},
    }
    Ok(references)
}

fn reachable_objects(
    package: &Package,
    root_identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<HashSet<u64>, Error> {
    let object_capacity = package.state.total_objects.max(1);
    let mut payload_bytes = 0usize;
    let mut metadata_references = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                payload_bytes = payload_bytes
                    .checked_add(message.data.len())
                    .ok_or(Error::InvalidSource)?;
            }
            for message in &object.archive_info.message_infos {
                metadata_references = metadata_references
                    .checked_add(message.object_references.len())
                    .ok_or(Error::InvalidSource)?;
                for field in &message.field_infos {
                    metadata_references = metadata_references
                        .checked_add(field.object_references.len())
                        .ok_or(Error::InvalidSource)?;
                }
            }
        }
    }
    let pending_capacity = object_capacity
        .checked_add(payload_bytes)
        .and_then(|value| value.checked_add(metadata_references))
        .ok_or(Error::InvalidSource)?;
    let scratch = object_capacity
        .checked_add(pending_capacity)
        .and_then(|value| value.checked_mul(size_of::<u64>()))
        .ok_or(Error::InvalidSource)?;
    budget
        .work(
            object_capacity
                .checked_add(payload_bytes)
                .and_then(|value| value.checked_add(metadata_references))
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|_| budget.allocations(2))
        .and_then(|_| budget.scratch(scratch))
        .map_err(map_transaction_error)?;
    let mut pending = Vec::new();
    pending
        .try_reserve_exact(pending_capacity)
        .map_err(|_| Error::Allocation {
            amount: pending_capacity,
        })?;
    pending.push(root_identifier);
    let mut reachable = HashSet::new();
    reachable
        .try_reserve(object_capacity)
        .map_err(|_| Error::Allocation {
            amount: object_capacity,
        })?;
    let limits = graph_limits(package)?;
    while let Some(identifier) = pending.pop() {
        if !reachable.insert(identifier) {
            continue;
        }
        let object = match package_object(package, identifier) {
            Some(object) => object,
            None => return Err(Error::InvalidSource),
        };
        let reference_capacity = object
            .archive_info
            .message_infos
            .iter()
            .map(|message| {
                message
                    .object_references
                    .len()
                    .checked_add(
                        message
                            .field_infos
                            .iter()
                            .map(|field| field.object_references.len())
                            .try_fold(0usize, |sum, count| sum.checked_add(count))
                            .ok_or(Error::InvalidSource)?,
                    )
                    .ok_or(Error::InvalidSource)
            })
            .try_fold(0usize, |sum, count| {
                count.and_then(|count| sum.checked_add(count).ok_or(Error::InvalidSource))
            })?;
        let work = object
            .messages
            .len()
            .checked_add(reference_capacity)
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::InvalidSource)?;
        let scratch = reference_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?;
        budget
            .work(work)
            .and_then(|_| budget.allocations(1))
            .and_then(|_| budget.scratch(scratch))
            .map_err(map_transaction_error)?;
        budget
            .allocations(1)
            .and_then(|_| {
                budget.scratch(
                    reference_capacity
                        .checked_mul(size_of::<u64>())
                        .ok_or(TransactionError::InvalidSource)?,
                )
            })
            .map_err(map_transaction_error)?;
        let mut references = ObjectReferenceCollector::default();
        references
            .identifiers
            .try_reserve_exact(reference_capacity)
            .map_err(|_| Error::Allocation {
                amount: reference_capacity,
            })?;
        object
            .inspect_references_with_policy(
                &mut references,
                ArchiveReferencePolicy::KnownReferences,
                limits,
            )
            .map_err(map_core_error)?;
        let mut payload_references = Vec::new();
        let wire_limits = budget.residual(package).map_err(map_transaction_error)?;
        for message in &object.messages {
            payload_references.extend(payload_object_references(message, wire_limits, budget)?);
        }
        budget
            .work(payload_references.len())
            .and_then(|_| budget.references(payload_references.len()))
            .map_err(map_transaction_error)?;
        budget
            .references(references.identifiers.len())
            .map_err(map_transaction_error)?;
        pending.extend(references.identifiers);
        pending.extend(payload_references);
    }
    Ok(reachable)
}

fn verify_locality_members(
    source: &Package,
    candidate: &Package,
    plan: &MutationPlan,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let locality_work = source_catalog
        .source_bytes()
        .len()
        .checked_add(candidate_catalog.source_bytes().len())
        .ok_or(Error::InvalidSource)?;
    budget.work(locality_work).map_err(map_transaction_error)?;
    let set_capacity = plan
        .members
        .len()
        .checked_add(plan.deleted_members.len())
        .and_then(|value| value.checked_add(plan.object_ids.len()))
        .ok_or(Error::InvalidSource)?;
    let set_scratch = plan
        .members
        .len()
        .checked_mul(size_of::<&str>())
        .and_then(|value| {
            value.checked_add(plan.deleted_members.len().checked_mul(size_of::<&str>())?)
        })
        .and_then(|value| value.checked_add(plan.object_ids.len().checked_mul(size_of::<u64>())?))
        .ok_or(Error::InvalidSource)?;
    budget
        .work(set_capacity)
        .and_then(|_| budget.allocations(3))
        .and_then(|_| budget.scratch(set_scratch))
        .map_err(map_transaction_error)?;
    let mut changed_members = HashSet::new();
    changed_members
        .try_reserve(plan.members.len())
        .map_err(|_| Error::Allocation {
            amount: plan.members.len(),
        })?;
    changed_members.extend(plan.members.iter().map(String::as_str));
    let mut deleted_members = HashSet::new();
    deleted_members
        .try_reserve(plan.deleted_members.len())
        .map_err(|_| Error::Allocation {
            amount: plan.deleted_members.len(),
        })?;
    deleted_members.extend(plan.deleted_members.iter().map(String::as_str));
    let mut allowed_objects = HashSet::new();
    allowed_objects
        .try_reserve(plan.object_ids.len())
        .map_err(|_| Error::Allocation {
            amount: plan.object_ids.len(),
        })?;
    allowed_objects.extend(plan.object_ids.iter().copied());
    let source_reachable = reachable_objects(source, plan.root_identifier, budget)?;
    let candidate_reachable = reachable_objects(candidate, plan.root_identifier, budget)?;
    for identifier in &allowed_objects {
        if *identifier != 2
            && package_object(source, *identifier).is_some()
            && !source_reachable.contains(identifier)
        {
            return Err(Error::InvalidSource);
        }
        if *identifier != 2
            && package_object(candidate, *identifier).is_some()
            && !candidate_reachable.contains(identifier)
        {
            return Err(Error::InvalidSource);
        }
    }

    for candidate_entry in candidate_catalog.package().iter() {
        if source_catalog
            .package()
            .iter()
            .all(|entry| entry.name() != candidate_entry.name())
            && !changed_members.contains(candidate_entry.name())
        {
            return Err(Error::InvalidSource);
        }
    }
    for name in &deleted_members {
        if candidate_catalog
            .package()
            .iter()
            .any(|entry| entry.name() == *name)
        {
            return Err(Error::InvalidSource);
        }
    }

    for source_entry in source_catalog.package().iter() {
        let name = source_entry.name();
        if deleted_members.contains(name) {
            continue;
        }
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or(Error::InvalidSource)?;
        if !changed_members.contains(name) {
            if source_entry.data() != candidate_entry.data()
                || source_entry.raw_name() != candidate_entry.raw_name()
                || source_entry.raw_record().local_record()
                    != candidate_entry.raw_record().local_record()
                || !central_record_preserved(source_entry, candidate_entry)
            {
                return Err(Error::InvalidSource);
            }
            continue;
        }
        if !zip_static_metadata_preserved(source_entry, candidate_entry) {
            return Err(Error::InvalidSource);
        }
        let source_component = source
            .state
            .source
            .components()
            .get(name)
            .ok_or(Error::InvalidSource)?;
        let candidate_component = candidate
            .state
            .source
            .components()
            .get(name)
            .ok_or(Error::InvalidSource)?;
        for source_object in &source_component.archive().objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource)?;
            let Some(candidate_object) = candidate_component
                .archive()
                .objects
                .iter()
                .find(|object| object.archive_info.identifier == Some(identifier))
            else {
                if allowed_objects.contains(&identifier) {
                    continue;
                }
                return Err(Error::InvalidSource);
            };
            if source_object.same_content_ignoring_offsets(candidate_object) {
                continue;
            }
            if !allowed_objects.contains(&identifier)
                || !source_object
                    .messages
                    .first()
                    .is_some_and(|message| is_chart_font_rewritable_type(message.type_))
            {
                return Err(Error::InvalidSource);
            }
            if source_object.messages.len() != candidate_object.messages.len() {
                return Err(Error::InvalidSource);
            }
            inspect_archive_metadata_strict(source, source_object, budget)?;
            inspect_archive_metadata_strict(candidate, candidate_object, budget)?;
            if !archive_metadata_preserved(
                source_object,
                candidate_object,
                &allowed_objects,
                &source_reachable,
                &candidate_reachable,
            ) {
                return Err(Error::InvalidSource);
            }
            for (index, source_message) in source_object.messages.iter().enumerate() {
                let candidate_message = candidate_object
                    .messages
                    .get(index)
                    .ok_or(Error::InvalidSource)?;
                if candidate_message.type_ != source_message.type_
                    || (!is_chart_font_rewritable_type(source_message.type_)
                        && candidate_message != source_message)
                    || (is_chart_font_rewritable_type(source_message.type_)
                        && !message_unknown_fields_preserved(
                            source_message.type_,
                            source_message.data.as_slice(),
                            candidate_message.data.as_slice(),
                            budget.residual(source).map_err(map_transaction_error)?,
                            budget,
                        )?)
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
        for candidate_object in &candidate_component.archive().objects {
            if source_component.archive().objects.iter().all(|object| {
                object.archive_info.identifier != candidate_object.archive_info.identifier
            }) && !candidate_object
                .archive_info
                .identifier
                .is_some_and(|id| allowed_objects.contains(&id))
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn zip_static_metadata_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_local = source.metadata().local();
    let candidate_local = candidate.metadata().local();
    let source_central = source.metadata().central();
    let candidate_central = candidate.metadata().central();
    source.raw_name() == candidate.raw_name()
        && central_static_metadata_preserved(source, candidate)
        && source.is_opaque() == candidate.is_opaque()
        && source_local.version_needed() == candidate_local.version_needed()
        && source_local.flags() == candidate_local.flags()
        && source_local.compression_method() == candidate_local.compression_method()
        && source_local.last_modified() == candidate_local.last_modified()
        && source_local.name() == candidate_local.name()
        && source_local.extra() == candidate_local.extra()
        && source_local.comment() == candidate_local.comment()
        && source_central.version_needed() == candidate_central.version_needed()
        && source_central.flags() == candidate_central.flags()
        && source_central.compression_method() == candidate_central.compression_method()
        && source_central.last_modified() == candidate_central.last_modified()
        && source_central.name() == candidate_central.name()
        && source_central.extra() == candidate_central.extra()
        && source_central.comment() == candidate_central.comment()
}

fn central_static_metadata_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    source_record.len() >= 46
        && candidate_record.len() >= 46
        && source_record[..16] == candidate_record[..16]
        && source_record[34..42] == candidate_record[34..42]
        && source_record[46..] == candidate_record[46..]
}

fn central_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    source_record.len() >= 46
        && candidate_record.len() >= 46
        && source_record[..42] == candidate_record[..42]
        && source_record[46..] == candidate_record[46..]
}

fn references_transition_preserved(
    source: &[u64],
    candidate: &[u64],
    allowed: &HashSet<u64>,
    source_reachable: &HashSet<u64>,
    candidate_reachable: &HashSet<u64>,
) -> bool {
    source.iter().all(|identifier| {
        candidate.contains(identifier)
            || allowed.contains(identifier)
            || source_reachable.contains(identifier)
            || !candidate_reachable.contains(identifier)
    }) && candidate.iter().all(|identifier| {
        source.contains(identifier)
            || allowed.contains(identifier)
            || source_reachable.contains(identifier)
    })
}

fn archive_metadata_preserved(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    allowed: &HashSet<u64>,
    source_reachable: &HashSet<u64>,
    candidate_reachable: &HashSet<u64>,
) -> bool {
    if source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return false;
    }
    for (source_info, candidate_info) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
    {
        if source_info.type_ != candidate_info.type_
            || source_info.versions != candidate_info.versions
            || source_info.field_infos.len() != candidate_info.field_infos.len()
            || source_info.data_references != candidate_info.data_references
            || source_info.base_message_index != candidate_info.base_message_index
            || source_info.diff_merge_version != candidate_info.diff_merge_version
            || source_info.diff_field_path != candidate_info.diff_field_path
            || source_info.fields_to_remove != candidate_info.fields_to_remove
            || source_info.diff_read_version != candidate_info.diff_read_version
            || !references_transition_preserved(
                &source_info.object_references,
                &candidate_info.object_references,
                allowed,
                source_reachable,
                candidate_reachable,
            )
        {
            return false;
        }
        for (source_field, candidate_field) in source_info
            .field_infos
            .iter()
            .zip(&candidate_info.field_infos)
        {
            if source_field.path != candidate_field.path
                || source_field.r#type != candidate_field.r#type
                || source_field.unknown_field_rule != candidate_field.unknown_field_rule
                || source_field.data_references != candidate_field.data_references
                || source_field.known_field_rule != candidate_field.known_field_rule
                || source_field.known_field_version != candidate_field.known_field_version
                || source_field.known_field_feature_identifier
                    != candidate_field.known_field_feature_identifier
                || !references_transition_preserved(
                    &source_field.object_references,
                    &candidate_field.object_references,
                    allowed,
                    source_reachable,
                    candidate_reachable,
                )
            {
                return false;
            }
        }
    }
    true
}

fn message_unknown_fields_preserved(
    message_type: u32,
    source: &[u8],
    candidate: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    match message_type {
        CHART_MESSAGE_TYPE => chart_unknown_fields_preserved(source, candidate, limits, budget),
        PARAGRAPH_STYLE_MESSAGE_TYPE => {
            wire_unknown_fields_preserved(source, candidate, limits, &[1, 10, 11, 12], budget)
        },
        STYLESHEET_MESSAGE_TYPE => {
            wire_unknown_fields_preserved(source, candidate, limits, &[1, 5], budget)
        },
        PACKAGE_METADATA_MESSAGE_TYPE => {
            wire_unknown_fields_preserved(source, candidate, limits, &[1, 3, 8, 11], budget)
        },
        _ => Ok(true),
    }
}

fn wire_unknown_fields_preserved(
    source: &[u8],
    candidate: &[u8],
    limits: litchi_iwa_common::WireLimits,
    known_fields: &[u32],
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    let source_fields = parse_locality_fields(source, limits, budget)?;
    let candidate_fields = parse_locality_fields(candidate, limits, budget)?;
    Ok(
        raw_fields_except_numbers(&source_fields, source, known_fields, budget)?
            == raw_fields_except_numbers(&candidate_fields, candidate, known_fields, budget)?,
    )
}

fn chart_unknown_fields_preserved(
    source: &[u8],
    candidate: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    let source_fields = parse_locality_fields(source, limits, budget)?;
    let candidate_fields = parse_locality_fields(candidate, limits, budget)?;
    let source_unknown = raw_fields_except(&source_fields, source, 10_000, budget)?;
    let candidate_unknown = raw_fields_except(&candidate_fields, candidate, 10_000, budget)?;
    if source_unknown != candidate_unknown {
        return Ok(false);
    }
    let mut source_extensions = source_fields
        .iter()
        .filter(|field| field.number() == 10_000);
    let mut candidate_extensions = candidate_fields
        .iter()
        .filter(|field| field.number() == 10_000);
    loop {
        let source_extension = source_extensions.next();
        let candidate_extension = candidate_extensions.next();
        match (source_extension, candidate_extension) {
            (None, None) => break,
            (Some(_), None) | (None, Some(_)) => return Ok(false),
            (Some(source_extension), Some(candidate_extension)) => {
                let source_extension = source_extension.payload(source).map_err(map_wire_error)?;
                let candidate_extension = candidate_extension
                    .payload(candidate)
                    .map_err(map_wire_error)?;
                let source_fields = parse_locality_fields(source_extension, limits, budget)?;
                let candidate_fields = parse_locality_fields(candidate_extension, limits, budget)?;
                let source_unknown =
                    raw_fields_except(&source_fields, source_extension, 20, budget)?;
                let candidate_unknown =
                    raw_fields_except(&candidate_fields, candidate_extension, 20, budget)?;
                if source_unknown != candidate_unknown {
                    return Ok(false);
                }
            },
        }
    }
    Ok(true)
}

fn mutation_plan(
    package: &Package,
    root_identifier: u64,
    changes: &graph::Changes,
    budget: &mut TransactionBudget,
) -> Result<MutationPlan, Error> {
    let member_capacity = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|value| value.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let mut object_capacity = 0usize;
    let mut name_bytes = 0usize;
    for (name, archive) in &changes.edits {
        name_bytes = name_bytes
            .checked_add(name.len())
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::InvalidSource)?;
        object_capacity = object_capacity
            .checked_add(archive.objects.len())
            .ok_or(Error::InvalidSource)?;
    }
    for insertion in &changes.insertions {
        name_bytes = name_bytes
            .checked_add(insertion.name.len())
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::InvalidSource)?;
        object_capacity = object_capacity
            .checked_add(insertion.archive.objects.len())
            .ok_or(Error::InvalidSource)?;
    }
    for name in &changes.deletions {
        name_bytes = name_bytes
            .checked_add(name.len())
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::InvalidSource)?;
    }
    let scratch = member_capacity
        .checked_mul(size_of::<String>())
        .and_then(|value| value.checked_add(object_capacity.checked_mul(size_of::<u64>())?))
        .ok_or(Error::InvalidSource)?;
    budget
        .work(
            name_bytes
                .checked_add(object_capacity)
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|_| budget.allocations(member_capacity.max(1).saturating_add(3)))
        .and_then(|_| budget.retained(name_bytes))
        .and_then(|_| budget.scratch(name_bytes))
        .and_then(|_| budget.scratch(scratch))
        .map_err(map_transaction_error)?;
    let mut members = HashSet::new();
    let mut deleted_members = HashSet::new();
    let mut object_ids = HashSet::new();
    members
        .try_reserve(member_capacity)
        .map_err(|_| Error::Allocation {
            amount: member_capacity,
        })?;
    deleted_members
        .try_reserve(changes.deletions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.deletions.len(),
        })?;
    object_ids
        .try_reserve(object_capacity)
        .map_err(|_| Error::Allocation {
            amount: object_capacity,
        })?;
    for (name, archive) in &changes.edits {
        members.insert(name.clone());
        let source_archive = package
            .state
            .source
            .components()
            .get(name)
            .ok_or(Error::InvalidSource)?
            .archive();
        for source_object in &source_archive.objects {
            let id = source_object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource)?;
            if archive
                .objects
                .iter()
                .all(|object| object.archive_info.identifier != Some(id))
            {
                object_ids.insert(id);
            }
        }
        for candidate_object in &archive.objects {
            let id = candidate_object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource)?;
            let source_object = source_archive
                .objects
                .iter()
                .find(|object| object.archive_info.identifier == Some(id));
            if source_object
                .is_none_or(|object| !object.same_content_ignoring_offsets(candidate_object))
            {
                object_ids.insert(id);
            }
        }
    }
    for insertion in &changes.insertions {
        members.insert(insertion.name.clone());
        for object in &insertion.archive.objects {
            object_ids.insert(object.archive_info.identifier.ok_or(Error::InvalidSource)?);
        }
    }
    for name in &changes.deletions {
        members.insert(name.clone());
        deleted_members.insert(name.clone());
    }
    let mut members_vec = Vec::new();
    members_vec
        .try_reserve_exact(members.len())
        .map_err(|_| Error::Allocation {
            amount: members.len(),
        })?;
    members_vec.extend(members);
    let mut deleted_members_vec = Vec::new();
    deleted_members_vec
        .try_reserve_exact(deleted_members.len())
        .map_err(|_| Error::Allocation {
            amount: deleted_members.len(),
        })?;
    deleted_members_vec.extend(deleted_members);
    let mut object_ids_vec = Vec::new();
    object_ids_vec
        .try_reserve_exact(object_ids.len())
        .map_err(|_| Error::Allocation {
            amount: object_ids.len(),
        })?;
    object_ids_vec.extend(object_ids);
    let mut members = members_vec;
    let mut deleted_members = deleted_members_vec;
    let mut object_ids = object_ids_vec;
    members.sort_unstable();
    deleted_members.sort_unstable();
    object_ids.sort_unstable();
    Ok(MutationPlan {
        root_identifier,
        members: members.into_boxed_slice(),
        deleted_members: deleted_members.into_boxed_slice(),
        object_ids: object_ids.into_boxed_slice(),
    })
}

const fn is_chart_font_rewritable_type(message_type: u32) -> bool {
    matches!(
        message_type,
        CHART_MESSAGE_TYPE
            | PARAGRAPH_STYLE_MESSAGE_TYPE
            | STYLESHEET_MESSAGE_TYPE
            | PACKAGE_METADATA_MESSAGE_TYPE
    )
}

fn publish(
    package: &Package,
    changes: graph::Changes,
    budget: &mut TransactionBudget,
    plan: &MutationPlan,
) -> Result<Package, Error> {
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
        admit_archive_encoding(archive, archive_limits, snappy_limits, budget)?;
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
        admit_archive_encoding(&insertion.archive, archive_limits, snappy_limits, budget)?;
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
    let edited_bytes = edited
        .iter()
        .map(|(_, data)| data.len())
        .chain(inserted.iter().map(|(_, data, _)| data.len()))
        .try_fold(0usize, |total, bytes| total.checked_add(bytes))
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_reassembly(catalog, edited_bytes, changes.deletions.len())
        .map_err(map_transaction_error)?;
    let prepared = catalog
        .package()
        .prepare_reassembly_with_changes(&insertions, &edits, &deletions, catalog.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .reassembly(requirements)
        .map_err(map_transaction_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let output_len = output.len();
    preflight_candidate(package, output_len, changes.insertions.len(), budget)?;
    let candidate = Package::from_source_with_options(output.into(), package.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    verify_locality_members(package, &candidate, plan, budget)?;
    Ok(candidate)
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

fn map_axis_error(error: AxisSupportError) -> Error {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            AxisSupportSelectorError::UnsupportedSource => Refusal::UnsupportedSource.into(),
            AxisSupportSelectorError::AmbiguousSelector => Refusal::AmbiguousSelector.into(),
            AxisSupportSelectorError::EmptySlideName => Refusal::EmptySlideName.into(),
            AxisSupportSelectorError::SlideNameNotFound => Refusal::SlideNameNotFound.into(),
            AxisSupportSelectorError::SlidePositionNotFound { position } => {
                Refusal::SlidePositionNotFound { position }.into()
            },
            AxisSupportSelectorError::ChartNameNotFound => Refusal::ChartNameNotFound.into(),
            AxisSupportSelectorError::ChartPositionNotFound { position } => {
                Refusal::ChartPositionNotFound { position }.into()
            },
            AxisSupportSelectorError::EmptyChartName => Refusal::EmptyChartName.into(),
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => LimitKind::InputBytes,
                chart_axis_support::AxisSupportLimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_axis_support::AxisSupportLimitKind::WireBytes => LimitKind::WireBytes,
                chart_axis_support::AxisSupportLimitKind::Entries => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::EntryBytes => LimitKind::EntryBytes,
                chart_axis_support::AxisSupportLimitKind::TotalBytes => LimitKind::TotalBytes,
                chart_axis_support::AxisSupportLimitKind::Slides => LimitKind::Slides,
                chart_axis_support::AxisSupportLimitKind::References => LimitKind::References,
                chart_axis_support::AxisSupportLimitKind::TextStorages
                | chart_axis_support::AxisSupportLimitKind::TextFragments => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::TextBytes => LimitKind::TextBytes,
                chart_axis_support::AxisSupportLimitKind::WireFields => LimitKind::WireFields,
                chart_axis_support::AxisSupportLimitKind::WireNesting => LimitKind::WireNesting,
                chart_axis_support::AxisSupportLimitKind::WireWork => LimitKind::WireWork,
                chart_axis_support::AxisSupportLimitKind::TitleBytes => LimitKind::SelectorBytes,
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => Error::Allocation { amount },
        AxisSupportError::InvalidSource => Error::InvalidSource,
    }
}

fn map_transaction_error(error: TransactionError) -> Error {
    match error {
        TransactionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                super::slide_table_core::LimitKind::InputBytes => LimitKind::InputBytes,
                super::slide_table_core::LimitKind::OutputBytes => LimitKind::OutputBytes,
                super::slide_table_core::LimitKind::WireFields => LimitKind::WireFields,
                super::slide_table_core::LimitKind::WireNesting => LimitKind::WireNesting,
                super::slide_table_core::LimitKind::WireWork => LimitKind::WireWork,
                super::slide_table_core::LimitKind::References => LimitKind::References,
                super::slide_table_core::LimitKind::Allocations => LimitKind::Allocations,
                super::slide_table_core::LimitKind::Retained => LimitKind::RetainedBytes,
                super::slide_table_core::LimitKind::Scratch => LimitKind::ScratchBytes,
                super::slide_table_core::LimitKind::Entries
                | super::slide_table_core::LimitKind::EntryBytes
                | super::slide_table_core::LimitKind::TotalBytes
                | super::slide_table_core::LimitKind::Components
                | super::slide_table_core::LimitKind::PayloadObjects
                | super::slide_table_core::LimitKind::PayloadMessages => LimitKind::Entries,
            },
            observed,
            maximum,
        },
        TransactionError::Allocation(amount) => Error::Allocation { amount },
        TransactionError::InvalidSource
        | TransactionError::UnsupportedSource
        | TransactionError::UnsupportedDependency
        | TransactionError::UnsupportedTopology
        | TransactionError::AmbiguousSelector
        | TransactionError::EmptySlideName
        | TransactionError::SlideNameNotFound
        | TransactionError::SlidePositionNotFound(_)
        | TransactionError::TablePositionNotFound(_)
        | TransactionError::Read
        | TransactionError::Wire
        | TransactionError::Codec
        | TransactionError::Archive
        | TransactionError::Verification => Error::InvalidSource,
    }
}

fn map_engine_error(error: engine::Error) -> Error {
    match error {
        engine::Error::InvalidSource(_) => Error::InvalidSource,
        engine::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                engine::LimitKind::InputBytes => LimitKind::InputBytes,
                engine::LimitKind::Fields => LimitKind::WireFields,
                engine::LimitKind::OutputBytes => LimitKind::OutputBytes,
                engine::LimitKind::Nesting => LimitKind::WireNesting,
                engine::LimitKind::Work => LimitKind::WireWork,
                engine::LimitKind::TextBytes => LimitKind::TextBytes,
                engine::LimitKind::Allocations => LimitKind::Allocations,
                engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                engine::LimitKind::Archive(kind) => map_core_limit(kind),
                engine::LimitKind::Wire(kind) => map_wire_limit(kind),
                engine::LimitKind::Codec(kind) => map_codec_limit(kind),
                _ => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        engine::Error::Refused(_) => Refusal::UnsupportedStyle.into(),
        engine::Error::Graph(error) => map_graph_error(*error),
        engine::Error::Codec(error) => map_codec_error(error),
        engine::Error::Wire(error) => map_wire_error(error),
        engine::Error::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
        _ => Error::InvalidSource,
    }
}

fn map_graph_error(error: graph::Error) -> Error {
    match error {
        graph::Error::Archive(error) => map_core_error(error),
        graph::Error::Wire(error) => map_wire_error(error),
        graph::Error::Codec(error) => map_codec_error(error),
        graph::Error::Registry(error) => map_registry_error(error),
        graph::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::GraphLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                graph::GraphLimitKind::AllocationEvents => LimitKind::Allocations,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
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

fn map_core_limit(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
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
    }
}

fn map_wire_limit(kind: litchi_iwa_common::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
        litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
        litchi_iwa_common::LimitKind::Fields
        | litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
        litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
        litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
    }
}

fn map_codec_limit(kind: litchi_iwa_protos::chart_style_codec::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_protos::chart_style_codec::LimitKind::Bytes => LimitKind::WireBytes,
        litchi_iwa_protos::chart_style_codec::LimitKind::Fields => LimitKind::WireFields,
        _ => LimitKind::WireBytes,
    }
}

fn map_codec_error(error: litchi_iwa_protos::chart_style_codec::Error) -> Error {
    match error {
        litchi_iwa_protos::chart_style_codec::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_codec_limit(kind),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_protos::chart_style_codec::Error::Allocation { bytes } => {
            Error::Allocation { amount: bytes }
        },
        _ => Error::InvalidSource,
    }
}

fn map_registry_error(error: litchi_iwa_protos::package_metadata_codec::RewriteError) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                observed,
                maximum,
            } => (LimitKind::InputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                observed,
                maximum,
            } => (LimitKind::OutputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                observed,
                maximum,
            } => (LimitKind::WireFields, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                observed,
                maximum,
            } => (LimitKind::WireNesting, observed as usize, maximum as usize),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                observed,
                maximum,
            } => (LimitKind::Entries, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                observed,
                maximum,
            } => (LimitKind::References, observed, maximum),
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    Error::InvalidSource
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

#[cfg(test)]
mod locality_tests {
    use super::*;

    const SOURCE: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/iwork/keynote/chart-data-edited-native.key"
    ));

    fn package_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }

    fn chart_root(package: &Package) -> Result<u64, Box<dyn std::error::Error>> {
        package
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .find(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == CHART_MESSAGE_TYPE)
            })
            .and_then(|object| object.archive_info.identifier)
            .ok_or_else(|| "chart root is absent".into())
    }

    fn locality_plan(
        package: &Package,
        member: &str,
        root_identifier: u64,
    ) -> Result<MutationPlan, Box<dyn std::error::Error>> {
        let component = package
            .state
            .source
            .components()
            .get(member)
            .ok_or("requested member is absent")?;
        let mut object_ids = Vec::new();
        object_ids.try_reserve_exact(component.archive().objects.len())?;
        for object in &component.archive().objects {
            object_ids.push(object.archive_info.identifier.ok_or("object has no id")?);
        }
        Ok(MutationPlan {
            root_identifier,
            members: vec![member.to_owned()].into_boxed_slice(),
            deleted_members: Box::default(),
            object_ids: object_ids.into_boxed_slice(),
        })
    }

    fn tamper_archive_member(
        package: &Package,
        message_type: u32,
        mutate: impl FnOnce(&mut Archive) -> Result<(), Box<dyn std::error::Error>>,
    ) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let catalog = physical_catalog(package)?;
        let archive_limits = catalog.limits().effective_archive_limits()?;
        let snappy_limits = catalog.limits().snappy_limits()?;
        for entry in catalog.package().iter() {
            if !entry.name().ends_with(".iwa") {
                continue;
            }
            let decompressed = SnappyStream::decompress(entry.data(), snappy_limits)?;
            let mut archive = Archive::parse(decompressed.as_bytes(), archive_limits)?;
            if !archive.objects.iter().any(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == message_type)
            }) {
                continue;
            }
            mutate(&mut archive)?;
            let rewritten = archive.to_bytes(archive_limits)?;
            let compressed = SnappyStream::compress(&rewritten, snappy_limits)?;
            return Ok(catalog.package().reassemble_to_bytes(
                &[EntryEdit::new(entry.name(), &compressed)],
                catalog.limits(),
            )?);
        }
        Err("requested message type is absent".into())
    }

    fn member_with_message(
        package: &Package,
        message_type: u32,
    ) -> Result<String, Box<dyn std::error::Error>> {
        let catalog = physical_catalog(package)?;
        let archive_limits = catalog.limits().effective_archive_limits()?;
        let snappy_limits = catalog.limits().snappy_limits()?;
        for entry in catalog.package().iter() {
            if !entry.name().ends_with(".iwa") {
                continue;
            }
            let decompressed = SnappyStream::decompress(entry.data(), snappy_limits)?;
            let archive = Archive::parse(decompressed.as_bytes(), archive_limits)?;
            if archive.objects.iter().any(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == message_type)
            }) {
                return Ok(entry.name().to_owned());
            }
        }
        Err("requested message type is absent".into())
    }

    fn append_unknown_message_field(
        archive: &mut Archive,
        message_type: u32,
    ) -> Result<(), Box<dyn std::error::Error>> {
        for object in &mut archive.objects {
            for (index, message) in object.messages.iter_mut().enumerate() {
                if message.type_ != message_type {
                    continue;
                }
                message.data.extend_from_slice(&[0x98, 0x06, 0x01]);
                object.archive_info.message_infos[index].length =
                    u32::try_from(message.data.len())?;
                return Ok(());
            }
        }
        Err("requested message type is absent".into())
    }

    fn append_unplanned_object(archive: &mut Archive) -> Result<(), Box<dyn std::error::Error>> {
        let identifier = archive
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .max()
            .ok_or("archive has no objects")?
            .checked_add(1)
            .ok_or("object identifier overflow")?;
        archive.objects.push(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: 9_999,
                data: vec![1],
            }],
            CoreLimits::default(),
        )?);
        Ok(())
    }

    fn mutate_central_version(
        mut bytes: Vec<u8>,
        member: &str,
    ) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let signature = b"PK\x01\x02";
        let member = member.as_bytes();
        let mut offset = 0;
        while let Some(relative) = bytes[offset..]
            .windows(signature.len())
            .position(|window| window == signature)
        {
            let start = offset + relative;
            if start + 46 > bytes.len() {
                break;
            }
            let name_length = usize::from(u16::from_le_bytes(
                bytes[start + 28..start + 30].try_into()?,
            ));
            let extra_length = usize::from(u16::from_le_bytes(
                bytes[start + 30..start + 32].try_into()?,
            ));
            let comment_length = usize::from(u16::from_le_bytes(
                bytes[start + 32..start + 34].try_into()?,
            ));
            let name_start = start + 46;
            let name_end = name_start + name_length;
            if name_end > bytes.len() {
                break;
            }
            if &bytes[name_start..name_end] == member {
                bytes[start + 4] ^= 1;
                return Ok(bytes);
            }
            offset = name_end
                .checked_add(extra_length)
                .and_then(|value| value.checked_add(comment_length))
                .ok_or("central directory overflow")?;
        }
        Err("requested ZIP member is absent".into())
    }

    fn assert_locality_refused(
        source: &Package,
        plan: &MutationPlan,
        bytes: &[u8],
    ) -> Result<(), Box<dyn std::error::Error>> {
        let candidate = Package::from_bytes(bytes)?;
        let mut budget = transaction_budget(source)?;
        admit_transaction(source, &mut budget)?;
        let result = verify_locality_members(source, &candidate, plan, &mut budget);
        assert!(
            result.is_err(),
            "locality verifier unexpectedly accepted: {result:?}"
        );
        Ok(())
    }

    #[test]
    fn locality_verifier_rejects_unplanned_member_object_unknown_and_zip_metadata()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = Package::from_bytes(SOURCE)?;
        let chart_member = member_with_message(&source, CHART_MESSAGE_TYPE)?;
        let chart_plan = locality_plan(&source, &chart_member, chart_root(&source)?)?;
        let candidate_bytes = package_bytes(&source)?;
        let catalog = SourceCatalog::from_bytes(&candidate_bytes, Default::default())?;
        let extra_member = catalog
            .package()
            .reassemble_with_insertions_to_bytes(
                &[EntryInsertion::new(
                    "Data/unplanned-chart-font.bin",
                    b"unplanned",
                )],
                catalog.limits(),
            )
            .map_err(|error| format!("insert member: {error}"))?;
        assert_locality_refused(&source, &chart_plan, &extra_member)
            .map_err(|error| format!("extra member: {error}"))?;

        let extra_object =
            tamper_archive_member(&source, CHART_MESSAGE_TYPE, append_unplanned_object)
                .map_err(|error| format!("tamper object: {error}"))?;
        assert_locality_refused(&source, &chart_plan, &extra_object)
            .map_err(|error| format!("extra object: {error}"))?;

        let stylesheet_member = member_with_message(&source, STYLESHEET_MESSAGE_TYPE)?;
        let stylesheet_plan = locality_plan(&source, &stylesheet_member, chart_root(&source)?)?;
        let unknown_stylesheet =
            tamper_archive_member(&source, STYLESHEET_MESSAGE_TYPE, |archive| {
                append_unknown_message_field(archive, STYLESHEET_MESSAGE_TYPE)
            })
            .map_err(|error| format!("tamper stylesheet: {error}"))?;
        assert_locality_refused(&source, &stylesheet_plan, &unknown_stylesheet)
            .map_err(|error| format!("unknown stylesheet: {error}"))?;

        for message_type in [PARAGRAPH_STYLE_MESSAGE_TYPE, PACKAGE_METADATA_MESSAGE_TYPE] {
            let message_member = member_with_message(&source, message_type)?;
            let message_plan = locality_plan(&source, &message_member, chart_root(&source)?)?;
            let unknown_message = tamper_archive_member(&source, message_type, |archive| {
                append_unknown_message_field(archive, message_type)
            })?;
            assert_locality_refused(&source, &message_plan, &unknown_message)?;
        }

        let metadata_tamper = mutate_central_version(candidate_bytes, &chart_member)?;
        assert_locality_refused(&source, &chart_plan, &metadata_tamper)
            .map_err(|error| format!("ZIP metadata: {error}"))?;
        Ok(())
    }

    #[test]
    fn selector_field_admission_precedes_scan_work() -> Result<(), Box<dyn std::error::Error>> {
        let package = Package::from_bytes(SOURCE)?;
        let mut transaction = transaction_budget(&package)?;
        let mut fields = FieldLedger {
            used: 0,
            maximum: 1,
        };
        let mut selector = ChartFontSelectionBudget {
            budget: &mut transaction,
            fields: &mut fields,
        };
        assert!(matches!(
            selector.finish_wire_scan(2),
            Err(AxisSupportError::LimitExceeded {
                kind: chart_axis_support::AxisSupportLimitKind::WireFields,
                ..
            })
        ));
        assert_eq!(fields.used, 0);
        Ok(())
    }
}

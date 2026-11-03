//! Selector-first title and caption transactions of sheet shapes, images,
//! movies and charts (record 0425) on the shared
//! `litchi_iwa_style::caption` engine.
//!
//! A drawable's title and caption are references in its
//! `TSD.DrawableArchive` (fields 10 and 11): a present label is a caption
//! info with a private text storage, shape style and placement in the
//! drawable's component, an absent one an empty stand-in. Setting a present
//! label rewrites its storage text; setting an absent one authors the
//! native inline graph; clearing a present label points the drawable at a
//! fresh stand-in and keeps the old graph, as the application does for
//! undo. New objects get UUIDs in the drawable's component and raise the
//! identifier watermark (the host's `set_sheet_{shape,image,movie}_{title,
//! caption}`, `remove_*` and `set_sheet_chart_caption`). A chart has a
//! caption only; its title is the chart-title family's.

use std::fmt;

use litchi_iwa_core::Limits;
use litchi_iwa_style::caption;
use litchi_iwa_style::graph::Workspace;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph as graph;
use super::structure::{self, references};
use super::style_graph as numbers_style_graph;
use crate::selector::SheetSelector;
use crate::sheet::drawable::{DrawableSelector, Kind};

structure::structural_family!("sheet drawable caption");

/// `TN.DocumentArchive.super` (`TSA.DocumentArchive`).
const DOCUMENT_SUPER_FIELD: u32 = 8;
/// `TSA.DocumentArchive.document_language`.
const DOCUMENT_LANGUAGE_FIELD: u32 = 3;
/// `TN.DocumentArchive.stylesheet`.
const DOCUMENT_STYLESHEET_FIELD: u32 = 4;
/// `TN.DocumentArchive.theme`.
const DOCUMENT_THEME_FIELD: u32 = 6;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `TN.ThemeArchive`.
const THEME_MESSAGE_TYPE: u32 = 12_009;
/// `TN.ThemeArchive.super` (`TSS.ThemeArchive`).
const THEME_SUPER_FIELD: u32 = 1;

/// Which of a drawable's two labels an operation targets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Label {
    /// The label above the drawable.
    Title,
    /// The label below the drawable.
    Caption,
}

impl Label {
    const fn shared(self) -> caption::Label {
        match self {
            Self::Title => caption::Label::Title,
            Self::Caption => caption::Label::Caption,
        }
    }
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The selected drawable has no such label in this family (text boxes,
    /// audio, tables and groups; a chart's title).
    #[error("a Numbers sheet {kind} has no {label:?} in this family")]
    NoLabel {
        /// The selected drawable kind.
        kind: &'static str,
        /// The requested label.
        label: Label,
    },
    /// The text contains an object-replacement, footnote or section-break
    /// marker, which only object families may write.
    #[error("the label text contains a reserved native text marker")]
    ReservedMarker,
    /// Replacing the present label's text would drop an object it
    /// references.
    #[error("the label text references a native object")]
    DependentContent,
    /// Authoring a new label needs a positive displayed width.
    #[error("the drawable has no displayed width for a new label")]
    NoWidth,
    /// A new chart caption in a chart stored beside the calculation engine
    /// needs the host's componentized graph (its style in the document
    /// stylesheet), which this family does not author; replacing or
    /// removing a present caption is supported.
    #[error("a new caption of this chart needs the componentized caption layout")]
    ComponentizedLayout,
    /// A second operation was staged on one edit.
    #[error("one label edit stages one operation")]
    OperationAlreadyStaged,
    /// Commit was requested without a staged change.
    #[error("no sheet-drawable label change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the graph being rewritten.
    #[error("a surviving object depends on the selected label")]
    UnsupportedDependency,
    /// The source has no exact physical artifact suitable for publication.
    #[error("source representation does not support sheet-drawable label edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::NoLabel { .. } => RefusalKind::UnsupportedTarget,
            Self::ReservedMarker => RefusalKind::InvalidValue,
            Self::DependentContent
            | Self::ComponentizedLayout
            | Self::NoWidth
            | Self::UnsupportedDependency
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::OperationAlreadyStaged | Self::NoStagedOperation => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

/// A typed, content-redacted sheet-drawable label failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or its label graph is malformed.
    #[error("the selected Numbers sheet-drawable label source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} sheet-drawable label units or bytes")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested labels.
    #[error("sheet-drawable label candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("sheet-drawable label patch conflicts with this package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_style::graph::Error> for Error {
    fn from(error: litchi_iwa_style::graph::Error) -> Self {
        match &error {
            litchi_iwa_style::graph::Error::Invalid(_) => Self::InvalidSource,
            _ => match numbers_style_graph::fault(&error) {
                numbers_style_graph::StyleGraphFault::Shared(fault) => fault.into(),
                numbers_style_graph::StyleGraphFault::Limit {
                    kind,
                    observed,
                    maximum,
                } => Self::Limit {
                    kind: match kind {
                        litchi_iwa_style::graph::GraphLimitKind::ScratchBytes => {
                            LimitKind::ScratchBytes
                        },
                        litchi_iwa_style::graph::GraphLimitKind::AllocationEvents => {
                            LimitKind::AllocationEvents
                        },
                    },
                    observed,
                    maximum,
                },
            },
        }
    }
}

impl From<caption::Error> for Error {
    fn from(error: caption::Error) -> Self {
        match error {
            caption::Error::Graph(error) => error.into(),
            caption::Error::NoWidth => Refusal::NoWidth.into(),
            caption::Error::DependentContent => Refusal::DependentContent.into(),
            caption::Error::Encode(_) | caption::Error::Storage => Self::InvalidSource,
        }
    }
}

/// One label change staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: caption::Target,
    chart: bool,
    before: Option<String>,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("label", &self.target.label)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The label text when the edit began; `None` when the label is absent.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Stage the label's text, creating the label when it is absent.
    ///
    /// # Errors
    ///
    /// [`Refusal::ReservedMarker`], [`Refusal::OperationAlreadyStaged`].
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        if text
            .chars()
            .any(|character| matches!(character, '\u{fffc}' | '\u{000e}' | '\u{0004}'))
        {
            return Err(Refusal::ReservedMarker.into());
        }
        self.operation = Some(Some(Box::from(text)));
        Ok(self)
    }

    /// Stage the label's removal; removing an absent label changes nothing.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`].
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.operation = Some(None);
        Ok(self)
    }

    /// Validate, publish and strictly reopen the staged change. A change
    /// that leaves the label as it is is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::NoWidth`],
    /// [`Refusal::DependentContent`], source, limit and verification
    /// errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self.operation.ok_or(Refusal::NoStagedOperation)?;
        let text = operation.as_deref();
        if text == self.before.as_deref() {
            return Commit::unchanged(self.source);
        }
        let (components, limits) = workspace(self.source)?;
        let mut working = Workspace::new(&components, limits)?;
        let other = caption::Target {
            label: self.target.label.other(),
            ..self.target
        };
        let other_before = caption::read(&working, other)?;
        if self.chart
            && text.is_some()
            && self.before.is_none()
            && caption::component_holds(
                &working,
                self.target.drawable,
                CALCULATION_ENGINE_MESSAGE_TYPE,
            )?
        {
            return Err(Refusal::ComponentizedLayout.into());
        }
        let source = self.source;
        let written = caption::write(&mut working, self.target, text, |_| {
            theme(source).map_err(|_| {
                caption::Error::Graph(litchi_iwa_style::graph::Error::Invalid(
                    "a document without a caption theme",
                ))
            })
        })?;
        if !written.changed {
            return Commit::unchanged(self.source);
        }
        let changes = working.finish();
        let published =
            numbers_style_graph::publish(self.source, changes, &[]).map_err(Error::from)?;
        let (components, limits) = workspace(&published.package)?;
        let reopened = Workspace::new(&components, limits)?;
        if caption::read(&reopened, self.target)?.as_deref() != text
            || caption::read(&reopened, other)? != other_before
        {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

fn workspace(
    package: &Package,
) -> Result<(numbers_style_graph::LeafComponents<'_>, Limits), Error> {
    let catalog = structure::exact_source(package).map_err(Error::from)?;
    let limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|error| Error::from(structure::map_archive(&error)))?;
    let components = numbers_style_graph::LeafComponents::new(package).map_err(Error::from)?;
    Ok((components, limits))
}

/// The document stylesheet, the theme's first caption paragraph-style
/// preset and the document language (the host's
/// `sheet_shape_caption_theme`).
fn theme(package: &Package) -> Result<(u64, u64, Option<String>), structure::Fault> {
    use structure::Fault;
    let reference = |payload: &[u8], path: &[u32]| -> Result<u64, Fault> {
        let records = references::records(payload, path)?;
        let [record] = records.as_slice() else {
            return Err(Fault::Invalid);
        };
        references::reference_identifier(record)
    };
    let document = graph::object(package, 1)?;
    let document = &document.messages[structure::document_message(document)?].data;
    let stylesheet = reference(document, &[DOCUMENT_STYLESHEET_FIELD])?;
    let theme = reference(document, &[DOCUMENT_THEME_FIELD])?;
    let language = references::records(document, &[DOCUMENT_SUPER_FIELD, DOCUMENT_LANGUAGE_FIELD])?
        .last()
        .map(|bytes| std::str::from_utf8(bytes).map(str::to_owned))
        .transpose()
        .map_err(|_| Fault::Invalid)?;
    let theme = graph::object(package, theme)?;
    let theme = &theme.messages[structure::unique_message(theme, &[THEME_MESSAGE_TYPE])?].data;
    let base = references::records(theme, &[THEME_SUPER_FIELD])?;
    let [base] = base.as_slice() else {
        return Err(Fault::Invalid);
    };
    let paragraph_style = caption::theme_paragraph_style(base).ok_or(Fault::Invalid)?;
    graph::object(package, stylesheet)?;
    graph::object(package, paragraph_style)?;
    Ok((stylesheet, paragraph_style, language))
}

fn resolve(
    package: &Package,
    sheet: SheetSelector<'_>,
    selector: DrawableSelector,
    label: Label,
) -> Result<(caption::Target, bool), Error> {
    let sheet = graph::sheet_position(package, sheet)
        .map_err(Error::from)?
        .ok_or(Refusal::SheetNotFound)?;
    let drawables = graph::sheet_drawables(package, sheet).map_err(Error::from)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    let nesting = match (entry.kind, label) {
        (Kind::Shape, _) => caption::Nesting::Shape,
        (Kind::Image | Kind::Movie, _) | (Kind::Chart, Label::Caption) => {
            caption::Nesting::Drawable
        },
        _ => {
            return Err(Refusal::NoLabel {
                kind: entry.kind.name(),
                label,
            }
            .into());
        },
    };
    Ok((
        caption::Target {
            drawable: entry.identifier,
            nesting,
            label: label.shared(),
        },
        entry.kind == Kind::Chart,
    ))
}

impl Package {
    /// Read one label (title or caption) of a sheet shape, image, movie or
    /// chart (caption only); `None` when the label is absent.
    ///
    /// # Errors
    ///
    /// Selector refusals, [`Refusal::NoLabel`], source and limit errors.
    pub fn sheet_drawable_caption<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        label: Label,
    ) -> Result<Option<String>, Error> {
        let (target, _) = resolve(self, sheet.into(), drawable.into(), label)?;
        let (components, limits) = workspace(self)?;
        let workspace = Workspace::new(&components, limits)?;
        Ok(caption::read(&workspace, target)?)
    }

    /// Begin a change of one label of a sheet shape, image, movie or chart.
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_drawable_caption`].
    pub fn edit_sheet_drawable_caption<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
        label: Label,
    ) -> Result<Edit<'_>, Error> {
        let target = resolve(self, sheet.into(), drawable.into(), label)?;
        let (components, limits) = workspace(self)?;
        let workspace = Workspace::new(&components, limits)?;
        let before = caption::read(&workspace, target.0)?;
        Ok(Edit {
            source: self,
            target: target.0,
            chart: target.1,
            before,
            operation: None,
        })
    }

    /// Apply a reversible exact-source label patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] for another source, verification errors.
    pub fn apply_sheet_drawable_caption(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

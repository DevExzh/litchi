//! Selector-first caption transactions of body charts (record 0427) on the
//! shared `litchi_iwa_style::caption` engine.
//!
//! A chart's caption is the `TSD.DrawableArchive` caption reference of its
//! chart message (field 1): a present caption is a caption info with a
//! private text storage, shape style and placement in the chart's
//! component, an absent one an empty stand-in. Setting a present caption
//! rewrites its storage text; setting an absent one authors the native
//! inline graph; clearing a present caption points the chart at a fresh
//! stand-in and keeps the old graph, as the application does for undo. The
//! chart message's metadata moves the reference to the end of its lists
//! (the host's chart caption edge). New objects get UUIDs in the chart's
//! component and raise the identifier watermark (the host's
//! `body_chart_caption`, `set_body_chart_caption` and
//! `remove_body_chart_caption`). A chart's title is the chart-title
//! family's ([`Package::edit_body_chart_title`]).

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;

use super::Package;
use super::body_drawable_caption::{self as label, Diagnostics, Error, Slot};
use super::body_drawable_graph as graph;
use crate::selector::BodyChartSelector;

/// One chart caption change staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selector: BodyChartSelector,
    slot: Slot,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The caption when the edit began; `None` when the chart has none.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.slot.text.as_deref()
    }

    /// Stage the caption's text, creating the caption when it is absent.
    ///
    /// # Errors
    ///
    /// Refuses a reserved text marker and a second operation.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        label::stage(&mut self.operation, Some(text))?;
        Ok(self)
    }

    /// Stage the caption's removal; removing an absent caption changes
    /// nothing.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        label::stage(&mut self.operation, None)?;
        Ok(self)
    }

    /// Validate, publish and strictly reopen the staged change. A change
    /// that leaves the caption as it is is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`super::body_drawable_caption::Refusal::NoWidth`],
    /// [`super::body_drawable_caption::Refusal::DependentContent`], source,
    /// limit and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let source = self.source;
        let after: Option<Arc<str>> = self.operation.clone().flatten().map(Arc::from);
        let published = if self.operation.is_some() && after != self.slot.text {
            label::publish(source, &self.slot, after.as_ref())?
        } else {
            None
        };
        let source_bytes = source.state.source.shared_source();
        let Some((candidate, created)) = published else {
            return Ok(Commit {
                package: source.clone(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    selector: self.selector,
                    before: self.slot.text.clone(),
                    after: self.slot.text,
                },
                diagnostics: Diagnostics {
                    changed: false,
                    created_objects: 0,
                },
            });
        };
        let target_bytes = candidate.state.source.shared_source();
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, target_bytes),
                selector: self.selector,
                before: self.slot.text,
                after,
            },
            diagnostics: Diagnostics {
                changed: true,
                created_objects: created,
            },
        })
    }
}

/// Exact-source reversible body-chart caption patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selector: BodyChartSelector,
    before: Option<Arc<str>>,
    after: Option<Arc<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("selector", &self.selector)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The selected chart.
    #[must_use]
    pub const fn selector(&self) -> BodyChartSelector {
        self.selector
    }

    /// The caption required before this patch applies.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// The caption this patch produces.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Whether both the caption and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selector: self.selector,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// The validated immutable result of one body-chart caption transaction.
#[must_use = "a body-chart caption commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the caption of a body chart; `None` when it has none.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when no chart matches or the graph is invalid.
    pub fn body_chart_caption(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Option<String>, Error> {
        Ok(label::resolve_chart(self, selector.into())?
            .text
            .map(|text| text.as_ref().to_owned()))
    }

    /// Begin a change of a body chart's caption.
    ///
    /// # Errors
    ///
    /// As [`Self::body_chart_caption`].
    pub fn edit_body_chart_caption(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        Ok(Edit {
            source: self,
            selector,
            slot: label::resolve_chart(self, selector)?,
            operation: None,
        })
    }

    /// Replay an exact-source body-chart caption patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the caption no longer matches its `before`
    /// state.
    pub fn apply_body_chart_caption(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        if label::resolve_chart(self, patch.selector)?.text != patch.before {
            return Err(Error::PatchConflict);
        }
        let changed = !patch.is_noop();
        let package = if changed {
            let candidate = graph::reopen(self, patch.artifacts.target())?;
            if label::resolve_chart(&candidate, patch.selector)?.text != patch.after {
                return Err(Error::Verification);
            }
            candidate
        } else {
            self.clone()
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed,
                created_objects: 0,
            },
        })
    }
}

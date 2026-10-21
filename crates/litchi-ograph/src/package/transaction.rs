//! Failure-atomic typed edits at the standalone Graph package boundary.

use crate::{Result, chart};

use std::sync::Arc;

use super::{Package, codec, patch};

/// Source-checked edit over the one standalone Graph Workbook chart.
///
/// Only existing fixed-width chart payloads exposed by the chart editor are
/// reachable here. The edit never recalculates formulas, follows links,
/// renders a chart, or activates an OLE object.
#[derive(Debug)]
pub struct Edit {
    source: Package,
    editor: Option<chart::Edit>,
}

impl Edit {
    pub(super) fn new(source: Package) -> Self {
        Self {
            source,
            editor: None,
        }
    }

    fn editor(&mut self) -> Result<&mut chart::Edit> {
        match &mut self.editor {
            Some(editor) => Ok(editor),
            slot @ None => Ok(slot.insert(codec::read_chart(&self.source)?.edit()?)),
        }
    }

    /// Immutable package source retained by this edit.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.source
    }

    /// Number of typed chart operations staged in this edit.
    #[must_use]
    pub fn len(&self) -> usize {
        self.editor.as_ref().map_or(0, chart::Edit::len)
    }

    /// Whether no chart payload has been staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.editor.as_ref().is_none_or(chart::Edit::is_empty)
    }

    /// Replace one existing Graph cache value without evaluating its source.
    pub fn set_cache_value<V>(&mut self, index: usize, value: V) -> Result<&mut Self>
    where
        V: Into<chart::CacheValue>,
    {
        self.editor()?.set_cache_value(index, value)?;
        Ok(self)
    }

    /// Replace the fixed-size Graph chart-area rectangle.
    pub fn set_rect(&mut self, value: chart::Rect) -> Result<&mut Self> {
        self.editor()?.set_rect(value)?;
        Ok(self)
    }

    /// Replace the fixed-size Graph `ShtProps` metadata payload.
    pub fn set_props(&mut self, value: chart::Props) -> Result<&mut Self> {
        self.editor()?.set_props(value)?;
        Ok(self)
    }

    /// Validate and publish the package edit atomically.
    pub fn commit(self) -> Result<patch::Commit> {
        let Self { source, editor } = self;
        let editor = match editor {
            Some(editor) => editor,
            None => codec::read_chart(&source)?.edit()?,
        };
        let chart_commit = editor.commit()?;
        let chart_patch = chart_commit.patch().clone();
        if chart_patch.is_noop() {
            let patch = patch::Patch::new(
                Arc::clone(source.bytes.shared()),
                Arc::clone(source.bytes.shared()),
                chart_patch,
            );
            return Ok(patch::Commit::new(
                source,
                patch,
                patch::Diagnostics::UNCHANGED,
            ));
        }

        let chart = chart_commit.into_package();
        let stream = chart.encode()?;
        // Publication owns the writer buffer; release its spare capacity before
        // history starts sharing this endpoint.
        let bytes = codec::replace_chart(&source, stream.as_bytes())?
            .into_boxed_slice()
            .into_vec();
        let package =
            Package::open_with_options(bytes, crate::OpenOptions::new().limits(source.limits))?;
        let patch = patch::Patch::new(
            Arc::clone(source.bytes.shared()),
            Arc::clone(package.bytes.shared()),
            chart_patch,
        );
        Ok(patch::Commit::new(
            package,
            patch,
            patch::Diagnostics::CHANGED,
        ))
    }
}

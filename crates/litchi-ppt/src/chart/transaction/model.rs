//! Semantic chart-package transaction and immutable snapshots.

use litchi_ograph::Limits;
use litchi_ograph::chart::{Chart as SemanticChart, Context};

use super::codec;
use crate::package::{Error, Result};

/// An immutable view of one transaction state.
///
/// The workbook bytes and chart reference borrow the editor, so a snapshot is
/// allocation-free and becomes unusable as soon as its editor is mutably
/// borrowed. The chart reference is still the host-neutral typed view from
/// `litchi-ograph`; no unsupported semantic authoring is implied.
#[derive(Debug, Clone, Copy)]
pub struct Snapshot<'a> {
    workbook: &'a [u8],
    chart: litchi_ograph::chart::Ref<'a>,
    dirty: bool,
}

impl Snapshot<'_> {
    /// Exact current `Workbook` stream bytes.
    #[must_use]
    pub fn workbook(&self) -> &[u8] {
        self.workbook
    }

    /// The one typed Graph chart substream in the workbook.
    #[must_use]
    pub const fn chart(&self) -> litchi_ograph::chart::Ref<'_> {
        self.chart
    }

    /// Validate and own the current chart through the MS-OGRAPH semantic
    /// model.
    ///
    /// The returned chart retains its exact source stream. Encoding it
    /// without mutation is therefore lossless; an attempted parsed mutation
    /// reports the neutral model's typed unsafe-edit error.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn semantic_chart(&self) -> Result<SemanticChart> {
        SemanticChart::parse(self.chart, Context::graph()).map_err(Into::into)
    }

    /// Whether an edit has been applied since the transaction was opened.
    #[must_use]
    pub const fn is_dirty(&self) -> bool {
        self.dirty
    }
}

/// Bounded transaction for a standalone Microsoft Graph OLE2 chart package.
///
/// `[MS-OGRAPH]` 2.1.7.3 requires exactly one `Workbook` stream containing a
/// globals substream followed by one chart-sheet substream. This editor
/// replaces only that chart-sheet byte range, then rebuilds and revalidates the
/// standalone OLE2 topology. The operation does not edit `[MS-PPT]`
/// `ExOleObjStg` records or the `[MS-ODRAW]` picture-frame shape that hosts the
/// package.
#[derive(Debug)]
pub struct PackageEditor {
    original: Option<Vec<u8>>,
    workbook: Vec<u8>,
    comp_obj: Option<Vec<u8>>,
    ole: Option<Vec<u8>>,
    chart_start: usize,
    chart_end: usize,
    limits: Limits,
    dirty: bool,
}

impl PackageEditor {
    /// Open a validated package with conservative `OGraph` limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Open a validated package under explicit resource bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let parts = codec::read(&bytes, validated_limits)?;
        Ok(Self {
            original: Some(bytes),
            workbook: parts.workbook,
            comp_obj: parts.comp_obj,
            ole: parts.ole,
            chart_start: parts.chart_start,
            chart_end: parts.chart_end,
            limits: validated_limits,
            dirty: false,
        })
    }

    /// Resource bounds applied to this transaction.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Whether a chart replacement has been staged.
    #[must_use]
    pub const fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Borrow the current typed, Graph-framed chart stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn chart(&self) -> Result<litchi_ograph::chart::Ref<'_>> {
        self.chart_ref()
    }

    /// Validate and own the current chart through the MS-OGRAPH semantic
    /// model without changing transaction state.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn semantic_chart(&self) -> Result<SemanticChart> {
        SemanticChart::parse_with(self.chart_ref()?, Context::graph(), self.limits)
            .map_err(Into::into)
    }

    /// Capture the current state without copying its bounded buffers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot<'_>> {
        Ok(Snapshot {
            workbook: &self.workbook,
            chart: self.chart_ref()?,
            dirty: self.dirty,
        })
    }

    /// Replace the one Graph-framed chart stream atomically.
    ///
    /// The input stream is checked for Graph BIFF framing and the complete
    /// candidate OLE2 package is rebuilt and validated before this editor is
    /// changed. This is a package transaction, not complete native chart
    /// authoring: the existing `OGraph` model remains the authority for the
    /// supported chart-stream grammar and opaque records remain inert.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_chart(&mut self, chart: litchi_ograph::chart::Stream) -> Result<()> {
        if chart.kind() != litchi_ograph::chart::Kind::Graph {
            return Err(Error::InvalidFormat(
                "standalone Graph packages require a Graph chart stream".to_string(),
            ));
        }
        let bytes = chart.into_bytes();
        let checked = litchi_ograph::chart::Ref::with_limits(&bytes, self.limits)?;
        if checked.kind() != litchi_ograph::chart::Kind::Graph {
            return Err(Error::InvalidFormat(
                "standalone Graph packages require a Graph chart stream".to_string(),
            ));
        }

        let start = self.chart_start;
        let end = self.chart_end;
        if self.workbook.get(start..end).is_none() {
            return Err(Error::Corrupted(
                "chart transaction range no longer matches Workbook".to_string(),
            ));
        }
        let suffix_len = self.workbook.len().checked_sub(end).ok_or_else(|| {
            Error::Corrupted("chart transaction range exceeds Workbook".to_string())
        })?;
        let new_end = start
            .checked_add(bytes.len())
            .ok_or(litchi_ograph::Error::SizeOverflow {
                resource: "edited chart substream",
            })?;
        let total = start
            .checked_add(bytes.len())
            .and_then(|length| length.checked_add(suffix_len))
            .ok_or(litchi_ograph::Error::SizeOverflow {
                resource: "edited Workbook bytes",
            })?;
        if total > self.limits.max_workbook_bytes {
            return Err(litchi_ograph::Error::LimitExceeded {
                resource: "Workbook bytes",
                observed: u64::try_from(total).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.limits.max_workbook_bytes).unwrap_or(u64::MAX),
            }
            .into());
        }

        let mut candidate = Vec::new();
        candidate
            .try_reserve_exact(total)
            .map_err(|_err| litchi_ograph::Error::Allocation {
                resource: "edited Workbook bytes",
            })?;
        candidate.extend_from_slice(&self.workbook[..start]);
        candidate.extend_from_slice(&bytes);
        candidate.extend_from_slice(&self.workbook[end..]);

        // Rebuilding before assignment is the atomicity boundary, so a CFB or
        // topology failure cannot partially update the editor.
        let _ = codec::write(
            &candidate,
            self.comp_obj.as_deref(),
            self.ole.as_deref(),
            self.limits,
        )?;

        self.workbook = candidate;
        self.chart_end = new_end;
        self.original.take();
        self.dirty = true;
        Ok(())
    }

    /// Replace the current chart with a semantically validated MS-OGRAPH
    /// chart.
    ///
    /// This keeps the existing raw [`Self::replace_chart`] API intact while
    /// making the stronger semantic path explicit. Untouched parsed charts
    /// replay byte-for-byte; parsed mutations and fresh charts are rejected by
    /// `litchi-ograph` until their opaque-record placement is proven.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_semantic_chart(&mut self, chart: SemanticChart) -> Result<()> {
        if chart.context().kind() != litchi_ograph::chart::Kind::Graph {
            return Err(Error::InvalidFormat(
                "standalone Graph packages require a Graph semantic chart".to_string(),
            ));
        }
        let stream = chart.encode_with(self.limits)?;
        self.replace_chart(stream)
    }

    /// Commit the current transaction into a standalone OLE2 package.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Vec<u8>> {
        if let Some(original) = self.original {
            return Ok(original);
        }
        codec::write(
            &self.workbook,
            self.comp_obj.as_deref(),
            self.ole.as_deref(),
            self.limits,
        )
    }

    /// Commit alias matching the package writers' move-owned finish pattern.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.commit()
    }

    fn chart_ref(&self) -> Result<litchi_ograph::chart::Ref<'_>> {
        let mut charts = litchi_ograph::chart::Refs::with_limits(&self.workbook, self.limits)?;
        let chart = charts
            .next()
            .ok_or_else(|| Error::Corrupted("Graph Workbook has no chart stream".to_string()))??;
        if let Some(extra) = charts.next() {
            let _ = extra?;
            return Err(Error::Corrupted(
                "Graph Workbook has more than one chart stream".to_string(),
            ));
        }
        let chart_len = self
            .chart_end
            .checked_sub(self.chart_start)
            .ok_or_else(|| Error::Corrupted("chart transaction range is reversed".to_string()))?;
        if chart.kind() != litchi_ograph::chart::Kind::Graph
            || chart.offset() != self.chart_start
            || chart.as_bytes().len() != chart_len
        {
            return Err(Error::Corrupted(
                "chart transaction range no longer matches Graph Workbook".to_string(),
            ));
        }
        Ok(chart)
    }
}

//! Immutable chart snapshots and detached, failure-atomic edits.

use super::model::DataLabelFlag;
use super::{codec, validation};
use crate::chart::data::TitleText;
use crate::chart::model::Chart;
use crate::chart::types::{AxisPosition, DisplayBlanks};
use crate::{Error, Result};
use std::sync::Arc;

/// An immutable source snapshot of one host-neutral chart-space part.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Chart,
}

impl Snapshot {
    /// Parse and retain one bounded chart part without normalizing its bytes.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Create a source snapshot from a detached typed chart value.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn new(value: Chart) -> Result<Self> {
        let mut xml = Vec::new();
        crate::chart::writer::write(&mut xml, &value).map_err(|error| {
            Error::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, error))
        })?;
        drop(value);
        Self::from_xml(xml)
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the host-neutral typed chart projection.
    #[must_use]
    pub const fn value(&self) -> &Chart {
        &self.value
    }

    /// Start an isolated edit from this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            working_xml: self.xml.as_ref().to_vec(),
            working: self.value.clone(),
        }
    }
}

/// A detached chart edit that is not visible until it commits.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    working_xml: Vec<u8>,
    working: Chart,
}

impl Transaction {
    /// Borrow the projected chart value after pending edits.
    #[must_use]
    pub const fn value(&self) -> &Chart {
        &self.working
    }

    /// Set or clear the chart title while retaining existing rich formatting
    /// and extension children when the representation is unchanged.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_title(&mut self, title: Option<TitleText>) -> Result<&mut Self> {
        let candidate = codec::set_chart_title(&self.working_xml, title.as_ref())?;
        drop(title);
        self.apply(candidate)
    }

    /// Set or clear the chart-space language.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_language(&mut self, language: Option<&str>) -> Result<&mut Self> {
        let candidate = codec::set_language(&self.working_xml, language)?;
        self.apply(candidate)
    }

    /// Set or clear the built-in chart style.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_style(&mut self, style: Option<u32>) -> Result<&mut Self> {
        let candidate = codec::set_style(&self.working_xml, style)?;
        self.apply(candidate)
    }

    /// Set the chart blank-cell display mode.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_display_blanks(&mut self, mode: DisplayBlanks) -> Result<&mut Self> {
        let candidate = codec::set_display_blanks(&self.working_xml, mode)?;
        self.apply(candidate)
    }

    /// Set or clear a series title selected by its stable index.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_series_title(
        &mut self,
        series_index: u32,
        title: Option<TitleText>,
    ) -> Result<&mut Self> {
        let candidate = codec::set_series_title(&self.working_xml, series_index, title.as_ref())?;
        drop(title);
        self.apply(candidate)
    }

    /// Set one shared data-label switch on a series.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_series_data_label_flag(
        &mut self,
        series_index: u32,
        flag: DataLabelFlag,
        value: bool,
    ) -> Result<&mut Self> {
        let candidate =
            codec::set_series_data_label_flag(&self.working_xml, series_index, flag, value)?;
        self.apply(candidate)
    }

    /// Remove one explicit shared data-label switch from a series.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn clear_series_data_label_flag(
        &mut self,
        series_index: u32,
        flag: DataLabelFlag,
    ) -> Result<&mut Self> {
        let candidate = codec::clear_series_data_label_flag(&self.working_xml, series_index, flag)?;
        self.apply(candidate)
    }

    /// Set or clear the shared data-label separator on a series.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_series_data_label_separator(
        &mut self,
        series_index: u32,
        separator: Option<String>,
    ) -> Result<&mut Self> {
        let candidate = codec::set_series_data_label_separator(
            &self.working_xml,
            series_index,
            separator.as_deref(),
        )?;
        drop(separator);
        self.apply(candidate)
    }

    /// Set or clear the numeric range of an axis selected by its ID.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_axis_range(
        &mut self,
        axis_id: u32,
        min: Option<f64>,
        max: Option<f64>,
    ) -> Result<&mut Self> {
        validation::validate_axis_range(min, max)?;
        let candidate = codec::set_axis_range(&self.working_xml, axis_id, min, max)?;
        self.apply(candidate)
    }

    /// Set the position of an axis selected by its ID.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn set_axis_position(&mut self, axis_id: u32, position: AxisPosition) -> Result<&mut Self> {
        let candidate = codec::set_axis_position(&self.working_xml, axis_id, position)?;
        self.apply(candidate)
    }

    /// Whether pending source bytes differ from the original snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.working_xml != self.base.xml.as_ref()
    }

    /// Publish the edit and return its exact snapshot plus reversible patch.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = if self.is_changed() {
            Snapshot::from_xml(self.working_xml)?
        } else {
            self.base.clone()
        };
        Ok(Commit {
            patch: Patch {
                before: self.base.xml.clone(),
                after: snapshot.xml.clone(),
            },
            snapshot,
        })
    }

    fn apply(&mut self, candidate: Vec<u8>) -> Result<&mut Self> {
        if candidate == self.working_xml {
            return Ok(self);
        }
        let value = codec::read(&candidate)?;
        self.working_xml = candidate;
        self.working = value;
        Ok(self)
    }
}

/// A successful source-preserving chart publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// An exact, source-preconditioned and reversible chart patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Borrow the exact source bytes required by this patch.
    #[must_use]
    pub fn before_xml(&self) -> &[u8] {
        &self.before
    }

    /// Borrow the exact bytes produced by this patch.
    #[must_use]
    pub fn after_xml(&self) -> &[u8] {
        &self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }

    /// Apply only to the exact source snapshot from which this patch came.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_ref() != self.before.as_ref() {
            return Err(Error::Invalid(
                "chart patch source does not match its byte precondition".into(),
            ));
        }
        Snapshot::from_xml(self.after.as_ref().to_vec())
    }
}

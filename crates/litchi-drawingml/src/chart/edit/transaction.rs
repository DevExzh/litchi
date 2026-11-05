//! Immutable chart snapshots and detached, failure-atomic edits.

use super::model::DataLabelFlag;
use super::{codec, validation};
use crate::Error;
use crate::Result;
use crate::chart::data::TitleText;
use crate::chart::model::Chart;
use crate::chart::types::{AxisPosition, DisplayBlanks, LegendPosition};
use std::sync::Arc;

pub use crate::LimitKind;
pub use litchi_core::edit::Diagnostics;

/// An immutable source snapshot of one host-neutral chart-space part.
#[derive(Debug, Clone)]
pub struct Package {
    xml: Arc<[u8]>,
    value: Chart,
}

impl Package {
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
        crate::chart::writer::write(&mut xml, &value, &crate::chart::Limits::DEFAULT)?;
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
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            working_xml: self.xml.as_ref().to_vec(),
            working: self.value.clone(),
        }
    }
}

/// A detached chart edit that is not visible until it commits.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Package,
    working_xml: Vec<u8>,
    working: Chart,
}

impl Edit {
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

    /// Show the legend at a position (`Some`) or remove it (`None`).
    ///
    /// An existing legend keeps its overlay, layout, entries and formatting;
    /// only its position changes.
    /// # Errors
    ///
    /// Returns an error when the chart XML is malformed or the edited
    /// chart fails to read back.
    pub fn set_legend(&mut self, position: Option<LegendPosition>) -> Result<&mut Self> {
        let candidate = codec::set_legend(&self.working_xml, position)?;
        self.apply(candidate)
    }

    /// Set or clear the title of an axis selected by its ID.
    /// # Errors
    ///
    /// Returns an error for a missing axis, a title whose representation
    /// would change (literal to reference or back), or malformed XML.
    pub fn set_axis_title(&mut self, axis_id: u32, title: Option<TitleText>) -> Result<&mut Self> {
        let candidate = codec::set_axis_title(&self.working_xml, axis_id, title.as_ref())?;
        drop(title);
        self.apply(candidate)
    }

    /// Set (`Some`) or clear (`None`) the color of a series selected by its
    /// stable index: its area fill, or for a line, scatter, stock or
    /// unfilled radar series its line color. The color is `RRGGBB` or
    /// `AARRGGBB`; the rest of the series' shape properties are kept.
    ///
    /// ```
    /// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
    /// use litchi_drawingml::chart::edit::Package;
    ///
    /// let spec = Spec::new(Variant::new(Kind::Column))
    ///     .with_series(Series::new("East", vec![1.0, 2.0]));
    /// let package = Package::from_xml(spec.to_xml(None)?)?;
    /// let mut edit = package.edit();
    /// edit.set_series_color(0, Some("C00000"))?;
    /// let xml = edit.commit()?.into_package().xml_bytes().to_vec();
    /// assert!(String::from_utf8_lossy(&xml).contains(r#"<a:srgbClr val="C00000"/>"#));
    /// # Ok::<(), litchi_drawingml::Error>(())
    /// ```
    /// # Errors
    ///
    /// Returns an error for a missing series, an invalid color, a series
    /// with more than one fill, or malformed XML.
    pub fn set_series_color(
        &mut self,
        series_index: u32,
        color: Option<&str>,
    ) -> Result<&mut Self> {
        let candidate = codec::set_series_color(&self.working_xml, series_index, color)?;
        self.apply(candidate)
    }

    /// Replace the values of a series selected by its stable index (its Y
    /// values for a scatter or bubble series).
    ///
    /// A literal is rewritten; a reference keeps its formula and gets a new
    /// cache. When the number of points changes, a single-column reference
    /// is resized from its first cell; any other reference is refused,
    /// because the cache would no longer match the range.
    /// # Errors
    ///
    /// Returns an error for a missing series, an empty or non-finite value
    /// list, a reference that cannot be resized, or malformed XML.
    pub fn set_series_values(&mut self, series_index: u32, values: &[f64]) -> Result<&mut Self> {
        let candidate = codec::set_series_values(&self.working_xml, series_index, values)?;
        self.apply(candidate)
    }

    /// Replace the category labels of every series (the X values of
    /// scatter and bubble series), following the rules of
    /// [`Self::set_series_values`].
    /// # Errors
    ///
    /// Returns an error for an empty label list, a text label in numeric
    /// data, a reference that cannot be resized, or malformed XML.
    pub fn set_categories(&mut self, labels: &[String]) -> Result<&mut Self> {
        let candidate = codec::set_categories(&self.working_xml, labels)?;
        self.apply(candidate)
    }

    /// Publish the edit and return its exact package plus reversible patch.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.working_xml != self.base.xml.as_ref();
        let package = if changed {
            Package::from_xml(self.working_xml)?
        } else {
            self.base.clone()
        };
        let patch = Patch {
            before: self.base.xml.clone(),
            after: package.xml.clone(),
        };
        Ok(Commit::new(package, patch, Diagnostics::new(changed)))
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
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

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

    /// Whether this patch preserves the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Apply only to the exact source package from which this patch came.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    fn apply(&self, source: &Package) -> Result<Package> {
        if source.xml.as_ref() != self.before.as_ref() {
            return Err(Error::PatchConflict);
        }
        Package::from_xml(self.after.as_ref().to_vec())
    }
}

impl Package {
    /// Replay a patch against this exact package and publish a commit triad.
    ///
    /// Exact no-op bytes remain unchanged; the resulting package is reparsed
    /// and validated before it becomes visible, just like a changed patch.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.apply(self)?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

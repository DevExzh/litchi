//! `OpenDocument` standalone chart (`.odc` and `.otc`) support.
//!
//! [`Document::edit`] stages an [`Edit`] of the chart definition, its axes
//! and its series; [`Edit::commit`] publishes a new [`Document`] with its
//! [`Patch`] and [`Diagnostics`], and [`Document::apply`] replays a patch
//! onto its exact source.

mod document;
mod mutation;

pub use document::Document;
pub use litchi_odf_common::chart::authoring::{
    AxisSpec, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, ChartClassKind,
    DataLabelSpec, DataPointSpec, Definition, DomainSpec, EquationSpec, ExtensionAttribute,
    ExtensionElement, Extensions, GridSpec, LegendSpec, PlotAreaSpec, RegressionSpec, SeriesSpec,
    StyleElement, Text, serialize_axis_fragment, serialize_content, serialize_series_fragment,
};
pub use litchi_odf_common::chart::{
    Attribute, Axis, Class, DataPoint, Dimension, Element, Grid, Kind, Labels, Legend, PlotArea,
    Position, Series,
};
pub use mutation::{AxisUpdate, SeriesUpdate};

crate::generic::staged::staged_family!(Document, "odc");

impl crate::generic::staged::Staged for Document {
    fn working_copy(&self) -> litchi_core::Result<Self> {
        self.reopen(self.as_bytes())
    }

    fn reopen(&self, bytes: &[u8]) -> litchi_core::Result<Self> {
        let options = litchi_odf_common::OpenOptions::new()
            .source_limits(self.package.owned_package().limits());
        Ok(Self::from_bytes_with_options(bytes, options)?)
    }

    fn staged_bytes(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl Document {
    /// Stage an edit of this chart; the document itself never changes.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Replay `patch` onto this exact chart.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this chart is not the patch
    /// source, or the validation failure of the patched package.
    pub fn apply(&self, patch: &Patch) -> Result<Commit, Error> {
        patch.replay(self)
    }
}

impl Edit<'_> {
    /// Stages replacing the typed chart content; other package entries are
    /// kept.
    pub fn set_definition(&mut self, definition: &Definition) -> &mut Self {
        let definition = definition.clone();
        self.stage(move |document| document.set_definition(&definition))
    }

    /// Stages appending one axis to the plot area.
    pub fn add_axis(&mut self, axis: &AxisSpec) -> &mut Self {
        let axis = axis.clone();
        self.stage(move |document| document.add_axis(&axis).map(drop))
    }

    /// Stages replacing the axis at `index` (plot-area order).
    pub fn replace_axis(&mut self, index: usize, axis: &AxisSpec) -> &mut Self {
        let axis = axis.clone();
        self.stage(move |document| document.replace_axis(index, &axis))
    }

    /// Stages updating selected attributes of the axis at `index`.
    pub fn update_axis(&mut self, index: usize, update: &AxisUpdate) -> &mut Self {
        let update = update.clone();
        self.stage(move |document| document.update_axis(index, &update))
    }

    /// Stages removing the axis at `index`.
    pub fn remove_axis(&mut self, index: usize) -> &mut Self {
        self.stage(move |document| document.remove_axis(index))
    }

    /// Stages moving the axis at `from` to `to`.
    pub fn reorder_axis(&mut self, from: usize, to: usize) -> &mut Self {
        self.stage(move |document| document.reorder_axis(from, to))
    }

    /// Stages appending one series to the plot area.
    pub fn add_series(&mut self, series: &SeriesSpec) -> &mut Self {
        let series = series.clone();
        self.stage(move |document| document.add_series(&series).map(drop))
    }

    /// Stages replacing the series at `index` (plot-area order).
    pub fn replace_series(&mut self, index: usize, series: &SeriesSpec) -> &mut Self {
        let series = series.clone();
        self.stage(move |document| document.replace_series(index, &series))
    }

    /// Stages updating selected attributes of the series at `index`.
    pub fn update_series(&mut self, index: usize, update: &SeriesUpdate) -> &mut Self {
        let update = update.clone();
        self.stage(move |document| document.update_series(index, &update))
    }

    /// Stages removing the series at `index`.
    pub fn remove_series(&mut self, index: usize) -> &mut Self {
        self.stage(move |document| document.remove_series(index))
    }

    /// Stages moving the series at `from` to `to`.
    pub fn reorder_series(&mut self, from: usize, to: usize) -> &mut Self {
        self.stage(move |document| document.reorder_series(from, to))
    }
}

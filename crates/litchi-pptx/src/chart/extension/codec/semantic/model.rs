//! Typed `ChartEx` document-facing model seam.

use super::super::super::super::style::{ColorDocument, Document as StyleDocument};
use super::super::super::model::{
    Axis, Chart, ChartSpaceFormatting, DataSet, Document, ExternalDataTarget, Info, PlotArea,
    SeriesDataReference,
};

impl Document {
    #[must_use]
    pub fn info(&self) -> &Info {
        &self.info
    }

    #[must_use]
    pub fn external_data_target(&self) -> Option<&ExternalDataTarget> {
        self.external_data_target.as_ref()
    }

    #[must_use]
    pub fn fallback_image_part_name(&self) -> Option<&str> {
        self.fallback_image_part_name.as_deref()
    }

    #[must_use]
    pub fn chart_style(&self) -> Option<&StyleDocument> {
        self.chart_style.as_ref()
    }

    #[must_use]
    pub fn chart_color_style(&self) -> Option<&ColorDocument> {
        self.chart_color_style.as_ref()
    }

    /// Return the validated source XML unchanged.
    #[must_use]
    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

pub(in crate::chart::extension::codec) type ParsedDataGraph = (
    Vec<DataSet>,
    Vec<SeriesDataReference>,
    Vec<Axis>,
    bool,
    Chart,
    PlotArea,
    ChartSpaceFormatting,
);

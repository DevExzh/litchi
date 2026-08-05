//! Typed ChartEx document-facing model seam.

use super::super::super::super::style::{ColorDocument, Document as StyleDocument};
use super::super::super::model::*;

impl Document {
    pub fn info(&self) -> &Info {
        &self.info
    }

    pub fn external_data_target(&self) -> Option<&ExternalDataTarget> {
        self.external_data_target.as_ref()
    }

    pub fn fallback_image_part_name(&self) -> Option<&str> {
        self.fallback_image_part_name.as_deref()
    }

    pub fn chart_style(&self) -> Option<&StyleDocument> {
        self.chart_style.as_ref()
    }

    pub fn chart_color_style(&self) -> Option<&ColorDocument> {
        self.chart_color_style.as_ref()
    }

    /// Return the validated source XML unchanged.
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

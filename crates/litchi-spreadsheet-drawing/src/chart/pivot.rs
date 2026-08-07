//! Private pivot-chart defaults used by [`super::Chart`].

use super::model::{Chart, for_each_series_mut};
use crate::Result;

const DEFAULT_FORMAT_ID: u32 = 0;
const DEFAULT_OPTIONS_EXTENSION_XML: &[u8] = br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8C8630D5DC}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:pivotOptions><c14:dropZoneVisible val="1"/><c14:dropZoneCategories val="1"/><c14:dropZoneData val="1"/><c14:dropZoneSeries val="1"/><c14:dropZoneAxis val="1"/><c14:dropZoneValues val="1"/></c14:pivotOptions></c:ext></c:extLst>"#;

impl Chart {
    /// Convert this chart into a pivot chart bound to `pivot_table_name`.
    ///
    /// # Errors
    ///
    /// Returns an error when the pivot-chart extension cannot be parsed.
    pub fn into_pivot_chart(mut self, pivot_table_name: &str) -> Result<Self> {
        self.chart.pivot_source = Some(litchi_drawingml::chart::model::PivotSource::new(
            pivot_table_name,
            DEFAULT_FORMAT_ID,
        ));
        let extension = litchi_drawingml::chart::ExtensionList::from_xml(
            DEFAULT_OPTIONS_EXTENSION_XML.to_vec(),
        )?;
        for_each_series_mut(&mut self.chart, |series| {
            if series.extension_list.is_none() {
                series.extension_list = Some(extension.clone());
            }
        });
        Ok(self)
    }
}

//! Application-specific object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;

impl ObjectIndex {
    pub(super) fn extract_chart_references(&mut self, object_id: u64, raw_msg: &RawMessage) {
        use prost::Message;

        match raw_msg.type_ {
            // TSCH (Chart) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSCHArchives.proto and libetonyek's chart parsing
            5000 => {
                // TSCH.PreUFF.ChartInfoArchive - legacy chart format
                // This is a pre-unified format chart, structure may vary
                // Attempt basic reference extraction but may fail gracefully
                if let Ok(chart_info) =
                    crate::protobuf::tsch::pre_uff::ChartInfoArchive::decode(&*raw_msg.data)
                {
                    // Extract chart style reference if present
                    if let Some(ref style) = chart_info.style {
                        self.extract_reference(object_id, style);
                    }
                    // Note: PreUFF ChartInfoArchive doesn't have a direct legend field
                    // Legend info is embedded in other structures
                }
            },
            5004 => {
                // TSCH.ChartMediatorArchive - mediator between chart and data
                if let Ok(mediator) =
                    crate::protobuf::tsch::ChartMediatorArchive::decode(&*raw_msg.data)
                {
                    // Extract info reference (points to the chart drawable)
                    if let Some(ref info) = mediator.info {
                        self.extract_reference(object_id, info);
                    }
                    // Note: local_series_indexes and remote_series_indexes are
                    // indices, not references to objects
                }
            },
            5020 => {
                // TSCH.ChartStylePreset - preset styles for charts
                if let Ok(preset) = crate::protobuf::tsch::ChartStylePreset::decode(&*raw_msg.data)
                {
                    // Extract chart style reference
                    if let Some(ref chart_style) = preset.chart_style {
                        self.extract_reference(object_id, chart_style);
                    }
                    // Extract legend style reference
                    if let Some(ref legend_style) = preset.legend_style {
                        self.extract_reference(object_id, legend_style);
                    }
                    // Note: ChartStylePreset has a complex nested structure
                    // Styles for series and axes are managed through different fields
                    // than what might be expected from the pre-UFF format
                }
            },
            5021 => {
                // TSCH.ChartDrawableArchive - main chart drawable
                if let Ok(chart_drawable) =
                    crate::protobuf::tsch::ChartDrawableArchive::decode(&*raw_msg.data)
                {
                    // Extract parent from super DrawableArchive
                    if let Some(ref drawable) = chart_drawable.super_
                        && let Some(ref parent) = drawable.parent
                    {
                        self.extract_reference(object_id, parent);
                    }
                    // Note: ChartArchive is embedded via protobuf extensions,
                    // which requires special handling. The chart data and preset
                    // references would be in the extension fields that we can't
                    // easily access through the standard decode.
                }
            },

            _ => {},
        }
    }
}

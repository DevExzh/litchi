//! Semantic `DrawingML` chart decoding facade.
//!
//! The chart-reader implementation is kept under [`chart_reader`] so the
//! codec boundary exposes only the typed entry points consumed by the XML
//! stream reader.

mod chart_reader;

pub(super) use chart_reader::{
    parse_chart_protection, parse_color_map_override, parse_external_data, parse_legend,
    parse_pivot_formats, parse_pivot_source, parse_plot_area, parse_print_settings, parse_title,
    parse_view_3d, parse_wall_floor, required_chart_relationship_id,
};

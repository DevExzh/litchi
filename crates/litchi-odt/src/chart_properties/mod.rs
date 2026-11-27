//! Complete typed ODF `style:chart-properties` support.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(super) const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const CHART: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
pub(super) const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
pub(super) const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const CHART_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
pub(super) const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
pub(super) const MAX_XML: usize = 32 * 1024 * 1024;
pub(super) const MAX_VALUE: usize = 1024 * 1024;
pub(super) const MAX_ATTRIBUTES: usize = 96;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_STYLES: usize = 65_536;
pub(super) const MAX_TOTAL: usize = 16 * 1024 * 1024;
pub(super) const MAX_EVENTS: usize = 1_000_000;

pub use codec::{parse_chart_style_properties, set_chart_style_properties_xml};
pub use model::{
    Angle, AxisLabelPosition, AxisPosition, DataLabelNumber, Direction, Double, EmptyCellTreatment,
    ErrorCategory, Integer, Interpolation, LabelArrangement, LabelPosition, LabelSeparator,
    NonNegativeInteger, NonNegativeLength, Percent, PositiveInteger, RegressionType, SeriesSource,
    SolidType, StyleProperties, StylePropertiesSet, StyleRecord, SymbolImage, SymbolName,
    SymbolType, TickMarkPosition,
};

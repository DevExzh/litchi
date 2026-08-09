//! Opaque, bounded XLSX workbooks generated for embedded chart data.
//!
//! The owner API deliberately does not depend on `PresentationML` chart-part
//! types. Callers provide a small chart data model and receive an OPC byte
//! buffer that can be attached to a presentation by a package writer.

mod model;
mod package;

pub use model::{Chart, Kind, Series, Workbook};
pub use package::generate;

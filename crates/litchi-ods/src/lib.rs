//! OpenDocument Spreadsheet (`.ods`) support.
//!
//! The crate is organized by responsibility: immutable spreadsheet vocabulary
//! in [`model`], XML codecs in [`codec`], package access in [`package`],
//! construction in [`authoring`], and the concise entry points in [`facade`].

pub mod authoring;
pub mod codec;
pub mod drawing;
pub mod embedded;
pub mod facade;
pub mod media;
pub mod model;
pub mod package;
pub mod styles;
pub mod worksheet;

pub use drawing::{Frame, Part};
pub use embedded::{Kind, Object, Parameter, Root};
pub use facade::{Builder, MutableSpreadsheet, Spreadsheet};
pub use litchi_odf_common::rdf;
pub use media::Image;
pub use model::names;
pub use worksheet::{Cell, CellValue, CellView, Merge, Row, Sheet};

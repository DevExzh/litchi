//! Canonical layered owner for PresentationML presentation properties.

mod codec;
mod model;
mod package;

/// Typed document-level math defaults and their bounded XML codec.
pub mod math;

/// Presentation/document metadata capabilities that share the presentation
/// properties facade until the crate-level API integrator publishes their
/// final top-level paths.
pub mod metadata;

pub use model::*;
pub use package::{
    load_from_package, load_math_from_package, put_math_to_package, remove_math_from_package,
};

pub(crate) const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const P_STRICT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const R_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const P14_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
pub(crate) const P15_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
pub(crate) const A14_NS: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
pub(crate) const DISCARD_IMAGE_EDIT_DATA_URI: &str = "{E76CE94A-603C-4142-B9EB-6D1370010A27}";
pub(crate) const DEFAULT_IMAGE_DPI_URI: &str = "{D31A062A-798A-4329-ABDD-BBA856620510}";
pub(crate) const CHART_TRACKING_REF_BASED_URI: &str = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}";
pub(crate) const BROWSE_MODE_URI: &str = "{F99C55AA-B7CB-42B0-86F8-08522FDF87E8}";
pub(crate) const LASER_COLOR_URI: &str = "{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}";
pub(crate) const SHOW_MEDIA_CONTROLS_URI: &str = "{2FDB2607-1784-4EEB-B798-7EB5836EED8A}";
pub(crate) const MATH_URI: &str = "{4599F94E-CEE6-441E-89CC-EB005ECD8F06}";
pub(crate) const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
pub(crate) const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
pub(crate) const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";
pub(crate) const MAX_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_NODES: usize = 100_000;
pub(crate) const MAX_STRING: usize = 1024 * 1024;
pub(crate) const MAX_EXTENSIONS: usize = 1024;

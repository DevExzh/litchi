//! `PowerPoint` projections over format-neutral `OfficeArt` shapes.
//!
//! [MS-ODRAW] deliberately leaves `ClientData` and `ClientTextbox` payloads
//! to the host application. Keeping their interpretation here prevents the
//! shared drawing crate from acquiring a `PowerPoint` dependency while giving
//! PPT callers a concise, typed shape API.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{anchor, parse, parse_drawing, text_from_drawing, text_from_textbox};
pub use litchi_odraw::image::LimitKind as ImageLimitKind;
pub use litchi_odraw::shape::{Bounds, Flags, Kind, Native, Shape};
pub use litchi_odraw::{Container, Error, LimitKind, Limits, Record, RecordKind};
pub(crate) use model::ShapeLimitsExt;
pub use model::{Anchor, Drawing, FrameKind, Placeholder, ShapeExt};

pub(crate) use codec::{
    TextRecords, limits_for_source, powerpoint12_shape_metadata, shape_labels,
    text_from_slide_drawing, text_from_textbox_payload, text_from_textbox_payload_bounded, textbox,
    textbox_text_records, validate_host_framing,
};

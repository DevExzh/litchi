//! Layered image and OfficeArt drawing writer for legacy DOC files.
//!
//! The facade keeps the ergonomic public image inputs small while the nested
//! model, validation, and codec modules own semantic data, bounded sniffing,
//! and [MS-DOC]/[MS-ODRAW] wire layouts respectively.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use codec::{
    FIRST_PICTURE_SHAPE_ID, HEADER_FIRST_SHAPE_ID, build_dgg_info, build_plcf_spa,
    write_opt_record, write_picture_block, write_record_header,
};
pub use model::{FloatingPosition, Picture};
pub(crate) use model::{FloatingShapeContent, FloatingShapeInfo};

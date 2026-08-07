//! PPT record generation system.
//!
//! PowerPoint files use a record-based format where each record has:
//! - Version and instance (1 byte combined)
//! - Type (2 bytes)
//! - Length (4 bytes)
//! - Data (variable length)
//!
//! Based on Microsoft's "[MS-PPT]" specification and Apache POI's Record classes.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{
    create_docinfo_list_container, create_docinfo_list_container_minimal, create_document_atom,
    create_document_container, create_end_document, create_environment_minimal,
    create_main_master_container, create_slide_container, create_slide_list_with_text_master,
    create_slide_list_with_text_notes, create_slide_list_with_text_slides, create_text_atom,
    wrap_dg_into_ppdrawing, wrap_dgg_into_ppdrawing_group,
};
pub(crate) use codec::{
    create_docinfo_list_container_with_extensions, create_document_atom_with_font_embedding,
    create_environment_with_font_collection,
};
pub use model::{Error, RecordBuilder, RecordHeader, record_type};

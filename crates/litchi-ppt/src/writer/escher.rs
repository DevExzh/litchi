//! PPT-specific Escher (Office Drawing) writer facade.
//!
//! The facade keeps the historical `crate::writer::escher` paths stable while
//! layering PPT semantic inputs over the shared `litchi-odraw` wire substrate.
//!
//! - `model` owns PPT property vocabulary and typed shape data.
//! - `codec` owns PPT-specific record assembly.
//! - `tests` exercises the facade and its byte-level contracts.

/// Error type for PPT operations.
pub(crate) type Error = std::io::Error;

// Shared OfficeArt wire vocabulary. The record grammar itself remains owned by
// the format-neutral OfficeArt substrate. The aliases are PPT-internal facade
// names retained only for the other PPT writer modules.
pub(crate) use litchi_odraw::shape::Flags as ShapeFlags;
pub(crate) use litchi_odraw::write::{
    COMPLEX as PROPERTY_FLAG_COMPLEX, Property as EscherProperty, Sp as EscherSpData, record_type,
    shape_type,
};

/// PPT-specific record types embedded in Escher.
pub(crate) mod ppt_record_type {
    /// OEPlaceholderAtom.
    pub(crate) const OE_PLACEHOLDER_ATOM: u16 = 0x0BC3;
}

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use model::FreeformGeometry;

pub(crate) use model::{
    BG_SHAPE_PROPERTIES, ChildAnchor, DGG_DEFAULT_PROPERTIES, EscherDgData, EscherDggHeader,
    EscherHeader, EscherSpgrData, FileIdCluster, SplitMenuColors, UserShapeData, header_version,
    ppt_prop_value, prop_id,
};

pub(crate) use codec::{
    EscherBuilder, build_client_textbox, create_dg_container_with_charts,
    create_dg_container_with_shapes, create_dgg_container, create_dgg_container_with_blips,
};

#[cfg(test)]
pub(crate) use codec::{
    build_client_data_with_hyperlink, build_client_data_with_placeholder,
    build_client_textbox_formatted, build_client_textbox_with_interactions, build_shape_properties,
    create_dg_container_with_tables, create_user_shape_container,
};

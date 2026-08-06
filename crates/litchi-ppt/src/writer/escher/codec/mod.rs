//! Layered PPT OfficeArt/Escher encoding and inspection.
//!
//! Each record family owns one responsibility:
//!
//! - [`wire`] contains the shared record builder and the zero-copy inspection
//!   view;
//! - [`drawing`] assembles drawing-group and drawing containers;
//! - [`shapes`] emits shape containers and host anchors;
//! - [`group`] emits the bounded nested-group grammar from [MS-ODRAW];
//! - [`properties`] owns the PPT shape property vocabulary;
//! - [`client_data`] and [`text`] own host-specific child records;
//! - [`validation`] keeps topology and resource checks out of encoders.

mod client_data;
mod drawing;
mod group;
mod properties;
mod shapes;
mod text;
mod validation;
pub(crate) mod wire;

#[cfg(test)]
pub(crate) use client_data::build_client_data_with_hyperlink;
#[cfg(test)]
pub(crate) use client_data::build_client_data_with_placeholder;
#[cfg(test)]
pub(crate) use drawing::create_dg_container_with_tables;
pub(crate) use drawing::{
    create_dg_container_with_charts, create_dg_container_with_shapes, create_dgg_container,
    create_dgg_container_with_blips,
};
#[cfg(test)]
pub(crate) use group::{create_dg_container_with_group, create_group_shape_container};
#[cfg(test)]
pub(crate) use properties::build_shape_properties;
#[cfg(test)]
pub(crate) use shapes::create_user_shape_container;
pub(crate) use text::build_client_textbox;
#[cfg(test)]
pub(crate) use text::{build_client_textbox_formatted, build_client_textbox_with_interactions};
pub(crate) use wire::EscherBuilder;

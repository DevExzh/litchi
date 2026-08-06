//! Bounded XML conversion for rich-value and feature-property-bag parts.

mod parser;
mod writer;
mod xml;

pub use parser::{
    parse_arrays, parse_data, parse_dxf_complement, parse_feature_property_bags,
    parse_rich_value_rels, parse_structures, parse_xf_complement,
};
pub use writer::{
    write_arrays, write_data, write_dxf_complement, write_feature_property_bags,
    write_rich_value_rels, write_structures, write_xf_complement,
};

pub(crate) use parser::parse_part;
pub(crate) use writer::write_part;
pub(crate) use xml::validate_fragment;

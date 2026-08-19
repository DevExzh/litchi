//! Transactional ODS construction.

mod builder;
mod mutable;

pub use builder::Builder;
pub use mutable::MutableSpreadsheet;

pub(crate) use builder::{
    ValidateHandler, is_office_namespace, validate_content_xml, validate_size,
};

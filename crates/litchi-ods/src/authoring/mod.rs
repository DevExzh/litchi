//! Transactional ODS construction.

mod builder;

pub use builder::Builder;

pub(crate) use builder::{
    ValidateHandler, is_office_namespace, validate_content_xml, validate_size,
};

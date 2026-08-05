//! ODT facade for shared master-page region semantics.

pub use litchi_odf_common::style::master::{Child, ChildKind, Kind, Master, Region};

pub(crate) use litchi_odf_common::style::master::reader::read;
pub(crate) use litchi_odf_common::style::master::writer::{replace_range, set_text, set_xml};

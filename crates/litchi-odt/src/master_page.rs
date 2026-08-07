//! ODT facade for shared master-page authoring and package access.

pub use litchi_odf_common::style::master::writer::{add, insert, remove, replace};
pub use litchi_odf_common::style::master::{Child, ChildKind, Master};

use litchi_core::Result;

impl crate::Package {
    /// Parse inert master-page metadata from packaged `styles.xml`.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        self.styles_xml()?.map_or_else(
            || Ok(Vec::new()),
            |xml| litchi_odf_common::style::master::reader::read(&xml),
        )
    }
}

impl crate::FlatDocument {
    /// Parse inert master-page metadata from a flat `OpenDocument`.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        litchi_odf_common::style::master::reader::read(self.xml())
    }
}

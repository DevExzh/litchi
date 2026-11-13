//! Contextual access to retained `styles.xml` metadata.

use super::super::model::MutableDocument;
use crate::master_page::Master;
use crate::page_layout::PageLayout;
use litchi_core::Result;

/// Read-only view of the retained styles layer.
pub(crate) struct Styles<'document> {
    pub(super) document: &'document MutableDocument,
}

/// Mutable view of targeted styles-layer edits.
pub(crate) struct StylesMut<'document> {
    pub(super) document: &'document mut MutableDocument,
}

impl Styles<'_> {
    /// Parse master pages and their authored header/footer regions.
    pub(crate) fn master_pages(&self) -> Result<Vec<Master>> {
        self.document.master_pages()
    }

    /// Parse automatic page layouts and their authored properties.
    pub(crate) fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.document.page_layouts()
    }
}

impl StylesMut<'_> {
    /// Add an empty master page and its referenced page layout.
    pub(crate) fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        self.document.add_master_page(name, page_layout_name)
    }
}

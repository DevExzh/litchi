//! Contextual access to retained `styles.xml` metadata.

use super::super::model::MutableDocument;
use crate::header_footer::Kind;
use crate::master_page::Master;
use crate::page_layout::PageLayout;
use litchi_core::Result;

/// Read-only view of the retained styles layer.
pub struct Styles<'document> {
    pub(super) document: &'document MutableDocument,
}

/// Mutable view of targeted styles-layer edits.
pub struct StylesMut<'document> {
    pub(super) document: &'document mut MutableDocument,
}

impl Styles<'_> {
    /// Parse master pages and their authored header/footer regions.
    pub fn master_pages(&self) -> Result<Vec<Master>> {
        self.document.master_pages()
    }

    /// Parse automatic page layouts and their authored properties.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.document.page_layouts()
    }

    /// Parse named ruby styles from `styles.xml`.
    pub fn ruby_styles(&self) -> Result<crate::ruby_family::Styles> {
        self.document.ruby_styles()
    }

    /// Parse note presentation configurations from `styles.xml`.
    pub fn notes_configurations(&self) -> Result<crate::notes_configuration::Configurations> {
        self.document.notes_configurations()
    }

    /// Parse outline numbering styles from `styles.xml`.
    pub fn outline_styles(&self) -> Result<crate::outline_style::Styles> {
        self.document.outline_styles()
    }
}

impl StylesMut<'_> {
    /// Reborrow this editor as a read-only styles view.
    pub fn read(&self) -> Styles<'_> {
        Styles {
            document: self.document,
        }
    }

    /// Add an empty master page and its referenced page layout.
    pub fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        self.document.add_master_page(name, page_layout_name)
    }

    /// Replace one page layout with a complete, validated XML fragment.
    pub fn set_page_layout_xml(
        &mut self,
        page_layout_name: &str,
        page_layout_xml: &str,
    ) -> Result<()> {
        self.document
            .set_page_layout_xml(page_layout_name, page_layout_xml)
    }

    /// Set plain text in one master-page header or footer.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: &str,
        kind: Kind,
        text: &str,
    ) -> Result<()> {
        self.document
            .set_header_footer_text(master_page_name, kind, text)
    }

    /// Replace one master-page header or footer with a complete XML fragment.
    pub fn set_header_footer_xml(
        &mut self,
        master_page_name: &str,
        kind: Kind,
        xml: &str,
    ) -> Result<()> {
        self.document
            .set_header_footer_xml(master_page_name, kind, xml)
    }

    /// Remove one master-page header or footer region.
    pub fn clear_header_footer(&mut self, master_page_name: &str, kind: Kind) -> Result<()> {
        self.document.clear_header_footer(master_page_name, kind)
    }

    /// Insert or replace a named outline style.
    pub fn set_outline_style(
        &mut self,
        style: &crate::outline_style::Style,
    ) -> Result<Option<crate::outline_style::Style>> {
        self.document.set_outline_style(style)
    }

    /// Insert or replace a named ruby style.
    pub fn set_ruby_style(
        &mut self,
        style: &crate::ruby_family::Style,
    ) -> Result<Option<crate::ruby_family::Style>> {
        self.document.set_ruby_style(style)
    }
}

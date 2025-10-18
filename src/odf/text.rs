//! OpenDocument Text (.odt) support.
//!
//! This module provides a unified API for working with OpenDocument text documents,
//! equivalent to Microsoft Word documents.

use crate::common::{Error, Result, Metadata};
use crate::odf::core::{Content, Meta, Package, Styles, Manifest};
use crate::odf::elements::text::{TextElements, Paragraph as ElementParagraph};
use crate::odf::elements::table::Table as ElementTable;
use crate::odf::elements::style::{StyleRegistry, StyleElements};
use std::io::Cursor;
use std::path::Path;

/// An OpenDocument text document (.odt)
#[allow(dead_code)]
pub struct Document {
    package: Package<Cursor<Vec<u8>>>,
    content: Content,
    styles: Option<Styles>,
    meta: Option<Meta>,
    manifest: Manifest,
    style_registry: StyleRegistry,
}

impl Document {
    /// Open an ODT document from a file path
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();

        // Read the entire file into memory
        let bytes = std::fs::read(path)?;
        Self::from_bytes(bytes)
    }

    /// Create a Document from a byte buffer
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut package = Package::from_reader(cursor)?;

        // Verify this is a text document
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.text") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODT file: MIME type is {}", mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        let manifest = package.manifest().clone();

        // Initialize style registry
        let mut style_registry = StyleRegistry::new();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles
            && let Ok(registry) = StyleElements::parse_styles(styles_part.xml_content()) {
                style_registry = registry;
            }

        // Also parse styles from content.xml (automatic styles)
        if let Ok(content_registry) = StyleElements::parse_styles(content.xml_content()) {
            // Merge content styles into main registry (content styles take precedence)
            for (_name, style) in content_registry.styles {
                style_registry.add_style(style);
            }
        }

        Ok(Self {
            package,
            content,
            styles,
            meta,
            manifest,
            style_registry,
        })
    }

    /// Extract all text content from the document
    pub fn text(&mut self) -> Result<String> {
        TextElements::extract_text(self.content.xml_content())
    }

    /// Get all paragraphs in the document
    pub fn paragraphs(&mut self) -> Result<Vec<ElementParagraph>> {
        TextElements::parse_paragraphs(self.content.xml_content())
    }

    /// Get all tables in the document
    pub fn tables(&mut self) -> Result<Vec<ElementTable>> {
        use crate::odf::elements::table::TableElements;
        TableElements::parse_tables_from_content(self.content.xml_content())
    }

    /// Get document metadata
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }

    /// Get the style registry
    pub fn styles(&self) -> &StyleRegistry {
        &self.style_registry
    }

    /// Get resolved style properties for a given style name
    pub fn get_style_properties(&self, style_name: &str) -> crate::odf::elements::style::StyleProperties {
        self.style_registry.get_resolved_properties(style_name)
    }
}


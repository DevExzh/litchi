//! Typed semantic values for WordprocessingML header and footer stories.

use crate::error::Result;
use crate::namespace::scan_word_element_ranges;
use crate::paragraph::{Paragraph, extract_word_text};
use crate::table::Table;
use crate::writer::Watermark;
use litchi_opc::part::Part;
use std::sync::{Arc, OnceLock};

use super::codec;

/// Maximum accepted header or footer XML size.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// Maximum XML nesting depth accepted by the story boundary.
pub const MAX_XML_DEPTH: usize = 128;
/// Maximum XML element count accepted by the story boundary.
pub const MAX_XML_NODES: usize = 1_000_000;

/// The WordprocessingML story root carried by a package part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Role {
    /// A `<w:hdr>` story.
    Header,
    /// A `<w:ftr>` story.
    Footer,
}

impl Role {
    /// The local XML root name for this story role.
    #[inline]
    pub const fn root(self) -> &'static str {
        match self {
            Self::Header => "hdr",
            Self::Footer => "ftr",
        }
    }
}

/// Header or footer definition kind within a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Kind {
    /// Header/footer for odd pages or all pages if no even definition exists.
    Primary = 1,
    /// Header/footer for the first page of a section.
    FirstPage = 2,
    /// Header/footer for even pages of a recto/verso section.
    EvenPage = 3,
}

impl Kind {
    /// Convert the kind to its WordprocessingML `w:type` value.
    #[inline]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Primary => "default",
            Self::FirstPage => "first",
            Self::EvenPage => "even",
        }
    }

    /// Parse a WordprocessingML `w:type` value.
    #[inline]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "default" => Some(Self::Primary),
            "first" => Some(Self::FirstPage),
            "even" => Some(Self::EvenPage),
            _ => None,
        }
    }
}

impl Default for Kind {
    #[inline]
    fn default() -> Self {
        Self::Primary
    }
}

impl std::fmt::Display for Kind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Primary => write!(formatter, "Primary"),
            Self::FirstPage => write!(formatter, "First Page"),
            Self::EvenPage => write!(formatter, "Even Page"),
        }
    }
}

/// A lossless, immutable view of one WordprocessingML header or footer story.
//
// The raw package bytes are shared with the OPC part. Markup-compatibility
// processing is lazy and cached, so repeated semantic queries do not copy or
// reprocess an unchanged story.
#[derive(Debug, Clone)]
pub struct Story {
    xml_bytes: Arc<Vec<u8>>,
    semantic: Arc<OnceLock<Arc<Vec<u8>>>>,
    role: Role,
    kind: Kind,
}

impl Story {
    /// Construct a story from validated XML bytes.
    pub fn from_xml_bytes(xml_bytes: Vec<u8>, kind: Kind) -> Result<Self> {
        let role = codec::validate(&xml_bytes)?;
        Ok(Self {
            xml_bytes: Arc::new(xml_bytes),
            semantic: Arc::new(OnceLock::new()),
            role,
            kind,
        })
    }

    /// Construct a story by sharing the XML allocation owned by an OPC part.
    pub(crate) fn from_part(part: &dyn Part, kind: Kind) -> Result<Self> {
        let role = codec::validate(part.blob())?;
        Ok(Self {
            xml_bytes: part.blob_arc(),
            semantic: Arc::new(OnceLock::new()),
            role,
            kind,
        })
    }

    /// Return whether this story is a header or footer.
    #[inline]
    pub const fn role(&self) -> Role {
        self.role
    }

    /// Return the section page kind associated with this story reference.
    #[inline]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the original XML bytes without MCE rewriting.
    #[inline]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml_bytes.as_slice()
    }

    /// Extract all text content from this story.
    pub fn text(&self) -> Result<String> {
        let xml = self.semantic_xml()?;
        extract_word_text(xml.as_slice())
    }

    /// Return all paragraphs in authored order.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let xml = self.semantic_xml()?;
        let mut paragraphs = Vec::new();
        scan_word_element_ranges(xml.as_slice(), &[b"p".as_slice()], |_, start, length| {
            paragraphs.push(Paragraph::from_arc_range(Arc::clone(&xml), start, length));
            Ok(())
        })?;
        Ok(paragraphs)
    }

    /// Return all tables in authored order.
    pub fn tables(&self) -> Result<Vec<Table>> {
        let xml = self.semantic_xml()?;
        let mut tables = Vec::new();
        scan_word_element_ranges(xml.as_slice(), &[b"tbl".as_slice()], |_, start, length| {
            tables.push(Table::from_arc_range(Arc::clone(&xml), start, length));
            Ok(())
        })?;
        Ok(tables)
    }

    /// Count paragraphs in this story.
    pub fn paragraph_count(&self) -> Result<usize> {
        let xml = self.semantic_xml()?;
        let mut count = 0;
        scan_word_element_ranges(xml.as_slice(), &[b"p".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Count tables in this story.
    pub fn table_count(&self) -> Result<usize> {
        let xml = self.semantic_xml()?;
        let mut count = 0;
        scan_word_element_ranges(xml.as_slice(), &[b"tbl".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Return standard VML text watermarks embedded in this story.
    pub fn watermarks(&self) -> Result<Vec<Watermark>> {
        let xml = self.semantic_xml()?;
        Watermark::from_header_xml(xml.as_slice())
    }

    /// Return picture watermark anchors embedded in this story.
    pub fn image_watermarks(&self) -> Result<Vec<crate::writer::ImageWatermarkAnchor>> {
        let xml = self.semantic_xml()?;
        crate::writer::ImageWatermarkAnchor::from_header_xml(xml.as_slice())
    }

    fn semantic_xml(&self) -> Result<Arc<Vec<u8>>> {
        if let Some(xml) = self.semantic.get() {
            return Ok(Arc::clone(xml));
        }
        let xml = codec::semantic_xml(&self.xml_bytes)?;
        let _ = self.semantic.set(Arc::clone(&xml));
        Ok(xml)
    }
}

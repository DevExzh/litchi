//! Web-template package authoring.

use litchi_core::{Error, Result};
use litchi_odf_common::{compact_xml, core::PackageWriter};

const CONTENT_PREFIX: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:text>"#;
const CONTENT_SUFFIX: &str = "</office:text></office:body></office:document-content>";

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    blocks: Vec<AuthoredBlock>,
    content_xml: Option<String>,
    meta_xml: Option<String>,
    settings_xml: Option<String>,
    styles_xml: Option<String>,
}

impl Builder {
    /// Creates an empty ODF 1.4 web-template builder.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            blocks: Vec::new(),
            content_xml: None,
            meta_xml: None,
            settings_xml: None,
            styles_xml: None,
        }
    }

    /// Adds a typed heading to freshly authored content.
    #[must_use]
    pub fn heading(mut self, heading: crate::heading::Heading) -> Self {
        self.blocks.push(AuthoredBlock::Heading(heading));
        self
    }

    /// Adds a compact `meta.xml` payload.
    #[must_use]
    pub fn meta_xml(mut self, xml: impl Into<String>) -> Self {
        self.meta_xml = Some(xml.into());
        self
    }

    /// Adds a typed paragraph to freshly authored content.
    #[must_use]
    pub fn paragraph(mut self, paragraph: crate::paragraph::Paragraph) -> Self {
        self.blocks.push(AuthoredBlock::Paragraph(paragraph));
        self
    }

    /// Adds a compact `settings.xml` payload.
    #[must_use]
    pub fn settings_xml(mut self, xml: impl Into<String>) -> Self {
        self.settings_xml = Some(xml.into());
        self
    }

    /// Adds a compact `styles.xml` payload.
    #[must_use]
    pub fn styles_xml(mut self, xml: impl Into<String>) -> Self {
        self.styles_xml = Some(xml.into());
        self
    }

    /// Replaces the `content.xml` payload.
    ///
    /// Raw content and typed blocks are mutually exclusive so call order
    /// cannot silently discard either representation.
    #[must_use]
    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = Some(xml.into());
        self
    }

    /// Validates, packages, reopens, and publishes the document bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation, package writing, or semantic
    /// readback fails, or raw content and typed blocks were both supplied.
    pub fn build(self) -> Result<Vec<u8>> {
        if self.content_xml.is_some() && !self.blocks.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH builder cannot combine raw content_xml with typed blocks".to_string(),
            ));
        }
        let content_xml = match self.content_xml {
            Some(xml) => xml,
            None => render_blocks(&self.blocks)?,
        };
        crate::codec::validate_authored(&content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        for (path, optional_xml) in [
            ("meta.xml", self.meta_xml),
            ("settings.xml", self.settings_xml),
            ("styles.xml", self.styles_xml),
        ] {
            if let Some(part_xml) = optional_xml {
                compact_xml::validate(part_xml.as_bytes()).map_err(Error::from)?;
                writer.add_file(path, part_xml.as_bytes())?;
            }
        }
        Ok(crate::package::Snapshot::from_bytes(writer.finish_to_bytes()?)?.into_bytes())
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Clone, Debug)]
enum AuthoredBlock {
    Heading(crate::heading::Heading),
    Paragraph(crate::paragraph::Paragraph),
}

fn render_blocks(blocks: &[AuthoredBlock]) -> Result<String> {
    let mut capacity = CONTENT_PREFIX.len() + CONTENT_SUFFIX.len();
    for block in blocks {
        let (markup_bytes, text_bytes) = match block {
            AuthoredBlock::Heading(heading) => (
                48_usize.saturating_add(heading.style_name().map_or(0, str::len)),
                heading.text().len(),
            ),
            AuthoredBlock::Paragraph(paragraph) => (
                32_usize.saturating_add(paragraph.style_name().map_or(0, str::len)),
                paragraph.text().len(),
            ),
        };
        capacity = capacity
            .checked_add(markup_bytes)
            .and_then(|size| size.checked_add(text_bytes.saturating_mul(6)))
            .ok_or_else(|| {
                Error::InvalidFormat("OTH authored content size overflow".to_string())
            })?;
    }
    if capacity > compact_xml::DEFAULT_MAX_BYTES {
        return Err(Error::InvalidFormat(
            "OTH authored content exceeds the output byte limit".to_string(),
        ));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "OTH authored content",
            source,
        })?;
    output.push_str(CONTENT_PREFIX);
    for block in blocks {
        match block {
            AuthoredBlock::Heading(heading) => {
                output.push_str("<text:h text:outline-level=\"");
                output.push_str(&heading.level().to_string());
                output.push('"');
                push_style(&mut output, heading.style_name());
                output.push('>');
                output.push_str(&quick_xml::escape::escape(heading.text()));
                output.push_str("</text:h>");
            },
            AuthoredBlock::Paragraph(paragraph) => {
                output.push_str("<text:p");
                push_style(&mut output, paragraph.style_name());
                output.push('>');
                output.push_str(&quick_xml::escape::escape(paragraph.text()));
                output.push_str("</text:p>");
            },
        }
    }
    output.push_str(CONTENT_SUFFIX);
    Ok(output)
}

fn push_style(output: &mut String, optional_style_name: Option<&str>) {
    if let Some(style_name) = optional_style_name {
        output.push_str(" text:style-name=\"");
        output.push_str(&quick_xml::escape::escape(style_name));
        output.push('"');
    }
}

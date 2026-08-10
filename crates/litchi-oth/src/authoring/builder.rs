//! Web-template package authoring.

use litchi_core::{Error, Result};
use litchi_odf_common::{compact_xml, core::PackageWriter};

const CONTENT_PREFIX: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:text>"#;
const CONTENT_SUFFIX: &str = "</office:text></office:body></office:document-content>";

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    blocks: Vec<crate::ContentBlock>,
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
        self.blocks.push(crate::ContentBlock::Heading(heading));
        self
    }

    /// Adds a typed flat list to freshly authored content.
    #[must_use]
    pub fn list(mut self, list: crate::list::List) -> Self {
        self.blocks.push(crate::ContentBlock::List(list));
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
        self.blocks.push(crate::ContentBlock::Paragraph(paragraph));
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

fn render_blocks(blocks: &[crate::ContentBlock]) -> Result<String> {
    let mut capacity = CONTENT_PREFIX.len() + CONTENT_SUFFIX.len();
    for block in blocks {
        let (markup_bytes, text_bytes) = match block {
            crate::ContentBlock::Heading(heading) => (
                48_usize.saturating_add(heading.style_name().map_or(0, str::len)),
                heading.text().len(),
            ),
            crate::ContentBlock::Paragraph(paragraph) => (
                32_usize.saturating_add(paragraph.style_name().map_or(0, str::len)),
                paragraph.text().len(),
            ),
            crate::ContentBlock::List(list) => {
                let text_bytes = list
                    .items()
                    .iter()
                    .flat_map(crate::list::Item::paragraphs)
                    .map(crate::paragraph::Paragraph::text)
                    .map(str::len)
                    .fold(0_usize, usize::saturating_add);
                (128_usize.saturating_mul(list.items().len()), text_bytes)
            },
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
            crate::ContentBlock::Heading(heading) => {
                output.push_str("<text:h text:outline-level=\"");
                output.push_str(&heading.level().to_string());
                output.push('"');
                push_style(&mut output, heading.style_name());
                output.push('>');
                output.push_str(&quick_xml::escape::escape(heading.text()));
                output.push_str("</text:h>");
            },
            crate::ContentBlock::Paragraph(paragraph) => {
                output.push_str("<text:p");
                push_style(&mut output, paragraph.style_name());
                output.push('>');
                output.push_str(&quick_xml::escape::escape(paragraph.text()));
                output.push_str("</text:p>");
            },
            crate::ContentBlock::List(list) => render_list(&mut output, list),
        }
    }
    output.push_str(CONTENT_SUFFIX);
    Ok(output)
}

pub(crate) fn render_fragment(blocks: &[crate::ContentBlock]) -> Result<String> {
    let document = render_blocks(blocks)?;
    document
        .strip_prefix(CONTENT_PREFIX)
        .and_then(|value| value.strip_suffix(CONTENT_SUFFIX))
        .map(str::to_owned)
        .ok_or_else(|| {
            Error::InvalidFormat("OTH rendered fragment envelope is invalid".to_string())
        })
}

pub(crate) fn render_inline(items: &[crate::inline::Content]) -> Result<(String, String)> {
    let mut xml = String::new();
    let mut text = String::new();
    for item in items {
        match item {
            crate::inline::Content::Text(value) => {
                xml.push_str(&quick_xml::escape::escape(value));
                text.push_str(value);
            },
            crate::inline::Content::Link(link) => {
                xml.push_str("<text:a xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xlink:href=\"");
                xml.push_str(&quick_xml::escape::escape(link.href()));
                xml.push_str("\">");
                xml.push_str(&quick_xml::escape::escape(link.label()));
                xml.push_str("</text:a>");
                text.push_str(link.label());
            },
            crate::inline::Content::Span(span) => {
                xml.push_str("<text:span xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" text:style-name=\"");
                xml.push_str(&quick_xml::escape::escape(span.style_name()));
                xml.push_str("\">");
                xml.push_str(&quick_xml::escape::escape(span.text()));
                xml.push_str("</text:span>");
                text.push_str(span.text());
            },
            crate::inline::Content::Field(field) => {
                let local = field_local(field.kind())?;
                xml.push_str("<text:");
                xml.push_str(local);
                xml.push_str(" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"");
                if let Some(name) = field.name() {
                    xml.push_str(" text:name=\"");
                    xml.push_str(&quick_xml::escape::escape(name));
                    xml.push('"');
                }
                if let Some(value) = field.value() {
                    xml.push_str(" office:string-value=\"");
                    xml.push_str(&quick_xml::escape::escape(value));
                    xml.push('"');
                }
                if field.is_fixed() {
                    xml.push_str(" text:fixed=\"true\"");
                }
                xml.push('>');
                xml.push_str(&quick_xml::escape::escape(field.display()));
                xml.push_str("</text:");
                xml.push_str(local);
                xml.push('>');
                text.push_str(field.display());
            },
            crate::inline::Content::BookmarkPoint(name) => {
                push_bookmark(&mut xml, "bookmark", name);
            },
            crate::inline::Content::BookmarkRangeStart(name) => {
                push_bookmark(&mut xml, "bookmark-start", name);
            },
            crate::inline::Content::BookmarkRangeEnd(name) => {
                push_bookmark(&mut xml, "bookmark-end", name);
            },
        }
    }
    Ok((xml, text))
}

pub(crate) fn render_forms(forms: &[crate::form::Form]) -> Result<String> {
    if forms.is_empty() {
        return Ok(String::new());
    }
    let mut xml = String::from(
        "<office:forms xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\">",
    );
    for form in forms {
        xml.push_str("<form:form");
        if let Some(name) = form.name() {
            xml.push_str(" form:name=\"");
            xml.push_str(&quick_xml::escape::escape(name));
            xml.push('"');
        }
        xml.push('>');
        for control in form.controls() {
            if control.kind().is_empty()
                || !control
                    .kind()
                    .bytes()
                    .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_'))
            {
                return Err(Error::InvalidFormat(
                    "OTH form control kind is not a safe XML local name".to_string(),
                ));
            }
            xml.push_str("<form:");
            xml.push_str(control.kind());
            if let Some(id) = control.id() {
                xml.push_str(" form:id=\"");
                xml.push_str(&quick_xml::escape::escape(id));
                xml.push('"');
            }
            if let Some(name) = control.name() {
                xml.push_str(" form:name=\"");
                xml.push_str(&quick_xml::escape::escape(name));
                xml.push('"');
            }
            xml.push_str("/>");
        }
        xml.push_str("</form:form>");
    }
    xml.push_str("</office:forms>");
    Ok(xml)
}

pub(crate) fn render_resource(resource: &crate::resource::Resource) -> String {
    let local = match resource.kind() {
        crate::resource::Kind::Image => "image",
        crate::resource::Kind::Object => "object",
        crate::resource::Kind::OleObject => "object-ole",
        crate::resource::Kind::Plugin => "plugin",
        crate::resource::Kind::FloatingFrame => "floating-frame",
    };
    format!(
        "<draw:{local} xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xlink:href=\"{}\"/>",
        quick_xml::escape::escape(resource.href())
    )
}

fn push_bookmark(output: &mut String, local: &str, name: &str) {
    output.push_str("<text:");
    output.push_str(local);
    output.push_str(" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"");
    output.push_str(" text:name=\"");
    output.push_str(&quick_xml::escape::escape(name));
    output.push_str("\"/>");
}

fn field_local(kind: &crate::field::Kind) -> Result<&str> {
    match kind {
        crate::field::Kind::Date => Ok("date"),
        crate::field::Kind::Time => Ok("time"),
        crate::field::Kind::PageNumber => Ok("page-number"),
        crate::field::Kind::PageCount => Ok("page-count"),
        crate::field::Kind::Title => Ok("title"),
        crate::field::Kind::Subject => Ok("subject"),
        crate::field::Kind::AuthorName => Ok("author-name"),
        crate::field::Kind::User => Ok("user-field-get"),
        crate::field::Kind::Variable => Ok("variable-get"),
        crate::field::Kind::Other(_) => Err(Error::InvalidFormat(
            "OTH authored inline field kind is unsupported".to_string(),
        )),
    }
}

fn render_list(output: &mut String, list: &crate::list::List) {
    output.push_str("<text:list");
    push_style(output, list.style_name());
    output.push('>');
    for item in list.items() {
        output.push_str("<text:list-item");
        if let Some(start_value) = item.start_value() {
            output.push_str(" text:start-value=\"");
            output.push_str(&start_value.to_string());
            output.push('"');
        }
        output.push('>');
        for paragraph in item.paragraphs() {
            output.push_str("<text:p");
            push_style(output, paragraph.style_name());
            output.push('>');
            output.push_str(&quick_xml::escape::escape(paragraph.text()));
            output.push_str("</text:p>");
        }
        for nested in item.nested_lists() {
            render_list(output, nested);
        }
        output.push_str("</text:list-item>");
    }
    output.push_str("</text:list>");
}

fn push_style(output: &mut String, optional_style_name: Option<&str>) {
    if let Some(style_name) = optional_style_name {
        output.push_str(" text:style-name=\"");
        output.push_str(&quick_xml::escape::escape(style_name));
        output.push('"');
    }
}

pub(crate) fn render_styles(styles: &[crate::style::Style]) -> Result<String> {
    let mut output = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"><office:styles>",
    );
    for style in styles {
        output.push_str("<style:style style:name=\"");
        output.push_str(&quick_xml::escape::escape(style.name()));
        output.push('"');
        if let Some(family) = style.family() {
            output.push_str(" style:family=\"");
            output.push_str(&quick_xml::escape::escape(family));
            output.push('"');
        }
        if let Some(parent) = style.parent_name() {
            output.push_str(" style:parent-style-name=\"");
            output.push_str(&quick_xml::escape::escape(parent));
            output.push('"');
        }
        if let Some(properties) = style.text_properties() {
            output.push_str("><style:text-properties");
            if let Some(color) = properties.color() {
                output.push_str(" fo:color=\"");
                output.push_str(color);
                output.push('"');
            }
            if let Some(color) = properties.background_color() {
                output.push_str(" fo:background-color=\"");
                output.push_str(color);
                output.push('"');
            }
            if let Some(weight) = properties.weight() {
                output.push_str(" fo:font-weight=\"");
                output.push_str(match weight {
                    crate::style::Weight::Normal => "normal",
                    crate::style::Weight::Bold => "bold",
                });
                output.push('"');
            }
            if let Some(slant) = properties.slant() {
                output.push_str(" fo:font-style=\"");
                output.push_str(match slant {
                    crate::style::Slant::Normal => "normal",
                    crate::style::Slant::Italic => "italic",
                });
                output.push('"');
            }
            output.push_str("/></style:style>");
        } else {
            output.push_str("/>");
        }
    }
    output.push_str("</office:styles></office:document-styles>");
    compact_xml::validate(output.as_bytes())?;
    Ok(output)
}

pub(crate) fn render_metadata(metadata: &litchi_core::Metadata) -> Result<String> {
    let mut output = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\"><office:meta>",
    );
    for (tag, optional_text) in [
        ("dc:title", metadata.title.as_deref()),
        ("dc:subject", metadata.subject.as_deref()),
        ("dc:creator", metadata.author.as_deref()),
        ("dc:description", metadata.description.as_deref()),
        ("dc:identifier", metadata.identifier.as_deref()),
        ("dc:language", metadata.language.as_deref()),
        ("meta:keyword", metadata.keywords.as_deref()),
        ("meta:generator", metadata.application.as_deref()),
        ("meta:category", metadata.category.as_deref()),
    ] {
        if let Some(text) = optional_text {
            output.push('<');
            output.push_str(tag);
            output.push('>');
            output.push_str(&quick_xml::escape::escape(text));
            output.push_str("</");
            output.push_str(tag);
            output.push('>');
        }
    }
    if let Some(created) = metadata.created {
        output.push_str("<meta:creation-date>");
        output.push_str(&created.to_rfc3339());
        output.push_str("</meta:creation-date>");
    }
    if let Some(modified) = metadata.modified {
        output.push_str("<dc:date>");
        output.push_str(&modified.to_rfc3339());
        output.push_str("</dc:date>");
    }
    if metadata.page_count.is_some()
        || metadata.word_count.is_some()
        || metadata.character_count.is_some()
    {
        output.push_str("<meta:document-statistic");
        for (name, optional_count) in [
            ("page-count", metadata.page_count),
            ("word-count", metadata.word_count),
            ("character-count", metadata.character_count),
        ] {
            if let Some(count) = optional_count {
                output.push_str(" meta:");
                output.push_str(name);
                output.push_str("=\"");
                output.push_str(&count.to_string());
                output.push('"');
            }
        }
        output.push_str("/>");
    }
    output.push_str("</office:meta></office:document-meta>");
    compact_xml::validate(output.as_bytes())?;
    Ok(output)
}

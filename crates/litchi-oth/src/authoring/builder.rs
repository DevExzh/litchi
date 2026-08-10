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
    forms: Vec<crate::form::Form>,
    heading_count: usize,
    inline_changes: Vec<(crate::facade::InlineBlock, Vec<crate::inline::Content>)>,
    meta_xml: Option<String>,
    metadata: Option<litchi_core::Metadata>,
    paragraph_count: usize,
    resources: Vec<FreshResource>,
    settings_xml: Option<String>,
    styles: Vec<crate::style::Style>,
    styles_xml: Option<String>,
}

#[derive(Clone, Debug)]
struct FreshResource {
    members: Vec<ResourceMember>,
    reference: crate::resource::Resource,
}

/// One checked package member below a freshly authored embedded object.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceMember {
    bytes: Vec<u8>,
    media_type: String,
    path: String,
}

impl ResourceMember {
    /// Creates one object-relative package member.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe path or empty media type.
    pub fn new(
        path_input: impl Into<String>,
        media_type_input: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<Self> {
        let path_value = path_input.into();
        validate_relative_path(&path_value)?;
        let media_type_value = media_type_input.into();
        if media_type_value.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH fresh resource media type cannot be empty".to_string(),
            ));
        }
        Ok(Self {
            bytes,
            media_type: media_type_value,
            path: path_value,
        })
    }

    /// Object-relative member path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Manifest media type.
    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Exact member bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

impl Builder {
    /// Creates an empty ODF 1.4 web-template builder.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            blocks: Vec::new(),
            content_xml: None,
            forms: Vec::new(),
            heading_count: 0,
            inline_changes: Vec::new(),
            meta_xml: None,
            metadata: None,
            paragraph_count: 0,
            resources: Vec::new(),
            settings_xml: None,
            styles: Vec::new(),
            styles_xml: None,
        }
    }

    /// Adds a typed heading to freshly authored content.
    #[must_use]
    pub fn heading(mut self, heading: crate::heading::Heading) -> Self {
        self.blocks.push(crate::ContentBlock::Heading(heading));
        self.heading_count = self.heading_count.saturating_add(1);
        self
    }

    /// Adds a typed nested list tree to freshly authored content.
    #[must_use]
    pub fn list(mut self, list: crate::list::List) -> Self {
        self.paragraph_count = self
            .paragraph_count
            .saturating_add(list_paragraph_count(&list));
        self.blocks.push(crate::ContentBlock::List(list));
        self
    }

    /// Adds one inert form and its controls.
    #[must_use]
    pub fn form(mut self, form: crate::form::Form) -> Self {
        self.forms.push(form);
        self
    }

    /// Adds a compact `meta.xml` payload.
    #[must_use]
    pub fn meta_xml(mut self, xml: impl Into<String>) -> Self {
        self.meta_xml = Some(xml.into());
        self
    }

    /// Adds typed document metadata.
    #[must_use]
    pub fn metadata(mut self, metadata: litchi_core::Metadata) -> Self {
        self.metadata = Some(metadata);
        self
    }

    /// Adds a typed paragraph to freshly authored content.
    #[must_use]
    pub fn paragraph(mut self, paragraph: crate::paragraph::Paragraph) -> Self {
        self.blocks.push(crate::ContentBlock::Paragraph(paragraph));
        self.paragraph_count = self.paragraph_count.saturating_add(1);
        self
    }

    /// Adds a rich paragraph whose inline content is semantically reopened.
    ///
    /// # Errors
    ///
    /// Returns an error when an inline item cannot be safely authored.
    pub fn rich_paragraph(
        mut self,
        content_input: impl IntoIterator<Item = crate::inline::Content>,
    ) -> Result<Self> {
        let content_items = content_input.into_iter().collect::<Vec<_>>();
        let (_xml, text) = render_inline(&content_items)?;
        self.inline_changes.push((
            crate::facade::InlineBlock::Paragraph(litchi_core::Position::new(self.paragraph_count)),
            content_items,
        ));
        self.blocks.push(crate::ContentBlock::Paragraph(
            crate::paragraph::Paragraph::new(text),
        ));
        self.paragraph_count = self.paragraph_count.saturating_add(1);
        Ok(self)
    }

    /// Adds a styled rich paragraph whose inline content is semantically reopened.
    ///
    /// # Errors
    ///
    /// Returns an error when an inline item cannot be safely authored.
    pub fn styled_rich_paragraph(
        mut self,
        style_name: impl Into<String>,
        content_input: impl IntoIterator<Item = crate::inline::Content>,
    ) -> Result<Self> {
        let content_items = content_input.into_iter().collect::<Vec<_>>();
        let (_xml, text) = render_inline(&content_items)?;
        self.inline_changes.push((
            crate::facade::InlineBlock::Paragraph(litchi_core::Position::new(self.paragraph_count)),
            content_items,
        ));
        self.blocks.push(crate::ContentBlock::Paragraph(
            crate::paragraph::Paragraph::styled(text, style_name),
        ));
        self.paragraph_count = self.paragraph_count.saturating_add(1);
        Ok(self)
    }

    /// Adds a rich heading whose inline content is semantically reopened.
    ///
    /// # Errors
    ///
    /// Returns an error for level zero or unsafe inline content.
    pub fn rich_heading(
        mut self,
        level: u8,
        content_input: impl IntoIterator<Item = crate::inline::Content>,
    ) -> Result<Self> {
        let content_items = content_input.into_iter().collect::<Vec<_>>();
        let (_xml, text) = render_inline(&content_items)?;
        let heading = crate::heading::Heading::new(level, text)?;
        self.inline_changes.push((
            crate::facade::InlineBlock::Heading(litchi_core::Position::new(self.heading_count)),
            content_items,
        ));
        self.blocks.push(crate::ContentBlock::Heading(heading));
        self.heading_count = self.heading_count.saturating_add(1);
        Ok(self)
    }

    /// Adds a styled rich heading whose inline content is semantically reopened.
    ///
    /// # Errors
    ///
    /// Returns an error for level zero or unsafe inline content.
    pub fn styled_rich_heading(
        mut self,
        level: u8,
        style_name: impl Into<String>,
        content_input: impl IntoIterator<Item = crate::inline::Content>,
    ) -> Result<Self> {
        let content_items = content_input.into_iter().collect::<Vec<_>>();
        let (_xml, text) = render_inline(&content_items)?;
        let heading = crate::heading::Heading::styled(level, text, style_name)?;
        self.inline_changes.push((
            crate::facade::InlineBlock::Heading(litchi_core::Position::new(self.heading_count)),
            content_items,
        ));
        self.blocks.push(crate::ContentBlock::Heading(heading));
        self.heading_count = self.heading_count.saturating_add(1);
        Ok(self)
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

    /// Adds one typed common style.
    #[must_use]
    pub fn style(mut self, style: crate::style::Style) -> Self {
        self.styles.push(style);
        self
    }

    /// Adds an external inert resource reference without resolving it.
    ///
    /// # Errors
    ///
    /// Returns an error for a package-relative reference.
    pub fn external_resource(mut self, resource: crate::resource::Resource) -> Result<Self> {
        if resource.is_embedded() || resource.href().starts_with('#') {
            return Err(Error::InvalidFormat(
                "OTH fresh external resource must contain an external URI".to_string(),
            ));
        }
        self.resources.push(FreshResource {
            members: Vec::new(),
            reference: resource,
        });
        Ok(self)
    }

    /// Adds one embedded file resource and its inert reference atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for an external or unsafe reference or empty media type.
    pub fn resource_with_payload(
        mut self,
        resource: crate::resource::Resource,
        media_type_input: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<Self> {
        let path = embedded_path(resource.href())?.to_owned();
        let media_type_value = media_type_input.into();
        if media_type_value.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH fresh resource media type cannot be empty".to_string(),
            ));
        }
        self.resources.push(FreshResource {
            members: vec![ResourceMember {
                bytes,
                media_type: media_type_value,
                path,
            }],
            reference: resource,
        });
        Ok(self)
    }

    /// Adds a directory-backed embedded object and all of its members atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for an image, external/unsafe root, or empty member set.
    pub fn object_with_members(
        mut self,
        resource: crate::resource::Resource,
        member_input: impl IntoIterator<Item = ResourceMember>,
    ) -> Result<Self> {
        if resource.kind() == crate::resource::Kind::Image {
            return Err(Error::InvalidFormat(
                "OTH fresh object members require an object resource kind".to_string(),
            ));
        }
        let root = embedded_path(resource.href())?;
        let member_values = member_input
            .into_iter()
            .map(|member| {
                let ResourceMember {
                    bytes,
                    media_type,
                    path,
                } = member;
                ResourceMember {
                    bytes,
                    media_type,
                    path: format!("{root}/{path}"),
                }
            })
            .collect::<Vec<_>>();
        if member_values.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH fresh embedded object requires at least one member".to_string(),
            ));
        }
        self.resources.push(FreshResource {
            members: member_values,
            reference: resource,
        });
        Ok(self)
    }

    /// Replaces the `content.xml` payload.
    ///
    /// Raw and typed content are mutually exclusive so call order
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
    /// readback fails, or raw and typed content were both supplied.
    pub fn build(self) -> Result<Vec<u8>> {
        if self.content_xml.is_some()
            && (!self.blocks.is_empty() || !self.forms.is_empty() || !self.resources.is_empty())
        {
            return Err(Error::InvalidFormat(
                "OTH builder cannot combine raw content_xml with typed content".to_string(),
            ));
        }
        if self.meta_xml.is_some() && self.metadata.is_some() {
            return Err(Error::InvalidFormat(
                "OTH builder cannot combine raw and typed metadata".to_string(),
            ));
        }
        if self.styles_xml.is_some() && !self.styles.is_empty() {
            return Err(Error::InvalidFormat(
                "OTH builder cannot combine raw and typed styles".to_string(),
            ));
        }
        let content_xml = match self.content_xml {
            Some(xml) => xml,
            None => render_document(&self.blocks, &self.forms, &self.resources)?,
        };
        crate::codec::validate_authored(&content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", content_xml.as_bytes())?;
        let typed_meta_xml = self.metadata.as_ref().map(render_metadata).transpose()?;
        let typed_styles_xml = (!self.styles.is_empty())
            .then(|| render_styles(&self.styles))
            .transpose()?;
        let mut paths = vec![
            "mimetype".to_string(),
            "content.xml".to_string(),
            "META-INF/manifest.xml".to_string(),
        ];
        for (path, optional_xml) in [
            ("meta.xml", self.meta_xml.or(typed_meta_xml)),
            ("settings.xml", self.settings_xml),
            ("styles.xml", self.styles_xml.or(typed_styles_xml)),
        ] {
            if let Some(part_xml) = optional_xml {
                compact_xml::validate(part_xml.as_bytes()).map_err(Error::from)?;
                writer.add_file(path, part_xml.as_bytes())?;
                paths.push(path.to_string());
            }
        }
        for resource in &self.resources {
            for member in &resource.members {
                if paths.iter().any(|path| path == &member.path) {
                    return Err(Error::InvalidFormat(
                        "OTH fresh resource package path is duplicated".to_string(),
                    ));
                }
                paths.push(member.path.clone());
                writer.add_file_with_media_type(&member.path, &member.bytes, &member.media_type)?;
            }
        }
        let bytes = crate::package::Snapshot::from_bytes(writer.finish_to_bytes()?)?.into_bytes();
        if self.inline_changes.is_empty() {
            return Ok(bytes);
        }
        let template = crate::facade::Template::from_bytes(bytes)?;
        let mut edit = template.edit();
        for (block, content) in &self.inline_changes {
            match *block {
                crate::facade::InlineBlock::Paragraph(position) => {
                    edit.set_paragraph_inline(position, content)?;
                },
                crate::facade::InlineBlock::Heading(position) => {
                    edit.set_heading_inline(position, content)?;
                },
            }
        }
        Ok(edit.commit()?.into_template().into_bytes())
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn render_blocks(blocks: &[crate::ContentBlock]) -> Result<String> {
    render_document(blocks, &[], &[])
}

fn render_document(
    blocks: &[crate::ContentBlock],
    forms: &[crate::form::Form],
    resources: &[FreshResource],
) -> Result<String> {
    let forms_xml = render_forms(forms)?;
    let mut capacity = CONTENT_PREFIX
        .len()
        .saturating_add(CONTENT_SUFFIX.len())
        .saturating_add(forms_xml.len());
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
            crate::ContentBlock::List(list) => list_authored_size(list),
        };
        capacity = capacity
            .checked_add(markup_bytes)
            .and_then(|size| size.checked_add(text_bytes.saturating_mul(6)))
            .ok_or_else(|| {
                Error::InvalidFormat("OTH authored content size overflow".to_string())
            })?;
    }
    for resource in resources {
        capacity = capacity
            .checked_add(64)
            .and_then(|size| size.checked_add(render_resource(&resource.reference).len()))
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
    output.push_str(&forms_xml);
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
    for resource in resources {
        output.push_str(
            "<text:p><draw:frame xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\">",
        );
        output.push_str(&render_resource(&resource.reference));
        output.push_str("</draw:frame></text:p>");
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
                if !value.is_empty() && value.bytes().all(|byte| byte == b' ') {
                    xml.push_str(
                        "<text:s xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"",
                    );
                    if value.len() > 1 {
                        xml.push_str(" text:c=\"");
                        xml.push_str(&value.len().to_string());
                        xml.push('"');
                    }
                    xml.push_str("/>");
                } else {
                    xml.push_str(&quick_xml::escape::escape(value));
                }
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

fn list_paragraph_count(list: &crate::list::List) -> usize {
    list.items()
        .iter()
        .map(|item| {
            item.nested_lists()
                .iter()
                .map(list_paragraph_count)
                .fold(item.paragraphs().len(), usize::saturating_add)
        })
        .fold(0_usize, usize::saturating_add)
}

fn list_authored_size(list: &crate::list::List) -> (usize, usize) {
    let mut markup_bytes = 64_usize.saturating_add(list.style_name().map_or(0, str::len));
    let mut text_bytes = 0_usize;
    for item in list.items() {
        markup_bytes = markup_bytes.saturating_add(48);
        for paragraph in item.paragraphs() {
            markup_bytes = markup_bytes
                .saturating_add(32)
                .saturating_add(paragraph.style_name().map_or(0, str::len));
            text_bytes = text_bytes.saturating_add(paragraph.text().len());
        }
        for nested in item.nested_lists() {
            let (nested_markup, nested_text) = list_authored_size(nested);
            markup_bytes = markup_bytes.saturating_add(nested_markup);
            text_bytes = text_bytes.saturating_add(nested_text);
        }
    }
    (markup_bytes, text_bytes)
}

fn embedded_path(href: &str) -> Result<&str> {
    if href.is_empty() || href.starts_with('#') || href.contains("://") || href.starts_with("data:")
    {
        return Err(Error::InvalidFormat(
            "OTH fresh resource reference must be package-relative".to_string(),
        ));
    }
    let path = href.strip_prefix("./").unwrap_or(href);
    validate_relative_path(path)?;
    Ok(path)
}

fn validate_relative_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.starts_with('/')
        || path.contains("://")
        || path.starts_with("data:")
        || path.contains('\\')
        || path.chars().any(char::is_control)
        || path
            .split('/')
            .any(|segment| matches!(segment, "" | "." | ".."))
    {
        return Err(Error::InvalidFormat(
            "OTH fresh resource path is not a safe package member".to_string(),
        ));
    }
    Ok(())
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

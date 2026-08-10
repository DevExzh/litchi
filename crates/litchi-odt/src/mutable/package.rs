//! Document/package snapshots and transactional ODT package output.

use super::model::{DocumentElement, MutableDocument};
use crate::Document;
use crate::core::{MetaXmlPatch, PackageWriter, Structure, patch_meta_xml};
use crate::elements::parser::OrderElement;
use crate::elements::text::Paragraph;
use litchi_core::{Result, xml::escape_xml};
use litchi_odf_common::package::xml_splice_publication;
use std::path::Path;

impl MutableDocument {
    /// Create a mutable document from an existing Document.
    ///
    /// This parses the document structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `doc` - The document to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Document;
    /// use litchi_odt::mutable::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let mut mutable_doc = MutableDocument::from_document(doc)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_document(doc: Document) -> Result<Self> {
        let content_xml = String::from_utf8(doc.get_file("content.xml")?).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!("content.xml is not UTF-8: {error}"))
        })?;
        let source_elements = doc.elements()?;
        let metadata = doc.metadata()?;

        // Get MIME type from package
        let mimetype = "application/vnd.oasis.opendocument.text".to_string();

        // Extract styles XML from the document's package
        let styles_xml = doc
            .get_file("styles.xml")
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok());

        let elements = source_elements
            .into_iter()
            .map(|element| match element {
                OrderElement::Paragraph(paragraph) => DocumentElement::Paragraph(paragraph),
                OrderElement::NumberedParagraph(paragraph) => {
                    DocumentElement::Paragraph(paragraph.into_paragraph())
                },
                OrderElement::Heading(heading) => DocumentElement::Heading(heading),
                OrderElement::Table(table) => DocumentElement::Table(table),
                OrderElement::List(list) => DocumentElement::List(list),
            })
            .collect();
        let source_package = Some(doc.into_package());

        Ok(Self {
            elements,
            metadata,
            mimetype,
            styles_xml,
            source_package,
            content_xml: Some(content_xml),
            pending_images: Vec::new(),
            next_frame_number: 1,
        })
    }
}

impl MutableDocument {
    /// Override the root MIME type written by `to_bytes`.
    ///
    /// Used by the web-template authoring model to emit the legacy
    /// `application/vnd.oasis.opendocument.text-web` MIME type.
    pub fn set_mimetype(&mut self, mimetype: impl Into<String>) {
        self.mimetype = mimetype.into();
    }
}

impl MutableDocument {
    /// Insert an image frame as a new paragraph at a specific index.
    ///
    /// The payload is sniffed (PNG, JPEG, and GIF are accepted), stored
    /// verbatim under `Pictures/` in the package, and referenced from a
    /// `draw:frame`/`draw:image` element with the given geometry and anchor.
    /// Returns the allocated package path of the picture part.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::frame::{Anchor, Length};
    /// use litchi_odt::mutable::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// let png = b"\x89PNG\r\n\x1a\n".as_slice();
    /// let path = doc.insert_image(
    ///     0,
    ///     png,
    ///     &Length::centimeters(10.0),
    ///     &Length::centimeters(4.0),
    ///     Anchor::AsChar,
    /// )?;
    /// assert!(path.starts_with("Pictures/"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_image(
        &mut self,
        index: usize,
        image: &[u8],
        width: &crate::frame::Length,
        height: &crate::frame::Length,
        anchor: crate::frame::Anchor,
    ) -> Result<String> {
        use crate::frame;
        let format = frame::validate_payload(image)?;
        let path = frame::allocate_picture_path(format.extension(), |candidate| {
            // Picture numbering is global: a stem taken by any supported
            // extension blocks the whole index.
            let taken = |path: &str| {
                self.pending_images
                    .iter()
                    .any(|pending| pending.path() == path)
                    || self
                        .source_package
                        .as_ref()
                        .is_some_and(|package| package.has_file(path).unwrap_or(false))
            };
            if taken(candidate) {
                return true;
            }
            let stem = candidate.trim_end_matches(format.extension());
            ["png", "jpg", "gif"]
                .iter()
                .any(|extension| taken(&format!("{stem}{extension}")))
        })?;
        let name = format!("Frame {}", self.next_frame_number);
        let frame = frame::Frame::new(&name, width.clone(), height.clone(), anchor)?;
        let frame_element = frame::image_element(&frame, &path);
        let mut paragraph_element = crate::elements::element::Element::new("text:p");
        paragraph_element.add_child(frame_element);
        let paragraph = Paragraph::from_element(paragraph_element)?;

        if index > self.elements.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )));
        }
        let pending = frame::Part::new(path.clone(), image.to_vec())?;
        self.invalidate_content_xml();
        self.elements
            .insert(index, DocumentElement::Paragraph(paragraph));
        self.pending_images.push(pending);
        self.next_frame_number += 1;
        Ok(path)
    }

    /// Insert a plain-text text-box frame as a new paragraph at a specific index.
    ///
    /// The box is a `draw:frame` wrapping `draw:text-box`; newlines in `text`
    /// become separate paragraphs in the box story. Returns the frame name.
    pub fn insert_text_box(
        &mut self,
        index: usize,
        text: &str,
        width: &crate::frame::Length,
        height: &crate::frame::Length,
        anchor: crate::frame::Anchor,
    ) -> Result<String> {
        use crate::frame;
        let name = format!("Text Box {}", self.next_frame_number);
        let frame = frame::Frame::new(&name, width.clone(), height.clone(), anchor)?;
        let frame_element = frame::text_box_element(&frame, text)?;

        if index > self.elements.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )));
        }
        self.invalidate_content_xml();
        self.elements
            .insert(index, DocumentElement::Frame(frame_element));
        self.next_frame_number += 1;
        Ok(name)
    }
}

impl MutableDocument {
    /// Generate content.xml from the current mutable state.
    pub(super) fn generate_content_xml(&self) -> String {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map_or(0, |t| t.len()),
                DocumentElement::Heading(h) => h.text().map_or(0, |t| t.len()),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
                DocumentElement::Frame(_) => 256,
            })
            .sum::<usize>();
        let mut body = String::with_capacity(estimated);

        // Add elements in their insertion order (paragraphs and tables mixed)
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Frame(frame) => {
                    body.push_str(&frame.to_xml_string());
                },
            }
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> Result<String> {
        if let Some(patched) = self.patched_source_meta_xml()? {
            return Ok(patched);
        }
        Ok(self.generate_meta_xml_from_scratch())
    }

    /// Patch the retained source meta.xml so metadata the edit did not change
    /// survives the save, while fields set through the mutable API are updated
    /// in place. Existing timestamps and generator metadata remain untouched.
    fn patched_source_meta_xml(&self) -> Result<Option<String>> {
        let Some(package) = &self.source_package else {
            return Ok(None);
        };
        let Ok(bytes) = package.get_file("meta.xml") else {
            return Ok(None);
        };
        let Ok(source) = String::from_utf8(bytes) else {
            return Ok(None);
        };
        let source_metadata = crate::Metadata::from_xml(&source)?;
        let patch =
            MetaXmlPatch::preserve_all().diff_simple_fields(&source_metadata, &self.metadata);
        patch_meta_xml(&source, &patch)
    }

    /// Generate meta.xml from the mutable metadata model alone.
    fn generate_meta_xml_from_scratch(&self) -> String {
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map_or(0, String::len);
        estimated += self.metadata.author.as_ref().map_or(0, String::len);
        estimated += self.metadata.subject.as_ref().map_or(0, String::len);
        estimated += self.metadata.description.as_ref().map_or(0, String::len);
        estimated += self.metadata.keywords.as_ref().map_or(0, String::len);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        if let Some(ref subject) = self.metadata.subject {
            let escaped_subject = escape_xml(subject);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:subject>{}</dc:subject>"#,
                escaped_subject
            ));
        }

        if let Some(ref description) = self.metadata.description {
            let escaped_description = escape_xml(description);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:description>{}</dc:description>"#,
                escaped_description
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            let escaped_keywords = escape_xml(keywords);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:keyword>{}</meta:keyword>"#,
                escaped_keywords
            ));
        }

        if let Some(ref identifier) = self.metadata.identifier {
            let escaped_identifier = escape_xml(identifier);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:identifier>{}</dc:identifier>"#,
                escaped_identifier
            ));
        }

        if let Some(ref language) = self.metadata.language {
            let escaped_language = escape_xml(language);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:language>{}</dc:language>"#,
                escaped_language
            ));
        }

        if let Some(created) = &self.metadata.created {
            let value = created.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:creation-date>{}</meta:creation-date>"#,
                value
            ));
        } else if let Some(created) = &self.metadata.created_local {
            let value = created.format("%Y-%m-%dT%H:%M:%S%.f").to_string();
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:creation-date>{}</meta:creation-date>"#,
                value
            ));
        }

        if let Some(modified) = &self.metadata.modified {
            let value = modified.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:date>{}</dc:date>"#,
                value
            ));
        } else if let Some(modified) = &self.metadata.modified_local {
            let value = modified.format("%Y-%m-%dT%H:%M:%S%.f").to_string();
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:date>{}</dc:date>"#,
                value
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator>{}</office:meta></office:document-meta>"#,
            meta_fields
        )
    }

    /// Save the modified document to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::mutable::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::mutable::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// let bytes = doc.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let generated_content_xml;
        let content_xml = if let Some(content_xml) = self.content_xml.as_deref() {
            content_xml
        } else {
            generated_content_xml = self.generate_content_xml();
            &generated_content_xml
        };
        let content_splice = self
            .source_package
            .as_ref()
            .map(|source| {
                crate::font_face::content_font_face_splice_publication(source, content_xml)
            })
            .transpose()?
            .flatten();
        if let Some(publication) = content_splice {
            publication.publish(&mut writer)?;
        } else {
            writer.add_file("content.xml", content_xml.as_bytes())?;
        }

        // Add styles.xml (preserved or default)
        let default_styles = Structure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        let styles_splice = if self.source_package.as_ref().is_some_and(|source| {
            source
                .package()
                .is_ok_and(|package| package.has_file("styles.xml"))
        }) {
            let font_face = self
                .source_package
                .as_ref()
                .map(|source| {
                    crate::font_face::styles_font_face_splice_publication(source, styles_xml)
                })
                .transpose()?
                .flatten();
            if font_face.is_some() {
                font_face
            } else {
                self.source_package
                    .as_ref()
                    .map(|source| crate::outline_style::splice_publication(source, styles_xml))
                    .transpose()?
                    .flatten()
            }
        } else {
            None
        };
        if let Some(publication) = styles_splice {
            publication.publish(&mut writer)?;
        } else {
            writer.add_file("styles.xml", styles_xml.as_bytes())?;
        }

        // Add meta.xml (patched from the source or regenerated with current metadata)
        let meta_xml = self.generate_meta_xml()?;
        let meta_splice = self
            .source_package
            .as_ref()
            .map(|source| xml_splice_publication(source, "meta.xml", &meta_xml))
            .transpose();
        if let Ok(Some(publication)) = meta_splice {
            publication.publish(&mut writer)?;
        } else {
            writer.add_file("meta.xml", meta_xml.as_bytes())?;
        }

        // Add authored picture payloads.
        for pending in &self.pending_images {
            writer.add_file(pending.path(), pending.bytes())?;
        }

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

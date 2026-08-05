//! OPC relationship and package-graph accessors for the document facade.

use crate::Variables;
use crate::bookmark::Bookmark;
use crate::comment::Comment;
use crate::custom_xml::Part as CustomXmlPart;
use crate::error::{Error, Result};
use crate::footnote::Note;
use crate::header_footer::{Kind, Story};
use crate::mail_merge::{Recipients, extract_recipients, is_settings_relationship};
use crate::numbering::{Collection, parse_part};
use crate::settings::{DocumentSettings, extract_document_variables};
use crate::theme::Theme;
use crate::web;
use crate::writer::Watermark;
use litchi_opc::constants::relationship_type;

use super::super::model::{Document, ImageWatermarkPart};

impl<'a> Document<'a> {
    /// Return the package's glossary/building-block catalog and dialect.
    pub fn glossary(
        &self,
    ) -> Result<Option<(crate::glossary::Catalog, crate::glossary::Conformance)>> {
        Ok(crate::glossary::load(self.opc)?)
    }

    /// Load the typed, inert SmartArt (DrawingML diagram) inventory anchored
    /// in this document.
    ///
    /// Each returned [`crate::smartart::Diagram`] carries the
    /// parsed data-model node tree, the layout/quick-style/colors part
    /// metadata, and the diagram part names. Both transitional and Strict
    /// namespace dialects are supported.
    pub fn smart_arts(&self) -> Result<Vec<crate::smartart::Diagram>> {
        crate::smartart::load_smart_arts(self.opc, self.part.part().partname())
    }

    /// Load the typed, inert text-box and WordArt inventory anchored in this
    /// document.
    ///
    /// Each returned [`crate::textbox::TextBox`] carries the shape
    /// identity, the `wps:bodyPr` text-body properties, the story as
    /// paragraphs with runs, and WordArt warp/styling presence flags. Both
    /// DrawingML shapes and legacy VML `w:pict` fallbacks are recognized, in
    /// both the transitional and Strict namespace dialects.
    pub fn text_boxes(&self) -> Result<Vec<crate::textbox::TextBox>> {
        crate::textbox::load_text_boxes(self.part.xml_bytes())
    }
    /// Get all headers in the document.
    ///
    /// Returns header stories in WordprocessingML section-reference order.
    /// Each story carries its typed [`Kind`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for header in doc.headers()? {
    ///     println!("{:?} header: {}", header.kind(), header.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn headers(&self) -> Result<Vec<Story>> {
        crate::header_footer::load_headers(self.opc, &self.part)
    }

    /// Return distinct standard VML text watermarks from document headers.
    ///
    /// Word commonly repeats the same watermark in default, first-page, and
    /// even-page headers; equivalent copies are returned once in section
    /// reference order.
    pub fn watermarks(&self) -> Result<Vec<Watermark>> {
        let mut watermarks = Vec::new();
        for header in self.headers()? {
            for watermark in header.watermarks()? {
                if !watermarks.contains(&watermark) {
                    watermarks.push(watermark);
                }
            }
        }
        Ok(watermarks)
    }

    /// Return picture watermarks from document headers with their media
    /// parts resolved.
    ///
    /// Each entry pairs a `v:imagedata` anchor discovered in a header with
    /// the relationship-resolved media part name and payload bytes. The
    /// payload is an inert byte view; it is never decoded or displayed.
    pub fn image_watermarks(&self) -> Result<Vec<ImageWatermarkPart<'_>>> {
        let mut parts = Vec::new();
        for image in crate::header_footer::image_watermarks(self.opc, &self.part)? {
            parts.push(ImageWatermarkPart {
                source_header_name: image.source_header_name,
                relationship_id: image.relationship_id,
                part_name: image.part_name,
                content_type: image.content_type,
                bytes: image.bytes,
            });
        }
        Ok(parts)
    }

    /// Get all footers in the document.
    ///
    /// Returns footer stories in WordprocessingML section-reference order.
    /// Each story carries its typed [`Kind`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for footer in doc.footers()? {
    ///     println!("{:?} footer: {}", footer.kind(), footer.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footers(&self) -> Result<Vec<Story>> {
        crate::header_footer::load_footers(self.opc, &self.part)
    }

    /// Get a specific header by type.
    ///
    /// # Arguments
    /// * `hdr_type` - The type of header to retrieve
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::{header_footer::Kind, Package};
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(header) = doc.header(Kind::Primary)? {
    ///     println!("Primary header: {}", header.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn header(&self, hdr_type: Kind) -> Result<Option<Story>> {
        let headers = self.headers()?;
        Ok(headers.into_iter().find(|header| header.kind() == hdr_type))
    }

    /// Get a specific footer by type.
    ///
    /// # Arguments
    /// * `ftr_type` - The type of footer to retrieve
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::{header_footer::Kind, Package};
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(footer) = doc.footer(Kind::Primary)? {
    ///     println!("Primary footer: {}", footer.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footer(&self, ftr_type: Kind) -> Result<Option<Story>> {
        let footers = self.footers()?;
        Ok(footers.into_iter().find(|footer| footer.kind() == ftr_type))
    }

    /// Get all `<w:hyperlink>` element hyperlinks in the document.
    ///
    /// Returns external URL and internal bookmark links stored as
    /// `<w:hyperlink>` elements. `HYPERLINK` field instructions are exposed
    /// separately by `Self::hyperlink_fields`.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for link in doc.hyperlinks()? {
    ///     println!("Link text: {}", link.text());
    ///     if let Some(url) = link.url() {
    ///         println!("  URL: {}", url);
    ///     }
    ///     if let Some(anchor) = link.anchor() {
    ///         println!("  Bookmark: {}", anchor);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<crate::hyperlink::Hyperlink>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();
        let xml_bytes = self.part.xml_bytes();

        Ok(crate::hyperlink::Hyperlink::extract_from_document(
            xml_bytes, rels,
        )?)
    }

    /// Get the number of `<w:hyperlink>` element hyperlinks in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} hyperlinks", doc.hyperlink_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn hyperlink_count(&self) -> Result<usize> {
        Ok(self.hyperlinks()?.len())
    }

    /// Get all footnotes in the document.
    ///
    /// Returns a vector of `Note` objects representing all footnotes
    /// in the document (excluding separators).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for note in doc.footnotes()? {
    ///     println!("Footnote {}: {}", note.id(), note.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footnotes(&self) -> Result<Vec<Note>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for footnotes relationship
        match rels.part_with_reltype(relationship_type::FOOTNOTES) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let footnotes_part = self.opc.get_part(&target)?;
                Note::extract_footnotes_from_part(footnotes_part)
            },
            Err(_) => {
                // No footnotes in document
                Ok(Vec::new())
            },
        }
    }

    /// Get all endnotes in the document.
    ///
    /// Returns a vector of `Note` objects representing all endnotes
    /// in the document (excluding separators).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for note in doc.endnotes()? {
    ///     println!("Endnote {}: {}", note.id(), note.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn endnotes(&self) -> Result<Vec<Note>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for endnotes relationship
        match rels.part_with_reltype(relationship_type::ENDNOTES) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let endnotes_part = self.opc.get_part(&target)?;
                Note::extract_endnotes_from_part(endnotes_part)
            },
            Err(_) => {
                // No endnotes in document
                Ok(Vec::new())
            },
        }
    }

    /// Get the number of footnotes in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} footnotes", doc.footnote_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn footnote_count(&self) -> Result<usize> {
        Ok(self.footnotes()?.len())
    }

    /// Get the number of endnotes in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} endnotes", doc.endnote_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn endnote_count(&self) -> Result<usize> {
        Ok(self.endnotes()?.len())
    }

    /// Get all comments in the document.
    ///
    /// Returns a vector of `Comment` objects representing all comments
    /// in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for comment in doc.comments()? {
    ///     println!("{} commented: {}", comment.author(), comment.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for comments relationship
        match rels.part_with_reltype(relationship_type::COMMENTS) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let comments_part = self.opc.get_part(&target)?;
                Comment::extract_from_part(comments_part)
            },
            Err(_) => {
                // No comments in document
                Ok(Vec::new())
            },
        }
    }

    /// Get the number of comments in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} comments", doc.comment_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn comment_count(&self) -> Result<usize> {
        Ok(self.comments()?.len())
    }

    /// Get all bookmarks in the document.
    ///
    /// Returns a vector of `Bookmark` objects representing all bookmarks
    /// in the document (excluding system bookmarks).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for bookmark in doc.bookmarks()? {
    ///     println!("Bookmark: {}", bookmark.name());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn bookmarks(&self) -> Result<Vec<Bookmark>> {
        let xml_bytes = self.part.xml_bytes();
        Bookmark::extract_from_document(xml_bytes)
    }

    /// Get the number of bookmarks in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// println!("Document has {} bookmarks", doc.bookmark_count()?);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn bookmark_count(&self) -> Result<usize> {
        Ok(self.bookmarks()?.len())
    }

    /// Get all fields in the document.
    ///
    /// Returns a vector of `Field` objects representing all fields
    /// in the document (PAGE, DATE, REF, etc.).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for field in doc.fields()? {
    ///     println!("Field {}: {}", field.field_type(), field.instruction());
    ///     if let Some(result) = field.result() {
    ///         println!("  Result: {}", result);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn numbering(&self) -> Result<Option<Collection>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Look for numbering relationship
        match rels.part_with_reltype(relationship_type::NUMBERING) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let numbering_part = self.opc.get_part(&target)?;
                Ok(Some(parse_part(numbering_part)?))
            },
            Err(_) => {
                // No numbering in document
                Ok(None)
            },
        }
    }

    /// Get the document settings including protection status.
    ///
    /// Returns a `Settings` object providing access to document settings
    /// such as protection status, track revisions, and zoom level.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(settings) = doc.settings()? {
    ///     if settings.is_protected() {
    ///         println!("Document is protected");
    ///         if let Some(ptype) = settings.protection_type() {
    ///             println!("Protection type: {:?}", ptype);
    ///         }
    ///     }
    ///     println!("Track revisions: {}", settings.track_revisions());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn settings(&self) -> Result<Option<DocumentSettings>> {
        let main_part = self.opc.main_document_part()?;
        let mut matches = main_part
            .rels()
            .iter()
            .filter(|rel| is_settings_relationship(rel.reltype()));
        let Some(rel) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        if rel.is_external() {
            return Err(Error::InvalidFormat(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = rel.target_partname()?;
        let settings_part = self.opc.get_part(&target)?;
        Ok(Some(DocumentSettings::extract_from_part(settings_part)?))
    }

    /// Load the ISO mail-merge recipient-data part referenced by `settings.xml`.
    pub fn mail_merge_recipients(&self) -> Result<Option<Recipients>> {
        let Some(settings) = self.settings()? else {
            return Ok(None);
        };
        let Some(relationship_id) = settings
            .mail_merge()
            .and_then(|merge| merge.odso())
            .and_then(|odso| odso.recipient_data_relationship_id())
        else {
            return Ok(None);
        };
        let main_part = self.opc.main_document_part()?;
        let settings_relationship = main_part
            .rels()
            .iter()
            .find(|rel| is_settings_relationship(rel.reltype()))
            .ok_or_else(|| Error::InvalidFormat("settings relationship is missing".into()))?;
        let settings_part = self
            .opc
            .get_part(&settings_relationship.target_partname()?)?;
        let relationship = settings_part.rels().get(relationship_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "recipient-data relationship '{relationship_id}' is missing"
            ))
        })?;
        let recipient_part = self.opc.get_part(&relationship.target_partname()?)?;
        Ok(Some(extract_recipients(recipient_part)?))
    }

    /// Read the document's typed web-output settings and conformance family.
    pub fn web(&self) -> Result<Option<(web::Settings, web::Conformance)>> {
        Ok(web::load(self.opc)?)
    }
    /// Get document variables.
    ///
    /// Returns document variables stored in the settings, which can be
    /// referenced by fields and used for mail merge.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(vars) = doc.document_variables()? {
    ///     for (name, value) in vars.iter() {
    ///         println!("{} = {}", name, value);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document_variables(&self) -> Result<Option<Variables>> {
        let main_part = self.opc.main_document_part()?;
        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let mut matches = main_part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::SETTINGS | STRICT_SETTINGS_RELATIONSHIP
            )
        });
        let Some(relationship) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::InvalidFormat(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let settings_part = self.opc.get_part(&target)?;
        Ok(Some(extract_document_variables(settings_part)?))
    }

    /// Get the document theme.
    ///
    /// Returns the theme containing color scheme, font scheme, and format scheme.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if let Some(theme) = doc.theme()? {
    ///     if let Some(name) = theme.name() {
    ///         println!("Theme: {}", name);
    ///     }
    ///     if let Some(major) = theme.major_font() {
    ///         println!("Major font: {}", major);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn theme(&self) -> Result<Option<Theme>> {
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        match rels.part_with_reltype(relationship_type::THEME) {
            Ok(rel) => {
                let target = rel.target_partname()?;
                let theme_part = self.opc.get_part(&target)?;
                Ok(Some(Theme::extract_from_part(theme_part)?))
            },
            Err(_) => Ok(None),
        }
    }
    /// Get custom XML parts from the document.
    ///
    /// Returns a vector of custom XML parts that store arbitrary XML data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for xml_part in doc.custom_xml()? {
    ///     println!("Custom XML part: {}", xml_part.id());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn custom_xml(&self) -> Result<Vec<CustomXmlPart>> {
        let mut custom_parts = Vec::new();

        // Custom XML parts are stored as relationships from the main document part
        let main_part = self.opc.main_document_part()?;
        let rels = main_part.rels();

        // Iterate through all relationships to find custom XML parts
        for rel in rels.iter() {
            if rel.reltype() == relationship_type::CUSTOM_XML {
                let target = rel.target_partname()?;
                let part = self.opc.get_part(&target)?;
                let id = rel.r_id().to_string();
                let custom_xml = CustomXmlPart::from_part(part, id)?;
                custom_parts.push(custom_xml);
            }
        }

        Ok(custom_parts)
    }
}

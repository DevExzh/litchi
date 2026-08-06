//! Word-owned XML part publication and lossless settings snapshots.

use super::super::model::*;

use crate::numbering::{Patch as NumberingPatch, Snapshot as NumberingSnapshot};

impl Package {
    /// Load the source-preserving numbering snapshot from the main document.
    pub fn numbering_snapshot(&self) -> Result<Option<NumberingSnapshot>> {
        use litchi_opc::constants::relationship_type as rt;

        let main = self.opc.main_document_part()?;
        let relationship = match main.rels().part_with_reltype(rt::NUMBERING) {
            Ok(relationship) => relationship,
            Err(_) => return Ok(None),
        };
        let target = relationship.target_partname()?;
        let part = self.opc.get_part(&target)?;
        Ok(Some(crate::numbering::parse_snapshot_part(part)?))
    }

    /// Apply a numbering patch to the current package graph atomically.
    ///
    /// The source snapshot must still match the authored numbering part. The
    /// candidate is validated before the existing OPC transaction publishes
    /// its replacement, and a failed publication leaves the package untouched.
    pub fn apply_numbering_patch(
        &mut self,
        source: &NumberingSnapshot,
        patch: &NumberingPatch,
    ) -> Result<NumberingSnapshot> {
        let current = self
            .numbering_snapshot()?
            .ok_or_else(|| Error::PartNotFound("numbering part".into()))?;
        if current.xml_bytes() != source.xml_bytes() {
            return Err(Error::InvalidFormat(
                "numbering package source does not match its snapshot precondition".into(),
            ));
        }
        let candidate = patch.apply(&current)?;
        let replacement = candidate.xml_bytes().to_vec();
        self.edit_opc(|opc| {
            let target = {
                let main = opc.main_document_part()?;
                let relationship = main
                    .rels()
                    .part_with_reltype(litchi_opc::constants::relationship_type::NUMBERING)?;
                relationship.target_partname()?
            };
            opc.get_part_mut(&target)?.set_blob(replacement);
            Ok(())
        })?;
        Ok(candidate)
    }

    /// Update the footnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_footnotes_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let footnotes_uri = PackURI::new("/word/footnotes.xml")
            .map_err(|e| Error::InvalidUri(format!("footnotes URI: {}", e)))?;

        let content_type = ct::WML_FOOTNOTES.to_string();
        let footnotes_part = BlobPart::new(footnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the footnotes part
        self.opc.add_part(Box::new(footnotes_part));

        // Create relationship from document to footnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("footnotes.xml", rt::FOOTNOTES);
        }

        Ok(())
    }

    /// Update the endnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_endnotes_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let endnotes_uri = PackURI::new("/word/endnotes.xml")
            .map_err(|e| Error::InvalidUri(format!("endnotes URI: {}", e)))?;

        let content_type = ct::WML_ENDNOTES.to_string();
        let endnotes_part = BlobPart::new(endnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the endnotes part
        self.opc.add_part(Box::new(endnotes_part));

        // Create relationship from document to endnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("endnotes.xml", rt::ENDNOTES);
        }

        Ok(())
    }

    /// Update or create the comments part with the given XML content.
    pub(in crate::package) fn update_comments_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let comments_uri = PackURI::new("/word/comments.xml")
            .map_err(|e| Error::InvalidUri(format!("comments URI: {}", e)))?;

        let content_type = ct::WML_COMMENTS.to_string();
        let comments_part = BlobPart::new(comments_uri.clone(), content_type, xml.into_bytes());

        // Add the comments part
        self.opc.add_part(Box::new(comments_part));

        // Create relationship from document to comments
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("/word/comments.xml", rt::COMMENTS);
        }

        Ok(())
    }

    /// Update the settings.xml part with new content.
    pub(in crate::package) fn update_settings_part(&mut self, xml: Vec<u8>) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&part)?;
        self.commit_settings_part(&snapshot, part)
    }

    pub(super) fn settings_part_snapshot(&self) -> Result<SettingsPartSnapshot> {
        use litchi_opc::constants::relationship_type as rt;

        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let document = self.opc.main_document_part()?;
        let document_uri = document.partname().clone();
        let mut matches = document.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::SETTINGS | STRICT_SETTINGS_RELATIONSHIP
            )
        });
        let relationship = matches.next();
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let (target, relationship_exists) = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(Error::InvalidFormat(
                    "settings relationship cannot be external".into(),
                ));
            },
            Some(relationship) => (relationship.target_partname()?, true),
            None => (
                PackURI::new("/word/settings.xml")
                    .map_err(|error| Error::InvalidUri(error.to_string()))?,
                false,
            ),
        };

        let (content_type, xml, relationships) = match self.opc.get_part(&target) {
            Ok(part) => {
                if part.content_type() != ct::WML_SETTINGS {
                    return Err(Error::InvalidFormat(format!(
                        "settings part has content type {:?}, expected {:?}",
                        part.content_type(),
                        ct::WML_SETTINGS
                    )));
                }
                (
                    part.content_type().to_owned(),
                    part.blob().to_vec(),
                    part.rels()
                        .iter()
                        .map(|relationship| StoredRelationship {
                            reltype: relationship.reltype().to_owned(),
                            target: relationship.target_ref().to_owned(),
                            id: relationship.r_id().to_owned(),
                            external: relationship.is_external(),
                        })
                        .collect(),
                )
            },
            Err(_) if relationship_exists => {
                return Err(Error::PartNotFound(format!("settings part {target}")));
            },
            Err(_) => (
                ct::WML_SETTINGS.to_owned(),
                crate::template::default_settings_xml().as_bytes().to_vec(),
                Vec::new(),
            ),
        };
        Ok(SettingsPartSnapshot {
            document_uri,
            target,
            relationship_exists,
            content_type,
            xml,
            relationships,
        })
    }

    pub(super) fn commit_settings_part(
        &mut self,
        snapshot: &SettingsPartSnapshot,
        part: BlobPart,
    ) -> Result<()> {
        use litchi_opc::constants::relationship_type as rt;

        if !snapshot.relationship_exists {
            // Acquire the only fallible mutable reference before changing package state.
            self.opc
                .get_part_mut(&snapshot.document_uri)?
                .relate_to("settings.xml", rt::SETTINGS);
        }
        self.opc.add_part(Box::new(part));
        Ok(())
    }

    pub(in crate::package) fn update_theme_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::part::BlobPart;

        let theme_uri = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| Error::InvalidUri(format!("theme URI: {}", e)))?;

        let content_type = "application/vnd.openxmlformats-officedocument.theme+xml".to_string();
        let theme_part = BlobPart::new(theme_uri.clone(), content_type, xml.into_bytes());

        // Add/replace the theme part
        self.opc.add_part(Box::new(theme_part));

        // Add relationship from document to theme if not exists
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            // Check if theme relationship already exists
            let has_theme_rel = doc_part.rels().iter().any(|rel| {
                rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            });

            if !has_theme_rel {
                doc_part.relate_to(
                    "theme/theme1.xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                );
            }
        }

        Ok(())
    }

    #[allow(unused)] // Kept for future use
    fn update_watermark_headers(&mut self, mutable_doc: &MutableDocument) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        // Get watermark if present
        // Access watermark through a temporary reference
        let has_watermark = mutable_doc.has_watermark();
        if !has_watermark {
            return Ok(());
        }

        // Get user header content if it exists
        let user_header_content = if mutable_doc.has_header() {
            mutable_doc.generate_header_xml()?
        } else {
            None
        };

        // Create three headers (default, first, even) with watermark
        let header_types = [
            ("/word/header1.xml", "default"),
            ("/word/header2.xml", "first"),
            ("/word/header3.xml", "even"),
        ];

        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        for (idx, (header_path, _header_type)) in header_types.iter().enumerate() {
            // Generate watermark XML for this header - need to get watermark again each iteration
            let watermark_xml = if let Some(wm) = mutable_doc.watermark.as_ref() {
                wm.to_header_xml((idx + 1) as u32)?
            } else {
                continue;
            };

            // Merge user header content with watermark for the default header
            let header_xml = if idx == 0
                && let Some(ref user_content) = user_header_content
            {
                // Extract user paragraphs from the <w:hdr>...</w:hdr> wrapper
                let user_paragraphs = if let Some(start) = user_content.find("<w:p") {
                    if let Some(end) = user_content.rfind("</w:hdr>") {
                        &user_content[start..end]
                    } else {
                        ""
                    }
                } else {
                    ""
                };

                // Combine watermark and user content
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}{}</w:hdr>"#,
                    watermark_xml, user_paragraphs
                )
            } else {
                // Just watermark for first and even headers
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}</w:hdr>"#,
                    watermark_xml
                )
            };

            let header_uri = PackURI::new(*header_path)
                .map_err(|e| Error::InvalidUri(format!("header URI: {}", e)))?;

            let header_part = BlobPart::new(
                header_uri,
                ct::WML_HEADER.to_string(),
                header_xml.into_bytes(),
            );

            self.opc.add_part(Box::new(header_part));

            // Add relationship from document to header (use relative path)
            // Extract filename from the absolute path (e.g., "/word/header1.xml" -> "header1.xml")
            let header_filename = header_path.rsplit('/').next().unwrap_or(header_path);
            if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
                doc_part.relate_to(header_filename, rt::HEADER);
            }
        }

        Ok(())
    }
}

pub(super) fn settings_part_from_snapshot(
    snapshot: &SettingsPartSnapshot,
    xml: Vec<u8>,
    omitted_relationship_id: Option<&str>,
) -> BlobPart {
    use litchi_opc::part::{BlobPart, Part};

    let mut part = BlobPart::new(snapshot.target.clone(), snapshot.content_type.clone(), xml);
    for relationship in &snapshot.relationships {
        if omitted_relationship_id == Some(relationship.id.as_str()) {
            continue;
        }
        part.rels_mut().add_relationship(
            relationship.reltype.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            relationship.external,
        );
    }
    part
}

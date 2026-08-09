//! Typed DOCX package state and semantic document-facing operations.

pub(super) use crate::Variables;
pub(super) use crate::alt::{Chunk, Conformance, Import, MAX_CHUNKS, Rel, is_relationship};
pub(super) use crate::bibliography::{
    BibliographySource, SourceStore, discover_bibliography_source_stores,
};
pub(super) use crate::custom_xml::{Binding, NewStore};
pub(super) use crate::document::Document;
#[cfg(feature = "encryption")]
pub(super) use crate::encryption::{Limits, Mode};
/// Package implementation for Word documents.
pub(super) use crate::error::{Error, Result};
pub(super) use crate::mail_merge::{
    self, Recipients, RelationshipId, Settings as MailMergeSettings, Source, Target,
    is_mail_merge_relationship_type, map_docx_error,
};
pub(super) use crate::parts::DocumentPart;
pub(super) use crate::settings::{
    ATTACHED_TEMPLATE_RELATIONSHIP, AttachedTemplate, DocumentSettings, extract_document_variables,
    patch_attached_template, patch_document_variables, patch_mail_merge,
    validate_attached_template_target,
};
#[cfg(feature = "vba-inspection")]
pub(super) use crate::vba_project::{
    VbaProject, VbaSupplementalData, discover_vba_project, matching_vba_project,
    remove_vba_project as clear_vba_graph_from_document,
    store_vba_project as store_vba_project_in_document,
};
pub(super) use crate::writer::MutableDocument;
pub(super) use crate::{font, glossary, web as docx_web};
pub(super) use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL,
};
pub(super) use litchi_ooxml_common::custom::{Host as CustomPropsHost, Props as CustomProps};
pub(super) use litchi_ooxml_common::custom_xml::{
    self, Item as CustomXmlItem, MAX_ITEMS, NewItem as NewCustomXmlItem,
    NewProps as NewCustomXmlProps, Props as CustomXmlProps,
};
pub(super) use litchi_ooxml_common::embedded;
pub(super) use litchi_ooxml_common::properties::{Props, Slot};
pub(super) use litchi_ooxml_common::ribbon;
pub(super) use litchi_ooxml_common::web;
pub(super) use litchi_opc::OpcPackage;
pub(super) use litchi_opc::constants::content_type as ct;
pub(super) use litchi_opc::packuri::PackURI;
pub(super) use litchi_opc::part::{BlobPart, Part};
pub(super) use litchi_opc::rel::TargetMode;
pub(super) use std::io::{Read, Seek, Write};
pub(super) use std::path::Path;

pub(super) const MAX_MAIL_MERGE_RELATIONSHIPS: usize = 65_536;

pub(super) fn validate_document_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Ok(());
    }

    Err(Error::InvalidContentType {
        expected: format!(
            "{}, {}, {}, or {}",
            ct::WML_DOCUMENT_MAIN,
            ct::WML_TEMPLATE_MAIN,
            ct::WML_DOCUMENT_MACRO_MAIN,
            ct::WML_TEMPLATE_MACRO_MAIN,
        ),
        got: content_type.to_string(),
    })
}

/// A WordprocessingML (.docx, .docm, or .dotm) package.
///
/// This is the main entry point for working with Word documents.
/// It wraps an OPC package and provides Word-specific functionality.
///
/// # Examples
///
/// ## Reading an existing document
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// // Open an existing document
/// let pkg = Package::open("document.docx")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
///
/// ## Creating a new document
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// // Create a new document
/// let mut pkg = Package::new()?;
/// let mut doc = pkg.document_mut()?;
///
/// // Add content
/// doc.add_paragraph_with_text("Hello, World!");
/// doc.add_heading("Chapter 1", 1)?;
///
/// // Save the document
/// pkg.save("output.docx")?;
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    pub(super) opc: OpcPackage,
    /// Mutable document for writing (cached)
    pub(super) mutable_doc: Option<MutableDocument>,
    /// Whether a committed raw edit has disabled the legacy document writer.
    pub(super) raw_edit_committed: bool,
    /// Authoritative, mutation-tracked core properties.
    pub(super) properties: Slot,
    /// Custom document properties
    pub(super) custom_props: CustomProps,
    /// Whether the custom-property facade has unmaterialized changes.
    pub(super) custom_props_dirty: bool,
    /// Optional managed font-publication policy.
    #[cfg(feature = "automatic-fonts")]
    pub(super) font_embedding: Option<litchi_fonts::embedding::Mode>,
    /// Encryption profile of the opened outer package, retained to prevent an
    /// accidental plaintext downgrade on save.
    #[cfg(feature = "encryption")]
    pub(super) source_encryption: Option<Mode>,
}

pub(super) struct StoredRelationship {
    pub(super) reltype: String,
    pub(super) target: String,
    pub(super) id: String,
    pub(super) external: bool,
}

pub(super) struct SettingsPartSnapshot {
    pub(super) document_uri: PackURI,
    pub(super) target: PackURI,
    pub(super) relationship_exists: bool,
    pub(super) content_type: String,
    pub(super) xml: Vec<u8>,
    pub(super) relationships: Vec<StoredRelationship>,
}

/// Owns the DOCX state that is unpublished until the sink accepts the package.
///
/// The OPC snapshot is structural: built-in part payloads remain shared through
/// their `Arc` allocations. The mutable document stays owned by this guard while
/// materialization runs, so an error or unwind cannot drop the retryable writer.
pub(super) struct WriteRollbackGuard {
    package_before: OpcPackage,
    mutable_doc: Option<MutableDocument>,
}

impl WriteRollbackGuard {
    pub(super) fn new(package: &mut Package) -> Self {
        Self {
            package_before: package.opc.clone(),
            mutable_doc: package.mutable_doc.take(),
        }
    }

    pub(super) fn mutable_doc_mut(&mut self) -> Option<&mut MutableDocument> {
        self.mutable_doc.as_mut()
    }

    pub(super) fn publish(self, package: &mut Package) {
        package.mutable_doc = self.mutable_doc;
    }

    pub(super) fn rollback(self, package: &mut Package) {
        package.opc = self.package_before;
        package.mutable_doc = self.mutable_doc;
    }
}

#[cfg(feature = "automatic-fonts")]
use litchi_core::id::generate_guid_bytes;
#[cfg(feature = "automatic-fonts")]
use litchi_fonts::CollectGlyphs;

#[cfg(feature = "automatic-fonts")]
impl Package {
    pub(super) fn embed_fonts(&mut self) -> Result<()> {
        let Some(mode) = self.font_embedding else {
            return Ok(());
        };
        let glyphs = {
            let document = self.mutable_doc.as_ref().ok_or(Error::UnsafeEdit {
                format: "DOCX",
                operation: "embed_fonts",
                reason: "font discovery is unavailable until the document has a complete mutable model",
            })?;
            if !document.glyphs_are_complete() {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "embed_fonts",
                    reason: "the mutable document preserves unscanned source XML; embedding could omit fonts or subset away live glyphs",
                });
            }
            document.collect_glyphs()
        };
        self.embed_fonts_with_glyphs(glyphs, mode)
    }
}

#[cfg(feature = "automatic-fonts")]
impl Package {
    pub(super) fn embed_fonts_for_document(&mut self, document: &MutableDocument) -> Result<()> {
        let Some(mode) = self.font_embedding else {
            return Ok(());
        };
        if !document.glyphs_are_complete() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "embed_fonts",
                reason: "the mutable document preserves unscanned source XML; embedding could omit fonts or subset away live glyphs",
            });
        }
        self.embed_fonts_with_glyphs(document.collect_glyphs(), mode)
    }

    pub(super) fn embed_fonts_with_glyphs(
        &mut self,
        glyphs: litchi_fonts::GlyphMap,
        mode: litchi_fonts::embedding::Mode,
    ) -> Result<()> {
        let prepared = litchi_fonts::embedding::prepare(glyphs, mode)
            .map_err(|error| Error::Other(error.to_string()))?;
        if prepared.is_empty() {
            return Ok(());
        }
        let subsetted = prepared.iter().any(|font| font.subsetted);
        let conformance = word_font_conformance(&self.opc)?;
        let mut staged = self.opc.clone();
        let mut table = font::read(&staged)?.unwrap_or_default();
        let mut fonts_changed = false;
        for prepared in prepared {
            fonts_changed |= merge_word_font(&mut table, prepared, conformance)?;
        }
        if fonts_changed {
            let _ = font::put(&mut staged, table, conformance)?;
        }
        let settings_changed = ensure_word_font_settings(&mut staged, conformance, subsetted)?;
        if fonts_changed || settings_changed {
            self.opc = staged;
        }
        Ok(())
    }
}

#[cfg(feature = "automatic-fonts")]
fn merge_word_font(
    table: &mut font::Table,
    prepared: litchi_fonts::Prepared,
    conformance: font::Conformance,
) -> Result<bool> {
    let litchi_fonts::Prepared {
        name,
        data,
        style,
        properties,
        subsetted,
    } = prepared;
    let current = table.get(name.as_str())?.cloned();
    let mut next = match &current {
        Some(font) => font.clone(),
        None => font::Font::new(&name)?,
    };
    next = next
        .with_panose(properties.panose().into_bytes())
        .with_family(word_family(properties.family()))
        .with_pitch(word_pitch(properties.pitch()))
        .with_signature(word_signature(properties.signature()));
    let charset = properties
        .charset()
        .map(|charset| font::Charset::from_legacy(charset.code()))
        .filter(|charset| {
            conformance == font::Conformance::Transitional || charset.strict_name().is_some()
        });
    let _ = next.set_charset(charset);

    let style = word_style(style);
    if !word_face_matches(&next, style, &data, subsetted)? {
        let key = font::FontKey::new(generate_guid_bytes());
        let mut obfuscated = data;
        font::obfuscate(&mut obfuscated, key)?;
        let mut embedded = font::Embed::new(style, key, font::Resource::new(obfuscated)?);
        if subsetted {
            embedded = embedded.with_subset(true);
        }
        let _ = next.put(embedded)?;
    }

    if current.as_ref() == Some(&next) {
        return Ok(false);
    }
    match current {
        Some(_) => {
            let replaced = table.replace(name.as_str(), next)?;
            if replaced.is_none() {
                return Err(Error::InvalidFormat(format!(
                    "font '{name}' disappeared during replacement"
                )));
            }
        },
        None => table.add(next)?,
    }
    Ok(true)
}

#[cfg(feature = "automatic-fonts")]
fn word_face_matches(
    font: &font::Font,
    style: font::Style,
    clear: &[u8],
    subsetted: bool,
) -> Result<bool> {
    let Some(embedded) = font
        .embeds()
        .iter()
        .find(|embedded| embedded.style() == style)
    else {
        return Ok(false);
    };
    let subset_matches = if subsetted {
        embedded.subsetted() == Some(true)
    } else {
        embedded.subsetted() != Some(true)
    };
    if !subset_matches {
        return Ok(false);
    }
    let (Some(key), Some(resource)) = (embedded.key(), embedded.resource()) else {
        return Ok(false);
    };
    let encoded = resource.bytes();
    if encoded.len() != clear.len() {
        return Ok(false);
    }
    let (Some(encoded_prefix), Some(clear_prefix)) = (encoded.get(..32), clear.get(..32)) else {
        return Ok(false);
    };
    let mut decoded_prefix = [0; 32];
    decoded_prefix.copy_from_slice(encoded_prefix);
    font::deobfuscate(&mut decoded_prefix, key)?;
    Ok(decoded_prefix == clear_prefix && encoded.get(32..) == clear.get(32..))
}

#[cfg(feature = "automatic-fonts")]
fn word_style(value: litchi_fonts::Style) -> font::Style {
    match value {
        litchi_fonts::Style::Regular => font::Style::Regular,
        litchi_fonts::Style::Bold => font::Style::Bold,
        litchi_fonts::Style::Italic => font::Style::Italic,
        litchi_fonts::Style::BoldItalic => font::Style::BoldItalic,
    }
}

#[cfg(feature = "automatic-fonts")]
fn word_family(value: litchi_fonts::Family) -> font::Family {
    match value {
        litchi_fonts::Family::Auto => font::Family::Auto,
        litchi_fonts::Family::Roman => font::Family::Roman,
        litchi_fonts::Family::Swiss => font::Family::Swiss,
        litchi_fonts::Family::Modern => font::Family::Modern,
        litchi_fonts::Family::Script => font::Family::Script,
        litchi_fonts::Family::Decorative => font::Family::Decorative,
    }
}

#[cfg(feature = "automatic-fonts")]
fn word_pitch(value: litchi_fonts::Pitch) -> font::Pitch {
    match value {
        litchi_fonts::Pitch::Default => font::Pitch::Default,
        litchi_fonts::Pitch::Fixed => font::Pitch::Fixed,
        litchi_fonts::Pitch::Variable => font::Pitch::Variable,
    }
}

#[cfg(feature = "automatic-fonts")]
fn word_signature(value: litchi_fonts::Signature) -> font::Signature {
    font::Signature::new(*value.unicode(), *value.code_pages())
}

#[cfg(feature = "automatic-fonts")]
fn word_font_conformance(package: &OpcPackage) -> Result<font::Conformance> {
    const STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::OFFICE_DOCUMENT | STRICT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| Error::InvalidFormat("main-document relationship is missing".into()))?;
    if relationships.next().is_some() {
        return Err(Error::InvalidFormat(
            "package has multiple main-document relationships".into(),
        ));
    }
    Ok(if relationship.reltype() == STRICT {
        font::Conformance::Strict
    } else {
        font::Conformance::Transitional
    })
}

#[cfg(feature = "automatic-fonts")]
fn ensure_word_font_settings(
    package: &mut OpcPackage,
    conformance: font::Conformance,
    subsetted: bool,
) -> Result<bool> {
    const STRICT_SETTINGS: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
    const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

    let (document_uri, target, exists) = {
        let document = package.main_document_part()?;
        let document_uri = document.partname().clone();
        let mut relationships = document.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::SETTINGS | STRICT_SETTINGS
            )
        });
        let relationship = relationships.next();
        if relationships.next().is_some() {
            return Err(Error::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let target = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(Error::InvalidFormat(
                    "settings relationship cannot be external".into(),
                ));
            },
            Some(relationship) => {
                let expected = match conformance {
                    font::Conformance::Transitional => {
                        litchi_opc::constants::relationship_type::SETTINGS
                    },
                    font::Conformance::Strict => STRICT_SETTINGS,
                };
                if relationship.reltype() != expected {
                    return Err(Error::InvalidFormat(
                        "settings relationship uses the wrong conformance namespace".into(),
                    ));
                }
                relationship.target_partname()?
            },
            None => PackURI::new("/word/settings.xml")
                .map_err(|error| Error::InvalidUri(error.to_string()))?,
        };
        (document_uri, target, relationship.is_some())
    };

    let original = match package.get_part(&target) {
        Ok(part) if exists => {
            if part.content_type() != ct::WML_SETTINGS {
                return Err(Error::InvalidFormat(format!(
                    "settings part has content type {:?}, expected {:?}",
                    part.content_type(),
                    ct::WML_SETTINGS
                )));
            }
            DocumentSettings::extract_from_part(part)?;
            part.blob().to_vec()
        },
        Ok(_) => {
            return Err(Error::InvalidFormat(format!(
                "unowned settings part collision at '{target}'"
            )));
        },
        Err(_) if exists => {
            return Err(Error::PartNotFound(format!("settings part {target}")));
        },
        Err(_) => {
            let word = match conformance {
                font::Conformance::Transitional => TRANSITIONAL_WORD,
                font::Conformance::Strict => STRICT_WORD,
            };
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{word}"/>"#
            )
            .into_bytes()
        },
    };
    let updated = crate::settings::patch_font_embedding(&original, subsetted)?;
    if updated == original {
        return Ok(false);
    }

    package.unsign();
    if exists {
        let part = package.get_part(&target)?;
        let mut checked = part.clone_part();
        checked.set_blob(updated.clone());
        DocumentSettings::extract_from_part(&*checked)?;
        package.get_part_mut(&target)?.set_blob(updated);
    } else {
        let part = BlobPart::new(target.clone(), ct::WML_SETTINGS.to_owned(), updated);
        DocumentSettings::extract_from_part(&part)?;
        package.try_add_part(Box::new(part))?;
        let relationship_type = match conformance {
            font::Conformance::Transitional => litchi_opc::constants::relationship_type::SETTINGS,
            font::Conformance::Strict => STRICT_SETTINGS,
        };
        let target_ref = target.relative_ref(document_uri.base_uri());
        package
            .get_part_mut(&document_uri)?
            .relate_to(&target_ref, relationship_type);
    }
    Ok(true)
}

impl Package {
    /// Get the main document for reading.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, styles, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn document(&self) -> Result<Document<'_>> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| Error::PartNotFound(format!("main document part: {}", e)))?;

        // Create DocumentPart wrapper
        let doc_part = DocumentPart::from_part(main_part)?;

        // Create and return Document with reference to OPC package
        Ok(Document::new(doc_part, &self.opc))
    }

    /// Discover the attached MS-OFFMACRO2 VBA project without inspecting its payloads.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, decompress, or execute the binary
    /// VBA project or Word supplemental-data bytes.
    #[cfg(feature = "vba-inspection")]
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let document = self.opc.main_document_part()?;
        discover_vba_project(&self.opc, document)
    }

    /// Attach a cache-free, inert MS-OVBA project with empty Word supplemental data.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        let payload = project.finish(&litchi_vba::Limits::default())?;
        self.put_vba_managed("set_vba", payload, &VbaSupplementalData::new())
    }

    /// Attach a cache-free project and typed Word document-event/macro metadata.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        supplemental_data: &VbaSupplementalData,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        let payload = project.finish(limits)?;
        self.put_vba_managed("set_vba_with", payload, supplemental_data)
    }

    /// Attach a prevalidated `vbaProject.bin` and typed Word supplemental data.
    #[cfg(feature = "vba-inspection")]
    pub fn put_vba(
        &mut self,
        payload: litchi_vba::Payload,
        supplemental_data: &VbaSupplementalData,
    ) -> Result<VbaProject> {
        self.put_vba_managed("put_vba", payload, supplemental_data)
    }

    /// Remove the VBA project and supplemental-data graph and restore DOCX/DOTX type.
    #[cfg(feature = "vba-inspection")]
    pub fn clear_vba(&mut self) -> Result<bool> {
        let source = self.opc.main_document_part()?.partname().clone();
        let source_part = self.opc.get_part(&source)?;
        if discover_vba_project(&self.opc, source_part)?.is_none() {
            return Ok(false);
        }
        self.edit_semantic_opc("clear_vba", move |candidate| {
            let source = candidate.main_document_part()?.partname().clone();
            clear_vba_graph_from_document(candidate, &source)
        })
    }

    #[cfg(feature = "vba-inspection")]
    fn put_vba_managed(
        &mut self,
        operation: &'static str,
        payload: litchi_vba::Payload,
        supplemental_data: &VbaSupplementalData,
    ) -> Result<VbaProject> {
        let source = self.opc.main_document_part()?.partname().clone();
        let supplemental_xml = supplemental_data.to_xml()?;
        if let Some(project) =
            matching_vba_project(&self.opc, &source, payload.bytes(), &supplemental_xml)?
        {
            return Ok(project);
        }
        let payload = std::sync::Arc::new(payload.into_bytes());
        self.edit_semantic_opc(operation, move |candidate| {
            let source = candidate.main_document_part()?.partname().clone();
            store_vba_project_in_document(candidate, &source, payload, supplemental_xml)
        })
    }

    /// Load inert persisted Office Add-in task panes.
    pub fn task_panes(&self) -> Result<Option<web::Panes>> {
        Ok(web::load(&self.opc)?)
    }

    /// Store a validated task-pane graph by moving it into package ownership.
    pub fn put_task_panes(
        &mut self,
        panes: web::Panes,
        conformance: web::Conformance,
    ) -> Result<&mut Self> {
        web::put(&mut self.opc, panes, conformance)?;
        Ok(self)
    }

    /// Remove task panes and graph resources no longer shared elsewhere.
    pub fn remove_task_panes(&mut self) -> Result<bool> {
        Ok(web::remove(&mut self.opc)?)
    }

    /// Read the fixed legacy and modern package-level Ribbon slots.
    ///
    /// [`ribbon::Set::effective`] applies modern-first precedence. XML remains
    /// inert; callbacks and commands are never invoked.
    pub fn ribbon(&self) -> Result<ribbon::Set<'_>> {
        Ok(ribbon::load(&self.opc)?)
    }

    /// Store opaque Ribbon XML by moving its `Vec` into package ownership.
    pub fn put_ribbon(&mut self, version: ribbon::Version, xml: Vec<u8>) -> Result<&mut Self> {
        ribbon::put(&mut self.opc, version, xml)?;
        Ok(self)
    }

    /// Remove one package-level Ribbon relationship family and its orphaned part.
    pub fn remove_ribbon(&mut self, family: ribbon::Family) -> Result<bool> {
        Ok(ribbon::remove(&mut self.opc, family)?)
    }

    /// Read typed font metadata and inert embedded-font resources.
    pub fn fonts(&self) -> Result<Option<font::Table>> {
        Ok(font::read(&self.opc)?)
    }

    /// Move a complete font table into this package.
    pub fn put_fonts(
        &mut self,
        table: font::Table,
        conformance: font::Conformance,
    ) -> Result<bool> {
        Ok(font::put(&mut self.opc, table, conformance)?)
    }

    /// Remove the font table and font resources that become unreferenced.
    pub fn remove_fonts(&mut self) -> Result<bool> {
        Ok(font::remove(&mut self.opc)?)
    }

    /// Select the font publication policy used by managed save operations.
    #[cfg(feature = "automatic-fonts")]
    pub fn set_font_embedding(&mut self, mode: litchi_fonts::embedding::Mode) -> Result<&mut Self> {
        if self.mutable_doc.is_none() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "set_font_embedding",
                reason: "font discovery requires a complete mutable document model",
            });
        }
        self.font_embedding = Some(mode);
        Ok(self)
    }

    /// Select the font publication policy and return this package by value.
    #[cfg(feature = "automatic-fonts")]
    pub fn with_font_embedding(mut self, mode: litchi_fonts::embedding::Mode) -> Result<Self> {
        self.set_font_embedding(mode)?;
        Ok(self)
    }

    /// Disable managed font publication for subsequent saves.
    #[cfg(feature = "automatic-fonts")]
    pub fn clear_font_embedding(&mut self) -> &mut Self {
        self.font_embedding = None;
        self
    }
}

impl Package {
    /// Borrows the document core properties, retaining package absence.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let props = pkg.props();
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn props(&self) -> Option<&Props> {
        self.properties.get()
    }

    /// Mutably borrows existing core properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// if let Some(props) = pkg.props_mut() {
    ///     props.title = Some("My Document".to_string());
    /// }
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn props_mut(&mut self) -> Option<&mut Props> {
        self.properties.get_mut()
    }

    /// Moves a present core-properties value into the package facade.
    pub fn put_props(&mut self, props: Props) -> Option<Props> {
        self.properties.put(props)
    }

    /// Marks core properties absent and moves out the previous value.
    pub fn clear_props(&mut self) -> Option<Props> {
        self.properties.clear()
    }

    /// Get a reference to the custom document properties.
    ///
    /// Custom properties allow you to attach arbitrary typed metadata to documents.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let custom = pkg.custom_props();
    ///
    /// if let Some(value) = custom.get("ProjectName") {
    ///     println!("Project: {:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn custom_props(&self) -> &CustomProps {
        &self.custom_props
    }

    /// Get a mutable reference to the custom document properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    /// use litchi_ooxml_common::custom::Value;
    ///
    /// let mut pkg = Package::new()?;
    /// let custom = pkg.custom_props_mut();
    ///
    /// custom.insert("ProjectName", "MyProject")?;
    /// custom.insert("Version", Value::I32(1))?;
    /// custom.insert("Budget", Value::F64(50_000.0))?;
    ///
    /// pkg.save("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn custom_props_mut(&mut self) -> &mut CustomProps {
        self.custom_props_dirty = true;
        &mut self.custom_props
    }
}

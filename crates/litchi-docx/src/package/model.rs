#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![cfg_attr(
    any(feature = "automatic-fonts", feature = "vba-inspection"),
    expect(
        clippy::shadow_reuse,
        reason = "parser bindings are intentionally refined after validation"
    )
)]
#![cfg_attr(
    feature = "vba-inspection",
    expect(
        clippy::shadow_unrelated,
        reason = "local parser names mirror the OOXML role currently being decoded"
    )
)]
//! Typed DOCX package state and semantic document-facing operations.

pub(super) use crate::Variables;
pub(super) use crate::alt::{Chunk, Conformance, Import, MAX_CHUNKS, Rel, is_relationship};
pub(super) use crate::bibliography::{
    BibliographySource, SourceStore, discover_bibliography_source_stores,
};
pub(super) use crate::builder::Document as DraftDocument;
pub(super) use crate::custom_xml::{Binding, NewStore};
pub(super) use crate::document::Document;
#[cfg(feature = "encryption")]
pub(super) use crate::encryption::Mode;
/// Package implementation for Word documents.
pub(super) use crate::error::{Error, Result};
pub(super) use crate::mail_merge::{
    self, Recipients, RelationshipId, Settings as MailMergeSettings, Source, Target,
    is_mail_merge_relationship_type, map_docx_error,
};
pub(super) use crate::parts::{DocumentPart, ParagraphIndexCache};
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
#[cfg(test)]
pub(super) use crate::web_extensions as web;
pub(super) use crate::{font, glossary, web as docx_web};
pub(super) use litchi_core::Metadata as Props;
pub(super) use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL,
};
pub(super) use litchi_ooxml_common::custom::{Host as CustomPropsHost, Props as CustomProps};
pub(super) use litchi_ooxml_common::custom_xml::{
    self, Item as CustomXmlItem, MAX_ITEMS, NewItem as NewCustomXmlItem,
    NewProps as NewCustomXmlProps, Props as CustomXmlProps,
};
pub(super) use litchi_ooxml_common::embedded;
pub(super) use litchi_ooxml_common::properties::Slot;
pub(super) use litchi_ooxml_common::ribbon;
pub(super) use litchi_opc::OpcPackage;
pub(super) use litchi_opc::constants::content_type as ct;
pub(super) use litchi_opc::packuri::PackURI;
pub(super) use litchi_opc::part::{BlobPart, Part};
pub(super) use litchi_opc::rel::TargetMode;
pub(super) use std::io::{Read, Write};
pub(super) use std::path::Path;

pub(super) const MAX_MAIL_MERGE_RELATIONSHIPS: usize = 65_536;

pub(crate) fn validate_document_main_content_type(content_type: &str) -> Result<()> {
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

/// A `WordprocessingML` (.docx, .docm, or .dotm) package.
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
/// Cloning shares the physical graph, metadata and retained writer state.
/// Mutating a clone detaches that state before writing; the other handles
/// and their owned document projections retain their original contents.
///
/// ## Creating a new document
///
/// Fresh content is authored with [`crate::Builder`]; an opened package
/// changes only through [`Package::edit`] and its scoped `edit_*` families.
///
/// ```rust,no_run
/// use litchi_docx::Builder;
///
/// // Author a new document
/// let mut builder = Builder::new()?;
/// let doc = builder.document()?;
///
/// // Add content
/// doc.add_paragraph_with_text("Hello, World!");
/// doc.add_heading("Chapter 1", 1)?;
///
/// // Finalize and save the document
/// let mut pkg = builder.build()?;
/// pkg.save("output.docx")?;
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
#[derive(Clone)]
pub struct Package {
    /// The underlying OPC package
    pub(super) opc: SharedOpc,
    /// Package-scoped clean paragraph memo; entries are invalidated by source
    /// allocation identity when the main relationship or payload changes.
    pub(super) paragraph_index_cache: std::sync::Arc<ParagraphIndexCache>,
    /// The main-document payload the markup-compatibility admission last
    /// proved for a transaction. Payloads are replaced, never mutated in
    /// place, so a live allocation that is still this one holds the same
    /// bytes; the weak reference keeps nothing alive.
    pub(super) main_admission: std::sync::Arc<std::sync::Mutex<std::sync::Weak<Vec<u8>>>>,
    /// Authoritative, mutation-tracked core properties.
    pub(super) properties: Shared<Slot>,
    /// Custom document properties
    pub(super) custom_props: Shared<CustomProps>,
    /// The semantic read budgets the package was opened or built with.
    pub(super) semantic_limits: std::sync::Arc<crate::SemanticLimits>,
    /// Encryption profile of the opened outer package, retained to prevent an
    /// accidental plaintext downgrade on save.
    #[cfg(feature = "encryption")]
    pub(super) source_encryption: litchi_ooxml_common::package_encryption::PackageEncryption,
    /// Outer encryption resource policy retained for every managed output.
    #[cfg(feature = "encryption")]
    pub(super) encryption_limits: crate::encryption::Limits,
}

/// Shared immutable state; legacy mutations detach before writing.
pub(super) struct Shared<T>(std::sync::Arc<T>);

impl<T> Clone for Shared<T> {
    fn clone(&self) -> Self {
        Self(std::sync::Arc::clone(&self.0))
    }
}

impl<T> From<T> for Shared<T> {
    fn from(value: T) -> Self {
        Self(std::sync::Arc::new(value))
    }
}

impl<T> std::ops::Deref for Shared<T> {
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl<T: Clone> std::ops::DerefMut for Shared<T> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        std::sync::Arc::make_mut(&mut self.0)
    }
}

/// The OPC graph's output memo belongs with that graph, rather than in a
/// package-wide one-entry table. Keeping the weak key and digest inline makes
/// package clones independent without allocating an invalidation counter for
/// every shared metadata field.
///
/// A clone copies the memo for the same immutable graph. `Arc::make_mut`
/// dissociates any live weak key before an in-place mutation, so a changed
/// branch cannot reuse its old digest while the original branch keeps its
/// memo. A failed or no-op mutable borrow may therefore recompute on that
/// branch, but it cannot publish a stale fingerprint.
pub(super) struct SharedOpc {
    value: std::sync::Arc<OpcPackage>,
    fingerprint: std::sync::Mutex<Option<(std::sync::Weak<OpcPackage>, [u8; 32])>>,
}

impl Clone for SharedOpc {
    fn clone(&self) -> Self {
        Self {
            value: std::sync::Arc::clone(&self.value),
            fingerprint: std::sync::Mutex::new(
                self.fingerprint
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .clone(),
            ),
        }
    }
}

impl SharedOpc {
    pub(super) fn shared(&self) -> std::sync::Arc<OpcPackage> {
        std::sync::Arc::clone(&self.value)
    }

    pub(super) fn fingerprint(&self) -> Option<[u8; 32]> {
        let current = std::sync::Arc::downgrade(&self.value);
        self.fingerprint
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .as_ref()
            .filter(|(cached, _)| cached.strong_count() != 0 && cached.ptr_eq(&current))
            .map(|(_, fingerprint)| *fingerprint)
    }

    pub(super) fn remember_fingerprint(&self, fingerprint: [u8; 32]) {
        *self
            .fingerprint
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner) =
            Some((std::sync::Arc::downgrade(&self.value), fingerprint));
    }
}

impl From<OpcPackage> for SharedOpc {
    fn from(value: OpcPackage) -> Self {
        Self {
            value: std::sync::Arc::new(value),
            fingerprint: std::sync::Mutex::new(None),
        }
    }
}

impl std::ops::Deref for SharedOpc {
    type Target = OpcPackage;

    fn deref(&self) -> &Self::Target {
        &self.value
    }
}

impl std::ops::DerefMut for SharedOpc {
    fn deref_mut(&mut self) -> &mut Self::Target {
        // `Arc::make_mut` dissociates live weak references before returning a
        // mutable graph, which fences the old digest without a generation
        // counter or a package-wide invalidation allocation.
        std::sync::Arc::make_mut(&mut self.value)
    }
}

pub(super) struct StoredRelationship {
    pub(super) reltype: String,
    pub(super) target: String,
    pub(super) id: String,
    pub(super) external: bool,
}

pub(crate) struct SettingsPartSnapshot {
    pub(super) document_uri: PackURI,
    pub(super) target: PackURI,
    pub(super) relationship_exists: bool,
    pub(super) content_type: String,
    pub(crate) xml: Vec<u8>,
    pub(super) relationships: Vec<StoredRelationship>,
}

#[cfg(feature = "automatic-fonts")]
use litchi_core::id::generate_guid_bytes;
#[cfg(feature = "automatic-fonts")]
use litchi_fonts::CollectGlyphs;

#[cfg(feature = "automatic-fonts")]
fn map_font_preparation_error(error: litchi_fonts::FontError) -> Error {
    match error {
        litchi_fonts::FontError::LimitExceeded { limit, actual, .. } => Error::FontLimit {
            kind: font::LimitKind::FontPreparation,
            actual,
            maximum: limit,
        },
        error => Error::Other(error.to_string()),
    }
}

#[cfg(feature = "automatic-fonts")]
impl Package {
    pub(super) fn embed_fonts_for_document(
        &mut self,
        document: &DraftDocument,
        mode: litchi_fonts::embedding::Mode,
    ) -> Result<()> {
        if !document.glyphs_are_complete() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "embed_fonts",
                "the mutable document preserves unscanned source XML; embedding could omit fonts or subset away live glyphs",
            ));
        }
        self.embed_fonts_with_glyphs(document.collect_glyphs(), mode)
    }

    pub(super) fn embed_fonts_with_glyphs(
        &mut self,
        glyphs: litchi_fonts::GlyphMap,
        mode: litchi_fonts::embedding::Mode,
    ) -> Result<()> {
        let prepared =
            litchi_fonts::embedding::prepare(glyphs, mode, self.semantic_limits.fonts_value())
                .map_err(map_font_preparation_error)?;
        if prepared.is_empty() {
            return Ok(());
        }
        let subsetted = prepared.iter().any(|font| font.subsetted);
        let conformance = word_font_conformance(&self.opc)?;
        let mut staged = self.opc.clone();
        let mut table = font::package::read(&staged)?.unwrap_or_default();
        let mut fonts_changed = false;
        for prepared in prepared {
            fonts_changed |= merge_word_font(&mut table, prepared, conformance)?;
        }
        if fonts_changed {
            let _ = font::package::put(&mut staged, table, conformance)?;
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
        let mut embedded = font::Embed::new(
            style,
            key,
            font::Resource::new(obfuscated, &font::Limits::DEFAULT)?,
        );
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
                .map_err(|error| Error::InvalidUri(error.clone()))?,
        };
        (document_uri, target, relationship.is_some())
    };

    // Presence comes from metadata alone (ADR 0030): only an owned settings
    // part is decoded, and its decode failure is reported as one rather than
    // as a missing part or as room for a new one.
    let original = match (package.part_metadata(&target).is_some(), exists) {
        (true, true) => {
            let part = package.get_part(&target)?;
            if part.content_type() != ct::WML_SETTINGS {
                return Err(Error::InvalidFormat(format!(
                    "settings part has content type {:?}, expected {:?}",
                    part.content_type(),
                    ct::WML_SETTINGS
                )));
            }
            let limits = mail_merge::mce_limits_for_source(&package.read_limits());
            DocumentSettings::extract_from_part_with_mce_limits(part, &limits)?;
            part.blob().to_vec()
        },
        (true, false) => {
            return Err(Error::InvalidFormat(format!(
                "unowned settings part collision at '{target}'"
            )));
        },
        (false, true) => {
            return Err(Error::PartNotFound(format!("settings part {target}")));
        },
        (false, false) => {
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

    crate::signature_policy::ensure_unsigned(package)?;
    if exists {
        let part = package.get_part(&target)?;
        let mut checked = part.clone_part();
        checked.set_blob(updated.clone());
        let limits = mail_merge::mce_limits_for_source(&package.read_limits());
        DocumentSettings::extract_from_part_with_mce_limits(&*checked, &limits)?;
        package.get_part_mut(&target)?.set_blob(updated);
    } else {
        let part = BlobPart::new(target.clone(), ct::WML_SETTINGS.to_owned(), updated);
        let limits = mail_merge::mce_limits_for_source(&package.read_limits());
        DocumentSettings::extract_from_part_with_mce_limits(&part, &limits)?;
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
    /// Get an owned, read-only projection of the current main document.
    ///
    /// The projection retains its physical snapshot and may outlive this
    /// package. Later package edits do not change its content or relationships.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn document(&self) -> Result<Document> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| Error::PartNotFound(format!("main document part: {e}")))?;

        // Create DocumentPart wrapper
        let doc_part = DocumentPart::from_owned_part_with_cache(
            main_part.clone_part(),
            std::sync::Arc::clone(&self.paragraph_index_cache),
        )?;

        // Everything repaired while the document was read, in discovery
        // order. The report starts with the packaging and metadata repairs
        // made at open; recording sites then build their own report at the
        // codec boundary and merge it into this aggregation with
        // `ToleranceReport::append`, the open-time aggregation idiom
        // litchi-doc established. See `crate::leniency`.
        let tolerance = self.tolerance_report();

        // Retain the same immutable graph used to select the main part.
        Ok(Document::new(
            doc_part,
            self.opc.shared(),
            tolerance,
            std::sync::Arc::clone(&self.semantic_limits),
        ))
    }

    /// Release the package's clean document memos: the visible main-document
    /// projection, the paragraph offsets and the body block index that no
    /// live document view or value still holds.
    ///
    /// Views keep what they hold; the next [`Self::document`] rebuilds what
    /// was released. Results never depend on what the memos hold.
    pub fn trim_document_memos(&self) {
        self.paragraph_index_cache.trim_clean_all();
    }

    /// A snapshot of the repairs made while this package was opened, in
    /// discovery order: the OPC packaging repairs
    /// ([`Defect::Package`](crate::Defect::Package)) followed by the core-
    /// and custom-properties repairs
    /// ([`Defect::Properties`](crate::Defect::Properties)), which the owning
    /// package reads at open.
    ///
    /// [`Package::document`] starts each document's report from this one.
    /// An empty report means the package and its metadata conformed; edits
    /// never add to it.
    #[must_use]
    pub fn tolerance_report(&self) -> crate::leniency::ToleranceReport {
        let mut report = crate::leniency::ToleranceReport::default();
        report.append_package(self.opc.tolerance_report());
        report.append_properties(self.properties.tolerance_report());
        report.append_properties(self.custom_props.tolerance_report());
        report
    }

    /// Discover the attached MS-OFFMACRO2 VBA project without inspecting its payloads.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, decompress, or execute the binary
    /// VBA project or Word supplemental-data bytes.
    #[cfg(feature = "vba-inspection")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let document = self.opc.main_document_part()?;
        discover_vba_project(&self.opc, document)
    }

    /// Attach a cache-free, inert MS-OVBA project with empty Word supplemental data.
    #[cfg(feature = "vba-inspection")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        let payload = project.finish(&litchi_vba::Limits::default())?;
        self.put_vba_managed("set_vba", payload, &VbaSupplementalData::new())
    }

    /// Attach a cache-free project and typed Word document-event/macro metadata.
    #[cfg(feature = "vba-inspection")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn set_vba_with(
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_vba(
        &mut self,
        payload: litchi_vba::Payload,
        supplemental_data: &VbaSupplementalData,
    ) -> Result<VbaProject> {
        self.put_vba_managed("put_vba", payload, supplemental_data)
    }

    /// Remove the VBA project and supplemental-data graph and restore DOCX/DOTX type.
    #[cfg(feature = "vba-inspection")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn clear_vba(&mut self) -> Result<bool> {
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
        let payload = payload.into_bytes();
        self.edit_semantic_opc(operation, move |candidate| {
            let source = candidate.main_document_part()?.partname().clone();
            store_vba_project_in_document(candidate, &source, payload, supplemental_xml)
        })
    }

    /// Read the fixed legacy and modern package-level Ribbon slots.
    ///
    /// [`ribbon::Set::effective`] applies modern-first precedence. XML remains
    /// inert; callbacks and commands are never invoked.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn ribbon(&self) -> Result<ribbon::Set<'_>> {
        Ok(ribbon::load(&self.opc, &ribbon::Limits::standard())?)
    }

    /// Store opaque Ribbon XML by moving its `Vec` into package ownership.
    ///
    /// The version and bytes the family already stores publish nothing, so a
    /// signed package keeps its signature byte for byte. Any other store
    /// changes the package, so a signed package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] before
    /// anything is changed (ADR 0006): remove the signature explicitly first
    /// with [`Self::unsign`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_ribbon(
        &mut self,
        version: ribbon::Version,
        xml: Vec<u8>,
    ) -> Result<&mut Self> {
        let mut edit = ribbon::Edit::new(&self.opc)?;
        edit.put(version, xml)?;
        self.opc = edit.commit()?.into_package().into();
        Ok(self)
    }

    /// Remove one package-level Ribbon relationship family and its orphaned part.
    ///
    /// An absent family leaves the package untouched and returns `false`,
    /// signed or not. Removing a family from a signed package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] (ADR 0006):
    /// remove the signature explicitly first with [`Self::unsign`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_ribbon(&mut self, family: ribbon::Family) -> Result<bool> {
        let mut edit = ribbon::Edit::new(&self.opc)?;
        let changed = edit.remove(family)?;
        self.opc = edit.commit()?.into_package().into();
        Ok(changed)
    }

    /// Read typed font metadata and inert embedded-font resources.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn fonts(&self) -> Result<Option<font::Table>> {
        font::package::read(&self.opc)
    }

    /// Move a complete font table into this package.
    ///
    /// The signature policy is [`font::raw::put`]'s: a change to a signed package
    /// is refused until [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_fonts(
        &mut self,
        table: font::Table,
        conformance: font::Conformance,
    ) -> Result<bool> {
        font::package::put(&mut self.opc, table, conformance)
    }

    /// Remove the font table and font resources that become unreferenced.
    ///
    /// The signature policy is [`font::raw::remove`]'s: removing a table from a
    /// signed package is refused until [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_fonts(&mut self) -> Result<bool> {
        font::package::remove(&mut self.opc)
    }
}

impl Package {
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
    #[must_use]
    pub fn custom_props(&self) -> &CustomProps {
        &self.custom_props
    }
}

#[cfg(all(test, feature = "automatic-fonts"))]
mod font_settings_tests {
    use litchi_fonts::FontError;
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    const SETTINGS_PAYLOAD: &[u8] =
        br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

    #[test]
    fn automatic_font_limits_remain_typed_at_the_docx_boundary() {
        let error = map_font_preparation_error(FontError::LimitExceeded {
            resource: "font subsetting output bytes",
            limit: 7,
            actual: 8,
        });
        assert!(matches!(
            error,
            Error::FontLimit {
                kind: font::LimitKind::FontPreparation,
                actual: 8,
                maximum: 7,
            }
        ));
    }

    /// A lazily opened package with a stored `/word/settings.xml`, owned by a
    /// settings relationship or not, and optionally corrupted so that only a
    /// decode of its payload fails (ADR 0030).
    fn package(settings: Option<(bool, bool)>) -> OpcPackage {
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{}"/><Override PartName="/word/settings.xml" ContentType="{}"/></Types>"#,
            ct::WML_DOCUMENT_MAIN,
            ct::WML_SETTINGS
        );
        let owned = settings.is_some_and(|(owned, _)| owned);
        let document_relationships = if owned {
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#
        } else {
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#
        };
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        for (name, payload) in [
            ("[Content_Types].xml", content_types.as_bytes()),
            (
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
            (
                "word/document.xml",
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#,
            ),
            ("word/_rels/document.xml.rels", document_relationships.as_bytes()),
        ] {
            writer.write_stored(name, payload).expect("archive member");
        }
        if settings.is_some() {
            writer
                .write_stored("word/settings.xml", SETTINGS_PAYLOAD)
                .expect("settings");
        }
        let mut bytes = writer.finish_to_bytes().expect("archive");
        if settings.is_some_and(|(_, corrupt)| corrupt) {
            let offset = bytes
                .windows(SETTINGS_PAYLOAD.len())
                .position(|window| window == SETTINGS_PAYLOAD)
                .expect("stored settings payload");
            bytes[offset] ^= 1;
        }
        OpcPackage::from_vec(bytes, litchi_opc::SourceLimits::default()).expect("deferred open")
    }

    // Found after the 0759 review: every settings lookup error read as
    // absence, so a present settings part that failed to decode was reported
    // as missing, or was taken for no settings part at all.
    #[test]
    fn font_settings_tell_absent_settings_from_undecodable_ones() {
        let settings = PackURI::new("/word/settings.xml").expect("settings name");
        let transitional = font::Conformance::Transitional;

        // An unowned settings part is a collision whether or not it decodes,
        // and the collision is found without decoding it.
        for corrupt in [false, true] {
            let mut unowned = package(Some((false, corrupt)));
            let result = ensure_word_font_settings(&mut unowned, transitional, false);
            assert!(
                matches!(
                    &result,
                    Err(Error::InvalidFormat(message))
                        if message.contains("unowned settings part collision")
                ),
                "corrupt = {corrupt}: {result:?}"
            );
            assert!(
                !unowned
                    .part_metadata(&settings)
                    .expect("settings stays present")
                    .payload_is_decoded()
            );
        }

        // An owned settings part that fails to decode reports the decode
        // failure, not a missing part.
        let mut owned = package(Some((true, true)));
        let result = ensure_word_font_settings(&mut owned, transitional, false);
        assert!(
            matches!(&result, Err(Error::Opc(litchi_opc::OpcError::ZipError(_)))),
            "{result:?}"
        );

        // Only a truly absent settings part is created.
        let mut absent = package(None);
        assert!(ensure_word_font_settings(&mut absent, transitional, false).expect("created"));
        assert_eq!(
            absent
                .get_part(&settings)
                .expect("created settings")
                .content_type(),
            ct::WML_SETTINGS
        );
    }

    /// `plain` with an OPC signature origin part and relationship, written
    /// out so the signature is an ingressed source signature.
    fn signed_source(plain: &OpcPackage) -> Vec<u8> {
        let mut package = OpcPackage::new();
        for relationship in plain.rels().iter() {
            package
                .rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("package relationship");
        }
        let names: Vec<_> = plain
            .iter_parts()
            .map(|part| part.partname().clone())
            .collect();
        for name in names {
            let part = plain.get_part(&name).expect("part");
            let mut copy = BlobPart::new(
                name.clone(),
                part.content_type().to_owned(),
                part.blob().to_vec(),
            );
            for relationship in part.rels().iter() {
                copy.rels_mut()
                    .try_add_relationship(
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.r_id().to_owned(),
                        relationship.target_mode(),
                    )
                    .expect("part relationship");
            }
            package.try_add_part(Box::new(copy)).expect("part copy");
        }
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").expect("origin name"),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                Vec::new(),
            )))
            .expect("origin part");
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        litchi_opc::PackageWriter::to_bytes(&package).expect("signed source")
    }

    // Record 0014b: the managed font settings used to remove a signature
    // before patching the settings part. A change to a signed package is
    // refused instead; an unchanged one keeps the signature byte for byte.
    #[test]
    fn font_settings_refuse_a_signed_package_until_it_is_unsigned() {
        let transitional = font::Conformance::Transitional;
        for name in ["hello-world-signed.docx", "ms-office-2010-signed.docx"] {
            let original = std::fs::read(format!(
                "{}/../../test-data/poi/test-data/xmldsign/{name}",
                env!("CARGO_MANIFEST_DIR")
            ))
            .expect("signed fixture");
            let mut signed =
                OpcPackage::from_vec(original.clone(), litchi_opc::SourceLimits::default())
                    .expect("signed open");
            let result = ensure_word_font_settings(&mut signed, transitional, true);
            assert!(
                matches!(
                    result,
                    Err(Error::Opc(
                        litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
                    ))
                ),
                "{name}: {result:?}"
            );
            assert_eq!(
                litchi_opc::PackageWriter::to_bytes(&signed).expect("unchanged"),
                original,
                "{name}"
            );

            signed.unsign();
            assert!(ensure_word_font_settings(&mut signed, transitional, true).expect("patched"));
            assert!(!signed.is_signed());

            // Settings that already carry the flags publish nothing, so a
            // signed package keeps its exact source.
            let resigned = signed_source(&signed);
            let mut package =
                OpcPackage::from_vec(resigned.clone(), litchi_opc::SourceLimits::default())
                    .expect("resigned open");
            assert!(package.is_signed());
            assert!(!ensure_word_font_settings(&mut package, transitional, true).expect("no-op"));
            assert_eq!(
                litchi_opc::PackageWriter::to_bytes(&package).expect("exact"),
                resigned,
                "{name}"
            );
        }
    }
}

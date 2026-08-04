use crate::docx::bibliography::{
    BibliographySource, BibliographySourceStore, discover_bibliography_source_stores,
};
use crate::docx::content_control::ContentControl;
use crate::docx::custom_xml::{Binding, NewStore};
use crate::docx::document::Document;
use crate::docx::mail_merge::{
    self, Recipients, Settings, Source, Target, is_mail_merge_relationship_type, map_docx_error,
};
use crate::docx::parts::DocumentPart;
use crate::docx::settings::{
    ATTACHED_TEMPLATE_RELATIONSHIP, AttachedTemplate, DocumentSettings, extract_document_variables,
    patch_attached_template, patch_document_variables, patch_mail_merge,
    validate_attached_template_target,
};
use crate::docx::vba_project::{
    VbaProject, VbaSupplementalData, discover_vba_project,
    remove_vba_project as clear_vba_graph_from_document,
    store_vba_project as store_vba_project_in_document,
};
use crate::docx::writer::MutableDocument;
#[cfg(feature = "encryption")]
use crate::encryption::{Limits, Mode};
/// Package implementation for Word documents.
use crate::error::{OoxmlError, Result};
use litchi_docx::DocumentVariables;
use litchi_docx::alt::{Chunk, Conformance, Import, MAX_CHUNKS, Rel, is_relationship};
use litchi_docx::{font, glossary, web as docx_web};
use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL,
};
use litchi_ooxml_common::custom::Props as CustomProps;
use litchi_ooxml_common::custom_xml::{
    self, Item as CustomXmlItem, MAX_ITEMS, NewItem as NewCustomXmlItem,
    NewProps as NewCustomXmlProps, Props as CustomXmlProps,
};
use litchi_ooxml_common::embedded;
use litchi_ooxml_common::properties::{Props, Slot};
use litchi_ooxml_common::ribbon;
use litchi_ooxml_common::web;
use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::rel::TargetMode;
use std::io::{Read, Seek, Write};
use std::path::Path;

const MAX_MAIL_MERGE_RELATIONSHIPS: usize = 65_536;

fn validate_document_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Ok(());
    }

    Err(OoxmlError::InvalidContentType {
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
/// use litchi_ooxml::docx::Package;
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
/// use litchi_ooxml::docx::Package;
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
    opc: OpcPackage,
    /// Mutable document for writing (cached)
    mutable_doc: Option<MutableDocument>,
    /// Whether a committed raw edit has disabled the legacy document writer.
    raw_edit_committed: bool,
    /// Authoritative, mutation-tracked core properties.
    properties: Slot,
    /// Custom document properties
    custom_props: CustomProps,
    /// Whether the custom-property facade has unmaterialized changes.
    custom_props_dirty: bool,
    /// Encryption profile of the opened outer package, retained to prevent an
    /// accidental plaintext downgrade on save.
    #[cfg(feature = "encryption")]
    source_encryption: Option<Mode>,
}

struct StoredRelationship {
    reltype: String,
    target: String,
    id: String,
    external: bool,
}

struct SettingsPartSnapshot {
    document_uri: PackURI,
    target: PackURI,
    relationship_exists: bool,
    content_type: String,
    xml: Vec<u8>,
    relationships: Vec<StoredRelationship>,
}

/// Owns the DOCX state that is unpublished until the sink accepts the package.
///
/// The OPC snapshot is structural: built-in part payloads remain shared through
/// their `Arc` allocations. The mutable document stays owned by this guard while
/// materialization runs, so an error or unwind cannot drop the retryable writer.
struct WriteRollbackGuard {
    package_before: OpcPackage,
    mutable_doc: Option<MutableDocument>,
}

impl WriteRollbackGuard {
    fn new(package: &mut Package) -> Self {
        Self {
            package_before: package.opc.clone(),
            mutable_doc: package.mutable_doc.take(),
        }
    }

    fn mutable_doc_mut(&mut self) -> Option<&mut MutableDocument> {
        self.mutable_doc.as_mut()
    }

    fn publish(self, package: &mut Package) {
        package.mutable_doc = self.mutable_doc;
    }

    fn rollback(self, package: &mut Package) {
        package.opc = self.package_before;
        package.mutable_doc = self.mutable_doc;
    }
}

#[cfg(feature = "fonts")]
use crate::fonts::{EmbedFonts, PreparedFont, prepare_fonts};
#[cfg(feature = "fonts")]
use litchi_core::id::generate_guid_bytes;
#[cfg(feature = "fonts")]
use litchi_fonts::CollectGlyphs;

#[cfg(feature = "fonts")]
impl EmbedFonts for Package {
    fn embed_fonts(&mut self) -> Result<()> {
        let options = self.opc.save_options().clone();
        let subset = match options.fonts {
            litchi_opc::FontEmbedding::None => return Ok(()),
            litchi_opc::FontEmbedding::Full => false,
            litchi_opc::FontEmbedding::Subset => true,
        };
        let glyphs = {
            let document = self.mutable_doc.as_ref().ok_or(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "embed_fonts",
                reason: "font discovery is unavailable until the document has a complete mutable model",
            })?;
            if !document.glyphs_are_complete() {
                return Err(OoxmlError::UnsafeEdit {
                    format: "DOCX",
                    operation: "embed_fonts",
                    reason: "the mutable document preserves unscanned source XML; embedding could omit fonts or subset away live glyphs",
                });
            }
            document.collect_glyphs()
        };
        self.embed_fonts_with_glyphs(glyphs, subset)
    }
}

#[cfg(feature = "fonts")]
impl Package {
    fn embed_fonts_for_document(&mut self, document: &MutableDocument) -> Result<()> {
        let options = self.opc.save_options().clone();
        let subset = match options.fonts {
            litchi_opc::FontEmbedding::None => return Ok(()),
            litchi_opc::FontEmbedding::Full => false,
            litchi_opc::FontEmbedding::Subset => true,
        };
        if !document.glyphs_are_complete() {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "embed_fonts",
                reason: "the mutable document preserves unscanned source XML; embedding could omit fonts or subset away live glyphs",
            });
        }
        self.embed_fonts_with_glyphs(document.collect_glyphs(), subset)
    }

    fn embed_fonts_with_glyphs(
        &mut self,
        glyphs: litchi_fonts::GlyphMap,
        subset: bool,
    ) -> Result<()> {
        let prepared = prepare_fonts(glyphs, subset)?;
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

#[cfg(feature = "fonts")]
fn merge_word_font(
    table: &mut font::Table,
    prepared: PreparedFont,
    conformance: font::Conformance,
) -> Result<bool> {
    let PreparedFont {
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
                return Err(OoxmlError::InvalidFormat(format!(
                    "font '{name}' disappeared during replacement"
                )));
            }
        },
        None => table.add(next)?,
    }
    Ok(true)
}

#[cfg(feature = "fonts")]
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

#[cfg(feature = "fonts")]
fn word_style(value: litchi_fonts::Style) -> font::Style {
    match value {
        litchi_fonts::Style::Regular => font::Style::Regular,
        litchi_fonts::Style::Bold => font::Style::Bold,
        litchi_fonts::Style::Italic => font::Style::Italic,
        litchi_fonts::Style::BoldItalic => font::Style::BoldItalic,
    }
}

#[cfg(feature = "fonts")]
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

#[cfg(feature = "fonts")]
fn word_pitch(value: litchi_fonts::Pitch) -> font::Pitch {
    match value {
        litchi_fonts::Pitch::Default => font::Pitch::Default,
        litchi_fonts::Pitch::Fixed => font::Pitch::Fixed,
        litchi_fonts::Pitch::Variable => font::Pitch::Variable,
    }
}

#[cfg(feature = "fonts")]
fn word_signature(value: litchi_fonts::Signature) -> font::Signature {
    font::Signature::new(*value.unicode(), *value.code_pages())
}

#[cfg(feature = "fonts")]
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
        .ok_or_else(|| OoxmlError::InvalidFormat("main-document relationship is missing".into()))?;
    if relationships.next().is_some() {
        return Err(OoxmlError::InvalidFormat(
            "package has multiple main-document relationships".into(),
        ));
    }
    Ok(if relationship.reltype() == STRICT {
        font::Conformance::Strict
    } else {
        font::Conformance::Transitional
    })
}

#[cfg(feature = "fonts")]
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
            return Err(OoxmlError::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let target = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(OoxmlError::InvalidFormat(
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
                    return Err(OoxmlError::InvalidFormat(
                        "settings relationship uses the wrong conformance namespace".into(),
                    ));
                }
                relationship.target_partname()?
            },
            None => PackURI::new("/word/settings.xml")
                .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?,
        };
        (document_uri, target, relationship.is_some())
    };

    let original = match package.get_part(&target) {
        Ok(part) if exists => {
            if part.content_type() != ct::WML_SETTINGS {
                return Err(OoxmlError::InvalidFormat(format!(
                    "settings part has content type {:?}, expected {:?}",
                    part.content_type(),
                    ct::WML_SETTINGS
                )));
            }
            DocumentSettings::extract_from_part(part)?;
            part.blob().to_vec()
        },
        Ok(_) => {
            return Err(OoxmlError::InvalidFormat(format!(
                "unowned settings part collision at '{target}'"
            )));
        },
        Err(_) if exists => {
            return Err(OoxmlError::PartNotFound(format!("settings part {target}")));
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
    let updated = crate::docx::settings::patch_font_embedding(&original, subsetted)?;
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
    /// Create a new empty .docx package.
    ///
    /// Creates a minimal valid Word document with default styles and settings.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Add content to the document...
    /// pkg.save("new_document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::docx::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::packuri::PackURI;
        use litchi_opc::part::BlobPart;

        let mut opc = OpcPackage::new();

        // Create document.xml part
        let doc_partname = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document partname: {}", e)))?;
        let doc_part = BlobPart::new(
            doc_partname.clone(),
            ct::WML_DOCUMENT_MAIN.to_string(),
            template::default_document_xml().as_bytes().to_vec(),
        );

        // Create relationship from package to document (use relative path for package-level rels)
        opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        opc.add_part(Box::new(doc_part));

        // Create styles.xml part with dynamic style generation
        let styles_partname = PackURI::new("/word/styles.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("styles partname: {}", e)))?;

        // Generate default styles dynamically
        use crate::docx::writer::style::{MutableStyle, generate_styles_xml};
        let default_styles = vec![
            MutableStyle::normal(),
            MutableStyle::heading_1(),
            MutableStyle::heading_2(),
            MutableStyle::heading_3(),
            MutableStyle::heading_4(),
            MutableStyle::heading_5(),
            MutableStyle::heading_6(),
            MutableStyle::heading_7(),
            MutableStyle::heading_8(),
            MutableStyle::heading_9(),
            MutableStyle::title(),
            MutableStyle::default_paragraph_font(),
            MutableStyle::toc_heading(),
            MutableStyle::toc1(),
            MutableStyle::toc2(),
            MutableStyle::toc3(),
            MutableStyle::hyperlink(),
            MutableStyle::header(),
            MutableStyle::footer(),
            MutableStyle::footnote_text(),
            MutableStyle::endnote_text(),
        ];
        let styles_xml = generate_styles_xml(&default_styles)?;

        let styles_part = BlobPart::new(
            styles_partname.clone(),
            ct::WML_STYLES.to_string(),
            styles_xml.as_bytes().to_vec(),
        );

        // Add relationship from document to styles (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("styles.xml", rt::STYLES);
        }
        opc.add_part(Box::new(styles_part));

        // Create settings.xml part
        let settings_partname = PackURI::new("/word/settings.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("settings partname: {}", e)))?;
        let settings_part = BlobPart::new(
            settings_partname,
            ct::WML_SETTINGS.to_string(),
            template::default_settings_xml().as_bytes().to_vec(),
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("settings.xml", rt::SETTINGS);
        }
        opc.add_part(Box::new(settings_part));

        // Create fontTable.xml part
        let font_table_partname = PackURI::new("/word/fontTable.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("fontTable partname: {}", e)))?;
        let font_table_part = BlobPart::new(
            font_table_partname,
            ct::WML_FONT_TABLE.to_string(),
            template::default_font_table_xml().as_bytes().to_vec(),
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("fontTable.xml", rt::FONT_TABLE);
        }
        opc.add_part(Box::new(font_table_part));

        // Create webSettings.xml part
        let web_settings_partname = PackURI::new("/word/webSettings.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("webSettings partname: {}", e)))?;
        let web_settings_xml = docx_web::write(
            &docx_web::Settings::default(),
            docx_web::Conformance::Transitional,
        )?;
        let web_settings_part = BlobPart::new(
            web_settings_partname,
            ct::WML_WEB_SETTINGS.to_string(),
            web_settings_xml,
        );

        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("webSettings.xml", rt::WEB_SETTINGS);
        }
        opc.add_part(Box::new(web_settings_part));

        // Create core.xml part (core properties)
        let core_props_partname = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml partname: {}", e)))?;
        let core_props_part = BlobPart::new(
            core_props_partname,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        opc.add_part(Box::new(core_props_part));

        // Create app.xml part (extended properties)
        let app_props_partname = PackURI::new("/docProps/app.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("app.xml partname: {}", e)))?;
        let app_props_part = BlobPart::new(
            app_props_partname,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        opc.add_part(Box::new(app_props_part));

        // Create theme1.xml part
        let theme_partname = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme partname: {}", e)))?;
        let theme_part = BlobPart::new(
            theme_partname,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );

        // Add relationship from document to theme (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("theme/theme1.xml", rt::THEME);
        }
        opc.add_part(Box::new(theme_part));

        // Create numbering.xml part
        let numbering_partname = PackURI::new("/word/numbering.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("numbering partname: {}", e)))?;
        let numbering_part = BlobPart::new(
            numbering_partname,
            ct::WML_NUMBERING.to_string(),
            template::default_numbering_xml().as_bytes().to_vec(),
        );

        // Add relationship from document to numbering (use relative path)
        if let Ok(doc_part) = opc.get_part_mut(&doc_partname) {
            doc_part.relate_to("numbering.xml", rt::NUMBERING);
        }
        opc.add_part(Box::new(numbering_part));

        // Create a mutable document for writing
        let mutable_doc = Some(MutableDocument::new());

        // Initialize document properties
        let properties = Slot::load(&opc)?;

        // Initialize custom properties
        let custom_props = CustomProps::new();

        Ok(Self {
            opc,
            mutable_doc,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Create a new empty macro-free Word template (`.dotx`) package.
    ///
    /// Template packages are the native container for reusable AutoText and
    /// other building blocks authored through [`Self::put_glossary`].
    pub fn new_template() -> Result<Self> {
        let mut package = Self::new()?;
        let main = package.opc.main_document_part()?.partname().clone();
        package
            .opc
            .get_part_mut(&main)?
            .set_content_type(ct::WML_TEMPLATE_MAIN.to_owned())?;
        Ok(package)
    }

    /// Open a .docx, .docm, .dotx, or .dotm package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .docx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        let custom_props = CustomProps::read(&opc)?;
        let properties = Slot::load(&opc)?;

        Ok(Self {
            opc,
            mutable_doc: None,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let file = std::fs::File::open(path.as_ref()).map_err(OoxmlError::Io)?;
        let opened = crate::encryption::load(file, password)?;
        Self::from_opened(opened)
    }

    /// Open with explicit outer-encryption limits.
    ///
    /// The decrypted OPC archive is parsed under its independently bounded
    /// default policy; a composite host policy remains migration work.
    #[cfg(feature = "encryption")]
    pub fn open_with<P: AsRef<Path>>(path: P, password: &str, limits: &Limits) -> Result<Self> {
        let file = std::fs::File::open(path.as_ref()).map_err(OoxmlError::Io)?;
        let opened = crate::encryption::load_with(file, password, limits)?;
        Self::from_opened(opened)
    }

    /// Create a Package from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `opc` - An already-parsed OPC package
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::{OpcPackage, docx::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.docx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        let custom_props = CustomProps::read(&opc)?;
        let properties = Slot::load(&opc)?;

        Ok(Self {
            opc,
            mutable_doc: None,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Create a .docx, .docm, or .dotm package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .docx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("document.docx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a Word document by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

        validate_document_main_content_type(main_part.content_type())?;

        let custom_props = CustomProps::read(&opc)?;
        let properties = Slot::load(&opc)?;

        Ok(Self {
            opc,
            mutable_doc: None,
            raw_edit_committed: false,
            properties,
            custom_props,
            custom_props_dirty: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    #[cfg(feature = "encryption")]
    pub fn from_reader_with_password<R: Read>(reader: R, password: &str) -> Result<Self> {
        let opened = crate::encryption::load(reader, password)?;
        Self::from_opened(opened)
    }

    /// Open a reader with explicit outer-encryption limits.
    ///
    /// The decrypted OPC archive uses its independently bounded defaults.
    #[cfg(feature = "encryption")]
    pub fn from_reader_with<R: Read>(reader: R, password: &str, limits: &Limits) -> Result<Self> {
        let opened = crate::encryption::load_with(reader, password, limits)?;
        Self::from_opened(opened)
    }

    #[cfg(feature = "encryption")]
    fn from_opened(opened: crate::encryption::Opened) -> Result<Self> {
        let source_encryption = opened.mode();
        let opc = OpcPackage::from_vec(opened.into_bytes())?;
        let mut package = Self::from_opc_package(opc)?;
        package.source_encryption = source_encryption;
        Ok(package)
    }

    /// Get the main document for reading.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, styles, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn document(&self) -> Result<Document<'_>> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main document part: {}", e)))?;

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
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let document = self.opc.main_document_part()?;
        discover_vba_project(&self.opc, document)
    }

    /// Attach a cache-free, inert MS-OVBA project with empty Word supplemental data.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        self.set_vba_with(
            project,
            &VbaSupplementalData::new(),
            &litchi_vba::Limits::default(),
        )
    }

    /// Attach a cache-free project and typed Word document-event/macro metadata.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        supplemental_data: &VbaSupplementalData,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        self.put_vba(project.finish(limits)?, supplemental_data)
    }

    /// Attach a prevalidated `vbaProject.bin` and typed Word supplemental data.
    pub fn put_vba(
        &mut self,
        payload: litchi_vba::Payload,
        supplemental_data: &VbaSupplementalData,
    ) -> Result<VbaProject> {
        let source = self.opc.main_document_part()?.partname().clone();
        store_vba_project_in_document(&mut self.opc, &source, payload, supplemental_data)
    }

    /// Remove the VBA project and supplemental-data graph and restore DOCX/DOTX type.
    pub fn clear_vba(&mut self) -> Result<bool> {
        let source = self.opc.main_document_part()?.partname().clone();
        clear_vba_graph_from_document(&mut self.opc, &source)
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
    #[cfg(feature = "fonts")]
    pub fn set_font_embedding(
        &mut self,
        embedding: litchi_opc::FontEmbedding,
    ) -> Result<&mut Self> {
        if embedding != litchi_opc::FontEmbedding::None && self.mutable_doc.is_none() {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "set_font_embedding",
                reason: "font discovery requires a complete mutable document model",
            });
        }

        if self.opc.save_options().fonts != embedding {
            self.opc.with_fonts(embedding);
        }
        Ok(self)
    }

    /// Select the font publication policy and return this package by value.
    #[cfg(feature = "fonts")]
    pub fn with_font_embedding(mut self, embedding: litchi_opc::FontEmbedding) -> Result<Self> {
        self.set_font_embedding(embedding)?;
        Ok(self)
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Return whether this document contains package signatures.
    #[must_use]
    #[inline]
    pub fn is_signed(&self) -> bool {
        self.opc.is_signed()
    }

    /// Verify package signatures with the safe strict policy.
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures_with(policy)
    }

    /// Add a signature while preserving every existing valid signature.
    pub fn sign(&mut self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<PackURI> {
        self.opc.sign(signer)
    }

    /// Add a signature with explicit authoring resource bounds.
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.opc.sign_with(signer, limits)
    }

    /// Atomically replace all signatures with one signature.
    pub fn resign(&mut self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<PackURI> {
        self.opc.resign(signer)
    }

    /// Atomically replace signatures with explicit authoring resource bounds.
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.opc.resign_with(signer, limits)
    }

    /// Remove all package signatures.
    pub fn unsign(&mut self) -> &mut Self {
        self.opc.unsign();
        self
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan_with`] with [`Self::opc_package`] when a lower
    /// layer needs explicitly tuned limits.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.opc)?)
    }

    /// Load the bounded, inert classic-chart graph owned by the main document.
    pub fn chart_graph(&self) -> Result<crate::docx::chart::DocxChartGraph> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::docx::chart::load_chart_graph(&self.opc, &document)
    }

    /// Load the typed, inert SmartArt (DrawingML diagram) inventory anchored
    /// in the main document.
    ///
    /// Each returned [`crate::docx::smartart::DocxSmartArt`] carries the
    /// parsed data-model node tree, the layout/quick-style/colors part
    /// metadata, and the diagram part names. Both transitional and Strict
    /// namespace dialects are supported.
    pub fn smart_arts(&self) -> Result<Vec<crate::docx::smartart::DocxSmartArt>> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::docx::smartart::load_smart_arts(&self.opc, &document)
    }

    /// Load the typed, inert text-box and WordArt inventory anchored in the
    /// main document.
    ///
    /// Each returned [`crate::docx::textbox::DocxTextBox`] carries the shape
    /// identity, the `wps:bodyPr` text-body properties, the story as
    /// paragraphs with runs, and WordArt warp/styling presence flags. Both
    /// DrawingML shapes and legacy VML `w:pict` fallbacks are recognized, in
    /// both the transitional and Strict namespace dialects.
    pub fn text_boxes(&self) -> Result<Vec<crate::docx::textbox::DocxTextBox>> {
        crate::docx::textbox::load_text_boxes(self.opc.main_document_part()?.blob())
    }

    /// Deterministically store an already coherent classic-chart graph.
    pub fn store_chart_graph(&mut self, graph: &crate::docx::chart::DocxChartGraph) -> Result<()> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::docx::chart::store_chart_graph(&mut self.opc, &document, graph)
    }

    /// Transactionally edit the current plaintext OPC graph.
    ///
    /// The closure receives a structural candidate whose built-in part payloads
    /// share immutable `Arc` storage. Returning an error or unwinding leaves
    /// this package's graph unpublished; custom `Part` implementations retain
    /// their own clone and interior-mutability policy. Before a successful
    /// commit, the candidate's Word main relationship, content type, core
    /// properties, and custom properties are validated and facade-owned state
    /// is reloaded. Committing a raw edit disables the legacy document writer
    /// so it cannot later erase the edit.
    pub fn edit_opc<T>(&mut self, edit: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.ensure_opc_current("edit_opc")?;
        if self.opc.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "edit_opc",
                reason: "raw OPC editing cannot honor an automatic font policy; use the managed font facade",
            });
        }

        let mut candidate = self.opc.clone();
        candidate.unsign();
        let value = edit(&mut candidate)?;

        if candidate.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "edit_opc",
                reason: "raw OPC transactions cannot configure automatic font embedding; use the managed font facade",
            });
        }
        let main_part = candidate
            .main_document_part()
            .map_err(|error| OoxmlError::PartNotFound(format!("main document part: {error}")))?;
        validate_document_main_content_type(main_part.content_type())?;
        let properties = Slot::load(&candidate)?;
        let custom_props = CustomProps::read(&candidate)?;

        self.opc = candidate;
        self.properties = properties;
        self.custom_props = custom_props;
        self.custom_props_dirty = false;
        self.mutable_doc = None;
        self.raw_edit_committed = true;
        Ok(value)
    }

    /// Get a mutable document for writing and modification.
    ///
    /// This returns a `MutableDocument` that allows you to add and modify
    /// paragraphs, tables, and other document elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut doc = pkg.document_mut()?;
    ///
    /// // Add content
    /// doc.add_paragraph_with_text("Hello, World!");
    /// let para = doc.add_paragraph();
    /// para.add_run_with_text("Bold text").bold(true);
    ///
    /// // Add a table
    /// let table = doc.add_table(3, 2);
    /// if let Some(cell) = table.cell(0, 0) {
    ///     cell.set_text("Header 1");
    /// }
    ///
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn document_mut(&mut self) -> Result<&mut MutableDocument> {
        if self.raw_edit_committed {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: "document_mut",
                reason: "a raw OPC edit committed; use edit_opc for further low-level changes",
            });
        }

        // If we don't have a mutable document, try to load it from the package
        if self.mutable_doc.is_none() {
            let doc_uri = PackURI::new("/word/document.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

            // Try to get existing document content
            if let Ok(part) = self.opc.get_part(&doc_uri) {
                let xml = std::str::from_utf8(part.blob())
                    .map_err(|e| OoxmlError::InvalidFormat(format!("Invalid UTF-8: {}", e)))?;
                self.mutable_doc = Some(MutableDocument::from_xml(xml)?);
            } else {
                // Create a new empty document
                self.mutable_doc = Some(MutableDocument::new());
            }
        }

        self.mutable_doc.as_mut().ok_or_else(|| {
            OoxmlError::InvalidFormat("mutable document initialization did not complete".into())
        })
    }

    /// Append a package-backed alternative-format import to the document body.
    pub fn add_alt(&mut self, import: Import, match_source: Option<bool>) -> Result<Chunk> {
        let index = self.document_mut()?.alts().len();
        self.insert_alt(index, import, match_source)
    }

    /// Insert a package-backed alternative-format import by anchor-relative index.
    ///
    /// Part, relationship, and body mutations are rolled back together on error.
    pub fn insert_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let count = self.document_mut()?.alts().len();
        if index > count {
            return Err(OoxmlError::InvalidFormat(format!(
                "altChunk index {index} is out of range"
            )));
        }
        if count >= MAX_CHUNKS {
            return Err(OoxmlError::InvalidFormat(format!(
                "alternative-format anchor limit of {MAX_CHUNKS} is exhausted"
            )));
        }
        let namespace = self.alt_chunk_namespace()?;
        let (chunk, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .insert_alt(index, chunk.clone(), namespace)
        {
            self.rollback_alt_chunk_target(chunk.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        Ok(chunk)
    }

    /// Replace an anchor and its relationship as one package mutation.
    pub fn replace_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let old = self
            .document_mut()?
            .alts()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        let namespace = self.alt_chunk_namespace()?;
        let (new, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .replace_alt(index, new.clone(), namespace)
        {
            self.rollback_alt_chunk_target(new.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        self.remove_alt_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Remove an anchor, its relationship, and an unreachable internal payload.
    pub fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        let old = self
            .document_mut()?
            .alts()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        self.document_mut()?.remove_alt(index)?;
        self.remove_alt_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Reorder body anchors without changing their package relationships.
    pub fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        self.document_mut()?.move_alt(from, to)
    }

    fn alt_chunk_namespace(&self) -> Result<Conformance> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let strict = self
            .opc
            .get_part(&document_uri)
            .map(|part| {
                part.blob()
                    .windows(b"http://purl.oclc.org/ooxml/wordprocessingml/main".len())
                    .any(|window| window == b"http://purl.oclc.org/ooxml/wordprocessingml/main")
            })
            .unwrap_or(false);
        Ok(if strict {
            Conformance::Strict
        } else {
            Conformance::Transitional
        })
    }

    fn install_alt_chunk_target(
        &mut self,
        import: Import,
        match_source: Option<bool>,
        namespace: Conformance,
    ) -> Result<(Chunk, Option<PackURI>)> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let document = self.opc.get_part(&document_uri)?;
        let relationship_id = (1usize..=MAX_CHUNKS)
            .map(|number| format!("rIdAltChunk{number}"))
            .find(|id| document.rels().get(id).is_none())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("altChunk relationship ID space is exhausted".into())
            })?;
        let relationship = Rel::new(relationship_id.clone())?;
        let relationship_type = namespace.relationship();
        let (target_ref, target_mode, installed_part) = match import {
            Import::Link(uri) => (uri.into_string(), TargetMode::External, None),
            Import::Data(data) => {
                data.validate()?;
                let media_type = data.media_type();
                let (uri, target_ref) = (1usize..=MAX_CHUNKS)
                    .find_map(|number| {
                        let target_ref = format!("afchunk{number}.{}", data.extension());
                        let uri = PackURI::new(format!("/word/{target_ref}")).ok()?;
                        self.opc
                            .get_part(&uri)
                            .is_err()
                            .then_some((uri, target_ref))
                    })
                    .ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "alternative-format part-name space is exhausted".into(),
                        )
                    })?;
                self.opc.try_add_part(Box::new(BlobPart::new(
                    uri.clone(),
                    media_type.to_string(),
                    data.into_bytes(),
                )))?;
                (target_ref, TargetMode::Internal, Some(uri))
            },
        };
        let relation_result = self
            .opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_string(),
                target_ref,
                relationship_id.clone(),
                target_mode,
            );
        if let Err(error) = relation_result {
            if let Some(uri) = &installed_part {
                self.opc.remove_part(uri);
            }
            return Err(error.into());
        }
        Ok((Chunk::new(relationship, match_source), installed_part))
    }

    fn validate_alt_chunk_relationship(&self, chunk: &Chunk) -> Result<Option<PackURI>> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        let relationship = self
            .opc
            .get_part(&document_uri)?
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(format!(
                    "altChunk relationship {:?} is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if !is_relationship(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "relationship {:?} is not an alternative-format import",
                chunk.relationship().as_str()
            )));
        }
        if relationship.is_external() {
            return Ok(None);
        }
        let target = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidFormat(format!(
                "invalid altChunk relationship {:?}: {error}",
                chunk.relationship().as_str()
            ))
        })?;
        self.opc.get_part(&target).map_err(|_| {
            OoxmlError::InvalidFormat(format!(
                "altChunk relationship {:?} targets a missing part",
                chunk.relationship().as_str()
            ))
        })?;
        Ok(Some(target))
    }

    fn rollback_alt_chunk_target(&mut self, id: &str, part: Option<&PackURI>) -> Result<()> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        self.opc.get_part_mut(&document_uri)?.rels_mut().remove(id);
        if let Some(part) = part {
            self.opc.remove_part(part);
        }
        Ok(())
    }

    fn remove_alt_relationship(&mut self, chunk: &Chunk, target: Option<&PackURI>) -> Result<()> {
        if self.mutable_doc.as_ref().is_some_and(|document| {
            document
                .alts()
                .iter()
                .any(|remaining| remaining.relationship() == chunk.relationship())
        }) {
            return Ok(());
        }
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| OoxmlError::InvalidUri(format!("document URI: {error}")))?;
        self.opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .remove(chunk.relationship().as_str());
        let Some(target) = target else {
            return Ok(());
        };
        let package_reference = self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        });
        let part_reference = self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|part| &part == target)
            })
        });
        if !package_reference && !part_reference {
            self.opc.remove_part(target);
        }
        Ok(())
    }

    /// Discover every validated Custom XML Data Storage relationship occurrence.
    pub fn custom_xml(&self) -> Result<Vec<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?)
    }

    /// Discover typed, inert bibliography source stores from Custom XML.
    ///
    /// Word stores its current bibliography source list in a document Custom
    /// XML data store. This method exposes stored source values and style
    /// metadata only. It never matches source tags to citations, resolves
    /// schemas or styles, runs transforms, refreshes fields, or changes data.
    pub fn bibliography_source_stores(&self) -> Result<Vec<BibliographySourceStore>> {
        let items = custom_xml::discover(&self.opc)?;
        discover_bibliography_source_stores(&items)
    }

    /// Discover typed, inert bibliography sources in package and XML order.
    ///
    /// This flattens [`Self::bibliography_source_stores`] without resolving
    /// `CITATION` fields or applying bibliography style rules.
    pub fn bibliography_sources(&self) -> Result<Vec<BibliographySource>> {
        let stores = self.bibliography_source_stores()?;
        Ok(stores
            .iter()
            .flat_map(|store| store.sources().iter().cloned())
            .collect())
    }

    /// Return the number of typed, inert bibliography sources.
    pub fn bibliography_source_count(&self) -> Result<usize> {
        Ok(self.bibliography_sources()?.len())
    }

    /// Find a Custom XML data store by its case-insensitive datastore item GUID.
    pub fn custom_xml_by_id(&self, id: &str) -> Result<Option<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?.into_iter().find(|item| {
            item.props()
                .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        }))
    }

    /// Add a collision-safe `/customXml/itemN.xml` data store to the main document.
    pub fn add_custom_xml(&mut self, store: NewStore) -> Result<CustomXmlItem> {
        custom_xml::validate_content_type(&store.content_type)?;
        custom_xml::validate_payload(&store.xml)?;
        let props = CustomXmlProps {
            id: store.id,
            schemas: store.schemas,
        };
        custom_xml::validate_props(&props)?;
        let source_part = self.opc.main_document_part()?.partname().clone();
        let source = self.opc.get_part(&source_part)?;
        let rel_id = (1usize..=MAX_ITEMS + 1)
            .map(|number| format!("rIdCustomXml{number}"))
            .find(|id| source.rels().get(id).is_none())
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("Custom XML relationship ID space is exhausted".into())
            })?;
        let mut part_names = None;
        for number in 1usize..=MAX_ITEMS + 1 {
            let data = PackURI::new(format!("/customXml/item{number}.xml"))
                .map_err(OoxmlError::InvalidUri)?;
            let props = PackURI::new(format!("/customXml/itemProps{number}.xml"))
                .map_err(OoxmlError::InvalidUri)?;
            let conflict = self.opc.iter_parts().any(|part| {
                part.partname().as_str().eq_ignore_ascii_case(data.as_str())
                    || part
                        .partname()
                        .as_str()
                        .eq_ignore_ascii_case(props.as_str())
            });
            if !conflict {
                part_names = Some((data, props));
                break;
            }
        }
        let (data_part, props_part) = part_names.ok_or_else(|| {
            OoxmlError::InvalidFormat("Custom XML part-name space is exhausted".into())
        })?;
        custom_xml::add(
            &mut self.opc,
            NewCustomXmlItem {
                source: source_part,
                rel_id,
                part: data_part.clone(),
                content_type: store.content_type,
                xml: store.xml,
                props: Some(NewCustomXmlProps {
                    part: props_part,
                    rel_id: "rIdProps1".to_string(),
                    value: props,
                }),
                conformance: store.conformance,
            },
        )?;
        custom_xml::discover(&self.opc)?
            .into_iter()
            .find(|item| item.part() == &data_part)
            .ok_or_else(|| {
                OoxmlError::InvalidFormat("new Custom XML data store was not discoverable".into())
            })
    }

    /// Replace only the inert XML payload of a data store.
    pub fn set_custom_xml(&mut self, id: &str, xml: Vec<u8>) -> Result<()> {
        custom_xml::validate_payload(&xml)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| OoxmlError::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        self.opc.get_part_mut(item.part())?.set_blob(xml);
        self.opc.unsign();
        Ok(())
    }

    /// Replace payload, content type, schema references, and canonical properties.
    pub fn replace_custom_xml(&mut self, id: &str, replacement: NewStore) -> Result<()> {
        custom_xml::validate_content_type(&replacement.content_type)?;
        custom_xml::validate_payload(&replacement.xml)?;
        if !replacement.id.eq_ignore_ascii_case(id) {
            return Err(OoxmlError::InvalidFormat(
                "replacement itemID must identify the existing data store".into(),
            ));
        }
        let props = CustomXmlProps {
            id: replacement.id,
            schemas: replacement.schemas,
        };
        let props_xml = custom_xml::write_props(&props, replacement.conformance)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| OoxmlError::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        let props_part = item.props_part().cloned().ok_or_else(|| {
            OoxmlError::InvalidFormat("Custom XML data store has no properties part".into())
        })?;
        let existing_relationships = self
            .opc
            .get_part(item.part())?
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.r_id().to_string(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        let mut data_part = BlobPart::new(
            item.part().clone(),
            replacement.content_type,
            replacement.xml,
        );
        for (reltype, target, id, external) in existing_relationships {
            data_part
                .rels_mut()
                .add_relationship(reltype, target, id, external);
        }
        self.opc.add_part(Box::new(data_part));
        self.opc.get_part_mut(&props_part)?.set_blob(props_xml);
        self.opc.unsign();
        Ok(())
    }

    /// Remove a data store unless an SDT still binds to its item GUID.
    pub fn remove_custom_xml(&mut self, id: &str) -> Result<bool> {
        let items = custom_xml::discover(&self.opc)?;
        let matching = items
            .iter()
            .filter(|item| {
                item.props()
                    .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
            })
            .collect::<Vec<_>>();
        let Some(first) = matching.first() else {
            return Ok(false);
        };
        if self
            .custom_xml_bindings()?
            .iter()
            .any(|binding| binding.store_id.eq_ignore_ascii_case(id))
        {
            return Err(OoxmlError::InvalidFormat(format!(
                "Custom XML itemID '{id}' is still referenced by a content control"
            )));
        }
        for item in &matching {
            self.opc
                .get_part_mut(item.source())?
                .rels_mut()
                .remove(item.rel_id());
        }
        let data_part = first.part().clone();
        let props_part = first.props_part().cloned();
        if !self.part_is_referenced(&data_part) {
            self.opc.remove_part(&data_part);
            if let Some(props_part) = props_part
                && !self.part_is_referenced(&props_part)
            {
                self.opc.remove_part(&props_part);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    /// Locate the document's bibliography source store, if one exists.
    ///
    /// Returns the Custom XML item GUID and the store payload. Word keeps a
    /// single current source list; when several stores exist the first in
    /// package order is used and the rest are left untouched.
    fn bibliography_store_item(&self) -> Result<Option<(String, Vec<u8>)>> {
        let stores = self.bibliography_source_stores()?;
        let Some(store) = stores.first() else {
            return Ok(None);
        };
        let item_id = store
            .data_store_item_id()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(
                    "bibliography source store has no Custom XML item GUID".into(),
                )
            })?
            .to_owned();
        let item = self.custom_xml_by_id(&item_id)?.ok_or_else(|| {
            OoxmlError::PartNotFound(format!("bibliography source store item '{item_id}'"))
        })?;
        Ok(Some((item_id, item.xml().to_vec())))
    }

    /// Add a typed bibliography source to the document's source store.
    ///
    /// When no store exists, one is created as a Custom XML data store with
    /// the bibliography namespace registered. Otherwise the source is
    /// appended in place, preserving untouched entries, style metadata, and
    /// the store's relationship/content-type graph. Duplicate tags are
    /// rejected. Returns the Custom XML item GUID of the store.
    pub fn add_bibliography_source(
        &mut self,
        source: crate::docx::bibliography_writer::BibliographySourceBuilder,
    ) -> Result<String> {
        if let Some((item_id, xml)) = self.bibliography_store_item()? {
            let updated = crate::docx::bibliography_writer::add_source_xml(&xml, &source)?;
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
            Ok(item_id)
        } else {
            let xml = crate::docx::bibliography_writer::new_store_xml(&[source])?;
            let item = self.add_custom_xml(NewStore {
                xml: xml.into_bytes(),
                content_type: "application/xml".to_string(),
                id: crate::docx::bibliography_writer::DEFAULT_STORE_ITEM_ID.to_string(),
                schemas: vec![crate::docx::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()],
                conformance: custom_xml::Conformance::Transitional,
            })?;
            item.props().map(|props| props.id.clone()).ok_or_else(|| {
                OoxmlError::InvalidFormat(
                    "new bibliography Custom XML store has no item GUID".into(),
                )
            })
        }
    }

    /// Remove the bibliography source with the given tag from the source
    /// store. Returns whether a source was removed.
    pub fn remove_bibliography_source(&mut self, tag: &str) -> Result<bool> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Ok(false);
        };
        let (updated, removed) = crate::docx::bibliography_writer::remove_source_xml(&xml, tag)?;
        if removed {
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
        }
        Ok(removed)
    }

    /// Replace the bibliography source with the given tag, preserving entry
    /// order and all untouched entries. Fails when the tag does not exist.
    pub fn replace_bibliography_source(
        &mut self,
        tag: &str,
        source: crate::docx::bibliography_writer::BibliographySourceBuilder,
    ) -> Result<()> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Err(OoxmlError::PartNotFound(
                "no bibliography source store exists".to_string(),
            ));
        };
        let updated = crate::docx::bibliography_writer::replace_source_xml(&xml, tag, &source)?;
        self.set_custom_xml(&item_id, updated.into_bytes())?;
        // Re-validate the mutated store through the read side.
        self.bibliography_source_stores()?;
        Ok(())
    }

    /// Reorder main-document data-store relationships by item GUID.
    pub fn order_custom_xml(&mut self, ordered_ids: &[String]) -> Result<()> {
        let source_part = self.opc.main_document_part()?.partname().clone();
        let items = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter(|item| item.source() == &source_part)
            .collect::<Vec<_>>();
        if items.len() != ordered_ids.len() {
            return Err(OoxmlError::InvalidFormat(
                "reorder list must contain every main-document Custom XML item".into(),
            ));
        }
        let mut by_id = std::collections::HashMap::new();
        for item in &items {
            let id = item
                .props()
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat("Custom XML item has no datastore itemID".into())
                })?
                .id
                .to_ascii_lowercase();
            if by_id.insert(id, item).is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "main-document Custom XML items are not uniquely reorderable".into(),
                ));
            }
        }
        let mut ordered = Vec::with_capacity(items.len());
        let mut seen = std::collections::HashSet::new();
        for id in ordered_ids {
            let key = id.to_ascii_lowercase();
            if !seen.insert(key.clone()) {
                return Err(OoxmlError::InvalidFormat("duplicate reorder itemID".into()));
            }
            let item = *by_id.get(&key).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("unknown reorder itemID '{id}'"))
            })?;
            let reltype = self
                .opc
                .get_part(&source_part)?
                .rels()
                .get(item.rel_id())
                .ok_or_else(|| {
                    OoxmlError::InvalidRelationship(format!(
                        "Custom XML relationship '{}' disappeared during reorder",
                        item.rel_id()
                    ))
                })?
                .reltype()
                .to_string();
            ordered.push((item, reltype));
        }
        let source = self.opc.get_part(&source_part)?;
        let reserved = source
            .rels()
            .iter()
            .filter(|relationship| {
                !items
                    .iter()
                    .any(|item| item.rel_id() == relationship.r_id())
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect::<std::collections::HashSet<_>>();
        let ids = (1usize..=MAX_ITEMS + 1)
            .filter_map(|batch| {
                let candidates = (0..ordered.len())
                    .map(|index| format!("rIdCustomXmlOrder{batch:04}_{index:06}"))
                    .collect::<Vec<_>>();
                candidates
                    .iter()
                    .all(|id| !reserved.contains(id))
                    .then_some(candidates)
            })
            .next()
            .ok_or_else(|| {
                OoxmlError::InvalidFormat(
                    "Custom XML reorder relationship ID space is exhausted".into(),
                )
            })?;
        let source = self.opc.get_part_mut(&source_part)?;
        let source_base_uri = source.partname().base_uri().to_string();
        for item in &items {
            source.rels_mut().remove(item.rel_id());
        }
        for ((item, reltype), id) in ordered.into_iter().zip(ids) {
            source.rels_mut().add_relationship(
                reltype,
                item.part().relative_ref(&source_base_uri),
                id,
                false,
            );
        }
        self.opc.unsign();
        Ok(())
    }

    /// Collect and lexically validate SDT bindings from every permitted Word container.
    pub fn custom_xml_bindings(&self) -> Result<Vec<Binding>> {
        let permitted = [
            ct::WML_DOCUMENT_MAIN,
            ct::WML_DOCUMENT_GLOSSARY,
            ct::WML_HEADER,
            ct::WML_FOOTER,
            ct::WML_FOOTNOTES,
            ct::WML_ENDNOTES,
        ];
        let mut bindings = Vec::new();
        for part in self
            .opc
            .iter_parts()
            .filter(|part| permitted.contains(&part.content_type()))
        {
            for control in ContentControl::extract_from_document(part.blob())? {
                control.validate_data_binding()?;
                if let (Some(xpath), Some(store_item_id)) = (
                    control.data_binding_xpath(),
                    control.data_binding_store_item_id(),
                ) {
                    bindings.push(Binding {
                        source: part.partname().clone(),
                        control_id: control.id(),
                        xpath: xpath.to_string(),
                        store_id: store_item_id.to_string(),
                        prefixes: control.data_binding_prefix_mappings().map(str::to_string),
                    });
                }
            }
        }
        bindings.sort_unstable_by(|left, right| {
            left.source
                .as_str()
                .cmp(right.source.as_str())
                .then_with(|| left.control_id.cmp(&right.control_id))
        });
        Ok(bindings)
    }

    /// Validate that every permitted SDT binding resolves to a datastore item GUID.
    pub fn validate_custom_xml_bindings(&self) -> Result<()> {
        let item_ids = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter_map(|item| item.props().map(|props| props.id.to_ascii_lowercase()))
            .collect::<std::collections::HashSet<_>>();
        for binding in self.custom_xml_bindings()? {
            if !item_ids.contains(&binding.store_id.to_ascii_lowercase()) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "content control {} in '{}' references missing Custom XML itemID '{}'",
                    binding.control_id,
                    binding.source.as_str(),
                    binding.store_id
                )));
            }
        }
        Ok(())
    }

    fn part_is_referenced(&self, target: &PackURI) -> bool {
        self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        }) || self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|part| &part == target)
            })
        })
    }

    /// Return the validated inert mail-merge settings, if configured.
    pub fn mail_merge_settings(&self) -> Result<Option<Settings>> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .mail_merge()
            .cloned())
    }

    /// Resolve a mail-merge relationship without opening or fetching its target.
    pub fn mail_merge_target(&self, relationship_id: &str) -> Result<Target> {
        let snapshot = self.settings_part_snapshot()?;
        let part = self.opc.get_part(&snapshot.target)?;
        let relationship = part.rels().get(relationship_id).ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "mail-merge relationship '{relationship_id}' is missing"
            ))
        })?;
        if !is_mail_merge_relationship_type(relationship.reltype()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "relationship '{relationship_id}' is not a mail-merge source"
            )));
        }
        if relationship.is_external() {
            return Ok(Target::External(relationship.target_ref().to_string()));
        }
        let target = relationship.target_partname()?;
        let target_part = self.opc.get_part(&target)?;
        Ok(Target::Internal {
            part_name: target,
            bytes: target_part.blob().to_vec(),
            content_type: target_part.content_type().to_string(),
        })
    }

    /// Set or replace the complete mail-merge graph atomically.
    pub fn set_mail_merge(
        &mut self,
        mut settings: Settings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let mut used_ids = snapshot
            .relationships
            .iter()
            .filter(|relationship| !is_mail_merge_relationship_type(&relationship.reltype))
            .map(|relationship| relationship.id.clone())
            .collect::<std::collections::HashSet<_>>();
        let mut staged_parts = Vec::new();
        let mut staged_relationships = Vec::new();

        let data_id = if let Some(source) = data_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Data",
                "mailMergeSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let header_id = if let Some(source) = header_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Header",
                "mailMergeHeaderSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let recipient_id = if let Some(recipients) = recipients {
            let xml = recipients
                .to_xml(conformance)
                .map_err(map_docx_error)?
                .into_bytes();
            let id = allocate_mail_merge_relationship_id("Recipients", &mut used_ids)?;
            let uri = self.allocate_mail_merge_part_name("recipientData", "xml")?;
            let target = uri.relative_ref(snapshot.target.base_uri());
            staged_parts.push(BlobPart::new(
                uri,
                Recipients::content_type().to_string(),
                xml,
            ));
            staged_relationships.push(StoredRelationship {
                reltype: mail_merge_relationship_type(conformance, "recipientData"),
                target,
                id: id.clone(),
                external: false,
            });
            Some(id)
        } else {
            None
        };
        settings.assign_package_relationships(data_id, header_id, recipient_id);
        let patched = patch_mail_merge(&snapshot.xml, Some(&settings), conformance)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let old_ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in old_ids {
            replacement.rels_mut().remove(&id);
        }
        for relationship in staged_relationships {
            replacement.rels_mut().add_relationship(
                relationship.reltype,
                relationship.target,
                relationship.id,
                relationship.external,
            );
        }
        DocumentSettings::extract_from_part(&replacement)?;

        let mut installed = Vec::new();
        for part in staged_parts {
            let name = part.partname().clone();
            if let Err(error) = self.opc.try_add_part(Box::new(part)) {
                for installed_name in installed {
                    self.opc.remove_part(&installed_name);
                }
                return Err(error.into());
            }
            installed.push(name);
        }
        if let Err(error) = self.commit_settings_part(&snapshot, replacement) {
            for installed_name in installed {
                self.opc.remove_part(&installed_name);
            }
            return Err(error);
        }
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(())
    }

    /// Update settings and sources using the same atomic replacement semantics.
    pub fn update_mail_merge(
        &mut self,
        settings: Settings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            recipients,
            conformance,
        )
    }

    /// Replace recipient inclusion flags while retaining inert source targets and settings.
    pub fn update_mail_merge_recipients(
        &mut self,
        recipients: Recipients,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let settings = self.mail_merge_settings()?.ok_or_else(|| {
            OoxmlError::InvalidFormat("document has no mail-merge settings".into())
        })?;
        let data_source = settings
            .data_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        let header_source = settings
            .header_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            Some(recipients),
            conformance,
        )
    }

    /// Clear mail-merge XML, relationships, and unreachable owned targets.
    pub fn clear_mail_merge(&mut self) -> Result<bool> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        if DocumentSettings::extract_from_part(&original)?
            .mail_merge()
            .is_none()
        {
            return Ok(false);
        }
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let patched = patch_mail_merge(&snapshot.xml, None, mail_merge::Conformance::Transitional)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in ids {
            replacement.rels_mut().remove(&id);
        }
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    fn stage_mail_merge_source(
        &self,
        source: Source,
        label: &str,
        relationship_suffix: &str,
        conformance: mail_merge::Conformance,
        used_ids: &mut std::collections::HashSet<String>,
    ) -> Result<(StoredRelationship, Option<BlobPart>)> {
        let id = allocate_mail_merge_relationship_id(label, used_ids)?;
        let settings_target = self.settings_part_snapshot()?.target;
        match source {
            Source::External(uri) => {
                validate_mail_merge_external_uri(&uri)?;
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target: uri,
                        id,
                        external: true,
                    },
                    None,
                ))
            },
            Source::Internal {
                bytes,
                content_type,
                extension,
            } => {
                validate_mail_merge_internal_source(&bytes, &content_type, &extension)?;
                let uri = self.allocate_mail_merge_part_name(label, &extension)?;
                let target = uri.relative_ref(settings_target.base_uri());
                let part = BlobPart::new(uri, content_type, bytes);
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target,
                        id,
                        external: false,
                    },
                    Some(part),
                ))
            },
        }
    }

    fn allocate_mail_merge_part_name(&self, stem: &str, extension: &str) -> Result<PackURI> {
        for number in 1usize.. {
            let candidate = PackURI::new(format!("/word/mailMerge/{stem}{number}.{extension}"))
                .map_err(OoxmlError::InvalidUri)?;
            if self.opc.iter_parts().all(|part| {
                !part
                    .partname()
                    .as_str()
                    .eq_ignore_ascii_case(candidate.as_str())
            }) {
                return Ok(candidate);
            }
        }
        unreachable!("the mail-merge part-name space is unbounded")
    }

    fn mail_merge_internal_targets(&self, snapshot: &SettingsPartSnapshot) -> Result<Vec<PackURI>> {
        let Ok(part) = self.opc.get_part(&snapshot.target) else {
            return Ok(Vec::new());
        };
        part.rels()
            .iter()
            .filter(|relationship| {
                is_mail_merge_relationship_type(relationship.reltype())
                    && !relationship.is_external()
            })
            .map(|relationship| relationship.target_partname().map_err(Into::into))
            .collect()
    }

    /// Read typed web-output settings and their conformance family.
    pub fn web(&self) -> Result<Option<(docx_web::Settings, docx_web::Conformance)>> {
        Ok(docx_web::load(&self.opc)?)
    }

    /// Move complete web-output settings into package ownership.
    pub fn put_web(
        &mut self,
        settings: docx_web::Settings,
        conformance: docx_web::Conformance,
    ) -> Result<bool> {
        Ok(docx_web::put(&mut self.opc, settings, conformance)?)
    }

    /// Remove the document-owned web-settings part.
    pub fn remove_web(&mut self) -> Result<bool> {
        Ok(docx_web::remove(&mut self.opc)?)
    }

    /// Inspect the external template associated with this document without dereferencing it.
    pub fn attached_template(&self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .attached_template()
            .cloned())
    }

    /// Associate this document with an external template URI.
    ///
    /// The URI is recorded inertly and is never fetched or executed.
    pub fn set_attached_template_uri(
        &mut self,
        target_uri: impl Into<String>,
    ) -> Result<&mut Self> {
        use litchi_opc::part::Part;

        let target_uri = target_uri.into();
        validate_attached_template_target(&target_uri)?;
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let old_id = settings
            .attached_template()
            .map(|template| template.relationship_id().to_owned());

        let mut replacement =
            settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), old_id.as_deref());
        let relationship_id = if let Some(id) = old_id {
            replacement.rels_mut().add_relationship(
                ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
                target_uri,
                id.clone(),
                true,
            );
            id
        } else {
            replacement.relate_to_ext(&target_uri, ATTACHED_TEMPLATE_RELATIONSHIP)
        };
        replacement.set_blob(patch_attached_template(
            &snapshot.xml,
            Some(&relationship_id),
        )?);
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(self)
    }

    /// Remove the attached-template element and its referenced relationship.
    pub fn remove_attached_template(&mut self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let Some(attached_template) = settings.attached_template().cloned() else {
            return Ok(None);
        };
        let replacement = settings_part_from_snapshot(
            &snapshot,
            patch_attached_template(&snapshot.xml, None)?,
            Some(attached_template.relationship_id()),
        );
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(attached_template))
    }

    /// Read the document variables stored in `settings.xml`.
    pub fn document_variables(&self) -> Result<Option<DocumentVariables>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(Some(extract_document_variables(&part)?))
    }

    /// Insert or replace one document variable atomically.
    pub fn set_document_variable(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let previous = variables.insert(name, value)?;
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(previous)
    }

    /// Remove one document variable atomically.
    pub fn remove_document_variable(&mut self, name: &str) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let Some(previous) = variables.remove(name) else {
            return Ok(None);
        };
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(previous))
    }

    /// Remove every document variable atomically and return the number removed.
    pub fn clear_document_variables(&mut self) -> Result<usize> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(0);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let count = variables.count();
        if count == 0 {
            return Ok(0);
        }
        variables.clear();
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(count)
    }

    /// Load the typed glossary/building-block catalog and its dialect.
    pub fn glossary(&self) -> Result<Option<(glossary::Catalog, glossary::Conformance)>> {
        Ok(glossary::load(&self.opc)?)
    }

    /// Move a complete semantic catalog into the package.
    pub fn put_glossary(
        &mut self,
        catalog: glossary::Catalog,
        conformance: glossary::Conformance,
    ) -> Result<bool> {
        Ok(glossary::put(&mut self.opc, catalog, conformance)?)
    }

    /// Load the complete low-level glossary OPC graph without copying payloads.
    pub fn glossary_graph(&self) -> Result<Option<glossary::raw::Graph>> {
        Ok(glossary::raw::load(&self.opc)?)
    }

    /// Publish a complete low-level glossary OPC graph into the package.
    ///
    /// Returns `false` when the graph is already identical, preserving package
    /// bytes and digital signatures.
    pub fn put_glossary_graph(&mut self, graph: &glossary::raw::Graph) -> Result<bool> {
        Ok(glossary::raw::put(&mut self.opc, graph)?)
    }

    /// Remove and return the complete low-level glossary OPC graph.
    pub fn take_glossary_graph(&mut self) -> Result<Option<glossary::raw::Graph>> {
        Ok(glossary::raw::remove(&mut self.opc)?)
    }

    /// Remove the complete glossary-owned graph.
    pub fn remove_glossary(&mut self) -> Result<bool> {
        Ok(glossary::remove(&mut self.opc)?)
    }

    /// Save the package to a file.
    ///
    /// Writes the complete Word document including all parts, relationships,
    /// and content types to a .docx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .docx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.ensure_plain_output("save")?;
        self.save_plain_impl(path)
    }

    /// Explicitly save a plaintext package, even when the source was encrypted.
    pub fn save_plain<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.save_plain_impl(path)
    }

    fn save_plain_impl<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        litchi_opc::atomic::replace_with::<OoxmlError>(path.as_ref(), |temporary| {
            self.write_plain(temporary)
        })
    }

    /// Save the package to a stream.
    ///
    /// Writes the complete Word document including all parts, relationships,
    /// and content types to a writer stream.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write + Seek
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
    /// use std::io::Cursor;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify document...
    /// let mut cursor = Cursor::new(Vec::new());
    /// pkg.to_stream(&mut cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn to_stream<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        self.ensure_plain_output("to_stream")?;
        self.write_plain(writer)
    }

    /// Explicitly write a plaintext package to a stream.
    pub fn to_plain_stream<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        self.write_plain(writer)
    }

    /// Serialize and encrypt this package entirely in memory.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        let mut output = std::io::Cursor::new(Vec::new());
        self.write_plain(&mut output)?;
        crate::encryption::encrypt(output.into_inner(), password, mode).map_err(Into::into)
    }

    /// Serialize and encrypt using the source package's retained profile.
    #[cfg(feature = "encryption")]
    pub fn to_reencrypted(&mut self, password: &str) -> Result<Vec<u8>> {
        let mode = self.preserved_mode("to_reencrypted")?;
        self.to_encrypted(password, mode)
    }

    /// Save with an explicit encryption profile and a borrowed password.
    #[cfg(feature = "encryption")]
    pub fn save_encrypted<P: AsRef<Path>>(
        &mut self,
        path: P,
        password: &str,
        mode: Mode,
    ) -> Result<()> {
        let output = self.to_encrypted(password, mode)?;
        litchi_opc::atomic::replace(path.as_ref(), |temporary| {
            temporary.write_all(&output)?;
            Ok(())
        })?;
        self.source_encryption = Some(mode);
        Ok(())
    }

    /// Save using the encrypted source's retained profile.
    #[cfg(feature = "encryption")]
    pub fn save_reencrypted<P: AsRef<Path>>(&mut self, path: P, password: &str) -> Result<()> {
        let mode = self.preserved_mode("save_reencrypted")?;
        self.save_encrypted(path, password, mode)
    }

    /// Encryption profile of the opened or most recently encrypted package.
    #[cfg(feature = "encryption")]
    pub const fn encryption(&self) -> Option<Mode> {
        self.source_encryption
    }

    fn ensure_opc_current(&self, operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "raw OPC access would expose an encrypted source as plaintext; use the managed encryption or explicit plaintext APIs",
            });
        }

        if self
            .mutable_doc
            .as_ref()
            .is_some_and(MutableDocument::is_modified)
        {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "the legacy document writer has unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }
        if self.properties.is_dirty() {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "core properties have unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }
        if self.custom_props_dirty {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "custom properties have unmaterialized changes; use a managed save or to_plain_stream first",
            });
        }

        Ok(())
    }

    fn ensure_plain_output(&self, _operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(OoxmlError::UnsafeEdit {
                format: "DOCX",
                operation: _operation,
                reason: "the source package was encrypted; use save_reencrypted or save_plain",
            });
        }
        Ok(())
    }

    #[cfg(feature = "encryption")]
    fn preserved_mode(&self, operation: &'static str) -> Result<Mode> {
        self.source_encryption.ok_or(OoxmlError::UnsafeEdit {
            format: "DOCX",
            operation,
            reason: "the source package has no encryption profile; supply an explicit Mode",
        })
    }

    fn write_plain<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        use crate::docx::writer::relmap::RelationshipMapper;
        use litchi_opc::constants::relationship_type as rt;

        // Keep both the source graph and the mutable semantic document
        // available until the complete publication succeeds. Materializing a
        // document rebuilds many related parts; an error half-way through must
        // leave the caller with the same retryable edit rather than a dropped
        // writer and a partially rewritten package.
        let mut rollback = WriteRollbackGuard::new(self);
        // A sink or a late writer hook may unwind instead of returning an
        // error. Catch the unwind long enough to put both owned pieces of the
        // host state back before resuming it; the staged-properties guard is
        // dropped while unwinding the closure and therefore keeps its dirty
        // intent for the next attempt.
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<()> {
            // If we have a mutable document, update the document.xml part
            if let Some(mutable_doc) = rollback.mutable_doc_mut()
                && mutable_doc.is_modified()
            {
                // Generate TOC if configured (must happen before serialization)
                mutable_doc.generate_toc_if_needed()?;

                // Step 1: Collect all content that needs relationships
                let hyperlink_urls = mutable_doc.collect_hyperlink_urls();
                let images = mutable_doc.collect_images();
                let ole_objects = mutable_doc.collect_ole_objects();
                let smart_arts = mutable_doc.collect_smart_arts();
                let has_header = mutable_doc.has_header();
                let has_footer = mutable_doc.has_footer();
                let section_header_footer_parts =
                    mutable_doc.collect_section_header_footer_parts()?;
                let explicit_section_relationships =
                    mutable_doc.collect_explicit_section_header_footer_relationships()?;
                let mut planned_section_parts = Vec::new();
                for (index, (header, part)) in section_header_footer_parts.into_iter().enumerate() {
                    let stem = if header {
                        "headerSection"
                    } else {
                        "footerSection"
                    };
                    let filename = format!("{stem}{}.xml", index + 1);
                    let uri = PackURI::new(format!("/word/{filename}"))
                        .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
                    if self.opc.get_part(&uri).is_ok() {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "section header/footer part {uri} already exists"
                        )));
                    }
                    planned_section_parts.push((header, part, uri, filename));
                }

                // Step 2: Create a relationship mapper and add relationships
                let mut rel_mapper = RelationshipMapper::new();

                // Create the document part first (we'll update it later)
                let doc_uri = PackURI::new("/word/document.xml")
                    .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

                if !explicit_section_relationships.is_empty() {
                    let existing_document = self.opc.get_part(&doc_uri).map_err(|_| {
                        OoxmlError::InvalidFormat(
                            "section references exist without a document part".to_string(),
                        )
                    })?;
                    for (id, header) in &explicit_section_relationships {
                        let relationship = existing_document.rels().get(id).ok_or_else(|| {
                            OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} is missing"
                            ))
                        })?;
                        let expected_type = if *header {
                            [
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                                "http://purl.oclc.org/ooxml/officeDocument/relationships/header",
                            ]
                        } else {
                            [
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                                "http://purl.oclc.org/ooxml/officeDocument/relationships/footer",
                            ]
                        };
                        if relationship.is_external()
                            || !expected_type.contains(&relationship.reltype())
                        {
                            return Err(OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} has the wrong type or target mode"
                            )));
                        }
                        let target = relationship.target_partname().map_err(|error| {
                            OoxmlError::InvalidFormat(format!(
                                "invalid section relationship {id:?}: {error}"
                            ))
                        })?;
                        let part = self.opc.get_part(&target).map_err(|_| {
                            OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} targets a missing part"
                            ))
                        })?;
                        let expected_content_type = if *header {
                            ct::WML_HEADER
                        } else {
                            ct::WML_FOOTER
                        };
                        if part.content_type() != expected_content_type {
                            return Err(OoxmlError::InvalidFormat(format!(
                                "section relationship {id:?} targets the wrong content type"
                            )));
                        }
                    }
                }

                // Get or create the document part to add relationships to
                let content_type = self
                    .opc
                    .get_part(&doc_uri)
                    .map(|p| p.content_type().to_string())
                    .unwrap_or_else(|_| ct::WML_DOCUMENT_MAIN.to_string());

                // Create new temporary part for relationships
                use litchi_opc::part::{BlobPart, Part};
                let mut temp_part =
                    BlobPart::new(doc_uri.clone(), content_type.clone(), Vec::new());

                // Copy existing relationships from the original document part (styles, settings, etc.)
                if let Ok(existing_part) = self.opc.get_part(&doc_uri) {
                    for rel in existing_part.rels().iter() {
                        // Skip relationships we're going to recreate dynamically
                        if !matches!(
                            rel.reltype(),
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                                | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                        ) {
                            temp_part.rels_mut().add_relationship(
                                rel.reltype().to_string(),
                                rel.target_ref().to_string(),
                                rel.r_id().to_string(),
                                rel.is_external(),
                            );
                        }
                    }
                }

                for (header, part, _, filename) in &planned_section_parts {
                    let relationship_type = if *header { rt::HEADER } else { rt::FOOTER };
                    let rid = temp_part.relate_to(filename, relationship_type);
                    rel_mapper.add_section_header_footer_id(part.key.clone(), rid);
                }

                // Add hyperlink relationships (external)
                for (i, url) in hyperlink_urls.iter().enumerate() {
                    let rid = temp_part.relate_to_ext(url, rt::HYPERLINK);
                    rel_mapper.add_hyperlink(i, rid);
                }

                // Add image parts and relationships
                for (i, (image_data, image_format)) in images.iter().enumerate() {
                    let image_num = i + 1;
                    let ext = image_format.extension();
                    let image_partname = format!("/word/media/image{}.{}", image_num, ext);
                    let image_uri = PackURI::new(&image_partname)
                        .map_err(|e| OoxmlError::InvalidUri(format!("image URI: {}", e)))?;

                    // Create and add image part
                    let image_part = BlobPart::new(
                        image_uri,
                        image_format.mime_type().to_string(),
                        image_data.to_vec(),
                    );
                    self.opc.add_part(Box::new(image_part));

                    // Create relationship from document to image
                    let rid = temp_part.relate_to(&image_partname, rt::IMAGE);
                    rel_mapper.add_image(i, rid);
                }

                // Add embedded OLE object parts and relationships. Payloads
                // are stored verbatim as inert binary parts; optional
                // previews are stored as ordinary media parts.
                for (i, object) in ole_objects.iter().enumerate() {
                    let object_num = i + 1;
                    let object_partname = format!("/word/embeddings/oleObject{object_num}.bin");
                    let object_uri = PackURI::new(&object_partname)
                        .map_err(|e| OoxmlError::InvalidUri(format!("OLE object URI: {}", e)))?;
                    self.opc.add_part(Box::new(BlobPart::new(
                        object_uri,
                        ct::OFC_OLE_OBJECT.to_string(),
                        object.payload().to_vec(),
                    )));
                    let rid = temp_part.relate_to(&object_partname, rt::OLE_OBJECT);
                    rel_mapper.add_ole_object(object.shape_id(), rid);

                    if let Some((preview_data, preview_format)) = object.preview() {
                        let preview_partname = format!(
                            "/word/media/oleObjectPreview{object_num}.{}",
                            preview_format.extension()
                        );
                        let preview_uri = PackURI::new(&preview_partname).map_err(|e| {
                            OoxmlError::InvalidUri(format!("OLE preview URI: {}", e))
                        })?;
                        self.opc.add_part(Box::new(BlobPart::new(
                            preview_uri,
                            preview_format.mime_type().to_string(),
                            preview_data.to_vec(),
                        )));
                        let rid = temp_part.relate_to(&preview_partname, rt::IMAGE);
                        rel_mapper.add_ole_preview(object.shape_id(), rid);
                    }
                }

                // Add SmartArt diagram parts (data, layout, quick style,
                // colors) and their relationships. The optional pre-rendered
                // drawing part is not generated; Word and LibreOffice
                // re-render from the layout and data parts.
                let mut diagram_index = 0u32;
                for smartart in &smart_arts {
                    // Allocate non-colliding part names under /word/diagrams/.
                    let (data_name, layout_name, quick_style_name, colors_name) = loop {
                        diagram_index = diagram_index.checked_add(1).ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "SmartArt diagram part name space exhausted".to_string(),
                            )
                        })?;
                        let names = (
                            format!("/word/diagrams/data{diagram_index}.xml"),
                            format!("/word/diagrams/layout{diagram_index}.xml"),
                            format!("/word/diagrams/quickStyle{diagram_index}.xml"),
                            format!("/word/diagrams/colors{diagram_index}.xml"),
                        );
                        let taken = [&names.0, &names.1, &names.2, &names.3].iter().any(|name| {
                            PackURI::new(*name)
                                .map(|uri| self.opc.get_part(&uri).is_ok())
                                .unwrap_or(true)
                        });
                        if !taken {
                            break names;
                        }
                    };
                    let parts = smartart.generate_parts();
                    for (partname, content_type, xml) in [
                        (&data_name, ct::DML_DIAGRAM_DATA, parts.data_xml),
                        (&layout_name, ct::DML_DIAGRAM_LAYOUT, parts.layout_xml),
                        (
                            &quick_style_name,
                            ct::DML_DIAGRAM_STYLE,
                            parts.quick_style_xml,
                        ),
                        (&colors_name, ct::DML_DIAGRAM_COLORS, parts.colors_xml),
                    ] {
                        let uri = PackURI::new(partname)
                            .map_err(|e| OoxmlError::InvalidUri(format!("diagram URI: {}", e)))?;
                        self.opc.add_part(Box::new(BlobPart::new(
                            uri,
                            content_type.to_string(),
                            xml.into_bytes(),
                        )));
                    }
                    let rel_ids = crate::docx::writer::smartart::SmartArtRelIds {
                        data: temp_part.relate_to(&data_name, DIAGRAM_DATA_REL),
                        layout: temp_part.relate_to(&layout_name, DIAGRAM_LAYOUT_REL),
                        quick_style: temp_part
                            .relate_to(&quick_style_name, DIAGRAM_QUICK_STYLE_REL),
                        colors: temp_part.relate_to(&colors_name, DIAGRAM_COLORS_REL),
                    };
                    rel_mapper.add_smart_art(smartart.anchor_key(), rel_ids);
                }

                // Add header/footer parts and relationships
                // Note: If watermark exists, headers will be handled by update_watermark_headers
                // which merges user content with watermark
                if has_header
                    && !mutable_doc.has_watermark()
                    && let Some(header_xml) = mutable_doc.generate_header_xml()?
                {
                    let header_uri = PackURI::new("/word/header1.xml")
                        .map_err(|e| OoxmlError::InvalidUri(format!("header URI: {}", e)))?;
                    let header_part = BlobPart::new(
                        header_uri,
                        ct::WML_HEADER.to_string(),
                        header_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(header_part));
                    // Use relative path for relationship (relative to document.xml location)
                    let rid = temp_part.relate_to("header1.xml", rt::HEADER);
                    rel_mapper.set_header_id(rid);
                }

                if has_footer && let Some(footer_xml) = mutable_doc.generate_footer_xml()? {
                    let footer_uri = PackURI::new("/word/footer1.xml")
                        .map_err(|e| OoxmlError::InvalidUri(format!("footer URI: {}", e)))?;
                    let footer_part = BlobPart::new(
                        footer_uri,
                        ct::WML_FOOTER.to_string(),
                        footer_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(footer_part));
                    // Use relative path for relationship (relative to document.xml location)
                    let rid = temp_part.relate_to("footer1.xml", rt::FOOTER);
                    rel_mapper.set_footer_id(rid);
                }

                // Add footnotes parts and relationships BEFORE document XML generation
                if let Some(footnotes_xml) = mutable_doc.generate_footnotes_xml()? {
                    let footnotes_uri = PackURI::new("/word/footnotes.xml")
                        .map_err(|e| OoxmlError::InvalidUri(format!("footnotes URI: {}", e)))?;
                    let footnotes_part = BlobPart::new(
                        footnotes_uri,
                        ct::WML_FOOTNOTES.to_string(),
                        footnotes_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(footnotes_part));
                    let rid = temp_part.relate_to("footnotes.xml", rt::FOOTNOTES);
                    rel_mapper.set_footnotes_id(rid);
                }

                // Add endnotes parts and relationships BEFORE document XML generation
                if let Some(endnotes_xml) = mutable_doc.generate_endnotes_xml()? {
                    let endnotes_uri = PackURI::new("/word/endnotes.xml")
                        .map_err(|e| OoxmlError::InvalidUri(format!("endnotes URI: {}", e)))?;
                    let endnotes_part = BlobPart::new(
                        endnotes_uri,
                        ct::WML_ENDNOTES.to_string(),
                        endnotes_xml.into_bytes(),
                    );
                    self.opc.add_part(Box::new(endnotes_part));
                    let rid = temp_part.relate_to("endnotes.xml", rt::ENDNOTES);
                    rel_mapper.set_endnotes_id(rid);
                }

                // Handle watermark headers before generating document XML
                // This ensures header relationships are properly set up
                if mutable_doc.has_watermark() || mutable_doc.has_image_watermark() {
                    // Generate user header content if exists (will be merged with watermark)
                    let user_header_content = if mutable_doc.has_header() {
                        mutable_doc.generate_header_xml()?
                    } else {
                        None
                    };

                    // Store the watermark image as an ordinary media part,
                    // shared by all three headers.
                    let image_media_name =
                        if let Some(image_watermark) = mutable_doc.image_watermark.as_ref() {
                            let media_name = format!(
                                "/word/media/watermarkImage1.{}",
                                image_watermark.format().extension()
                            );
                            let media_uri = PackURI::new(&media_name).map_err(|e| {
                                OoxmlError::InvalidUri(format!("watermark image URI: {}", e))
                            })?;
                            self.opc.add_part(Box::new(BlobPart::new(
                                media_uri,
                                image_watermark.format().mime_type().to_string(),
                                image_watermark.data().to_vec(),
                            )));
                            Some(media_name)
                        } else {
                            None
                        };

                    // Create three headers (default, first, even) with watermark
                    let header_types = [
                        ("/word/header1.xml", "header1.xml"),
                        ("/word/header2.xml", "header2.xml"),
                        ("/word/header3.xml", "header3.xml"),
                    ];

                    for (idx, (header_uri_path, header_filename)) in header_types.iter().enumerate()
                    {
                        let mut watermark_xml = String::new();
                        if let Some(wm) = mutable_doc.watermark.as_ref() {
                            watermark_xml.push_str(&wm.to_header_xml((idx + 1) as u32)?);
                        }

                        let header_uri = PackURI::new(*header_uri_path)
                            .map_err(|e| OoxmlError::InvalidUri(format!("header URI: {}", e)))?;
                        let mut header_part =
                            BlobPart::new(header_uri, ct::WML_HEADER.to_string(), Vec::new());

                        // The image watermark references the media part
                        // through a relationship owned by this header part.
                        if let (Some(image_watermark), Some(media_name)) = (
                            mutable_doc.image_watermark.as_ref(),
                            image_media_name.as_deref(),
                        ) {
                            let media_target =
                                media_name.strip_prefix("/word/").unwrap_or(media_name);
                            let rel_id = header_part.relate_to(media_target, rt::IMAGE);
                            watermark_xml.push_str(
                                &image_watermark.to_header_xml((idx + 1) as u32, &rel_id)?,
                            );
                        }

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
                                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}{}</w:hdr>"#,
                                watermark_xml, user_paragraphs
                            )
                        } else {
                            // Just watermark for first and even headers
                            format!(
                                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{}</w:hdr>"#,
                                watermark_xml
                            )
                        };

                        header_part.set_blob(header_xml.into_bytes());
                        self.opc.add_part(Box::new(header_part));

                        // Add relationship for the default header
                        if idx == 0 {
                            let rid = temp_part.relate_to(header_filename, rt::HEADER);
                            rel_mapper.set_header_id(rid);
                        } else {
                            // Other headers are added but not set in rel_mapper (they're referenced in sectPr)
                            temp_part.relate_to(header_filename, rt::HEADER);
                        }
                    }
                }

                // Step 3: Generate XML with actual relationship IDs
                let xml = mutable_doc.to_xml_with_rels(&rel_mapper)?;

                // Step 4: Update the document part with final XML and relationships
                for (header, part, uri, _) in planned_section_parts {
                    let content_type = if header {
                        ct::WML_HEADER
                    } else {
                        ct::WML_FOOTER
                    };
                    self.opc.add_part(Box::new(BlobPart::new(
                        uri,
                        content_type.to_string(),
                        part.xml.into_bytes(),
                    )));
                }
                temp_part.set_blob(xml.into_bytes());
                self.opc.add_part(Box::new(temp_part));

                // Note: Footnotes and endnotes are already handled above (before document XML generation)
                // so they appear in sectPr with proper relationship IDs

                // Update comments if present
                if let Some(comments_xml) = mutable_doc.generate_comments_xml()? {
                    self.update_comments_part(comments_xml)?;
                }

                // Patch only explicitly changed protection, preserving every other setting.
                if mutable_doc.protection_is_dirty() {
                    let settings_uri = PackURI::new("/word/settings.xml").map_err(|error| {
                        OoxmlError::InvalidUri(format!("settings URI: {error}"))
                    })?;
                    let existing_settings = self
                        .opc
                        .get_part(&settings_uri)
                        .ok()
                        .map(|part| part.blob().to_vec());
                    let settings_xml =
                        mutable_doc.generate_settings_xml(existing_settings.as_deref())?;
                    self.update_settings_part(settings_xml)?;
                }

                // Update theme if present
                if let Some(theme_xml) = mutable_doc.generate_theme_xml()? {
                    self.update_theme_part(theme_xml)?;
                }
            }

            // Update or remove the custom-properties package graph atomically.
            self.custom_props.write(&mut self.opc)?;

            // Embed fonts if feature enabled and requested in options
            #[cfg(feature = "fonts")]
            {
                if let Some(mutable_doc) = rollback.mutable_doc_mut() {
                    self.embed_fonts_for_document(mutable_doc)?;
                } else {
                    self.embed_fonts()?;
                }
            }

            // Stage only an explicitly edited core-properties slot. The guard
            // keeps edit intent dirty until the output sink accepts the complete
            // package, so a failed stream remains retryable.
            let staged_properties = self.properties.stage(&mut self.opc)?;

            self.opc.to_stream(writer).map_err(|e| {
                OoxmlError::Io(std::io::Error::other(format!(
                    "Failed to save package: {}",
                    e
                )))
            })?;
            staged_properties.commit();
            self.custom_props_dirty = false;
            Ok(())
        }));

        match result {
            Ok(Ok(())) => {
                rollback.publish(self);
                Ok(())
            },
            Ok(Err(error)) => {
                rollback.rollback(self);
                Err(error)
            },
            Err(payload) => {
                rollback.rollback(self);
                std::panic::resume_unwind(payload);
            },
        }
    }

    /// Borrows the document core properties, retaining package absence.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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
    /// use litchi_ooxml::docx::Package;
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

    /// Update the footnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_footnotes_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let footnotes_uri = PackURI::new("/word/footnotes.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("footnotes URI: {}", e)))?;

        let content_type = ct::WML_FOOTNOTES.to_string();
        let footnotes_part = BlobPart::new(footnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the footnotes part
        self.opc.add_part(Box::new(footnotes_part));

        // Create relationship from document to footnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

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
            .map_err(|e| OoxmlError::InvalidUri(format!("endnotes URI: {}", e)))?;

        let content_type = ct::WML_ENDNOTES.to_string();
        let endnotes_part = BlobPart::new(endnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the endnotes part
        self.opc.add_part(Box::new(endnotes_part));

        // Create relationship from document to endnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("endnotes.xml", rt::ENDNOTES);
        }

        Ok(())
    }

    /// Update or create the comments part with the given XML content.
    fn update_comments_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let comments_uri = PackURI::new("/word/comments.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("comments URI: {}", e)))?;

        let content_type = ct::WML_COMMENTS.to_string();
        let comments_part = BlobPart::new(comments_uri.clone(), content_type, xml.into_bytes());

        // Add the comments part
        self.opc.add_part(Box::new(comments_part));

        // Create relationship from document to comments
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("/word/comments.xml", rt::COMMENTS);
        }

        Ok(())
    }

    /// Update the settings.xml part with new content.
    fn update_settings_part(&mut self, xml: Vec<u8>) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&part)?;
        self.commit_settings_part(&snapshot, part)
    }

    fn settings_part_snapshot(&self) -> Result<SettingsPartSnapshot> {
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
            return Err(OoxmlError::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let (target, relationship_exists) = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(OoxmlError::InvalidFormat(
                    "settings relationship cannot be external".into(),
                ));
            },
            Some(relationship) => (relationship.target_partname()?, true),
            None => (
                PackURI::new("/word/settings.xml")
                    .map_err(|error| OoxmlError::InvalidUri(error.to_string()))?,
                false,
            ),
        };

        let (content_type, xml, relationships) = match self.opc.get_part(&target) {
            Ok(part) => {
                if part.content_type() != ct::WML_SETTINGS {
                    return Err(OoxmlError::InvalidFormat(format!(
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
                return Err(OoxmlError::PartNotFound(format!("settings part {target}")));
            },
            Err(_) => (
                ct::WML_SETTINGS.to_owned(),
                crate::docx::template::default_settings_xml()
                    .as_bytes()
                    .to_vec(),
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

    fn commit_settings_part(
        &mut self,
        snapshot: &SettingsPartSnapshot,
        part: litchi_opc::part::BlobPart,
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

    fn update_theme_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::part::BlobPart;

        let theme_uri = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme URI: {}", e)))?;

        let content_type = "application/vnd.openxmlformats-officedocument.theme+xml".to_string();
        let theme_part = BlobPart::new(theme_uri.clone(), content_type, xml.into_bytes());

        // Add/replace the theme part
        self.opc.add_part(Box::new(theme_part));

        // Add relationship from document to theme if not exists
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

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
    fn update_watermark_headers(
        &mut self,
        mutable_doc: &crate::docx::writer::MutableDocument,
    ) -> Result<()> {
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
            .map_err(|e| OoxmlError::InvalidUri(format!("document URI: {}", e)))?;

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
                .map_err(|e| OoxmlError::InvalidUri(format!("header URI: {}", e)))?;

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

fn settings_part_from_snapshot(
    snapshot: &SettingsPartSnapshot,
    xml: Vec<u8>,
    omitted_relationship_id: Option<&str>,
) -> litchi_opc::part::BlobPart {
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

fn mail_merge_relationship_type(conformance: mail_merge::Conformance, suffix: &str) -> String {
    let base = match conformance {
        mail_merge::Conformance::Transitional => {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        },
        mail_merge::Conformance::Strict => {
            "http://purl.oclc.org/ooxml/officeDocument/relationships"
        },
    };
    format!("{base}/{suffix}")
}

fn allocate_mail_merge_relationship_id(
    label: &str,
    used: &mut std::collections::HashSet<String>,
) -> Result<String> {
    (1usize..=MAX_MAIL_MERGE_RELATIONSHIPS)
        .map(|number| format!("rIdMailMerge{label}{number}"))
        .find(|id| used.insert(id.clone()))
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("mail-merge relationship ID space is exhausted".into())
        })
}

fn validate_mail_merge_external_uri(uri: &str) -> Result<()> {
    if uri.is_empty() || uri.len() > 32 * 1024 || uri.chars().any(char::is_control) {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge external target is empty or exceeds URI limits".into(),
        ));
    }
    Ok(())
}

fn validate_mail_merge_internal_source(
    bytes: &[u8],
    content_type: &str,
    extension: &str,
) -> Result<()> {
    if bytes.len() > 128 * 1024 * 1024 {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source exceeds the 128 MiB authoring limit".into(),
        ));
    }
    if content_type.is_empty()
        || content_type.len() > 1024
        || content_type.chars().any(char::is_control)
    {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source content type is invalid".into(),
        ));
    }
    if extension.is_empty()
        || extension.len() > 16
        || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(OoxmlError::InvalidFormat(
            "mail-merge source extension is invalid".into(),
        ));
    }
    Ok(())
}

fn mail_merge_target_as_source(target: Target) -> Source {
    match target {
        Target::External(uri) => Source::External(uri),
        Target::Internal {
            part_name,
            bytes,
            content_type,
        } => {
            let extension = part_name
                .as_str()
                .rsplit_once('.')
                .map(|(_, extension)| extension)
                .filter(|extension| {
                    !extension.is_empty()
                        && extension.len() <= 16
                        && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                })
                .unwrap_or("bin")
                .to_string();
            Source::Internal {
                bytes,
                content_type,
                extension,
            }
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Seek, Write};
    use std::panic::{AssertUnwindSafe, catch_unwind};
    use tempfile::NamedTempFile;

    struct FailingWriter;

    impl Write for FailingWriter {
        fn write(&mut self, _buffer: &[u8]) -> std::io::Result<usize> {
            Err(std::io::Error::other("injected DOCX sink failure"))
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    impl Seek for FailingWriter {
        fn seek(&mut self, _position: std::io::SeekFrom) -> std::io::Result<u64> {
            Err(std::io::Error::other("injected DOCX seek failure"))
        }
    }

    struct PanickingWriter;

    impl Write for PanickingWriter {
        fn write(&mut self, _buffer: &[u8]) -> std::io::Result<usize> {
            panic!("injected DOCX sink panic")
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    impl Seek for PanickingWriter {
        fn seek(&mut self, _position: std::io::SeekFrom) -> std::io::Result<u64> {
            panic!("injected DOCX seek panic")
        }
    }

    #[test]
    fn saves_and_reopens_package() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("round-trip text");
        package.save(file.path()).unwrap();

        let mut reopened = Package::open(file.path()).unwrap();
        assert!(
            reopened
                .document()
                .unwrap()
                .text()
                .unwrap()
                .contains("round-trip text")
        );

        reopened
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("appended after reopen");
        reopened.save(file.path()).unwrap();
        let reopened_again = Package::open(file.path()).unwrap();
        let text = reopened_again.document().unwrap().text().unwrap();
        assert!(text.contains("round-trip text"));
        assert!(text.contains("appended after reopen"));
    }

    #[test]
    fn failed_stream_keeps_document_and_properties_retryable() {
        let mut package = Package::new().unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("retryable document");
        package.put_props(Props::new().title("retryable properties"));
        let document_before = package.opc.main_document_part().unwrap().blob_arc();
        let core_properties_uri = PackURI::new("/docProps/core.xml").unwrap();
        let core_properties_before = package
            .opc
            .get_part(&core_properties_uri)
            .unwrap()
            .blob_arc();

        assert!(package.to_plain_stream(FailingWriter).is_err());
        assert_eq!(
            package.opc.main_document_part().unwrap().blob(),
            document_before.as_slice()
        );
        assert!(std::sync::Arc::ptr_eq(
            &document_before,
            &package.opc.main_document_part().unwrap().blob_arc()
        ));
        assert_eq!(
            package.opc.get_part(&core_properties_uri).unwrap().blob(),
            core_properties_before.as_slice()
        );
        assert!(std::sync::Arc::ptr_eq(
            &core_properties_before,
            &package
                .opc
                .get_part(&core_properties_uri)
                .unwrap()
                .blob_arc()
        ));
        assert!(
            package
                .mutable_doc
                .as_ref()
                .is_some_and(MutableDocument::is_modified)
        );
        assert!(package.properties.is_dirty());

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("second attempt");
        let mut output = Cursor::new(Vec::new());
        package.to_plain_stream(&mut output).unwrap();
        assert!(!output.into_inner().is_empty());
        assert!(!package.properties.is_dirty());
    }

    #[test]
    fn panicking_stream_restores_package_and_retryable_state() {
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_heading("panic-safe heading", 1).unwrap();
            document
                .add_toc(crate::docx::TableOfContents::new())
                .unwrap();
        }
        package.put_props(Props::new().title("panic-safe properties"));
        package
            .custom_props_mut()
            .insert("RetryMarker", "panic-safe custom property")
            .unwrap();

        let package_before = litchi_opc::PackageWriter::to_bytes(&package.opc).unwrap();
        let document_before = package.mutable_doc.as_ref().unwrap().to_xml().unwrap();

        let unwind = catch_unwind(AssertUnwindSafe(|| {
            package.to_plain_stream(PanickingWriter).unwrap();
        }));
        assert!(unwind.is_err());
        assert_eq!(
            litchi_opc::PackageWriter::to_bytes(&package.opc).unwrap(),
            package_before
        );
        let document_after = package.mutable_doc.as_ref().unwrap().to_xml().unwrap();
        // TOC materialization is a one-shot semantic mutation: it consumes the
        // pending configuration and inserts the generated field before the sink
        // runs. The rollback restores that same writer value and its edit
        // intent, so retryability is checked through the retained heading/TOC
        // semantics rather than an impossible byte-identical writer snapshot.
        assert_ne!(document_after, document_before);
        assert!(document_after.contains("panic-safe heading"));
        assert!(document_after.contains("TOC"));
        assert!(
            package
                .mutable_doc
                .as_ref()
                .is_some_and(MutableDocument::is_modified)
        );
        assert!(package.properties.is_dirty());

        let mut output = Cursor::new(Vec::new());
        package.to_plain_stream(&mut output).unwrap();
        let output = output.into_inner();
        assert!(!output.is_empty());
        assert!(!package.properties.is_dirty());
        let reopened = Package::from_opc_package(OpcPackage::from_bytes(&output).unwrap()).unwrap();
        assert!(
            reopened
                .document()
                .unwrap()
                .text()
                .unwrap()
                .contains("panic-safe heading")
        );
        assert_eq!(
            reopened
                .document()
                .unwrap()
                .table_of_contents_count()
                .unwrap(),
            1
        );
        assert!(reopened.custom_props().contains("RetryMarker"));
    }

    #[test]
    fn unchanged_stream_preserves_exact_bytes_and_part_payload_sharing() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut source = Package::new().unwrap();
        source
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("unchanged package");
        source.save(file.path()).unwrap();

        let mut package = Package::open(file.path()).unwrap();
        let before = litchi_opc::PackageWriter::to_bytes(package.opc_package()).unwrap();
        let document_uri = PackURI::new("/word/document.xml").unwrap();
        let core_properties_uri = PackURI::new("/docProps/core.xml").unwrap();
        let document_before = package.opc.get_part(&document_uri).unwrap().blob_arc();
        let core_properties_before = package
            .opc
            .get_part(&core_properties_uri)
            .unwrap()
            .blob_arc();

        let mut output = Cursor::new(Vec::new());
        package.to_plain_stream(&mut output).unwrap();
        assert_eq!(output.into_inner(), before);
        assert!(std::sync::Arc::ptr_eq(
            &document_before,
            &package.opc.get_part(&document_uri).unwrap().blob_arc()
        ));
        assert!(std::sync::Arc::ptr_eq(
            &core_properties_before,
            &package
                .opc
                .get_part(&core_properties_uri)
                .unwrap()
                .blob_arc()
        ));
    }

    #[cfg(feature = "fonts")]
    #[test]
    fn raw_opc_rejects_automatic_font_embedding_policy() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut source = Package::new().unwrap();
        source
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("preserved text");
        source.save(file.path()).unwrap();

        let mut opened = Package::open(file.path()).unwrap();
        opened.opc.with_fonts(litchi_opc::FontEmbedding::Subset);
        assert!(matches!(
            opened.edit_opc(|_| Ok(())),
            Err(OoxmlError::UnsafeEdit {
                operation: "edit_opc",
                ..
            })
        ));
    }

    #[test]
    fn raw_opc_transaction_publishes_candidate_and_disables_writer() {
        let mut package = Package::new().unwrap();
        let marker = PackURI::new("/custom/raw-edit-marker.bin").unwrap();

        package
            .edit_opc(|candidate| {
                candidate.try_add_part(Box::new(BlobPart::new(
                    marker.clone(),
                    "application/octet-stream".to_string(),
                    b"raw edit".to_vec(),
                )))?;
                Ok::<_, OoxmlError>(())
            })
            .unwrap();

        assert_eq!(
            package.opc_package().get_part(&marker).unwrap().blob(),
            b"raw edit"
        );
        assert!(matches!(
            package.document_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "document_mut",
                ..
            })
        ));
    }

    #[test]
    fn failed_raw_opc_transaction_preserves_graph_and_writer_state() {
        let mut package = Package::new().unwrap();
        let document_uri = PackURI::new("/word/document.xml").unwrap();
        let original = package
            .opc_package()
            .get_part(&document_uri)
            .unwrap()
            .blob_arc();

        let error = package
            .edit_opc(|candidate| {
                candidate.remove_part(&document_uri);
                Ok::<_, OoxmlError>(())
            })
            .unwrap_err();
        assert!(matches!(error, OoxmlError::PartNotFound(_)));
        assert!(std::sync::Arc::ptr_eq(
            &original,
            &package
                .opc_package()
                .get_part(&document_uri)
                .unwrap()
                .blob_arc()
        ));
        assert!(package.document_mut().is_ok());
    }

    #[test]
    fn raw_opc_transaction_rejects_pending_managed_state() {
        let mut package = Package::new().unwrap();
        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("managed edit");

        assert!(matches!(
            package.edit_opc(|_| Ok::<_, OoxmlError>(())),
            Err(OoxmlError::UnsafeEdit {
                operation: "edit_opc",
                ..
            })
        ));
    }

    #[test]
    fn saves_and_reopens_inline_and_display_office_math() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let inline = crate::docx::OfficeMath::text("x + y");
        let display =
            crate::docx::OfficeMath::from_xml("<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>")
                .unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document
                .add_paragraph()
                .add_inline_office_math(inline.clone());
            document.add_display_office_math(display.clone());
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document_uri = PackURI::new("/word/document.xml").unwrap();
        let document_xml =
            std::str::from_utf8(reopened.opc.get_part(&document_uri).unwrap().blob()).unwrap();
        let document_opening = &document_xml[..document_xml.find("><w:body>").unwrap()];
        assert!(
            document_opening
                .contains("xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"")
        );
        let paragraphs = reopened.document().unwrap().paragraphs().unwrap();
        assert_eq!(paragraphs[0].inline_office_math().unwrap(), vec![inline]);
        assert_eq!(paragraphs[1].display_office_math().unwrap(), vec![display]);
    }

    #[test]
    fn writes_and_rediscovers_distinct_watermarks() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let mut watermark = crate::docx::Watermark::text("INTERNAL");
        watermark.set_font("Aptos");
        watermark.set_color("808080");
        package.document_mut().unwrap().set_watermark(watermark);
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let watermarks = reopened.document().unwrap().watermarks().unwrap();

        assert_eq!(watermarks.len(), 1);
        assert_eq!(watermarks[0].get_text(), "INTERNAL");
        assert_eq!(watermarks[0].font(), "Aptos");
        assert_eq!(watermarks[0].color(), "#808080");
    }

    #[test]
    fn writes_and_discovers_typed_table_of_contents_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_heading("Overview", 1).unwrap();
            document
                .add_toc(
                    crate::docx::TableOfContents::new()
                        .heading_levels(1, 4)
                        .hyperlinks(true),
                )
                .unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let toc = document.table_of_contents().unwrap();
        assert_eq!(document.table_of_contents_count().unwrap(), 1);
        assert_eq!(toc.len(), 1);
        assert!(toc[0].includes_hyperlinks());
        assert!(toc[0].hides_page_numbers_in_web_layout());
        assert_eq!(
            toc[0].heading_style_levels().unwrap(),
            vec![crate::docx::TocLevelRange::new(1, 4).unwrap()]
        );
    }

    #[test]
    fn all_supported_headings_resolve_in_the_saved_style_catalog() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for level in 0..=9 {
                document
                    .add_heading(&format!("Heading {level}"), level)
                    .unwrap();
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let ids = document
            .paragraphs()
            .unwrap()
            .into_iter()
            .map(|paragraph| paragraph.style_id().unwrap().unwrap())
            .collect::<Vec<_>>();
        let expected = std::iter::once("Title".to_string())
            .chain((1..=9).map(|level| format!("Heading{level}")))
            .collect::<Vec<_>>();
        assert_eq!(ids, expected);

        let mut styles = document.styles().unwrap();
        let outlines = [
            None,
            Some(crate::docx::styles::Outline::H1),
            Some(crate::docx::styles::Outline::H2),
            Some(crate::docx::styles::Outline::H3),
            Some(crate::docx::styles::Outline::H4),
            Some(crate::docx::styles::Outline::H5),
            Some(crate::docx::styles::Outline::H6),
            Some(crate::docx::styles::Outline::H7),
            Some(crate::docx::styles::Outline::H8),
            Some(crate::docx::styles::Outline::H9),
        ];
        for (id, outline) in expected.into_iter().zip(outlines) {
            let style = styles
                .get_by_id(&id)
                .unwrap()
                .unwrap_or_else(|| panic!("missing {id}"));
            assert_eq!(style.outline(), outline, "wrong outline for {id}");
        }
    }

    #[test]
    fn writes_and_discovers_typed_table_of_contents_entry_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"TC "Illustration 1" \f i \l 4 \n"#.to_string(),
                "cached entry".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let entries = document.table_of_contents_entries().unwrap();
        assert_eq!(document.table_of_contents_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "Illustration 1");
        assert_eq!(entries[0].cached_result(), Some("cached entry"));
        assert_eq!(entries[0].list_identifier().unwrap(), Some("i"));
        assert_eq!(entries[0].level().unwrap(), Some("4"));
        assert!(entries[0].omits_page_number());
    }

    #[test]
    fn writes_and_discovers_typed_table_of_authorities_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let authority_table = document.add_paragraph();
            authority_table.add_field(crate::docx::writer::MutableField::with_result(
                r#"TOA \c 2 \b "Authorities" \p \h"#.to_string(),
                "Statutes\t3".to_string(),
            ));
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"TA \l "Example Statute" \s "Example" \c 2 \b"#.to_string(),
                "hidden marker".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let authorities = document.tables_of_authorities().unwrap();
        assert_eq!(document.table_of_authorities_count().unwrap(), 1);
        assert_eq!(authorities.len(), 1);
        assert_eq!(authorities[0].category().unwrap(), Some(2));
        assert_eq!(authorities[0].bookmark().unwrap(), Some("Authorities"));
        assert!(authorities[0].uses_passim());
        assert!(authorities[0].includes_category_headers());

        let entries = document.table_of_authorities_entries().unwrap();
        assert_eq!(document.table_of_authorities_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].long_citation().unwrap(), Some("Example Statute"));
        assert_eq!(entries[0].short_citation().unwrap(), Some("Example"));
        assert_eq!(entries[0].category().unwrap(), Some(2));
        assert!(entries[0].is_bold());
    }

    #[test]
    fn writes_and_discovers_typed_citation_and_bibliography_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let citation = document.add_paragraph();
            let mut primary = crate::docx::CitationSource::new("Doe2024").unwrap();
            primary.set_prefix(Some("qtd. in".to_string())).unwrap();
            let mut citation_spec = crate::docx::CitationFieldSpec::new(primary);
            citation_spec.set_locale(Some(1033));
            citation_spec
                .add_source(crate::docx::CitationSource::new("Smith2025").unwrap())
                .unwrap();
            citation_spec
                .set_cached_result(Some("(Doe, 2024; Smith, 2025)".to_string()))
                .unwrap();
            citation_spec.set_dirty(false);
            citation
                .add_field(crate::docx::writer::MutableField::citation(&citation_spec).unwrap());
            let bibliography = document.add_paragraph();
            let mut bibliography_spec = crate::docx::BibliographyFieldSpec::new();
            bibliography_spec.set_locale(Some(1033));
            bibliography_spec.set_filter(Some(crate::docx::BibliographyFilter::Locale(1036)));
            bibliography_spec.add_source_tag("Doe2024").unwrap();
            bibliography_spec.add_source_tag("Smith2025").unwrap();
            bibliography_spec
                .set_cached_result(Some("Doe. Example work.".to_string()))
                .unwrap();
            bibliography_spec.set_dirty(false);
            bibliography.add_field(
                crate::docx::writer::MutableField::bibliography(&bibliography_spec).unwrap(),
            );
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let citations = document.citations().unwrap();
        assert_eq!(document.citation_count().unwrap(), 1);
        assert_eq!(citations.len(), 1);
        assert_eq!(citations[0].primary_source_tag(), "Doe2024");
        assert_eq!(citations[0].source_tags(), ["Doe2024", "Smith2025"]);
        assert!(citations[0].has_switch('l'));
        assert!(citations[0].has_switch('f'));
        assert!(!citations[0].is_dirty());
        assert_eq!(
            citations[0].cached_result(),
            Some("(Doe, 2024; Smith, 2025)")
        );

        let bibliographies = document.bibliographies().unwrap();
        assert_eq!(document.bibliography_count().unwrap(), 1);
        assert_eq!(bibliographies.len(), 1);
        assert_eq!(
            bibliographies[0].cached_result(),
            Some("Doe. Example work.")
        );
        assert!(bibliographies[0].has_switch('l'));
        assert!(bibliographies[0].has_switch('f'));
        assert!(bibliographies[0].has_switch('m'));
        assert!(!bibliographies[0].is_dirty());
        assert_eq!(bibliographies[0].switches()[0].argument(), Some("1033"));
        assert_eq!(bibliographies[0].switches()[1].argument(), Some("1036"));
    }

    #[test]
    fn writes_and_discovers_inert_document_variable_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"DOCVARIABLE CustomerName \* MERGEFORMAT"#.to_string(),
                "cached customer".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_variable_fields().unwrap();
        assert_eq!(document.document_variable_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].variable_name(), "CustomerName");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].has_switch('*'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert!(
            document
                .document_variables()
                .unwrap()
                .is_none_or(|variables| variables.get("CustomerName").is_none())
        );
    }

    #[test]
    fn writes_and_discovers_inert_document_property_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy""#.to_string(),
                "cached project".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_property_fields().unwrap();
        assert_eq!(document.document_property_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].property_name(), "Project Name");
        assert_eq!(fields[0].cached_result(), Some("cached project"));
        assert!(fields[0].has_switch('*'));
        assert!(fields[0].has_switch('@'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(fields[0].switches()[1].argument(), Some("MMMM d, yyyy"));
    }

    #[test]
    fn writes_and_discovers_inert_document_information_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let title = document.add_paragraph();
            title.add_field(crate::docx::writer::MutableField::with_result(
                r#"TITLE \* MERGEFORMAT"#.to_string(),
                "cached title".to_string(),
            ));
            let author = document.add_paragraph();
            author.add_field(crate::docx::writer::MutableField::with_result(
                r#"AUTHOR \@ "opaque format""#.to_string(),
                "cached author".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_information_fields().unwrap();
        assert_eq!(document.document_information_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), crate::docx::InformationKind::Title);
        assert_eq!(fields[0].cached_result(), Some("cached title"));
        assert!(fields[0].has_switch('*'));
        assert_eq!(fields[0].switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(fields[1].kind(), crate::docx::InformationKind::Author);
        assert_eq!(fields[1].cached_result(), Some("cached author"));
        assert!(fields[1].has_switch('@'));
        assert_eq!(fields[1].switches()[0].argument(), Some("opaque format"));
    }

    #[test]
    fn writes_and_discovers_inert_document_context_fields_without_resolution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let file_name = document.add_paragraph();
            file_name.add_field(crate::docx::writer::MutableField::with_result(
                r#"FILENAME \p"#.to_string(),
                "cached file name".to_string(),
            ));
            let page = document.add_paragraph();
            page.add_field(crate::docx::writer::MutableField::with_result(
                r#"PAGE \* MERGEFORMAT"#.to_string(),
                "cached page".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.document_context_fields().unwrap();
        assert_eq!(document.document_context_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), crate::docx::ContextKind::FileName);
        assert_eq!(fields[0].cached_result(), Some("cached file name"));
        assert!(fields[0].has_switch('p'));
        assert_eq!(fields[1].kind(), crate::docx::ContextKind::Page);
        assert_eq!(fields[1].cached_result(), Some("cached page"));
        assert!(fields[1].has_switch('*'));
        assert_eq!(fields[1].switches()[0].argument(), Some("MERGEFORMAT"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_merge_fields_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"MERGEFIELD "Customer Region" \b "Dear " \f "!" \m \v \* MERGEFORMAT"#
                    .to_string(),
                "cached customer".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let fields = document.typed_merge_fields().unwrap();
        assert_eq!(document.typed_merge_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].field_name(), "Customer Region");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].has_switch('b'));
        assert!(fields[0].has_switch('f'));
        assert!(fields[0].has_switch('m'));
        assert!(fields[0].has_switch('v'));
        assert_eq!(fields[0].switches()[0].argument(), Some("Dear "));
        assert_eq!(fields[0].switches()[1].argument(), Some("!"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_mail_merge_counters_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document
                .add_paragraph()
                .add_field(crate::docx::writer::MutableField::with_result(
                    "MERGEREC".to_string(),
                    "12".to_string(),
                ));
            document
                .add_paragraph()
                .add_field(crate::docx::writer::MutableField::with_result(
                    "MERGESEQ".to_string(),
                    "3".to_string(),
                ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let counters = document.mail_merge_counters().unwrap();
        assert_eq!(document.mail_merge_counter_count().unwrap(), 2);
        assert_eq!(counters.len(), 2);
        assert_eq!(counters[0].kind(), crate::docx::MergeCounterKind::Record);
        assert_eq!(counters[0].cached_result(), Some("12"));
        assert_eq!(counters[1].kind(), crate::docx::MergeCounterKind::Sequence);
        assert_eq!(counters[1].cached_result(), Some("3"));
    }

    #[test]
    fn writes_and_discovers_inert_mail_merge_next_fields_without_advancing_records() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                "NEXT".to_string(),
                "cached next".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let fields = document.mail_merge_next_fields().unwrap();
        assert_eq!(document.mail_merge_next_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].cached_result(), Some("cached next"));
    }

    #[test]
    fn writes_and_discovers_inert_conditional_mail_merge_controls_without_merging() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                r#"SKIPIF MERGEFIELD Order < 100"#.to_string(),
                "cached skipif".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert!(reopened.mail_merge_settings().unwrap().is_none());
        let document = reopened.document().unwrap();
        let controls = document.mail_merge_conditional_controls().unwrap();
        assert_eq!(document.mail_merge_conditional_control_count().unwrap(), 1);
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].kind(), crate::docx::MergeControlKind::SkipIf);
        assert_eq!(controls[0].comparison(), "MERGEFIELD Order < 100");
        assert_eq!(controls[0].cached_result(), Some("cached skipif"));
    }

    #[test]
    fn writes_and_discovers_inert_if_fields_without_evaluation() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.document_mut().unwrap().add_paragraph().add_field(
            crate::docx::writer::MutableField::with_result(
                r#"IF 1 = 1 "yes" "no""#.to_string(),
                "yes".to_string(),
            ),
        );
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.if_fields().unwrap();
        assert_eq!(document.if_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].expression(), r#"1 = 1 "yes" "no""#);
        assert_eq!(fields[0].cached_result(), Some("yes"));
    }

    #[test]
    fn writes_and_discovers_inert_document_state_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (
                    r#"SET RecipientName "North America" \* MERGEFORMAT"#,
                    "cached recipient",
                ),
                (r#"SEQ Figure FigureChapter \r 3 \* ARABIC"#, "3"),
                (r#"=SUM(ABOVE) \* MERGEFORMAT"#, "42"),
                (r#"STYLEREF "Heading 1" \n \p"#, "1 above"),
            ] {
                document
                    .add_paragraph()
                    .add_field(crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ));
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let sets = document.set_fields().unwrap();
        assert_eq!(document.set_field_count().unwrap(), 1);
        assert_eq!(sets.len(), 1);
        assert_eq!(sets[0].target_name(), "RecipientName");
        assert_eq!(sets[0].expression(), r#""North America" \* MERGEFORMAT"#);
        assert_eq!(sets[0].cached_result(), Some("cached recipient"));

        let sequences = document.sequence_fields().unwrap();
        assert_eq!(document.sequence_field_count().unwrap(), 1);
        assert_eq!(sequences.len(), 1);
        assert_eq!(sequences[0].identifier(), "Figure");
        assert_eq!(sequences[0].bookmark(), Some("FigureChapter"));
        assert_eq!(sequences[0].tail(), r#"\r 3 \* ARABIC"#);
        assert_eq!(sequences[0].cached_result(), Some("3"));

        let formulas = document.formula_fields().unwrap();
        assert_eq!(document.formula_field_count().unwrap(), 1);
        assert_eq!(formulas.len(), 1);
        assert_eq!(formulas[0].formula(), r#"SUM(ABOVE) \* MERGEFORMAT"#);
        assert_eq!(formulas[0].cached_result(), Some("42"));

        let style_references = document.style_reference_fields().unwrap();
        assert_eq!(document.style_reference_field_count().unwrap(), 1);
        assert_eq!(style_references.len(), 1);
        assert_eq!(style_references[0].style_name(), "Heading 1");
        assert_eq!(
            style_references[0].options(),
            &[
                crate::docx::StyleOption::ParagraphNumber,
                crate::docx::StyleOption::RelativePosition,
            ]
        );
        assert_eq!(style_references[0].cached_result(), Some("1 above"));
    }

    #[test]
    fn writes_and_discovers_inert_bookmark_reference_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (
                    r#"REF "Target Bookmark" \d "-" \f \h \n \p \r \t \w"#,
                    "cached reference",
                ),
                (r#"PAGEREF PageTarget \h \p"#, "12 above"),
                (r#"FTNREF FootnoteTarget \p \f"#, "1 above"),
                (r#"NOTEREF EndnoteTarget \p \f"#, "i above"),
            ] {
                document
                    .add_paragraph()
                    .add_field(crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ));
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let references = document.reference_fields().unwrap();
        assert_eq!(document.reference_field_count().unwrap(), 4);
        assert_eq!(references.len(), 4);

        assert_eq!(references[0].kind(), crate::docx::ReferenceKind::Reference);
        assert_eq!(references[0].bookmark(), "Target Bookmark");
        assert_eq!(
            references[0].options(),
            &[
                crate::docx::ReferenceOption::SequencePageSeparator("-".to_string()),
                crate::docx::ReferenceOption::ReferencedNoteContent,
                crate::docx::ReferenceOption::Hyperlink,
                crate::docx::ReferenceOption::ParagraphNumberWithoutContext,
                crate::docx::ReferenceOption::RelativePosition,
                crate::docx::ReferenceOption::ParagraphNumberRelativeContext,
                crate::docx::ReferenceOption::SuppressNonNumberText,
                crate::docx::ReferenceOption::ParagraphNumberFullContext,
            ]
        );
        assert_eq!(references[0].cached_result(), Some("cached reference"));

        assert_eq!(
            references[1].kind(),
            crate::docx::ReferenceKind::PageReference
        );
        assert_eq!(references[1].bookmark(), "PageTarget");
        assert_eq!(
            references[1].options(),
            &[
                crate::docx::ReferenceOption::Hyperlink,
                crate::docx::ReferenceOption::RelativePosition,
            ]
        );
        assert_eq!(references[1].cached_result(), Some("12 above"));

        assert_eq!(
            references[2].kind(),
            crate::docx::ReferenceKind::FootnoteReference
        );
        assert_eq!(references[2].bookmark(), "FootnoteTarget");
        assert_eq!(
            references[2].options(),
            &[
                crate::docx::ReferenceOption::RelativePosition,
                crate::docx::ReferenceOption::NoteMarkFormatting,
            ]
        );
        assert_eq!(references[2].cached_result(), Some("1 above"));

        assert_eq!(
            references[3].kind(),
            crate::docx::ReferenceKind::NoteReference
        );
        assert_eq!(references[3].bookmark(), "EndnoteTarget");
        assert_eq!(
            references[3].options(),
            &[
                crate::docx::ReferenceOption::RelativePosition,
                crate::docx::ReferenceOption::NoteMarkFormatting,
            ]
        );
        assert_eq!(references[3].cached_result(), Some("i above"));
    }

    #[test]
    fn writes_and_discovers_inert_equation_fields_without_calculation_or_rendering() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (r#"EQ \o\ac(\fs24 Q,\fs16 R)"#, "cached equation"),
                (r#"EQ \f(1,2)"#, "1/2"),
                ("EQ", ""),
            ] {
                document
                    .add_paragraph()
                    .add_field(crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ));
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let equations = document.equations().unwrap();
        assert_eq!(document.equation_count().unwrap(), 3);
        assert_eq!(equations.len(), 3);
        assert_eq!(equations[0].expression(), r#"\o\ac(\fs24 Q,\fs16 R)"#);
        assert_eq!(equations[0].cached_result(), Some("cached equation"));
        assert_eq!(equations[1].expression(), r#"\f(1,2)"#);
        assert_eq!(equations[1].cached_result(), Some("1/2"));
        assert_eq!(equations[2].expression(), "");
    }

    #[test]
    fn writes_and_discovers_inert_hyperlink_fields_without_opening_them() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                (
                    r#"HYPERLINK "https://example.test/a b" \l "_Toc1" \o "Stored tip" \t "_blank" \m \n"#,
                    "cached external link",
                ),
                (r#"HYPERLINK \l "JumpTarget""#, "cached internal link"),
            ] {
                document
                    .add_paragraph()
                    .add_field(crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ));
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        assert_eq!(document.hyperlink_count().unwrap(), 0);
        let fields = document.hyperlink_fields().unwrap();
        assert_eq!(document.hyperlink_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(
            fields[0].external_target(),
            Some("https://example.test/a b")
        );
        assert_eq!(fields[0].bookmark(), Some("_Toc1"));
        assert_eq!(fields[0].screen_tip(), Some("Stored tip"));
        assert_eq!(fields[0].target_frame(), Some("_blank"));
        assert!(fields[0].appends_image_map_coordinates());
        assert!(fields[0].opens_new_window());
        assert_eq!(fields[0].cached_result(), Some("cached external link"));
        assert_eq!(fields[1].external_target(), None);
        assert_eq!(fields[1].bookmark(), Some("JumpTarget"));
        assert_eq!(fields[1].cached_result(), Some("cached internal link"));
    }

    #[test]
    fn writes_and_discovers_inert_prompt_fields_without_displaying_prompts() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let ask = document.add_paragraph();
            ask.add_field(crate::docx::writer::MutableField::with_result(
                r#"ASK AskResponse "What is your first name?" \d "" \o"#.to_string(),
                "cached ask response".to_string(),
            ));
            let fill_in = document.add_paragraph();
            fill_in.add_field(crate::docx::writer::MutableField::with_result(
                r#"FILLIN "Enter appointment time" \d "09:00""#.to_string(),
                "10:30".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.prompt_fields().unwrap();
        assert_eq!(document.prompt_field_count().unwrap(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), crate::docx::PromptKind::Ask);
        assert_eq!(fields[0].bookmark(), Some("AskResponse"));
        assert_eq!(fields[0].default_response(), Some(""));
        assert!(fields[0].prompts_once_per_mail_merge());
        assert_eq!(fields[0].cached_result(), Some("cached ask response"));
        assert_eq!(fields[1].kind(), crate::docx::PromptKind::FillIn);
        assert_eq!(fields[1].bookmark(), None);
        assert_eq!(fields[1].prompt(), Some("Enter appointment time"));
        assert_eq!(fields[1].default_response(), Some("09:00"));
        assert_eq!(fields[1].cached_result(), Some("10:30"));
    }

    #[test]
    fn writes_and_discovers_inert_macro_button_fields_without_execution() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let paragraph = document.add_paragraph();
            paragraph.add_field(crate::docx::writer::MutableField::with_result(
                r#"MACROBUTTON NeverRun "Click here""#.to_string(),
                "cached button".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let fields = document.macro_button_fields().unwrap();
        assert_eq!(document.macro_button_field_count().unwrap(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].macro_name(), "NeverRun");
        assert_eq!(fields[0].display_text(), "Click here");
        assert_eq!(fields[0].cached_result(), Some("cached button"));
    }

    #[test]
    fn writes_and_discovers_inert_active_and_building_block_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for (instruction, cached_result) in [
                ("ADDIN opaque-add-in-data", "cached add-in"),
                ("CONTROL opaque-control-data", "cached control"),
                ("HTMLCONTROL opaque-html-data", "cached HTML control"),
                (
                    r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT"#,
                    "cached glossary",
                ),
                (
                    r#"AUTOTEXT "Reusable Clause" \q opaque"#,
                    "cached auto text",
                ),
                (
                    r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Select one""#,
                    "cached auto text list",
                ),
            ] {
                document
                    .add_paragraph()
                    .add_field(crate::docx::writer::MutableField::with_result(
                        instruction.to_string(),
                        cached_result.to_string(),
                    ));
            }
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let active_content = document.active_content_fields().unwrap();
        assert_eq!(document.active_content_field_count().unwrap(), 3);
        assert_eq!(active_content.len(), 3);
        assert_eq!(
            active_content[0].kind(),
            crate::docx::ActiveContentKind::AddIn
        );
        assert_eq!(
            active_content[1].kind(),
            crate::docx::ActiveContentKind::OcxControl
        );
        assert_eq!(
            active_content[2].kind(),
            crate::docx::ActiveContentKind::HtmlControl
        );
        assert_eq!(
            active_content[2].cached_result(),
            Some("cached HTML control")
        );

        let auto_text = document.auto_text_fields().unwrap();
        assert_eq!(document.auto_text_field_count().unwrap(), 2);
        assert_eq!(auto_text.len(), 2);
        assert_eq!(auto_text[0].kind(), crate::docx::AutoTextKind::Glossary);
        assert_eq!(auto_text[0].entry_name(), "Legacy Clause");
        assert_eq!(auto_text[1].kind(), crate::docx::AutoTextKind::AutoText);
        assert_eq!(auto_text[1].entry_name(), "Reusable Clause");

        let auto_text_lists = document.auto_text_list_fields().unwrap();
        assert_eq!(document.auto_text_list_field_count().unwrap(), 1);
        assert_eq!(auto_text_lists.len(), 1);
        assert_eq!(auto_text_lists[0].display_text(), Some("Choose a name"));
        assert_eq!(
            auto_text_lists[0].cached_result(),
            Some("cached auto text list")
        );
    }

    #[test]
    fn writes_and_discovers_typed_inert_dde_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let manual = document.add_paragraph();
            manual.add_field(crate::docx::writer::MutableField::with_result(
                r#"DDE Excel "missing.xlsx" "Sheet1!A1" \a \p"#.to_string(),
                "cached DDE link".to_string(),
            ));
            let automatic = document.add_paragraph();
            automatic.add_field(crate::docx::writer::MutableField::with_result(
                r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \t"#.to_string(),
                "cached DDE auto link".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let links = document.dde_links().unwrap();
        assert_eq!(document.dde_link_count().unwrap(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].kind(), crate::docx::DdeKind::Dde);
        assert_eq!(links[0].application(), "Excel");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].representation(),
            Some(crate::docx::DdeFormat::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached DDE link"));
        assert_eq!(links[1].kind(), crate::docx::DdeKind::DdeAuto);
        assert_eq!(links[1].item(), Some("Sheet1!A2"));
        assert!(links[1].requests_automatic_updates());
        assert_eq!(
            links[1].representation(),
            Some(crate::docx::DdeFormat::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached DDE auto link"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_external_include_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let text = document.add_paragraph();
            text.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDETEXT "file:///no-contact/source.docx" Summary \! \c Word8 \x /resume/name"#
                    .to_string(),
                "cached included text".to_string(),
            ));
            let picture = document.add_paragraph();
            picture.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDEPICTURE "file:///no-contact/picture.gif" \c Pictim32 \d"#.to_string(),
                "cached picture".to_string(),
            ));
            let legacy_text = document.add_paragraph();
            legacy_text.add_field(crate::docx::writer::MutableField::with_result(
                r#"INCLUDE "file:///no-contact/legacy.docx" LegacySection \!"#.to_string(),
                "cached legacy text".to_string(),
            ));
            let legacy_picture = document.add_paragraph();
            legacy_picture.add_field(crate::docx::writer::MutableField::with_result(
                r#"IMPORT "file:///no-contact/legacy.wmf" \c GraphicsFilter \d"#.to_string(),
                "cached legacy picture".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let includes = document.external_includes().unwrap();
        assert_eq!(document.external_include_count().unwrap(), 4);
        assert_eq!(includes.len(), 4);
        assert_eq!(includes[0].kind(), crate::docx::IncludeKind::Text);
        assert_eq!(includes[0].source(), "file:///no-contact/source.docx");
        assert_eq!(includes[0].bookmark(), Some("Summary"));
        assert!(includes[0].suppresses_nested_field_updates());
        assert_eq!(
            includes[0].options(),
            &[
                crate::docx::IncludeOption::Converter("Word8".to_string()),
                crate::docx::IncludeOption::XPath("/resume/name".to_string()),
            ]
        );
        assert_eq!(includes[0].cached_result(), Some("cached included text"));
        assert_eq!(includes[1].kind(), crate::docx::IncludeKind::Picture);
        assert_eq!(includes[1].source(), "file:///no-contact/picture.gif");
        assert!(includes[1].omits_picture_data());
        assert_eq!(
            includes[1].options(),
            &[crate::docx::IncludeOption::Converter(
                "Pictim32".to_string()
            )]
        );
        assert_eq!(includes[1].cached_result(), Some("cached picture"));
        assert_eq!(includes[2].kind(), crate::docx::IncludeKind::Text);
        assert_eq!(includes[2].source(), "file:///no-contact/legacy.docx");
        assert_eq!(includes[2].bookmark(), Some("LegacySection"));
        assert!(includes[2].suppresses_nested_field_updates());
        assert_eq!(includes[2].cached_result(), Some("cached legacy text"));
        assert_eq!(includes[3].kind(), crate::docx::IncludeKind::Picture);
        assert_eq!(includes[3].source(), "file:///no-contact/legacy.wmf");
        assert!(includes[3].omits_picture_data());
        assert_eq!(
            includes[3].options(),
            &[crate::docx::IncludeOption::Converter(
                "GraphicsFilter".to_string()
            )]
        );
        assert_eq!(includes[3].cached_result(), Some("cached legacy picture"));
    }

    #[test]
    fn writes_and_discovers_typed_inert_referenced_document_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let relative = document.add_paragraph();
            relative.add_field(crate::docx::writer::MutableField::with_result(
                r#"RD "C:\\Manual\\Chapters\\Chapter 1.docx" \f"#.to_string(),
                "cached relative reference".to_string(),
            ));
            let absolute = document.add_paragraph();
            absolute.add_field(crate::docx::writer::MutableField::with_result(
                r#"RD "file:///no-contact/appendix.docx""#.to_string(),
                "cached absolute reference".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let references = document.referenced_documents().unwrap();
        assert_eq!(document.referenced_document_count().unwrap(), 2);
        assert_eq!(references.len(), 2);
        assert_eq!(references[0].source(), r"C:\Manual\Chapters\Chapter 1.docx");
        assert!(references[0].uses_relative_path());
        assert_eq!(
            references[0].cached_result(),
            Some("cached relative reference")
        );
        assert_eq!(references[1].source(), "file:///no-contact/appendix.docx");
        assert!(!references[1].uses_relative_path());
        assert_eq!(
            references[1].cached_result(),
            Some("cached absolute reference")
        );
    }

    #[test]
    fn writes_and_discovers_typed_inert_link_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let spreadsheet = document.add_paragraph();
            spreadsheet.add_field(crate::docx::writer::MutableField::with_result(
                r#"LINK Excel.Sheet.8 "missing.xlsx" "Sheet1!A1" \a \f 4 \p"#.to_string(),
                "cached spreadsheet link".to_string(),
            ));
            let text = document.add_paragraph();
            text.add_field(crate::docx::writer::MutableField::with_result(
                r#"LINK Word.Document.8 "missing.docx" Bookmark \t"#.to_string(),
                "cached text link".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let links = document.link_fields().unwrap();
        assert_eq!(document.link_field_count().unwrap(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].application_type(), "Excel.Sheet.8");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].formatting_modes(),
            &[crate::docx::LinkFormat::SpreadsheetSource]
        );
        assert_eq!(
            links[0].effective_result_option(),
            Some(crate::docx::LinkResult::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached spreadsheet link"));
        assert_eq!(links[1].application_type(), "Word.Document.8");
        assert_eq!(links[1].item(), Some("Bookmark"));
        assert_eq!(
            links[1].effective_result_option(),
            Some(crate::docx::LinkResult::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached text link"));
    }

    #[test]
    fn saves_and_discovers_typed_inert_bibliography_source_stores() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .add_custom_xml(NewStore {
                xml: br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA"><b:Source><b:Tag>Doe2024</b:Tag><b:SourceType>Book</b:SourceType><b:Title>Stored source</b:Title></b:Source></b:Sources>"#.to_vec(),
                content_type: "application/xml".to_string(),
                id: "{22222222-2222-2222-2222-222222222222}".to_string(),
                schemas: vec![
                    crate::docx::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string(),
                ],
                conformance: custom_xml::Conformance::Transitional,
            })
            .unwrap();
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let stores = reopened.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        assert_eq!(
            stores[0].data_store_item_id(),
            Some("{22222222-2222-2222-2222-222222222222}")
        );
        assert_eq!(stores[0].selected_style(), Some("/APA.XSL"));
        assert_eq!(stores[0].style_name(), Some("APA"));
        assert_eq!(stores[0].source_count(), 1);
        assert_eq!(stores[0].sources()[0].tag(), Some("Doe2024"));
        assert_eq!(stores[0].sources()[0].source_type(), Some("Book"));
        assert_eq!(stores[0].sources()[0].title(), Some("Stored source"));

        let sources = reopened.bibliography_sources().unwrap();
        assert_eq!(reopened.bibliography_source_count().unwrap(), 1);
        assert_eq!(sources.len(), 1);
        assert_eq!(sources[0].tag(), Some("Doe2024"));
    }

    #[test]
    fn writes_and_discovers_typed_index_fields() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let index = document.add_paragraph();
            index.add_field(crate::docx::writer::MutableField::with_result(
                r#"INDEX \c 2 \f "topics" \r"#.to_string(),
                "Topic\t3".to_string(),
            ));
            let entry = document.add_paragraph();
            entry.add_field(crate::docx::writer::MutableField::with_result(
                r#"XE "Topic" \f "topics" \r TopicRange \b"#.to_string(),
                "hidden marker".to_string(),
            ));
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        let indexes = document.indexes().unwrap();
        assert_eq!(document.index_count().unwrap(), 1);
        assert_eq!(indexes.len(), 1);
        assert_eq!(indexes[0].columns().unwrap(), Some(2));
        assert_eq!(indexes[0].entry_identifier().unwrap(), Some("topics"));
        assert!(indexes[0].runs_subentries_inline());

        let entries = document.index_entries().unwrap();
        assert_eq!(document.index_entry_count().unwrap(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "Topic");
        assert_eq!(entries[0].entry_identifier().unwrap(), Some("topics"));
        assert_eq!(
            entries[0].page_range_bookmark().unwrap(),
            Some("TopicRange")
        );
        assert!(entries[0].is_bold());
    }

    #[test]
    fn body_edits_preserve_settings_part_byte_for_byte() {
        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let before = package.opc.get_part(&settings_uri).unwrap().blob().to_vec();

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("body-only edit");
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        assert_eq!(package.opc.get_part(&settings_uri).unwrap().blob(), before);
    }

    #[test]
    fn new_package_uses_leaf_owned_web_settings_bytes() {
        use litchi_docx::web::{Conformance, Settings, write};

        let package = Package::new().unwrap();
        let uri = PackURI::new("/word/webSettings.xml").unwrap();
        let expected = write(&Settings::default(), Conformance::Transitional).unwrap();

        assert_eq!(package.opc.get_part(&uri).unwrap().blob(), expected);
    }

    #[test]
    fn body_edits_preserve_web_settings_part_byte_for_byte() {
        let mut package = Package::new().unwrap();
        let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
        let before = package
            .opc
            .get_part(&web_settings_uri)
            .unwrap()
            .blob()
            .to_vec();

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("body-only edit");
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        assert_eq!(
            package.opc.get_part(&web_settings_uri).unwrap().blob(),
            before
        );
    }

    #[test]
    fn edits_web_settings_without_rewriting_document_content() {
        use litchi_docx::web::{Conformance, Screen};

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let (mut settings, conformance) = package.web().unwrap().unwrap();
        settings
            .set_allow_png(false)
            .set_optimize_for_browser(true)
            .set_target_screen_size(Screen::Pixels1600x1200);
        assert_eq!(conformance, Conformance::Transitional);
        assert!(package.put_web(settings, conformance).unwrap());
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let (settings, conformance) = reopened.document().unwrap().web().unwrap().unwrap();
        assert_eq!(conformance, Conformance::Transitional);
        assert_eq!(settings.allow_png(), Some(false));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.target_screen_size(), Some(Screen::Pixels1600x1200));
    }

    #[test]
    fn web_settings_updates_preserve_frame_relationship_ids() {
        use litchi_docx::web::{Frameset, Layout};

        const FRAME_RELATIONSHIP: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame";
        let mut package = Package::new().unwrap();
        let web_settings_uri = PackURI::new("/word/webSettings.xml").unwrap();
        let relationship_id = package
            .opc
            .get_part_mut(&web_settings_uri)
            .unwrap()
            .relate_to("frame1.html", FRAME_RELATIONSHIP);

        let mut frameset = Frameset::default();
        frameset.set_layout(Layout::Rows);
        frameset
            .add_frame()
            .unwrap()
            .set_name("main")
            .unwrap()
            .set_rel(&relationship_id)
            .unwrap();
        let (mut settings, conformance) = package.web().unwrap().unwrap();
        settings.set_frameset(frameset);
        assert!(package.put_web(settings, conformance).unwrap());
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let part = package.opc.get_part(&web_settings_uri).unwrap();
        let relationship = part.rels().get(&relationship_id).unwrap();
        assert_eq!(relationship.reltype(), FRAME_RELATIONSHIP);
        assert_eq!(relationship.target_ref(), "frame1.html");
        assert!(package.document().unwrap().web().unwrap().is_some());
    }

    #[test]
    fn creates_a_web_settings_relationship_when_missing() {
        use litchi_docx::web::{Conformance, Settings};
        use litchi_opc::constants::relationship_type as rt;

        let mut package = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        assert!(package.remove_web().unwrap());
        let mut settings = Settings::default();
        settings.set_encoding("utf-8").unwrap();
        assert!(
            package
                .put_web(settings, Conformance::Transitional)
                .unwrap()
        );
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let relationship = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::WEB_SETTINGS)
            .unwrap();
        assert_eq!(relationship.target_ref(), "webSettings.xml");
        assert_eq!(
            package
                .document()
                .unwrap()
                .web()
                .unwrap()
                .unwrap()
                .0
                .encoding(),
            Some("utf-8")
        );
    }

    #[test]
    fn reads_and_updates_strict_web_settings_relationships() {
        use litchi_docx::web::{Conformance, Settings};
        use litchi_opc::constants::relationship_type as rt;

        let mut package = Package::new().unwrap();
        assert!(package.remove_web().unwrap());
        let (relationship_id, target_ref) = {
            let relationship = package
                .opc
                .rels()
                .part_with_reltype(rt::OFFICE_DOCUMENT)
                .unwrap();
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
            )
        };
        package.opc.rels_mut().remove(&relationship_id);
        package.opc.rels_mut().add_relationship(
            rt::STRICT_OFFICE_DOCUMENT.to_owned(),
            target_ref,
            relationship_id.clone(),
            false,
        );

        let mut settings = Settings::default();
        settings.set_save_smart_tags_as_xml(true);
        assert!(package.put_web(settings, Conformance::Strict).unwrap());
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let (_, conformance) = package.document().unwrap().web().unwrap().unwrap();
        assert_eq!(conformance, Conformance::Strict);
        assert_eq!(
            package
                .document()
                .unwrap()
                .web()
                .unwrap()
                .unwrap()
                .0
                .save_smart_tags_as_xml(),
            Some(true)
        );
    }

    #[test]
    fn rejects_ambiguous_or_external_web_settings_relationships() {
        use litchi_docx::web::{Conformance, Settings};
        use litchi_opc::constants::relationship_type as rt;

        let mut duplicate = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        duplicate
            .opc
            .get_part_mut(&doc_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                Conformance::Strict.relationship().to_owned(),
                "webSettings.xml".to_owned(),
                "rIdDuplicateWebSettings".to_owned(),
                false,
            );
        assert!(duplicate.document().unwrap().web().is_err());
        assert!(duplicate.web().is_err());
        assert!(
            duplicate
                .put_web(Settings::default(), Conformance::Transitional)
                .is_err()
        );

        let mut external = Package::new().unwrap();
        let relationship_id = external
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::WEB_SETTINGS)
            .unwrap()
            .r_id()
            .to_owned();
        let document_part = external.opc.get_part_mut(&doc_uri).unwrap();
        document_part.rels_mut().remove(&relationship_id);
        document_part.rels_mut().add_relationship(
            rt::WEB_SETTINGS.to_owned(),
            "https://example.invalid/webSettings.xml".to_owned(),
            relationship_id,
            true,
        );
        assert!(external.document().unwrap().web().is_err());
        assert!(external.web().is_err());
    }

    fn settings_state(package: &Package) -> (Vec<u8>, String) {
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        (part.blob().to_vec(), part.rels().to_xml())
    }

    #[test]
    fn adds_replaces_removes_and_reopens_attached_template() {
        use crate::docx::settings::is_attached_template_relationship;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package
            .set_attached_template_uri("file:///templates/Corporate.dotx")
            .unwrap();
        let attached = package.attached_template().unwrap().unwrap();
        let relationship_id = attached.relationship_id().to_owned();
        assert_eq!(attached.target_uri(), "file:///templates/Corporate.dotx");

        package
            .set_attached_template_uri("https://example.test/New.dotx?a=1&b=2")
            .unwrap();
        let replacement = package.attached_template().unwrap().unwrap();
        assert_eq!(replacement.relationship_id(), relationship_id);
        assert_eq!(
            replacement.target_uri(),
            "https://example.test/New.dotx?a=1&b=2"
        );
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let settings_part = package.opc.get_part(&settings_uri).unwrap();
        assert_eq!(
            settings_part
                .rels()
                .iter()
                .filter(|relationship| is_attached_template_relationship(relationship.reltype()))
                .count(),
            1
        );

        package.save(file.path()).unwrap();
        let mut reopened = Package::open(file.path()).unwrap();
        assert_eq!(
            reopened.attached_template().unwrap().unwrap().target_uri(),
            "https://example.test/New.dotx?a=1&b=2"
        );
        let removed = reopened.remove_attached_template().unwrap().unwrap();
        assert_eq!(removed.relationship_id(), relationship_id);
        assert!(reopened.attached_template().unwrap().is_none());
        let part = reopened.opc.get_part(&settings_uri).unwrap();
        assert!(!String::from_utf8_lossy(part.blob()).contains("attachedTemplate"));
        assert!(
            !part
                .rels()
                .iter()
                .any(|relationship| is_attached_template_relationship(relationship.reltype()))
        );
    }

    #[test]
    fn attached_template_mutation_preserves_unrelated_xml_and_relationships() {
        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part_mut(&settings_uri).unwrap();
        part.set_blob(br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque></q:settings>"#.to_vec());
        part.rels_mut().add_relationship(
            "urn:unrelated".to_owned(),
            "https://example.test/keep?a=1&b=2".to_owned(),
            "customRelationship".to_owned(),
            true,
        );

        package
            .set_attached_template_uri("file:///templates/Keep.dotx")
            .unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let xml = String::from_utf8_lossy(part.blob());
        assert!(xml.contains(
            r#"<!--keep--><q:zoom q:percent="137"/><x:opaque><![CDATA[a < b]]></x:opaque>"#
        ));
        let unrelated = part.rels().get("customRelationship").unwrap();
        assert_eq!(unrelated.reltype(), "urn:unrelated");
        assert_eq!(unrelated.target_ref(), "https://example.test/keep?a=1&b=2");
    }

    #[test]
    fn reads_strict_prefixed_attached_template_and_rewrites_word_compatible_type() {
        use crate::docx::settings::STRICT_ATTACHED_TEMPLATE_RELATIONSHIP;

        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        let part = package.opc.get_part_mut(&settings_uri).unwrap();
        part.set_blob(br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:attachedTemplate rel:id="arbitrary-id"/></s:settings>"#.to_vec());
        part.rels_mut().add_relationship(
            STRICT_ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
            "file:///strict.dotx".to_owned(),
            "arbitrary-id".to_owned(),
            true,
        );
        assert_eq!(
            package.attached_template().unwrap().unwrap().target_uri(),
            "file:///strict.dotx"
        );

        package
            .set_attached_template_uri("file:///compatible.dotx")
            .unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let relationship = part.rels().get("arbitrary-id").unwrap();
        assert_eq!(relationship.reltype(), ATTACHED_TEMPLATE_RELATIONSHIP);
        assert!(relationship.is_external());
        assert!(
            String::from_utf8_lossy(part.blob())
                .contains(r#"<s:attachedTemplate rel:id="arbitrary-id"/>"#)
        );
    }

    #[test]
    fn attached_template_failures_are_atomic() {
        let mut invalid_target = Package::new().unwrap();
        let before = settings_state(&invalid_target);
        assert!(
            invalid_target
                .set_attached_template_uri("file:///bad path.dotx")
                .is_err()
        );
        assert_eq!(settings_state(&invalid_target), before);

        let mut malformed = Package::new().unwrap();
        malformed
            .set_attached_template_uri("file:///valid.dotx")
            .unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        malformed
            .opc
            .get_part_mut(&settings_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
                "file:///duplicate.dotx".to_owned(),
                "duplicate-id".to_owned(),
                true,
            );
        let before = settings_state(&malformed);
        assert!(
            malformed
                .set_attached_template_uri("file:///replacement.dotx")
                .is_err()
        );
        assert_eq!(settings_state(&malformed), before);
        assert!(malformed.remove_attached_template().is_err());
        assert_eq!(settings_state(&malformed), before);
    }

    #[test]
    fn protection_rewrite_preserves_attached_template_relationship() {
        use crate::docx::settings::ProtectionType;

        let mut package = Package::new().unwrap();
        package
            .set_attached_template_uri("file:///templates/Protected.dotx")
            .unwrap();
        let relationship_id = package
            .attached_template()
            .unwrap()
            .unwrap()
            .relationship_id()
            .to_owned();
        package
            .document_mut()
            .unwrap()
            .set_protection(ProtectionType::ReadOnly);
        package.to_stream(Cursor::new(Vec::new())).unwrap();

        let attached = package.attached_template().unwrap().unwrap();
        assert_eq!(attached.relationship_id(), relationship_id);
        assert_eq!(attached.target_uri(), "file:///templates/Protected.dotx");
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        assert!(
            package
                .opc
                .get_part(&settings_uri)
                .unwrap()
                .rels()
                .get(&relationship_id)
                .is_some()
        );
    }

    #[test]
    fn document_variable_package_lifecycle_is_deterministic_and_reopens() {
        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        assert_eq!(
            package
                .set_document_variable("Company & Team", "A < B")
                .unwrap(),
            None
        );
        package.set_document_variable("second", "two").unwrap();
        assert_eq!(
            package
                .set_document_variable("Company & Team", "updated")
                .unwrap(),
            Some("A < B".into())
        );
        assert_eq!(
            package.document_variables().unwrap().unwrap().names(),
            vec!["Company & Team", "second"]
        );
        package.save(file.path()).unwrap();

        let mut reopened = Package::open(file.path()).unwrap();
        let variables = reopened
            .document()
            .unwrap()
            .document_variables()
            .unwrap()
            .unwrap();
        assert_eq!(variables.get("Company & Team"), Some("updated"));
        assert_eq!(variables.get("second"), Some("two"));
        assert_eq!(
            reopened.remove_document_variable("Company & Team").unwrap(),
            Some("updated".into())
        );
        assert_eq!(reopened.clear_document_variables().unwrap(), 1);
        assert!(reopened.document_variables().unwrap().unwrap().is_empty());
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        assert!(
            !String::from_utf8_lossy(reopened.opc.get_part(&settings_uri).unwrap().blob())
                .contains("docVars")
        );
    }

    #[test]
    fn document_variable_mutation_preserves_xml_relationships_and_protection() {
        use crate::docx::settings::ProtectionType;

        let mut package = Package::new().unwrap();
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<?xml version="1.0"?><q:settings xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:opaque"><!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#.to_vec(),
        );
        package
            .opc
            .get_part_mut(&settings_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:unrelated".to_owned(),
                "https://example.test/keep".to_owned(),
                "keep-id".to_owned(),
                true,
            );
        package
            .set_attached_template_uri("file:///templates/Variables.dotx")
            .unwrap();
        let attached_id = package
            .attached_template()
            .unwrap()
            .unwrap()
            .relationship_id()
            .to_owned();
        package.set_document_variable("project", "Litchi").unwrap();
        let part = package.opc.get_part(&settings_uri).unwrap();
        let xml = String::from_utf8_lossy(part.blob());
        assert!(xml.contains(
            r#"<!--keep--><q:zoom q:percent="133"/><x:opaque><![CDATA[a < b]]></x:opaque>"#
        ));
        assert!(part.rels().get("keep-id").is_some());
        assert!(part.rels().get(&attached_id).is_some());

        package
            .document_mut()
            .unwrap()
            .set_protection(ProtectionType::ReadOnly);
        package.to_stream(Cursor::new(Vec::new())).unwrap();
        assert_eq!(
            package
                .document_variables()
                .unwrap()
                .unwrap()
                .get("project"),
            Some("Litchi")
        );
        let part = package.opc.get_part(&settings_uri).unwrap();
        assert!(part.rels().get("keep-id").is_some());
        assert!(part.rels().get(&attached_id).is_some());
    }

    #[test]
    fn document_variable_mutation_failures_are_atomic() {
        let mut package = Package::new().unwrap();
        let before = settings_state(&package);
        assert!(package.set_document_variable("", "invalid").is_err());
        assert_eq!(settings_state(&package), before);
        assert!(
            package
                .set_document_variable("too-long", "x".repeat(65_281))
                .is_err()
        );
        assert_eq!(settings_state(&package), before);

        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVars><w:docVar w:name="duplicate" w:val="one"/><w:docVar w:name="duplicate" w:val="two"/></w:docVars></w:settings>"#.to_vec(),
        );
        let malformed = settings_state(&package);
        assert!(package.set_document_variable("new", "value").is_err());
        assert_eq!(settings_state(&package), malformed);
        assert!(package.remove_document_variable("duplicate").is_err());
        assert_eq!(settings_state(&package), malformed);
        assert!(package.clear_document_variables().is_err());
        assert_eq!(settings_state(&package), malformed);
    }

    #[test]
    fn reads_strict_settings_relationship_and_mce_fallback_variables() {
        use litchi_opc::constants::relationship_type as rt;

        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let mut package = Package::new().unwrap();
        let doc_uri = PackURI::new("/word/document.xml").unwrap();
        let relationship_id = package
            .opc
            .get_part(&doc_uri)
            .unwrap()
            .rels()
            .part_with_reltype(rt::SETTINGS)
            .unwrap()
            .r_id()
            .to_owned();
        let document = package.opc.get_part_mut(&doc_uri).unwrap();
        let target = document
            .rels_mut()
            .remove(&relationship_id)
            .unwrap()
            .target_ref()
            .to_owned();
        document.rels_mut().add_relationship(
            STRICT_SETTINGS_RELATIONSHIP.to_owned(),
            target,
            relationship_id,
            false,
        );
        let settings_uri = PackURI::new("/word/settings.xml").unwrap();
        package.opc.get_part_mut(&settings_uri).unwrap().set_blob(
            br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><s:docVars><s:docVar s:name="choice" s:val="ignored"/></s:docVars></mc:Choice><mc:Fallback><s:docVars><s:docVar s:name="fallback" s:val="selected"/></s:docVars></mc:Fallback></mc:AlternateContent></s:settings>"#.to_vec(),
        );

        let package_variables = package.document_variables().unwrap().unwrap();
        assert_eq!(package_variables.get("fallback"), Some("selected"));
        assert!(!package_variables.contains("choice"));
        let document_variables = package
            .document()
            .unwrap()
            .document_variables()
            .unwrap()
            .unwrap();
        assert_eq!(document_variables.get("fallback"), Some("selected"));
    }
}

#[cfg(feature = "encryption")]
use crate::encryption::{Limits, Mode};
use crate::error::{OoxmlError, Result};
use crate::pptx::parts::PresentationPart;
use crate::pptx::presentation::{PptxChart, Presentation};
use crate::pptx::show_events::PptxSlideShowEvent;
use crate::pptx::slide::Key as SlideKey;
use crate::pptx::vba_project::{
    VbaProject, discover_vba_project, remove_vba_project as clear_presentation_vba,
    store_vba_project as store_presentation_vba_project,
};
use crate::pptx::writer::MutablePresentation;
use litchi_ooxml_common::embedded;
/// Package implementation for PowerPoint presentations.
use litchi_ooxml_common::properties::{Props, Slot};
use litchi_ooxml_common::ribbon;
use litchi_ooxml_common::web;
use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part;
use litchi_pptx::tag;
use std::io::{Read, Seek};
use std::path::Path;

pub(crate) const STALE_NOTES_REASON: &str = "the legacy writer has unflushed changes that could replace slide and notes relationships; save and reopen before reading or editing notes";
const STALE_SLIDE_GRAPH_REASON: &str = "the legacy writer has unflushed changes that could replace slide parts or relationships; save and reopen before editing the canonical slide graph";
const STALE_TAG_GRAPH_REASON: &str = "the legacy writer has unflushed changes that could replace slide relationships; save and reopen before editing tags";
const STALE_PRESENTATION_GRAPH_REASON: &str = "the legacy writer has unflushed changes that could replace presentation parts or relationships; publish before switching to canonical graph editing";

/// Default media poster image - a simple 1x1 gray PNG.
/// This is used as a placeholder for media shapes that don't have a custom poster frame.
/// It's a valid minimal PNG image (67 bytes).
const DEFAULT_MEDIA_POSTER: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 pixel
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
    0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
    0x08, 0xD7, 0x63, 0x78, 0x78, 0x78, 0x00, 0x00, // Compressed gray pixel
    0x00, 0x85, 0x00, 0x82, 0x3E, 0x8F, 0xFE, 0xB6, // CRC
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
    0xAE, 0x42, 0x60, 0x82, // CRC
];

fn validate_presentation_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        return Ok(());
    }

    Err(OoxmlError::InvalidContentType {
        expected: format!(
            "{}, {}, {}, {}, {}, or {}",
            ct::PML_PRESENTATION_MAIN,
            ct::PML_SLIDESHOW_MAIN,
            ct::PML_TEMPLATE_MAIN,
            ct::PML_PRES_MACRO_MAIN,
            ct::PML_SLIDESHOW_MACRO_MAIN,
            ct::PML_TEMPLATE_MACRO_MAIN,
        ),
        got: content_type.to_string(),
    })
}

/// A PowerPoint (.pptx, .pptm, .ppsm, or .potm) package.
///
/// This is the main entry point for working with PowerPoint presentations.
/// It wraps an OPC package and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::pptx::Package;
///
/// // Open an existing presentation
/// let pkg = Package::open("presentation.pptx")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Access slides
/// println!("Presentation has {} slides", pres.slide_count()?);
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
    /// Mutable presentation for writing (cached)
    mutable_pres: Option<MutablePresentation>,
    /// Authoritative, mutation-tracked core properties.
    properties: Slot,
    /// Whether managed publication must synchronize embedded-font state.
    #[cfg(feature = "fonts")]
    font_sync_pending: bool,
    /// Encryption profile of the opened outer package.
    #[cfg(feature = "encryption")]
    source_encryption: Option<Mode>,
}

#[cfg(feature = "fonts")]
use crate::fonts::{EmbedFonts, powerpoint_data, prepare_fonts};
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
        let presentation = self.mutable_pres.as_ref().ok_or(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "embed_fonts",
            reason: "font discovery is unavailable for an opened presentation without a complete mutable model",
        })?;
        let glyphs = presentation.collect_glyphs();
        let prepared = prepare_fonts(glyphs, subset)?;
        if prepared.is_empty() {
            return Ok(());
        }
        let mut fonts = litchi_pptx::font::load(&self.opc)?.unwrap_or_default();
        let mut changed = false;
        for mut prepared in prepared {
            let current = match fonts.get(prepared.name.as_str()) {
                Ok(font) => Some(font.clone()),
                Err(litchi_pptx::Error::FontNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            let mut font = match &current {
                Some(font) => font.clone(),
                None => litchi_pptx::font::Font::new(&prepared.name)?,
            };
            font.set_panose(Some(litchi_pptx::font::Panose::new(
                prepared.properties.panose().into_bytes(),
            )));
            font.set_pitch_family(Some(litchi_pptx::font::PitchFamily::new(
                prepared_pitch(prepared.properties.pitch()),
                prepared_family(prepared.properties.family()),
            )));
            font.set_charset(
                prepared
                    .properties
                    .charset()
                    .map(|value| litchi_pptx::font::Charset::new(value.code())),
            );

            let style = prepared_style(prepared.style);
            let data = litchi_pptx::font::Data::powerpoint(powerpoint_data(&mut prepared)?)?;
            if font.get(style).is_none_or(|face| face.data() != &data) {
                let _ = font.put(litchi_pptx::font::Face::new(style, data))?;
            }
            if current.as_ref() == Some(&font) {
                continue;
            }
            match current {
                Some(_) => {
                    let _ = fonts.replace(prepared.name.as_str(), font)?;
                },
                None => fonts.add(font)?,
            }
            changed = true;
        }
        if changed {
            let _ = litchi_pptx::font::put(&mut self.opc, fonts)?;
        }
        Ok(())
    }
}

#[cfg(feature = "fonts")]
fn prepared_style(value: litchi_fonts::Style) -> litchi_pptx::font::Style {
    match value {
        litchi_fonts::Style::Regular => litchi_pptx::font::Style::Regular,
        litchi_fonts::Style::Bold => litchi_pptx::font::Style::Bold,
        litchi_fonts::Style::Italic => litchi_pptx::font::Style::Italic,
        litchi_fonts::Style::BoldItalic => litchi_pptx::font::Style::BoldItalic,
    }
}

#[cfg(feature = "fonts")]
fn prepared_pitch(value: litchi_fonts::Pitch) -> litchi_pptx::font::Pitch {
    match value {
        litchi_fonts::Pitch::Fixed => litchi_pptx::font::Pitch::Fixed,
        litchi_fonts::Pitch::Variable => litchi_pptx::font::Pitch::Variable,
        litchi_fonts::Pitch::Default => litchi_pptx::font::Pitch::Default,
    }
}

#[cfg(feature = "fonts")]
fn prepared_family(value: litchi_fonts::Family) -> litchi_pptx::font::Family {
    match value {
        litchi_fonts::Family::Roman => litchi_pptx::font::Family::Roman,
        litchi_fonts::Family::Swiss => litchi_pptx::font::Family::Swiss,
        litchi_fonts::Family::Modern => litchi_pptx::font::Family::Modern,
        litchi_fonts::Family::Script => litchi_pptx::font::Family::Script,
        litchi_fonts::Family::Decorative => litchi_pptx::font::Family::Decorative,
        litchi_fonts::Family::Auto => litchi_pptx::font::Family::None,
    }
}

impl Package {
    /// Create a new empty .pptx package.
    ///
    /// Creates a minimal valid PowerPoint presentation with default master slide and layout.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Add slides to the presentation...
    /// pkg.save("new_presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::pptx::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::packuri::PackURI;
        use litchi_opc::part::BlobPart;

        let mut opc = OpcPackage::new();

        // Create presentation.xml part
        let pres_partname = PackURI::new("/ppt/presentation.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presentation partname: {}", e)))?;
        let pres_part = BlobPart::new(
            pres_partname.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            template::default_presentation_xml().as_bytes().to_vec(),
        );

        // Create relationship from package to presentation (use relative path for package-level rels)
        opc.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
        opc.add_part(Box::new(pres_part));

        // Create slideMaster.xml
        let master_partname = PackURI::new("/ppt/slideMasters/slideMaster1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("slideMaster partname: {}", e)))?;
        let master_part = BlobPart::new(
            master_partname.clone(),
            ct::PML_SLIDE_MASTER.to_string(),
            template::default_slide_master_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to slideMaster
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);
        }
        opc.add_part(Box::new(master_part));

        // Create all 11 slide layouts
        // Each layout needs:
        // 1. A relationship FROM slideMaster TO the layout
        // 2. A relationship FROM the layout back TO slideMaster
        let layout_xmls = template::all_slide_layouts();
        for (i, layout_xml) in layout_xmls.iter().enumerate() {
            let layout_num = i + 1;
            let layout_partname_str = format!("/ppt/slideLayouts/slideLayout{}.xml", layout_num);
            let layout_partname = PackURI::new(&layout_partname_str).map_err(|e| {
                OoxmlError::InvalidUri(format!("slideLayout{} partname: {}", layout_num, e))
            })?;

            let mut layout_part = BlobPart::new(
                layout_partname.clone(),
                ct::PML_SLIDE_LAYOUT.to_string(),
                layout_xml.as_bytes().to_vec(),
            );

            // Add relationship from slideMaster to this slideLayout
            if let Ok(master_part) = opc.get_part_mut(&master_partname) {
                let layout_rel_target = format!("../slideLayouts/slideLayout{}.xml", layout_num);
                master_part.relate_to(&layout_rel_target, rt::SLIDE_LAYOUT);
            }

            // Add relationship from slideLayout back to slideMaster
            // This bidirectional relationship is required by PowerPoint
            layout_part.relate_to("../slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

            opc.add_part(Box::new(layout_part));
        }

        // Create theme.xml
        let theme_partname = PackURI::new("/ppt/theme/theme1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("theme partname: {}", e)))?;
        let theme_part = BlobPart::new(
            theme_partname,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );

        // Add relationship from slideMaster to theme
        if let Ok(master_part) = opc.get_part_mut(&master_partname) {
            master_part.relate_to("../theme/theme1.xml", rt::THEME);
        }
        opc.add_part(Box::new(theme_part));

        // Notes masters require their own theme part. Sharing the slide-master
        // theme makes desktop PowerPoint repair the package by creating this
        // second part and repointing the notes master.
        let notes_theme_partname = PackURI::new("/ppt/theme/theme2.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("notes theme partname: {e}")))?;
        let notes_theme_part = BlobPart::new(
            notes_theme_partname,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );
        opc.add_part(Box::new(notes_theme_part));

        // Create tableStyles.xml
        let table_styles_partname = PackURI::new("/ppt/tableStyles.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("tableStyles partname: {}", e)))?;
        let table_styles_part = BlobPart::new(
            table_styles_partname,
            ct::PML_TABLE_STYLES.to_string(),
            litchi_pptx::table::style::default_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to tableStyles
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("tableStyles.xml", rt::TABLE_STYLES);
        }
        opc.add_part(Box::new(table_styles_part));

        // Create viewProps.xml
        let view_props_partname = PackURI::new("/ppt/viewProps.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("viewProps partname: {}", e)))?;
        let view_props_part = BlobPart::new(
            view_props_partname,
            ct::PML_VIEW_PROPS.to_string(),
            template::default_view_props_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to viewProps
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("viewProps.xml", rt::VIEW_PROPS);
        }
        opc.add_part(Box::new(view_props_part));

        // Create presProps.xml
        let pres_props_partname = PackURI::new("/ppt/presProps.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presProps partname: {}", e)))?;
        let pres_props_part = BlobPart::new(
            pres_props_partname,
            ct::PML_PRES_PROPS.to_string(),
            template::default_pres_props_xml().as_bytes().to_vec(),
        );

        // Add relationship from presentation to presProps
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("presProps.xml", rt::PRES_PROPS);
        }
        opc.add_part(Box::new(pres_props_part));

        // Create notesMaster.xml
        let notes_master_partname = PackURI::new("/ppt/notesMasters/notesMaster1.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("notesMaster partname: {}", e)))?;
        let mut notes_master_part = BlobPart::new(
            notes_master_partname.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
                .to_string(),
            litchi_pptx::notes::master_xml().as_bytes().to_vec(),
        );

        // Add relationship from notesMaster to theme
        notes_master_part.relate_to("../theme/theme2.xml", rt::THEME);

        // Add relationship from presentation to notesMaster and retain its
        // relationship ID in the required presentation-root reference.
        let notes_master_relationship_id = opc
            .get_part_mut(&pres_partname)
            .map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "default presentation part for notes master: {error}"
                ))
            })?
            .relate_to("notesMasters/notesMaster1.xml", rt::NOTES_MASTER);
        {
            let presentation = opc.get_part_mut(&pres_partname).map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "default presentation part for notes-master reference: {error}"
                ))
            })?;
            let xml = std::str::from_utf8(presentation.blob()).map_err(|error| {
                OoxmlError::InvalidFormat(format!("default presentation XML is not UTF-8: {error}"))
            })?;
            let marker = "</p:sldMasterIdLst>";
            let replacement = format!(
                "{marker}<p:notesMasterIdLst><p:notesMasterId r:id=\"{notes_master_relationship_id}\"/></p:notesMasterIdLst>"
            );
            let updated = xml.replacen(marker, &replacement, 1);
            if updated == xml {
                return Err(OoxmlError::InvalidFormat(
                    "default presentation XML is missing its slide-master list".to_string(),
                ));
            }
            presentation.set_blob(updated.into_bytes());
        }
        opc.add_part(Box::new(notes_master_part));

        // Create core.xml (core properties)
        let core_props_partname = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml partname: {}", e)))?;
        let core_props_part = BlobPart::new(
            core_props_partname,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        opc.add_part(Box::new(core_props_part));

        // Create app.xml (extended properties)
        let app_props_partname = PackURI::new("/docProps/app.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("app.xml partname: {}", e)))?;
        let app_props_part = BlobPart::new(
            app_props_partname,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );

        opc.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        opc.add_part(Box::new(app_props_part));

        // Create a mutable presentation for writing
        let mutable_pres = Some(MutablePresentation::new());

        // Initialize document properties
        let properties = Slot::load(&opc)?;

        Ok(Self {
            opc,
            mutable_pres,
            properties,
            #[cfg(feature = "fonts")]
            font_sync_pending: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Open a .pptx, .pptm, .ppsm, or .potm package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .pptx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        validate_presentation_main_content_type(main_part.content_type())?;

        let properties = Slot::load(&opc)?;
        Ok(Self {
            opc,
            mutable_pres: None,
            properties,
            #[cfg(feature = "fonts")]
            font_sync_pending: false,
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
    /// use litchi_ooxml::{OpcPackage, pptx::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.pptx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        validate_presentation_main_content_type(main_part.content_type())?;

        let properties = Slot::load(&opc)?;
        Ok(Self {
            opc,
            mutable_pres: None,
            properties,
            #[cfg(feature = "fonts")]
            font_sync_pending: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Create a .pptx, .pptm, .ppsm, or .potm package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .pptx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("presentation.pptx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        validate_presentation_main_content_type(main_part.content_type())?;

        let properties = Slot::load(&opc)?;
        Ok(Self {
            opc,
            mutable_pres: None,
            properties,
            #[cfg(feature = "fonts")]
            font_sync_pending: false,
            #[cfg(feature = "encryption")]
            source_encryption: None,
        })
    }

    /// Open a reader with a password and safe default resource limits.
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

    /// Get the main presentation for reading.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// // Access slides
    /// for slide in pres.slides()? {
    ///     println!("Slide text: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        let notes_current = !self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified);
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        // Create PresentationPart wrapper
        let pres_part = PresentationPart::from_part(main_part)?;

        // Create and return Presentation
        Ok(Presentation::new(pres_part, &self.opc, notes_current))
    }

    /// Discover all native chart parts reachable from presentation slides.
    ///
    /// The returned items retain their owning slide, relationship ID, and OPC
    /// part name alongside basic inert chart metadata.
    pub fn charts(&self) -> Result<Vec<PptxChart>> {
        self.presentation()?.charts()
    }

    /// Discover inert InkML annotation content parts on presentation slides.
    ///
    /// Ink payloads are never rendered, recognized, interpreted, or executed.
    pub fn ink_annotations(&self) -> Result<Vec<crate::pptx::PptxInkAnnotation>> {
        self.presentation()?.ink_annotations()
    }

    /// Store an InkML annotation onto a slide as an inert `p:contentPart`.
    ///
    /// The payload is validated as InkML and stored verbatim under
    /// `/ppt/ink/`; the slide gains a `customXml` relationship and a
    /// `p:contentPart` reference at the end of its shape tree, preserving
    /// the slide's namespace dialect. The ink is never rendered, recognized,
    /// or executed.
    /// A successful edit disables subsequent legacy-writer access on this value.
    pub fn add_ink_annotation(
        &mut self,
        slide_name: &PackURI,
        inkml: &[u8],
    ) -> Result<crate::pptx::StoredInkAnnotation> {
        self.edit_canonical("add_ink_annotation", STALE_SLIDE_GRAPH_REASON, |package| {
            crate::pptx::store_slide_ink_annotation(package, slide_name, inkml)
        })
    }

    /// Discover persisted laser-pointer traces from presentation slides.
    ///
    /// Trace points are returned as inert stored data and are never replayed,
    /// rendered, interpolated, modified, or executed.
    pub fn laser_traces(&self) -> Result<Vec<crate::pptx::PptxLaserTrace>> {
        self.presentation()?.laser_traces()
    }

    /// Store one laser-pointer trace onto a slide as a PowerPoint 2010
    /// `p14:laserTraceLst` extension.
    ///
    /// Typed points are serialized canonically; the slide gains the
    /// `p:ext` extension block (creating `p:extLst` when absent) while
    /// preserving its namespace dialect. Slides that already carry a laser
    /// extension are rejected. Traces are never replayed, rendered,
    /// interpolated, or executed.
    /// A successful edit disables subsequent legacy-writer access on this value.
    pub fn add_laser_trace(
        &mut self,
        slide_name: &PackURI,
        points: &[crate::pptx::PptxLaserTracePoint],
    ) -> Result<()> {
        self.edit_canonical("add_laser_trace", STALE_SLIDE_GRAPH_REASON, |package| {
            crate::pptx::store_slide_laser_trace(package, slide_name, points)
        })
    }

    /// Discover persisted slide-show event records from presentation slides.
    ///
    /// Event records are returned as inert historical metadata only. This
    /// never replays triggers, seeks media, opens targets, or changes slide-show state.
    pub fn show_events(&self) -> Result<Vec<PptxSlideShowEvent>> {
        self.presentation()?.show_events()
    }

    /// Store slide-show event records onto a slide as a PowerPoint 2010
    /// `p14:showEvtLst` extension.
    ///
    /// Typed events are serialized canonically in caller order; the
    /// slide gains the `p:ext` extension block while preserving its
    /// namespace dialect. Slides that already carry a show-event extension
    /// are rejected. Events are never replayed, rendered, or executed.
    /// A successful edit disables subsequent legacy-writer access on this value.
    pub fn add_slide_show_events(
        &mut self,
        slide_name: &PackURI,
        events: &[crate::pptx::PptxSlideShowEventDraft],
    ) -> Result<()> {
        self.edit_canonical(
            "add_slide_show_events",
            STALE_SLIDE_GRAPH_REASON,
            |package| crate::pptx::store_slide_show_events(package, slide_name, events),
        )
    }

    /// Discover bounded, inert click and hover action settings on slides.
    ///
    /// Declared targets are never followed, opened, activated, or executed.
    pub fn action_settings(&self) -> Result<Vec<crate::pptx::PptxActionSetting>> {
        self.presentation()?.action_settings()
    }

    /// Discover bounded, inert OLE object shapes and declared payload targets.
    ///
    /// This never parses, opens, activates, renders, or executes an embedded
    /// object or package payload.
    pub fn ole_objects(&self) -> Result<Vec<crate::pptx::PptxOleObject>> {
        self.presentation()?.ole_objects()
    }

    /// Discover bounded, inert slide controls (ActiveX/OCX) and their
    /// resolved controls-part descriptors.
    ///
    /// This never instantiates a control, resolves a CLSID, decodes binary
    /// control state, executes a macro, or follows an external relationship.
    pub fn controls(&self) -> Result<Vec<crate::pptx::PptxSlideControl>> {
        self.presentation()?.controls()
    }

    /// Read the programmable-tag list attached directly to one selected slide.
    ///
    /// A producer-visible slide name is the ordinary selector and a checked
    /// zero-based position is available for ordered workflows. Relationship
    /// IDs and part names remain below this facade. The returned list owns its
    /// bounded inert strings and retained extension attributes; tag values are
    /// never interpreted.
    ///
    /// Shape-owned tag lists are intentionally not flattened into this result;
    /// use the lower-level slide inventory when inspecting producer markup.
    pub fn tags<'a>(&self, slide: impl Into<SlideKey<'a>>) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(tag::load(&self.opc, &slide_name)?.map(tag::Source::into_list))
    }

    /// Put the direct programmable-tag list on one selected slide.
    ///
    /// The common path selects the slide by its producer-visible name. The
    /// list is consumed by value and encoded without exposing relationship IDs
    /// or part names.
    ///
    /// Existing content is replaced atomically and returned by value. `None`
    /// means the selected slide did not previously own a direct tag list.
    pub fn put_tags<'a>(
        &mut self,
        slide: impl Into<SlideKey<'a>>,
        list: tag::List,
    ) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("put_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_canonical("put_tags", STALE_TAG_GRAPH_REASON, move |package| {
            tag::put(package, &slide_name, list).map_err(Into::into)
        })
    }

    /// Remove the direct tag list from one selected slide.
    ///
    /// Orphaned parts are collected only after a package-wide inbound-edge
    /// scan proves that no other owner still references them.
    ///
    /// Returns `None` when the slide has no direct tag list, making repeated
    /// removal safe and idempotent.
    pub fn remove_tags<'a>(&mut self, slide: impl Into<SlideKey<'a>>) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("remove_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_canonical("remove_tags", STALE_TAG_GRAPH_REASON, move |package| {
            tag::remove(package, &slide_name).map_err(Into::into)
        })
    }

    /// Read the programmable-tag list attached to one semantic slide shape.
    ///
    /// Producer-visible shape names are the ordinary selector. A checked
    /// depth-first numeric position is also accepted for source-order repair;
    /// non-visual IDs and relationship IDs remain below the safe facade.
    pub fn shape_tags<'s, 'k>(
        &self,
        slide: impl Into<SlideKey<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        Ok(tag::shape::load(&self.opc, &slide_name, shape)?.map(tag::Source::into_list))
    }

    /// Create or replace one semantic slide shape's programmable-tag list.
    ///
    /// The list is moved into a staged part. Shape XML, its relationship, and
    /// the target part commit together after source-preserving validation.
    pub fn put_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<SlideKey<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        list: tag::List,
    ) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("put_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_canonical("put_shape_tags", STALE_TAG_GRAPH_REASON, move |package| {
            tag::shape::put(package, &slide_name, shape, list).map_err(Into::into)
        })
    }

    /// Remove one semantic slide shape's programmable-tag list.
    ///
    /// Absence is an idempotent `Ok(None)`. Shared relationships and targets
    /// are retained until no active anchor or package edge uses them.
    pub fn remove_shape_tags<'s, 'k>(
        &mut self,
        slide: impl Into<SlideKey<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<tag::List>> {
        self.ensure_tag_graph_current("remove_shape_tags")?;
        let slide_name = self.resolve_slide(slide.into())?;
        self.edit_canonical(
            "remove_shape_tags",
            STALE_TAG_GRAPH_REASON,
            move |package| tag::shape::remove(package, &slide_name, shape).map_err(Into::into),
        )
    }

    fn ensure_tag_graph_current(&self, operation: &'static str) -> Result<()> {
        self.ensure_canonical_edit_ready(operation, STALE_TAG_GRAPH_REASON)
    }

    fn ensure_notes_graph_current(&self, operation: &'static str) -> Result<()> {
        self.ensure_canonical_edit_ready(operation, STALE_NOTES_REASON)
    }

    fn ensure_canonical_edit_ready(
        &self,
        operation: &'static str,
        stale_reason: &'static str,
    ) -> Result<()> {
        if self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
        {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: stale_reason,
            });
        }

        if self.opc.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "automatic font embedding requires the legacy presentation model; disable the font policy before switching to canonical graph editing",
            });
        }

        #[cfg(feature = "fonts")]
        if self.font_sync_pending {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "font embedding is pending; publish through save or to_bytes before switching to canonical graph editing",
            });
        }

        Ok(())
    }

    fn edit_canonical<T>(
        &mut self,
        operation: &'static str,
        stale_reason: &'static str,
        edit: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.ensure_canonical_edit_ready(operation, stale_reason)?;

        // Built-in part payloads are Arc-backed. Custom `Part` implementations
        // retain their declared clone and interior-mutability policy.
        let mut candidate = self.opc.clone();
        let value = edit(&mut candidate)?;
        self.opc = candidate;
        self.mutable_pres = None;
        Ok(value)
    }

    fn edit_signature<T>(
        &mut self,
        operation: &'static str,
        edit: impl FnOnce(&mut OpcPackage) -> litchi_opc::sign::Result<T>,
    ) -> litchi_opc::sign::Result<T> {
        self.ensure_canonical_edit_ready(operation, STALE_PRESENTATION_GRAPH_REASON)
            .map_err(|error| litchi_opc::sign::Error::Graph(error.to_string()))?;
        if self.properties.is_dirty() {
            return Err(litchi_opc::sign::Error::Graph(format!(
                "PPTX {operation} requires current core properties; publish through save or to_bytes before signing"
            )));
        }

        let mut candidate = self.opc.clone();
        let value = edit(&mut candidate)?;
        self.opc = candidate;
        self.mutable_pres = None;
        Ok(value)
    }

    fn resolve_slide(&self, key: SlideKey<'_>) -> Result<PackURI> {
        match key {
            SlideKey::Index(index) => {
                let presentation = self.presentation()?;
                if let Some(slide) = presentation.slide(index)? {
                    return Ok(slide.part().part().partname().clone());
                }
                Err(litchi_pptx::Error::SlideIndexOutOfBounds {
                    index,
                    len: presentation.slide_count()?,
                }
                .into())
            },
            SlideKey::Name(name) => {
                let slides = self.presentation()?.slides()?;
                let mut selected = None;
                let mut matches = 0usize;
                for slide in &slides {
                    if slide.name()?.as_str() == name.as_ref() {
                        matches = matches.saturating_add(1);
                        if selected.is_none() {
                            selected = Some(slide.part().part().partname().clone());
                        }
                    }
                }
                match (selected, matches) {
                    (Some(part), 1) => Ok(part),
                    (None, 0) => {
                        Err(litchi_pptx::Error::SlideNameNotFound(name.into_owned()).into())
                    },
                    (_, matches) => Err(litchi_pptx::Error::AmbiguousSlideName {
                        name: name.into_owned(),
                        matches,
                    }
                    .into()),
                }
            },
        }
    }

    /// Load typed presentation-view settings, if the package contains them.
    ///
    /// View settings are returned as stored document data only; this does not
    /// alter the application's display state or follow outline-slide targets.
    pub fn view_properties(&self) -> Result<Option<crate::pptx::view_properties::ViewProperties>> {
        crate::pptx::view_properties::load_from_package(&self.opc)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Load typed presentation settings, if the package contains them.
    ///
    /// Declared HTML publishing targets remain inert metadata and are never
    /// opened, fetched, or otherwise activated.
    pub fn presentation_properties(
        &self,
    ) -> Result<Option<crate::pptx::presentation_properties::PresentationProperties>> {
        crate::pptx::presentation_properties::load_from_package(&self.opc)
            .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
    }

    /// Load the presentation's typed embedded-font collection.
    pub fn fonts(&self) -> Result<Option<litchi_pptx::font::Fonts>> {
        Ok(litchi_pptx::font::load(&self.opc)?)
    }

    /// Select the font publication policy used by managed save operations.
    #[cfg(feature = "fonts")]
    pub fn set_font_embedding(
        &mut self,
        embedding: litchi_opc::FontEmbedding,
    ) -> Result<&mut Self> {
        if embedding != litchi_opc::FontEmbedding::None && self.mutable_pres.is_none() {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "set_font_embedding",
                reason: "font discovery requires a complete mutable presentation model",
            });
        }

        if self.opc.save_options().fonts != embedding {
            self.opc.with_fonts(embedding);
            self.font_sync_pending = embedding != litchi_opc::FontEmbedding::None;
        }
        Ok(self)
    }

    /// Select the font publication policy and return this package by value.
    #[cfg(feature = "fonts")]
    pub fn with_font_embedding(mut self, embedding: litchi_opc::FontEmbedding) -> Result<Self> {
        self.set_font_embedding(embedding)?;
        Ok(self)
    }

    /// Atomically publish a complete embedded-font collection by value.
    ///
    /// Returns `false` for an exact no-op, preserving valid signatures.
    pub fn put_fonts(&mut self, fonts: litchi_pptx::font::Fonts) -> Result<bool> {
        self.flush_presentation()?;
        Ok(litchi_pptx::font::put(&mut self.opc, fonts)?)
    }

    /// Atomically remove and return the embedded-font collection.
    pub fn remove_fonts(&mut self) -> Result<Option<litchi_pptx::font::Fonts>> {
        self.flush_presentation()?;
        Ok(litchi_pptx::font::remove(&mut self.opc)?)
    }

    /// Load the presentation's typed, bounded table-style catalog.
    pub fn styles(&self) -> Result<Option<litchi_pptx::table::style::List>> {
        Ok(litchi_pptx::table::style::load(&self.opc)?)
    }

    /// Atomically create or replace the table-style catalog by value.
    ///
    /// Returns `true` when package bytes changed and `false` for a semantic
    /// no-op, allowing callers to retain signatures on unchanged documents.
    pub fn put_styles(&mut self, styles: litchi_pptx::table::style::List) -> Result<bool> {
        Ok(litchi_pptx::table::style::put(&mut self.opc, styles)?)
    }

    /// Atomically remove and return the optional table-style catalog.
    pub fn remove_styles(&mut self) -> Result<Option<litchi_pptx::table::style::List>> {
        Ok(litchi_pptx::table::style::remove(&mut self.opc)?)
    }

    /// Load the PowerPoint Revision Information part, if present.
    ///
    /// Revision extension XML remains inert metadata and is never executed or
    /// used to resolve relationships.
    pub fn revision_information(
        &self,
    ) -> Result<Option<crate::pptx::revision_information::RevisionInformationPart>> {
        crate::pptx::revision_information::load_revision_information(&self.opc)
    }

    /// Add a validated PowerPoint Revision Information part.
    ///
    /// This operation deliberately rejects replacing an existing part.
    pub fn store_revision_information(
        &mut self,
        value: &crate::pptx::revision_information::RevisionInformationPart,
    ) -> Result<()> {
        self.edit_canonical(
            "store_revision_information",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::revision_information::store_revision_information(package, value)?;
                package.unsign();
                Ok(())
            },
        )
    }

    /// Load the PowerPoint Changes Information part, if present.
    ///
    /// Nested change descriptors remain inert XML and are never executed or
    /// used to resolve relationships.
    pub fn changes_information(
        &self,
    ) -> Result<Option<crate::pptx::changes_information::ChangesInformationPart>> {
        crate::pptx::changes_information::load_changes_information(&self.opc)
    }

    /// Add a validated PowerPoint Changes Information part.
    ///
    /// This operation deliberately rejects replacing an existing part.
    pub fn store_changes_information(
        &mut self,
        value: &crate::pptx::changes_information::ChangesInformationPart,
    ) -> Result<()> {
        self.edit_canonical(
            "store_changes_information",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::changes_information::store_changes_information(package, value)?;
                package.unsign();
                Ok(())
            },
        )
    }

    /// Discover the attached MS-OFFMACRO2 VBA project without inspecting its payload.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// type. It does not inspect, parse, decompress, or execute the binary
    /// VBA project bytes.
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let presentation = self.opc.main_document_part()?;
        discover_vba_project(&self.opc, presentation)
    }

    /// Attach a cache-free, inert MS-OVBA project and convert this package to PPTM/PPSM/POTM.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Attach a cache-free project with explicit resource limits.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        self.put_vba(project.finish(limits)?)
    }

    /// Attach a prevalidated `vbaProject.bin` payload without executing it.
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> Result<VbaProject> {
        let source = self.opc.main_document_part()?.partname().clone();
        store_presentation_vba_project(&mut self.opc, &source, payload)
    }

    /// Remove the VBA project graph and restore the corresponding non-macro main type.
    pub fn clear_vba(&mut self) -> Result<bool> {
        let source = self.opc.main_document_part()?.partname().clone();
        clear_presentation_vba(&mut self.opc, &source)
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

    /// Borrow the current plaintext OPC graph for focused low-level inspection.
    ///
    /// This rejects access while facade-owned changes still need managed
    /// materialization, or when exposing the graph would disclose an encrypted
    /// source as plaintext. Use [`Self::to_bytes`] for managed serialization.
    #[inline]
    pub fn opc(&self) -> Result<&OpcPackage> {
        self.ensure_opc_current("opc")?;
        Ok(&self.opc)
    }

    /// Return whether this presentation contains package signatures.
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
        self.edit_signature("sign", |package| package.sign(signer))
    }

    /// Add a signature with explicit authoring resource bounds.
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.edit_signature("sign_with", |package| package.sign_with(signer, limits))
    }

    /// Atomically replace all signatures with one signature.
    pub fn resign(&mut self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<PackURI> {
        self.edit_signature("resign", |package| package.resign(signer))
    }

    /// Atomically replace signatures with explicit authoring resource bounds.
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.edit_signature("resign_with", |package| package.resign_with(signer, limits))
    }

    /// Remove all package signatures.
    pub fn unsign(&mut self) -> &mut Self {
        self.opc.unsign();
        self
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan_with`] with [`Self::opc`] when a lower layer needs
    /// explicitly tuned limits and the raw graph is current.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.opc)?)
    }

    /// Load the bounded, inert notes-slide/notes-master graph.
    pub fn notes(&self) -> Result<Option<litchi_pptx::notes::Graph>> {
        self.ensure_notes_graph_current("notes")?;
        let presentation = self.opc.main_document_part()?.partname().clone();
        Ok(litchi_pptx::notes::load(&self.opc, &presentation)?)
    }

    /// Deterministically store an already coherent bounded notes graph.
    pub fn put_notes(&mut self, graph: litchi_pptx::notes::Graph) -> Result<()> {
        self.ensure_notes_graph_current("put_notes")?;
        let presentation = self.opc.main_document_part()?.partname().clone();
        self.edit_canonical("put_notes", STALE_NOTES_REASON, move |package| {
            Ok(litchi_pptx::notes::put(package, &presentation, graph)?)
        })
    }

    /// Remove the speaker notes owned by one selected slide.
    ///
    /// Exact producer-visible slide names are the primary selector. A checked
    /// zero-based index is also accepted for ordered workflows; native slide
    /// and relationship IDs remain below the safe facade. Missing notes are an
    /// idempotent `Ok(false)`.
    pub fn remove_notes<'a>(&mut self, slide: impl Into<SlideKey<'a>>) -> Result<bool> {
        self.ensure_notes_graph_current("remove_notes")?;
        let slide_name = self.resolve_slide(slide.into())?;
        let presentation = self.opc.main_document_part()?.partname().clone();
        self.edit_canonical("remove_notes", STALE_NOTES_REASON, move |package| {
            Ok(litchi_pptx::notes::remove(
                package,
                &presentation,
                &slide_name,
            )?)
        })
    }

    /// Remove speaker notes from every slide, returning the number removed.
    ///
    /// The complete Strict or Transitional notes graph is validated before
    /// any relationship or part is changed. Shared notes-master and theme
    /// resources remain available for presentation layout.
    pub fn clear_notes(&mut self) -> Result<usize> {
        self.ensure_notes_graph_current("clear_notes")?;
        let presentation = self.opc.main_document_part()?.partname().clone();
        self.edit_canonical("clear_notes", STALE_NOTES_REASON, move |package| {
            Ok(litchi_pptx::notes::clear(package, &presentation)?)
        })
    }

    /// Create a new slide master with default text styles and reference it
    /// from the presentation (`sldMasterIdLst` + relationship + part).
    ///
    /// The master is assigned a spec-compliant unique ID (≥ 2^31) and is
    /// related to an existing theme part when one is available.
    pub fn add_slide_master(&mut self) -> Result<crate::pptx::master_layout::AuthoredSlideMaster> {
        self.edit_canonical(
            "add_slide_master",
            STALE_PRESENTATION_GRAPH_REASON,
            crate::pptx::master_layout::add_slide_master,
        )
    }

    /// Create a new slide layout of the given kind, attached to an existing
    /// slide master identified by its part name (for example
    /// `/ppt/slideMasters/slideMaster1.xml`).
    ///
    /// The master's `sldLayoutIdLst`, relationships, and content types stay
    /// consistent, and the layout carries the required relationship back to
    /// its owning master. Optional placeholder shapes are inventoried back by
    /// the read side.
    pub fn add_slide_layout(
        &mut self,
        master_part_name: &str,
        kind: crate::pptx::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::pptx::master_layout::PlaceholderSpec],
    ) -> Result<crate::pptx::master_layout::AuthoredSlideLayout> {
        self.edit_canonical(
            "add_slide_layout",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::master_layout::add_slide_layout(
                    package,
                    master_part_name,
                    kind,
                    name,
                    placeholders,
                )
            },
        )
    }

    /// Add or replace a placeholder shape (and its prompt text) on a slide
    /// master or slide layout part.
    ///
    /// The placeholder is matched by its `p:ph` type and index; an existing
    /// match is replaced in place, otherwise a new shape is appended.
    pub fn store_placeholder_shape(
        &mut self,
        part_name: &str,
        spec: &crate::pptx::master_layout::PlaceholderSpec,
    ) -> Result<()> {
        self.edit_canonical(
            "store_placeholder_shape",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| crate::pptx::master_layout::store_placeholder_shape(package, part_name, spec),
        )
    }

    /// Delete a slide layout that is not referenced by any slide.
    ///
    /// The owning master's `sldLayoutIdLst` entry and relationship are
    /// removed together with the layout part.
    pub fn remove_slide_layout(&mut self, layout_part_name: &str) -> Result<()> {
        self.edit_canonical(
            "remove_slide_layout",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| crate::pptx::master_layout::remove_slide_layout(package, layout_part_name),
        )
    }

    /// Validate the slide master and slide layout relationship graph with the
    /// same rules the read side applies.
    pub fn validate_master_layout_graph(&self) -> Result<()> {
        crate::pptx::master_layout::validate_master_layout_graph(&self.opc)
    }

    /// Create a new theme part with a caller-supplied color scheme and font
    /// scheme, registered with the Office theme content type.
    ///
    /// The theme is written to the next free `/ppt/theme/themeN.xml` part
    /// name with the twelve-slot color scheme, the major/minor font scheme,
    /// and the default format scheme; serialization is deterministic.
    pub fn add_theme(
        &mut self,
        name: &str,
        color_scheme: &crate::pptx::theme::ThemeColorScheme,
        font_scheme: &crate::pptx::theme::ThemeFontScheme,
    ) -> Result<crate::pptx::theme::AuthoredTheme> {
        self.edit_canonical("add_theme", STALE_PRESENTATION_GRAPH_REASON, |package| {
            crate::pptx::theme::add_theme(package, name, color_scheme, font_scheme)
        })
    }

    /// Attach a theme part to a slide master through a theme relationship.
    ///
    /// The master part must exist with the slide-master content type and no
    /// existing theme relationship; the theme part must exist with the
    /// Office theme content type. The master/layout/theme graph is
    /// re-validated afterwards.
    pub fn attach_theme_to_master(
        &mut self,
        master_part_name: &str,
        theme_part_name: &str,
    ) -> Result<String> {
        self.edit_canonical(
            "attach_theme_to_master",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::theme::attach_theme_to_master(
                    package,
                    master_part_name,
                    theme_part_name,
                )
            },
        )
    }

    /// Store a theme override on a slide layout or slide part.
    ///
    /// The override is validated and serialized deterministically; an
    /// existing override relationship on the parent is reused, otherwise a
    /// new `/ppt/theme/themeOverrideN.xml` part and relationship are
    /// created. Returns the override part name.
    pub fn store_theme_override(
        &mut self,
        parent_part_name: &str,
        value: &crate::pptx::ThemeOverride,
    ) -> Result<String> {
        self.edit_canonical(
            "store_theme_override",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| crate::pptx::store_theme_override(package, parent_part_name, value),
        )
    }

    /// Read the theme override attached to a slide layout or slide part.
    pub fn theme_override(
        &self,
        parent_part_name: &str,
    ) -> Result<Option<crate::pptx::ThemeOverride>> {
        crate::pptx::theme_override(&self.opc, parent_part_name)
    }

    /// Remove the theme override from a slide layout or slide part,
    /// deleting the override part when it becomes orphaned.
    pub fn remove_theme_override(&mut self, parent_part_name: &str) -> Result<bool> {
        self.edit_canonical(
            "remove_theme_override",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| crate::pptx::remove_theme_override(package, parent_part_name),
        )
    }

    /// Replace the color scheme (`a:clrScheme`) of an existing theme part,
    /// leaving the rest of the part untouched.
    pub fn store_theme_color_scheme(
        &mut self,
        theme_part_name: &str,
        color_scheme: &crate::pptx::theme::ThemeColorScheme,
    ) -> Result<()> {
        self.edit_canonical(
            "store_theme_color_scheme",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::theme::store_theme_color_scheme(package, theme_part_name, color_scheme)
            },
        )
    }

    /// Replace the font scheme (`a:fontScheme`) of an existing theme part,
    /// leaving the rest of the part untouched.
    pub fn store_theme_font_scheme(
        &mut self,
        theme_part_name: &str,
        font_scheme: &crate::pptx::theme::ThemeFontScheme,
    ) -> Result<()> {
        self.edit_canonical(
            "store_theme_font_scheme",
            STALE_PRESENTATION_GRAPH_REASON,
            |package| {
                crate::pptx::theme::store_theme_font_scheme(package, theme_part_name, font_scheme)
            },
        )
    }

    /// Validate the master/layout/theme relationship graph with the same
    /// rules the read side applies.
    pub fn validate_theme_graph(&self) -> Result<()> {
        crate::pptx::theme::validate_theme_graph(&self.opc)
    }

    /// Embed an inert binary payload into a slide as an OLE object shape.
    ///
    /// The payload is stored verbatim in the next free
    /// `/ppt/embeddings/oleObjectN.bin` part, the slide gains the matching
    /// OLE/package relationship, and a `p:graphicFrame` carrying `p:oleObj`
    /// is appended to the slide's shape tree. The patched slide is verified
    /// through the read-side OLE inventory before the operation returns.
    /// Payloads are never parsed, activated, rendered, or executed.
    /// A successful edit disables subsequent legacy-writer access on this value.
    pub fn add_ole_object(
        &mut self,
        slide_part_name: &str,
        kind: crate::pptx::ole::PptxOlePayloadKind,
        prog_id: Option<&str>,
        name: Option<&str>,
        frame: crate::pptx::ole_object::OleObjectFrame,
        payload: &[u8],
    ) -> Result<crate::pptx::ole_object::AuthoredOleObject> {
        self.edit_canonical("add_ole_object", STALE_SLIDE_GRAPH_REASON, |package| {
            crate::pptx::ole_object::add_ole_object(
                package,
                slide_part_name,
                kind,
                prog_id,
                name,
                frame,
                payload,
            )
        })
    }

    /// Transactionally edit the current plaintext OPC graph.
    ///
    /// The closure receives a structural candidate whose built-in part payloads
    /// share immutable `Arc` storage. Returning an error or unwinding leaves
    /// this package's graph unpublished; custom `Part` implementations retain
    /// their own clone and interior-mutability policy. Before a successful
    /// commit, the candidate's PowerPoint main relationship, content type, and
    /// core properties are validated and facade-owned state is reloaded.
    /// Committing a raw edit disables the legacy writer so it cannot later
    /// erase the edit.
    pub fn edit_opc<T>(&mut self, edit: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.ensure_opc_current("edit_opc")?;
        if self.opc.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "edit_opc",
                reason: "raw OPC editing cannot honor an automatic font policy; disable font embedding before entering the transaction",
            });
        }

        let mut candidate = self.opc.clone();
        candidate.unsign();
        let value = edit(&mut candidate)?;

        if candidate.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "edit_opc",
                reason: "raw OPC transactions cannot configure automatic font embedding; use the typed package facade",
            });
        }
        let main_part = candidate.main_document_part().map_err(|error| {
            OoxmlError::PartNotFound(format!("main presentation part: {error}"))
        })?;
        validate_presentation_main_content_type(main_part.content_type())?;
        let properties = Slot::load(&candidate)?;

        self.opc = candidate;
        self.properties = properties;
        self.mutable_pres = None;
        #[cfg(feature = "fonts")]
        {
            self.font_sync_pending = false;
        }
        Ok(value)
    }

    /// Get a mutable presentation for writing and modification.
    ///
    /// This returns a `MutablePresentation` that allows you to add and modify
    /// slides, shapes, and other presentation elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut pres = pkg.presentation_mut()?;
    ///
    /// // Add a slide
    /// let slide = pres.add_slide()?;
    /// slide.set_title("My Presentation");
    /// slide.add_text_box("Hello, World!", 914400, 914400, 2743200, 914400);
    ///
    /// pkg.save("output.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn presentation_mut(&mut self) -> Result<&mut MutablePresentation> {
        if self.mutable_pres.is_none() {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "presentation_mut",
                reason: "the legacy writer cannot hydrate an existing presentation losslessly",
            });
        }

        #[cfg(feature = "fonts")]
        if self.opc.save_options().fonts != litchi_opc::FontEmbedding::None {
            self.font_sync_pending = true;
        }

        match self.mutable_pres.as_mut() {
            Some(presentation) => Ok(presentation),
            None => Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "presentation_mut",
                reason: "the legacy writer cannot hydrate an existing presentation losslessly",
            }),
        }
    }

    /// Borrows the presentation core properties, retaining package absence.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
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
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// if let Some(props) = pkg.props_mut() {
    ///     props.title = Some("My Presentation".to_string());
    /// }
    /// pkg.save("presentation.pptx")?;
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

    /// Save the package to a file.
    ///
    /// Writes the complete PowerPoint presentation including all parts, relationships,
    /// and content types to a .pptx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .pptx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify presentation...
    /// pkg.save("output.pptx")?;
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
        self.write_with(|package| package.save(path).map_err(Into::into))
    }

    /// Serialize the complete plaintext package through the managed staged
    /// publication path.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.ensure_plain_output("to_bytes")?;
        self.to_plain_bytes()
    }

    /// Explicitly serialize a plaintext package, even when the source was
    /// encrypted.
    pub fn to_plain_bytes(&mut self) -> Result<Vec<u8>> {
        use litchi_opc::pkgwriter::PackageWriter;

        self.write_with(|package| PackageWriter::to_bytes(package).map_err(Into::into))
    }

    /// Serialize and encrypt this package entirely in memory.
    #[cfg(feature = "encryption")]
    pub fn to_encrypted(&mut self, password: &str, mode: Mode) -> Result<Vec<u8>> {
        use litchi_opc::pkgwriter::PackageWriter;

        self.write_with(|package| {
            let package = PackageWriter::to_bytes(package)?;
            crate::encryption::encrypt(package, password, mode).map_err(Into::into)
        })
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
        use litchi_opc::pkgwriter::PackageWriter;
        use std::io::Write;

        self.write_with(|package| {
            litchi_opc::atomic::replace_with::<OoxmlError>(path.as_ref(), |temporary| {
                let plain = PackageWriter::to_bytes(package)?;
                let encrypted = crate::encryption::encrypt(plain, password, mode)?;
                temporary.write_all(&encrypted)?;
                Ok(())
            })
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

    fn ensure_plain_output(&self, _operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: _operation,
                reason: "the source package was encrypted; use save_reencrypted or save_plain",
            });
        }
        Ok(())
    }

    fn ensure_opc_current(&self, operation: &'static str) -> Result<()> {
        #[cfg(feature = "encryption")]
        if self.source_encryption.is_some() {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "raw OPC access would expose an encrypted source as plaintext; use the managed encryption or explicit plaintext APIs",
            });
        }

        if self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
        {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "the legacy writer has unmaterialized presentation changes; use a managed save or to_bytes first",
            });
        }
        if self.properties.is_dirty() {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "core properties have unmaterialized changes; use a managed save or to_bytes first",
            });
        }

        #[cfg(feature = "fonts")]
        if self.font_sync_pending {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation,
                reason: "font embedding is pending; use a managed save or to_bytes first",
            });
        }

        Ok(())
    }

    #[cfg(feature = "encryption")]
    fn preserved_mode(&self, operation: &'static str) -> Result<Mode> {
        self.source_encryption.ok_or(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation,
            reason: "the source package has no encryption profile; supply an explicit Mode",
        })
    }

    fn write_with<T>(&mut self, write: impl FnOnce(&OpcPackage) -> Result<T>) -> Result<T> {
        let presentation_dirty = self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified);
        let properties_dirty = self.properties.is_dirty();
        #[cfg(feature = "fonts")]
        let fonts_requested = self.font_sync_pending;
        #[cfg(not(feature = "fonts"))]
        let fonts_requested = false;

        // An unchanged package needs no rollback graph. This is the common
        // opened-document save path and retains shared payload identity.
        if !presentation_dirty && !properties_dirty && !fonts_requested {
            return write(&self.opc);
        }

        // Built-in OPC parts retain immutable payloads behind `Arc`, so this
        // transaction copies bounded graph metadata rather than package bytes.
        let package_before = self.opc.clone();
        let result = (|| {
            let presentation_staged = self.materialize_presentation()?;

            // Font publication may inspect the complete writer model and is
            // therefore staged after presentation materialization.
            #[cfg(feature = "fonts")]
            self.embed_fonts()?;

            // This guard keeps core-property intent dirty until the output
            // sink has accepted the complete staged package.
            let staged_properties = self.properties.stage(&mut self.opc)?;
            let value = write(&self.opc)?;
            staged_properties.commit();
            if presentation_staged && let Some(presentation) = self.mutable_pres.as_mut() {
                presentation.mark_clean();
            }
            #[cfg(feature = "fonts")]
            {
                self.font_sync_pending = false;
            }
            Ok(value)
        })();

        match result {
            Ok(value) => Ok(value),
            Err(error) => {
                self.opc = package_before;
                Err(error)
            },
        }
    }

    /// Materialize pending legacy-writer state without allowing it to erase
    /// font relationships or leave orphan font parts behind.
    ///
    /// OPC part payloads are `Arc` backed, so the rollback snapshot shares
    /// immutable bytes and copies only bounded package metadata.
    fn flush_presentation(&mut self) -> Result<()> {
        let should_update = self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified);
        if !should_update {
            return Ok(());
        }

        let package_before = self.opc.clone();
        match self.materialize_presentation() {
            Ok(true) => {
                if let Some(presentation) = self.mutable_pres.as_mut() {
                    presentation.mark_clean();
                }
                Ok(())
            },
            Ok(false) => Ok(()),
            Err(error) => {
                self.opc = package_before;
                Err(error)
            },
        }
    }

    /// Stage pending legacy-writer state into the current OPC graph.
    ///
    /// The caller owns rollback and decides when successful publication may
    /// mark the writer model clean.
    fn materialize_presentation(&mut self) -> Result<bool> {
        let should_update = self
            .mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified);
        if !should_update {
            return Ok(false);
        }

        let Some(presentation) = self.mutable_pres.take() else {
            return Ok(false);
        };
        let result = (|| {
            // The legacy writer replaces presentation.xml and its
            // relationships. Detach the typed graph first, then republish it
            // after materialization so source payload allocations are moved
            // through the operation without copying their bytes.
            let fonts = litchi_pptx::font::remove(&mut self.opc)?;
            self.update_presentation_parts(&presentation)?;
            if let Some(fonts) = fonts {
                litchi_pptx::font::put(&mut self.opc, fonts)?;
            }
            Ok(())
        })();

        self.mutable_pres = Some(presentation);
        result.map(|()| true)
    }

    /// Update presentation parts with modified data.
    fn update_presentation_parts(&mut self, pres: &MutablePresentation) -> Result<()> {
        use crate::pptx::parts::CommentAuthor;
        use crate::pptx::parts::{generate_comment_authors_xml, generate_comments_xml};
        use crate::pptx::template;
        use crate::pptx::writer::relmap::RelationshipMapper;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::{BlobPart, Part};

        if litchi_pptx::table::style::conformance(&self.opc)?
            == litchi_pptx::table::style::Conformance::Strict
        {
            return Err(OoxmlError::UnsafeEdit {
                format: "PPTX",
                operation: "materialize_presentation",
                reason: "the legacy presentation writer emits Transitional markup and cannot safely rewrite a Strict package",
            });
        }
        let table_styles = litchi_pptx::table::style::link(&self.opc)?.map(|link| {
            (
                link.kind().to_owned(),
                link.target().to_owned(),
                link.id().to_owned(),
            )
        });

        // Initialize relationship mapper
        let mut rel_mapper = RelationshipMapper::new();

        // Collect all images from all slides (shapes)
        let all_images = pres.collect_all_images();

        // Collect all background images
        let all_bg_images = pres.collect_all_background_images();

        // Collect all media (audio/video) from all slides
        let all_media = pres.collect_all_media();

        // Collect all comments from all slides
        let all_comments = pres.collect_all_comments();

        // Track the total number of images for unique numbering
        let mut total_image_count = 0;

        // Track the total number of media files for unique numbering
        let mut total_media_count = 0;

        // Create image parts for shape images first and add to package
        for (_slide_index, image_data, image_format) in &all_images {
            total_image_count += 1;
            let ext = image_format.extension();

            // Create image part URI
            let image_partname = format!("/ppt/media/image{}.{}", total_image_count, ext);
            let image_uri = PackURI::new(&image_partname)
                .map_err(|e| OoxmlError::InvalidUri(format!("image URI: {}", e)))?;

            // Create image part
            let image_part = BlobPart::new(
                image_uri,
                image_format.mime_type().to_string(),
                image_data.to_vec(),
            );

            // Add image part to package
            self.opc.add_part(Box::new(image_part));
        }

        // Create image parts for background images
        for (_slide_index, image_data, image_format) in &all_bg_images {
            total_image_count += 1;
            let ext = image_format.extension();

            // Create image part URI
            let image_partname = format!("/ppt/media/image{}.{}", total_image_count, ext);
            let image_uri = PackURI::new(&image_partname)
                .map_err(|e| OoxmlError::InvalidUri(format!("image URI: {}", e)))?;

            // Create image part
            let image_part = BlobPart::new(
                image_uri,
                image_format.mime_type().to_string(),
                image_data.to_vec(),
            );

            // Add image part to package
            self.opc.add_part(Box::new(image_part));
        }

        // Create media parts (audio/video) and poster images, add to package
        for (_slide_index, _media_index, media_data, media_format) in &all_media {
            total_media_count += 1;
            let ext = media_format.extension();

            // Create media part URI
            let media_partname = format!("/ppt/media/media{}.{}", total_media_count, ext);
            let media_uri = PackURI::new(&media_partname)
                .map_err(|e| OoxmlError::InvalidUri(format!("media URI: {}", e)))?;

            // Create media part
            let media_part = BlobPart::new(
                media_uri,
                media_format.mime_type().to_string(),
                media_data.to_vec(),
            );

            // Add media part to package
            self.opc.add_part(Box::new(media_part));

            // Create poster image part for this media
            // Each media needs a poster image for blipFill/blip
            let poster_partname = format!("/ppt/media/poster{}.png", total_media_count);
            let poster_uri = PackURI::new(&poster_partname)
                .map_err(|e| OoxmlError::InvalidUri(format!("poster URI: {}", e)))?;

            let poster_part = BlobPart::new(
                poster_uri,
                "image/png".to_string(),
                DEFAULT_MEDIA_POSTER.to_vec(),
            );

            self.opc.add_part(Box::new(poster_part));
        }

        // Create comment authors part if there are any comments
        if !all_comments.is_empty() {
            // Create a default author for now (could be extended to support multiple authors)
            let authors = vec![CommentAuthor::new(0, "Author", "A")];
            let authors_xml = generate_comment_authors_xml(&authors);

            let authors_uri = PackURI::new("/ppt/commentAuthors.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("commentAuthors URI: {}", e)))?;

            let authors_part = BlobPart::new(
                authors_uri,
                ct::PML_COMMENT_AUTHORS.to_string(),
                authors_xml.into_bytes(),
            );

            self.opc.add_part(Box::new(authors_part));
        }

        // Create chart parts and add to package
        for (chart_idx, chart_parts) in &pres.charts {
            // Create chart XML part
            let chart_uri = PackURI::new(format!("/ppt/charts/chart{}.xml", chart_idx))
                .map_err(|e| OoxmlError::InvalidUri(format!("chart{} URI: {}", chart_idx, e)))?;

            let mut chart_part = BlobPart::new(
                chart_uri,
                ct::DML_CHART.to_string(),
                chart_parts.chart_xml.as_bytes().to_vec(),
            );

            // Add relationship from chart to embedded Excel data
            chart_part.relate_to(
                &format!("../embeddings/Microsoft_Excel_Worksheet{}.xlsx", chart_idx),
                rt::PACKAGE,
            );

            self.opc.add_part(Box::new(chart_part));

            // Create embedded Excel workbook part
            let excel_uri = PackURI::new(format!(
                "/ppt/embeddings/Microsoft_Excel_Worksheet{}.xlsx",
                chart_idx
            ))
            .map_err(|e| OoxmlError::InvalidUri(format!("excel{} URI: {}", chart_idx, e)))?;

            let excel_part = BlobPart::new(
                excel_uri,
                ct::SML_SHEET.to_string(),
                chart_parts.excel_data.clone(),
            );

            self.opc.add_part(Box::new(excel_part));
        }

        // Collect SmartArt shape positions from slides for drawing generation
        let mut smartart_positions: std::collections::HashMap<u32, (i64, i64, i64, i64)> =
            std::collections::HashMap::new();
        for slide in &pres.slides {
            for shape in &slide.shapes {
                if let crate::pptx::writer::shape::ShapeType::SmartArt {
                    x,
                    y,
                    width,
                    height,
                    diagram_idx,
                    ..
                } = &shape.shape_type
                {
                    smartart_positions.insert(*diagram_idx, (*x, *y, *width, *height));
                }
            }
        }

        // Create SmartArt diagram parts and add to package
        for (diagram_idx, smartart_parts) in &pres.smartarts {
            // Get position/size for drawing generation (use defaults if not found)
            let (x, y, width, height) = smartart_positions
                .get(diagram_idx)
                .copied()
                .unwrap_or((0, 0, 5486400, 3657600)); // Default 6" x 4"

            // Generate drawing XML with actual position/size
            let drawing_xml = crate::pptx::smartart::generate_smartart_drawing_xml(
                &smartart_parts.smartart,
                x,
                y,
                width,
                height,
            );

            // Create diagram data XML part with relationship to drawing
            let data_uri =
                PackURI::new(format!("/ppt/diagrams/data{}.xml", diagram_idx)).map_err(|e| {
                    OoxmlError::InvalidUri(format!("diagram data{} URI: {}", diagram_idx, e))
                })?;

            // Start with the generated data XML and attach the diagramDrawing relationship so we
            // can embed a dataModelExt extLst referencing it (matches Apache POI / PowerPoint).
            let mut data_xml = smartart_parts.data_xml.clone();

            let mut data_part = BlobPart::new(
                data_uri,
                ct::DML_DIAGRAM_DATA.to_string(),
                data_xml.clone().into_bytes(),
            );

            // Add relationship from data to drawing and capture its Id
            let drawing_rel_id = data_part.relate_to(
                &format!("drawing{}.xml", diagram_idx),
                "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
            );

            // Inject extLst with dsp:dataModelExt referencing the drawing relationship, if possible.
            // This mirrors the structure produced by PowerPoint and Apache POI.
            if let Some(pos) = data_xml.rfind("</dgm:dataModel>") {
                let ext = format!(
                    concat!(
                        "<dgm:extLst>",
                        "<a:ext xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
                        "uri=\"http://schemas.microsoft.com/office/drawing/2008/diagram\">",
                        "<dsp:dataModelExt xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" ",
                        "relId=\"{}\" ",
                        "minVer=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>",
                        "</a:ext>",
                        "</dgm:extLst>",
                    ),
                    drawing_rel_id,
                );
                data_xml.insert_str(pos, &ext);
                data_part.set_blob(data_xml.into_bytes());
            }

            self.opc.add_part(Box::new(data_part));

            // Create diagram drawing XML part
            let drawing_uri = PackURI::new(format!("/ppt/diagrams/drawing{}.xml", diagram_idx))
                .map_err(|e| {
                    OoxmlError::InvalidUri(format!("diagram drawing{} URI: {}", diagram_idx, e))
                })?;

            let drawing_part = BlobPart::new(
                drawing_uri,
                ct::DML_DIAGRAM_DRAWING.to_string(),
                drawing_xml.as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(drawing_part));

            // Create diagram layout XML part
            let layout_uri = PackURI::new(format!("/ppt/diagrams/layout{}.xml", diagram_idx))
                .map_err(|e| {
                    OoxmlError::InvalidUri(format!("diagram layout{} URI: {}", diagram_idx, e))
                })?;

            let layout_part = BlobPart::new(
                layout_uri,
                ct::DML_DIAGRAM_LAYOUT.to_string(),
                smartart_parts.layout_xml.as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(layout_part));

            // Create diagram quick style XML part
            let style_uri = PackURI::new(format!("/ppt/diagrams/quickStyle{}.xml", diagram_idx))
                .map_err(|e| {
                    OoxmlError::InvalidUri(format!("diagram style{} URI: {}", diagram_idx, e))
                })?;

            let style_part = BlobPart::new(
                style_uri,
                ct::DML_DIAGRAM_STYLE.to_string(),
                smartart_parts.style_xml.as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(style_part));

            // Create diagram colors XML part
            let colors_uri = PackURI::new(format!("/ppt/diagrams/colors{}.xml", diagram_idx))
                .map_err(|e| {
                    OoxmlError::InvalidUri(format!("diagram colors{} URI: {}", diagram_idx, e))
                })?;

            let colors_part = BlobPart::new(
                colors_uri,
                ct::DML_DIAGRAM_COLORS.to_string(),
                smartart_parts.colors_xml.as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(colors_part));
        }

        // Create presentation part and add relationships
        let pres_uri = PackURI::new("/ppt/presentation.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("presentation URI: {}", e)))?;

        let (presentation_content_type, vba_project_target) = {
            let existing = self.opc.get_part(&pres_uri)?;
            discover_vba_project(&self.opc, existing)?;
            let mut projects = existing
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == rt::VBA_PROJECT);
            let target = match projects.next() {
                Some(relationship) if relationship.is_external() => {
                    return Err(OoxmlError::InvalidFormat(
                        "presentation VBA Project relationship cannot be external".to_string(),
                    ));
                },
                Some(relationship) => Some(relationship.target_ref().to_string()),
                None => None,
            };
            if projects.next().is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "presentation has multiple VBA Project relationships".to_string(),
                ));
            }
            (existing.content_type().to_string(), target)
        };

        // Create a temporary presentation part to manage relationships
        let mut temp_pres_part =
            BlobPart::new(pres_uri.clone(), presentation_content_type, Vec::new());

        // Preserve the validated optional table-style edge exactly. Adding it
        // first reserves a producer-selected relationship ID; generated edges
        // then select free IDs around it.
        if let Some((kind, target, id)) = table_styles {
            temp_pres_part.rels_mut().try_add_relationship(
                kind,
                target,
                id,
                litchi_opc::TargetMode::Internal,
            )?;
        }

        // Add relationship to slideMaster and retain its allocated ID because
        // a producer is allowed to use rId1 for another relationship.
        let slide_master_rel_id =
            temp_pres_part.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

        // Add other required relationships (in the order they were created in Package::new())
        // These relationships should be added even if not modified, as they're required for a valid PPTX
        temp_pres_part.relate_to("viewProps.xml", rt::VIEW_PROPS);
        temp_pres_part.relate_to("presProps.xml", rt::PRES_PROPS);
        temp_pres_part.relate_to("theme/theme1.xml", rt::THEME);
        if let Some(target) = vba_project_target {
            temp_pres_part.relate_to(&target, rt::VBA_PROJECT);
        }

        // Add relationship to notesMaster (required when we have notesSlides)
        let notes_master_rel_id =
            temp_pres_part.relate_to("notesMasters/notesMaster1.xml", rt::NOTES_MASTER);

        // Add relationship to commentAuthors if there are comments
        if !all_comments.is_empty() {
            temp_pres_part.relate_to("commentAuthors.xml", rt::COMMENT_AUTHORS);
        }

        // Track slide relationship IDs for presentation.xml generation
        let mut slide_rel_ids: Vec<String> = Vec::new();

        // Process each slide: create relationships first, then generate XML
        // Note: We process ALL slides, not just modified ones, because when creating a new
        // presentation or when slides have been reordered, we need to regenerate everything
        for (slide_index, slide) in pres.slides.iter().enumerate() {
            let slide_num = slide_index + 1;
            let slide_uri = PackURI::new(format!("/ppt/slides/slide{}.xml", slide_num))
                .map_err(|e| OoxmlError::InvalidUri(format!("slide{} URI: {}", slide_num, e)))?;

            // Create a temporary slide part to manage relationships
            let mut temp_slide_part =
                BlobPart::new(slide_uri.clone(), ct::PML_SLIDE.to_string(), Vec::new());

            // Add relationship from slide to slide layout (always first relationship)
            temp_slide_part.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);

            // Collect images for this slide and create relationships
            let slide_images = slide.collect_images();
            for (img_index_in_slide, (_, image_format)) in slide_images.iter().enumerate() {
                // Find the global image index for this slide's image
                let mut global_img_idx = 0;
                for (global_idx, (s_idx, _, _)) in all_images.iter().enumerate() {
                    if *s_idx == slide_index {
                        if global_img_idx == img_index_in_slide {
                            let img_num = global_idx + 1;
                            let ext = image_format.extension();
                            let image_rel_target = format!("../media/image{}.{}", img_num, ext);
                            let rid = temp_slide_part.relate_to(&image_rel_target, rt::IMAGE);
                            rel_mapper.add_image(slide_index, img_index_in_slide, rid);
                            break;
                        }
                        global_img_idx += 1;
                    }
                }
            }

            // Add relationship for background image if present
            if slide.get_background_image().is_some() {
                // Find the background image for this slide in all_bg_images
                for (bg_idx, (bg_slide_idx, _, bg_format)) in all_bg_images.iter().enumerate() {
                    if *bg_slide_idx == slide_index {
                        // Calculate the image number (after all shape images)
                        let bg_img_num = all_images.len() + bg_idx + 1;
                        let ext = bg_format.extension();
                        let bg_rel_target = format!("../media/image{}.{}", bg_img_num, ext);
                        let rid = temp_slide_part.relate_to(&bg_rel_target, rt::IMAGE);
                        rel_mapper.add_background(slide_index, rid);
                        break;
                    }
                }
            }

            // Add relationship from slide to notes slide if notes exist
            if slide.has_notes() {
                let notes_rel_target = format!("../notesSlides/notesSlide{}.xml", slide_num);
                let rid = temp_slide_part.relate_to(&notes_rel_target, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide");
                rel_mapper.add_notes(slide_index, rid);
            }

            // Add relationships for media (audio/video) on this slide
            // PowerPoint requires THREE relationships per media file:
            // 1. OOXML video/audio type (for r:link in a:videoFile/a:audioFile)
            // 2. Microsoft media type (for r:embed in p14:media extension)
            // 3. Poster image type (for r:embed in blipFill/blip)
            let slide_media = slide.collect_media();
            for (media_index_in_slide, (_, media_format)) in slide_media.iter().enumerate() {
                // Find the global media index for this slide's media
                for (global_idx, (s_idx, m_idx, _, _)) in all_media.iter().enumerate() {
                    if *s_idx == slide_index && *m_idx == media_index_in_slide {
                        let media_num = global_idx + 1;
                        let ext = media_format.extension();
                        let media_rel_target = format!("../media/media{}.{}", media_num, ext);

                        // Add OOXML video/audio relationship (for r:link in a:videoFile/a:audioFile)
                        let video_rel_type = match media_format.media_type() {
                            crate::pptx::media::MediaType::Audio => rt::AUDIO,
                            crate::pptx::media::MediaType::Video => rt::VIDEO,
                        };
                        let video_rid =
                            temp_slide_part.relate_to(&media_rel_target, video_rel_type);

                        // Add Microsoft media relationship (for r:embed in p14:media)
                        let media_rid = temp_slide_part.relate_to(&media_rel_target, rt::MEDIA);

                        // Add poster image for this media (required for blipFill/blip)
                        // Use a default placeholder image - shared across all media on this slide
                        let poster_image_path = format!("../media/poster{}.png", media_num);
                        let poster_rid = temp_slide_part.relate_to(&poster_image_path, rt::IMAGE);

                        rel_mapper.add_media(
                            slide_index,
                            media_index_in_slide,
                            video_rid,
                            media_rid,
                            poster_rid,
                        );
                        break;
                    }
                }
            }

            // Add relationship for comments if this slide has comments
            if !slide.comments().is_empty() {
                let comments_rel_target = format!("../comments/comment{}.xml", slide_num);
                let rid = temp_slide_part.relate_to(&comments_rel_target, rt::COMMENTS);
                rel_mapper.add_comments(slide_index, rid);
            }

            // Add relationships for charts on this slide
            // We need to scan the slide's shapes for Chart types and create relationships
            for shape in &slide.shapes {
                if let crate::pptx::writer::shape::ShapeType::Chart { chart_idx, .. } =
                    &shape.shape_type
                {
                    let chart_rel_target = format!("../charts/chart{}.xml", chart_idx);
                    let rid = temp_slide_part.relate_to(&chart_rel_target, rt::CHART);
                    rel_mapper.add_chart(slide_index, *chart_idx, rid);
                }
            }

            // Add relationships for SmartArt diagrams on this slide
            for shape in &slide.shapes {
                if let crate::pptx::writer::shape::ShapeType::SmartArt { diagram_idx, .. } =
                    &shape.shape_type
                {
                    // SmartArt requires 4 standard relationships plus an optional diagramDrawing
                    // extension relationship used by PowerPoint/Apache POI for pre-rendered shapes.
                    let data_rel_target = format!("../diagrams/data{}.xml", diagram_idx);
                    let layout_rel_target = format!("../diagrams/layout{}.xml", diagram_idx);
                    let style_rel_target = format!("../diagrams/quickStyle{}.xml", diagram_idx);
                    let colors_rel_target = format!("../diagrams/colors{}.xml", diagram_idx);
                    let drawing_rel_target = format!("../diagrams/drawing{}.xml", diagram_idx);

                    let data_rid = temp_slide_part.relate_to(
                        &data_rel_target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                    );
                    let layout_rid = temp_slide_part.relate_to(
                        &layout_rel_target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
                    );
                    let style_rid = temp_slide_part.relate_to(
                        &style_rel_target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
                    );
                    let colors_rid = temp_slide_part.relate_to(
                        &colors_rel_target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
                    );

                    // Slide-level relationship to the Microsoft-specific diagramDrawing part.
                    // This matches the structure produced by PowerPoint and Apache POI and is
                    // used by tools to locate the pre-rendered SmartArt shapes.
                    let _drawing_rid = temp_slide_part.relate_to(
                        &drawing_rel_target,
                        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                    );

                    rel_mapper.add_smartart(
                        slide_index,
                        *diagram_idx,
                        data_rid,
                        layout_rid,
                        style_rid,
                        colors_rid,
                    );
                }
            }

            // Now generate slide XML with actual relationship IDs
            let slide_xml = slide.to_xml_with_rels(Some(slide_index), Some(&rel_mapper))?;

            // Update the temp part with the actual XML content
            temp_slide_part.set_blob(slide_xml.into_bytes());

            // Add the slide part to the package
            self.opc.add_part(Box::new(temp_slide_part));

            // Create notes slide if notes exist
            if let Some(notes) = slide.notes() {
                let notes_xml = litchi_pptx::notes::write_text(notes)?;
                let notes_uri =
                    PackURI::new(format!("/ppt/notesSlides/notesSlide{}.xml", slide_num)).map_err(
                        |e| OoxmlError::InvalidUri(format!("notesSlide{} URI: {}", slide_num, e)),
                    )?;

                let mut notes_part = BlobPart::new(
                    notes_uri,
                    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
                        .to_string(),
                    notes_xml,
                );

                // Add relationship from notes to slide
                notes_part.relate_to(&format!("../slides/slide{}.xml", slide_num), rt::SLIDE);

                // Add relationship from notes to notesMaster (REQUIRED by PowerPoint!)
                notes_part.relate_to("../notesMasters/notesMaster1.xml", rt::NOTES_MASTER);

                self.opc.add_part(Box::new(notes_part));
            }

            // Create comments part if this slide has comments
            if !slide.comments().is_empty() {
                let comments_xml = generate_comments_xml(slide.comments());
                let comments_uri = PackURI::new(format!("/ppt/comments/comment{}.xml", slide_num))
                    .map_err(|e| {
                        OoxmlError::InvalidUri(format!("comment{} URI: {}", slide_num, e))
                    })?;

                let comments_part = BlobPart::new(
                    comments_uri,
                    ct::PML_COMMENTS.to_string(),
                    comments_xml.into_bytes(),
                );

                self.opc.add_part(Box::new(comments_part));
            }

            // Add relationship from presentation to this slide and track the ID
            let rel_target = format!("slides/slide{}.xml", slide_num);
            let slide_rid = temp_pres_part.relate_to(&rel_target, rt::SLIDE);
            slide_rel_ids.push(slide_rid);
        }

        // Create custom handout master if one is set
        // We need to get the relationship ID BEFORE generating the presentation XML
        let handout_rel_id = if let Some(handout_master) = pres.handout_master() {
            // Notes and authored masters may already own additional themes;
            // allocate from the actual graph instead of assuming theme3.xml.
            let handout_theme_uri = crate::pptx::theme::next_theme_part_uri(&self.opc)?;
            let handout_theme_target = format!("../theme/{}", handout_theme_uri.filename());
            let handout_theme_part = BlobPart::new(
                handout_theme_uri,
                ct::OFC_THEME.to_string(),
                template::default_theme_xml().as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(handout_theme_part));

            let handout_uri = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("handoutMaster URI: {}", e)))?;

            let mut handout_part = BlobPart::new(
                handout_uri,
                "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
                    .to_string(),
                handout_master.to_xml().into_bytes(),
            );

            // Add relationship from handoutMaster to its own theme.
            handout_part.relate_to(&handout_theme_target, rt::THEME);

            // Add relationship from presentation to handoutMaster and capture the ID
            let rel_id =
                temp_pres_part.relate_to("handoutMasters/handoutMaster1.xml", rt::HANDOUT_MASTER);

            self.opc.add_part(Box::new(handout_part));

            // Note: presProps.xml with prnPr for handout layout is already added in Package::new()
            // We don't add it again here to avoid duplicate parts which causes corruption

            Some(rel_id)
        } else {
            None
        };

        // Now generate presentation XML with actual relationship IDs
        // Note: notesMasterIdLst is NOT required for handout master (per python-pptx reference)
        let pres_xml = pres.generate_presentation_xml_with_rels(
            Some(&slide_master_rel_id),
            Some(&slide_rel_ids),
            Some(&notes_master_rel_id),
            handout_rel_id.as_deref(),
        )?;
        temp_pres_part.set_blob(pres_xml.into_bytes());

        // Add the presentation part to the package
        self.opc.add_part(Box::new(temp_pres_part));

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::NamedTempFile;

    fn workspace() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
    }

    #[test]
    fn saves_and_reopens_package() {
        let file = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 1);
        assert!(
            !package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
    }

    #[test]
    fn handout_theme_allocation_does_not_overwrite_an_existing_theme() {
        let output = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut package = Package::new().unwrap();
        let default_theme_uri = PackURI::new("/ppt/theme/theme1.xml").unwrap();
        let occupied_theme_uri = PackURI::new("/ppt/theme/theme3.xml").unwrap();
        let default_theme = package
            .opc
            .get_part(&default_theme_uri)
            .unwrap()
            .blob()
            .to_vec();
        package
            .opc
            .add_part(Box::new(litchi_opc::part::BlobPart::new(
                occupied_theme_uri,
                ct::OFC_THEME.to_owned(),
                default_theme,
            )));
        package
            .presentation_mut()
            .unwrap()
            .set_handout_master(crate::pptx::HandoutMaster::new());
        package.save(output.path()).unwrap();

        let reopened = Package::open(output.path()).unwrap();
        let handout_uri = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();
        let handout = reopened.opc().unwrap().get_part(&handout_uri).unwrap();
        let theme = handout
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype() == litchi_opc::constants::relationship_type::THEME
            })
            .expect("handout theme relationship");
        assert_eq!(theme.target_ref(), "../theme/theme4.xml");
        assert!(
            reopened
                .opc()
                .unwrap()
                .contains_part(&PackURI::new("/ppt/theme/theme4.xml").unwrap())
        );
    }

    #[test]
    fn failed_publication_restores_package_and_retains_dirty_intent() {
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        package.put_props(Props::new().title("Retry title"));
        let original = package.opc.get_part(&presentation_uri).unwrap().blob_arc();

        let error = package
            .write_with::<()>(|staged| {
                let presentation = staged.get_part(&presentation_uri)?.blob();
                assert!(memchr::memmem::find(presentation, b"<p:sldId ").is_some());
                assert_eq!(
                    litchi_ooxml_common::properties::read(staged)?.and_then(|props| props.title),
                    Some("Retry title".to_owned())
                );
                Err(OoxmlError::Io(std::io::Error::other(
                    "injected sink failure",
                )))
            })
            .unwrap_err();
        assert!(matches!(error, OoxmlError::Io(_)));

        let restored = package.opc.get_part(&presentation_uri).unwrap().blob_arc();
        assert!(std::sync::Arc::ptr_eq(&original, &restored));
        assert!(
            package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
        assert!(package.properties.is_dirty());

        let file = NamedTempFile::with_suffix(".pptx").unwrap();
        package.save(file.path()).unwrap();
        assert!(
            !package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
        assert!(!package.properties.is_dirty());
        let reopened = Package::open(file.path()).unwrap();
        assert_eq!(
            reopened.props().and_then(|props| props.title.as_deref()),
            Some("Retry title")
        );
    }

    #[test]
    fn raw_opc_access_rejects_pending_facade_state() {
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();

        assert!(matches!(
            package.opc(),
            Err(OoxmlError::UnsafeEdit {
                operation: "opc",
                ..
            })
        ));
        assert!(matches!(
            package.edit_opc(|_| Ok(())),
            Err(OoxmlError::UnsafeEdit {
                operation: "edit_opc",
                ..
            })
        ));

        let bytes = package.to_bytes().unwrap();
        assert!(!bytes.is_empty());
        assert!(package.opc().is_ok());

        package.put_props(Props::new().title("pending"));
        assert!(matches!(
            package.opc(),
            Err(OoxmlError::UnsafeEdit {
                operation: "opc",
                ..
            })
        ));
    }

    #[test]
    fn raw_transaction_disables_the_legacy_writer_after_commit() {
        let mut package = Package::new().unwrap();
        package.edit_opc(|_| Ok(())).unwrap();
        assert!(matches!(
            package.presentation_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "presentation_mut",
                ..
            })
        ));
    }

    #[test]
    fn failed_raw_transaction_rolls_back_and_keeps_the_writer_available() {
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let mut package = Package::new().unwrap();
        let original = package.opc.get_part(&presentation_uri).unwrap().blob_arc();

        let error = package
            .edit_opc(|candidate| {
                candidate
                    .get_part_mut(&presentation_uri)?
                    .set_blob(b"not a presentation".to_vec());
                Err::<(), _>(OoxmlError::InvalidFormat("injected raw failure".to_owned()))
            })
            .unwrap_err();
        assert!(matches!(error, OoxmlError::InvalidFormat(_)));
        assert!(std::sync::Arc::ptr_eq(
            &original,
            &package.opc.get_part(&presentation_uri).unwrap().blob_arc()
        ));
        assert!(package.presentation_mut().is_ok());
    }

    #[test]
    fn raw_transaction_reloads_core_properties_before_commit() {
        let mut package = Package::new().unwrap();
        package
            .edit_opc(|candidate| {
                litchi_ooxml_common::properties::write(
                    candidate,
                    Props::new().title("Raw transaction title"),
                )?;
                Ok(())
            })
            .unwrap();

        assert_eq!(
            package.props().and_then(|props| props.title.as_deref()),
            Some("Raw transaction title")
        );
        assert!(matches!(
            package.presentation_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "presentation_mut",
                ..
            })
        ));
    }

    #[test]
    fn raw_transaction_rejects_invalid_main_graph_without_committing() {
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let mut package = Package::new().unwrap();
        let error = package
            .edit_opc(|candidate| {
                assert!(candidate.remove_part(&presentation_uri));
                Ok(())
            })
            .unwrap_err();

        assert!(matches!(error, OoxmlError::PartNotFound(_)));
        assert!(package.opc.get_part(&presentation_uri).is_ok());
        assert!(package.presentation_mut().is_ok());
    }

    #[test]
    fn canonical_tag_and_master_edits_disable_the_legacy_writer() {
        let mut tags = Package::new().unwrap();
        tags.presentation_mut().unwrap().add_slide().unwrap();
        tags.to_bytes().unwrap();
        tags.put_tags(0, tag::List::new()).unwrap();
        assert!(matches!(
            tags.presentation_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "presentation_mut",
                ..
            })
        ));

        let mut masters = Package::new().unwrap();
        masters.add_slide_master().unwrap();
        assert!(matches!(
            masters.presentation_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "presentation_mut",
                ..
            })
        ));
    }

    #[cfg(feature = "fonts")]
    #[test]
    fn canonical_edit_rejects_even_a_synchronized_automatic_font_policy() {
        let mut package = Package::new().unwrap();
        package
            .set_font_embedding(litchi_opc::FontEmbedding::Subset)
            .unwrap();
        package.to_bytes().unwrap();
        assert!(!package.font_sync_pending);

        assert!(matches!(
            package.add_slide_master(),
            Err(OoxmlError::UnsafeEdit {
                operation: "add_slide_master",
                ..
            })
        ));
        package
            .set_font_embedding(litchi_opc::FontEmbedding::None)
            .unwrap();
        package.add_slide_master().unwrap();
    }

    #[test]
    fn canonical_slide_mutators_reject_a_dirty_legacy_writer() {
        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();

        assert!(matches!(
            package.add_ink_annotation(&slide, b""),
            Err(OoxmlError::UnsafeEdit {
                operation: "add_ink_annotation",
                ..
            })
        ));
        assert!(matches!(
            package.add_laser_trace(&slide, &[]),
            Err(OoxmlError::UnsafeEdit {
                operation: "add_laser_trace",
                ..
            })
        ));
        assert!(matches!(
            package.add_slide_show_events(&slide, &[]),
            Err(OoxmlError::UnsafeEdit {
                operation: "add_slide_show_events",
                ..
            })
        ));
        assert!(matches!(
            package.add_ole_object(
                slide.as_str(),
                crate::pptx::ole::PptxOlePayloadKind::OleObject,
                None,
                None,
                crate::pptx::ole_object::OleObjectFrame::new(0, 0, 1, 1),
                b"",
            ),
            Err(OoxmlError::UnsafeEdit {
                operation: "add_ole_object",
                ..
            })
        ));
    }

    #[test]
    fn successful_canonical_slide_edit_disables_the_legacy_writer() {
        let file = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        package.save(file.path()).unwrap();

        package
            .add_ole_object(
                "/ppt/slides/slide1.xml",
                crate::pptx::ole::PptxOlePayloadKind::Package,
                None,
                None,
                crate::pptx::ole_object::OleObjectFrame::new(0, 0, 1, 1),
                b"opaque",
            )
            .unwrap();
        assert!(matches!(
            package.presentation_mut(),
            Err(OoxmlError::UnsafeEdit {
                operation: "presentation_mut",
                ..
            })
        ));
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn failed_encrypted_file_sink_keeps_edits_retryable() {
        const PASSWORD: &str = "transaction retry password";

        let directory = tempfile::tempdir().unwrap();
        let missing_parent = directory.path().join("missing").join("failed.pptx");
        let output = directory.path().join("retry.pptx");
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        package.put_props(Props::new().title("Encrypted retry"));

        assert!(
            package
                .save_encrypted(&missing_parent, PASSWORD, Mode::Agile)
                .is_err()
        );
        assert!(!missing_parent.exists());
        assert_eq!(package.encryption(), None);
        assert!(
            package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
        assert!(package.properties.is_dirty());

        package
            .save_encrypted(&output, PASSWORD, Mode::Agile)
            .unwrap();
        assert_eq!(package.encryption(), Some(Mode::Agile));
        assert!(
            !package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
        assert!(!package.properties.is_dirty());

        let reopened = Package::open_with_password(&output, PASSWORD).unwrap();
        assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 1);
        assert_eq!(
            reopened.props().and_then(|props| props.title.as_deref()),
            Some("Encrypted retry")
        );
    }

    #[cfg(feature = "encryption")]
    #[test]
    fn encrypted_source_rejects_raw_plaintext_access() {
        const PASSWORD: &str = "raw access guard password";

        let mut source = Package::new().unwrap();
        let encrypted = source.to_encrypted(PASSWORD, Mode::Agile).unwrap();
        let mut package =
            Package::from_reader_with_password(std::io::Cursor::new(encrypted), PASSWORD).unwrap();

        assert!(matches!(
            package.opc(),
            Err(OoxmlError::UnsafeEdit {
                operation: "opc",
                ..
            })
        ));
        assert!(matches!(
            package.edit_opc(|_| Ok(())),
            Err(OoxmlError::UnsafeEdit {
                operation: "edit_opc",
                ..
            })
        ));
        assert!(matches!(
            package.to_bytes(),
            Err(OoxmlError::UnsafeEdit {
                operation: "to_bytes",
                ..
            })
        ));
        assert!(!package.to_plain_bytes().unwrap().is_empty());
    }

    #[cfg(feature = "fonts")]
    #[test]
    fn opened_presentation_rejects_font_embedding_policy_immediately() {
        let file = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut source = Package::new().unwrap();
        source.presentation_mut().unwrap().add_slide().unwrap();
        source.save(file.path()).unwrap();

        let mut opened = Package::open(file.path()).unwrap();
        assert!(matches!(
            opened.set_font_embedding(litchi_opc::FontEmbedding::Subset),
            Err(OoxmlError::UnsafeEdit {
                operation: "set_font_embedding",
                ..
            })
        ));
        assert_eq!(
            opened.opc.save_options().fonts,
            litchi_opc::FontEmbedding::None
        );
        assert!(!opened.font_sync_pending);
    }

    #[cfg(feature = "fonts")]
    #[test]
    fn font_policy_syncs_once_and_rearms_only_when_needed() {
        let mut package = Package::new().unwrap();
        package
            .set_font_embedding(litchi_opc::FontEmbedding::Subset)
            .unwrap();
        assert!(package.font_sync_pending);
        assert!(matches!(
            package.opc(),
            Err(OoxmlError::UnsafeEdit {
                operation: "opc",
                ..
            })
        ));

        package.to_bytes().unwrap();
        assert!(!package.font_sync_pending);
        assert!(package.opc().is_ok());

        package
            .set_font_embedding(litchi_opc::FontEmbedding::Subset)
            .unwrap();
        assert!(!package.font_sync_pending);
        package.to_bytes().unwrap();
        assert!(!package.font_sync_pending);

        let _ = package.presentation_mut().unwrap();
        assert!(package.font_sync_pending);
        package.to_bytes().unwrap();
        assert!(!package.font_sync_pending);

        package
            .set_font_embedding(litchi_opc::FontEmbedding::Full)
            .unwrap();
        assert!(package.font_sync_pending);
        let error = package
            .write_with::<()>(|_| {
                Err(OoxmlError::Io(std::io::Error::other(
                    "injected font publication failure",
                )))
            })
            .unwrap_err();
        assert!(matches!(error, OoxmlError::Io(_)));
        assert!(package.font_sync_pending);
        package.to_bytes().unwrap();
        assert!(!package.font_sync_pending);
    }

    #[test]
    fn pending_writer_flush_preserves_font_graph_across_later_saves() {
        let source = workspace()
            .join("test-data/libreoffice-core/sd/qa/unit/data/BoldonseFontEmbedded.pptx");
        let source = Package::open(source).unwrap();
        let fonts = source.fonts().unwrap().unwrap();

        let first = NamedTempFile::with_suffix(".pptx").unwrap();
        let second = NamedTempFile::with_suffix(".pptx").unwrap();
        let mut package = Package::new().unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        assert!(package.put_fonts(fonts).unwrap());
        assert!(
            !package
                .mutable_pres
                .as_ref()
                .is_some_and(MutablePresentation::is_modified)
        );
        package.save(first.path()).unwrap();

        package.presentation_mut().unwrap().add_slide().unwrap();
        package.save(second.path()).unwrap();
        package.save(second.path()).unwrap();

        let reopened = Package::open(second.path()).unwrap();
        assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 2);
        let fonts = reopened.fonts().unwrap().unwrap();
        assert_eq!(fonts.len(), 1);
        assert_eq!(fonts.get("boldonse").unwrap().faces().len(), 1);
    }
}

/// Package implementation for PowerPoint presentations.
use crate::ooxml::common::DocumentProperties;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::packuri::PackURI;
use crate::ooxml::opc::part::Part;
use crate::ooxml::pptx::parts::PresentationPart;
use crate::ooxml::pptx::presentation::Presentation;
use crate::ooxml::pptx::writer::MutablePresentation;
use std::io::{Read, Seek};
use std::path::Path;

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

/// A PowerPoint (.pptx) package.
///
/// This is the main entry point for working with PowerPoint presentations.
/// It wraps an OPC package and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::pptx::Package;
///
/// // Open an existing presentation
/// let pkg = Package::open("presentation.pptx")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Access slides
/// println!("Presentation has {} slides", pres.slide_count()?);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    /// The underlying OPC package
    opc: OpcPackage,
    /// Mutable presentation for writing (cached)
    mutable_pres: Option<MutablePresentation>,
    /// Document properties (metadata)
    properties: DocumentProperties,
}

impl Package {
    /// Create a new empty .pptx package.
    ///
    /// Creates a minimal valid PowerPoint presentation with default master slide and layout.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::new()?;
    /// // Add slides to the presentation...
    /// pkg.save("new_presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn new() -> Result<Self> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::packuri::PackURI;
        use crate::ooxml::opc::part::BlobPart;
        use crate::ooxml::pptx::template;

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

        // Create tableStyles.xml
        let table_styles_partname = PackURI::new("/ppt/tableStyles.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("tableStyles partname: {}", e)))?;
        let table_styles_part = BlobPart::new(
            table_styles_partname,
            ct::PML_TABLE_STYLES.to_string(),
            template::default_table_styles_xml().as_bytes().to_vec(),
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
            template::default_notes_master_xml().as_bytes().to_vec(),
        );

        // Add relationship from notesMaster to theme
        notes_master_part.relate_to("../theme/theme1.xml", rt::THEME);

        // Add relationship from presentation to notesMaster
        if let Ok(pres_part) = opc.get_part_mut(&pres_partname) {
            pres_part.relate_to("notesMasters/notesMaster1.xml", rt::NOTES_MASTER);
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
        let properties = DocumentProperties::new();

        Ok(Self {
            opc,
            mutable_pres,
            properties,
        })
    }

    /// Open a .pptx package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .pptx file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let opc = OpcPackage::open(path)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    #[cfg(feature = "ooxml_encryption")]
    pub fn open_with_password<P: AsRef<Path>>(path: P, password: &str) -> Result<Self> {
        let data = std::fs::read(path.as_ref()).map_err(OoxmlError::Io)?;
        let decrypted = crate::ooxml::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
        let opc = OpcPackage::from_bytes(&decrypted.package_bytes)?;
        Self::from_opc_package(opc)
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
    /// use litchi::ooxml::{OpcPackage, pptx::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.pptx")?;
    /// let opc = OpcPackage::from_reader(Cursor::new(bytes))?;
    /// let pkg = Package::from_opc_package(opc)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    /// Create a .pptx package from a reader.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .pptx file data (must implement Read + Seek)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    /// use std::io::Cursor;
    ///
    /// let data = std::fs::read("presentation.pptx")?;
    /// let cursor = Cursor::new(data);
    /// let pkg = Package::from_reader(cursor)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let opc = OpcPackage::from_reader(reader)?;

        // Verify it's a PowerPoint presentation by checking the main part's content type
        let main_part = opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        let content_type = main_part.content_type();
        // Support both regular and macro-enabled presentations
        if content_type != ct::PML_PRESENTATION_MAIN && content_type != ct::PML_PRES_MACRO_MAIN {
            return Err(OoxmlError::InvalidContentType {
                expected: format!(
                    "{} or {}",
                    ct::PML_PRESENTATION_MAIN,
                    ct::PML_PRES_MACRO_MAIN
                ),
                got: content_type.to_string(),
            });
        }

        Ok(Self {
            opc,
            mutable_pres: None,
            properties: DocumentProperties::new(),
        })
    }

    /// Get the main presentation for reading.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// // Access slides
    /// for slide in pres.slides()? {
    ///     println!("Slide text: {}", slide.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        let main_part = self
            .opc
            .main_document_part()
            .map_err(|e| OoxmlError::PartNotFound(format!("main presentation part: {}", e)))?;

        // Create PresentationPart wrapper
        let pres_part = PresentationPart::from_part(main_part)?;

        // Create and return Presentation
        Ok(Presentation::new(pres_part, &self.opc))
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Get mutable access to the underlying OPC package.
    ///
    /// This provides access to lower-level package operations for modification.
    #[inline]
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        &mut self.opc
    }

    /// Get a mutable presentation for writing and modification.
    ///
    /// This returns a `MutablePresentation` that allows you to add and modify
    /// slides, shapes, and other presentation elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation_mut(&mut self) -> Result<&mut MutablePresentation> {
        // If we don't have a mutable presentation, create one
        if self.mutable_pres.is_none() {
            self.mutable_pres = Some(MutablePresentation::new());
        }

        Ok(self.mutable_pres.as_mut().unwrap())
    }

    /// Get a reference to the presentation properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let props = pkg.properties();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the presentation properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// pkg.properties_mut().title = Some("My Presentation".to_string());
    /// pkg.properties_mut().creator = Some("John Doe".to_string());
    /// pkg.save("presentation.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
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
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// // Modify presentation...
    /// pkg.save("output.pptx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        // If we have a mutable presentation, update the presentation parts
        let should_update = self
            .mutable_pres
            .as_ref()
            .map(|p| p.is_modified())
            .unwrap_or(false);

        if should_update {
            // Take mutable_pres temporarily to avoid borrow issues
            if let Some(mutable_pres) = self.mutable_pres.take() {
                self.update_presentation_parts(&mutable_pres)?;
                self.mutable_pres = Some(mutable_pres);
            }
        }

        // Update core properties
        self.update_core_properties()?;

        #[cfg(feature = "ooxml_encryption")]
        #[allow(clippy::collapsible_if)]
        {
            if let Some(pres) = self.mutable_pres.as_ref() {
                let prot = pres.protection();
                if prot.open_password_protected {
                    if let Some(password) = prot.open_password() {
                        use crate::ooxml::crypto::{
                            encrypt_ooxml_package_agile, encrypt_ooxml_package_standard_2007,
                        };
                        use crate::ooxml::opc::pkgwriter::PackageWriter;
                        use crate::ooxml::pptx::OpenPasswordEncryption;

                        let pkg_bytes = PackageWriter::to_bytes(&self.opc)?;

                        let ole_bytes = match prot.open_password_encryption() {
                            OpenPasswordEncryption::Standard2007 => {
                                encrypt_ooxml_package_standard_2007(&pkg_bytes, password)?
                            },
                            OpenPasswordEncryption::Agile => {
                                encrypt_ooxml_package_agile(&pkg_bytes, password)?
                            },
                        };

                        std::fs::write(&path, ole_bytes).map_err(|e| {
                            OoxmlError::IoError(std::io::Error::other(format!(
                                "Failed to save encrypted package: {}",
                                e
                            )))
                        })?;

                        return Ok(());
                    }
                }
            }
        }

        self.opc.save(path).map_err(|e| {
            OoxmlError::IoError(std::io::Error::other(format!(
                "Failed to save package: {}",
                e
            )))
        })
    }

    /// Update presentation parts with modified data.
    fn update_presentation_parts(&mut self, pres: &MutablePresentation) -> Result<()> {
        use crate::ooxml::opc::constants::content_type as ct;
        use crate::ooxml::opc::constants::relationship_type as rt;
        use crate::ooxml::opc::part::{BlobPart, Part};
        use crate::ooxml::pptx::parts::CommentAuthor;
        use crate::ooxml::pptx::parts::{generate_comment_authors_xml, generate_comments_xml};
        use crate::ooxml::pptx::template;
        use crate::ooxml::pptx::writer::relmap::RelationshipMapper;

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
                if let crate::ooxml::pptx::writer::shape::ShapeType::SmartArt {
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
            let drawing_xml = crate::ooxml::pptx::smartart::generate_smartart_drawing_xml(
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

        // Create a temporary presentation part to manage relationships
        let mut temp_pres_part = BlobPart::new(
            pres_uri.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            Vec::new(),
        );

        // Add relationship to slideMaster (this should be rId1)
        temp_pres_part.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

        // Add other required relationships (in the order they were created in Package::new())
        // These relationships should be added even if not modified, as they're required for a valid PPTX
        temp_pres_part.relate_to("tableStyles.xml", rt::TABLE_STYLES);
        temp_pres_part.relate_to("viewProps.xml", rt::VIEW_PROPS);
        temp_pres_part.relate_to("presProps.xml", rt::PRES_PROPS);
        temp_pres_part.relate_to("theme/theme1.xml", rt::THEME);

        // Add relationship to notesMaster (required when we have notesSlides)
        let _notes_master_rel_id =
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
                            crate::ooxml::pptx::media::MediaType::Audio => rt::AUDIO,
                            crate::ooxml::pptx::media::MediaType::Video => rt::VIDEO,
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
                if let crate::ooxml::pptx::writer::shape::ShapeType::Chart { chart_idx, .. } =
                    &shape.shape_type
                {
                    let chart_rel_target = format!("../charts/chart{}.xml", chart_idx);
                    let rid = temp_slide_part.relate_to(&chart_rel_target, rt::CHART);
                    rel_mapper.add_chart(slide_index, *chart_idx, rid);
                }
            }

            // Add relationships for SmartArt diagrams on this slide
            for shape in &slide.shapes {
                if let crate::ooxml::pptx::writer::shape::ShapeType::SmartArt {
                    diagram_idx, ..
                } = &shape.shape_type
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
            if let Some(notes_xml_result) = slide.generate_notes_xml() {
                let notes_xml = notes_xml_result?;
                let notes_uri =
                    PackURI::new(format!("/ppt/notesSlides/notesSlide{}.xml", slide_num)).map_err(
                        |e| OoxmlError::InvalidUri(format!("notesSlide{} URI: {}", slide_num, e)),
                    )?;

                let mut notes_part = BlobPart::new(
                    notes_uri,
                    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
                        .to_string(),
                    notes_xml.into_bytes(),
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
            // Create theme2.xml for handout master (required - handout needs its own theme)
            let theme2_uri = PackURI::new("/ppt/theme/theme2.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("theme2 URI: {}", e)))?;
            let theme2_part = BlobPart::new(
                theme2_uri,
                ct::OFC_THEME.to_string(),
                template::default_theme_xml().as_bytes().to_vec(),
            );
            self.opc.add_part(Box::new(theme2_part));

            let handout_uri = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml")
                .map_err(|e| OoxmlError::InvalidUri(format!("handoutMaster URI: {}", e)))?;

            let mut handout_part = BlobPart::new(
                handout_uri,
                "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
                    .to_string(),
                handout_master.to_xml().into_bytes(),
            );

            // Add relationship from handoutMaster to its own theme (theme2.xml)
            handout_part.relate_to("../theme/theme2.xml", rt::THEME);

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
            Some(&slide_rel_ids),
            None, // notesMasterIdLst not needed
            handout_rel_id.as_deref(),
        )?;
        temp_pres_part.set_blob(pres_xml.into_bytes());

        // Add the presentation part to the package
        self.opc.add_part(Box::new(temp_pres_part));

        Ok(())
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> Result<()> {
        use crate::ooxml::opc::part::BlobPart;

        let core_uri = PackURI::new("/docProps/core.xml")
            .map_err(|e| OoxmlError::InvalidUri(format!("core.xml URI: {}", e)))?;

        // Generate XML from properties
        let xml = self.properties.to_xml();

        // Create or update the core properties part
        let core_part = BlobPart::new(
            core_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.opc.add_part(Box::new(core_part));

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[ignore] // Requires test file
    fn test_open_package() {
        let result = Package::open("test.pptx");
        assert!(result.is_ok());
    }
}

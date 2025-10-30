/// Package implementation for PowerPoint presentations.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::OpcPackage;
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::part::Part;
use crate::ooxml::pptx::parts::PresentationPart;
use crate::ooxml::pptx::presentation::Presentation;
use std::io::{Read, Seek};
use std::path::Path;

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

        Ok(Self { opc })
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

        Ok(Self { opc })
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

        Ok(Self { opc })
    }

    /// Get the main presentation.
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
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.opc.save(path).map_err(|e| {
            OoxmlError::IoError(std::io::Error::other(format!(
                "Failed to save package: {}",
                e
            )))
        })
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

//! Lossless-bounded PPTX package facade.

use std::io::Read;
use std::path::Path;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackageWriter};

use crate::parts::PresentationPart;
use crate::presentation::Presentation;
use crate::writer::MutablePresentation;
use crate::{Error, Result};

const PRESENTATION_PART: &str = "/ppt/presentation.xml";
const MASTER_PART: &str = "/ppt/slideMasters/slideMaster1.xml";
const LAYOUT_PART: &str = "/ppt/slideLayouts/slideLayout1.xml";
const THEME_PART: &str = "/ppt/theme/theme1.xml";

const STALE_PRESENTATION_GRAPH_REASON: &str = "the mutable presentation model has unflushed changes; save and reopen before reading the canonical package graph";

/// Main entry point for PresentationML package ownership.
pub struct Package {
    opc: OpcPackage,
    mutable_pres: Option<MutablePresentation>,
}

impl Package {
    /// Create a minimal valid, mutable PresentationML package.
    pub fn new() -> Result<Self> {
        let mut package = OpcPackage::new();
        let presentation_name = pack_uri(PRESENTATION_PART)?;
        let master_name = pack_uri(MASTER_PART)?;
        let layout_name = pack_uri(LAYOUT_PART)?;
        let theme_name = pack_uri(THEME_PART)?;

        let mut presentation = BlobPart::new(
            presentation_name.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            MutablePresentation::new()
                .generate_presentation_xml()?
                .into_bytes(),
        );
        presentation.relate_to("slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

        let mut master = BlobPart::new(
            master_name,
            ct::PML_SLIDE_MASTER.to_string(),
            default_slide_master_xml().as_bytes().to_vec(),
        );
        master.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);
        master.relate_to("../theme/theme1.xml", rt::THEME);

        let mut layout = BlobPart::new(
            layout_name,
            ct::PML_SLIDE_LAYOUT.to_string(),
            default_slide_layout_xml().as_bytes().to_vec(),
        );
        layout.relate_to("../slideMasters/slideMaster1.xml", rt::SLIDE_MASTER);

        let theme = BlobPart::new(
            theme_name,
            ct::OFC_THEME.to_string(),
            default_theme_xml().as_bytes().to_vec(),
        );

        package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(master));
        package.add_part(Box::new(layout));
        package.add_part(Box::new(theme));

        Ok(Self {
            opc: package,
            mutable_pres: Some(MutablePresentation::new()),
        })
    }

    /// Open a PPTX from a filesystem path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_opc_package(OpcPackage::open(path)?)
    }

    /// Parse a PPTX from a reader.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_reader(reader)?)
    }

    /// Parse a PPTX from an owned ZIP buffer.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_vec(bytes)?)
    }

    /// Parse a PPTX from a borrowed ZIP buffer.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_opc_package(OpcPackage::from_bytes(bytes)?)
    }

    /// Adopt an already parsed OPC graph without hydrating a mutable model.
    pub fn from_opc_package(opc: OpcPackage) -> Result<Self> {
        let main = opc.main_document_part()?;
        PresentationPart::from_part(main)?;
        Ok(Self {
            opc,
            mutable_pres: None,
        })
    }

    /// Borrow the canonical presentation graph when no mutable state is stale.
    pub fn presentation(&self) -> Result<Presentation<'_>> {
        self.ensure_graph_current("presentation")?;
        Ok(Presentation::new(
            PresentationPart::from_package(&self.opc)?,
            &self.opc,
        ))
    }

    /// Borrow the mutable presentation model for a newly authored package.
    pub fn presentation_mut(&mut self) -> Result<&mut MutablePresentation> {
        self.mutable_pres.as_mut().ok_or(Error::UnsafeEdit {
            operation: "presentation_mut",
            reason: "the lossless facade cannot hydrate an opened package into the mutable writer",
        })
    }

    /// Whether a mutable model is currently pending managed publication.
    pub fn is_modified(&self) -> bool {
        self.mutable_pres
            .as_ref()
            .is_some_and(MutablePresentation::is_modified)
    }

    /// Borrow the underlying OPC graph when it is current.
    pub fn opc(&self) -> Result<&OpcPackage> {
        self.ensure_graph_current("opc")?;
        Ok(&self.opc)
    }

    /// Save the package atomically through the OPC writer.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        self.flush_presentation()?;
        PackageWriter::write(path, &self.opc)?;
        Ok(())
    }

    /// Serialize the package into a new ZIP buffer.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush_presentation()?;
        Ok(PackageWriter::to_bytes(&self.opc)?)
    }

    /// Run a read-only operation against the current OPC graph.
    pub fn with_opc<T>(&self, operation: impl FnOnce(&OpcPackage) -> Result<T>) -> Result<T> {
        self.ensure_graph_current("with_opc")?;
        operation(&self.opc)
    }

    fn ensure_graph_current(&self, operation: &'static str) -> Result<()> {
        if self.is_modified() {
            return Err(Error::UnsafeEdit {
                operation,
                reason: STALE_PRESENTATION_GRAPH_REASON,
            });
        }
        Ok(())
    }

    fn flush_presentation(&mut self) -> Result<()> {
        if !self.is_modified() {
            return Ok(());
        }
        let Some(presentation) = self.mutable_pres.as_ref().cloned() else {
            return Err(Error::UnsafeEdit {
                operation: "save",
                reason: "the lossless facade cannot publish an opened package's mutable graph",
            });
        };
        let before = self.opc.clone();
        match self.materialize_presentation(&presentation) {
            Ok(()) => {
                if let Some(presentation) = self.mutable_pres.as_mut() {
                    presentation.mark_clean();
                }
                Ok(())
            },
            Err(error) => {
                self.opc = before;
                Err(error)
            },
        }
    }

    fn materialize_presentation(&mut self, presentation: &MutablePresentation) -> Result<()> {
        let presentation_name = pack_uri(PRESENTATION_PART)?;
        let old_slide_relationships: Vec<String> = {
            let part = self.opc.get_part(&presentation_name)?;
            part.rels()
                .iter()
                .filter(|relationship| {
                    crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
                })
                .map(|relationship| relationship.r_id().to_string())
                .collect()
        };
        {
            let part = self.opc.get_part_mut(&presentation_name)?;
            for relationship_id in old_slide_relationships {
                part.rels_mut().remove(&relationship_id);
            }
        }

        let old_slides: Vec<_> = self
            .opc
            .iter_parts()
            .filter(|part| part.partname().as_str().starts_with("/ppt/slides/"))
            .map(|part| part.partname().clone())
            .collect();
        for part in old_slides {
            self.opc.remove_part(&part);
        }

        let mut relationship_ids = Vec::with_capacity(presentation.slide_count());
        for (index, slide) in presentation.slides().iter().enumerate() {
            let target = format!("slides/slide{}.xml", index.saturating_add(1));
            let relationship_id = self
                .opc
                .get_part_mut(&presentation_name)?
                .relate_to(&target, rt::SLIDE);
            relationship_ids.push(relationship_id);

            let slide_name = pack_uri(&format!("/ppt/{target}"))?;
            let mut slide_part = BlobPart::new(
                slide_name,
                ct::PML_SLIDE.to_string(),
                slide.generate_slide_xml()?.into_bytes(),
            );
            slide_part.relate_to("../slideLayouts/slideLayout1.xml", rt::SLIDE_LAYOUT);
            self.opc.add_part(Box::new(slide_part));
        }

        let xml = presentation.generate_presentation_xml_with(&relationship_ids)?;
        self.opc
            .get_part_mut(&presentation_name)?
            .set_blob(xml.into_bytes());
        Ok(())
    }
}

fn pack_uri(value: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(Error::Uri)
}

fn default_slide_master_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Office Theme"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles/></p:sldMaster>"#
}

fn default_slide_layout_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"#
}

fn default_theme_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>"#
}

#[cfg(test)]
mod tests {
    use super::Package;

    #[test]
    fn new_writer_round_trips_the_bounded_slide_graph() {
        let mut package = Package::new().expect("new package");
        {
            let presentation = package.presentation_mut().expect("mutable presentation");
            let slide = presentation.add_slide().expect("slide");
            slide.set_title("Canonical owner");
            slide.add_text_box("Hello & goodbye", 914_400, 914_400, 2_743_200, 914_400);
            presentation.set_widescreen_slide_size();
        }

        let bytes = package.to_bytes().expect("serialize package");
        let reopened = Package::from_bytes(&bytes).expect("reopen package");
        let presentation = reopened.presentation().expect("presentation");
        assert_eq!(presentation.slide_count().expect("slide count"), 1);
        assert_eq!(
            presentation.slide_size().expect("slide size"),
            (9_144_000, 5_143_500)
        );
        let slide = presentation.slide(0).expect("slide lookup").expect("slide");
        assert_eq!(slide.name().expect("slide name"), "Slide 256");
        assert!(
            slide
                .text()
                .expect("slide text")
                .contains("Hello & goodbye")
        );
        assert_eq!(slide.shape_count().expect("shape count"), 2);
        assert_eq!(presentation.slide_masters().expect("masters").len(), 1);
        assert_eq!(presentation.slide_layouts().expect("layouts").len(), 1);
    }

    #[test]
    fn opened_package_refuses_unsafe_mutable_hydration() {
        let mut package = Package::new().expect("new package");
        let bytes = package.to_bytes().expect("serialize package");
        let mut opened = Package::from_bytes(&bytes).expect("reopen package");
        assert!(opened.presentation_mut().is_err());
    }
}

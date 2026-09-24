//! Source-backed attachment and detachment checks for the optional SVG blip.
//!
//! These tests deliberately build a small OPC package at the physical package
//! boundary.  The slide XML is kept in the fixture so each lifecycle assertion
//! can compare the exact source member and every unrelated physical member.

use std::collections::BTreeMap;
use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits, TargetMode};
use litchi_pptx::{
    Error, SourceBackedPresentation, SourceBackedPresentationEditor, SourceSvgAttachmentReplacement,
};
use sha2::{Digest, Sha256};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const STRICT_DRAWINGML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_IMAGE: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const RASTER: &str = "/ppt/media/raster.png";
const SVG: &str = "/ppt/media/vector.svg";
const OPAQUE: &str = "/ppt/media/opaque.bin";
const UTF8_BOM: &[u8] = b"\xEF\xBB\xBF";

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            911,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

/// Build a slide whose direct picture is a raster-only owner.  `self_closing`
/// selects the lexical form of `a:blip`; the semantic owner is identical.
fn raster_slide_xml(self_closing: bool, ext_lst: bool) -> Vec<u8> {
    let extensions = if ext_lst {
        r#"<a:extLst><a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext></a:extLst>"#
    } else {
        ""
    };
    let blip = if self_closing {
        r#"<a:blip r:embed="rIdRaster"/>"#.to_owned()
    } else {
        format!(r#"<a:blip r:embed="rIdRaster">{extensions}</a:blip>"#)
    };
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="Raster photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
    .into_bytes()
}

fn paired_blip_slide_xml(self_closing: bool) -> Vec<u8> {
    paired_blip_slide_xml_with_uri(self_closing, SVG_URI)
}

fn paired_blip_slide_xml_with_uri(self_closing: bool, svg_uri: &str) -> Vec<u8> {
    let blip = if self_closing {
        r#"<a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext><a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext></a:extLst></a:blip>"#
    } else {
        r#"<a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"></asvg:svgBlip></a:ext><a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext></a:extLst></a:blip>"#
    };
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="Paired SVG photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
    .replace("{SVG_URI}", svg_uri)
    .into_bytes()
}

/// Retain the source tree while moving every PresentationML namespace to a
/// noncanonical inherited prefix.  The default-drawing variant exercises
/// unprefixed DrawingML names under an inherited default namespace.
fn namespace_variant_slide_xml(default_drawing: bool, strict: bool) -> Vec<u8> {
    let mut xml = String::from_utf8(paired_blip_slide_xml(true)).unwrap();
    xml = xml
        .replace(
            &format!(r#"xmlns:p="{PML}""#),
            &format!(r#"xmlns:ppt="{PML}""#),
        )
        .replace(
            &format!(r#"xmlns:a="{DRAWINGML}""#),
            &format!(r#"xmlns:draw="{DRAWINGML}""#),
        )
        .replace(
            &format!(r#"xmlns:r="{REL}""#),
            &format!(r#"xmlns:rel="{REL}""#),
        )
        .replace(
            &format!(r#"xmlns:asvg="{SVG_NS}""#),
            r#"xmlns:svg="http://schemas.microsoft.com/office/drawing/2016/SVG/main""#,
        )
        .replace("<p:", "<ppt:")
        .replace("</p:", "</ppt:")
        .replace("<a:", "<draw:")
        .replace("</a:", "</draw:")
        .replace("<asvg:", "<svg:")
        .replace("</asvg:", "</svg:")
        .replace(" r:", " rel:");
    if default_drawing {
        xml = xml
            .replace(
                &format!(r#"xmlns:draw="{DRAWINGML}""#),
                &format!(r#"xmlns="{DRAWINGML}""#),
            )
            .replace("<draw:", "<")
            .replace("</draw:", "</");
    }
    if strict {
        xml = xml
            .replace(PML, STRICT_PML)
            .replace(DRAWINGML, STRICT_DRAWINGML)
            .replace(REL, STRICT_REL);
    }
    xml.into_bytes()
}

fn package_from_namespace_variant(default_drawing: bool, strict: bool) -> Vec<u8> {
    source_package(
        namespace_variant_slide_xml(default_drawing, strict),
        true,
        false,
    )
}

fn package_from_strict_raster() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(false, false))
        .unwrap()
        .replace(PML, STRICT_PML)
        .replace(DRAWINGML, STRICT_DRAWINGML)
        .replace(REL, STRICT_REL)
        .into_bytes();
    let mut package = OpcPackage::from_bytes(&source_package(slide, false, false)).unwrap();
    let slide = package.get_part_mut(&PackURI::new(SLIDE).unwrap()).unwrap();
    slide.rels_mut().remove("rIdRaster");
    slide
        .rels_mut()
        .try_add_relationship(
            rt::STRICT_IMAGE.to_owned(),
            "../media/raster.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn foreign_nested_picture_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_pair(true), |xml| {
        xml.replace(
            "</future:unknown>",
            r#"<x:foreign xmlns:x="urn:litchi:foreign"><p:pic xmlns:p="urn:litchi:foreign"/></x:foreign></future:unknown>"#,
        )
    })
}

fn source_package(slide_xml: Vec<u8>, include_svg: bool, signed: bool) -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            ct::PML_PRESENTATION_MAIN.to_owned(),
            format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SLIDE).unwrap(),
            ct::PML_SLIDE.to_owned(),
            slide_xml,
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(RASTER).unwrap(),
            ct::PNG.to_owned(),
            b"old-raster-payload".to_vec(),
        )))
        .unwrap();
    if include_svg {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SVG).unwrap(),
                "image/svg+xml".to_owned(),
                br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>"#.to_vec(),
            )))
            .unwrap();
    }
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(OPAQUE).unwrap(),
            "application/octet-stream".to_owned(),
            b"untouched opaque member".to_vec(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::SLIDE.to_owned(),
            "slides/slide1.xml".to_owned(),
            "rIdSlide".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SLIDE).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/raster.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    if include_svg {
        package
            .get_part_mut(&PackURI::new(SLIDE).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/vector.svg".to_owned(),
                "rIdSvg".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package
        .get_part_mut(&PackURI::new(SLIDE).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:future-image-metadata".to_owned(),
            "../media/opaque.bin".to_owned(),
            "rIdFuture".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn open_editor(bytes: &[u8]) -> SourceBackedPresentationEditor {
    SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(bytes.to_vec())))
        .unwrap()
}

fn physical_member(bytes: &[u8], name: &str) -> Vec<u8> {
    PhysPkgReader::new(bytes)
        .unwrap()
        .read_member(name)
        .unwrap()
}

fn physical_member_names(bytes: &[u8]) -> Vec<String> {
    PhysPkgReader::new(bytes).unwrap().member_names().unwrap()
}

fn physical_members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = PhysPkgReader::new(bytes).unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn assert_unrelated_members_unchanged(before: &[u8], after: &[u8], excluded: &[&str]) {
    let before = physical_members(before);
    let after = physical_members(after);
    for (name, bytes) in before {
        if excluded.contains(&name.as_str()) {
            continue;
        }
        assert_eq!(after.get(&name), Some(&bytes), "member {name} changed");
    }
}

fn slide_member(bytes: &[u8]) -> Vec<u8> {
    physical_member(bytes, "ppt/slides/slide1.xml")
}

fn slide_relationships_member(bytes: &[u8]) -> Vec<u8> {
    physical_member(bytes, "ppt/slides/_rels/slide1.xml.rels")
}

fn native_fixture(name: &str) -> Vec<u8> {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sd/qa/unit/data/pptx")
        .join(name);
    std::fs::read(&fixture).unwrap_or_else(|error| {
        panic!(
            "native SVG fixture {} is unavailable: {error}",
            fixture.display()
        )
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    digest
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn direct_picture_count(bytes: &[u8]) -> usize {
    physical_member(bytes, "ppt/slides/slide1.xml")
        .windows(b"<p:pic>".len())
        .filter(|window| *window == b"<p:pic>")
        .count()
}

fn attach_and_publish(source: &[u8], svg: &[u8]) -> (Vec<u8>, String, String) {
    let editor = open_editor(source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(edit.source().svg().is_none());
    let replacement = SourceSvgAttachmentReplacement::new(svg.to_vec());
    assert!(edit.attach(&replacement).unwrap());
    let commit = edit.commit_checked().unwrap();
    let relationship_id = commit
        .snapshot()
        .svg()
        .unwrap()
        .relationship_id()
        .to_owned();
    let part_uri = commit
        .snapshot()
        .svg()
        .unwrap()
        .part_uri()
        .as_str()
        .to_owned();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    (output, relationship_id, part_uri)
}

fn detach_and_publish(source: &[u8]) -> Vec<u8> {
    let editor = open_editor(source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(edit.source().svg().is_some());
    assert!(edit.detach().unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    output
}

fn package_from_raster(self_closing: bool, ext_lst: bool) -> Vec<u8> {
    source_package(raster_slide_xml(self_closing, ext_lst), false, false)
}

fn package_from_raster_with_prefixed_selfclosing_ext_lst() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(false, false)).unwrap();
    let old = r#"<a:blip r:embed="rIdRaster"></a:blip>"#;
    let replacement = format!(
        r#"<a:blip r:embed="rIdRaster"><d:extLst xmlns:d="{DRAWINGML}" xmlns:a="urn:litchi:foreign" xmlns:x="urn:litchi:future" x:future="keep"/></a:blip>"#
    );
    assert!(slide.contains(old));
    source_package(slide.replace(old, &replacement).into_bytes(), false, false)
}

fn package_from_raster_with_namespace_only_selfclosing_ext_lst() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(false, false)).unwrap();
    let old = r#"<a:blip r:embed="rIdRaster"></a:blip>"#;
    let replacement = format!(
        r#"<a:blip r:embed="rIdRaster"><d:extLst xmlns:d="{DRAWINGML}" xmlns:a="urn:litchi:foreign" xmlns:x="urn:litchi:future"/></a:blip>"#
    );
    assert!(slide.contains(old));
    source_package(slide.replace(old, &replacement).into_bytes(), false, false)
}

fn package_from_bom_raster() -> Vec<u8> {
    let source = source_package(raster_slide_xml(true, false), false, false);
    let reader = PhysPkgReader::new(&source).unwrap();
    let mut writer = PhysPkgWriter::new();
    for member in reader.member_names().unwrap() {
        let uri = PackURI::new(format!("/{member}")).unwrap();
        let mut bytes = reader.read_member(&member).unwrap();
        if member == "ppt/slides/slide1.xml" {
            let mut with_bom = Vec::with_capacity(UTF8_BOM.len() + bytes.len());
            with_bom.extend_from_slice(UTF8_BOM);
            with_bom.append(&mut bytes);
            bytes = with_bom;
        }
        writer.write(&uri, &bytes).unwrap();
    }
    writer.finish().unwrap()
}

fn package_from_raster_with_opaque_pi_and_mce() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(false, true)).unwrap();
    let old = r#"<future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown>"#;
    let replacement = r#"<future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><?future data?><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="future"><future:opaque/></mc:Choice><mc:Fallback><future:opaque/></mc:Fallback></mc:AlternateContent></future:unknown>"#;
    assert!(slide.contains(old));
    source_package(slide.replace(old, replacement).into_bytes(), false, false)
}

fn package_from_raster_with_large_inherited_namespace() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(true, false)).unwrap();
    let marker = format!(r#"xmlns:r="{REL}""#);
    let huge_uri = format!("urn:litchi:inherited:{}", "x".repeat(3_000));
    let replacement = format!(r#"xmlns:r="{REL}" xmlns:huge="{huge_uri}""#);
    assert!(slide.contains(&marker));
    source_package(
        slide.replace(&marker, &replacement).into_bytes(),
        false,
        false,
    )
}

fn package_from_raster_with_standard_namespace_lookalike() -> Vec<u8> {
    let slide = String::from_utf8(raster_slide_xml(false, true)).unwrap();
    let old = r#"<future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown>"#;
    let replacement = format!(
        r#"<future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><p:pic xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:blipFill><a:blip r:embed="rIdFuture"/></p:blipFill></p:pic></future:unknown>"#
    );
    assert!(slide.contains(old));
    source_package(slide.replace(old, &replacement).into_bytes(), false, false)
}

fn package_from_pair(self_closing: bool) -> Vec<u8> {
    source_package(paired_blip_slide_xml(self_closing), true, false)
}

fn package_from_pair_with_uri(self_closing: bool, svg_uri: &str) -> Vec<u8> {
    source_package(
        paired_blip_slide_xml_with_uri(self_closing, svg_uri),
        true,
        false,
    )
}

fn rewrite_slide_member(source: &[u8], rewrite: impl FnOnce(String) -> String) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(source).unwrap();
    let uri = PackURI::new(SLIDE).unwrap();
    let slide = package.get_part_mut(&uri).unwrap();
    let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
    slide.set_blob(rewrite(xml).into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn external_svg_package() -> Vec<u8> {
    let source = package_from_pair(true);
    let mut package = OpcPackage::from_bytes(&source).unwrap();
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let slide = package.get_part_mut(&slide_uri).unwrap();
    let xml = String::from_utf8(slide.blob().to_vec())
        .unwrap()
        .replace("r:embed=\"rIdSvg\"", "r:link=\"rIdSvg\"");
    slide.set_blob(xml.into_bytes());
    let rels = slide.rels_mut();
    rels.remove("rIdSvg");
    rels.try_add_relationship(
        rt::IMAGE.to_owned(),
        "https://example.invalid/vector.svg".to_owned(),
        "rIdSvg".to_owned(),
        TargetMode::External,
    )
    .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn missing_svg_target_package() -> Vec<u8> {
    let source = package_from_pair(true);
    let mut package = OpcPackage::from_bytes(&source).unwrap();
    package
        .get_part_mut(&PackURI::new(SLIDE).unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdSvg");
    PackageWriter::to_bytes(&package).unwrap()
}

fn duplicate_svg_owner_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_raster(false, true), |xml| {
        xml.replace(
            "</a:extLst>",
            &format!(
                r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdDuplicate1"/></a:ext><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdDuplicate2"/></a:ext></a:extLst>"#
            ),
        )
        .replace(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            &format!(
                "xmlns:r=\"{REL}\" xmlns:asvg=\"{SVG_NS}\""
            ),
        )
    })
}

fn foreign_extension_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_raster(false, true), |xml| {
        xml.replace(
            r#"<a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext>"#,
            r#"<x:ext xmlns:x="urn:litchi:foreign" uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></x:ext>"#,
        )
    })
}

fn mce_extension_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_raster(false, true), |xml| {
        xml.replace(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            &format!(
                r#"xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#
            ),
        )
        .replace(
            r#"<a:extLst><a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext></a:extLst>"#,
            r#"<mc:AlternateContent><mc:Choice Requires="a"><a:extLst><a:ext uri="{future-extension}"><!--future-comment--><future:unknown xmlns:future="urn:litchi:future" future:keep="yes"><![CDATA[<?future]]></future:unknown></a:ext></a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        )
    })
}

fn unrelated_transition_mce_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_raster(true, false), |xml| {
        xml.replace(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            &format!(
                r#"xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">"#
            ),
        )
        .replace(
            "<p:clrMapOvr/>",
            r#"<p:clrMapOvr/><mc:AlternateContent><mc:Choice Requires="p15"><p:transition><p:fade/></p:transition></mc:Choice><mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>"#,
        )
    })
}

fn wrong_namespace_package(prefix: &str, namespace: &str) -> Vec<u8> {
    rewrite_slide_member(&package_from_raster(true, false), |xml| {
        let expected = match prefix {
            "p" => PML,
            "a" => DRAWINGML,
            "r" => REL,
            _ => unreachable!("fixture only uses known PresentationML prefixes"),
        };
        let marker = format!(r#"xmlns:{prefix}="{expected}""#);
        xml.replace(&marker, &format!(r#"xmlns:{prefix}="{namespace}""#))
    })
}

fn shared_svg_package() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&package_from_pair(true)).unwrap();
    let slide2 = PackURI::new("/ppt/slides/slide2.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            slide2.clone(),
            ct::PML_SLIDE.to_owned(),
            paired_blip_slide_xml(true),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .set_blob(
            format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/><p:sldId id="257" r:id="rIdSlide2"/></p:sldIdLst></p:presentation>"#
            )
            .into_bytes(),
        );
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::SLIDE.to_owned(),
            "slides/slide2.xml".to_owned(),
            "rIdSlide2".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&slide2)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/raster.png".to_owned(),
            "rIdRaster".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&slide2)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/vector.svg".to_owned(),
            "rIdSvg".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn same_owner_shared_svg_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_pair(true), |xml| {
        let start = xml.find("<p:pic>").unwrap();
        let end = xml[start..].find("</p:pic>").unwrap() + start + "</p:pic>".len();
        let duplicate = xml[start..end].replace("id=\"42\"", "id=\"43\"");
        xml.replace("</p:spTree>", &format!("{duplicate}</p:spTree>"))
    })
}

fn same_owner_shared_svg_entity_reference_package() -> Vec<u8> {
    rewrite_slide_member(&package_from_pair(true), |xml| {
        let start = xml.find("<p:pic>").unwrap();
        let end = xml[start..].find("</p:pic>").unwrap() + start + "</p:pic>".len();
        let duplicate = xml[start..end]
            .replace("id=\"42\"", "id=\"43\"")
            .replace("r:embed=\"rIdSvg\"", "r:embed=\"rId&#83;vg\"");
        xml.replace("</p:spTree>", &format!("{duplicate}</p:spTree>"))
    })
}

fn malformed_shared_svg_owner_package() -> Vec<u8> {
    rewrite_slide_member(&same_owner_shared_svg_package(), |mut xml| {
        let marker = format!(r#"<a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>"#);
        let first = xml.find(&marker).unwrap();
        let second = first
            + marker.len()
            + xml[first + marker.len()..]
                .find(&marker)
                .expect("shared fixture must contain a second SVG owner");
        let replacement = format!(r#"<a:ext uri="{SVG_URI}"></a:ext>"#);
        xml.replace_range(second..second + marker.len(), &replacement);
        xml
    })
}

const NEW_SVG: &[u8] =
    br#"<svg xmlns="http://www.w3.org/2000/svg"><circle cx="4" cy="5" r="2"/></svg>"#;

#[test]
fn attach_save_reopen_detach_save_reopen_restores_source_and_opaque_members() {
    for (self_closing, label) in [(true, "self-closing"), (false, "paired")] {
        let source = package_from_raster(self_closing, !self_closing);
        let source_slide = slide_member(&source);
        let (attached, relationship_id, part_uri) = attach_and_publish(&source, NEW_SVG);
        let part_member = part_uri.trim_start_matches('/');

        assert_ne!(
            attached, source,
            "{label} attach must produce a changed package"
        );
        assert!(
            slide_member(&attached)
                .windows(SVG_URI.len())
                .any(|window| window == SVG_URI.as_bytes())
        );
        assert!(
            slide_member(&attached)
                .windows(b"asvg:svgBlip".len())
                .any(|window| window == b"asvg:svgBlip")
        );
        if !self_closing {
            assert!(
                slide_member(&attached)
                    .windows(b"future:keep=\"yes\"".len())
                    .any(|window| window == b"future:keep=\"yes\"")
            );
            assert!(
                slide_member(&attached)
                    .windows(b"<!--future-comment-->".len())
                    .any(|window| window == b"<!--future-comment-->")
            );
            assert!(
                slide_member(&attached)
                    .windows(b"<![CDATA[<?future]]>".len())
                    .any(|window| window == b"<![CDATA[<?future]]>")
            );
        }
        assert_eq!(physical_member(&attached, part_member), NEW_SVG);
        assert!(physical_member_names(&attached).contains(&part_member.to_owned()));
        assert!(
            String::from_utf8(slide_relationships_member(&attached))
                .unwrap()
                .contains(&format!("Id=\"{relationship_id}\""))
        );
        assert!(
            String::from_utf8(slide_relationships_member(&attached))
                .unwrap()
                .contains(&format!(
                    "Target=\"../media/{}\"",
                    part_uri.rsplit('/').next().unwrap()
                ))
        );
        assert_unrelated_members_unchanged(
            &source,
            &attached,
            &[
                "[Content_Types].xml",
                "ppt/slides/slide1.xml",
                "ppt/slides/_rels/slide1.xml.rels",
            ],
        );
        assert_eq!(
            physical_member(&attached, "ppt/media/raster.png"),
            b"old-raster-payload"
        );
        assert_eq!(
            physical_member(&attached, "ppt/media/opaque.bin"),
            b"untouched opaque member"
        );

        let reopened = SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(
            attached.clone(),
        )))
        .unwrap();
        let svg = reopened.slide(0).unwrap().read_svg_image(0).unwrap();
        assert_eq!(svg.bytes(), NEW_SVG);
        assert_eq!(svg.descriptor().relationship_id(), relationship_id);

        let detached = detach_and_publish(&attached);
        if self_closing {
            assert_unrelated_members_unchanged(
                &source,
                &detached,
                &["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"],
            );
        } else {
            assert_eq!(
                detached, source,
                "{label} attach inverse must restore source bytes"
            );
        }
        let reopened =
            SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(detached)))
                .unwrap();
        assert!(
            reopened.slide(0).unwrap().images().unwrap()[0]
                .svg()
                .is_none()
        );
        assert_eq!(slide_member(&source), source_slide);
    }
}

#[test]
fn bom_prefixed_slide_attach_publish_and_reopen_preserves_the_bom() {
    let source = package_from_bom_raster();
    assert!(slide_member(&source).starts_with(UTF8_BOM));
    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    assert!(slide_member(&attached).starts_with(UTF8_BOM));
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(attached.clone())))
            .unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );

    let detached = detach_and_publish(&attached);
    assert!(slide_member(&detached).starts_with(UTF8_BOM));
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(detached))).unwrap();
    assert!(
        reopened.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .is_none()
    );
}

#[test]
fn attach_expands_selfclosing_ext_lst_without_rebinding_or_dropping_source_metadata() {
    let source = package_from_raster_with_prefixed_selfclosing_ext_lst();
    let source_slide = slide_member(&source);
    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    let attached_slide = String::from_utf8(slide_member(&attached)).unwrap();
    assert!(
        attached_slide.contains(&format!(
            r#"<d:extLst xmlns:d="{DRAWINGML}" xmlns:a="urn:litchi:foreign" xmlns:x="urn:litchi:future" x:future="keep">"#
        )),
        "the self-closing extension-list opening tag must retain its local declarations and attributes"
    );
    assert!(attached_slide.contains("<d:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">"));
    assert!(
        !attached_slide.contains("<a:extLst")
            && !attached_slide.contains("<a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">",),
        "the generated extension must use the extLst owner prefix, not the parent blip prefix"
    );
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(attached.clone())))
            .unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );

    let detached = detach_and_publish(&attached);
    let detached_slide = String::from_utf8(slide_member(&detached)).unwrap();
    assert!(detached_slide.contains("<d:extLst"));
    assert!(detached_slide.contains("xmlns:a=\"urn:litchi:foreign\""));
    assert!(detached_slide.contains("xmlns:x=\"urn:litchi:future\" x:future=\"keep\""));
    assert!(!detached_slide.contains(SVG_URI));
    assert_unrelated_members_unchanged(
        &source,
        &detached,
        &["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"],
    );
    assert_eq!(
        physical_member(&detached, RASTER),
        physical_member(&source, RASTER)
    );
    assert!(
        source_slide
            .windows(b"<d:extLst".len())
            .any(|window| window == b"<d:extLst")
    );
}

#[test]
fn fresh_detach_keeps_namespace_only_ext_lst_wrapper() {
    let source = package_from_raster_with_namespace_only_selfclosing_ext_lst();
    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    let attached_slide = String::from_utf8(slide_member(&attached)).unwrap();
    assert!(attached_slide.contains(&format!(
        r#"<d:extLst xmlns:d="{DRAWINGML}" xmlns:a="urn:litchi:foreign" xmlns:x="urn:litchi:future">"#
    )));

    let detached = detach_and_publish(&attached);
    let detached_slide = String::from_utf8(slide_member(&detached)).unwrap();
    assert!(detached_slide.contains("<d:extLst"));
    assert!(detached_slide.contains("xmlns:d=\""));
    assert!(detached_slide.contains("xmlns:a=\"urn:litchi:foreign\""));
    assert!(detached_slide.contains("xmlns:x=\"urn:litchi:future\""));
    assert!(!detached_slide.contains(SVG_URI));
}

#[test]
fn opaque_ext_pi_and_mce_payload_survive_attach_and_detach() {
    let source = package_from_raster_with_opaque_pi_and_mce();
    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    let attached_slide = String::from_utf8(slide_member(&attached)).unwrap();
    assert!(attached_slide.contains("<?future data?>"));
    assert!(attached_slide.contains("<mc:AlternateContent"));
    assert!(attached_slide.contains("<future:opaque/>"));

    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(attached.clone())))
            .unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );
    let detached = detach_and_publish(&attached);
    let detached_slide = String::from_utf8(slide_member(&detached)).unwrap();
    assert!(detached_slide.contains("<?future data?>"));
    assert!(detached_slide.contains("<mc:AlternateContent"));
    assert!(detached_slide.contains("<future:opaque/>"));
    assert!(!detached_slide.contains(SVG_URI));

    let typed_mce = mce_extension_package();
    assert!(
        open_editor(&typed_mce).edit_svg_attachment(0, 0).is_err(),
        "MCE ancestry around the typed owner remains refused"
    );
}

#[test]
fn standard_namespace_picture_lookalikes_inside_unknown_ext_stay_opaque() {
    let source = package_from_raster_with_standard_namespace_lookalike();
    let lookalike = format!(
        r#"<p:pic xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:blipFill><a:blip r:embed="rIdFuture"/></p:blipFill></p:pic>"#
    );
    assert!(
        String::from_utf8(slide_member(&source))
            .unwrap()
            .contains(&lookalike)
    );

    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    let attached_slide = String::from_utf8(slide_member(&attached)).unwrap();
    assert!(attached_slide.contains(&lookalike));
    assert_eq!(
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(attached.clone())))
            .unwrap()
            .slide(0)
            .unwrap()
            .images()
            .unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );

    let detached = detach_and_publish(&attached);
    let detached_slide = String::from_utf8(slide_member(&detached)).unwrap();
    assert!(detached_slide.contains(&lookalike));
    assert!(!detached_slide.contains(SVG_URI));
}

#[test]
fn paired_svg_detach_preserves_unknown_extension_and_shared_svg_part() {
    let source = shared_svg_package();
    let source_slide = slide_member(&source);
    let source_svg = physical_member(&source, "ppt/media/vector.svg");
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
    assert_eq!(edit.source().svg().unwrap().part_uri().as_str(), SVG);
    assert!(edit.detach().unwrap());
    let commit = edit.commit_checked().unwrap();
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert!(restored.svg().is_some());
    assert_eq!(restored.svg().unwrap().bytes(), source_svg.as_slice());

    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(
        !slide_member(&output)
            .windows(SVG_URI.len())
            .any(|window| window == SVG_URI.as_bytes())
    );
    assert!(
        slide_member(&output)
            .windows(b"future:keep=\"yes\"".len())
            .any(|window| window == b"future:keep=\"yes\"")
    );
    assert!(
        slide_member(&output)
            .windows(b"<!--future-comment-->".len())
            .any(|window| window == b"<!--future-comment-->")
    );
    assert!(
        slide_member(&output)
            .windows(b"<![CDATA[<?future]]>".len())
            .any(|window| window == b"<![CDATA[<?future]]>")
    );
    assert_unrelated_members_unchanged(
        &source,
        &output,
        &["ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"],
    );
    assert_eq!(physical_member(&output, "ppt/media/vector.svg"), source_svg);
    assert_ne!(slide_member(&output), source_slide);
    assert!(
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output)))
            .unwrap()
            .slide(0)
            .unwrap()
            .images()
            .unwrap()[0]
            .svg()
            .is_none()
    );
}

#[test]
fn detach_shared_svg_relationship_keeps_the_other_picture_owner() {
    let source = same_owner_shared_svg_package();
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
    edit.detach().unwrap();
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    let slide = slide_member(&output);
    assert_eq!(
        slide
            .windows(SVG_URI.len())
            .filter(|w| *w == SVG_URI.as_bytes())
            .count(),
        1
    );
    assert_eq!(
        physical_member(&output, "ppt/media/vector.svg"),
        br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>"#
    );
    let relationships = String::from_utf8(slide_relationships_member(&output)).unwrap();
    assert!(relationships.contains("Id=\"rIdSvg\""));
}

#[test]
fn detach_entity_encoded_shared_svg_relationship_keeps_media_and_other_owner() {
    let source = same_owner_shared_svg_entity_reference_package();
    let source_svg = physical_member(&source, "ppt/media/vector.svg");
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
    assert!(edit.detach().unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();

    let relationships = String::from_utf8(slide_relationships_member(&output)).unwrap();
    assert!(relationships.contains("Id=\"rIdSvg\""));
    assert_eq!(physical_member(&output, "ppt/media/vector.svg"), source_svg);
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    let images = reopened.slide(0).unwrap().images().unwrap();
    assert!(images[0].svg().is_none());
    assert_eq!(images[1].svg().unwrap().relationship_id(), "rIdSvg");
}

#[test]
fn malformed_shared_svg_detach_cannot_commit() {
    let source = malformed_shared_svg_owner_package();
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
    assert!(edit.detach().unwrap());
    assert!(matches!(
        edit.commit(),
        Err(Error::Invalid(_)) | Err(Error::Relationship(_)) | Err(Error::UnsafeEdit { .. })
    ));
}

#[test]
fn attach_rejects_existing_svg_owner_shared_id_and_unsafe_targets_before_publication() {
    let duplicate = duplicate_svg_owner_package();
    assert!(matches!(
        open_editor(&duplicate).edit_svg_attachment(0, 0),
        Err(Error::Relationship(_)) | Err(Error::Invalid(_))
    ));

    let missing = missing_svg_target_package();
    assert!(matches!(
        open_editor(&missing).edit_svg_attachment(0, 0),
        Err(Error::Relationship(_))
    ));

    let external = external_svg_package();
    assert!(matches!(
        open_editor(&external).edit_svg_attachment(0, 0),
        Err(Error::Relationship(_))
    ));

    let foreign = foreign_extension_package();
    assert!(matches!(
        open_editor(&foreign).edit_svg_attachment(0, 0),
        Err(Error::Invalid(_)) | Err(Error::Relationship(_))
    ));

    let mce = mce_extension_package();
    assert!(matches!(
        open_editor(&mce).edit_svg_attachment(0, 0),
        Err(Error::Invalid(_)) | Err(Error::Relationship(_)) | Err(Error::UnsafeEdit { .. })
    ));

    let transition_mce = unrelated_transition_mce_package();
    let editor = open_editor(&transition_mce);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(edit.attach_svg(NEW_SVG).unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(
        slide_member(&output)
            .windows(b"mc:AlternateContent".len())
            .any(|window| window == b"mc:AlternateContent")
    );

    let source = package_from_raster(true, false);
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    let collision =
        SourceSvgAttachmentReplacement::new(NEW_SVG.to_vec()).with_relationship_id("rIdRaster");
    assert!(matches!(
        edit.attach(&collision),
        Err(Error::Relationship(_)) | Err(Error::Invalid(_))
    ));
    assert_eq!(collision.svg(), NEW_SVG);
    assert!(edit.source().svg().is_none());
    assert!(edit.attach_svg(NEW_SVG).unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(!output.is_empty());

    let occupied = source_package(raster_slide_xml(true, false), true, false);
    let editor = open_editor(&occupied);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    let collision = SourceSvgAttachmentReplacement::new(NEW_SVG.to_vec())
        .with_part_uri(PackURI::new(SVG).unwrap());
    assert!(matches!(
        edit.attach(&collision),
        Err(Error::Relationship(_)) | Err(Error::Invalid(_))
    ));
    assert_eq!(collision.svg(), NEW_SVG);
    assert!(edit.source().svg().is_none());
    assert!(edit.attach_svg(NEW_SVG).is_ok());

    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(matches!(edit.attach_svg(&[]), Err(Error::Invalid(_))));
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(matches!(
        edit.attach(
            &SourceSvgAttachmentReplacement::new(NEW_SVG.to_vec())
                .with_part_uri(PackURI::new("/ppt/other/vector.svg").unwrap(),)
        ),
        Err(Error::Relationship(_)) | Err(Error::Invalid(_))
    ));

    for (prefix, namespace) in [
        ("p", "urn:litchi:wrong-presentation"),
        ("a", "urn:litchi:wrong-drawing"),
        ("r", "urn:litchi:wrong-relationships"),
    ] {
        let malformed = wrong_namespace_package(prefix, namespace);
        let operation =
            SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(malformed)))
                .and_then(|editor| editor.edit_svg_attachment(0, 0).map(|_| ()));
        assert!(
            operation.is_err(),
            "wrong {prefix} namespace must fail closed"
        );
    }

    let wrong_asvg = rewrite_slide_member(&package_from_pair(true), |xml| {
        xml.replace(
            &format!(r#"xmlns:asvg="{SVG_NS}""#),
            r#"xmlns:asvg="urn:litchi:wrong-svg""#,
        )
    });
    let operation =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(wrong_asvg)))
            .and_then(|editor| editor.edit_svg_attachment(0, 0).map(|_| ()));
    assert!(operation.is_err(), "wrong asvg namespace must fail closed");
}

#[test]
fn stale_and_signed_lifecycle_publication_refuse_before_writer_bytes() {
    let source = package_from_raster(true, false);
    let versioned = Arc::new(VersionedSource::new(source.clone()));
    let editor = SourceBackedPresentationEditor::from_read_at(versioned.clone()).unwrap();
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    edit.attach_svg(NEW_SVG).unwrap();
    let commit = edit.commit_checked().unwrap();
    versioned.changed();
    let mut output = Vec::new();
    let result = editor.publish_svg_attachment_commit_to_stream(&mut output, &commit);
    assert!(matches!(
        result,
        Err(Error::StaleSource) | Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let signed = source_package(raster_slide_xml(true, false), false, true);
    let editor = open_editor(&signed);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    edit.attach_svg(NEW_SVG).unwrap();
    let output: Vec<u8> = Vec::new();
    assert!(matches!(
        edit.commit(),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn attachment_limits_are_checked_before_lifecycle_publication() {
    let source = package_from_raster(true, false);
    let limits = ReadLimits::builder()
        .max_part_bytes(4096)
        .unwrap()
        .build()
        .unwrap();
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(VersionedSource::new(source.clone())),
        limits,
    )
    .unwrap();
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    let oversized = vec![b'X'; 4097];
    assert!(matches!(
        edit.attach_svg(&oversized),
        Err(Error::Limit { .. })
    ));
    assert_eq!(oversized.len(), 4097);
    assert!(oversized.iter().all(|byte| *byte == b'X'));
    assert!(edit.source().svg().is_none());
    assert!(edit.attach_svg(NEW_SVG).unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_ne!(output, source);
    assert_eq!(
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output)))
            .unwrap()
            .slide(0)
            .unwrap()
            .images()
            .unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );
}

#[test]
fn inherited_namespace_context_respects_source_part_limit_and_is_reusable() {
    let source = package_from_raster_with_large_inherited_namespace();
    let source_slide = slide_member(&source);
    // The inherited declarations enlarge the detached picture copy, while
    // the complete copy can still be smaller than the whole slide because
    // the outer slide wrappers are omitted. Keep the accepted budget just
    // above the physical source member and exercise the public source limit
    // at the materialization boundary below it.
    let source_limit = ReadLimits::builder()
        .max_part_bytes(u64::try_from(source_slide.len() + 1).unwrap())
        .unwrap()
        .build()
        .unwrap();

    let source_for_presentation = Arc::new(VersionedSource::new(source.clone()));
    let presentation =
        SourceBackedPresentation::from_read_at_with_limits(source_for_presentation, source_limit)
            .unwrap();
    let slide = presentation.slide(0).unwrap();
    assert!(slide.images().is_ok());

    let source_for_editor = Arc::new(VersionedSource::new(source.clone()));
    let editor =
        SourceBackedPresentationEditor::from_read_at_with_limits(source_for_editor, source_limit)
            .unwrap();
    assert!(editor.edit_svg_attachment(0, 0).is_ok());

    let tight_limit = ReadLimits::builder()
        .max_part_bytes(u64::try_from(source_slide.len() - 1).unwrap())
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPresentation::from_read_at_with_limits(
            Arc::new(VersionedSource::new(source.clone())),
            tight_limit,
        ),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));
    assert!(matches!(
        SourceBackedPresentationEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(source.clone())),
            tight_limit,
        ),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));
    assert_eq!(slide_member(&source), source_slide);
}

#[test]
fn selfclosing_blip_limit_rejects_the_real_expanded_slide_size() {
    let source = package_from_raster(true, false);
    let (attached, _, _) = attach_and_publish(&source, NEW_SVG);
    let attached_slide_len = slide_member(&attached).len();
    assert!(attached_slide_len > slide_member(&source).len());
    let limits = ReadLimits::builder()
        .max_part_bytes(u64::try_from(attached_slide_len - 1).unwrap())
        .unwrap()
        .build()
        .unwrap();
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(VersionedSource::new(source.clone())),
        limits,
    )
    .unwrap();
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(matches!(edit.attach_svg(NEW_SVG), Err(Error::Limit { .. })));
    assert!(edit.source().svg().is_none());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
}

#[test]
fn svg_uri_profile_collapses_xml_whitespace_but_not_nbsp() {
    let xml_spaced = format!("  {SVG_URI}\t\n ");
    let editor = open_editor(&package_from_pair_with_uri(true, &xml_spaced));
    let edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(
        edit.source().svg().is_some(),
        "xsd:token XML whitespace must classify the native SVG extension"
    );

    let nbsp_spaced = format!("\u{00a0}{SVG_URI}\u{00a0}");
    let editor = open_editor(&package_from_pair_with_uri(true, &nbsp_spaced));
    let edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(
        edit.source().svg().is_none(),
        "NBSP is not XML whitespace and must retain the extension as opaque"
    );
}

#[test]
fn raster_fixture_has_no_svg_owner_and_preserves_physical_catalog() {
    let source = package_from_raster(false, true);
    let reader = PhysPkgReader::new(&source).unwrap();
    assert!(
        reader
            .member_names()
            .unwrap()
            .iter()
            .any(|name| name == "ppt/slides/slide1.xml")
    );
    assert!(
        reader
            .member_names()
            .unwrap()
            .iter()
            .any(|name| name == "ppt/media/opaque.bin")
    );
    let presentation =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(source))).unwrap();
    assert_eq!(presentation.slide(0).unwrap().images().unwrap().len(), 1);
    assert!(
        presentation.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .is_none()
    );
}

#[test]
fn native_hidden_graphic_shared_svg_accepts_default_undeclaration_and_detaches_one_owner() {
    let source = native_fixture("tdf169496_hidden_graphic.pptx");

    let presentation =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .unwrap();
    let images = presentation.slide(0).unwrap().images().unwrap();
    assert_eq!(images.len(), 2);
    assert_eq!(
        images[0].svg().unwrap().relationship_id(),
        "rId3",
        "the hidden picture must retain the shared native SVG relationship"
    );
    assert_eq!(images[1].svg().unwrap().relationship_id(), "rId3");

    let source_svg = physical_member(&source, "ppt/media/image2.svg");
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rId3");
    assert!(edit.detach().unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();

    let output_slide = physical_member(&output, "ppt/slides/slide1.xml");
    assert_eq!(
        output_slide
            .windows(SVG_URI.len())
            .filter(|window| *window == SVG_URI.as_bytes())
            .count(),
        1,
        "the other picture must retain the native SVG extension"
    );
    let output_relationships = String::from_utf8(slide_relationships_member(&output)).unwrap();
    assert!(output_relationships.contains(r#"Id="rId3""#));
    assert!(output_relationships.contains(r#"Target="../media/image2.svg""#));
    assert_eq!(physical_member(&output, "ppt/media/image2.svg"), source_svg);

    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output.clone())))
            .unwrap();
    let images = reopened.slide(0).unwrap().images().unwrap();
    assert_eq!(images.len(), 2);
    assert!(images[0].svg().is_none());
    assert_eq!(images[1].svg().unwrap().relationship_id(), "rId3");

    let reopened_editor = open_editor(&output);
    let second = reopened_editor.edit_svg_attachment(0, 1).unwrap();
    assert_eq!(second.source().svg().unwrap().relationship_id(), "rId3");
}

#[test]
fn native_corpus_records_strict_refusals_and_full_hashes() {
    for (name, expected_hash, expected_pictures) in [
        (
            "tdf163852.pptx",
            "2525665a1e430a9ee543faf456424bd7b994cce0a2d76470a5a07453e3e2ab39",
            1,
        ),
        (
            "tdf164622.pptx",
            "33b86388804554e35cb3b07e81cf5db5cb6bb02e76716c71f233b57325c2d199",
            1,
        ),
    ] {
        let source = native_fixture(name);
        assert_eq!(sha256_hex(&source), expected_hash, "{name} fixture changed");
        assert_eq!(direct_picture_count(&source), expected_pictures);
        let editor = open_editor(&source);
        let result = editor.edit_svg_attachment(0, 0);
        assert!(
            matches!(result, Err(Error::Invalid(message)) if message.contains("fillRect")),
            "{name} must be refused for its malformed stretch grammar"
        );
    }
}

#[test]
fn custom_inherited_namespaces_default_drawingml_and_strict_dialect_round_trip() {
    for (label, source) in [
        (
            "custom-prefixes",
            package_from_namespace_variant(false, false),
        ),
        (
            "default-drawingml",
            package_from_namespace_variant(true, false),
        ),
        (
            "strict-custom-prefixes",
            package_from_namespace_variant(false, true),
        ),
    ] {
        let presentation =
            SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(source.clone())))
                .unwrap_or_else(|error| panic!("{label} source open failed: {error}"));
        let images = presentation
            .slide(0)
            .unwrap()
            .images()
            .unwrap_or_else(|error| panic!("{label} image inventory failed: {error}"));
        assert_eq!(images.len(), 1, "{label} picture count");
        assert_eq!(
            images[0].svg().unwrap().relationship_id(),
            "rIdSvg",
            "{label} SVG relationship"
        );

        let editor = open_editor(&source);
        let mut edit = editor
            .edit_svg_attachment(0, 0)
            .unwrap_or_else(|error| panic!("{label} lifecycle selection failed: {error}"));
        assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
        assert!(edit.detach().unwrap());
        let commit = edit.commit_checked().unwrap();
        let mut output = Vec::new();
        editor
            .publish_svg_attachment_commit_to_stream(&mut output, &commit)
            .unwrap_or_else(|error| panic!("{label} lifecycle publication failed: {error}"));
        assert!(
            !slide_member(&output)
                .windows(SVG_URI.len())
                .any(|window| window == SVG_URI.as_bytes()),
            "{label} must remove the selected SVG owner"
        );
        let reopened =
            SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output)))
                .unwrap_or_else(|error| panic!("{label} readback failed: {error}"));
        assert!(
            reopened.slide(0).unwrap().images().unwrap()[0]
                .svg()
                .is_none(),
            "{label} readback must retain the raster picture"
        );
    }
}

#[test]
fn strict_host_attach_uses_transitional_svg_attribute_namespace_and_strict_core_relationships() {
    let source = package_from_strict_raster();
    let source_slide = slide_member(&source);
    let source_relationships = slide_relationships_member(&source);
    assert!(
        source_slide
            .windows(STRICT_REL.len())
            .any(|window| { window == STRICT_REL.as_bytes() })
    );
    let source_relationships_text = String::from_utf8(source_relationships.clone()).unwrap();
    assert!(source_relationships_text.contains(&format!(
        r#"Type="{STRICT_IMAGE}" Target="../media/raster.png""#
    )));

    let presentation =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(source.clone())))
            .unwrap();
    assert!(
        presentation.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .is_none()
    );

    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().raster_relationship_type(), rt::STRICT_IMAGE);
    assert!(edit.attach_svg(NEW_SVG).unwrap());
    let commit = edit.commit_checked().unwrap();
    assert_eq!(
        commit.snapshot().svg().unwrap().relationship_type(),
        rt::STRICT_IMAGE,
        "the physical SVG relationship keeps the host package's Strict image type"
    );
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert!(restored.svg().is_none());

    let mut attached = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut attached, &commit)
        .unwrap();
    let attached_slide = String::from_utf8(slide_member(&attached)).unwrap();
    assert!(attached_slide.contains(&format!(
        r#"<asvg:svgBlip xmlns:asvg="{SVG_NS}" xmlns:r="{REL}" r:embed="rIdSvg"/>"#
    )));
    assert!(attached_slide.contains(&format!(r#"xmlns:r="{STRICT_REL}""#)));
    assert!(attached_slide.contains(r#"r:embed="rIdRaster""#));
    let attached_relationships = String::from_utf8(slide_relationships_member(&attached)).unwrap();
    assert!(attached_relationships.contains(&format!(
        r#"Type="{STRICT_IMAGE}" Target="../media/raster.png""#
    )));
    assert!(attached_relationships.contains(&format!(
        r#"Type="{STRICT_IMAGE}" Target="../media/vector.svg""#
    )));
    assert!(!attached_relationships.contains(&format!(
        r#"Type="{REL}/image" Target="../media/vector.svg""#
    )));

    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(attached.clone())))
            .unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().images().unwrap()[0]
            .svg()
            .unwrap()
            .relationship_id(),
        "rIdSvg"
    );

    let reopened_editor = open_editor(&attached);
    let mut detach = reopened_editor.edit_svg_attachment(0, 0).unwrap();
    assert!(detach.detach().unwrap());
    let detach_commit = detach.commit_checked().unwrap();
    let detach_inverse = detach_commit.patch().inverse();
    let restored_attached = detach_inverse.apply(detach_commit.snapshot()).unwrap();
    assert!(restored_attached.svg().is_some());
    let mut detached = Vec::new();
    reopened_editor
        .publish_svg_attachment_commit_to_stream(&mut detached, &detach_commit)
        .unwrap();
    assert_eq!(slide_member(&detached), source_slide);
    assert_eq!(slide_relationships_member(&detached), source_relationships);
    assert_eq!(
        physical_member(&detached, RASTER),
        physical_member(&source, RASTER)
    );
    assert_eq!(
        physical_member(&detached, OPAQUE),
        physical_member(&source, OPAQUE)
    );
    assert_eq!(
        physical_member_names(&detached),
        physical_member_names(&source)
    );
}

#[test]
fn foreign_nested_picture_does_not_change_direct_picture_selection() {
    let source = foreign_nested_picture_package();
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(edit.source().svg().unwrap().relationship_id(), "rIdSvg");
    assert!(edit.detach().unwrap());
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut output, &commit)
        .unwrap();
    let slide = slide_member(&output);
    assert!(
        !slide
            .windows(SVG_URI.len())
            .any(|window| window == SVG_URI.as_bytes())
    );
    assert!(
        slide
            .windows(
                br#"<x:foreign xmlns:x="urn:litchi:foreign"><p:pic xmlns:p="urn:litchi:foreign"/></x:foreign>"#
                    .len()
            )
            .any(|window| {
                window
                    == br#"<x:foreign xmlns:x="urn:litchi:foreign"><p:pic xmlns:p="urn:litchi:foreign"/></x:foreign>"#
            })
    );
}

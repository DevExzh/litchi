use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits, TargetMode};
use litchi_pptx::{
    Error, SourceBackedPresentation, SourceBackedPresentationEditor, SourceSvgReplacement,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const RASTER: &str = "/ppt/media/raster.png";
const SVG: &str = "/ppt/media/vector.svg";

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
            717,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn slide_xml() -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:asvg="{SVG_NS}" xmlns:future="urn:litchi:future"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="42" name="SVG photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{SVG_URI}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext><a:ext uri="{{future-extension}}"><future:unknown future:keep="yes"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
    .into_bytes()
}

fn source_package(signed: bool) -> Vec<u8> {
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
            slide_xml(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(RASTER).unwrap(),
            "image/png".to_owned(),
            b"old-raster-payload".to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SVG).unwrap(),
            "image/svg+xml".to_owned(),
            br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>"#.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/media/opaque.bin").unwrap(),
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

fn package_with_orphan_relationship_member() -> Vec<u8> {
    let source = source_package(false);
    let archive = ArchiveReader::new(&source).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let data = archive.read(name).unwrap();
        writer.write_stored(name, &data).unwrap();
    }
    let orphan = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rOrphan" Type="urn:litchi:orphan" Target="../media/raster.png"/></Relationships>"#;
    writer
        .write_stored("ppt/slides/_rels/orphan.xml.rels", orphan)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn package_with_unselected_dangling_picture() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&source_package(false)).unwrap();
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let slide = package.get_part_mut(&slide_uri).unwrap();
    let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
    let picture_start = xml.find("<p:pic>").unwrap();
    let picture_end = xml.find("</p:pic>").unwrap() + "</p:pic>".len();
    let mut duplicate = xml[picture_start..picture_end].to_owned();
    duplicate = duplicate.replace("rIdRaster", "rIdMissing");
    duplicate = duplicate.replace("rIdSvg", "rIdSvgMissing");
    let xml = xml.replace("</p:spTree>", &format!("{duplicate}</p:spTree>"));
    slide.set_blob(xml.into_bytes());
    PackageWriter::to_bytes(&package).unwrap()
}

fn open_editor(bytes: &[u8]) -> SourceBackedPresentationEditor {
    SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(bytes.to_vec())))
        .unwrap()
}

fn external_svg_package() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&source_package(false)).unwrap();
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let slide = package.get_part_mut(&slide_uri).unwrap();
    let xml = String::from_utf8(slide.blob().to_vec())
        .unwrap()
        .replace("r:embed=\"rIdSvg\"", "r:link=\"rIdSvg\"")
        .into_bytes();
    slide.set_blob(xml);
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

fn non_png_raster_package() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&source_package(false)).unwrap();
    package
        .get_part_mut(&PackURI::new(RASTER).unwrap())
        .unwrap()
        .set_content_type("image/jpeg".to_owned())
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

#[test]
fn svg_transaction_retargets_media_and_preserves_opaque_slide_bytes() {
    let source = source_package(false);
    let source_package = OpcPackage::from_bytes(&source).unwrap();
    let source_slide = source_package
        .get_part(&PackURI::new(SLIDE).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    assert_eq!(edit.source().svg_relationship_id(), "rIdSvg");
    assert_eq!(edit.source().raster_relationship_id(), "rIdRaster");
    let replacement = SourceSvgReplacement::new(
        br#"<svg xmlns="http://www.w3.org/2000/svg"><circle cx="2" cy="3" r="1"/></svg>"#.to_vec(),
        b"new-raster-payload".to_vec(),
    )
    .with_part_uris(
        PackURI::new("/ppt/media/replaced.dat").unwrap(),
        PackURI::new("/ppt/media/replaced.png").unwrap(),
    );
    assert!(edit.replace(replacement).unwrap());
    let commit = edit.commit_checked().unwrap();
    assert!(commit.is_changed());
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored.svg_bytes(),
        br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>"#
    );
    assert_eq!(restored.raster_bytes(), b"old-raster-payload");
    assert_eq!(restored.svg_part_uri().as_str(), SVG);
    assert_eq!(restored.raster_part_uri().as_str(), RASTER);

    let mut output = Vec::new();
    editor
        .publish_svg_commit_to_stream(&mut output, &commit)
        .unwrap();
    let package = OpcPackage::from_bytes(&output).unwrap();
    let slide = package.get_part(&PackURI::new(SLIDE).unwrap()).unwrap();
    assert_eq!(slide.blob(), source_slide.as_slice());
    assert_eq!(
        slide.rels().get("rIdSvg").unwrap().target_ref(),
        "../media/replaced.dat"
    );
    assert_eq!(
        slide.rels().get("rIdRaster").unwrap().target_ref(),
        "../media/replaced.png"
    );
    assert_eq!(
        slide.rels().get("rIdFuture").unwrap().reltype(),
        "urn:litchi:future-image-metadata"
    );
    assert_eq!(
        slide.rels().get("rIdFuture").unwrap().target_ref(),
        "../media/opaque.bin"
    );
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/media/replaced.dat").unwrap())
            .unwrap()
            .blob(),
        br#"<svg xmlns="http://www.w3.org/2000/svg"><circle cx="2" cy="3" r="1"/></svg>"#
    );
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/media/replaced.png").unwrap())
            .unwrap()
            .blob(),
        b"new-raster-payload"
    );
    assert!(package.get_part(&PackURI::new(SVG).unwrap()).is_err());
    assert!(package.get_part(&PackURI::new(RASTER).unwrap()).is_err());
    assert_eq!(
        package
            .get_part(&PackURI::new("/ppt/media/opaque.bin").unwrap())
            .unwrap()
            .blob(),
        b"untouched opaque member"
    );

    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    let slide = reopened.slide(0).unwrap();
    assert_eq!(
        slide.read_svg_image(0).unwrap().bytes(),
        br#"<svg xmlns="http://www.w3.org/2000/svg"><circle cx="2" cy="3" r="1"/></svg>"#
    );
}

#[test]
fn svg_transaction_noop_inverse_stale_and_signed_are_atomic() {
    let source = source_package(false);
    let editor = open_editor(&source);
    let original_svg = editor.edit_svg_image(0, 0).unwrap();
    let replacement = SourceSvgReplacement::new(
        original_svg.source().svg_bytes().to_vec(),
        original_svg.source().raster_bytes().to_vec(),
    );
    let mut edit = original_svg;
    assert!(!edit.replace(replacement).unwrap());
    let noop = edit.commit_checked().unwrap();
    assert!(!noop.is_changed());
    let mut exact = Vec::new();
    editor
        .publish_svg_commit_to_stream(&mut exact, &noop)
        .unwrap();
    assert_eq!(exact, source);

    let versioned = Arc::new(VersionedSource::new(source.clone()));
    let editor = SourceBackedPresentationEditor::from_read_at(versioned.clone()).unwrap();
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    edit.replace(SourceSvgReplacement::new(
        b"<svg xmlns=\"http://www.w3.org/2000/svg\"><rect/></svg>".to_vec(),
        b"changed-raster".to_vec(),
    ))
    .unwrap();
    let commit = edit.commit_checked().unwrap();
    let inverse = commit.patch().inverse();
    assert!(inverse.apply(commit.snapshot()).is_ok());
    versioned.changed();
    let mut stale_output = Vec::new();
    assert!(matches!(
        editor.publish_svg_commit_to_stream(&mut stale_output, &commit),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(stale_output.is_empty());

    let signed = source_package(true);
    let editor = open_editor(&signed);
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    edit.replace(SourceSvgReplacement::new(
        b"<svg xmlns=\"http://www.w3.org/2000/svg\"><rect/></svg>".to_vec(),
        b"signed-raster".to_vec(),
    ))
    .unwrap();
    let commit = edit.commit_checked().unwrap();
    let mut signed_output = Vec::new();
    assert!(matches!(
        editor.publish_svg_commit_to_stream(&mut signed_output, &commit),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(signed_output.is_empty());
}

#[test]
fn svg_transaction_refuses_shared_media_before_output() {
    let source = source_package(false);
    let mut package = OpcPackage::from_bytes(&source).unwrap();
    let second = PackURI::new("/ppt/slides/slide2.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            second.clone(),
            ct::PML_SLIDE.to_owned(),
            slide_xml(),
        )))
        .unwrap();
    package
        .get_part_mut(&second)
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
        .get_part_mut(&second)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/vector.svg".to_owned(),
            "rIdSvg".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let shared = PackageWriter::to_bytes(&package).unwrap();
    let editor = open_editor(&shared);
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    edit.replace(SourceSvgReplacement::new(
        b"<svg xmlns=\"http://www.w3.org/2000/svg\"><rect/></svg>".to_vec(),
        b"new-raster".to_vec(),
    ))
    .unwrap();
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_svg_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::InvalidRelationship(_))) | Err(Error::Relationship(_))
    ));
    assert!(output.is_empty());
}

#[test]
fn svg_transaction_refuses_linked_svg_before_edit_creation() {
    let editor = open_editor(&external_svg_package());
    assert!(matches!(
        editor.edit_svg_image(0, 0),
        Err(Error::Relationship(_))
    ));
}

#[test]
fn svg_transaction_requires_the_specified_png_raster_fallback() {
    let editor = open_editor(&non_png_raster_package());
    assert!(matches!(
        editor.edit_svg_image(0, 0),
        Err(Error::ContentType { expected, .. }) if expected == "image/png"
    ));
}

#[test]
fn svg_transaction_checks_replacement_limits_before_publication() {
    let source = source_package(false);
    let max_part_bytes = 2048_usize;
    let limits = ReadLimits::builder()
        .max_part_bytes(max_part_bytes as u64)
        .unwrap()
        .build()
        .unwrap();
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(VersionedSource::new(source.clone())),
        limits,
    )
    .unwrap();
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    let mut oversized = vec![b'X'; max_part_bytes + 1];
    oversized[..4].copy_from_slice(b"<svg");
    let error = edit
        .replace(SourceSvgReplacement::new(oversized, b"new-raster".to_vec()))
        .unwrap_err();
    assert!(matches!(error, Error::Limit { .. }));
    let mut output = Vec::new();
    let commit = edit.commit_checked().unwrap();
    editor
        .publish_svg_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
}

#[test]
fn svg_transaction_refuses_retained_orphan_relationship_members_before_output() {
    let source = package_with_orphan_relationship_member();
    let editor = open_editor(&source);
    let mut edit = editor.edit_svg_image(0, 0).unwrap();
    edit.replace(
        SourceSvgReplacement::new(
            b"<svg xmlns=\"http://www.w3.org/2000/svg\"><rect/></svg>".to_vec(),
            b"changed-raster".to_vec(),
        )
        .with_part_uris(
            PackURI::new("/ppt/media/replaced.svg").unwrap(),
            PackURI::new("/ppt/media/replaced.png").unwrap(),
        ),
    )
    .unwrap();
    let commit = edit.commit_checked().unwrap();
    let mut output = Vec::new();
    let result = editor.publish_svg_commit_to_stream(&mut output, &commit);
    assert!(
        matches!(result, Err(Error::Relationship(_))),
        "opaque orphan relationship members must fail closed before media removal"
    );
    assert!(output.is_empty());
}

#[test]
fn svg_transaction_validates_every_picture_owner_before_edit_creation() {
    let editor = open_editor(&package_with_unselected_dangling_picture());
    let result = editor.edit_svg_image(0, 0);
    assert!(
        matches!(result, Err(Error::Relationship(_))),
        "an unselected picture with unresolved relationships must fail closed"
    );
}

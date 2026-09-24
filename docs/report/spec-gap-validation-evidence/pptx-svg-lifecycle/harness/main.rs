use std::path::{Path, PathBuf};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_pptx::{SourceBackedPresentation, SourceBackedPresentationEditor};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const SVG: &[u8] = br#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1"/></svg>"#;

fn source() -> Vec<u8> {
    let mut package = OpcPackage::new();
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let slide = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}" xmlns:future="urn:litchi:future"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="2" name="Raster picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed='rIdRaster'><a:extLst><!-- retained before --><a:ext uri="urn:future"><future:payload><![CDATA[<?future]]></future:payload></a:ext><!-- retained after --></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>
<!-- retained trailing -->"#
    );
    // Valid 1x1 PNG fixture. SVG conversion/rendering is not part of the API.
    let png: Vec<u8> = [
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 4,
        0, 0, 0, 181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84, 120, 156, 99, 96, 248, 15, 0, 1, 2,
        1, 0, 66, 190, 188, 104, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ]
    .to_vec();
    for (name, kind, bytes) in [
        (
            "/ppt/presentation.xml",
            ct::PML_PRESENTATION_MAIN,
            presentation.into_bytes(),
        ),
        (SLIDE, ct::PML_SLIDE, b"<sld/>".to_vec()),
        ("/ppt/media/raster.png", "image/png", png),
        (
            "/ppt/media/opaque.bin",
            "application/octet-stream",
            b"unchanged opaque member".to_vec(),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                kind.to_owned(),
                bytes,
            )))
            .unwrap();
    }
    for (owner, kind, target, id) in [
        (
            "/ppt/presentation.xml",
            rt::SLIDE,
            "slides/slide1.xml",
            "rIdSlide",
        ),
        (SLIDE, rt::IMAGE, "../media/raster.png", "rIdRaster"),
    ] {
        package
            .get_part_mut(&PackURI::new(owner).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                kind.to_owned(),
                target.to_owned(),
                id.to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    let generated = PackageWriter::to_bytes(&package).unwrap();
    let reader = PhysPkgReader::new(&generated).unwrap();
    let mut writer = PhysPkgWriter::new();
    for name in reader.member_names().unwrap() {
        let uri = PackURI::new(format!("/{name}")).unwrap();
        let bytes = if uri.as_str() == SLIDE {
            slide.as_bytes().to_vec()
        } else {
            reader.read_member(&name).unwrap()
        };
        writer.write(&uri, &bytes).unwrap();
    }
    writer.finish().unwrap()
}

fn emit(output: &Path, name: &str, bytes: &[u8]) {
    std::fs::write(output.join(format!("{name}.pptx")), bytes).unwrap();
    let package = OpcPackage::from_bytes(bytes).unwrap();
    std::fs::write(
        output.join(format!("{name}.xml")),
        package
            .get_part(&PackURI::new(SLIDE).unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
}

fn unchanged_parts(before: &[u8], after: &[u8]) {
    let before = OpcPackage::from_bytes(before).unwrap();
    let after = OpcPackage::from_bytes(after).unwrap();
    for name in [
        "/ppt/media/raster.png",
        "/ppt/media/opaque.bin",
        "/ppt/presentation.xml",
    ] {
        let uri = PackURI::new(name).unwrap();
        assert_eq!(
            before.get_part(&uri).unwrap().blob(),
            after.get_part(&uri).unwrap().blob()
        );
    }
}

fn main() {
    let output = PathBuf::from(std::env::args().nth(1).expect("output directory"));
    std::fs::create_dir_all(&output).unwrap();
    let source = source();
    let view = SourceBackedPresentation::from_reader(source.as_slice()).unwrap();
    let images = view.slide(0).unwrap().images().unwrap();
    assert_eq!(images.len(), 1);
    assert!(images[0].svg().is_none());
    emit(&output, "source", &source);

    let editor = SourceBackedPresentationEditor::from_reader(source.as_slice()).unwrap();
    let noop = editor.edit_svg_attachment(0, 0).unwrap().commit().unwrap();
    let mut unchanged = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut unchanged, &noop)
        .unwrap();
    assert_eq!(unchanged, source);

    let editor = SourceBackedPresentationEditor::from_reader(source.as_slice()).unwrap();
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(edit.attach_svg(SVG).unwrap());
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    assert!(inverse.apply(commit.snapshot()).unwrap().svg().is_none());
    let mut attached = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut attached, &commit)
        .unwrap();
    unchanged_parts(&source, &attached);
    let view = SourceBackedPresentation::from_reader(attached.as_slice()).unwrap();
    assert_eq!(
        view.slide(0).unwrap().read_svg_image(0).unwrap().bytes(),
        SVG
    );
    emit(&output, "attached", &attached);

    let editor = SourceBackedPresentationEditor::from_reader(attached.as_slice()).unwrap();
    let mut edit = editor.edit_svg_attachment(0, 0).unwrap();
    assert!(edit.detach().unwrap());
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .svg()
            .unwrap()
            .bytes(),
        SVG
    );
    let mut detached = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut detached, &commit)
        .unwrap();
    unchanged_parts(&source, &detached);
    let view = SourceBackedPresentation::from_reader(detached.as_slice()).unwrap();
    assert!(view.slide(0).unwrap().images().unwrap()[0].svg().is_none());
    emit(&output, "detached", &detached);
    assert_eq!(
        std::fs::read(output.join("source.xml")).unwrap(),
        std::fs::read(output.join("detached.xml")).unwrap()
    );
    println!(
        "Passed exact no-op, attach/publish/reopen, detach/publish/reopen, in-memory inverses, raster/opaque preservation, and exact selected slide restoration"
    );
}

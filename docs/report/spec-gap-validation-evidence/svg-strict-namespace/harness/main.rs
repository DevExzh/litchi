use std::fs;
use std::path::{Path, PathBuf};

use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_pptx::{SourceBackedPresentation, SourceBackedPresentationEditor};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_URI: &str = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const SLIDE: &str = "ppt/slides/slide1.xml";
const SLIDE_RELS: &str = "ppt/slides/_rels/slide1.xml.rels";
const SVG: &[u8] =
    br#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1"/></svg>"#;

fn strict_xml(bytes: Vec<u8>) -> Vec<u8> {
    let text = String::from_utf8(bytes).expect("all XML members are UTF-8");
    text.replace(PML, STRICT_PML)
        .replace(DML, STRICT_DML)
        .replace(REL, STRICT_REL)
        .into_bytes()
}

fn strictify_package(source: &[u8]) -> Vec<u8> {
    let reader = PhysPkgReader::new(source).unwrap();
    let mut writer = PhysPkgWriter::new();
    for name in reader.member_names().unwrap() {
        let bytes = reader.read_member(&name).unwrap();
        let bytes = if name.ends_with(".xml") || name.ends_with(".rels") {
            strict_xml(bytes)
        } else {
            bytes
        };
        let uri = PackURI::new(format!("/{name}")).unwrap();
        writer.write(&uri, &bytes).unwrap();
    }
    writer.finish().unwrap()
}

fn member(package: &[u8], name: &str) -> Vec<u8> {
    PhysPkgReader::new(package)
        .unwrap()
        .read_member(name)
        .unwrap()
}

fn write_output(output: &Path, name: &str, package: &[u8]) {
    fs::write(output.join(format!("{name}.pptx")), package).unwrap();
    fs::write(output.join(format!("{name}.xml")), member(package, SLIDE)).unwrap();
    fs::write(
        output.join(format!("{name}.rels.xml")),
        member(package, SLIDE_RELS),
    )
    .unwrap();
}

fn assert_strict_package(package: &[u8], attached: bool) {
    let root_rels = String::from_utf8(member(package, "_rels/.rels")).unwrap();
    assert!(root_rels.contains(&format!(r#"Type="{STRICT_REL}/officeDocument""#)));
    assert!(!root_rels.contains(&format!(r#"Type="{REL}/officeDocument""#)));

    let presentation_rels =
        String::from_utf8(member(package, "ppt/_rels/presentation.xml.rels")).unwrap();
    assert!(presentation_rels.contains(&format!(r#"Type="{STRICT_REL}/slide""#)));
    assert!(!presentation_rels.contains(&format!(r#"Type="{REL}/slide""#)));

    let slide_rels = String::from_utf8(member(package, SLIDE_RELS)).unwrap();
    assert!(slide_rels.contains(&format!(r#"Type="{STRICT_REL}/image""#)));
    assert!(!slide_rels.contains(&format!(r#"Type="{REL}/image""#)));
    if attached {
        assert!(
            slide_rels
                .matches(&format!(r#"Type="{STRICT_REL}/image""#))
                .count()
                >= 2
        );
    }

    let slide = String::from_utf8(member(package, SLIDE)).unwrap();
    assert!(slide.contains(&format!(r#"xmlns:p="{STRICT_PML}""#)));
    assert!(slide.contains(&format!(r#"xmlns:a="{STRICT_DML}""#)));
    assert!(slide.contains(&format!(r#"xmlns:r="{STRICT_REL}""#)));
    assert!(!slide.contains(&format!(r#"xmlns:p="{PML}""#)));
    assert!(!slide.contains(&format!(r#"xmlns:a="{DML}""#)));
    if attached {
        let expected =
            format!(r#"<asvg:svgBlip xmlns:asvg="{SVG_NS}" xmlns:r="{REL}" r:embed="rIdSvg"/>"#);
        assert!(
            slide.contains(&expected),
            "missing Transitional SVG attribute: {expected}"
        );
        assert!(slide.contains(SVG_URI));
    }
}

fn assert_same_member_set(left: &[u8], right: &[u8]) {
    let mut left_names = PhysPkgReader::new(left).unwrap().member_names().unwrap();
    let mut right_names = PhysPkgReader::new(right).unwrap().member_names().unwrap();
    left_names.sort();
    right_names.sort();
    assert_eq!(left_names, right_names);
}

fn direct_controls(output: &Path) {
    let transitional =
        format!(r#"<asvg:svgBlip xmlns:asvg="{SVG_NS}" xmlns:r="{REL}" r:embed="rIdSvg"/>"#);
    let strict =
        format!(r#"<asvg:svgBlip xmlns:asvg="{SVG_NS}" xmlns:r="{STRICT_REL}" r:embed="rIdSvg"/>"#);
    fs::write(output.join("direct-transitional-valid.xml"), transitional).unwrap();
    fs::write(output.join("direct-strict-invalid.xml"), strict).unwrap();
}

fn main() {
    let mut args = std::env::args().skip(1);
    let source_path = PathBuf::from(args.next().expect("source.pptx path"));
    let output = PathBuf::from(args.next().expect("output directory"));
    assert!(args.next().is_none(), "unexpected harness argument");
    fs::create_dir_all(&output).unwrap();

    let source = fs::read(&source_path).unwrap();
    let strict_source = strictify_package(&source);
    assert_strict_package(&strict_source, false);
    write_output(&output, "strict-source", &strict_source);
    direct_controls(&output);

    let source_view = SourceBackedPresentation::from_reader(strict_source.as_slice()).unwrap();
    let source_images = source_view.slide(0).unwrap().images().unwrap();
    assert_eq!(source_images.len(), 1);
    assert!(source_images[0].svg().is_none());

    let editor = SourceBackedPresentationEditor::from_reader(strict_source.as_slice()).unwrap();
    let mut attach = editor.edit_svg_attachment(0, 0).unwrap();
    assert_eq!(attach.source().raster_relationship_type(), rt::STRICT_IMAGE);
    assert!(attach.attach_svg(SVG).unwrap());
    let attach_commit = attach.commit().unwrap();
    let restored = attach_commit
        .patch()
        .inverse()
        .apply(attach_commit.snapshot())
        .unwrap();
    assert!(restored.svg().is_none());
    let mut attached = Vec::new();
    editor
        .publish_svg_attachment_commit_to_stream(&mut attached, &attach_commit)
        .unwrap();
    assert_strict_package(&attached, true);
    write_output(&output, "strict-attached", &attached);

    let attached_view = SourceBackedPresentation::from_reader(attached.as_slice()).unwrap();
    assert_eq!(
        attached_view
            .slide(0)
            .unwrap()
            .read_svg_image(0)
            .unwrap()
            .bytes(),
        SVG
    );
    let reopened_editor = SourceBackedPresentationEditor::from_reader(attached.as_slice()).unwrap();
    let mut detach = reopened_editor.edit_svg_attachment(0, 0).unwrap();
    assert!(detach.detach().unwrap());
    let detach_commit = detach.commit().unwrap();
    let restored_attached = detach_commit
        .patch()
        .inverse()
        .apply(detach_commit.snapshot())
        .unwrap();
    assert!(restored_attached.svg().is_some());
    let mut detached = Vec::new();
    reopened_editor
        .publish_svg_attachment_commit_to_stream(&mut detached, &detach_commit)
        .unwrap();
    assert_strict_package(&detached, false);
    write_output(&output, "strict-detached", &detached);

    assert_eq!(member(&strict_source, SLIDE), member(&detached, SLIDE));
    assert_eq!(
        member(&strict_source, SLIDE_RELS),
        member(&detached, SLIDE_RELS)
    );
    for name in [
        "ppt/media/raster.png",
        "ppt/media/opaque.bin",
        "ppt/presentation.xml",
    ] {
        assert_eq!(
            member(&strict_source, name),
            member(&detached, name),
            "{name}"
        );
    }
    assert_same_member_set(&strict_source, &detached);
    println!(
        "passed strict package conversion, Transitional SVG attribute proof, ordinary attach/commit/save/reopen/detach/inverse, Strict physical relationship preservation, and exact detached restoration"
    );
}

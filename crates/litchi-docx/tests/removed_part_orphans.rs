//! A removed picture or header takes the parts only it reached with it, and
//! keeps a part something else still reaches, also when one edit removes
//! several of them. The inverse restores the package exactly.

use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::{ElementPath, ElementStep};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

const IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HEADER: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture""#;

fn pack(name: &str) -> PackURI {
    PackURI::new(name).unwrap()
}

/// An inline picture whose blip holds `blip` (its attributes and children).
fn picture(id: u32, blip: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="{id}" name="P{id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="P{id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip {blip}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn png(tag: u8) -> Vec<u8> {
    let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
    png.push(tag);
    png
}

/// A package whose main document holds `paragraphs` (one picture each) and
/// relates `relationships` (id, type, target); `parts` are added beside it,
/// each with its own relationships.
fn package(
    paragraphs: &[String],
    section: &str,
    relationships: &[(&str, &str, &str)],
    parts: &[(&str, &str, Vec<u8>, &[(&str, &str, &str)])],
) -> Package {
    let body: String = paragraphs
        .iter()
        .map(|run| format!("<w:p><w:r><w:t>T</w:t></w:r>{run}</w:p>"))
        .collect();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NAMESPACES}><w:body>{body}<w:p><w:r><w:t>End</w:t></w:r></w:p><w:sectPr>{section}</w:sectPr></w:body></w:document>"#
    );
    let mut opc = OpcPackage::new();
    let mut main = BlobPart::new(
        pack("/word/document.xml"),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document.into_bytes(),
    );
    for (id, kind, target) in relationships {
        main.rels_mut()
            .try_add_relationship(
                (*kind).to_owned(),
                (*target).to_owned(),
                (*id).to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    opc.try_add_part(Box::new(main)).unwrap();
    for (name, content_type, bytes, relationships) in parts {
        let mut part = BlobPart::new(pack(name), (*content_type).to_owned(), bytes.clone());
        for (id, kind, target) in *relationships {
            part.rels_mut()
                .try_add_relationship(
                    (*kind).to_owned(),
                    (*target).to_owned(),
                    (*id).to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        opc.try_add_part(Box::new(part)).unwrap();
    }
    opc.relate_to("word/document.xml", OFFICE_DOCUMENT);
    Package::from_bytes(&PackageWriter::to_bytes(&opc).unwrap()).unwrap()
}

fn paragraph(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

fn part_names(package: &Package) -> Vec<String> {
    let mut names: Vec<String> = package
        .opc_package()
        .iter_parts()
        .map(|part| part.partname().as_str().to_owned())
        .collect();
    names.sort();
    names
}

#[test]
fn a_picture_naming_one_image_twice_removes_it_once() {
    let base = package(
        &[picture(
            1,
            r#"r:embed="rId10"><a:extLst><a:ext uri="{X0}"><x:b xmlns:x="urn:x" r:embed="rId11"/></a:ext><a:ext uri="{X1}"><x:b xmlns:x="urn:x" r:embed="rId12"/></a:ext></a:extLst>"#,
        )],
        "",
        &[
            ("rId10", IMAGE, "media/image1.png"),
            ("rId11", IMAGE, "media/image2.png"),
            ("rId12", IMAGE, "media/image1.png"),
        ],
        &[
            ("/word/media/image1.png", "image/png", png(1), &[]),
            ("/word/media/image2.png", "image/png", png(2), &[]),
        ],
    );
    let mut edit = base.edit().unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(part_names(&changed), ["/word/document.xml"]);
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
}

#[test]
fn one_edit_removing_every_user_of_a_shared_image_removes_it() {
    let pictures = [
        picture(1, r#"r:embed="rId10">"#),
        picture(2, r#"r:embed="rId11">"#),
    ];
    let images = [("/word/media/image1.png", "image/png", png(1), &[][..])];
    // Two pictures, two relationships, one image.
    let base = package(
        &pictures,
        "",
        &[
            ("rId10", IMAGE, "media/image1.png"),
            ("rId11", IMAGE, "media/image1.png"),
        ],
        &images,
    );
    let mut edit = base.edit().unwrap();
    edit.remove_picture_at(&paragraph(1), Position::new(0))
        .unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(part_names(&changed), ["/word/document.xml"]);
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());

    // Removing one of them keeps the image the other still shows.
    let mut edit = base.edit().unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        part_names(&changed),
        ["/word/document.xml", "/word/media/image1.png"]
    );

    // The pictures and a header showing the same image, all in one edit.
    let header = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:hdr {NAMESPACES}><w:p>{}</w:p></w:hdr>"#,
        picture(9, r#"r:embed="rIdX">"#)
    );
    let base = package(
        &pictures,
        r#"<w:headerReference w:type="default" r:id="rIdH"/>"#,
        &[
            ("rId10", IMAGE, "media/image1.png"),
            ("rId11", IMAGE, "media/image1.png"),
            ("rIdH", HEADER, "header1.xml"),
        ],
        &[
            images[0].clone(),
            (
                "/word/header1.xml",
                ct::WML_HEADER,
                header.into_bytes(),
                &[("rIdX", IMAGE, "media/image1.png")],
            ),
        ],
    );
    let mut edit = base.edit().unwrap();
    edit.remove_header_footer("/word/header1.xml").unwrap();
    edit.remove_picture_at(&paragraph(1), Position::new(0))
        .unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(part_names(&changed), ["/word/document.xml"]);
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), base.to_bytes().unwrap());
}

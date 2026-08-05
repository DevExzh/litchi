use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::{Error, Package};
use litchi_pptx::modern_comments::{
    Authors, List, MODERN_COMMENT_AUTHOR_CONTENT_TYPE, MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE,
    MODERN_COMMENT_CONTENT_TYPE, MODERN_COMMENT_RELATIONSHIP_TYPE,
};

const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/modern-comments/slide.xml");

#[test]
fn presentation_loads_the_modern_comment_graph() {
    let package = package_with_modern_comments();
    let graph = litchi_pptx::modern_comments::load_modern_comment_graph(
        package.opc().unwrap(),
    )
    .unwrap();

    let authors = graph.authors.unwrap();
    assert_eq!(authors.relationship_id, "rIdModernAuthors");
    assert_eq!(authors.part_name, "/ppt/commentAuthors.xml");
    assert!(authors.authors.authors.is_empty());
    assert_eq!(graph.comments.len(), 1);
    assert_eq!(graph.comments[0].slide_part_name, "/ppt/slides/slide1.xml");
    assert_eq!(graph.comments[0].relationship_id, "rIdModernComments");
    assert_eq!(graph.comments[0].part_name, "/ppt/comments/comment1.xml");
    assert!(graph.comments[0].comments.comments.is_empty());
}

#[test]
fn presentation_modern_comments_reject_external_author_relationships() {
    let mut package = package_with_modern_comments();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package = edit_package(package, |opc| {
            let presentation = opc.get_part_mut(&presentation_name).unwrap();
            presentation.rels_mut().remove("rIdModernAuthors");
            presentation.rels_mut().add_relationship(
                MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.to_string(),
                "https://example.invalid/authors.xml".to_string(),
                "rIdModernAuthors".to_string(),
                true,
            );
    });

    assert!(matches!(
        litchi_pptx::modern_comments::load_modern_comment_graph(package.opc().unwrap()),
        Err(Error::Invalid(message)) if message.contains("cannot be external")
    ));
}

fn package_with_modern_comments() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let authors_name = PackURI::new("/ppt/commentAuthors.xml").unwrap();
    let comments_name = PackURI::new("/ppt/comments/comment1.xml").unwrap();

    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    {
            opc.get_part_mut(&presentation_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE.to_string(),
                    "commentAuthors.xml".to_string(),
                    "rIdModernAuthors".to_string(),
                    false,
                );
            opc.add_part(Box::new(BlobPart::new(
                slide_name.clone(),
                ct::PML_SLIDE.to_string(),
                SLIDE_XML.to_vec(),
            )));
            opc.add_part(Box::new(BlobPart::new(
                authors_name,
                MODERN_COMMENT_AUTHOR_CONTENT_TYPE.to_string(),
                Authors::default().to_xml().unwrap(),
            )));
            opc.add_part(Box::new(BlobPart::new(
                comments_name,
                MODERN_COMMENT_CONTENT_TYPE.to_string(),
                List::default().to_xml().unwrap(),
            )));
            opc.get_part_mut(&slide_name).unwrap().rels_mut().add_relationship(
                MODERN_COMMENT_RELATIONSHIP_TYPE.to_string(),
                "../comments/comment1.xml".to_string(),
                "rIdModernComments".to_string(),
                false,
            );
    }
    Package::from_opc_package(opc).unwrap()
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}

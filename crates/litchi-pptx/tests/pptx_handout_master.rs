#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Error;
use litchi_pptx::Package;
use litchi_pptx::presentation_properties::metadata::handout::Master;
use quick_xml::Reader;
use quick_xml::events::Event;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/presentation.xml");
const HANDOUT_MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/handout-master.xml");
const WRONG_ROOT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/handout-master/wrong-root.xml");
const STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster";

#[test]
fn presentation_handout_master_is_resolved_by_the_owner_layers() {
    let package = package_with_handout_master();
    let handout_master = load_handout_master(&package).unwrap().unwrap();

    assert!(handout_master.header_footer.show_header);
    assert!(handout_master.header_footer.show_footer);
    assert!(handout_master.header_footer.show_slide_number);
    assert!(handout_master.header_footer.show_date_time);
    assert_eq!(handout_master.background_color.as_deref(), Some("112233"));
}

#[test]
fn handout_master_relationship_is_validated() {
    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        rt::THEME,
        "handoutMasters/handoutMaster1.xml",
        false,
    );
    assert!(matches!(
        load_handout_master(&package),
        Err(Error::Relationship(message))
            if message.contains("is not a handout-master relationship")
    ));

    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        rt::HANDOUT_MASTER,
        "https://example.invalid/handoutMaster.xml",
        true,
    );
    assert!(matches!(
        load_handout_master(&package),
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));

    let mut package = package_with_handout_master();
    let notes_name = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        notes_name,
        ct::PML_NOTES_MASTER.to_string(),
        b"<p:notesMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
            .to_vec(),
    )));
    replace_handout_relationship(
        &mut package,
        rt::HANDOUT_MASTER,
        "notesMasters/notesMaster1.xml",
        false,
    );
    assert!(matches!(
        load_handout_master(&package),
        Err(Error::ContentType { expected, actual })
            if expected == ct::PML_HANDOUT_MASTER && actual == ct::PML_NOTES_MASTER
    ));
}

#[test]
fn strict_handout_master_relationship_is_supported() {
    let mut package = package_with_handout_master();
    replace_handout_relationship(
        &mut package,
        STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE,
        "handoutMasters/handoutMaster1.xml",
        false,
    );

    assert!(load_handout_master(&package).unwrap().is_some());
}

#[test]
fn handout_master_root_is_validated() {
    let mut package = package_with_handout_master();
    let handout_name = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();
    package
        .get_part_mut(&handout_name)
        .unwrap()
        .set_blob(WRONG_ROOT_XML.to_vec());

    assert!(matches!(
        load_handout_master(&package),
        Err(Error::Invalid(message)) if message.contains("handoutMaster root")
    ));
}

fn load_handout_master(package: &OpcPackage) -> litchi_pptx::Result<Option<Master>> {
    let presentation = package.main_document_part()?;
    let Some(relationship) = presentation
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == "rIdHandout")
    else {
        return Ok(None);
    };
    if !matches!(
        relationship.reltype(),
        rt::HANDOUT_MASTER | STRICT_HANDOUT_MASTER_RELATIONSHIP_TYPE
    ) {
        return Err(Error::Relationship(
            "relationship is not a handout-master relationship".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(Error::Relationship(
            "handout-master relationship must be internal".to_string(),
        ));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != ct::PML_HANDOUT_MASTER {
        return Err(Error::ContentType {
            expected: ct::PML_HANDOUT_MASTER.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    let xml = std::str::from_utf8(part.blob()).map_err(|error| Error::Xml(error.to_string()))?;
    validate_root(xml)?;
    Ok(Some(Master::parse_xml(xml)?))
}

fn validate_root(xml: &str) -> litchi_pptx::Result<()> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"handoutMaster" {
                    return Err(Error::Invalid("handoutMaster root is invalid".to_string()));
                }
                return Ok(());
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => return Err(Error::Invalid("handoutMaster root is missing".to_string())),
            _ => {},
        }
    }
}

fn package_with_handout_master() -> OpcPackage {
    let mut package = Package::new().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_vec(bytes).unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let handout_name = PackURI::new("/ppt/handoutMasters/handoutMaster1.xml").unwrap();

    let presentation = package.get_part_mut(&presentation_name).unwrap();
    presentation.set_blob(PRESENTATION_XML.to_vec());
    presentation.rels_mut().add_relationship(
        rt::HANDOUT_MASTER.to_string(),
        "handoutMasters/handoutMaster1.xml".to_string(),
        "rIdHandout".to_string(),
        false,
    );
    package.add_part(Box::new(BlobPart::new(
        handout_name,
        ct::PML_HANDOUT_MASTER.to_string(),
        HANDOUT_MASTER_XML.to_vec(),
    )));
    package
}

fn replace_handout_relationship(
    package: &mut OpcPackage,
    relationship_type: &str,
    target: &str,
    is_external: bool,
) {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package.get_part_mut(&presentation_name).unwrap();
    presentation.rels_mut().remove("rIdHandout");
    presentation.rels_mut().add_relationship(
        relationship_type.to_string(),
        target.to_string(),
        "rIdHandout".to_string(),
        is_external,
    );
}

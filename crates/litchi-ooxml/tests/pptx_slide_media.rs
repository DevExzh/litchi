use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/media/presentation.xml");
const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/media/slide.xml");

#[test]
fn slide_loads_inert_audio_resources() {
    let package = package_with_audio(false);
    let slide = package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .into_iter()
        .next()
        .unwrap();

    let media = slide.media().unwrap();
    assert_eq!(media.pictures.len(), 1);
    assert_eq!(media.pictures[0].shape_id, 7);
    assert_eq!(media.pictures[0].name, "clip.mp3");
    assert_eq!(media.pictures[0].relationship_id, "rIdAudio");
    assert_eq!(
        media.pictures[0].resource.as_ref().unwrap().part_name,
        "/ppt/media/media1.mp3"
    );
    assert_eq!(
        media.pictures[0].resource.as_ref().unwrap().data,
        b"opaque audio payload"
    );
}

#[test]
fn slide_rejects_external_media_resources() {
    let package = package_with_audio(true);
    let slide = package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .into_iter()
        .next()
        .unwrap();

    assert!(matches!(
        slide.media(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("not fetched")
    ));
}

fn package_with_audio(external: bool) -> Package {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let media_name = PackURI::new("/ppt/media/media1.mp3").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.to_string(),
        PRESENTATION_XML.to_vec(),
    );
    presentation.rels_mut().add_relationship(
        rt::SLIDE.to_string(),
        "slides/slide1.xml".to_string(),
        "rIdSlide1".to_string(),
        false,
    );
    let mut slide = BlobPart::new(slide_name, ct::PML_SLIDE.to_string(), SLIDE_XML.to_vec());
    slide.rels_mut().add_relationship(
        rt::AUDIO.to_string(),
        if external {
            "https://example.invalid/clip.mp3"
        } else {
            "../media/media1.mp3"
        }
        .to_string(),
        "rIdAudio".to_string(),
        external,
    );

    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(presentation));
    opc.add_part(Box::new(slide));
    if !external {
        opc.add_part(Box::new(BlobPart::new(
            media_name,
            "audio/mpeg".to_string(),
            b"opaque audio payload".to_vec(),
        )));
    }
    opc.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    Package::from_opc_package(opc).unwrap()
}

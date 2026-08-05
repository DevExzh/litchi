use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Error;
use litchi_pptx::presentation_properties::metadata::tracks::{
    Block, CONTENT_TYPE, RELATIONSHIP_TYPE, Target, load,
};

const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/caption-tracks/slide.xml");
const TRACK_DATA: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/caption-tracks/captions.vtt");

#[test]
fn presentation_caption_tracks_load_internal_webvtt() {
    let package = package_with_internal_track();
    let tracks = load(&package).unwrap();

    assert_eq!(tracks.len(), 1);
    assert_eq!(tracks[0].source_part_name, "/ppt/slides/slide1.xml");
    assert_eq!(tracks[0].relationship_id, "rIdCaptions");
    let Target::Internal { part_name, track } = &tracks[0].target else {
        panic!("expected an internal WebVTT track");
    };
    assert_eq!(part_name, "/ppt/media/captions.vtt");
    assert_eq!(track.blocks.len(), 1);
    let Block::Cue(cue) = &track.blocks[0] else {
        panic!("expected a cue block");
    };
    assert_eq!(cue.start_milliseconds, 0);
    assert_eq!(cue.end_milliseconds, 2_000);
    assert_eq!(cue.payload, ["Welcome"]);
}

#[test]
fn presentation_caption_tracks_reject_invalid_sources() {
    let mut package = OpcPackage::new();
    let presentation = BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        ct::PML_PRESENTATION_MAIN.to_string(),
        br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
            .to_vec(),
    );
    package.add_part(Box::new(presentation));
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.to_string(),
            "https://example.invalid/captions.vtt".to_string(),
            "rIdCaptions".to_string(),
            true,
        );

    assert!(matches!(
        load(&package),
        Err(Error::Invalid(message)) if message.contains("invalid source")
    ));
}

fn package_with_internal_track() -> OpcPackage {
    let mut package = OpcPackage::new();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let track_name = PackURI::new("/ppt/media/captions.vtt").unwrap();

    package.add_part(Box::new(BlobPart::new(
        slide_name.clone(),
        ct::PML_SLIDE.to_string(),
        SLIDE_XML.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        track_name.clone(),
        CONTENT_TYPE.to_string(),
        TRACK_DATA.to_vec(),
    )));
    package
        .get_part_mut(&slide_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.to_string(),
            "../media/captions.vtt".to_string(),
            "rIdCaptions".to_string(),
            false,
        );
    package
}

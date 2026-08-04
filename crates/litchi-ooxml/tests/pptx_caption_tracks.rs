use litchi_ooxml::pptx::Package;
use litchi_ooxml::pptx::tracks::{
    TRACK_CONTENT_TYPE, TRACK_RELATIONSHIP_TYPE, TrackTarget, WebVttBlock,
};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;

const SLIDE_XML: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/caption-tracks/slide.xml");
const TRACK_DATA: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/caption-tracks/captions.vtt");

#[test]
fn presentation_caption_tracks_load_internal_webvtt() {
    let package = package_with_internal_track();
    let tracks = package.presentation().unwrap().caption_tracks().unwrap();

    assert_eq!(tracks.len(), 1);
    assert_eq!(tracks[0].source_part_name, "/ppt/slides/slide1.xml");
    assert_eq!(tracks[0].relationship_id, "rIdCaptions");
    let TrackTarget::Internal { part_name, track } = &tracks[0].target else {
        panic!("expected an internal WebVTT track");
    };
    assert_eq!(part_name, "/ppt/media/captions.vtt");
    assert_eq!(track.blocks.len(), 1);
    let WebVttBlock::Cue(cue) = &track.blocks[0] else {
        panic!("expected a cue block");
    };
    assert_eq!(cue.start_milliseconds, 0);
    assert_eq!(cue.end_milliseconds, 2_000);
    assert_eq!(cue.payload, ["Welcome"]);
}

#[test]
fn presentation_caption_tracks_reject_invalid_sources() {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&presentation_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    TRACK_RELATIONSHIP_TYPE.to_string(),
                    "https://example.invalid/captions.vtt".to_string(),
                    "rIdCaptions".to_string(),
                    true,
                );
            Ok(())
        })
        .unwrap();

    assert!(matches!(
        package.presentation().unwrap().caption_tracks(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("invalid source")
    ));
}

fn package_with_internal_track() -> Package {
    let mut package = Package::new().unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let track_name = PackURI::new("/ppt/media/captions.vtt").unwrap();

    package
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                slide_name.clone(),
                ct::PML_SLIDE.to_string(),
                SLIDE_XML.to_vec(),
            )));
            opc.add_part(Box::new(BlobPart::new(
                track_name,
                TRACK_CONTENT_TYPE.to_string(),
                TRACK_DATA.to_vec(),
            )));
            opc.get_part_mut(&slide_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    TRACK_RELATIONSHIP_TYPE.to_string(),
                    "../media/captions.vtt".to_string(),
                    "rIdCaptions".to_string(),
                    false,
                );
            Ok(())
        })
        .unwrap();
    package
}

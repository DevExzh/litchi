use super::model::{Caption, CaptionTarget, DisplayLocation, MediaKey, MediaMetadata, TracksInfo};
use super::tracks_info;
use super::transaction::Snapshot;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const P173: &str = "http://schemas.microsoft.com/office/powerpoint/2017/3/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ID: &str = "{11111111-1111-1111-1111-111111111111}";

fn source() -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14}" xmlns:p15="{P15}" xmlns:p173="{P173}" xmlns:r="{REL}">
  <p:pic>
    <p:nvPicPr><p:cNvPr id="7"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
    <p:blipFill><p:blip r:embed="rIdImage"/></p:blipFill>
    <p:spPr/>
    <p:extLst><p:ext uri="{{opaque}}"><p14:media r:embed="rIdMedia">
      <p173:tracksInfo displayLoc="media">
        <p173:track id="{ID}" label="English" lang="en" r:embed="rIdTrack"/>
      </p173:tracksInfo>
      <x:future xmlns:x="urn:future"><x:data>keep</x:data></x:future>
    </p14:media><p15:isNarration val="true"/></p:ext></p:extLst>
  </p:pic>
</p:sld>"#
    )
    .into_bytes()
}

fn snapshot() -> Snapshot {
    let bytes = source();
    let key = MediaKey {
        slide_part_name: "/ppt/slides/slide1.xml".into(),
        shape_id: 7,
    };
    let found = tracks_info::discover(&bytes, &key)
        .unwrap()
        .expect("fixture media shape");
    let metadata = MediaMetadata {
        key,
        media_relationship_id: Some("rIdMedia".into()),
        tracks_info: Some(TracksInfo {
            display_location: DisplayLocation::Media,
            captions: vec![Caption {
                id: ID.into(),
                label: "English".into(),
                language: Some("en".into()),
                target: CaptionTarget::Internal {
                    part_name: "/ppt/media/captions1.vtt".into(),
                    content_type: "text/vtt".into(),
                },
            }],
        }),
        narration: Some(true),
    };
    Snapshot::from_wire("/ppt/slides/slide1.xml".into(), bytes, found, metadata).unwrap()
}

#[test]
fn no_op_commit_preserves_exact_slide_source() {
    let snapshot = snapshot();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.is_changed());
    assert_eq!(commit.snapshot().source_xml(), snapshot.source_xml());
}

#[test]
fn typed_edits_replace_only_modeled_attributes_and_preserve_unknown_xml() {
    let snapshot = snapshot();
    let mut edit = snapshot.edit();
    edit.set_display_location(DisplayLocation::Slide).unwrap();
    edit.set_caption_identity(0, ID, "English captions")
        .unwrap();
    edit.set_caption_language(0, Some("en-US".into())).unwrap();
    edit.set_narration(Some(false)).unwrap();
    let commit = edit.commit().unwrap();
    let output = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(output.contains(r#"displayLoc="slide""#));
    assert!(output.contains(r#"label="English captions""#));
    assert!(output.contains(r#"lang="en-US""#));
    assert!(output.contains(r#"isNarration val="false""#));
    assert!(output.contains("<x:data>keep</x:data>"));
}

#[test]
fn invalid_caption_edits_are_staged_atomically() {
    let snapshot = snapshot();
    let before = snapshot.edit().snapshot().clone();
    let mut edit = snapshot.edit();
    assert!(edit.set_caption_identity(0, "not-a-guid", "bad").is_err());
    assert_eq!(edit.snapshot(), &before);
    assert!(edit.set_caption_language(9, Some("en".into())).is_err());
    assert_eq!(edit.snapshot(), &before);
}

#[test]
fn malformed_track_metadata_is_rejected_without_a_partial_snapshot() {
    let mut bytes = source();
    let needle = br#"r:embed="rIdTrack""#;
    let position = bytes
        .windows(needle.len())
        .position(|window| window == needle)
        .unwrap();
    bytes.splice(
        position..position + needle.len(),
        br#"r:embed=""#.iter().copied(),
    );
    let key = MediaKey {
        slide_part_name: "/ppt/slides/slide1.xml".into(),
        shape_id: 7,
    };
    assert!(tracks_info::discover(&bytes, &key).is_err());
}

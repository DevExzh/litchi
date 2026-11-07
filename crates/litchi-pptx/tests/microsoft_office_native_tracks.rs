//! Native caption metadata editing; media playback and fresh Office validation are absent.
use litchi_pptx::presentation_properties::metadata::tracks::{
    self, Block, DisplayLocation, Target,
};
use litchi_pptx::{Error, Package};
use std::{collections::BTreeMap, io::Cursor};
const FILE: &str = "powerpoint-modern-media.pptx";
const SLIDE: &str = "ppt/slides/slide2.xml";
const TRACK_ID: &str = "{22E87EF0-5BD7-B344-840B-341EC7730D78}";
const NEW_ID: &str = "{11111111-2222-4333-8444-555555555555}";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

// Inverse replay after reopening restores exact OPC payloads. The edited ZIP
// member is recompressed; its producer-specific compression metadata is not owned.
#[track_caller]
fn assert_members_equal(a: &[u8], b: &[u8]) {
    assert_eq!(changed(a, b), Vec::<String>::new());
}

fn package() -> Package {
    Package::from_reader(Cursor::new(native(FILE))).unwrap()
}
fn save(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    p.save(file.path()).unwrap();
    Package::open(file.path()).unwrap()
}
fn cues(p: &Package) {
    let e = p.edit_caption_tracks().unwrap();
    let t = e.caption_tracks().unwrap();
    assert_eq!(t.len(), 1);
    assert_eq!(t[0].source_part_name, "/ppt/slides/slide2.xml");
    assert_eq!(t[0].relationship_id, "rId4");
    let Target::Internal { part_name, track } = &t[0].target else {
        panic!("internal captions")
    };
    assert_eq!(part_name, "/ppt/media/track1.vtt");
    assert_eq!(track.header_text, None);
    assert!(track.headers.is_empty());
    assert_eq!(track.blocks.len(), 2);
    for (b, (start, end, text)) in track.blocks.iter().zip([
        (0, 1000, "CAPTION_FIRST_007"),
        (1000, 2000, "CAPTION_SECOND_007"),
    ]) {
        let Block::Cue(c) = b else { panic!("cue") };
        assert_eq!((c.start_milliseconds, c.end_milliseconds), (start, end));
        assert_eq!(c.identifier, None);
        assert!(c.settings.is_empty());
        assert_eq!(c.payload, [text]);
    }
    assert_eq!(
        tracks::File::parse(
            &track.to_bytes(tracks::Limits::DEFAULT).unwrap(),
            tracks::Limits::DEFAULT
        )
        .unwrap(),
        *track
    );
}
#[test]
fn native_tracks_keep_exact_caption_identity_cues_and_absent_narration() {
    let p = package();
    let before = p.to_bytes().unwrap();
    cues(&p);
    let e = p.edit_caption_tracks().unwrap();
    let m = e.media(1, 4).unwrap().unwrap();
    assert_eq!(m.source_xml(), members(&before)[SLIDE]);
    assert_eq!(m.metadata().key.slide_part_name, "/ppt/slides/slide2.xml");
    assert_eq!(m.metadata().key.shape_id, 4);
    assert_eq!(m.metadata().media_relationship_id.as_deref(), Some("rId2"));
    assert_eq!(m.metadata().narration, None);
    let t = m.metadata().tracks_info.as_ref().unwrap();
    assert_eq!(t.display_location, DisplayLocation::Media);
    assert_eq!(t.captions.len(), 1);
    assert_eq!(t.captions[0].id, TRACK_ID);
    assert_eq!(t.captions[0].label, "English Fixture 007");
    assert_eq!(t.captions[0].language.as_deref(), Some("en-us"));
    assert_eq!(
        t.captions[0].target,
        tracks::CaptionTarget::Internal {
            part_name: "/ppt/media/track1.vtt".into(),
            content_type: tracks::CONTENT_TYPE.into()
        }
    );
    assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), before);
    assert_eq!(before, native(FILE));
}
#[test]
fn native_caption_metadata_update_preserves_trim_fade_cues_and_every_other_member() {
    let p = package();
    let before = p.to_bytes().unwrap();
    let mut e = p.edit_caption_tracks().unwrap();
    e.update(1, 4, |d| {
        d.set_display_location(DisplayLocation::Slide)?;
        d.set_caption_identity(0, NEW_ID, "French & English <captions>")?;
        d.set_caption_language(0, Some("fr-FR".into()))
    })
    .unwrap();
    let staged = e.media(1, 4).unwrap().unwrap().metadata().clone();
    assert!(matches!(
        e.update(1, 4, |d| d.set_caption_identity(0, "bad-guid", "bad")),
        Err(Error::Invalid(_))
    ));
    assert!(matches!(
        e.update(1, 4, |d| d.set_caption_language(99, None)),
        Err(Error::Invalid(_))
    ));
    assert!(
        matches!(e.update(1,4,|d| d.set_narration(Some(true))),Err(Error::Invalid(message)) if message.contains("cannot create isNarration"))
    );
    assert_eq!(e.media(1, 4).unwrap().unwrap().metadata(), &staged);
    let c = e.commit().unwrap();
    let reopened = save(c.package());
    let after = reopened.to_bytes().unwrap();
    assert_eq!(changed(&before, &after), [SLIDE]);
    let expected = String::from_utf8(members(&before)[SLIDE].clone())
        .unwrap()
        .replace(r#"displayLoc="media""#, r#"displayLoc="slide""#)
        .replace(TRACK_ID, NEW_ID)
        .replace(
            "English Fixture 007",
            "French &amp; English &lt;captions&gt;",
        )
        .replace(r#"lang="en-us""#, r#"lang="fr-FR""#);
    assert_eq!(members(&after)[SLIDE], expected.as_bytes());
    assert_eq!(
        reopened
            .edit_caption_tracks()
            .unwrap()
            .media(1, 4)
            .unwrap()
            .unwrap()
            .metadata(),
        &staged
    );
    cues(&reopened);
    assert!(matches!(
        reopened.apply_caption_tracks(c.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        p.apply_caption_tracks(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        after
    );
    assert_members_equal(
        &reopened
            .apply_caption_tracks(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    let mut same = reopened.edit_caption_tracks().unwrap();
    same.update(1, 4, |d| d.set_caption_language(0, Some("fr-FR".into())))
        .unwrap();
    assert_eq!(same.commit().unwrap().package().to_bytes().unwrap(), after);
    assert_eq!(p.to_bytes().unwrap(), before);
}
#[test]
fn native_caption_language_remove_and_restore_keep_a_valid_owner_and_exact_inverse() {
    let p = package();
    let before = p.to_bytes().unwrap();
    let mut e = p.edit_caption_tracks().unwrap();
    e.update(1, 4, |d| d.set_caption_language(0, None)).unwrap();
    let c = e.commit().unwrap();
    let removed = save(c.package());
    let bytes = removed.to_bytes().unwrap();
    assert_eq!(changed(&before, &bytes), [SLIDE]);
    assert_eq!(
        removed
            .edit_caption_tracks()
            .unwrap()
            .media(1, 4)
            .unwrap()
            .unwrap()
            .metadata()
            .tracks_info
            .as_ref()
            .unwrap()
            .captions[0]
            .language,
        None
    );
    cues(&removed);
    let mut e = removed.edit_caption_tracks().unwrap();
    e.update(1, 4, |d| d.set_caption_language(0, None)).unwrap();
    assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
    let mut e = removed.edit_caption_tracks().unwrap();
    e.update(1, 4, |d| d.set_caption_language(0, Some("en-us".into())))
        .unwrap();
    let added = e.commit().unwrap();
    let restored = save(added.package());
    assert_eq!(
        restored
            .edit_caption_tracks()
            .unwrap()
            .media(1, 4)
            .unwrap()
            .unwrap()
            .metadata(),
        p.edit_caption_tracks()
            .unwrap()
            .media(1, 4)
            .unwrap()
            .unwrap()
            .metadata()
    );
    assert_eq!(changed(&before, &restored.to_bytes().unwrap()), [SLIDE]); // whitespace from removal is retained
    assert_eq!(
        restored
            .apply_caption_tracks(&added.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert_members_equal(
        &removed
            .apply_caption_tracks(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert_eq!(native(FILE), before);
}

#[test]
fn native_source_can_add_an_inert_webvtt_part_without_replacing_the_native_track() {
    let p = package();
    let before = p.to_bytes().unwrap();
    let mut e = p.edit_caption_tracks().unwrap();
    let original = e.caption_tracks().unwrap().remove(0);
    assert!(
        matches!(e.set(&original),Err(Error::Invalid(message)) if message=="Track relationship ID already exists")
    );
    let mut added = original.clone();
    added.relationship_id = "rIdLitchiCaptions".into();
    assert!(
        matches!(e.set(&added),Err(Error::Invalid(message)) if message.contains("already exists"))
    );
    let Target::Internal { part_name, track } = &mut added.target else {
        panic!("internal")
    };
    *part_name = "/ppt/media/litchi-captions.vtt".into();
    let Block::Cue(cue) = &mut track.blocks[0] else {
        panic!("cue")
    };
    cue.payload = vec!["LITCHI_DERIVATIVE_CAPTION".into()];
    e.set(&added).unwrap();
    assert!(
        matches!(e.set(&added),Err(Error::Invalid(message)) if message=="Track relationship ID already exists")
    );
    let c = e.commit().unwrap();
    let out = save(c.package());
    let bytes = out.to_bytes().unwrap();
    assert_eq!(
        changed(&before, &bytes),
        [
            "ppt/media/litchi-captions.vtt",
            "ppt/slides/_rels/slide2.xml.rels"
        ]
    );
    let all = out.edit_caption_tracks().unwrap().caption_tracks().unwrap();
    assert_eq!(all.len(), 2);
    assert_eq!(
        all.iter().find(|v| v.relationship_id == "rId4"),
        Some(&original)
    );
    assert_eq!(
        all.iter()
            .find(|v| v.relationship_id == "rIdLitchiCaptions"),
        Some(&added)
    );
    // Creating a related WebVTT part does not create another media descriptor.
    let m = out
        .edit_caption_tracks()
        .unwrap()
        .media(1, 4)
        .unwrap()
        .unwrap();
    assert_eq!(m.metadata().tracks_info.as_ref().unwrap().captions.len(), 1);
    assert_eq!(
        m.metadata().tracks_info.as_ref().unwrap().captions[0].id,
        TRACK_ID
    );
    assert_members_equal(
        &out.apply_caption_tracks(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert_eq!(p.to_bytes().unwrap(), before);
}

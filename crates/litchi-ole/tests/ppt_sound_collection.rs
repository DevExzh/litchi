use litchi_ole::ppt::writer::build_sound_collection;
use litchi_ole::ppt::{Package, PowerPointSoundCollection, PptRecord};
use std::collections::HashSet;

#[test]
fn parses_poi_sound_fixture_byte_for_byte_when_available() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow");
    let ppt = root.join("sound.ppt");
    let wav = root.join("ringin.wav");
    if !ppt.exists() || !wav.exists() {
        return;
    }
    let mut package = Package::open(ppt).unwrap();
    let presentation = package.presentation().unwrap();
    let collection = presentation.embedded_sounds().unwrap().unwrap();
    let sound = collection
        .sounds
        .iter()
        .find(|sound| sound.name.eq_ignore_ascii_case("ringin.wav"))
        .unwrap();
    assert_eq!(sound.data, std::fs::read(wav).unwrap());
}

#[test]
fn existing_writer_round_trips_through_strict_reader() {
    let mut ids = HashSet::new();
    ids.insert(8);
    let (bytes, mapping) = build_sound_collection(&ids).unwrap();
    let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    let collection = PowerPointSoundCollection::parse(&record).unwrap();
    assert_eq!(collection.sounds.len(), 1);
    assert_eq!(collection.sounds[0].id, mapping[&8]);
    assert_eq!(&collection.sounds[0].data[..4], b"RIFF");
}

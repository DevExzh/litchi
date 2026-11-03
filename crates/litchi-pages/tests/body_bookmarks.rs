//! Ranged body bookmarks (`litchi_pages::bookmark`, record 0371).

use litchi_pages::Package;
use litchi_pages::bookmark::{Error, Name, Refusal, Settings, TextRange, Visibility};

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn named(name: &str) -> Settings {
    Settings::new().with_name(Name::new(name).unwrap())
}

#[test]
fn add_update_remove_round_trip_and_invert() {
    let package = Package::from_bytes(BASIC).unwrap();
    assert!(package.body_bookmarks().unwrap().is_empty());

    let mut edit = package.edit_body_bookmarks().unwrap();
    edit.add(TextRange::from_utf16_indexes(0, 4).unwrap(), named("Intro"))
        .unwrap();
    let added = edit.commit().unwrap();
    let bookmarks = added.package().body_bookmarks().unwrap();
    assert_eq!(bookmarks.len(), 1);
    assert_eq!(
        bookmarks[0].range(),
        TextRange::from_utf16_indexes(0, 4).unwrap()
    );
    assert_eq!(bookmarks[0].settings(), &named("Intro"));
    let restored = added
        .package()
        .apply_body_bookmarks(&added.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), BASIC);

    let mut edit = added.package().edit_body_bookmarks().unwrap();
    edit.update(
        "Intro",
        TextRange::from_utf16_indexes(1, 3).unwrap(),
        named("Intro").with_visibility(Visibility::Hidden),
    )
    .unwrap();
    let updated = edit.commit().unwrap();
    let bookmarks = updated.package().body_bookmarks().unwrap();
    assert_eq!(
        bookmarks[0].range(),
        TextRange::from_utf16_indexes(1, 3).unwrap()
    );
    assert_eq!(bookmarks[0].settings().visibility(), Visibility::Hidden);

    let mut edit = updated.package().edit_body_bookmarks().unwrap();
    edit.remove(0).unwrap();
    let removed = edit.commit().unwrap();
    assert!(removed.package().body_bookmarks().unwrap().is_empty());
    assert_eq!(removed.package().text().unwrap(), package.text().unwrap());
}

#[test]
fn overlaps_missing_selections_and_bad_ranges_are_refused() {
    let package = Package::from_bytes(BASIC).unwrap();
    let mut edit = package.edit_body_bookmarks().unwrap();
    edit.add(TextRange::from_utf16_indexes(0, 4).unwrap(), named("A"))
        .unwrap();
    let added = edit.commit().unwrap().into_package();

    let mut edit = added.edit_body_bookmarks().unwrap();
    edit.add(TextRange::from_utf16_indexes(2, 6).unwrap(), named("B"))
        .unwrap();
    assert_eq!(edit.commit().unwrap_err(), Error::Refused(Refusal::Overlap));

    let mut edit = added.edit_body_bookmarks().unwrap();
    assert_eq!(
        edit.remove("missing").unwrap_err(),
        Error::Refused(Refusal::NotFound)
    );
    edit.add(
        TextRange::from_utf16_indexes(5, 100_000).unwrap(),
        named("C"),
    )
    .unwrap();
    assert_eq!(
        edit.commit().unwrap_err(),
        Error::Refused(Refusal::InvalidRange)
    );
}

/// Record 0428: a new bookmark's text-attribute UUID is derived, not drawn,
/// so the same edit on the same input is byte-identical.
#[test]
fn added_bookmark_output_is_deterministic() {
    let run = || {
        let package = Package::from_bytes(BASIC).unwrap();
        let mut edit = package.edit_body_bookmarks().unwrap();
        edit.add(TextRange::from_utf16_indexes(0, 4).unwrap(), named("Intro"))
            .unwrap();
        bytes(edit.commit().unwrap().package())
    };
    assert_eq!(run(), run());
}

use litchi_ole::ppt::{Package, PowerPointHeaderFooterParent, PowerPointHeaderFooterScope};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/slideshow")
        .join(name)
}

fn read(name: &str) -> litchi_ole::ppt::PowerPointHeaderFooters {
    let mut package = Package::open(fixture(name)).expect("open POI fixture");
    package
        .presentation()
        .expect("parse presentation")
        .header_footers()
        .expect("parse header/footer records")
}

#[test]
fn poi_headers_footers_ppt_covers_all_scopes() {
    let values = read("headers_footers.ppt");
    let slides = values.presentation_slides().expect("slide defaults");
    assert_eq!(slides.display_footer(), Some("Global Slide Footer"));
    assert!(slides.options.show_footer);
    assert!(slides.options.show_slide_number);

    let notes = values.notes_and_handouts().expect("notes defaults");
    assert_eq!(notes.display_header(), Some("Notes Header"));
    assert_eq!(notes.display_footer(), Some("Notes Footer"));
    assert!(notes.options.show_header);
    assert!(notes.options.show_footer);

    assert!(values.entries().iter().any(|entry| {
        matches!(
            entry.scope,
            PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::Slide,
                ..
            }
        ) && entry.display_footer() == Some("per-slide footer")
            && entry.display_user_date() == Some("custom date format")
    }));
}

#[test]
fn poi_headers_footers_2007_ppt_preserves_dates_and_overrides() {
    let values = read("headers_footers_2007.ppt");
    let slides = values.presentation_slides().expect("slide defaults");
    assert_eq!(slides.display_footer(), Some("THE FOOTER TEXT"));
    assert_eq!(
        slides.display_user_date(),
        Some("Wednesday, August 06, 2008")
    );

    let notes = values.notes_and_handouts().expect("notes defaults");
    assert_eq!(notes.display_header(), Some("THE NOTES HEADER TEXT"));
    assert_eq!(notes.display_footer(), Some("THE NOTES FOOTER TEXT"));

    assert!(values.placeholder_displays().iter().any(|display| {
        matches!(display.scope, PowerPointHeaderFooterScope::Local { .. })
            && display.text.footer.as_deref() == Some("THE FOOTER TEXT FOR SLIDE 2")
            && display.text.user_date.as_deref() == Some("August 06, 2008")
    }));
}

#[test]
fn poi_bug_58144_2003_footer_is_discoverable() {
    let values = read("bug58144-headers-footers-2003.ppt");
    assert!(values.entries().iter().any(|entry| {
        entry.display_footer() == Some("Confidential") && entry.display_header().is_none()
    }));
}

#[test]
fn poi_bug_58144_2007_footer_is_discoverable() {
    let values = read("bug58144-headers-footers-2007.ppt");
    assert!(
        values
            .entries()
            .iter()
            .any(|entry| entry.display_footer() == Some("Slide footer"))
    );
}

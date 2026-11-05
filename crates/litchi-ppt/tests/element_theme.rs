#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "theme tests use panic-driven assertions"
)]

//! The `PowerPoint` 2007 theme of a legacy presentation (record 0323).

use litchi_ppt::Package;
use std::path::PathBuf;

fn package(path: &str) -> Package {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    Package::open(path).expect("PPT fixture")
}

#[test]
fn round_trip_theme_reads_colors_and_fonts() {
    let theme = package("poi/test-data/slideshow/headers_footers_2007.ppt")
        .presentation()
        .unwrap()
        .theme()
        .unwrap()
        .expect("theme atom");
    assert_eq!(theme.name.as_deref(), Some("Office"));
    assert!(theme.colors.iter().all(Option::is_some));
    assert_eq!(theme.color("dk2"), Some([0x1F, 0x49, 0x7D]));
    assert_eq!(theme.color("folHlink"), Some([0x80, 0x00, 0x80]));
    assert_eq!(theme.major.latin.as_deref(), Some("Calibri"));
    assert_eq!(theme.minor.latin.as_deref(), Some("Calibri"));
}

#[test]
fn a_presentation_without_a_theme_atom_reads_none() {
    assert!(
        package("poi/test-data/slideshow/basic_test_ppt_file.ppt")
            .presentation()
            .unwrap()
            .theme()
            .unwrap()
            .is_none()
    );
}

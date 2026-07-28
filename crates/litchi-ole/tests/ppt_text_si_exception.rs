//! Tests for TextSIExceptionAtom defaults and OutlineTextRefAtom references
//! against real PowerPoint fixtures.

use litchi_ole::ppt::Package;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

#[test]
fn reads_text_special_info_defaults() {
    let mut package = Package::open(fixture("basic_test_ppt_file.ppt")).unwrap();
    let defaults = package
        .presentation()
        .unwrap()
        .text_special_info_defaults()
        .unwrap()
        .expect("TextSIExceptionAtom present");
    assert!(defaults.spelling().unwrap().clean());
    assert!(!defaults.spelling().unwrap().error());
    assert_eq!(defaults.language(), Some(0x0809));
}

#[test]
fn reads_outline_text_refs_from_shape_textboxes() {
    let mut package = Package::open(fixture("basic_test_ppt_file.ppt")).unwrap();
    let indices: Vec<u32> = package
        .presentation()
        .unwrap()
        .outline_text_refs()
        .unwrap()
        .iter()
        .map(|reference| reference.get())
        .collect();
    assert_eq!(indices.len(), 6);
    assert!(indices.iter().all(|index| *index < 100));
}

#[test]
fn slides_without_refs_report_empty() {
    let mut package = Package::open(fixture("datetime.ppt")).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let total: usize = slides
        .iter()
        .map(|slide| slide.outline_text_refs().len())
        .sum();
    assert_eq!(total, 0);
}

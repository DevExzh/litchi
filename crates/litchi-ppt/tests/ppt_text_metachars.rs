//! Tests for header/footer metacharacter atoms (MS-PPT 2.9.47-2.9.52)
//! against real PowerPoint fixtures.

use litchi_ppt::{Package, PowerPointMetacharKind};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

fn metachars(name: &str) -> Vec<litchi_ppt::PowerPointTextMetachar> {
    let mut package = Package::open(fixture(name)).expect("open POI fixture");
    package
        .presentation()
        .expect("parse presentation")
        .text_metachars()
        .expect("parse metachar atoms")
}

fn count(kinds: &[PowerPointMetacharKind], kind: PowerPointMetacharKind) -> usize {
    kinds.iter().filter(|value| **value == kind).count()
}

#[test]
fn reads_slide_number_header_and_footer_metachars() {
    let values = metachars("headers_footers.ppt");
    let kinds = values.iter().map(|value| value.kind()).collect::<Vec<_>>();
    assert_eq!(count(&kinds, PowerPointMetacharKind::SlideNumber), 3);
    assert_eq!(count(&kinds, PowerPointMetacharKind::Footer), 3);
    assert_eq!(count(&kinds, PowerPointMetacharKind::Header), 2);
    assert_eq!(count(&kinds, PowerPointMetacharKind::GenericDate), 3);
    assert!(values.iter().all(|value| value.datetime_format().is_none()));
}

#[test]
fn reads_datetime_metachars_with_format_ids() {
    let values = metachars("datetime.ppt");
    let date_times = values
        .iter()
        .filter(|value| value.kind() == PowerPointMetacharKind::DateTime)
        .collect::<Vec<_>>();
    assert_eq!(date_times.len(), 13);
    assert!(date_times.iter().all(|value| {
        value
            .datetime_format()
            .is_some_and(|format| format.get() <= 12)
    }));
    assert!(values.iter().all(|value| value.rtf_format().is_none()));
}

#[test]
fn single_placeholder_presentations_parse_cleanly() {
    let values = metachars("incorrect_slide_order.ppt");
    let kinds = values.iter().map(|value| value.kind()).collect::<Vec<_>>();
    assert_eq!(count(&kinds, PowerPointMetacharKind::SlideNumber), 1);
    assert_eq!(count(&kinds, PowerPointMetacharKind::Footer), 1);
    assert_eq!(count(&kinds, PowerPointMetacharKind::GenericDate), 1);
}

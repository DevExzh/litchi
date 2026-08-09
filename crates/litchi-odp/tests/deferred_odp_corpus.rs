#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::Presentation;
use std::path::{Path, PathBuf};

fn corpus() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/corpus")
}

#[test]
fn three_impress_packages_open_and_expose_slides() {
    for (file, minimum) in [
        ("impress-basic.odp", 1),
        ("impress-embedded-spreadsheet.odp", 1),
        ("impress-master-layouts.odp", 5),
    ] {
        let presentation = Presentation::open(corpus().join(file)).unwrap();
        assert!(presentation.slides().unwrap().len() >= minimum, "{file}");
        presentation.declarations().unwrap();
        presentation.layouts().unwrap();
        presentation.pages().unwrap();
        presentation.metadata().unwrap();
    }

    let presentation =
        Presentation::open(corpus().join("impress-embedded-spreadsheet.odp")).unwrap();
    assert!(
        presentation
            .slides()
            .unwrap()
            .iter()
            .flat_map(|slide| slide.shapes().unwrap())
            .next()
            .is_some()
    );
}

//! Slide-library metadata must stay out of public presentation text views.
#![cfg(feature = "ppt")]
use litchi::ppt::{
    Package,
    slide_sync::{Synchronization, SystemTime},
};
use litchi_core::Position;
use litchi_ops::{OpenOptions, SourceMode, view::Mode};
use std::path::Path;

#[test]
fn native_ppt_sync_metadata_is_invisible_in_operations_and_facade_text_views() {
    for name in [
        "powerpoint-visual-objects.ppt",
        "powerpoint-visual-objects.pps",
        "powerpoint-visual-objects.pot",
        "powerpoint-passive-macro.ppt",
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name);
        let bytes = std::fs::read(&path).unwrap();
        let source = Package::from_bytes(&bytes).unwrap();
        let time = SystemTime::new(2026, 11, 6, 7, 12, 34, 56, 789).unwrap();
        let mut e = source.edit_slide_sync(Position::new(0)).unwrap();
        e.set(
            Synchronization::new(
                "INERT_METADATA_SERVER_074",
                "http://example.invalid/INERT_METADATA_URL_074",
                time,
                time,
            )
            .unwrap(),
        )
        .unwrap();
        let c = e.commit().unwrap();
        let output = c.package().to_bytes().unwrap();
        let before = litchi::open_bytes(bytes.clone()).unwrap();
        let after = litchi::open_bytes(output.clone()).unwrap();
        assert_eq!(after.text().unwrap(), before.text().unwrap());
        let slides = |f: &litchi::File| {
            f.presentation()
                .unwrap()
                .slides()
                .unwrap()
                .into_iter()
                .map(|slide| slide.text().unwrap())
                .collect::<Vec<_>>()
        };
        assert_eq!(slides(&after), slides(&before));
        let file = tempfile::NamedTempFile::with_suffix(".ppt").unwrap();
        c.package().save(file.path()).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let options = OpenOptions::new().source_mode(mode);
            let original = litchi_ops::open_bytes_with(bytes.clone(), &options).unwrap();
            let current = if mode == SourceMode::Auto {
                litchi_ops::open_with(file.path(), &options).unwrap()
            } else {
                litchi_ops::open_bytes_with(output.clone(), &options).unwrap()
            };
            assert_eq!(
                current.view(Mode::Text).unwrap(),
                original.view(Mode::Text).unwrap()
            );
            assert_eq!(
                serde_json::to_value(current.view_json(Mode::Text).unwrap()).unwrap(),
                serde_json::to_value(original.view_json(Mode::Text).unwrap()).unwrap()
            );
            assert_eq!(
                current.presentation().unwrap().slides().unwrap()[0]
                    .text()
                    .unwrap(),
                original.presentation().unwrap().slides().unwrap()[0]
                    .text()
                    .unwrap()
            );
            assert_eq!(current.to_bytes().unwrap(), output);
        }
        assert_eq!(source.to_bytes().unwrap(), bytes);
        assert_eq!(std::fs::read(path).unwrap(), bytes);
    }
}

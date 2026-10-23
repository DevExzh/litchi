//! ADR 0034 W3 evidence for the immutable ODC `Package` owner: atomic save,
//! typed caller-selected limits, deterministic fresh creation and exact no-op
//! publication.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odc::{
    AxisSpec, AxisUpdate, Builder, ChartClass, Definition, Error, LimitKind, OpenOptions, Package,
    SemanticLimitKind, SemanticLimits, SeriesSpec, SourceLimitKind, SourceLimits, chart::Dimension,
};
use std::path::{Path, PathBuf};

/// A fresh, empty scratch directory under the system temporary directory,
/// removed when dropped.
struct Scratch(PathBuf);

impl Scratch {
    fn new(name: &str) -> Self {
        let path = std::env::temp_dir().join(format!(
            "litchi-odc-package-owner-{name}-{}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&path);
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

fn definition(series: usize) -> Definition {
    let mut value = Definition::new(ChartClass::bar());
    let mut axis = AxisSpec::new(Dimension::X);
    axis.name = Some("categories".into());
    value.plot_area.axes.push(axis);
    for _ in 0..series {
        value.plot_area.series.push(SeriesSpec::default());
    }
    value
}

fn built(series: usize) -> Package {
    Builder::new()
        .with_definition(definition(series))
        .build()
        .unwrap()
}

#[test]
fn fresh_creation_is_deterministic_without_ambient_state() {
    let first = Package::new().unwrap();
    assert_eq!(first.as_bytes(), Package::new().unwrap().as_bytes());
    assert_eq!(built(2).as_bytes(), built(2).as_bytes());
    assert!(first.chart().plot_area().is_some());
    assert!(first.tolerance_report().is_clean());
}

#[test]
fn from_bytes_round_trips_the_exact_source() {
    let bytes = built(1).to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    assert_eq!(package.as_bytes(), bytes.as_slice());
    let streamed = Package::from_reader(bytes.as_slice()).unwrap();
    assert_eq!(streamed.as_bytes(), bytes.as_slice());
    let mut written = Vec::new();
    package.write_to(&mut written).unwrap();
    assert_eq!(written, bytes);
}

#[test]
fn an_exact_no_op_commit_is_the_source_byte_for_byte() {
    let bytes = built(1).to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    let commit = package.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), bytes.as_slice());
    let (published, patch) = commit.into_parts();
    assert_eq!(
        package.apply(&patch).unwrap().into_package().as_bytes(),
        published.as_bytes()
    );
}

#[test]
fn a_real_edit_changes_bytes_and_its_patch_reproduces_them() {
    let bytes = built(1).to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit();
    edit.update_axis(0, AxisUpdate::named("renamed")).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_ne!(commit.package().as_bytes(), bytes.as_slice());
    assert_eq!(
        package.as_bytes(),
        bytes.as_slice(),
        "the source never changes"
    );
    let (published, patch) = commit.into_parts();
    assert_eq!(
        published
            .chart()
            .plot_area()
            .and_then(|plot_area| plot_area.axes().next())
            .and_then(|axis| axis.name().map(str::to_owned))
            .as_deref(),
        Some("renamed")
    );
    let replayed = package.apply(&patch).unwrap();
    assert_eq!(replayed.into_package().as_bytes(), published.as_bytes());
    assert_eq!(
        published
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        bytes.as_slice()
    );
    assert!(
        published.apply(&patch).is_err(),
        "a stale source is refused"
    );
}

#[test]
fn save_atomically_replaces_the_destination_and_leaves_no_staging_file() {
    let scratch = Scratch::new("atomic");
    let path = scratch.path().join("out.odc");
    std::fs::write(&path, b"previous bytes").unwrap();
    let package = built(1);
    package.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 1);
    package
        .save_with_durability(&path, litchi_core::Durability::NoSync)
        .unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), package.as_bytes());
    assert_eq!(Package::open(&path).unwrap().as_bytes(), package.as_bytes());
}

#[cfg(unix)]
#[test]
fn a_save_through_a_symlink_is_refused_and_the_target_keeps_its_bytes() {
    let scratch = Scratch::new("symlink");
    let target = scratch.path().join("target.odc");
    std::fs::write(&target, b"untouched").unwrap();
    let link = scratch.path().join("link.odc");
    std::os::unix::fs::symlink(&target, &link).unwrap();
    assert!(built(1).save(&link).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), b"untouched");
    assert_eq!(std::fs::read_dir(scratch.path()).unwrap().count(), 2);
}

#[test]
fn caller_selected_source_limits_are_typed_at_every_ingress() {
    let bytes = built(1).to_bytes().unwrap();
    let length = u64::try_from(bytes.len()).unwrap();
    let exact = SourceLimits::builder()
        .max_input_bytes(length)
        .unwrap()
        .build()
        .unwrap();
    let opened =
        Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(exact)).unwrap();
    assert_eq!(opened.as_bytes(), bytes.as_slice());

    let below = SourceLimits::builder()
        .max_input_bytes(length - 1)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(below);
    for error in [
        Package::from_bytes_with_options(&bytes, options).unwrap_err(),
        Package::from_reader_with_options(bytes.as_slice(), options).unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                maximum,
                ..
            } if maximum == length - 1
        ));
    }
}

#[test]
fn semantic_limits_bound_content_and_series() {
    let bytes = built(2).to_bytes().unwrap();
    let content = SemanticLimits::builder()
        .max_content_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(content))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
            ..
        }
    ));

    let one = SemanticLimits::builder()
        .max_units(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(&bytes, OpenOptions::new().semantic_limits(one))
            .unwrap_err(),
        Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Units),
            observed: 2,
            maximum: 1,
        }
    ));
    let single = built(1).to_bytes().unwrap();
    assert!(
        Package::from_bytes_with_options(&single, OpenOptions::new().semantic_limits(one)).is_ok()
    );
}

#[test]
fn a_prepared_detection_result_is_adopted_and_other_families_are_refused() {
    let bytes = built(1).to_bytes().unwrap();
    let prepared = litchi_odf_common::detect::prepared(bytes.clone(), SourceLimits::default())
        .expect("ODC is detected");
    assert_eq!(
        Package::from_prepared_package(prepared).unwrap().as_bytes(),
        bytes.as_slice()
    );

    let mut text = litchi_odf_common::core::PackageWriter::new(
        litchi_odf_common::core::StreamingArchiveLimits::default(),
    );
    text.set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    text.add_file(
        "content.xml",
        br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.4"><office:body><office:text/></office:body></office:document-content>"#,
    )
    .unwrap();
    let text = text.finish_to_bytes().unwrap();
    let prepared = litchi_odf_common::detect::prepared(text, SourceLimits::default())
        .expect("ODT is detected");
    assert!(matches!(
        Package::from_prepared_package(prepared).unwrap_err(),
        Error::Core(litchi_core::Error::UnexpectedFormat { .. })
    ));
}

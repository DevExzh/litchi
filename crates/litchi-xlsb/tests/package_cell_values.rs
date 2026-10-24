#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "package-owner regressions use compact generated fixtures and immediate diagnostics"
)]

//! Regressions for the package-owned cell-value transaction.
//!
//! The worksheet codec remains available under [`cell_values::worksheet`],
//! while package publication owns the full OPC image and package policies.

use std::io::Cursor;

use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::{Reference, StyleIndex, TransferLimits, Value};
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::{Package, SourceLimits};

fn a1() -> Reference {
    Reference::new(0, 0).expect("A1 reference")
}

fn style_zero() -> StyleIndex {
    StyleIndex::new(0).expect("default style")
}

fn changed_commit(source: &Package) -> litchi_xlsb::cell_values::Commit {
    let mut edit = source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    edit.insert(a1(), style_zero(), Value::Number(42.0))
        .expect("insert numeric cell");
    edit.commit().expect("package cell-value commit")
}

fn long_inline_string() -> String {
    (0..32_000)
        .map(|index| char::from(b'a' + u8::try_from(index % 26).expect("alphabet offset")))
        .collect()
}

fn exact_input_limits(bytes: &[u8]) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len()).expect("package length"))
        .expect("input-byte limit")
        .build()
        .expect("source limits")
}

fn chart_sheet_package() -> Package {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .expect("chart fixture");
    let mut writer = Builder::new();
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .expect("chart sheet");
    writer.add_worksheet(builder::Worksheet::new("Data"));
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut output)
        .expect("write chart fixture");
    Package::from_bytes(&output.into_inner()).expect("read chart package")
}

#[test]
fn package_cell_name_selection_rejects_ambiguous_names() {
    let mut writer = Builder::new();
    writer.add_worksheet(builder::Worksheet::new("DATA"));
    writer.add_worksheet(builder::Worksheet::new("Else"));
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut output)
        .expect("valid distinct names");
    let mut opc =
        OpcPackage::from_bytes(&output.into_inner(), SourceLimits::default()).expect("OPC");
    let uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let original = opc.get_part(&uri).expect("workbook part").blob().to_vec();
    let old: Vec<_> = "Else".encode_utf16().flat_map(u16::to_le_bytes).collect();
    let replacement: Vec<_> = "data".encode_utf16().flat_map(u16::to_le_bytes).collect();
    let mut rewritten = Vec::new();
    let mut records = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
    let mut replaced = false;
    for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = record.expect("record");
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BUNDLE_SH && payload.ends_with(&old) {
            let start = payload.len() - old.len();
            payload[start..].copy_from_slice(&replacement);
            replaced = true;
        }
        records
            .write_record(record.kind(), &payload)
            .expect("record");
    }
    assert!(replaced);
    opc.get_part_mut(&uri)
        .expect("workbook part")
        .set_blob(rewritten);
    // Even a raw OPC escape into Package must validate the catalog before
    // selecting a name; imported ambiguity must never choose the first tab.
    let package = Package::from(opc);
    assert!(
        package
            .edit_cell_values("DaTa")
            .expect("cell-value edit entry")
            .resolve()
            .is_err()
    );
    assert!(package.workbook().is_err());
}

#[test]
fn package_cell_commit_is_immutable_and_exposes_the_commit_triad() {
    let source = Package::new().expect("source package");
    let source_bytes = source.to_bytes().expect("source bytes");
    let commit = changed_commit(&source);
    let candidate_bytes = commit.package().to_bytes().expect("candidate bytes");

    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().before(), source_bytes.as_slice());
    assert_eq!(commit.patch().after(), candidate_bytes.as_slice());
    assert_eq!(
        source.to_bytes().expect("source remains unchanged"),
        source_bytes
    );

    let applied = source
        .apply_cell_values(commit.patch())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect("replay against the exact source");
    assert_eq!(applied.to_bytes().expect("replayed bytes"), candidate_bytes);

    let (package, patch) = commit.into_parts();
    assert_eq!(package.to_bytes().expect("consumed package"), patch.after());
}

#[test]
fn package_cell_patch_inverse_replays_to_the_original_image() {
    let source = Package::new().expect("source package");
    let original = source.to_bytes().expect("original bytes");
    let commit = changed_commit(&source);
    let changed = source
        .apply_cell_values(commit.patch())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect("forward package patch");

    let restored = changed
        .apply_cell_values(&commit.patch().inverse())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect("inverse package patch");
    assert_eq!(restored.to_bytes().expect("restored bytes"), original);
}

#[test]
fn package_cell_patch_rejects_stale_sources_including_noop_patches() {
    let source = Package::new().expect("source package");
    let changed = changed_commit(&source);
    let changed_package = source
        .apply_cell_values(changed.patch())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect("changed package");

    let error = changed_package
        .apply_cell_values(changed.patch())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect_err("a changed patch must reject its already-published destination");
    assert!(matches!(
        error,
        litchi_xlsb::package::PackageError::PatchConflict {
            part: "cell values"
        }
    ));

    let noop = source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet")
        .commit()
        .expect("no-op package commit");
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    let error = changed_package
        .apply_cell_values(noop.patch())
        .map(litchi_xlsb::cell_values::Commit::into_package)
        .expect_err("even a no-op patch is source-checked against its exact package image");
    assert!(matches!(
        error,
        litchi_xlsb::package::PackageError::PatchConflict {
            part: "cell values"
        }
    ));
}

#[test]
fn package_cell_noop_preserves_exact_bytes() {
    let source = Package::new().expect("source package");
    let original = source.to_bytes().expect("original bytes");
    let noop = source
        .edit_cell_values("Sheet1")
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet")
        .commit()
        .expect("no-op package commit");

    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(), original.as_slice());
    assert_eq!(noop.patch().after(), original.as_slice());
    assert_eq!(
        noop.into_package().to_bytes().expect("no-op bytes"),
        original
    );
}

#[test]
fn package_cell_commit_rejects_invalid_style_in_context() {
    let source = Package::new().expect("source package");
    let original = source.to_bytes().expect("original bytes");
    let mut edit = source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    edit.insert(
        a1(),
        StyleIndex::new(0x00FF_FFFF).expect("wire style index"),
        Value::Number(1.0),
    )
    .expect("stage invalid style");

    assert!(
        edit.commit().is_err(),
        "publication must validate the style index against the package Styles part"
    );
    assert_eq!(
        source.to_bytes().expect("source bytes after refusal"),
        original
    );
}

#[test]
fn package_cell_commit_and_replay_honor_finite_retained_input_limits() {
    let source = Package::new().expect("source package");
    let original = source.to_bytes().expect("original bytes");

    let mut edit = source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    edit.insert(
        a1(),
        style_zero(),
        Value::InlineString(long_inline_string()),
    )
    .expect("stage growing cell");
    let commit = edit.commit().expect("default policy permits growth");
    assert!(
        commit.patch().after().len() > commit.patch().before().len(),
        "the fixture operation must produce a larger package image"
    );

    let bounded_source = Package::from_bytes_with_options(
        &original,
        litchi_xlsb::OpenOptions::new().source_limits(exact_input_limits(&original)),
    )
    .expect("reopen source at its exact input budget");
    assert!(
        bounded_source
            .apply_cell_values(commit.patch())
            .map(litchi_xlsb::cell_values::Commit::into_package)
            .is_err(),
        "destination retained input limits must reject a patch that grows the package"
    );

    let mut bounded_edit = bounded_source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    bounded_edit
        .insert(
            a1(),
            style_zero(),
            Value::InlineString(long_inline_string()),
        )
        .expect("stage growing cell under bounded source");
    assert!(
        bounded_edit.commit().is_err(),
        "the source owner's retained input budget must fail before publication"
    );
}

#[test]
fn package_cell_selector_uses_the_complete_catalog() {
    let package = chart_sheet_package();

    assert!(
        package
            .edit_cell_values("missing")
            .expect("missing selector")
            .resolve()
            .expect("missing worksheet resolution")
            .is_none()
    );
    assert!(
        package
            .edit_cell_values("sales chart")
            .expect("chart edit entry")
            .resolve()
            .is_err(),
        "a chart tab selected by case-insensitive name is not a worksheet"
    );
    assert!(
        package
            .edit_cell_values(0_usize)
            .expect("chart edit entry")
            .resolve()
            .is_err(),
        "the chart tab's complete-catalog position is not a worksheet"
    );
    assert!(
        package
            .edit_cell_values(1_usize)
            .expect("worksheet selector")
            .resolve()
            .expect("worksheet resolution")
            .is_some()
    );
    assert!(
        package
            .edit_cell_values("Data")
            .expect("worksheet selector")
            .resolve()
            .expect("worksheet resolution")
            .is_some()
    );
}

#[test]
fn raw_worksheet_transaction_stays_under_an_explicit_namespace() {
    let source = Package::new().expect("source package");
    let workbook = source.workbook().expect("workbook projection");
    let snapshot = workbook.cell_values(0).expect("worksheet snapshot");
    let mut draft: litchi_xlsb::cell_values::worksheet::Draft = snapshot.draft();
    draft
        .insert(a1(), style_zero(), Value::Number(7.0))
        .expect("raw worksheet insert");
    let (_, patch): (_, litchi_xlsb::cell_values::worksheet::Patch) =
        draft.finish().expect("raw worksheet draft");
    assert!(!patch.is_noop());
    assert_eq!(patch.before(), snapshot.source_bytes());
}

#[test]
fn package_cell_edit_accepts_explicit_worksheet_and_transfer_limits() {
    let source = Package::new().expect("source package");
    let edit = source
        .edit_cell_values_with_options(0_usize, litchi_xlsb::cell_values::EditOptions::new())
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    assert!(edit.commit().expect("bounded no-op").patch().is_noop());
}

#[test]
fn package_cell_transfer_quotas_cover_images_and_changes() {
    use litchi_xlsb::cell_values::Error;
    let source = Package::new().expect("source");
    let original = source.to_bytes().expect("bytes");
    let quota = |bytes, changes| TransferLimits::new(bytes, changes, 1, bytes);
    let edit = source
        .edit_cell_values_with_options(
            0,
            litchi_xlsb::cell_values::EditOptions::new().transfer_limits(quota(1, 1)),
        )
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    assert!(matches!(
        edit.commit(),
        Err(Error::Limit {
            kind: litchi_xlsb::package::LimitKind::CellValues(
                litchi_xlsb::cell_values::LimitKind::PatchBytes
            ),
            ..
        })
    ));
    let image_bytes = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet")
        .commit()
        .expect("baseline patch")
        .patch()
        .retained_bytes();
    assert!(
        image_bytes > original.len() * 2,
        "quota includes worksheet proof and metadata"
    );
    let edit = source
        .edit_cell_values_with_options(
            0,
            litchi_xlsb::cell_values::EditOptions::new().transfer_limits(quota(image_bytes - 1, 1)),
        )
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    assert!(matches!(
        edit.commit(),
        Err(Error::Limit {
            kind: litchi_xlsb::package::LimitKind::CellValues(
                litchi_xlsb::cell_values::LimitKind::PatchBytes
            ),
            ..
        })
    ));
    let edit = source
        .edit_cell_values_with_options(
            0,
            litchi_xlsb::cell_values::EditOptions::new().transfer_limits(quota(image_bytes, 1)),
        )
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    assert!(
        edit.commit()
            .expect("exact retained quota")
            .patch()
            .is_noop()
    );
    let mut edit = source
        .edit_cell_values_with_options(
            0,
            litchi_xlsb::cell_values::EditOptions::new().transfer_limits(quota(image_bytes * 2, 1)),
        )
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("worksheet");
    for column in [0, 1] {
        edit.insert(
            Reference::new(0, column).expect("reference"),
            style_zero(),
            Value::Number(1.0),
        )
        .expect("insert");
    }
    assert!(matches!(
        edit.commit(),
        Err(Error::Limit {
            kind: litchi_xlsb::package::LimitKind::CellValues(
                litchi_xlsb::cell_values::LimitKind::Changes
            ),
            ..
        })
    ));
    assert_eq!(source.to_bytes().expect("source unchanged"), original);
}

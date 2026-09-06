//! Source-preserving Numbers chart Arrange coverage.
//!
//! The fixture keeps the rooted sheet ownership list in `Document.iwa` and
//! the chart graph in `CalculationEngine.iwa`, matching the cross-component
//! profile emitted by Numbers.  Public calls use only semantic sheet/chart
//! selectors and [`ChartArrangement`]; native identifiers are confined to the
//! test oracle that checks graph locality and malformed input.

use std::{fmt::Debug, io};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_numbers::{
    ChartArrangement, ChartSelector, MAX_OBJECTS, Package, PackageReadOptions,
    PackageSemanticLimits, SheetChartArrangementError, SheetSelector,
};

#[path = "support/chart_arrangement_fixture.rs"]
mod fixture;

use fixture::{
    CALCULATION_MEMBER, CHARTS, DOCUMENT_MEMBER, SENTINEL_MEMBER, SHEETS, STYLESHEET_MEMBER,
    TestResult, UNRELATED_MEMBER,
};

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn arrangement(package: &Package, sheet: usize, chart: usize) -> TestResult<ChartArrangement> {
    Ok(
        package
            .sheet_chart_arrangement(SheetSelector::index(sheet), ChartSelector::index(chart))?,
    )
}

fn commit_arrangement(
    package: &Package,
    sheet: usize,
    chart: usize,
    value: ChartArrangement,
) -> TestResult<litchi_numbers::SheetChartArrangementCommit> {
    Ok(package
        .edit_sheet_chart_arrangement(SheetSelector::index(sheet), ChartSelector::index(chart))?
        .set(value)
        .commit()?)
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("Arrange removed a package member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact ZIP local record",
                entry.name()
            );
            assert_eq!(entry.metadata(), candidate.metadata());
        }
    }
    changed.sort_unstable();
    assert_eq!(changed, [CALCULATION_MEMBER.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_error_is_redacted<E: Debug + std::fmt::Display>(error: &E) {
    let text = format!("{error:?} {error}");
    for identifier in CHARTS
        .into_iter()
        .chain(SHEETS)
        .chain([fixture::CALCULATION_ENGINE_ID, fixture::FORMULA_OWNER_ID])
    {
        assert!(
            !text.contains(&identifier.to_string()),
            "native chart identifier leaked from semantic error: {text}"
        );
    }
}

#[test]
fn chart_arrangement_public_handles_are_typed_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<ChartSelector>();
    assert_send_sync_debug::<SheetSelector<'static>>();
    assert_send_sync_debug::<ChartArrangement>();
    assert_send_sync_debug::<litchi_numbers::SheetChartArrangementCommit>();
    assert_send_sync_debug::<litchi_numbers::SheetChartArrangementEdit<'static>>();
    assert_send_sync_debug::<litchi_numbers::SheetChartArrangementPatch>();
    assert_send_sync_debug::<litchi_numbers::SheetChartArrangementDiagnostics>();
    assert_send_sync_debug::<SheetChartArrangementError>();

    let package = Package::from_bytes(&fixture::fixture()?)?;
    let before = exact_bytes(&package)?;
    let error = package
        .sheet_chart_arrangement(SheetSelector::name("missing"), ChartSelector::index(0))
        .expect_err("missing semantic sheet must fail");
    assert_error_is_redacted(&error);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn chart_arrangement_read_and_exact_noop_cover_physical_boolean_presence() -> TestResult {
    let states = [
        (None, None),
        (Some(false), None),
        (None, Some(false)),
        (Some(false), Some(false)),
        (Some(false), Some(true)),
        (Some(true), None),
        (Some(true), Some(false)),
        (None, Some(true)),
        (Some(true), Some(true)),
    ];
    for state in states {
        let source = fixture::fixture_with_states([state, (None, None)], true)?;
        let package = Package::from_bytes(&source)?;
        let expected = ChartArrangement::new(state.0.unwrap_or(false), state.1.unwrap_or(false));
        assert_eq!(arrangement(&package, 0, 0)?, expected);
        let commit = commit_arrangement(&package, 0, 0, expected)?;
        assert!(commit.patch().is_noop());
        assert_eq!(exact_bytes(commit.package())?, source);
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
    }
    Ok(())
}

#[test]
fn chart_arrangement_change_is_local_preserves_unknowns_and_keeps_other_chart() -> TestResult {
    let source = fixture::fixture_with_states(
        [(Some(false), Some(false)), (Some(true), Some(false))],
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let target = ChartArrangement::new(true, true);
    let commit = commit_arrangement(&package, 0, 0, target)?;
    let target_bytes = exact_bytes(commit.package())?;

    assert_eq!(arrangement(commit.package(), 0, 0)?, target);
    assert_eq!(
        arrangement(commit.package(), 1, 0)?,
        ChartArrangement::new(true, false)
    );
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_locality(&source, &target_bytes)?;
    assert_eq!(
        fixture::chart_payload_from_source(&source, 1)?,
        fixture::chart_payload_from_source(&target_bytes, 1)?,
        "an Arrange edit must not rewrite the unselected chart graph"
    );
    assert_eq!(
        fixture::raw_drawable_fields(&source, 0, fixture::UNKNOWN_DRAWABLE_FIELD)?,
        fixture::raw_drawable_fields(&target_bytes, 0, fixture::UNKNOWN_DRAWABLE_FIELD)?,
        "unknown selected-drawable fields must survive the lazy rewrite"
    );
    assert_eq!(
        fixture::raw_drawable_fields(&source, 0, fixture::UNKNOWN_DRAWABLE_BYTES_FIELD)?,
        fixture::raw_drawable_fields(&target_bytes, 0, fixture::UNKNOWN_DRAWABLE_BYTES_FIELD)?,
    );
    let reopened = Package::from_bytes(&target_bytes)?;
    assert_eq!(arrangement(&reopened, 0, 0)?, target);
    assert_eq!(exact_bytes(&reopened)?, target_bytes);
    Ok(())
}

#[test]
fn chart_arrangement_inverse_and_exact_apply_are_source_bound() -> TestResult {
    let source =
        fixture::fixture_with_states([(None, Some(false)), (Some(false), Some(true))], true)?;
    let package = Package::from_bytes(&source)?;
    let target = ChartArrangement::new(true, true);
    let commit = commit_arrangement(&package, 0, 0, target)?;
    let target_bytes = exact_bytes(commit.package())?;

    let applied = package.apply_sheet_chart_arrangement(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target_bytes);

    let restored = commit
        .package()
        .apply_sheet_chart_arrangement(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        arrangement(&restored.package(), 0, 0)?,
        ChartArrangement::default()
    );

    let stale = commit
        .package()
        .apply_sheet_chart_arrangement(commit.patch())
        .expect_err("a forward patch cannot apply to its own target");
    assert!(matches!(stale, SheetChartArrangementError::PatchConflict));
    assert_eq!(exact_bytes(commit.package())?, target_bytes);

    let foreign = Package::from_bytes(&fixture::fixture_with_states(
        [(None, Some(false)), (Some(false), Some(true))],
        false,
    )?)?;
    let foreign_before = exact_bytes(&foreign)?;
    let conflict = foreign
        .apply_sheet_chart_arrangement(commit.patch())
        .expect_err("a patch must reject a distinct source snapshot");
    assert!(matches!(
        conflict,
        SheetChartArrangementError::PatchConflict
    ));
    assert_eq!(exact_bytes(&foreign)?, foreign_before);
    Ok(())
}

#[test]
fn chart_arrangement_rejects_malformed_ownership_parent_and_message_graphs_atomically() -> TestResult
{
    let source = fixture::fixture()?;
    let malformed = [
        (
            "duplicate sheet owner",
            fixture::with_duplicate_sheet_owner(&source)?,
        ),
        (
            "duplicate referenced drawable",
            fixture::with_duplicate_chart_reference(&source)?,
        ),
        ("parent mismatch", fixture::with_parent_mismatch(&source)?),
        (
            "wrong chart message type",
            fixture::with_wrong_chart_message_type(&source)?,
        ),
        (
            "duplicate chart message",
            fixture::with_duplicate_chart_message(&source)?,
        ),
    ];
    for (label, bytes) in malformed {
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let before = exact_bytes(&package)?;
        let error = package
            .sheet_chart_arrangement(SheetSelector::index(0), ChartSelector::index(0))
            .expect_err(label);
        assert_error_is_redacted(&error);
        assert_eq!(exact_bytes(&package)?, before, "{label} mutated its source");
    }
    Ok(())
}

#[test]
fn chart_arrangement_refuses_finite_reference_budget_without_mutation() -> TestResult {
    let source = fixture::with_single_rooted_sheet(&fixture::fixture()?)?;
    let semantic = PackageSemanticLimits::new(
        MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        3,
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(Limits::default(), semantic),
    )?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_sheet_chart_arrangement(SheetSelector::index(0), ChartSelector::index(0))
        .expect("the first source selection fits the finite budget")
        .set(ChartArrangement::new(true, true))
        .commit()
        .expect_err("the source-plus-candidate transaction must honor the finite budget");
    assert_error_is_redacted(&error);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn chart_arrangement_fixture_has_cross_component_profile_and_untouched_assets() -> TestResult {
    let source = fixture::fixture()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.sheets().len(), 2);
    assert_eq!(package.sheets()[0].name(), "Sheet 1");
    assert_eq!(package.sheets()[1].name(), "Sheet 2");
    assert_eq!(
        fixture::object_message(
            &source,
            DOCUMENT_MEMBER,
            fixture::DOCUMENT_ID,
            fixture::DOCUMENT_MESSAGE_TYPE
        )?
        .len()
            > 0,
        true
    );
    let catalog = Catalog::from_bytes(&source)?;
    for member in [
        DOCUMENT_MEMBER,
        CALCULATION_MEMBER,
        STYLESHEET_MEMBER,
        fixture::METADATA_MEMBER,
        UNRELATED_MEMBER,
        SENTINEL_MEMBER,
    ] {
        assert!(
            catalog.iter().any(|entry| entry.name() == member),
            "missing {member}"
        );
    }
    assert_eq!(
        catalog
            .iter()
            .find(|entry| entry.name() == SENTINEL_MEMBER)
            .map(|entry| entry.data()),
        Some(b"unrelated ZIP sentinel".as_slice())
    );
    Ok(())
}

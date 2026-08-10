//! Native integration coverage for exact-source Numbers table-header edits.

use std::{io, sync::Arc};

use litchi_iwa_archive::package::Catalog;
use litchi_numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    table::{
        headers::{
            Count, Settings,
            transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
        },
        lock::State as LockState,
    },
};

#[path = "support/table_headers_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

fn fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers")
}

fn changed_settings(base: Settings) -> Settings {
    Settings {
        header_rows: base.header_rows,
        header_columns: base.header_columns,
        footer_rows: base.footer_rows,
        header_rows_frozen: Some(!base.header_rows_are_frozen()),
        header_columns_frozen: Some(!base.header_columns_are_frozen()),
        repeating_header_rows_enabled: base.repeating_header_rows_enabled,
        repeating_header_columns_enabled: base.repeating_header_columns_enabled,
    }
}

#[test]
fn native_selectors_presence_noop_change_apply_inverse_and_locality() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let by_index = package.table_header_settings(0usize, 0usize)?;
    let by_name = package.table_header_settings(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
    )?;
    assert_eq!(by_name, by_index);

    let noop_edit = package.edit_table_headers(0usize, 0usize)?;
    assert_eq!(noop_edit.settings(), by_index);
    let noop = noop_edit.set(by_index).commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().exact_bytes(), source);

    let changed = changed_settings(by_index);
    let commit = package
        .edit_table_headers(0usize, 0usize)?
        .set(changed)
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert_eq!(commit.patch().path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(commit.patch().before(), by_index);
    assert_eq!(commit.patch().after(), changed);
    assert_eq!(
        commit.package().table_header_settings(0usize, 0usize)?,
        changed
    );

    let target = commit.package().exact_bytes();
    let source_catalog = Catalog::from_bytes(&source)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    for (before, after) in source_catalog.iter().zip(target_catalog.iter()) {
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record()
            );
        }
    }
    assert_eq!(
        package
            .apply_table_headers(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened
            .apply_table_headers(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    assert!(matches!(
        reopened.apply_table_headers(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let mut lock = package.edit_table_lock(0usize, 0usize)?;
    lock.lock();
    let cross_source = lock.commit()?.into_package();
    assert!(matches!(
        cross_source.apply_table_headers(commit.patch()),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn selector_bounds_and_public_values_are_checked_and_thread_safe() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    assert!(matches!(
        package.table_header_settings("missing", 0usize),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_header_settings(0usize, "missing"),
        Err(Error::TableNotFound)
    ));
    assert_eq!(Count::new(1)?, Count::ONE);
    assert!(Count::new(0).is_err());
    assert!(Count::new(6).is_err());

    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Package>();
    assert_send_sync::<Edit<'static>>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<Diagnostics>();
    assert_send_sync::<Error>();
    assert_send_sync::<LimitKind>();
    assert_send_sync::<Settings>();
    assert_send_sync::<Count>();
    assert_send_sync::<Arc<[u8]>>();
    Ok(())
}

#[test]
fn native_manager_backed_count_change_refuses_without_publication() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let base = package.table_header_settings(0usize, 0usize)?;
    let requested = Settings {
        header_rows: Some(if base.header_rows == Some(Count::ONE) {
            Count::TWO
        } else {
            Count::ONE
        }),
        ..base
    };
    assert_ne!(requested, base);
    let error = package
        .edit_table_headers(0usize, 0usize)?
        .set(requested)
        .commit()
        .expect_err("rooted HeaderNameMgr count dependency must refuse publication");
    assert!(matches!(error, Error::UnsupportedDependency { .. }));
    assert_eq!(package.exact_bytes(), source);
    Ok(())
}

#[test]
fn locked_table_allows_exact_header_noop_and_refuses_change() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let mut lock = package.edit_table_lock(0usize, 0usize)?;
    lock.lock();
    let locked = lock.commit()?.into_package();
    assert_eq!(locked.table_lock(0usize, 0usize)?, LockState::Locked);
    let base = locked.table_header_settings(0usize, 0usize)?;
    let noop = locked
        .edit_table_headers(0usize, 0usize)?
        .set(base)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), locked.exact_bytes());
    let error = locked
        .edit_table_headers(0usize, 0usize)?
        .set(changed_settings(base))
        .commit()
        .expect_err("locked table must refuse a changed header transaction");
    assert!(matches!(
        error,
        Error::TableLocked {
            path: Path::Table { sheet: 0, table: 0 }
        }
    ));
    Ok(())
}

#[test]
fn synthetic_bounds_and_dependency_role_aliases_fail_atomically() -> TestResult<()> {
    let base_bytes = fixture::synthetic_package()?;
    let package = Package::from_bytes(&base_bytes)?;
    let base = package.table_header_settings(0usize, 0usize)?;

    let invalid_bounds = Settings {
        header_rows: Some(Count::ONE),
        footer_rows: Some(Count::ONE),
        ..base
    };
    assert!(matches!(
        package
            .edit_table_headers(0usize, 0usize)?
            .set(invalid_bounds)
            .commit(),
        Err(Error::InvalidSettings {
            path: Path::Table { sheet: 0, table: 0 },
            ..
        })
    ));
    assert_eq!(package.exact_bytes(), base_bytes);

    let changed = Settings {
        header_rows_frozen: Some(true),
        ..base
    };
    let pivot_self = fixture::rewrite_tables(&base_bytes, |archive| {
        let model = archive
            .object_mut(fixture::TABLE_MODEL)
            .ok_or_else(|| io::Error::other("missing synthetic table model"))?;
        fixture::append_reference_field(
            model,
            fixture::TABLE_MODEL_MESSAGE_TYPE,
            85,
            fixture::TABLE_MODEL,
            true,
        )
    })?;
    assert_invalid_change_preserves(&pivot_self, changed)?;

    let category_late_alias = fixture::rewrite_tables(&base_bytes, |archive| {
        let model = archive
            .object_mut(fixture::TABLE_MODEL)
            .ok_or_else(|| io::Error::other("missing synthetic table model"))?;
        fixture::append_reference_field(
            model,
            fixture::TABLE_MODEL_MESSAGE_TYPE,
            86,
            fixture::CATEGORY_OWNER,
            true,
        )?;
        fixture::push_object(
            archive,
            fixture::category_owner_object(&[fixture::GROUP_BY, fixture::TABLE_MODEL])?,
        );
        fixture::push_object(archive, fixture::enabled_group_object()?);
        Ok(())
    })?;
    assert_invalid_change_preserves(&category_late_alias, changed)?;

    for field in [4, 5, 15, 17] {
        let aliased = fixture::rewrite_tables(&base_bytes, |archive| {
            let info = archive
                .object_mut(fixture::TABLE_INFO)
                .ok_or_else(|| io::Error::other("missing synthetic table info"))?;
            fixture::append_reference_field(
                info,
                fixture::TABLE_INFO_MESSAGE_TYPE,
                field,
                fixture::TABLE_MODEL,
                false,
            )
        })?;
        assert_invalid_change_preserves(&aliased, changed)?;
    }
    Ok(())
}

#[test]
fn inverse_apply_preflights_the_retained_target_work_before_reopen() -> TestResult<()> {
    let source = fixture::synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let before = package.table_header_settings(0usize, 0usize)?;
    let after = Settings {
        header_rows_frozen: Some(true),
        ..before
    };
    let commit = package
        .edit_table_headers(0usize, 0usize)?
        .set(after)
        .commit()?;
    let candidate = commit.package().exact_bytes();
    assert!(
        candidate.len() < source.len(),
        "changed output deletes previews"
    );
    let inverse = commit.patch().inverse();
    let observed = fixture::transaction_work_precharge(&candidate, &source)?;
    let maximum = observed
        .checked_sub(1)
        .ok_or_else(|| io::Error::other("transaction work must be non-zero"))?;
    let archive_limits = PackageLimits::new(
        PackageLimits::MAX_INPUT_BYTES,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        u64::try_from(maximum)?,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    let restricted = Package::from_bytes_with_options(
        &candidate,
        PackageReadOptions::new(archive_limits, PackageSemanticLimits::default()),
    )?;
    let unchanged = restricted.exact_bytes();
    let error = restricted
        .apply_table_headers(&inverse)
        .expect_err("retained inverse target must be included in transaction work");
    assert!(matches!(
        error,
        Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: actual,
            maximum: limit,
            path: Path::Package,
        } if actual == u64::try_from(observed)? && limit == u64::try_from(maximum)?
    ));
    assert_eq!(restricted.exact_bytes(), unchanged);
    Ok(())
}

fn assert_invalid_change_preserves(bytes: &[u8], settings: Settings) -> TestResult<()> {
    let package = Package::from_bytes(bytes)?;
    let source = package.exact_bytes();
    assert!(matches!(
        package
            .edit_table_headers(0usize, 0usize)?
            .set(settings)
            .commit(),
        Err(Error::InvalidSource { .. }) | Err(Error::UnsupportedDependency { .. })
    ));
    assert_eq!(package.exact_bytes(), source);
    Ok(())
}

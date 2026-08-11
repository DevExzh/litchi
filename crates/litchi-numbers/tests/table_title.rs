//! Native integration coverage for exact-source Numbers table-title edits.

use std::{fmt::Debug, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_numbers::table::title::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path, Settings,
};
use litchi_numbers::{Package, SheetSelector, TableSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const FIXTURE_MARKER: &str = "Litchi native Numbers fixture";
const CANONICAL_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

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

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn settings_preserve_each_optional_boolean_presence_state() {
    for visible in [None, Some(false), Some(true)] {
        for outlined in [None, Some(false), Some(true)] {
            let settings = Settings::new(visible, outlined);
            assert_eq!(settings.visible(), visible);
            assert_eq!(settings.outlined(), outlined);
            assert_eq!(settings.is_visible(), visible == Some(true));
            assert_eq!(settings.is_outlined(), outlined == Some(true));
        }
    }
    assert_eq!(Settings::default(), Settings::new(None, None));
}

#[test]
fn native_read_noop_apply_and_selectors_are_exact() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let by_index = package.table_title_settings(0usize, 0usize)?;
    let by_name = package.table_title_settings(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
    )?;
    assert_eq!(by_name, by_index);
    assert_eq!(by_index.visible(), Some(true));
    assert_eq!(by_index.outlined(), None);

    let edit = package.edit_table_title(0usize, 0usize)?;
    assert_eq!(edit.path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(edit.settings(), by_index);
    let commit = edit.set(by_index).commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), by_index);
    assert_eq!(commit.patch().after(), by_index);
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.package().exact_bytes(), source);

    let applied = package.apply_table_title(commit.patch())?;
    assert!(applied.patch().is_noop());
    assert_eq!(applied.package().exact_bytes(), source);
    let restored = applied
        .package()
        .apply_table_title(&applied.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);

    assert!(matches!(
        package.table_title_settings("missing", 0usize),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_title_settings(0usize, "missing"),
        Err(Error::TableNotFound)
    ));
    Ok(())
}

#[test]
fn native_hide_is_local_reversible_and_exact_source_bound() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let before = package.table_title_settings(0usize, 0usize)?;
    assert_eq!(before, Settings::new(Some(true), None));
    let hidden = Settings::new(Some(false), before.outlined());

    let commit = package
        .edit_table_title(0usize, 0usize)?
        .set(hidden)
        .commit()?;
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(commit.patch().before(), before);
    assert_eq!(commit.patch().after(), hidden);
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit.package().table_title_settings(0usize, 0usize)?,
        hidden
    );
    assert_eq!(
        commit
            .package()
            .table_title_settings("Sheet 1", "Table 1")?,
        hidden
    );

    let target = commit.package().exact_bytes();
    assert_ne!(target, source);
    assert_exact_locality(&source, &target)?;

    let applied = package.apply_table_title(commit.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let reopened = Package::from_bytes(&target)?;
    assert!(matches!(
        reopened.apply_table_title(commit.patch()),
        Err(Error::PatchConflict)
    ));

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), hidden);
    assert_eq!(inverse.after(), before);
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = reopened.apply_table_title(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored.package().table_title_settings(0usize, 0usize)?,
        before
    );
    Ok(())
}

#[test]
fn native_absence_and_explicit_false_are_distinct_changes() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source = package.exact_bytes();
    let before = package.table_title_settings(0usize, 0usize)?;
    assert_eq!(before.outlined(), None);
    let explicit_false = Settings::new(before.visible(), Some(false));

    let commit = package
        .edit_table_title(0usize, 0usize)?
        .set(explicit_false)
        .commit()?;
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.package().table_title_settings(0usize, 0usize)?,
        explicit_false
    );
    assert_eq!(
        commit
            .package()
            .apply_table_title(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult<()> {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    let mut changed_members = Vec::new();

    for before in source.iter() {
        let after = target
            .iter()
            .find(|candidate| candidate.name() == before.name());
        if CANONICAL_PREVIEWS.contains(&before.name()) {
            assert!(after.is_none(), "preview {} was retained", before.name());
            continue;
        }
        let after = after.ok_or_else(|| {
            std::io::Error::other(format!("member {} was unexpectedly deleted", before.name()))
        })?;
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
                "unchanged member {} lost its exact local record",
                before.name()
            );
        } else {
            changed_members.push(before.name());
        }
    }
    assert_eq!(changed_members, ["Index/CalculationEngine.iwa"]);
    assert_eq!(target.len() + CANONICAL_PREVIEWS.len(), source.len());
    Ok(())
}

#[test]
fn public_transaction_values_are_thread_safe_and_debug_redacted() -> TestResult<()> {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Settings>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::open(fixture_path())?;
    let edit = package.edit_table_title(0usize, 0usize)?;
    let edit_debug = format!("{edit:?}");
    assert!(edit_debug.contains("Table"));
    assert!(!edit_debug.contains(FIXTURE_MARKER));

    let commit = edit.commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    let commit_debug = format!("{commit:?}");
    for rendered in [&patch_debug, &commit_debug] {
        assert!(!rendered.contains(FIXTURE_MARKER));
        assert!(!rendered.contains("Index/CalculationEngine.iwa"));
        assert!(!rendered.contains("preview.jpg"));
    }
    Ok(())
}

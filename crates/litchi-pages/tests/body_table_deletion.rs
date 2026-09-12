//! Focused package-level resource, locality, and patch-boundary coverage for
//! Pages body-table deletion.

use std::error::Error as StdError;

use litchi_iwa_archive::Limits as PhysicalLimits;
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_core::Limits as ArchiveLimits;
use litchi_pages::{BodyTableDeletionError, BodyTableSelector, Limits, Package};

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-catalog-native.pages"
));
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SHARED_IWA_MEMBER: &str = "Index/ViewState.iwa";

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn entry_data(bytes: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes)?;
    catalog
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| format!("missing package member {name}").into())
}

fn remove_metadata_member(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes)?;
    if !catalog.iter().any(|entry| entry.name() == METADATA_MEMBER) {
        return Err("fixture has no PackageMetadata member to remove".into());
    }
    Ok(catalog.reassemble_with_deletions_to_bytes(
        &[],
        &[METADATA_MEMBER],
        PhysicalLimits::default(),
    )?)
}

fn member_names(bytes: &[u8]) -> TestResult<Vec<String>> {
    Ok(Catalog::from_bytes(bytes)?
        .iter()
        .map(|entry| entry.name().to_owned())
        .collect())
}

fn iwa_object_count(bytes: &[u8], name: &str) -> TestResult<usize> {
    let data = entry_data(bytes, name)?;
    let stream = litchi_iwa_core::SnappyStream::decompress(&data)?;
    let archive = litchi_iwa_core::Archive::parse(stream.as_bytes())?;
    Ok(archive.objects.len())
}

#[test]
fn lowered_transaction_budget_fails_atomically() -> TestResult {
    // Ingress success is monotonic in the header-field ceiling. Find the
    // smallest profile that opens the package, then spend that same strict
    // profile on deletion. This keeps the resource test bounded to logarithmic
    // full-package parses instead of probing every ceiling.
    let mut lower = 1usize;
    let mut upper = 4_096usize;
    let mut smallest_success = None;
    while lower <= upper {
        let fields = lower + (upper - lower) / 2;
        let archive_limits = ArchiveLimits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        match Package::from_bytes_with_limits(SOURCE, limits) {
            Ok(package) => {
                smallest_success = Some(package);
                upper = fields.saturating_sub(1);
            },
            Err(_) => {
                lower = fields.saturating_add(1);
            },
        }
    }
    let package = smallest_success.ok_or("no successful ingress limit boundary found")?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_body_table(BodyTableSelector::index(0)),
        Err(BodyTableDeletionError::LimitExceeded { .. })
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn deletion_preserves_unrelated_zip_members_byte_for_byte() -> TestResult {
    let source = Package::from_bytes(SOURCE)?;
    let commit = source.remove_body_table(BodyTableSelector::index(0))?;
    let before = Catalog::from_bytes(SOURCE)?;
    let after_bytes = exact_bytes(commit.package())?;
    let after = Catalog::from_bytes(&after_bytes)?;

    for entry in before
        .iter()
        .filter(|entry| !entry.name().ends_with(".iwa") && !PREVIEWS.contains(&entry.name()))
    {
        let retained = after
            .iter()
            .find(|candidate| candidate.name() == entry.name())
            .ok_or_else(|| format!("unrelated member {} was removed", entry.name()))?;
        assert_eq!(
            retained.data(),
            entry.data(),
            "member {} changed",
            entry.name()
        );
    }
    Ok(())
}

#[test]
fn deletion_removes_empty_iwa_members_without_package_metadata_registration() -> TestResult {
    let metadata_free = remove_metadata_member(SOURCE)?;

    for (label, source, metadata_expected) in [
        ("registered", SOURCE, true),
        ("metadata-free", metadata_free.as_slice(), false),
    ] {
        let before_names = member_names(source)?;
        assert_eq!(
            before_names
                .iter()
                .filter(|name| *name == METADATA_MEMBER)
                .count(),
            if metadata_expected { 1 } else { 0 },
            "{label}: metadata registration shape",
        );

        let package = Package::from_bytes(source)?;
        let commit = package.remove_body_table(BodyTableSelector::index(0))?;
        let target = exact_bytes(commit.package())?;
        let after_names = member_names(&target)?;
        assert_eq!(
            after_names
                .iter()
                .filter(|name| *name == METADATA_MEMBER)
                .count(),
            if metadata_expected { 1 } else { 0 },
            "{label}: deletion changed the metadata registration shape",
        );
        let removed_iwa = before_names
            .iter()
            .filter(|name| name.ends_with(".iwa") && !after_names.contains(name))
            .collect::<Vec<_>>();

        assert!(
            !removed_iwa.is_empty(),
            "{label}: body-table deletion must physically remove an emptied IWA member"
        );
        assert!(
            removed_iwa
                .iter()
                .all(|name| name.starts_with("Index/Tables/")),
            "{label}: only table-owned IWA members may disappear: {removed_iwa:?}"
        );
        for name in removed_iwa {
            assert!(
                iwa_object_count(source, name)? > 0,
                "{label}: removed member {name} had no source objects"
            );
        }

        // ViewState is a separate shared component in both fixture shapes.
        // Exact member bytes prove that its unrelated objects survive the
        // physical deletion unchanged.
        assert_eq!(
            entry_data(&target, SHARED_IWA_MEMBER)?,
            entry_data(source, SHARED_IWA_MEMBER)?,
            "{label}: unrelated shared IWA member changed"
        );
        for entry in Catalog::from_bytes(source)?
            .iter()
            .filter(|entry| !entry.name().ends_with(".iwa") && !PREVIEWS.contains(&entry.name()))
        {
            assert_eq!(
                entry_data(&target, entry.name())?,
                entry.data(),
                "{label}: unrelated ZIP member {} changed",
                entry.name()
            );
        }
    }
    Ok(())
}

#[test]
fn deleted_previews_are_restored_by_the_inverse_patch() -> TestResult {
    let source = Package::from_bytes(SOURCE)?;
    let commit = source.remove_body_table(BodyTableSelector::index(0))?;
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    let changed = exact_bytes(commit.package())?;
    let changed_catalog = Catalog::from_bytes(&changed)?;
    assert!(
        PREVIEWS
            .iter()
            .all(|name| changed_catalog.iter().all(|entry| entry.name() != *name))
    );

    let restored = commit
        .package()
        .apply_body_table_deletion(&commit.patch().inverse())?;
    let restored_bytes = exact_bytes(restored.package())?;
    for name in PREVIEWS {
        assert_eq!(
            entry_data(&restored_bytes, name)?,
            entry_data(SOURCE, name)?
        );
    }
    assert_eq!(restored_bytes, SOURCE);
    Ok(())
}

#[test]
fn patch_rejects_an_independently_reallocated_source_with_different_bytes() -> TestResult {
    let source = Package::from_bytes(SOURCE)?;
    let commit = source.remove_body_table(BodyTableSelector::index(0))?;

    let source_catalog = Catalog::from_bytes(SOURCE)?;
    let altered_bytes = source_catalog.reassemble_to_bytes(
        &[EntryEdit::new("preview.jpg", b"independent source bytes")],
        PhysicalLimits::default(),
    )?;
    let altered = Package::from_bytes(&altered_bytes)?;
    assert!(matches!(
        altered.apply_body_table_deletion(commit.patch()),
        Err(BodyTableDeletionError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&altered)?, altered_bytes);
    Ok(())
}

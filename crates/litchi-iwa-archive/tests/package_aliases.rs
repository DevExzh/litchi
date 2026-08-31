use litchi_iwa_archive::package::{
    PackageEntry, PackageEntryChangeKind, PackageEntryStoreError, PackagePatch,
};
use litchi_iwa_package::EntryStore;

#[test]
fn archive_aliases_reach_neutral_package_primitives_directly() {
    // These annotations are intentional: they prove that the archive names
    // are direct re-exports, rather than facade wrappers that copy or convert
    // neutral package values.
    let source_entry: litchi_iwa_package::Entry =
        PackageEntry::new("Index/Document.iwa".to_owned(), vec![1, 2, 3]);
    let source = EntryStore::try_from_entries(vec![source_entry])
        .unwrap_or_else(|error| panic!("valid source entry should be accepted: {error}"));
    let target = EntryStore::try_from_entries(vec![PackageEntry::new(
        "Index/Document.iwa".to_owned(),
        vec![4, 5, 6],
    )])
    .unwrap_or_else(|error| panic!("valid target entry should be accepted: {error}"));

    let patch: litchi_iwa_package::Patch = PackagePatch::between(&source, &target);
    assert_eq!(patch.version(), PackagePatch::VERSION);
    assert_eq!(patch.changes().len(), 1);
    assert_eq!(patch.changes()[0].kind(), PackageEntryChangeKind::Replaced);

    assert_eq!(
        [
            PackageEntryChangeKind::Added,
            PackageEntryChangeKind::Removed,
            PackageEntryChangeKind::Replaced,
            PackageEntryChangeKind::Reordered,
        ],
        [
            litchi_iwa_package::EntryChangeKind::Added,
            litchi_iwa_package::EntryChangeKind::Removed,
            litchi_iwa_package::EntryChangeKind::Replaced,
            litchi_iwa_package::EntryChangeKind::Reordered,
        ]
    );

    // Keep every neutral store error constructor available through the
    // archive-owned name, including the error emitted by the direct patch.
    let duplicate = PackageEntryStoreError::DuplicateEntry("duplicate".to_owned());
    let invalid_position = PackageEntryStoreError::InvalidPosition {
        position: 2,
        len: 1,
    };
    let allocation = PackageEntryStoreError::Allocation { requested: 3 };
    let mismatch = PackageEntryStoreError::PatchSourceMismatch;
    assert_eq!(
        duplicate,
        litchi_iwa_package::Error::DuplicateEntry("duplicate".to_owned())
    );
    assert_eq!(
        invalid_position,
        litchi_iwa_package::Error::InvalidPosition {
            position: 2,
            len: 1,
        }
    );
    assert_eq!(
        allocation,
        litchi_iwa_package::Error::Allocation { requested: 3 }
    );
    assert_eq!(mismatch, litchi_iwa_package::Error::PatchSourceMismatch);

    let unrelated = EntryStore::try_from_entries(vec![PackageEntry::new(
        "Index/Other.iwa".to_owned(),
        vec![9],
    )])
    .unwrap_or_else(|error| panic!("valid unrelated entry should be accepted: {error}"));
    assert_eq!(patch.apply(&unrelated), Err(mismatch));
}

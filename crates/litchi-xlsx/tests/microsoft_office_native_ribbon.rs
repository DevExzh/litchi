//! XLSX uses the public shared OPC Ribbon owner; callback XML remains inert.
use litchi_ooxml_common::{
    Error,
    ribbon::{self, Family, Version},
};
use litchi_opc::OpcPackage;
use litchi_xlsx::Package;
use std::path::Path;
#[path = "../../../tools/test_support/native_ribbon.rs"]
mod shared;
struct Host;
impl shared::Host for Host {
    type Package = Package;
    type Patch = ribbon::Patch;
    const SUFFIX: &'static str = ".xlsx";
    fn open(bytes: &[u8]) -> Package {
        Package::from_bytes(bytes).unwrap()
    }
    fn opc(p: &Package) -> OpcPackage {
        p.clone().into_plain_opc()
    }
    fn from_opc(p: OpcPackage) -> Package {
        Package::from_opc(p).unwrap()
    }
    fn bytes(p: &Package) -> Vec<u8> {
        p.to_bytes().unwrap()
    }
    fn save(p: &Package, path: &Path) {
        p.save(path).unwrap();
    }
    fn reopen(path: &Path) -> Package {
        Package::open(path).unwrap()
    }
    fn read(p: &Package) -> shared::Snapshot {
        shared::snapshot(ribbon::load(&Self::opc(p), &ribbon::Limits::standard()).unwrap())
    }
    fn edit(p: &Package, actions: &[shared::Action]) -> (Package, Self::Patch) {
        let opc = Self::opc(p);
        let mut e = ribbon::Edit::new(&opc).unwrap();
        for action in actions {
            match action {
                shared::Action::Put(version, bytes) => {
                    e.put(*version, bytes.clone()).unwrap();
                },
                shared::Action::Remove(family) => {
                    e.remove(*family).unwrap();
                },
            }
        }
        let c = e.commit().unwrap();
        let patch = c.patch().clone();
        (Self::from_opc(c.into_package()), patch)
    }
    fn apply(p: &Package, patch: &Self::Patch) -> Package {
        Self::from_opc(patch.apply(&Self::opc(p)).unwrap().into_package())
    }
    fn inverse(patch: &Self::Patch) -> Self::Patch {
        patch.inverse()
    }
    fn noop(patch: &Self::Patch) -> bool {
        patch.is_noop()
    }
    fn assert_stale(p: &Package, patch: &Self::Patch) {
        assert!(matches!(
            patch.apply(&Self::opc(p)),
            Err(Error::PatchConflict)
        ));
    }
}

#[test]
fn native_xlsx_ribbon_families_round_trip_with_callbacks_and_native_parts_inert() {
    shared::lifecycle::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        &[
            "excel-metadata.xlsx",
            "excel-recorded-macro.xlsm",
            "excel-visual-analytics-strict.xlsx",
        ],
    );
}

#[test]
fn native_xlsx_ribbon_image_collection_and_shared_owner_fork_preserve_native_assets() {
    shared::image_ownership::<Host>(env!("CARGO_MANIFEST_DIR"), "excel-metadata.xlsx");
}

#[test]
fn native_xlsx_ribbon_invalid_xml_and_version_mismatch_preserve_valid_staging() {
    shared::rejected_edit::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        "excel-recorded-macro.xlsm",
        |p, valid| {
            let opc = p.clone().into_plain_opc();
            let mut e = ribbon::Edit::new(&opc).unwrap();
            assert!(e.put(Version::V2010, valid.clone()).unwrap());
            assert!(
                matches!(e.put(Version::V2007, valid.clone()), Err(Error::Invalid(m)) if m == "Ribbon XML root namespace does not match its package relationship")
            );
            assert!(
                matches!(e.put(Version::Ui2, valid.clone()), Err(Error::Invalid(m)) if m == "Ribbon root namespace does not match requested Ui2 version")
            );
            assert!(
                matches!(e.put(Version::V2010, b"<!DOCTYPE customUI><customUI/>".to_vec()), Err(Error::Invalid(m)) if m == "DTD is forbidden in Ribbon XML")
            );
            assert_eq!(
                e.get(Family::Modern).unwrap(),
                Some((Version::V2010, valid))
            );
            assert_eq!(e.get(Family::Legacy).unwrap(), None);
            let c = e.commit().unwrap();
            let patch = c.patch().clone();
            (Package::from_opc(c.into_package()).unwrap(), patch)
        },
    );
}

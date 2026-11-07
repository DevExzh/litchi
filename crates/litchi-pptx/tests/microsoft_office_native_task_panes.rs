//! Native parents with explicit inert task-pane additions; no add-in activation.
use litchi_ooxml_common::web::{Conformance, Panes};
use litchi_pptx::Package;
use std::path::Path;
#[path = "../../../tools/test_support/native_task_panes.rs"]
mod shared;
struct Host;
impl shared::Host for Host {
    type Package = Package;
    type Patch = litchi_pptx::web::Patch;
    fn open(b: &[u8]) -> Package {
        Package::from_bytes(b).unwrap()
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
    fn panes(p: &Package) -> Option<Panes> {
        p.task_panes().unwrap()
    }
    fn put(p: &Package, panes: Panes, conformance: Conformance) -> (Package, Self::Patch) {
        let mut e = p.edit_task_panes().unwrap();
        e.set(panes, conformance).unwrap();
        let c = e.commit().unwrap();
        {
            let patch = c.patch().clone();
            (c.into_package(), patch)
        }
    }
    fn remove(p: &Package) -> (Package, Self::Patch) {
        let mut e = p.edit_task_panes().unwrap();
        e.remove().unwrap();
        let c = e.commit().unwrap();
        {
            let patch = c.patch().clone();
            (c.into_package(), patch)
        }
    }
    fn apply(p: &Package, patch: &Self::Patch) -> Package {
        p.apply_task_panes(patch).unwrap().into_package()
    }
    fn inverse(patch: &Self::Patch) -> Self::Patch {
        patch.inverse()
    }
    fn noop(patch: &Self::Patch) -> bool {
        patch.is_noop()
    }
    fn assert_stale(p: &Package, patch: &Self::Patch) {
        assert!(matches!(
            p.apply_task_panes(patch),
            Err(litchi_pptx::Error::PatchConflict)
        ));
    }
}
#[test]
fn native_pptx_task_pane_lifecycle_preserves_native_parts_and_shared_snapshots() {
    shared::lifecycle::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        &[
            (
                "powerpoint/powerpoint-media-objects.pptx",
                Conformance::Transitional,
            ),
            (
                "powerpoint/powerpoint-passive-macro.pptm",
                Conformance::Transitional,
            ),
        ],
    );
}

#[test]
fn native_pptx_task_pane_resource_collision_preserves_valid_staging() {
    shared::collision::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        "powerpoint/powerpoint-media-objects.pptx",
        |p, valid, invalid, reason| {
            let mut e = p.edit_task_panes().unwrap();
            e.set(valid, Conformance::Transitional).unwrap();
            assert!(
                matches!(e.set(invalid,Conformance::Transitional),Err(litchi_pptx::Error::CustomProperties(litchi_ooxml_common::Error::Invalid(message))) if message==reason)
            );
            let c = e.commit().unwrap();
            let patch = c.patch().clone();
            (c.into_package(), patch)
        },
    );
}

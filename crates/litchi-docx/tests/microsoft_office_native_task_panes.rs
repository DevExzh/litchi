//! Native parents with explicit inert task-pane additions; no add-in activation.
use litchi_docx::Package;
use litchi_ooxml_common::web::{Conformance, Panes};
use std::path::Path;
#[path = "../../../tools/test_support/native_task_panes.rs"]
mod shared;
struct Host;
impl shared::Host for Host {
    type Package = Package;
    type Patch = litchi_docx::task_panes::Patch;
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
        let mut e = p.edit_task_panes();
        e.put(panes, conformance).unwrap();
        let c = e.commit().unwrap();
        {
            let patch = c.patch().clone();
            (c.into_package(), patch)
        }
    }
    fn remove(p: &Package) -> (Package, Self::Patch) {
        let mut e = p.edit_task_panes();
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
            Err(litchi_docx::Error::PatchConflict { .. })
        ));
    }
}
#[test]
fn native_docx_task_pane_lifecycle_preserves_native_parts_and_shared_snapshots() {
    shared::lifecycle::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        &[
            ("word/word-visual-objects.docx", Conformance::Transitional),
            ("word/word-recorded-macro.docm", Conformance::Transitional),
        ],
    );
}

#[test]
fn native_docx_task_pane_resource_collision_preserves_valid_staging() {
    shared::collision::<Host>(
        env!("CARGO_MANIFEST_DIR"),
        "word/word-visual-objects.docx",
        |p, valid, invalid, reason| {
            let mut e = p.edit_task_panes();
            e.put(valid, Conformance::Transitional).unwrap();
            assert!(
                matches!(e.put(invalid,Conformance::Transitional),Err(litchi_docx::Error::Common(litchi_ooxml_common::Error::Invalid(message))) if message==reason)
            );
            let c = e.commit().unwrap();
            let patch = c.patch().clone();
            (c.into_package(), patch)
        },
    );
}

#[test]
fn native_docx_task_pane_limit_refusal_keeps_prior_valid_staging() {
    use litchi_docx::{task_panes::EditOptions, web_extensions as web};
    let bytes = shared::native(env!("CARGO_MANIFEST_DIR"), "word/word-recorded-macro.docm");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut limits = web::Limits::standard();
    limits.items = 2;
    let mut e = source.edit_task_panes_with_options(EditOptions::default().limits(limits));
    let valid = shared::panes(None);
    e.put(valid.clone(), Conformance::Transitional).unwrap();
    let mut too_many = valid;
    too_many
        .push(web::Pane::new(
            web::AddIn::new(
                "third-066",
                web::Reference::new("third", "1", web::Store::Omex).unwrap(),
            )
            .unwrap(),
        ))
        .unwrap();
    let error = e.put(too_many, Conformance::Transitional).unwrap_err();
    assert!(
        matches!(
            &error,
            litchi_docx::Error::Common(litchi_ooxml_common::Error::Limit {
                kind: litchi_ooxml_common::LimitKind::Web(web::LimitKind::TaskPaneItems),
                maximum: 2,
                observed: 3
            })
        ),
        "{error:?}"
    );
    let c = e.commit().unwrap();
    let p = Package::from_bytes(&c.package().to_bytes().unwrap()).unwrap();
    assert_eq!(p.task_panes().unwrap().unwrap().len(), 2);
    assert!(p.task_panes().unwrap().unwrap().get("third-066").is_none());
    assert_eq!(p.task_panes().unwrap(), c.package().task_panes().unwrap());
    let before = shared::members(&bytes);
    let after = shared::members(&p.to_bytes().unwrap());
    for (name, data) in before {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert!(after.get(&name) == Some(&data), "changed {name}");
        }
    }
    assert!(source.to_bytes().unwrap() == bytes);
}

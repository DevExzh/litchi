use litchi_ole::ppt::Package;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

#[test]
fn bundled_poi_ole_presentations_expose_inert_metadata_and_storage() {
    for name in [
        "ppt_with_embeded.ppt",
        "testPPT_oleWorkbook.ppt",
        "ole2-embedding-2003.ppt",
    ] {
        let mut package = Package::open(fixture(name)).expect("open POI OLE fixture");
        let presentation = package.presentation().expect("parse POI OLE fixture");
        let objects = presentation
            .ole_objects()
            .expect("parse external-object list")
            .expect("fixture has external-object list");
        assert!(
            !objects.objects.is_empty(),
            "{name} has no typed OLE objects"
        );
        for object in &objects.objects {
            let storage = presentation
                .ole_storage(object.persist_id())
                .expect("resolve inert ExOleObjStg")
                .expect("OLE object has persisted storage");
            assert!(
                !storage.data.is_empty(),
                "{name} has empty persisted storage"
            );
        }
    }
}

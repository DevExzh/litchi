//! A source-backed shape removal drops the Parts only the removed shape
//! owned, including those the slide XML never names: a SmartArt's cached
//! drawing (named by its data part) and a legacy VML drawing that held only
//! the removed OLE object's shape. The patch replays on the source, its
//! inverse restores the exact package after reopening, and the local inverse
//! restores the exact bytes.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_pptx::source::{Package, Patch};

fn data(relative: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative),
    )
    .unwrap()
}

fn open(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = archive.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

/// Remove `shape` from slide 0 and return the removed member names after
/// checking replay and both inverses.
fn removed(input: &[u8], shape: &str) -> Vec<String> {
    let source = open(input);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, shape).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let output = changed.to_bytes().unwrap();
    let forward = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let replay = open(input)
        .apply(&forward)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert!(members(&replay) == members(&output));
    let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let restored = open(&output)
        .apply(&inverse)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert!(members(&restored) == members(input));
    assert!(
        changed
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == input
    );
    let after = members(&output);
    members(input)
        .into_keys()
        .filter(|name| !after.contains_key(name))
        .collect()
}

#[test]
fn a_removed_smartart_takes_its_cached_drawing_and_pictures() {
    let gone = removed(&data("ooxml/pptx/smartart-simple.pptx"), "Diagram 3");
    assert!(
        gone.iter()
            .any(|name| name.starts_with("ppt/diagrams/drawing")),
        "{gone:?}"
    );
    assert!(
        gone.iter().any(|name| name == "ppt/media/image1.jpeg"),
        "{gone:?}"
    );
}

#[test]
fn a_removed_ole_object_takes_its_legacy_vml_drawing_and_preview() {
    let input = data("libreoffice-core/sd/qa/unit/data/pptx/ole.pptx");
    let gone = removed(&input, "Object 3");
    assert_eq!(
        gone,
        [
            "ppt/drawings/_rels/vmlDrawing1.vml.rels",
            "ppt/drawings/vmlDrawing1.vml",
            "ppt/embeddings/Microsoft_Word_Document.docx",
            "ppt/media/image1.emf",
        ]
    );
}

/// LibreOffice relates each slide's SmartArt drawing without writing the
/// Part, and keeps its content-types override. Removing two frames in one
/// edit, in either order, succeeds, replays and inverts exactly.
#[test]
fn smartart_removals_beside_a_missing_drawing_part_compose() {
    let input = data("ooxml/pptx/smartart-autoTxRot.pptx");
    for order in [[0_usize, 1], [1, 0]] {
        let source = open(&input);
        let mut edit = source.edit().unwrap();
        for slide in order {
            edit.remove_shape(slide, "Diagram 3")
                .unwrap_or_else(|error| panic!("{order:?} slide {slide}: {error}"));
        }
        let (changed, patch) = edit.commit().unwrap().into_parts();
        let output = changed.to_bytes().unwrap();
        let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
        let restored = open(&output)
            .apply(&inverse)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap();
        assert!(members(&restored) == members(&input), "{order:?}");
        let forward = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
        let replay = open(&input)
            .apply(&forward)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap();
        assert!(members(&replay) == members(&output), "{order:?}");
        assert!(
            changed
                .apply(&patch.inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
                == input
        );
    }
}

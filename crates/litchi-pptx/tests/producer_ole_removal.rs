#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the corpus fixture and assertions intentionally panic"
)]

//! Removing a whole `p:graphicFrame` whose `a:graphicData` encloses an
//! `mc:AlternateContent` (an embedded OLE object), on the source bytes
//! (record 0303, record 0290 F3).
//!
//! The slide guard admits markup compatibility strictly inside top-level
//! shape-tree children for whole-shape removal only; every other edit on
//! such a slide stays refused.

use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::Error;
use litchi_pptx::source::{Package, Patch};
use soapberry_zip::office::ArchiveReader;

const OLE: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/ole.pptx");

fn open(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn slide(bytes: &[u8]) -> String {
    String::from_utf8(
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .read("ppt/slides/slide1.xml")
            .unwrap(),
    )
    .unwrap()
}

#[test]
fn an_ole_frame_under_markup_compatibility_is_removed_on_the_source_bytes() {
    let source = open(OLE);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 0_usize).unwrap());
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();

    // Exactly the frame's source bytes leave; the rectangle and every other
    // byte stay.
    let before = slide(OLE);
    let start = before.find("<p:graphicFrame>").unwrap();
    let end = before.find("</p:graphicFrame>").unwrap() + "</p:graphicFrame>".len();
    let expected = format!("{}{}", &before[..start], &before[end..]);
    assert_eq!(slide(&output), expected);
    assert!(slide(&output).contains("Rectangle 4"));

    // The local inverse, the transported forward patch and the durable
    // inverse after reopening all restore exactly.
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    let forward = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    assert_eq!(
        slide(
            &open(OLE)
                .apply(&forward)
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
        ),
        expected
    );
    let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let restored = open(&output)
        .apply(&inverse)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert_eq!(slide(&restored), before);
}

#[test]
fn other_edits_on_a_slide_with_enclosed_markup_compatibility_stay_refused() {
    let source = open(OLE);
    let mut edit = source.edit().unwrap();
    let error = edit.set_shape_text(0, 1_usize, "changed").unwrap_err();
    assert!(
        matches!(error, Error::Refused(refusal) if refusal.reason().contains("markup-compatibility")),
        "{error:?}"
    );
    let error = edit
        .set_slide_properties(Position::new(0), Some("renamed"), None)
        .unwrap_err();
    assert!(matches!(error, Error::Refused(_)), "{error:?}");
    assert!(edit.commit().unwrap().into_parts().1.is_noop());
}

#[test]
fn a_plain_shape_beside_the_frame_is_removed_too() {
    // Removing the rectangle keeps the frame and its markup compatibility.
    let source = open(OLE);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(0, 1_usize).unwrap());
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = slide(&updated.to_bytes().unwrap());
    assert!(output.contains("mc:AlternateContent"));
    assert!(!output.contains("Rectangle 4"));
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

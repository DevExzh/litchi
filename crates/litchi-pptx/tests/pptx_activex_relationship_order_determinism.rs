#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
//! An `ActiveX` control snapshot's public revision, and whether its patch is
//! accepted, must be functions of the package and not of a hash seed.
//!
//! `slide::package::relationship_states` turns a `Relationships` iteration —
//! a `HashMap` walk — into a `Vec<RelationshipState>` and used to end in a bare
//! `.collect()`. `Snapshot::from_parts` passes the slide and descriptor arrays
//! through `sorted_relationships`, so those two were already canonical, but two
//! uses were not:
//!
//! * the array captured for the `ActiveX` **binary** part
//!   (`load_binary`) is stored in `BinarySource::relationships` and fed
//!   straight into `fingerprint`, so the public [`Snapshot::revision`] of a
//!   binary part owning two relationships differed between loads of the same
//!   bytes, and `Snapshot::same_source` — which compares `binary` by value —
//!   refused the matching patch at random with `invalid_revision()`;
//! * `install_patch` and `ensure_binary_part` rebuild the array from the staged
//!   package and compare it with `!=` against the patch target, one side sorted
//!   and the other not, so a patch that matched its own package was refused
//!   with "descriptor relationship lifecycle does not match the patch target"
//!   or "ActiveX binary part relationships are stale".
//!
//! ADR 0006 requires serialization to be deterministic unless a `Clock`, actor
//! identity, or cryptographic RNG is explicitly supplied; a snapshot revision
//! that a caller may persist and compare is exactly such an output. ADR 0003's
//! source-checked patches must conflict on a real overlap, not on a seed.
//!
//! Each `Relationships::new()` builds a fresh `HashMap`, whose `RandomState` is
//! seeded from a per-thread counter that advances on every construction, so
//! rebuilding the same package in one process varies the visit order exactly as
//! separate processes do.

use std::collections::BTreeSet;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_pptx::presentation::embedded::controls::{Limits, slide};

/// Repetitions per determinism assertion; 2^-128 accidental agreement.
const REPEATS: usize = 128;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const AX: &str = "http://schemas.microsoft.com/office/2006/activeX";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const DESCRIPTOR: &str = "/ppt/activeX/activeX1.xml";
const BINARY: &str = "/ppt/activeX/activeX1.bin";
const IMAGE_ONE: &str = "/ppt/media/image1.png";
const IMAGE_TWO: &str = "/ppt/media/image2.png";

const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const DESCRIPTOR_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const BINARY_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX";
const CONTROL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
const BINARY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const IMAGE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// Build the control graph. `companions` adds one extra inert image
/// relationship to the descriptor and to the binary part, which is the shape
/// that gives the captured arrays an order to get wrong; without it every
/// captured array holds at most one relationship and no order exists.
fn control_package(companions: bool) -> OpcPackage {
    let slide_xml = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:controls><p:control name="OldName" r:id="rIdControl" showAsIcon="0" imgW="10" imgH="20"><x:opaque xmlns:x="urn:opaque">retain</x:opaque></p:control></p:controls></p:cSld></p:sld>"#
    );
    let descriptor_xml = format!(
        r#"<ax:ocx xmlns:ax="{AX}" xmlns:r="{REL}" ax:classid="old-class" ax:persistence="persistStream" r:id="rIdBinary"><x:opaque xmlns:x="urn:opaque"><x:value>retain</x:value></x:opaque></ax:ocx>"#
    );

    let mut slide = BlobPart::new(
        PackURI::new(SLIDE).unwrap(),
        SLIDE_CONTENT_TYPE.to_owned(),
        slide_xml.into_bytes(),
    );
    slide
        .rels_mut()
        .try_add_relationship(
            CONTROL_RELATIONSHIP.to_owned(),
            "../activeX/activeX1.xml".to_owned(),
            "rIdControl".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();

    let mut descriptor = BlobPart::new(
        PackURI::new(DESCRIPTOR).unwrap(),
        DESCRIPTOR_CONTENT_TYPE.to_owned(),
        descriptor_xml.into_bytes(),
    );
    descriptor
        .rels_mut()
        .try_add_relationship(
            BINARY_RELATIONSHIP.to_owned(),
            "activeX1.bin".to_owned(),
            "rIdBinary".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();

    let mut binary = BlobPart::new(
        PackURI::new(BINARY).unwrap(),
        BINARY_CONTENT_TYPE.to_owned(),
        vec![0, 1, 2, 255],
    );

    if companions {
        // `rIdCompanion` sorts after `rIdBinary`, so a `HashMap` walk that
        // happens to visit it first produces the reversed array.
        descriptor
            .rels_mut()
            .try_add_relationship(
                IMAGE_RELATIONSHIP.to_owned(),
                "../media/image1.png".to_owned(),
                "rIdCompanion".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        for (r_id, target) in [
            ("rIdAlpha", "../media/image1.png"),
            ("rIdBeta", "../media/image2.png"),
        ] {
            binary
                .rels_mut()
                .try_add_relationship(
                    IMAGE_RELATIONSHIP.to_owned(),
                    target.to_owned(),
                    r_id.to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
    }

    let mut package = OpcPackage::new();
    package.try_add_part(Box::new(slide)).unwrap();
    package.try_add_part(Box::new(descriptor)).unwrap();
    package.try_add_part(Box::new(binary)).unwrap();
    for image in [IMAGE_ONE, IMAGE_TWO] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(image).unwrap(),
                "image/png".to_owned(),
                b"\x89PNG\r\n\x1a\n".to_vec(),
            )))
            .unwrap();
    }
    package
}

fn snapshot(package: &OpcPackage) -> slide::Snapshot {
    let slide_uri = PackURI::new(SLIDE).unwrap();
    let part = package.get_part(&slide_uri).unwrap();
    slide::load(package, 0, part, 0, &mut Limits::default()).unwrap()
}

/// Capture against one build of the package and publish against an independent
/// build of the same graph, the way a caller that keeps a patch across a reopen
/// does.
fn rename_across_two_builds(companions: bool) -> litchi_pptx::Result<slide::Snapshot> {
    let source = control_package(companions);
    let mut target = control_package(companions);
    let mut edit = snapshot(&source).edit();
    edit.set_name(Some("Renamed".to_owned())).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed(), "the rename changes the slide XML");
    slide::apply_patch(&mut target, commit.patch())
}

#[test]
fn snapshot_revision_is_a_function_of_the_package_not_of_the_hash_seed() {
    let mut revisions = BTreeSet::new();
    for _ in 0..REPEATS {
        revisions.insert(snapshot(&control_package(true)).revision());
    }
    assert_eq!(
        revisions.len(),
        1,
        "the same control graph produced {} distinct public revisions: {revisions:?}",
        revisions.len()
    );
}

#[test]
fn a_matching_patch_is_accepted_against_every_build_of_the_same_package() {
    for attempt in 0..REPEATS {
        rename_across_two_builds(true).unwrap_or_else(|error| {
            panic!(
                "attempt {attempt}: a patch captured from one build of the package was refused \
                 against an identical build: {error}"
            )
        });
    }
}

#[test]
fn single_relationship_graphs_keep_their_revision_and_accept_their_patch() {
    let mut revisions = BTreeSet::new();
    for attempt in 0..REPEATS {
        revisions.insert(snapshot(&control_package(false)).revision());
        rename_across_two_builds(false).unwrap_or_else(|error| {
            panic!("attempt {attempt}: single-relationship patch was refused: {error}")
        });
    }
    assert_eq!(revisions.len(), 1);
}

#[test]
fn a_genuinely_different_package_is_still_refused() {
    let source = control_package(true);
    let mut target = control_package(true);
    target
        .get_part_mut(&PackURI::new(BINARY).unwrap())
        .unwrap()
        .set_blob(vec![9, 9, 9, 9]);
    let mut edit = snapshot(&source).edit();
    edit.set_name(Some("Renamed".to_owned())).unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        slide::apply_patch(&mut target, commit.patch()).is_err(),
        "a changed binary payload must still be refused"
    );
}

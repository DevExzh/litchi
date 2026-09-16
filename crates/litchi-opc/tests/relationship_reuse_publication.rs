//! Publishing after a relationship call that establishes nothing.
//!
//! Change 0593 gives every relationship collection an open-time capture of its
//! canonical serialization, and publication copies the source `.rels` member
//! verbatim while that capture still stands. Change 0647 stops
//! [`litchi_opc::Relationships::add_relationship`] from dropping the capture on
//! an occupied identifier, which is the path
//! [`litchi_opc::Relationships::get_or_add`] takes when it reuses a
//! relationship the collection already carries.
//!
//! These tests pin the contract that change makes: a reusing call must publish
//! exactly what taking the mutable seam alone publishes, and exactly what the
//! pre-change serialize-and-byte-compare route published, **including on a
//! package whose `.rels` members are not spelled the way this crate would
//! serialize them**. That last condition is what makes the test meaningful:
//! 2,189 of the 2,216 relationships members in this repository's OOXML corpus
//! are spelled differently from their canonical form, so if keeping the capture
//! could move a published byte, it would move one here.
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::{OpcPackage, PackURI, PackageWriter};

/// A 132-member producer-written workbook. Changes 0593 and 0628 both measure
/// it, and neither its package member nor its drawing member is spelled the
/// way [`litchi_opc::Relationships::to_xml`] would spell it.
const FIXTURE: &str = "ooxml/xlsx/ConditionalFormattingSamples.xlsx";

/// A part whose relationships member carries fifteen internal relationships
/// and is 68 bytes shorter in the source than in canonical form.
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const DRAWING_MEMBER: &str = "xl/drawings/_rels/drawing1.xml.rels";
const PACKAGE_MEMBER: &str = "_rels/.rels";

fn fixture() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(FIXTURE);
    std::fs::read(&path).unwrap_or_else(|error| {
        panic!("read corpus package {}: {error}", path.display());
    })
}

fn member(archive: &[u8], name: &str) -> Vec<u8> {
    PhysPkgReader::new(archive)
        .expect("read archive")
        .read_member(name)
        .unwrap_or_else(|error| panic!("read member {name}: {error}"))
}

/// The smallest internal (type, target) pair of one collection, so the
/// selection is a function of the package rather than of a hash order.
fn smallest_internal_pair(package: &OpcPackage, owner: Option<&PackURI>) -> (String, String) {
    let rels = match owner {
        None => package.rels(),
        Some(owner) => package.get_part(owner).expect("owner part").rels(),
    };
    let mut pairs: Vec<(String, String)> = rels
        .iter()
        .filter(|relationship| !relationship.is_external())
        .map(|relationship| {
            (
                relationship.reltype().to_string(),
                relationship.target_ref().to_string(),
            )
        })
        .collect();
    pairs.sort();
    pairs.into_iter().next().expect("an internal relationship")
}

#[test]
fn the_fixture_is_not_spelled_the_way_this_crate_serializes_relationships() {
    let bytes = fixture();
    let package = OpcPackage::from_vec(bytes.clone()).expect("open fixture");
    let drawing = PackURI::new(DRAWING).expect("drawing URI");

    assert_ne!(
        member(&bytes, PACKAGE_MEMBER),
        package.rels().to_xml().into_bytes(),
        "the package relationships member is already canonical, so this file \
         cannot witness the byte question change 0647 asks"
    );
    assert_ne!(
        member(&bytes, DRAWING_MEMBER),
        package
            .get_part(&drawing)
            .expect("drawing part")
            .rels()
            .to_xml()
            .into_bytes(),
        "the drawing relationships member is already canonical"
    );
}

#[test]
fn reusing_a_package_relationship_publishes_what_taking_the_seam_publishes() {
    let bytes = fixture();
    let probe = OpcPackage::from_vec(bytes.clone()).expect("open fixture");
    let (reltype, target) = smallest_internal_pair(&probe, None);
    drop(probe);

    // Taking the seam revokes exact-source authorization without changing a
    // value, so all three legs publish through the preservation route.
    let mut seam = OpcPackage::from_vec(bytes.clone()).expect("open seam leg");
    let _seam = seam.rels_mut();
    let seam_output = PackageWriter::to_bytes(&seam).expect("publish seam leg");

    let mut reuse = OpcPackage::from_vec(bytes.clone()).expect("open reuse leg");
    let reused = reuse
        .rels_mut()
        .get_or_add(&reltype, &target)
        .r_id()
        .to_string();
    let reuse_output = PackageWriter::to_bytes(&reuse).expect("publish reuse leg");

    // Removing an absent identifier keeps every value identical but drops
    // every open-time proof, which is the route the reuse leg took before
    // change 0647. It is the oracle for "the capture is only ever a shortcut".
    let mut compared = OpcPackage::from_vec(bytes.clone()).expect("open compared leg");
    compared.rels_mut().remove("rIdAbsent0647");
    let compared_output = PackageWriter::to_bytes(&compared).expect("publish compared leg");

    assert!(
        reuse.rels().iter().any(|rel| rel.r_id() == reused),
        "reuse returned an identifier the collection does not carry"
    );
    assert_eq!(reuse.rels().len(), seam.rels().len());
    assert_eq!(reuse_output, seam_output);
    assert_eq!(reuse_output, compared_output);
    assert_eq!(
        member(&reuse_output, PACKAGE_MEMBER),
        member(&bytes, PACKAGE_MEMBER),
        "the published package relationships member must still be the source's \
         own spelling"
    );
}

#[test]
fn reusing_a_part_relationship_publishes_what_taking_the_seam_publishes() {
    let bytes = fixture();
    let drawing = PackURI::new(DRAWING).expect("drawing URI");
    let probe = OpcPackage::from_vec(bytes.clone()).expect("open fixture");
    let (reltype, target) = smallest_internal_pair(&probe, Some(&drawing));
    let established = probe.get_part(&drawing).expect("drawing part").rels().len();
    drop(probe);

    let mut seam = OpcPackage::from_vec(bytes.clone()).expect("open seam leg");
    let _seam = seam.get_part_mut(&drawing).expect("drawing part");
    let seam_output = PackageWriter::to_bytes(&seam).expect("publish seam leg");

    let mut reuse = OpcPackage::from_vec(bytes.clone()).expect("open reuse leg");
    let reused = reuse
        .get_part_mut(&drawing)
        .expect("drawing part")
        .rels_mut()
        .get_or_add(&reltype, &target)
        .r_id()
        .to_string();
    let reuse_output = PackageWriter::to_bytes(&reuse).expect("publish reuse leg");

    let mut compared = OpcPackage::from_vec(bytes.clone()).expect("open compared leg");
    compared
        .get_part_mut(&drawing)
        .expect("drawing part")
        .rels_mut()
        .remove("rIdAbsent0647");
    let compared_output = PackageWriter::to_bytes(&compared).expect("publish compared leg");

    let republished = reuse.get_part(&drawing).expect("drawing part").rels();
    assert!(republished.iter().any(|rel| rel.r_id() == reused));
    assert_eq!(republished.len(), established);
    assert_eq!(reuse_output, seam_output);
    assert_eq!(reuse_output, compared_output);
    assert_eq!(
        member(&reuse_output, DRAWING_MEMBER),
        member(&bytes, DRAWING_MEMBER),
        "the published drawing relationships member must still be the source's \
         own spelling"
    );
}

#[test]
fn establishing_a_new_relationship_still_republishes_the_member() {
    let bytes = fixture();
    let drawing = PackURI::new(DRAWING).expect("drawing URI");
    let (reltype, _target) = {
        let probe = OpcPackage::from_vec(bytes.clone()).expect("open fixture");
        smallest_internal_pair(&probe, Some(&drawing))
    };

    let mut package = OpcPackage::from_vec(bytes.clone()).expect("open fixture");
    let established = package
        .get_part_mut(&drawing)
        .expect("drawing part")
        .rels_mut()
        .get_or_add(&reltype, "../media/image0647.png")
        .r_id()
        .to_string();
    let output = PackageWriter::to_bytes(&package).expect("publish topology add");

    assert!(!established.is_empty());
    assert_ne!(
        member(&output, DRAWING_MEMBER),
        member(&bytes, DRAWING_MEMBER),
        "a relationship that was actually established must move the member"
    );
    let reopened = OpcPackage::from_bytes(&output).expect("reopen published package");
    assert!(
        reopened
            .get_part(&drawing)
            .expect("republished drawing part")
            .rels()
            .iter()
            .any(|rel| rel.target_ref() == "../media/image0647.png"),
        "the established relationship must survive the round trip"
    );
}

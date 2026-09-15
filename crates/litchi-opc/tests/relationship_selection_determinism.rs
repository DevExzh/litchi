#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
//! Reusing an existing relationship must pick the same one on every run.
//!
//! `Relationships` stores its relationships in a `HashMap` keyed by rId, so
//! `Relationships::iter` visits them in an order seeded per map instance. Every
//! serializer in the crate sorts before emitting — `Relationships::to_xml` and
//! `try_to_xml_bytes` sort by rId, `PublicationPlan` sorts its parts by
//! partname, and the signature resolver sorts its references and relationship
//! ids — so the `.rels` bytes themselves never depended on that order.
//!
//! `get_or_add` and `get_or_add_ext_rel` did. Both scanned `self.rels.values()`
//! and reused the *first* relationship whose type and target matched. When a
//! part owns two relationships with the same type and target under different
//! rIds, which of them was reused depended on nothing but the hash seed, and the
//! caller writes the returned rId straight into the part markup
//! (`Part::relate_to`, `Part::relate_to_ext`, `OpcPackage::relate_to`,
//! `OpcPackage::relate_to_external`), so the published bytes varied per process.
//!
//! That input shape is not hypothetical: four packages in this repository's own
//! corpus carry it, all of them external hyperlink or mail-merge relationships
//! (`docs/performance/results/change-0628/corpus-duplicate-relationships.txt`).
//!
//! ADR 0006 requires serialization to be deterministic unless a `Clock`, actor
//! identity, or cryptographic RNG is explicitly supplied, and none of these
//! paths is supplied any of the three. Both methods now reuse the
//! lexicographically smallest matching rId — the relationship that the sorted
//! `.rels` member emits first — so the choice is a function of the collection.
//!
//! Each `Relationships::new()` builds a fresh `HashMap`, whose `RandomState` is
//! seeded from a per-thread counter that advances on every construction, so
//! repeating a build in one process varies the visit order exactly as separate
//! processes do.

use litchi_opc::constants::relationship_type;
use litchi_opc::{OpcPackage, PackURI, Relationships, TargetMode};
use std::collections::BTreeSet;

/// Repetitions per determinism assertion. Change 0617 saw twelve distinct
/// outcomes in twelve processes at eight hash keys; with two candidates a
/// single build agrees by chance about half the time, so the count has to be
/// large enough that accidental agreement is not worth considering. At 128
/// builds that probability is below 2^-127.
const REPEATS: usize = 128;

const HYPERLINK: &str = relationship_type::HYPERLINK;
const IMAGE: &str = relationship_type::IMAGE;
const MAIL_MERGE_SOURCE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource";

/// Build a collection holding `ids` relationships that all share one type and
/// one target, inserted in the given order.
fn duplicates(ids: &[&str], reltype: &str, target: &str, mode: TargetMode) -> Relationships {
    let mut relationships = Relationships::new("/xl/worksheets".to_string());
    for id in ids {
        relationships
            .try_add_relationship(
                reltype.to_string(),
                target.to_string(),
                (*id).to_string(),
                mode,
            )
            .expect("distinct relationship identifiers");
    }
    relationships
}

#[test]
fn external_reuse_picks_the_same_relationship_on_every_build() {
    for order in [["rId1", "rId2"], ["rId2", "rId1"]] {
        let mut seen = BTreeSet::new();
        for _ in 0..REPEATS {
            let mut relationships = duplicates(
                &order,
                HYPERLINK,
                "http://www.apache.org/",
                TargetMode::External,
            );
            seen.insert(relationships.get_or_add_ext_rel(HYPERLINK, "http://www.apache.org/"));
        }
        assert_eq!(
            seen,
            BTreeSet::from(["rId1".to_string()]),
            "get_or_add_ext_rel reused a hash-order-dependent rId for insertion order {order:?}"
        );
    }
}

/// Reuse must not depend on which of the duplicates was inserted first
/// either, so the two insertion orders have to agree with each other and with
/// the canonical answer across repeated builds. `drawing.docx` carries this
/// pair (rId21 and rId23 both point at `http://www.regnum.ru/`).
#[test]
fn external_reuse_is_independent_of_insertion_order() {
    let mut seen = BTreeSet::new();
    for _ in 0..REPEATS {
        for order in [["rId21", "rId23"], ["rId23", "rId21"]] {
            let mut relationships = duplicates(
                &order,
                HYPERLINK,
                "http://www.regnum.ru/",
                TargetMode::External,
            );
            seen.insert(relationships.get_or_add_ext_rel(HYPERLINK, "http://www.regnum.ru/"));
        }
    }
    assert_eq!(
        seen,
        BTreeSet::from(["rId21".to_string()]),
        "reuse depends on the insertion order of the duplicate pair"
    );
}

#[test]
fn internal_reuse_picks_the_same_relationship_on_every_build() {
    let mut seen = BTreeSet::new();
    for _ in 0..REPEATS {
        let mut relationships = duplicates(
            &["rId4", "rId7", "rId9"],
            IMAGE,
            "../media/image1.png",
            TargetMode::Internal,
        );
        seen.insert(
            relationships
                .get_or_add(IMAGE, "../media/image1.png")
                .r_id()
                .to_string(),
        );
    }
    assert_eq!(
        seen,
        BTreeSet::from(["rId4".to_string()]),
        "get_or_add reused a hash-order-dependent rId"
    );
}

#[test]
fn reuse_of_a_unique_match_is_unchanged() {
    let mut relationships = Relationships::new("/word".to_string());
    relationships
        .try_add_relationship(
            IMAGE.to_string(),
            "media/image1.png".to_string(),
            "rId5".to_string(),
            TargetMode::Internal,
        )
        .expect("add");
    relationships
        .try_add_relationship(
            IMAGE.to_string(),
            "media/image2.png".to_string(),
            "rId6".to_string(),
            TargetMode::Internal,
        )
        .expect("add");
    assert_eq!(
        relationships.get_or_add(IMAGE, "media/image2.png").r_id(),
        "rId6"
    );
    assert_eq!(relationships.len(), 2);
}

#[test]
fn a_fresh_target_still_takes_the_next_free_identifier() {
    let mut relationships = duplicates(
        &["rId1", "rId2"],
        HYPERLINK,
        "http://www.apache.org/",
        TargetMode::External,
    );
    assert_eq!(
        relationships.get_or_add_ext_rel(HYPERLINK, "http://example.invalid/"),
        "rId3"
    );
    assert_eq!(relationships.len(), 3);
}

#[test]
fn an_external_duplicate_does_not_satisfy_an_internal_reuse() {
    let mut relationships = duplicates(
        &["rId1", "rId2"],
        HYPERLINK,
        "http://www.apache.org/",
        TargetMode::External,
    );
    // `get_or_add` only reuses internal relationships, so the external pair is
    // invisible to it and a new internal relationship is created.
    let created = relationships
        .get_or_add(HYPERLINK, "http://www.apache.org/")
        .r_id()
        .to_string();
    assert_eq!(created, "rId3");
    assert_eq!(
        relationships.get(&created).expect("created").target_mode(),
        TargetMode::Internal
    );
}

/// The four corpus packages whose `.rels` members carry a duplicate
/// (Type, Target, TargetMode) group, with the part that owns it.
const CORPUS: &[(&str, &str, &str, &str)] = &[
    (
        "ooxml/xlsx/sharedhyperlink.xlsx",
        "/xl/worksheets/sheet1.xml",
        "http://www.apache.org/",
        "rId1",
    ),
    (
        "poi/test-data/openxml4j/50154.xlsx",
        "/xl/worksheets/sheet1.xml",
        "..\\..\\..\\..\\..\\..\\..\\cygwin\\home\\yegor\\dinom\\%5baccess%5d.2010-10-26.log",
        "rId2",
    ),
    (
        "ooxml/docx/drawing.docx",
        "/word/document.xml",
        "http://www.regnum.ru/",
        "rId21",
    ),
    (
        "libreoffice-core/sw/qa/extras/ooxmlexport/data/mailmerge.docx",
        "/word/settings.xml",
        "file:///D:\\bug\\data%20source.xls",
        "rId1",
    ),
];

#[test]
fn corpus_packages_reuse_the_same_hyperlink_relationship_on_every_open() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    for (relative, owner, target, expected) in CORPUS {
        let path = root.join(relative);
        let bytes = std::fs::read(&path).unwrap_or_else(|error| {
            panic!("read corpus package {}: {error}", path.display());
        });
        let partname = PackURI::new(*owner).expect("part name");
        let reltype = if relative.ends_with("mailmerge.docx") {
            MAIL_MERGE_SOURCE
        } else {
            HYPERLINK
        };
        let mut seen = BTreeSet::new();
        for _ in 0..16 {
            let mut package = OpcPackage::from_bytes(&bytes).expect("open corpus package");
            let part = package.get_part_mut(&partname).expect("owner part");
            seen.insert(part.relate_to_ext(target, reltype));
        }
        assert_eq!(
            seen,
            BTreeSet::from([(*expected).to_string()]),
            "{relative}: reusing the duplicate relationship on {owner} is not deterministic"
        );
    }
}

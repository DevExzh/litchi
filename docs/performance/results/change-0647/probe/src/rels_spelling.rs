//! Change 0647 probe A: how far the source spelling of a `.rels` member is
//! from the canonical serialization change 0593 captures at open, and how
//! large the reuse surface of `Relationships::get_or_add` is.
//!
//! For every OOXML fixture this prints, per relationships member that the
//! source archive actually carries:
//!
//! * `src` — the member's decompressed source bytes;
//! * `can` — `Relationships::try_to_xml_bytes` for the collection parsed from
//!   it, which is exactly the `CanonicalRelationshipsXml` change 0593 binds as
//!   the open-time capture;
//! * `int` — the number of *internal* relationships in the collection, which
//!   is the number of distinct identifiers a reusing `get_or_add` can return
//!   for that member;
//! * `tri` — the number of distinct internal (type, target) pairs, which is
//!   the number of reuse calls probe B has to make for that member.
//!
//! The `SPELLING` verdict is `same` when the source bytes equal the canonical
//! bytes and `differs` otherwise. Change 0647's design question is what
//! happens to a member whose spelling differs when a no-op reuse keeps the
//! open-time capture instead of dropping it.
//!
//! Usage: rels_spelling <test-data-root>
mod corpus;

use litchi_opc::{OpcPackage, PackURI};
use std::collections::BTreeSet;
use std::path::PathBuf;

fn main() {
    let root = PathBuf::from(std::env::args().nth(1).expect("test-data root"));
    let files = corpus::fixtures(&root);

    let mut fixtures_opened = 0usize;
    let mut members_total = 0usize;
    let mut members_same = 0usize;
    let mut members_differ = 0usize;
    let mut internal_total = 0usize;
    let mut triples_total = 0usize;
    let mut internal_on_differing = 0usize;
    let mut triples_on_differing = 0usize;
    let mut max_rels_in_member = 0usize;
    let mut fixtures_with_any_reuse = 0usize;

    for path in &files {
        let relative = path
            .strip_prefix(&root)
            .unwrap_or(path)
            .display()
            .to_string();
        let Ok(bytes) = std::fs::read(path) else {
            println!("{relative}\tREAD-ERROR");
            continue;
        };
        let package = match OpcPackage::from_vec(bytes.clone()) {
            Ok(package) => package,
            Err(error) => {
                println!("{relative}\tOPEN-REFUSED\t{error}");
                continue;
            },
        };
        let physical = match litchi_opc::phys_pkg::PhysPkgReader::new(&bytes) {
            Ok(physical) => physical,
            Err(error) => {
                println!("{relative}\tZIP-REFUSED\t{error}");
                continue;
            },
        };
        fixtures_opened += 1;

        // (member name, canonical bytes, internal count, distinct internal
        // (type, target) count), package relationships first then parts in
        // partname order so the listing is a function of the package.
        let mut rows: Vec<(String, Vec<u8>, usize, usize)> = Vec::new();
        let package_uri = PackURI::new("/").expect("package uri");
        let mut collections: Vec<(PackURI, &litchi_opc::Relationships)> =
            vec![(package_uri, package.rels())];
        let mut parts: Vec<&dyn litchi_opc::Part> = package.iter_parts().collect();
        parts.sort_unstable_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
        for part in parts {
            collections.push((part.partname().clone(), part.rels()));
        }
        for (owner, rels) in collections {
            let Ok(rels_uri) = owner.rels_uri() else {
                continue;
            };
            // `Relationships::try_to_xml_bytes`, the fallible form change 0593
            // captures, is crate-private; `to_xml` is its public twin and the
            // crate's own `rel.rs:973` test asserts the two produce the same
            // bytes for the same collection.
            let canonical = rels.to_xml().into_bytes();
            let internal = rels.iter().filter(|rel| !rel.is_external()).count();
            let triples: BTreeSet<(&str, &str)> = rels
                .iter()
                .filter(|rel| !rel.is_external())
                .map(|rel| (rel.reltype(), rel.target_ref()))
                .collect();
            rows.push((
                rels_uri.membername().to_string(),
                canonical,
                internal,
                triples.len(),
            ));
        }

        let mut fixture_members = 0usize;
        let mut fixture_reuse = 0usize;
        for (membername, canonical, internal, triples) in rows {
            let Ok(source) = physical.read_member(&membername) else {
                // No member in the source archive: change 0593 binds no
                // capture for it, so change 0647 cannot reach it either.
                continue;
            };
            members_total += 1;
            fixture_members += 1;
            internal_total += internal;
            triples_total += triples;
            fixture_reuse += triples;
            max_rels_in_member = max_rels_in_member.max(canonical_rel_count(&source));
            let verdict = if source == canonical {
                members_same += 1;
                "same"
            } else {
                members_differ += 1;
                internal_on_differing += internal;
                triples_on_differing += triples;
                "differs"
            };
            println!(
                "{relative}\t{membername}\tSPELLING={verdict}\tsrc={}\tcan={}\tint={internal}\ttri={triples}",
                source.len(),
                canonical.len()
            );
        }
        if fixture_reuse > 0 {
            fixtures_with_any_reuse += 1;
        }
        println!(
            "#FIXTURE\t{relative}\tmembers={fixture_members}\treuse_pairs={fixture_reuse}"
        );
    }

    println!("#TOTAL files={}", files.len());
    println!("#TOTAL fixtures_opened={fixtures_opened}");
    println!("#TOTAL fixtures_with_any_reuse={fixtures_with_any_reuse}");
    println!("#TOTAL rels_members_present={members_total}");
    println!("#TOTAL spelling_same={members_same}");
    println!("#TOTAL spelling_differs={members_differ}");
    println!("#TOTAL internal_relationships={internal_total}");
    println!("#TOTAL internal_type_target_pairs={triples_total}");
    println!("#TOTAL internal_relationships_on_differing_members={internal_on_differing}");
    println!("#TOTAL internal_pairs_on_differing_members={triples_on_differing}");
    println!("#TOTAL max_relationship_elements_in_one_member={max_rels_in_member}");
}

/// Count `<Relationship` elements in a raw source member.
///
/// Used only to report how far the corpus is from the roughly 62,500
/// relationships at which the publication auditor's aggregate attribute
/// ceiling would refuse the canonical serialization (change 0593).
fn canonical_rel_count(source: &[u8]) -> usize {
    let needle = b"<Relationship ";
    source
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

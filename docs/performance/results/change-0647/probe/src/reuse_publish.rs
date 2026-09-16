//! Change 0647 probe B: the published-bytes oracle for a no-op reusing
//! `Relationships::get_or_add`.
//!
//! For every OOXML fixture, every relationship collection (the package's own
//! and every part's), and every distinct internal (type, target) pair the
//! collection already carries, this performs two owned-source publications:
//!
//! * `baseline` — take the same mutable seam (`OpcPackage::rels_mut` for the
//!   package, `OpcPackage::get_part_mut` for a part) and publish without
//!   touching the collection. Taking the seam is what revokes exact-source
//!   authorization, so both legs publish through the same route;
//! * `reuse` — take the seam and call `get_or_add(type, target)` with a pair
//!   the collection already carries, which must return an established
//!   identifier and change nothing, then publish.
//!
//! Every line records whether the two publications are byte-identical, and the
//! per-fixture summary carries rolling SHA-256 digests over all baseline and
//! all reuse outputs so that a cross-leg `diff` of two runs is a complete
//! oracle: if any published byte moved between the before and the after leg,
//! the digest for that fixture moves.
//!
//! `reuse_fired` counts the calls that actually returned an established
//! identifier without growing the collection; `unexpected_growth` counts any
//! call that did not, and must stay at zero.
//!
//! Usage: reuse_publish <test-data-root> [--parts-only]
mod corpus;

use litchi_opc::{OpcPackage, PackURI, PackageWriter};
use sha2::{Digest, Sha256};
use std::collections::BTreeSet;
use std::path::PathBuf;

fn publish(package: &OpcPackage) -> std::result::Result<Vec<u8>, String> {
    PackageWriter::to_bytes(package).map_err(|error| error.to_string())
}

fn digest(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    corpus::hex(&hasher.finalize())
}

/// One publication outcome: the digest of the published bytes, or the typed
/// refusal's rendered message. Refusals are outcomes, not failures.
fn outcome(result: &std::result::Result<Vec<u8>, String>) -> String {
    match result {
        Ok(bytes) => format!("sha256={};len={}", digest(bytes), bytes.len()),
        Err(error) => format!("refused={error}"),
    }
}

fn main() {
    let root = PathBuf::from(std::env::args().nth(1).expect("test-data root"));
    let parts_only = std::env::args().any(|arg| arg == "--parts-only");
    let files = corpus::fixtures(&root);

    let mut fixtures_opened = 0usize;
    let mut owners_probed = 0usize;
    let mut pairs_probed = 0usize;
    let mut pairs_same = 0usize;
    let mut pairs_diff = 0usize;
    let mut reuse_fired = 0usize;
    let mut unexpected_growth = 0usize;
    let mut baseline_refusals = 0usize;
    let mut reuse_refusals = 0usize;
    let mut refusal_kinds: std::collections::BTreeMap<String, usize> = Default::default();

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
        fixtures_opened += 1;

        // Snapshot every reuse target as owned strings, in a deterministic
        // order, then drop the reading package.
        let mut owners: Vec<(Option<String>, Vec<(String, String)>, usize)> = Vec::new();
        if !parts_only {
            let pairs: BTreeSet<(String, String)> = package
                .rels()
                .iter()
                .filter(|rel| !rel.is_external())
                .map(|rel| (rel.reltype().to_string(), rel.target_ref().to_string()))
                .collect();
            if !pairs.is_empty() {
                owners.push((None, pairs.into_iter().collect(), package.rels().len()));
            }
        }
        let mut parts: Vec<&dyn litchi_opc::Part> = package.iter_parts().collect();
        parts.sort_unstable_by(|left, right| {
            left.partname().as_str().cmp(right.partname().as_str())
        });
        for part in parts {
            let pairs: BTreeSet<(String, String)> = part
                .rels()
                .iter()
                .filter(|rel| !rel.is_external())
                .map(|rel| (rel.reltype().to_string(), rel.target_ref().to_string()))
                .collect();
            if pairs.is_empty() {
                continue;
            }
            owners.push((
                Some(part.partname().as_str().to_string()),
                pairs.into_iter().collect(),
                part.rels().len(),
            ));
        }
        drop(package);

        let mut baseline_roll = Sha256::new();
        let mut reuse_roll = Sha256::new();
        let mut fixture_pairs = 0usize;
        let mut fixture_diff = 0usize;

        for (owner, pairs, owner_len) in &owners {
            owners_probed += 1;
            // Baseline: the same seam, no mutation of the collection.
            let baseline = {
                let mut package = match OpcPackage::from_vec(bytes.clone()) {
                    Ok(package) => package,
                    Err(error) => {
                        println!("{relative}\tREOPEN-REFUSED\t{error}");
                        break;
                    },
                };
                match owner {
                    None => {
                        let _seam = package.rels_mut();
                    },
                    Some(partname) => {
                        let uri = match PackURI::new(partname) {
                            Ok(uri) => uri,
                            Err(error) => {
                                println!("{relative}\t{partname}\tURI-REFUSED\t{error}");
                                continue;
                            },
                        };
                        if let Err(error) = package.get_part_mut(&uri) {
                            println!("{relative}\t{partname}\tSEAM-REFUSED\t{error}");
                            continue;
                        }
                    },
                }
                publish(&package)
            };
            if let Err(error) = &baseline {
                baseline_refusals += 1;
                *refusal_kinds.entry(format!("baseline: {error}")).or_default() += 1;
            }
            let baseline_outcome = outcome(&baseline);
            baseline_roll.update(baseline_outcome.as_bytes());

            for (reltype, target) in pairs {
                pairs_probed += 1;
                fixture_pairs += 1;
                let mut package = match OpcPackage::from_vec(bytes.clone()) {
                    Ok(package) => package,
                    Err(error) => {
                        println!("{relative}\tREOPEN-REFUSED\t{error}");
                        break;
                    },
                };
                let (r_id, new_len) = match owner {
                    None => {
                        let rels = package.rels_mut();
                        let r_id = rels.get_or_add(reltype, target).r_id().to_string();
                        (r_id, package.rels().len())
                    },
                    Some(partname) => {
                        let uri = match PackURI::new(partname) {
                            Ok(uri) => uri,
                            Err(_) => continue,
                        };
                        let r_id = match package.get_part_mut(&uri) {
                            Ok(part) => {
                                part.rels_mut().get_or_add(reltype, target).r_id().to_string()
                            },
                            Err(_) => continue,
                        };
                        let new_len = match package.get_part(&uri) {
                            Ok(part) => part.rels().len(),
                            Err(_) => continue,
                        };
                        (r_id, new_len)
                    },
                };
                if new_len == *owner_len {
                    reuse_fired += 1;
                } else {
                    unexpected_growth += 1;
                }
                let reuse = publish(&package);
                if let Err(error) = &reuse {
                    reuse_refusals += 1;
                    *refusal_kinds.entry(format!("reuse: {error}")).or_default() += 1;
                }
                let reuse_outcome = outcome(&reuse);
                reuse_roll.update(reuse_outcome.as_bytes());
                let verdict = if reuse_outcome == baseline_outcome {
                    pairs_same += 1;
                    "SAME"
                } else {
                    pairs_diff += 1;
                    fixture_diff += 1;
                    "DIFF"
                };
                let owner_label = owner.as_deref().unwrap_or("/");
                if verdict == "DIFF" {
                    println!(
                        "{relative}\t{owner_label}\t{r_id}\tDIFF\tbaseline[{baseline_outcome}]\treuse[{reuse_outcome}]"
                    );
                } else {
                    let kind = if reuse.is_ok() { "published" } else { "refused" };
                    println!("{relative}\t{owner_label}\t{r_id}\tSAME\tlen={new_len}\t{kind}");
                }
            }
        }

        println!(
            "#FIXTURE\t{relative}\towners={}\tpairs={fixture_pairs}\tdiff={fixture_diff}\tbaselines={}\treuses={}",
            owners.len(),
            corpus::hex(&baseline_roll.finalize()),
            corpus::hex(&reuse_roll.finalize())
        );
    }

    println!("#TOTAL files={}", files.len());
    println!("#TOTAL fixtures_opened={fixtures_opened}");
    println!("#TOTAL owners_probed={owners_probed}");
    println!("#TOTAL pairs_probed={pairs_probed}");
    println!("#TOTAL pairs_same={pairs_same}");
    println!("#TOTAL pairs_diff={pairs_diff}");
    println!("#TOTAL reuse_fired={reuse_fired}");
    println!("#TOTAL unexpected_growth={unexpected_growth}");
    println!("#TOTAL baseline_refusals={baseline_refusals}");
    println!("#TOTAL reuse_refusals={reuse_refusals}");
    for (kind, count) in &refusal_kinds {
        println!("#REFUSAL {count}\t{kind}");
    }
}

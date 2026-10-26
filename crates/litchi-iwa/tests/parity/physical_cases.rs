//! Raw physical access moved to `litchi_iwa_archive::unstable` (record 0363
//! D5, wave W0-A, record 0367).
//!
//! The host `Package::{entry_names, iwa_entry_names, entry, archive}` and
//! `Document::stats` with `DocumentStats::{most_common_message_type,
//! message_type_summary}` delegate to the archive-owned functions. These
//! read-only cases run on native-saved fixtures and compare every host read
//! with the same read through the unstable functions over a package state
//! built from the leaf `SourceCatalog`.

use std::collections::HashMap;

use litchi_iwa_archive::package::{PackageEntry, PackageState};
use litchi_iwa_archive::unstable;
use litchi_iwa_archive::{Limits, SourceCatalog};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Everything the D5 routes read from one package, through one side.
#[derive(Debug, PartialEq)]
struct PhysicalReads {
    entry_names: Vec<String>,
    iwa_entry_names: Vec<String>,
    entries: Vec<(String, Option<usize>, u64)>,
    archives: Vec<(String, Vec<Option<u64>>, usize)>,
    message_types: Vec<(u32, usize)>,
    most_common: Option<(u32, usize)>,
    summary: String,
}

fn digest(bytes: &[u8]) -> u64 {
    use std::hash::Hasher;
    let mut hasher = std::hash::DefaultHasher::new();
    hasher.write(bytes);
    hasher.finish()
}

fn sorted_counts(counts: &HashMap<u32, usize>) -> Vec<(u32, usize)> {
    let mut counts: Vec<_> = counts.iter().map(|(&key, &value)| (key, value)).collect();
    counts.sort_unstable();
    counts
}

fn host_reads(bytes: &[u8]) -> TestResult<PhysicalReads> {
    let package = litchi_iwa::Package::from_bytes(bytes)?;
    let entry_names: Vec<String> = package.entry_names().map(str::to_owned).collect();
    let iwa_entry_names: Vec<String> = package.iwa_entry_names().map(str::to_owned).collect();
    let mut entries = Vec::new();
    for name in &entry_names {
        let data = package.entry(name);
        entries.push((name.clone(), data.map(<[u8]>::len), data.map_or(0, digest)));
    }
    let mut archives = Vec::new();
    for name in &iwa_entry_names {
        let archive = package.archive(name)?;
        archives.push((
            name.clone(),
            archive
                .objects
                .iter()
                .map(|object| object.archive_info.identifier)
                .collect(),
            archive
                .objects
                .iter()
                .map(|object| object.messages.len())
                .sum(),
        ));
    }
    let stats = litchi_iwa::Document::from_bytes(bytes)?.stats()?;
    Ok(PhysicalReads {
        entry_names,
        iwa_entry_names,
        entries,
        archives,
        message_types: sorted_counts(&stats.message_type_counts),
        most_common: stats.most_common_message_type(),
        summary: stats.message_type_summary(),
    })
}

fn archive_reads(bytes: &[u8]) -> TestResult<PhysicalReads> {
    let catalog = SourceCatalog::from_bytes(bytes, Limits::default())?;
    let mut members = Vec::new();
    for entry in catalog.package().iter() {
        members.push(PackageEntry::new(
            entry.name().to_owned(),
            entry.data().to_vec(),
        ));
    }
    let state = PackageState::from_entries(members, Default::default())?;
    let entry_names: Vec<String> = unstable::entry_names(&state).map(str::to_owned).collect();
    let iwa_entry_names: Vec<String> = unstable::iwa_entry_names(&state)
        .map(str::to_owned)
        .collect();
    let mut entries = Vec::new();
    for name in &entry_names {
        let data = unstable::entry(&state, name);
        entries.push((name.clone(), data.map(<[u8]>::len), data.map_or(0, digest)));
    }
    let mut archives = Vec::new();
    let mut types = Vec::new();
    for name in &iwa_entry_names {
        let archive = unstable::cached_component(&state, name, Default::default())
            .map_err(|error| error.to_string())?;
        archives.push((
            name.clone(),
            archive
                .objects
                .iter()
                .map(|object| object.archive_info.identifier)
                .collect(),
            archive
                .objects
                .iter()
                .map(|object| object.messages.len())
                .sum(),
        ));
        types.extend(
            archive
                .objects
                .iter()
                .flat_map(|object| object.messages.iter().map(|message| message.type_)),
        );
    }
    let counts = unstable::count_message_types(types);
    Ok(PhysicalReads {
        entry_names,
        iwa_entry_names,
        entries,
        archives,
        message_types: sorted_counts(&counts),
        most_common: unstable::most_common_message_type(&counts),
        summary: unstable::message_type_summary(&counts),
    })
}

/// One read that is true only when the host and archive-owned reads agree,
/// so a divergence fails the case on both outputs.
fn reads_agree(bytes: &[u8]) -> TestResult<bool> {
    let host = host_reads(bytes)?;
    let archive = archive_reads(bytes)?;
    if host != archive {
        return Err(format!(
            "host and archive reads differ:\nhost {host:#?}\narchive {archive:#?}"
        )
        .into());
    }
    Ok(!host.entry_names.is_empty() && !host.archives.is_empty())
}

fn physical_case(
    name: &'static str,
    format: Format,
    input: &Fixture,
    tier: Tier,
    evidence: Evidence,
) -> TestResult {
    Case::new(name, format, tier, evidence, input)
        .read_only()
        .host(|source| Ok(litchi_iwa::Package::from_bytes(source)?.to_bytes()?))
        .leaf(move |source| {
            let bytes = match format {
                Format::Numbers => leaf_bytes!(litchi_numbers::Package::from_bytes(source)?),
                Format::Pages => leaf_bytes!(litchi_pages::Package::from_bytes(source)?),
                Format::Keynote => leaf_bytes!(litchi_keynote::Package::from_bytes(source)?),
            };
            Ok(LeafRun::output(bytes))
        })
        .read("host physical reads", host_reads)
        .read("archive unstable physical reads", archive_reads)
        .read("host and archive reads agree", reads_agree)
        .run()?;
    Ok(())
}

#[test]
fn numbers_physical_access_parity() -> TestResult {
    let input = fixture("iwork/numbers/table-merges-native.numbers")?;
    physical_case(
        "numbers_physical_access_parity",
        Format::Numbers,
        &input,
        Tier::A,
        Evidence::Native("test-data/iwork/numbers/table-merges-native-receipt.json"),
    )
}

#[test]
fn pages_physical_access_parity() -> TestResult {
    let input = fixture("iwork/pages/body-table-merges-native.pages")?;
    physical_case(
        "pages_physical_access_parity",
        Format::Pages,
        &input,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-merges-native-receipt.json"),
    )
}

#[test]
fn keynote_physical_access_parity() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-merges-native.key")?;
    physical_case(
        "keynote_physical_access_parity",
        Format::Keynote,
        &input,
        Tier::A,
        Evidence::Native("test-data/iwork/keynote/slide-table-merges-native-receipt.json"),
    )
}

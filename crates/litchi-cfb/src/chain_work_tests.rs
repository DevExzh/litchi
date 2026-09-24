//! Change 0767: the reader's reusable chain maps keep every verdict and
//! clear only what a walk set.
//!
//! - Random call sequences hold both reusable chain collectors to the
//!   allocating helpers they replace, call for call, and require an all-clear
//!   map after every call.
//! - Stream-allocation validation (A5 in record 0749) is held, on fault
//!   injected states, to a copy of its per-stream fresh-map form.
//! - Reading every stream through one reader, whose chain buffers persist,
//!   is held to a fresh reader per stream.
//! - The work tests count the words the chain maps write and bound them by
//!   the tables plus the chains: the work per stream does not grow with the
//!   stream count.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::writer::OleWriter;
use std::io::Cursor;

/// A small deterministic generator (xorshift64*), so every case replays.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Self {
        Self(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1)
    }

    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value >> 12;
        value ^= value << 25;
        value ^= value >> 27;
        self.0 = value;
        value.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, bound: usize) -> usize {
        if bound == 0 {
            return 0;
        }
        usize::try_from(self.next() % u64::try_from(bound).unwrap()).unwrap()
    }

    fn u32_below(&mut self, bound: u32) -> u32 {
        u32::try_from(self.below(usize::try_from(bound).unwrap())).unwrap()
    }
}

fn outcome<T: Clone>(result: Result<T, OleError>) -> Result<T, String> {
    result.map_err(|error| error.to_string())
}

fn all_clear(map: &CheckedBitSet) -> bool {
    map.words.iter().all(|&word| word == 0)
}

/// A table of chains over a shuffled set of sectors, with a few links
/// corrupted: joined into another chain, looped back, cut short, freed, set
/// to a marker, or pointed outside the table. Returns the table and each
/// chain's head and length before corruption.
fn random_table(rng: &mut Rng) -> (Vec<u32>, Vec<(u32, usize)>) {
    let len = 1 + rng.below(300);
    let mut order: Vec<u32> = (0..u32::try_from(len).unwrap()).collect();
    for index in (1..order.len()).rev() {
        order.swap(index, rng.below(index + 1));
    }
    let mut table = vec![FREESECT; len];
    let mut chains = Vec::new();
    let mut position = 0;
    // Leave some sectors free, as a real table does.
    let used = len - rng.below(len / 4 + 1);
    while position < used {
        let chain_len = (1 + rng.below(40)).min(used - position);
        let chain = &order[position..position + chain_len];
        for pair in chain.windows(2) {
            table[usize::try_from(pair[0]).unwrap()] = pair[1];
        }
        table[usize::try_from(chain[chain_len - 1]).unwrap()] = ENDOFCHAIN;
        chains.push((chain[0], chain_len));
        position += chain_len;
    }
    for _ in 0..rng.below(4) {
        let slot = rng.below(len);
        let len_u32 = u32::try_from(len).unwrap();
        table[slot] = match rng.below(8) {
            // Join another chain, or loop back into this one.
            0..=2 => order[rng.below(used.max(1)).min(len - 1)],
            3 => ENDOFCHAIN,
            4 => FREESECT,
            5 => [FATSECT, DIFSECT, MAXREGSECT, MAXREGSECT + 1][rng.below(4)],
            6 => len_u32 + rng.u32_below(3),
            _ => u32::try_from(slot).unwrap(),
        };
    }
    (table, chains)
}

/// A start sector for one query: usually a chain head, sometimes anything.
fn random_start(rng: &mut Rng, table: &[u32], chains: &[(u32, usize)]) -> (u32, usize) {
    let len_u32 = u32::try_from(table.len()).unwrap();
    match rng.below(10) {
        0..=6 if !chains.is_empty() => chains[rng.below(chains.len())],
        7 => (rng.u32_below(len_u32), 1 + rng.below(8)),
        8 => (
            [ENDOFCHAIN, FREESECT, MAXREGSECT, len_u32][rng.below(4)],
            rng.below(3),
        ),
        _ => (0, rng.below(table.len() + 2)),
    }
}

#[test]
fn sector_chain_scratch_matches_fresh_maps_over_random_call_sequences() {
    let mut rng = Rng::new(0x0767_0001);
    let mut scratch = SectorChainScratch::default();
    let (mut accepted, mut refused) = (0usize, 0usize);
    for case in 0..400 {
        // Tables of different lengths through one scratch, so a short table
        // follows a long one and must not see its stale words.
        let (table, chains) = random_table(&mut rng);
        for query in 0..12 {
            let (start, chain_len) = random_start(&mut rng, &table, &chains);
            // The declared length: the chain's own, or one off either way.
            let expected_count = match rng.below(4) {
                0 => chain_len.saturating_sub(1),
                1 => chain_len + 1,
                _ => chain_len,
            };
            let oracle = outcome(collect_sector_chain_exact(
                &table,
                start,
                expected_count,
                "regular stream",
            ));
            let actual = outcome(
                scratch
                    .collect_exact(&table, start, expected_count, "regular stream")
                    .map(|()| scratch.sectors().to_vec()),
            );
            assert_eq!(actual, oracle, "case {case} query {query}");
            assert!(
                all_clear(&scratch.visited),
                "case {case} query {query} leaves visited bits behind"
            );
            if oracle.is_ok() {
                accepted += 1;
            } else {
                refused += 1;
                assert!(scratch.sectors().is_empty());
                assert_eq!(scratch.visited.bit_len, 0);
            }
        }
    }
    // Both verdicts are well represented.
    assert!(accepted > 800 && refused > 800, "{accepted} / {refused}");
}

#[test]
fn end_chain_scratch_matches_fresh_maps_over_random_call_sequences() {
    let mut rng = Rng::new(0x0767_0002);
    let mut scratch = EndChainScratch::default();
    let (mut accepted, mut refused) = (0usize, 0usize);
    for case in 0..400 {
        let (table, chains) = random_table(&mut rng);
        for query in 0..12 {
            let (start, _) = random_start(&mut rng, &table, &chains);
            let oracle = outcome(collect_sector_chain(&table, start, "FAT"));
            let actual = outcome(
                scratch
                    .collect(&table, start, "FAT")
                    .map(|()| scratch.sectors().to_vec()),
            );
            assert_eq!(actual, oracle, "case {case} query {query}");
            assert!(
                all_clear(&scratch.visited),
                "case {case} query {query} leaves visited bits behind"
            );
            if oracle.is_ok() {
                accepted += 1;
            } else {
                refused += 1;
                assert!(scratch.sectors().is_empty());
            }
        }
    }
    assert!(accepted > 800 && refused > 800, "{accepted} / {refused}");
}

/// A compound file of `sizes.len()` streams, the i-th of `sizes[i]` bytes;
/// every fourth stream sits in one of two storages, so the directory has more
/// than one sibling tree.
fn build_file(sector_size: usize, sizes: &[usize]) -> Vec<u8> {
    let mut writer = OleWriter::with_sector_size(sector_size).unwrap();
    for (index, &size) in sizes.iter().enumerate() {
        let payload: Vec<u8> = (0..size)
            .map(|byte| u8::try_from((byte * 7 + index * 13) % 251).unwrap())
            .collect();
        let name = format!("S{index:05}");
        match index % 8 {
            3 => writer
                .create_stream_owned(&["Left", name.as_str()], payload)
                .unwrap(),
            7 => writer
                .create_stream_owned(&["Right", name.as_str()], payload)
                .unwrap(),
            _ => writer
                .create_stream_owned(&[name.as_str()], payload)
                .unwrap(),
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Stream sizes that mix empty, mini and regular streams.
fn mixed_sizes(rng: &mut Rng, count: usize) -> Vec<usize> {
    (0..count)
        .map(|_| match rng.below(6) {
            0 => 0,
            1..=3 => 1 + rng.below(4095),
            _ => 4096 + rng.below(12_000),
        })
        .collect()
}

/// The state `validate_stream_allocations` starts from: the roles it claims
/// released and the root chain it records cleared.
fn before_allocation_validation(file: &mut OleFile<Cursor<Vec<u8>>>) {
    for role in &mut file.sector_roles {
        if matches!(
            role,
            PhysicalSectorRole::MiniStream | PhysicalSectorRole::RegularStream
        ) {
            *role = PhysicalSectorRole::Unclaimed;
        }
    }
    file.root_chain.clear();
}

/// `validate_stream_allocations` as it was before this change: a fresh
/// table-sized map for every stream, through the owned exact helper, which
/// the base's scratch was tested to match.
fn fresh_map_allocation_validation(file: &mut OleFile<Cursor<Vec<u8>>>) -> Result<(), OleError> {
    let root = file
        .root
        .as_ref()
        .ok_or_else(|| OleError::CorruptedFile("Missing root directory entry".to_string()))?;
    let root_start = root.start_sector;
    let root_size = root.size;
    if file.sector_size == SECTOR_SIZE_V4 && root_start == RANGE_LOCK_SECTOR_V4 {
        return Err(OleError::CorruptedFile(
            "root mini stream points to the range-lock sector".to_string(),
        ));
    }
    let root_sector_count = usize::try_from(root_size.div_ceil(file.sector_size as u64))
        .map_err(|_err| OleError::CorruptedFile("Root mini stream is too large".to_string()))?;
    let root_chain =
        collect_sector_chain_exact(&file.fat, root_start, root_sector_count, "root mini stream")?;
    file.claim_chain(&root_chain, PhysicalSectorRole::MiniStream)?;
    file.root_chain = root_chain;
    let mini_sector_capacity = usize::try_from(root_size.div_ceil(file.mini_sector_size as u64))
        .map_err(|_err| OleError::CorruptedFile("Root mini stream is too large".to_string()))?;
    let mut claimed_mini_sectors =
        CheckedBitSet::try_with_capacity(mini_sector_capacity, "mini-sector ownership map")?;
    for index in 0..file.dir_entries.len() {
        let Some(entry) = file.dir_entries[index].as_ref() else {
            continue;
        };
        if entry.entry_type != STGTY_STREAM {
            continue;
        }
        let (is_minifat, start_sector, size) = (entry.is_minifat, entry.start_sector, entry.size);
        if file.sector_size == SECTOR_SIZE_V4 && !is_minifat && start_sector == RANGE_LOCK_SECTOR_V4
        {
            return Err(OleError::CorruptedFile(
                "regular stream points to the range-lock sector".to_string(),
            ));
        }
        if is_minifat {
            let sector_count = usize::try_from(size.div_ceil(file.mini_sector_size as u64))
                .map_err(|_err| OleError::CorruptedFile("Mini stream is too large".to_string()))?;
            let chain = collect_sector_chain_exact(
                &file.minifat,
                start_sector,
                sector_count,
                "mini stream",
            )?;
            for &sector in &chain {
                let sector_index = usize::try_from(sector).map_err(|_err| {
                    OleError::CorruptedFile(
                        "mini stream sector index does not fit usize".to_string(),
                    )
                })?;
                if sector_index >= mini_sector_capacity {
                    return Err(OleError::CorruptedFile(
                        "Mini stream references storage outside the root mini stream".to_string(),
                    ));
                }
                if claimed_mini_sectors.contains(sector_index) {
                    return Err(OleError::CorruptedFile(format!(
                        "Mini sector {sector} is claimed by multiple streams"
                    )));
                }
                claimed_mini_sectors.insert(sector_index)?;
            }
        } else {
            let sector_count =
                usize::try_from(size.div_ceil(file.sector_size as u64)).map_err(|_err| {
                    OleError::CorruptedFile("Regular stream is too large".to_string())
                })?;
            let chain = collect_sector_chain_exact(
                &file.fat,
                start_sector,
                sector_count,
                "regular stream",
            )?;
            file.claim_chain(&chain, PhysicalSectorRole::RegularStream)?;
        }
    }
    Ok(())
}

/// The streams of a parsed file: SID, whether mini, start sector and size.
fn stream_entries(file: &OleFile<Cursor<Vec<u8>>>) -> Vec<(usize, bool, u32, u64)> {
    file.dir_entries
        .iter()
        .enumerate()
        .filter_map(|(sid, entry)| {
            let entry = entry.as_ref()?;
            (entry.entry_type == STGTY_STREAM).then_some((
                sid,
                entry.is_minifat,
                entry.start_sector,
                entry.size,
            ))
        })
        .collect()
}

/// A chain as the tables hold it, stopping at a marker, a revisit or the end
/// of the table.
fn chain_of(table: &[u32], start: u32) -> Vec<u32> {
    let mut chain = Vec::new();
    let mut sector = start;
    while let Ok(index) = usize::try_from(sector) {
        if index >= table.len() || chain.contains(&sector) {
            break;
        }
        chain.push(sector);
        sector = table[index];
    }
    chain
}

/// Applies one to three seeded faults to the in-memory tables and directory
/// of `file`: links joined into another chain, looped back, cut short,
/// freed, set to a marker or pointed outside the table, in the FAT and the
/// MiniFAT; streams sharing a start, moved, resized across sector and
/// cutoff boundaries, or reclassified; the root mini stream moved or
/// resized.
fn inject_faults(file: &mut OleFile<Cursor<Vec<u8>>>, seed: u64) {
    let mut rng = Rng::new(seed);
    let streams = stream_entries(file);
    for _ in 0..1 + rng.below(3) {
        let pick = streams[rng.below(streams.len())];
        let (sid, is_mini, start, size) = pick;
        let other = streams[rng.below(streams.len())];
        match rng.below(6) {
            // A FAT link: in a regular chain, the root chain or anywhere.
            0 | 1 => {
                let root_start = file.root.as_ref().unwrap().start_sector;
                let chain = match rng.below(3) {
                    0 if !is_mini => chain_of(&file.fat, start),
                    1 => chain_of(&file.fat, root_start),
                    _ => vec![u32_of(rng.below(file.fat.len()))],
                };
                let Some(&slot) = chain.get(rng.below(chain.len().max(1))) else {
                    continue;
                };
                let target = fault_target(&mut rng, &file.fat, &chain);
                file.fat[usize::try_from(slot).unwrap()] = target;
            },
            // A MiniFAT link.
            2 if !file.minifat.is_empty() => {
                let chain = if is_mini {
                    chain_of(&file.minifat, start)
                } else {
                    vec![u32_of(rng.below(file.minifat.len()))]
                };
                let Some(&slot) = chain.get(rng.below(chain.len().max(1))) else {
                    continue;
                };
                let target = fault_target(&mut rng, &file.minifat, &chain);
                file.minifat[usize::try_from(slot).unwrap()] = target;
            },
            // A directory start sector, or a second entry for another
            // stream's whole allocation.
            3 => {
                let entry = file.dir_entries[sid].as_mut().unwrap();
                match rng.below(6) {
                    0 | 1 => {
                        entry.start_sector = other.2;
                        entry.size = other.3;
                        entry.is_minifat = other.1;
                    },
                    2 => entry.start_sector = other.2,
                    3 => entry.start_sector = ENDOFCHAIN,
                    4 => entry.start_sector = FREESECT,
                    _ => entry.start_sector = u32::try_from(rng.below(4096)).unwrap(),
                }
            },
            // A declared size, reclassified as the directory parser would.
            4 => {
                let entry = file.dir_entries[sid].as_mut().unwrap();
                let unit = if is_mini { 64 } else { file.sector_size as u64 };
                entry.size = match rng.below(6) {
                    0 => size + 1,
                    1 => size.saturating_sub(1),
                    2 => size + unit,
                    3 => size.saturating_sub(unit),
                    4 => 0,
                    _ => [4095, 4096][rng.below(2)],
                };
                entry.is_minifat = entry.size < 4096;
            },
            // The root mini stream.
            _ => {
                let root = file.root.as_mut().unwrap();
                match rng.below(3) {
                    0 => root.size += [1, 64, file.sector_size as u64][rng.below(3)],
                    1 => root.size = root.size.saturating_sub(64),
                    _ => root.start_sector = other.2,
                }
            },
        }
    }
}

/// A replacement link: usually into `own` or another chain of `table`,
/// otherwise a terminator, a marker or an index outside the table.
fn fault_target(rng: &mut Rng, table: &[u32], own: &[u32]) -> u32 {
    let len = u32::try_from(table.len()).unwrap();
    match rng.below(9) {
        0 | 1 if !own.is_empty() => own[rng.below(own.len())],
        2 | 3 => rng.u32_below(len.max(1)),
        4 => ENDOFCHAIN,
        5 => FREESECT,
        6 => [FATSECT, DIFSECT, MAXREGSECT][rng.below(3)],
        _ => len + rng.u32_below(4),
    }
}

fn u32_of(index: usize) -> u32 {
    u32::try_from(index).unwrap()
}

#[test]
fn stream_allocation_validation_matches_the_fresh_map_form_on_faulted_files() {
    let mut rng = Rng::new(0x0767_0003);
    let mut outcomes = std::collections::BTreeMap::<String, usize>::new();
    for (file_index, sector_size) in [SECTOR_SIZE_V3, SECTOR_SIZE_V4]
        .repeat(6)
        .into_iter()
        .enumerate()
    {
        let count = 20 + rng.below(60);
        let sizes = mixed_sizes(&mut rng, count);
        let bytes = build_file(sector_size, &sizes);
        for case in 0..150u64 {
            let seed = (u64::try_from(file_index).unwrap() << 32) | case;
            let mut reused = OleFile::open(Cursor::new(bytes.clone())).unwrap();
            let mut fresh = OleFile::open(Cursor::new(bytes.clone())).unwrap();
            for file in [&mut reused, &mut fresh] {
                before_allocation_validation(file);
                inject_faults(file, seed);
            }
            let actual = outcome(reused.validate_stream_allocations());
            let oracle = outcome(fresh_map_allocation_validation(&mut fresh));
            assert_eq!(actual, oracle, "file {file_index} case {case}");
            // Claimed roles and the recorded root chain agree too, including
            // the prefix a failed validation leaves behind.
            assert_eq!(
                reused.sector_roles, fresh.sector_roles,
                "file {file_index} case {case}"
            );
            assert_eq!(
                reused.root_chain, fresh.root_chain,
                "file {file_index} case {case}"
            );
            // The message with its numbers masked, so cases group by check.
            let key = match &oracle {
                Ok(()) => "accepted".to_string(),
                Err(message) => message
                    .chars()
                    .map(|character| {
                        if character.is_ascii_digit() {
                            '#'
                        } else {
                            character
                        }
                    })
                    .collect(),
            };
            *outcomes.entry(key).or_default() += 1;
        }
    }
    // The faults reach the cycle, overlap, truncation, excess, marker,
    // bounds and ownership refusals, and some leave the file valid.
    let seen = |needle: &str| outcomes.keys().any(|key| key.contains(needle));
    for needle in [
        "accepted",
        "Cycle detected in regular stream chain",
        "Cycle detected in mini stream chain",
        "Cycle detected in root mini stream chain",
        "is claimed by both regular stream and regular stream",
        "Mini sector # is claimed by multiple streams",
        "ends before its declared length",
        "exceeds its declared length",
        "Invalid sector index",
        "Invalid sector marker",
        "Invalid start marker",
        "outside the root mini stream",
        "chain must start with ENDOFCHAIN",
    ] {
        assert!(seen(needle), "no case reached {needle:?}: {outcomes:?}");
    }
}

/// A copy of `file`'s parsed state with fresh read state: no loaded mini
/// stream and empty chain buffers, as a newly opened reader has.
fn fresh_reader(file: &OleFile<Cursor<Vec<u8>>>) -> OleFile<Cursor<Vec<u8>>> {
    OleFile {
        reader: file.reader.clone(),
        file_size: file.file_size,
        sector_size: file.sector_size,
        mini_sector_size: file.mini_sector_size,
        mini_stream_cutoff: file.mini_stream_cutoff,
        fat: file.fat.clone(),
        minifat: file.minifat.clone(),
        root_chain: file.root_chain.clone(),
        first_dir_sector: file.first_dir_sector,
        root: file.root.clone(),
        dir_entries: file.dir_entries.clone(),
        dir_name_data: file.dir_name_data.clone(),
        ministream: None,
        sector_roles: file.sector_roles.clone(),
        stream_chain: EndChainScratch::default(),
    }
}

/// Every stream of `file`, read in list order on the one reader.
fn read_every_stream(
    file: &mut OleFile<Cursor<Vec<u8>>>,
    paths: &[Vec<String>],
) -> Vec<Result<Vec<u8>, String>> {
    paths
        .iter()
        .map(|path| {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            outcome(file.open_stream(&refs))
        })
        .collect()
}

#[test]
fn reads_through_one_reader_match_a_fresh_reader_per_stream_on_faulted_tables() {
    let mut rng = Rng::new(0x0767_0004);
    let (mut read, mut refused) = (0usize, 0usize);
    for (file_index, sector_size) in [SECTOR_SIZE_V3, SECTOR_SIZE_V4]
        .repeat(4)
        .into_iter()
        .enumerate()
    {
        let count = 20 + rng.below(40);
        let sizes = mixed_sizes(&mut rng, count);
        let bytes = build_file(sector_size, &sizes);
        let paths = OleFile::open(Cursor::new(bytes.clone()))
            .unwrap()
            .list_streams();
        let opened = OleFile::open(Cursor::new(bytes)).unwrap();
        for case in 0..60u64 {
            let seed = (u64::try_from(file_index).unwrap() << 32) | case;
            // The faults are applied after the validating open, so the reads
            // meet chains the open never saw: the reader's own checks decide.
            let mut faulted = fresh_reader(&opened);
            inject_faults(&mut faulted, seed);
            let mut reused = fresh_reader(&faulted);
            let actual = read_every_stream(&mut reused, &paths);
            let oracle: Vec<_> = paths
                .iter()
                .map(|path| {
                    read_every_stream(&mut fresh_reader(&faulted), std::slice::from_ref(path))
                        .remove(0)
                })
                .collect();
            assert_eq!(actual, oracle, "file {file_index} case {case}");
            assert!(all_clear(&reused.stream_chain.visited));
            for result in &actual {
                match result {
                    Ok(_) => read += 1,
                    Err(_) => refused += 1,
                }
            }
        }
    }
    assert!(read > 1_000 && refused > 100, "{read} / {refused}");
}

/// The words a validating open's chain maps write, and the file's tables.
fn open_work(bytes: &[u8]) -> (u64, OleFile<Cursor<Vec<u8>>>) {
    visited_map_work::take();
    let file = OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    (visited_map_work::take(), file)
}

fn table_words(table: &[u32]) -> u64 {
    u64::try_from(table.len().div_ceil(BITSET_WORD_BITS)).unwrap()
}

/// The sectors of every stream's chain: what a walk must at least visit.
fn chain_sectors(file: &OleFile<Cursor<Vec<u8>>>) -> u64 {
    stream_entries(file)
        .iter()
        .map(|&(_, is_mini, _, size)| {
            let unit = if is_mini { 64 } else { file.sector_size as u64 };
            size.div_ceil(unit)
        })
        .sum()
}

#[test]
fn validating_open_clears_chain_maps_in_proportion_to_the_chains() {
    for sector_size in [SECTOR_SIZE_V3, SECTOR_SIZE_V4] {
        for stream_size in [2_000usize, 5_000] {
            let mut per_stream = Vec::new();
            for count in [64usize, 256, 1_024] {
                let bytes = build_file(sector_size, &vec![stream_size; count]);
                let (work, file) = open_work(&bytes);
                let tables = table_words(&file.fat) + table_words(&file.minifat);
                let chains = chain_sectors(&file);
                // Growth covers each table once; each stream then clears at
                // most its own chain.
                assert!(
                    work <= tables + chains,
                    "{sector_size}/{stream_size}/{count}: {work} words > {tables} + {chains}"
                );
                // Clearing a table-sized map per stream, as before, costs
                // the stream count times the table.
                let table = if stream_size < 4096 {
                    &file.minifat
                } else {
                    &file.fat
                };
                let per_stream_table = u64::try_from(count).unwrap() * table_words(table);
                if count == 1_024 {
                    assert!(
                        work * 4 < per_stream_table,
                        "{sector_size}/{stream_size}/{count}: {work} vs {per_stream_table}"
                    );
                }
                per_stream.push(work as f64 / count as f64);
            }
            // Flat: the work per stream at 1,024 streams is within 10% of
            // the work per stream at 64.
            assert!(
                per_stream[2] <= per_stream[0] * 1.1,
                "{sector_size}/{stream_size}: {per_stream:?}"
            );
        }
    }
}

#[test]
fn reading_every_stream_collects_chains_in_proportion_to_the_chains() {
    for sector_size in [SECTOR_SIZE_V3, SECTOR_SIZE_V4] {
        for stream_size in [2_000usize, 5_000] {
            let mut per_stream = Vec::new();
            for count in [64usize, 256, 1_024] {
                let bytes = build_file(sector_size, &vec![stream_size; count]);
                let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
                let paths = file.list_streams();
                visited_map_work::take();
                let contents = read_every_stream(&mut file, &paths);
                let work = visited_map_work::take();
                assert!(contents.iter().all(Result::is_ok));
                // The root mini stream's chain is read once, with the first
                // mini stream.
                let root_sectors = file
                    .root
                    .as_ref()
                    .unwrap()
                    .size
                    .div_ceil(file.sector_size as u64);
                let tables = table_words(&file.fat) + table_words(&file.minifat);
                let chains = chain_sectors(&file) + root_sectors;
                assert!(
                    work <= tables + chains,
                    "{sector_size}/{stream_size}/{count}: {work} words > {tables} + {chains}"
                );
                per_stream.push(work as f64 / count as f64);
            }
            assert!(
                per_stream[2] <= per_stream[0] * 1.1,
                "{sector_size}/{stream_size}: {per_stream:?}"
            );
        }
    }
}

#[test]
fn a_reader_reuses_its_chain_buffers_across_reads_and_errors() {
    let sizes: Vec<usize> = (0..40)
        .map(|index| if index % 2 == 0 { 700 } else { 9_000 })
        .collect();
    let bytes = build_file(SECTOR_SIZE_V3, &sizes);
    let mut file = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let paths = file.list_streams();
    let expected = read_every_stream(&mut OleFile::open(Cursor::new(bytes)).unwrap(), &paths);
    assert!(expected.iter().all(Result::is_ok));

    // Read every stream once, so both tables are covered and the longest
    // chain is held.
    assert_eq!(read_every_stream(&mut file, &paths), expected);
    let words = file.stream_chain.visited.words.as_ptr();
    let word_count = file.stream_chain.visited.words.len();
    let capacity = file.stream_chain.sectors.capacity();

    // A looped regular chain fails with the reader's own error, and the
    // reader then reads every stream as before, in the same buffers.
    let regular = stream_entries(&file)
        .into_iter()
        .find(|&(_, is_mini, _, _)| !is_mini)
        .unwrap();
    let saved = file.fat.clone();
    let chain = chain_of(&file.fat, regular.2);
    file.fat[usize::try_from(chain[chain.len() - 1]).unwrap()] = chain[0];
    // Stream names are unique across the storages.
    let name = file.dir_entries[regular.0].as_ref().unwrap().name.clone();
    let path = paths
        .iter()
        .find(|path| path.last() == Some(&name))
        .unwrap()
        .clone();
    let refs: Vec<&str> = path.iter().map(String::as_str).collect();
    assert_eq!(
        outcome(file.open_stream(&refs)),
        Err(format!(
            "Corrupted file: Cycle detected in FAT chain at sector {}",
            chain[0]
        ))
    );
    assert!(all_clear(&file.stream_chain.visited));
    file.fat = saved;
    assert_eq!(read_every_stream(&mut file, &paths), expected);
    assert_eq!(file.stream_chain.visited.words.as_ptr(), words);
    assert_eq!(file.stream_chain.visited.words.len(), word_count);
    assert_eq!(file.stream_chain.sectors.capacity(), capacity);
}

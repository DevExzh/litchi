//! The in-place Reuse-plan validation held to the readback validation it
//! replaced (change 0749).
//!
//! Every test builds a real plan, corrupts it through a test-only fault
//! injection, and requires [`ReusePlan::validate`] and the retained readback
//! oracle [`ReusePlan::validate_by_readback`] to reach the same verdict. The
//! corruptions that would publish a wrong or malformed artifact must be
//! refused by both; the ones that leave every stream's bytes intact must be
//! accepted by both, so the in-place comparison neither misses a defect nor
//! refuses more than the readback did.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use super::super::OleWriter;
use super::super::core::plan_validation_declines;
use super::*;
use std::io::Cursor;
use std::path::{Path, PathBuf};

type Model = Vec<(Vec<String>, Vec<u8>)>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Verdict {
    Accepted,
    Declined,
}

fn build(streams: &[(&[&str], Vec<u8>)]) -> Vec<u8> {
    build_with(512, streams)
}

fn build_with(sector_size: usize, streams: &[(&[&str], Vec<u8>)]) -> Vec<u8> {
    let mut writer = OleWriter::with_sector_size(sector_size).unwrap();
    for (path, data) in streams {
        writer.create_stream(path, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read_model(bytes: &[u8]) -> Model {
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&refs).unwrap();
            (path, data)
        })
        .collect()
}

fn inputs(model: &Model) -> Vec<StreamInput<'_>> {
    model
        .iter()
        .map(|(path, bytes)| StreamInput { path, bytes })
        .collect()
}

fn plan(layout: &SourceLayout, streams: &[StreamInput<'_>]) -> ReusePlan {
    let storages: BTreeSet<Vec<String>> = layout.storage_paths.keys().cloned().collect();
    let class_ids = BTreeMap::new();
    let model = ModelInputs {
        sector_size: layout.sector_size,
        mini_sector_size: layout.mini_sector_size,
        mini_stream_cutoff: layout.mini_stream_cutoff,
        streams,
        storages: &storages,
        storage_class_ids: &class_ids,
        root_class_id: None,
    };
    match plan_reuse(layout, &model).unwrap() {
        Outcome::Planned(plan) => *plan,
        Outcome::Declined(reason) => panic!("fixture plan declined: {reason:?}"),
    }
}

fn verdict(result: Result<(), OleError>, which: &str, case: &str) -> Verdict {
    match result {
        Ok(()) => Verdict::Accepted,
        Err(error) => {
            assert!(
                plan_validation_declines(&error),
                "{case}: {which} validation raised a non-declining error: {error}"
            );
            Verdict::Declined
        },
    }
}

/// Both validators' verdicts on `plan`, which must agree.
fn agreed(plan: &ReusePlan, streams: &[StreamInput<'_>], case: &str) -> Verdict {
    let in_place = verdict(plan.validate(streams), "in-place", case);
    let readback = verdict(plan.validate_by_readback(streams), "readback", case);
    assert_eq!(
        in_place, readback,
        "{case}: in-place and readback validation disagree"
    );
    in_place
}

fn assert_verdict(plan: &ReusePlan, streams: &[StreamInput<'_>], expected: Verdict, case: &str) {
    assert_eq!(agreed(plan, streams, case), expected, "{case}");
}

// --- test-only fault injection over a planned artifact ---

fn fat_entry(plan: &ReusePlan, sector: u32) -> u32 {
    let offset = sector as usize * 4;
    read_u32(&plan.fat_image, offset).unwrap()
}

fn set_fat_entry(plan: &mut ReusePlan, sector: u32, value: u32) {
    write_u32(&mut plan.fat_image, sector as usize * 4, value).unwrap();
}

fn minifat_entry(plan: &ReusePlan, mini: u32) -> u32 {
    read_u32(&plan.minifat_image, mini as usize * 4).unwrap()
}

fn set_minifat_entry(plan: &mut ReusePlan, mini: u32, value: u32) {
    write_u32(&mut plan.minifat_image, mini as usize * 4, value).unwrap();
}

/// The physical sectors planned for model stream `index`, in chunk order.
fn stream_sectors(plan: &ReusePlan, index: u32) -> Vec<u32> {
    let mut found: Vec<(u32, u32)> = plan
        .sectors
        .iter()
        .enumerate()
        .filter_map(|(sector, planned)| match planned {
            PlannedSector::Stream {
                index: owner,
                chunk,
            } if *owner == index => Some((*chunk, u32::try_from(sector).unwrap())),
            _ => None,
        })
        .collect();
    found.sort_unstable();
    found.into_iter().map(|(_, sector)| sector).collect()
}

fn entry_start(plan: &ReusePlan, sid: u32) -> u32 {
    read_u32(
        &plan.directory_image,
        sid as usize * DIRENTRY_SIZE + ENTRY_START_SECTOR_OFFSET,
    )
    .unwrap()
}

fn entry_size(plan: &ReusePlan, sid: u32) -> u64 {
    let offset = sid as usize * DIRENTRY_SIZE + ENTRY_STREAM_SIZE_OFFSET;
    let mut raw = [0u8; 8];
    raw.copy_from_slice(&plan.directory_image[offset..offset + 8]);
    u64::from_le_bytes(raw)
}

fn set_entry_size(plan: &mut ReusePlan, sid: u32, size: u64) {
    let start = entry_start(plan, sid);
    patch_entry(&mut plan.directory_image, sid, start, size).unwrap();
}

fn swap_sectors(plan: &mut ReusePlan, left: u32, right: u32) {
    plan.sectors.swap(left as usize, right as usize);
}

fn sid_of(layout: &SourceLayout, path: &[&str]) -> u32 {
    let owned: Vec<String> = path.iter().map(|part| (*part).to_string()).collect();
    *layout
        .stream_paths
        .get(&owned)
        .expect("stream path is in the source")
}

fn index_of(model: &Model, path: &[&str]) -> u32 {
    let position = model
        .iter()
        .position(|(candidate, _)| {
            candidate
                .iter()
                .map(String::as_str)
                .eq(path.iter().copied())
        })
        .expect("stream path is in the model");
    u32::try_from(position).unwrap()
}

fn pattern(len: usize, seed: u8) -> Vec<u8> {
    (0..len)
        .map(|index| (index % 251) as u8 ^ seed.wrapping_mul(37))
        .collect()
}

const BIG: &[&str] = &["Big"];
const OTHER: &[&str] = &["Other"];
const SMALL: &[&str] = &["Small"];
const TINY: &[&str] = &["Tiny"];
const NESTED: &[&str] = &["Storage", "Nested"];
const REPEAT: &[&str] = &["Repeat"];

/// A source with regular, mini, empty, nested and repetitive streams, and a
/// length-changing model over it: growth that appends, a shrink that
/// reclaims, mini growth and a mini-to-regular migration.
fn synthetic() -> (Vec<u8>, Model) {
    synthetic_with(512)
}

fn synthetic_with(sector_size: usize) -> (Vec<u8>, Model) {
    let source = build_with(
        sector_size,
        &[
            (BIG, pattern(10_000, 1)),
            (OTHER, pattern(6_000, 2)),
            (SMALL, pattern(700, 3)),
            (TINY, pattern(100, 4)),
            (NESTED, pattern(300, 5)),
            (&["Empty"], Vec::new()),
            (REPEAT, vec![0xAB; 4_096]),
        ],
    );
    let mut model = read_model(&source);
    for (path, bytes) in &mut model {
        match path.join("/").as_str() {
            "Big" => *bytes = pattern(12_345, 11),
            "Other" => bytes.truncate(4_500),
            "Small" => *bytes = pattern(1_000, 13),
            "Tiny" => *bytes = pattern(5_000, 14),
            _ => {},
        }
    }
    (source, model)
}

#[test]
fn the_synthetic_plan_reuses_the_source_and_both_validators_accept_it() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let plan = plan(&layout, &streams);
    let report = plan.report();
    assert!(report.reused_source_layout());
    assert!(report.appended_sectors() > 0, "{report:?}");
    assert_verdict(&plan, &streams, Verdict::Accepted, "unmodified plan");
}

#[test]
fn overlapping_and_cyclic_chains_are_refused() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    let big = stream_sectors(&base, index_of(&model, BIG));
    let other = stream_sectors(&base, index_of(&model, OTHER));
    assert!(big.len() > 2 && other.len() > 2);

    // One chain runs on into another stream's allocation.
    let mut overlap = base.clone();
    set_fat_entry(&mut overlap, *big.last().unwrap(), other[0]);
    assert_verdict(
        &overlap,
        &streams,
        Verdict::Declined,
        "chain continues into another",
    );

    // One chain jumps into the middle of another.
    let mut shared = base.clone();
    set_fat_entry(&mut shared, big[0], other[1]);
    assert_verdict(&shared, &streams, Verdict::Declined, "chain joins another");

    // Two directory entries start on the same chain.
    let mut aliased = base.clone();
    let other_sid = sid_of(&layout, OTHER);
    let other_size = entry_size(&aliased, other_sid);
    patch_entry(&mut aliased.directory_image, other_sid, big[0], other_size).unwrap();
    assert_verdict(
        &aliased,
        &streams,
        Verdict::Declined,
        "two entries share a chain",
    );

    let mut self_cycle = base.clone();
    set_fat_entry(&mut self_cycle, big[1], big[1]);
    assert_verdict(&self_cycle, &streams, Verdict::Declined, "self cycle");

    let mut back_cycle = base.clone();
    set_fat_entry(&mut back_cycle, *big.last().unwrap(), big[0]);
    assert_verdict(
        &back_cycle,
        &streams,
        Verdict::Declined,
        "cycle to chain start",
    );

    let mut early_end = base.clone();
    set_fat_entry(&mut early_end, big[1], ENDOFCHAIN);
    assert_verdict(&early_end, &streams, Verdict::Declined, "chain ends early");

    let mut freed_link = base.clone();
    set_fat_entry(&mut freed_link, big[1], FREESECT);
    assert_verdict(
        &freed_link,
        &streams,
        Verdict::Declined,
        "chain links a free marker",
    );
    assert_eq!(fat_entry(&base, big[1]), big[2]);
}

#[test]
fn wrong_lengths_are_refused() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    for path in [BIG, OTHER, SMALL, TINY, NESTED] {
        let sid = sid_of(&layout, path);
        let size = entry_size(&base, sid);
        assert!(size > 1);
        for wrong in [size - 1, size + 1] {
            let mut plan = base.clone();
            set_entry_size(&mut plan, sid, wrong);
            assert_verdict(
                &plan,
                &streams,
                Verdict::Declined,
                &format!("{path:?} declared {wrong} instead of {size}"),
            );
        }
    }

    // The model, rather than the plan, disagreeing about a length.
    for path in [BIG, SMALL] {
        let index = index_of(&model, path) as usize;
        for delta in [-1isize, 1] {
            let mut changed = model.clone();
            let bytes = &mut changed[index].1;
            if delta < 0 {
                bytes.pop();
            } else {
                bytes.push(0);
            }
            let changed_streams = inputs(&changed);
            assert_verdict(
                &base,
                &changed_streams,
                Verdict::Declined,
                &format!("{path:?} model length {delta:+}"),
            );
        }
    }
}

#[test]
fn wrong_contents_are_refused() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    let big = stream_sectors(&base, index_of(&model, BIG));
    let other = stream_sectors(&base, index_of(&model, OTHER));

    let mut swapped = base.clone();
    swap_sectors(&mut swapped, big[0], other[0]);
    assert_verdict(
        &swapped,
        &streams,
        Verdict::Declined,
        "chunks of two streams swapped",
    );

    let mut reordered = base.clone();
    swap_sectors(&mut reordered, big[0], big[1]);
    assert_verdict(
        &reordered,
        &streams,
        Verdict::Declined,
        "chunks of one stream reordered",
    );

    let mut shifted = base.clone();
    shifted.sectors[big[2] as usize] = PlannedSector::Stream {
        index: index_of(&model, BIG),
        chunk: 3,
    };
    assert_verdict(&shifted, &streams, Verdict::Declined, "chunk index shifted");

    let mut freed = base.clone();
    freed.sectors[big[2] as usize] = PlannedSector::Free;
    assert_verdict(
        &freed,
        &streams,
        Verdict::Declined,
        "payload sector left unallocated",
    );

    let mut missing = base.clone();
    missing.sectors[big[0] as usize] = PlannedSector::Stream {
        index: 99,
        chunk: 0,
    };
    assert_verdict(
        &missing,
        &streams,
        Verdict::Declined,
        "missing model stream",
    );

    // The model changed after planning, one byte at the start, the middle
    // and the end. A mini stream's bytes are copied into the planned mini
    // stream image, so the change is a mismatch. A regular stream's sectors
    // are planned as chunks of whatever payload the model supplies, so both
    // the readback and the emission see the changed bytes consistently and
    // there is nothing to refuse.
    for (path, expected) in [
        (BIG, Verdict::Accepted),
        (OTHER, Verdict::Accepted),
        (TINY, Verdict::Accepted),
        (SMALL, Verdict::Declined),
        (NESTED, Verdict::Declined),
    ] {
        let index = index_of(&model, path) as usize;
        let len = model[index].1.len();
        for position in [0, len / 2, len - 1] {
            let mut changed = model.clone();
            changed[index].1[position] ^= 0x5A;
            let changed_streams = inputs(&changed);
            assert_verdict(
                &base,
                &changed_streams,
                expected,
                &format!("{path:?} model byte {position} changed"),
            );
        }
    }
}

#[test]
fn mini_stream_faults_are_refused_and_padding_is_not_compared() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    let small_sid = sid_of(&layout, SMALL);
    let nested_sid = sid_of(&layout, NESTED);
    let small_start = entry_start(&base, small_sid);
    let nested_start = entry_start(&base, nested_sid);
    let mini = layout.mini_sector_size;

    // A payload byte of a mini stream, in the planned mini-stream image.
    let mut payload = base.clone();
    payload.ministream_image[small_start as usize * mini + 5] ^= 0x01;
    assert_verdict(&payload, &streams, Verdict::Declined, "mini payload byte");

    // The padding after a mini stream's last byte is not stream content.
    let small_len = model[index_of(&model, SMALL) as usize].1.len();
    let mut last = small_start;
    for _ in 1..small_len.div_ceil(mini) {
        last = minifat_entry(&base, last);
    }
    let padding_offset = last as usize * mini + small_len % mini;
    assert_ne!(small_len % mini, 0, "the fixture leaves padding");
    let mut padding = base.clone();
    padding.ministream_image[padding_offset] ^= 0xFF;
    assert_verdict(&padding, &streams, Verdict::Accepted, "mini padding byte");

    // One mini chain runs on into another's.
    let mut overlap = base.clone();
    set_minifat_entry(&mut overlap, last, nested_start);
    assert_verdict(&overlap, &streams, Verdict::Declined, "mini chains overlap");

    let mut cycle = base.clone();
    set_minifat_entry(&mut cycle, small_start, small_start);
    assert_verdict(&cycle, &streams, Verdict::Declined, "mini self cycle");

    // The root mini stream loses its last sector's bytes. The readback
    // loads the whole root chain, so the in-place comparison must refuse
    // this too even though no stream's payload lies there.
    let mut truncated = base.clone();
    let keep = truncated.ministream_image.len() - truncated.sector_size;
    truncated.ministream_image.truncate(keep);
    assert_verdict(
        &truncated,
        &streams,
        Verdict::Declined,
        "mini stream image truncated",
    );
}

#[test]
fn header_and_table_faults_are_refused() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);

    let mut directory = base.clone();
    let first = read_u32(&directory.header_sector, FIRST_DIR_SECTOR_OFFSET).unwrap();
    write_u32(
        &mut directory.header_sector,
        FIRST_DIR_SECTOR_OFFSET,
        first + 1,
    )
    .unwrap();
    assert_verdict(
        &directory,
        &streams,
        Verdict::Declined,
        "directory start moved",
    );

    let mut minifat = base.clone();
    write_u32(&mut minifat.header_sector, NUM_MINIFAT_SECTORS_OFFSET, 0).unwrap();
    assert_verdict(
        &minifat,
        &streams,
        Verdict::Declined,
        "MiniFAT count cleared",
    );

    let mut fat_marker = base.clone();
    let fat_sector = fat_marker
        .sectors
        .iter()
        .position(|planned| matches!(planned, PlannedSector::Fat(_)))
        .unwrap();
    set_fat_entry(
        &mut fat_marker,
        u32::try_from(fat_sector).unwrap(),
        FREESECT,
    );
    assert_verdict(
        &fat_marker,
        &streams,
        Verdict::Declined,
        "FAT sector not marked",
    );

    let mut truncated_fat = base.clone();
    truncated_fat.fat_image.truncate(8);
    assert_verdict(
        &truncated_fat,
        &streams,
        Verdict::Declined,
        "FAT image truncated",
    );
}

#[test]
fn byte_identical_placements_are_accepted_by_both() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);

    // Two chunks of a stream whose sectors hold identical bytes: the reader
    // sees the same bytes either way, so both validators accept.
    let repeat = stream_sectors(&base, index_of(&model, REPEAT));
    let mut swapped = base.clone();
    swap_sectors(&mut swapped, repeat[0], repeat[1]);
    assert_verdict(
        &swapped,
        &streams,
        Verdict::Accepted,
        "identical chunks swapped",
    );

    // Two paths that share one payload allocation, as `create_stream_shared`
    // allows: a chunk planned from the other path is the same memory.
    let shared = pattern(5_000, 21);
    let paths = [vec!["Twin1".to_string()], vec!["Twin2".to_string()]];
    let twin_source = build(&[
        (&["Twin1"], pattern(4_600, 1)),
        (&["Twin2"], pattern(4_600, 2)),
    ]);
    let twin_layout = SourceLayout::parse(&twin_source).unwrap();
    let twin_streams = [
        StreamInput {
            path: &paths[0],
            bytes: &shared,
        },
        StreamInput {
            path: &paths[1],
            bytes: &shared,
        },
    ];
    let twin_plan = plan(&twin_layout, &twin_streams);
    assert_verdict(
        &twin_plan,
        &twin_streams,
        Verdict::Accepted,
        "shared payload",
    );
    let first = stream_sectors(&twin_plan, 0);
    let second = stream_sectors(&twin_plan, 1);
    let mut crossed = twin_plan.clone();
    swap_sectors(&mut crossed, first[0], second[0]);
    assert_verdict(
        &crossed,
        &twin_streams,
        Verdict::Accepted,
        "shared payload crossed",
    );
    let mut shifted = twin_plan.clone();
    swap_sectors(&mut shifted, first[0], first[1]);
    assert_verdict(
        &shifted,
        &twin_streams,
        Verdict::Declined,
        "shared payload reordered",
    );
}

/// A deterministic xorshift generator; the sweep must be reproducible.
struct Sweep(u64);

impl Sweep {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound.max(1)).unwrap()).unwrap()
    }
}

/// Applies one random fault and returns its description.
fn random_fault(plan: &mut ReusePlan, sweep: &mut Sweep, stream_count: usize) -> String {
    let flip = |image: &mut Vec<u8>, sweep: &mut Sweep, name: &str| {
        if image.is_empty() {
            return format!("{name} empty");
        }
        let position = sweep.below(image.len());
        let bit = 1u8 << sweep.below(8);
        image[position] ^= bit;
        format!("{name} byte {position} bit {bit}")
    };
    match sweep.below(9) {
        0 => flip(&mut plan.header_sector, sweep, "header"),
        1 => flip(&mut plan.fat_image, sweep, "FAT"),
        2 => flip(&mut plan.minifat_image, sweep, "MiniFAT"),
        3 => flip(&mut plan.directory_image, sweep, "directory"),
        4 => flip(&mut plan.ministream_image, sweep, "mini stream"),
        5 => {
            let left = sweep.below(plan.sectors.len());
            let right = sweep.below(plan.sectors.len());
            plan.sectors.swap(left, right);
            format!("sectors {left} and {right} swapped")
        },
        6 => {
            let sector = sweep.below(plan.sectors.len());
            plan.sectors[sector] = PlannedSector::Free;
            format!("sector {sector} freed")
        },
        7 => {
            let sector = sweep.below(plan.sectors.len());
            let index = u32::try_from(sweep.below(stream_count + 1)).unwrap();
            let chunk = u32::try_from(sweep.below(4)).unwrap();
            plan.sectors[sector] = PlannedSector::Stream { index, chunk };
            format!("sector {sector} planned as stream {index} chunk {chunk}")
        },
        _ => {
            let entry = sweep.below(plan.fat_image.len() / 4);
            let value = match sweep.below(4) {
                0 => ENDOFCHAIN,
                1 => FREESECT,
                2 => FATSECT,
                _ => u32::try_from(sweep.below(plan.sectors.len() + 2)).unwrap(),
            };
            write_u32(&mut plan.fat_image, entry * 4, value).unwrap();
            format!("FAT entry {entry} set to {value:#x}")
        },
    }
}

fn sweep_agreement(
    label: &str,
    plan: &ReusePlan,
    streams: &[StreamInput<'_>],
    rounds: usize,
    seed: u64,
) {
    let mut sweep = Sweep(seed);
    let mut accepted = 0usize;
    let mut declined = 0usize;
    for round in 0..rounds {
        let mut faulted = plan.clone();
        let fault = random_fault(&mut faulted, &mut sweep, streams.len());
        match agreed(
            &faulted,
            streams,
            &format!("{label} round {round}: {fault}"),
        ) {
            Verdict::Accepted => accepted += 1,
            Verdict::Declined => declined += 1,
        }
    }
    assert!(
        declined > 0,
        "{label}: the sweep never produced a refused fault"
    );
    assert!(
        accepted > 0,
        "{label}: the sweep never produced a harmless fault"
    );
    assert_eq!(
        accepted + declined,
        rounds,
        "{label}: every round has a verdict"
    );
    eprintln!(
        "0749 sweep {label}: {rounds} faults, {accepted} accepted by both, {declined} refused by both"
    );
}

#[test]
fn random_faults_reach_the_readback_verdict() {
    let (source, model) = synthetic();
    let layout = SourceLayout::parse(&source).unwrap();
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    sweep_agreement("synthetic", &base, &streams, 2_000, 0x0749_5eed);
}

/// Version 4 geometry: 4096-byte sectors hold 64 mini sectors each and
/// 1,024 FAT entries, so runs, mini-stream offsets and table images differ.
#[test]
fn version_4_plans_reach_the_readback_verdict() {
    let (source, model) = synthetic_with(4096);
    let layout = SourceLayout::parse(&source).unwrap();
    assert_eq!(layout.sector_size, 4096);
    let streams = inputs(&model);
    let base = plan(&layout, &streams);
    assert!(base.report().reused_source_layout());
    assert_verdict(&base, &streams, Verdict::Accepted, "version 4 plan");
    let big = stream_sectors(&base, index_of(&model, BIG));
    let other = stream_sectors(&base, index_of(&model, OTHER));
    let mut swapped = base.clone();
    swap_sectors(&mut swapped, big[0], other[0]);
    assert_verdict(
        &swapped,
        &streams,
        Verdict::Declined,
        "version 4 chunks swapped",
    );
    let mut truncated = base.clone();
    let keep = truncated.ministream_image.len() - truncated.sector_size;
    truncated.ministream_image.truncate(keep);
    assert_verdict(
        &truncated,
        &streams,
        Verdict::Declined,
        "version 4 mini image truncated",
    );
    sweep_agreement("version 4", &base, &streams, 1_000, 0x0749_4096);
}

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("crate lives two levels below the repository root")
        .to_path_buf()
}

/// Real fixtures with a length-changing model: the largest stream grows by a
/// few sectors and the smallest non-empty stream grows by a few bytes.
#[test]
fn real_fixture_plans_reach_the_readback_verdict() {
    let fixtures = [
        "test-data/poi/test-data/slideshow/45543.ppt",
        "test-data/poi/test-data/slideshow/41246-1.ppt",
        "test-data/ole/doc/FloatingPictures.doc",
        "test-data/ole/doc/NoHeadFoot.doc",
    ];
    for (number, fixture) in fixtures.iter().enumerate() {
        let source = std::fs::read(repository_root().join(fixture)).unwrap();
        let mut model = read_model(&source);
        let largest = (0..model.len())
            .max_by_key(|index| model[*index].1.len())
            .unwrap();
        let smallest = (0..model.len())
            .filter(|index| !model[*index].1.is_empty())
            .min_by_key(|index| model[*index].1.len())
            .unwrap();
        model[largest].1.extend(pattern(1_500, 7));
        model[smallest].1.extend(pattern(10, 8));
        let layout = SourceLayout::parse(&source).unwrap();
        let streams = inputs(&model);
        let base = plan(&layout, &streams);
        assert!(base.report().reused_source_layout(), "{fixture}");
        assert_verdict(&base, &streams, Verdict::Accepted, fixture);
        sweep_agreement(fixture, &base, &streams, 500, 0x0749 + number as u64);
    }
}

/// Hundreds of mini streams, in both geometries: mini growth that appends,
/// a shrink that reclaims, and a mini-to-regular migration, then random
/// faults, all with identical verdicts.
#[test]
fn many_mini_stream_plans_reach_the_readback_verdict() {
    for sector_size in [512, 4096] {
        let names: Vec<String> = (0..240).map(|index| format!("M{index:04}")).collect();
        let paths: Vec<[&str; 1]> = names.iter().map(|name| [name.as_str()]).collect();
        let specs: Vec<(&[&str], Vec<u8>)> = paths
            .iter()
            .enumerate()
            .map(|(index, path)| {
                let seed = u8::try_from(index % 251).unwrap();
                (path.as_slice(), pattern(1_000 + index % 7, seed))
            })
            .collect();
        let source = build_with(sector_size, &specs);
        let mut model = read_model(&source);
        model[3].1.extend(pattern(900, 1));
        model[40].1.truncate(100);
        model[77].1 = pattern(6_000, 2);
        let layout = SourceLayout::parse(&source).unwrap();
        let streams = inputs(&model);
        let base = plan(&layout, &streams);
        assert!(base.report().reused_source_layout(), "{sector_size}");
        assert_verdict(&base, &streams, Verdict::Accepted, "many mini streams");
        let mut payload = base.clone();
        let middle = payload.ministream_image.len() / 2;
        payload.ministream_image[middle] ^= 0x10;
        // Payload or padding, the two validators must agree on it.
        agreed(&payload, &streams, "mini stream image byte");
        sweep_agreement(
            &format!("many mini streams, {sector_size}-byte sectors"),
            &base,
            &streams,
            300,
            0x0749_0240 + sector_size as u64,
        );
    }
}

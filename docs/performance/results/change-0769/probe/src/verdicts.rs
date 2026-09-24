//! The cross-build verdict lane: seeded byte-level faults on one compound
//! file, and every verdict the reader reaches on each faulted copy.
//!
//! The clean file's layout is located by walking its header, DIFAT and FAT
//! directly (the public API does not expose sector positions). Each case
//! applies one to three faults, chosen by a seeded generator:
//!
//! - a FAT link rewritten to another sector (joining or looping a chain),
//!   to itself, to `ENDOFCHAIN`, `FREESECT`, `FATSECT` or `DIFSECT`, or to an
//!   index outside the file;
//! - a MiniFAT link rewritten the same way within the MiniFAT;
//! - a stream or root directory entry given another stream's start sector,
//!   another stream's whole allocation (start and size), a size one byte,
//!   one mini sector or one sector off, a size across the 4096-byte cutoff,
//!   a zero size, or a terminator as its start.
//!
//! For each case it records `OleFile::open`'s verdict, then `open_stream` of
//! every listed stream on that reader in order, and the same for
//! `SharedOleFile::open_owned`. Every error is kept as its full display
//! string; every stream read as its length and SHA-256. The output depends
//! only on the input and the seed, so two builds' outputs compare byte for
//! byte.

use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;

use litchi_cfb::{OleFile, SharedOleFile};
use litchi_core::SourceVersion;

use crate::{BoxError, fail, json_string, sha};

const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
const FREESECT: u32 = 0xFFFF_FFFF;
const FATSECT: u32 = 0xFFFF_FFFD;
const DIFSECT: u32 = 0xFFFF_FFFC;
const MAXREGSECT: u32 = 0xFFFF_FFFA;

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
        (self.next() % bound as u64) as usize
    }
}

fn u32_at(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(bytes.get(offset..offset + 4)?.try_into().ok()?))
}

fn put_u32(bytes: &mut [u8], offset: usize, value: u32) {
    if let Some(slot) = bytes.get_mut(offset..offset + 4) {
        slot.copy_from_slice(&value.to_le_bytes());
    }
}

/// Where the clean file keeps its tables and directory.
struct Layout {
    sector_size: usize,
    sectors: u32,
    /// Byte offset of each FAT entry, by entry index.
    fat: Vec<usize>,
    /// Byte offset of each MiniFAT entry, by entry index.
    minifat: Vec<usize>,
    /// Byte offset of each directory entry, by SID.
    directory: Vec<usize>,
}

impl Layout {
    fn locate(bytes: &[u8]) -> Result<Self, BoxError> {
        let shift = u16::from_le_bytes([bytes[0x1E], bytes[0x1F]]);
        if !matches!(shift, 9 | 12) {
            return fail("unsupported sector shift");
        }
        let sector_size = 1usize << shift;
        let sectors = u32::try_from((bytes.len() / sector_size).saturating_sub(1))?;
        let offset_of = |sector: u32| (sector as usize + 1) * sector_size;
        let fat_count = u32_at(bytes, 0x2C).ok_or("header")? as usize;
        let mut fat_sectors = Vec::new();
        for index in 0..fat_count.min(109) {
            fat_sectors.push(u32_at(bytes, 0x4C + 4 * index).ok_or("header DIFAT")?);
        }
        let mut difat = u32_at(bytes, 0x44).ok_or("header")?;
        let per_difat = sector_size / 4 - 1;
        let mut guard = 0;
        while fat_sectors.len() < fat_count && difat < MAXREGSECT && guard < 1 << 20 {
            let base = offset_of(difat);
            for index in 0..per_difat {
                if fat_sectors.len() < fat_count {
                    fat_sectors.push(u32_at(bytes, base + 4 * index).ok_or("DIFAT")?);
                }
            }
            difat = u32_at(bytes, base + 4 * per_difat).ok_or("DIFAT")?;
            guard += 1;
        }
        let mut fat = Vec::new();
        for sector in fat_sectors {
            let base = offset_of(sector);
            fat.extend((0..sector_size / 4).map(|index| base + 4 * index));
        }
        let fat_value = |index: u32| -> Option<u32> {
            u32_at(bytes, *fat.get(index as usize)?)
        };
        let chain = |start: u32| -> Vec<u32> {
            let mut out = Vec::new();
            let mut sector = start;
            while sector < MAXREGSECT && out.len() <= fat.len() && !out.contains(&sector) {
                out.push(sector);
                sector = fat_value(sector).unwrap_or(ENDOFCHAIN);
            }
            out
        };
        let mut minifat = Vec::new();
        for sector in chain(u32_at(bytes, 0x3C).ok_or("header")?) {
            let base = offset_of(sector);
            minifat.extend((0..sector_size / 4).map(|index| base + 4 * index));
        }
        let mut directory = Vec::new();
        for sector in chain(u32_at(bytes, 0x30).ok_or("header")?) {
            let base = offset_of(sector);
            directory.extend((0..sector_size / 128).map(|index| base + 128 * index));
        }
        fat.retain(|&offset| offset + 4 <= bytes.len());
        minifat.retain(|&offset| offset + 4 <= bytes.len());
        directory.retain(|&offset| offset + 128 <= bytes.len());
        Ok(Self {
            sector_size,
            sectors,
            fat,
            minifat,
            directory,
        })
    }
}

/// A replacement link for an entry of a table of `len` entries.
fn link(rng: &mut Rng, len: usize, links: &[usize], own: usize) -> u32 {
    let len = u32::try_from(len).unwrap_or(u32::MAX);
    match rng.below(20) {
        0..=6 if !links.is_empty() => links[rng.below(links.len())] as u32,
        7 => own as u32,
        8..=10 => ENDOFCHAIN,
        11 | 12 => FREESECT,
        13 => [FATSECT, DIFSECT][rng.below(2)],
        14..=16 => rng.below(len as usize + 4) as u32,
        17 => len.saturating_add(rng.below(64) as u32),
        _ => rng.next() as u32,
    }
}

/// Applies one to three faults to `bytes`; returns their descriptions.
fn inject(bytes: &mut [u8], layout: &Layout, rng: &mut Rng) -> Vec<String> {
    let links = |bytes: &[u8], offsets: &[usize]| -> Vec<usize> {
        offsets
            .iter()
            .enumerate()
            .filter(|&(_, &offset)| {
                u32_at(bytes, offset).is_some_and(|value| value < MAXREGSECT || value == ENDOFCHAIN)
            })
            .map(|(index, _)| index)
            .collect()
    };
    let streams: Vec<usize> = layout
        .directory
        .iter()
        .copied()
        .filter(|&offset| matches!(bytes[offset + 0x42], 2 | 5))
        .collect();
    let mut faults = Vec::new();
    for _ in 0..1 + rng.below(3) {
        match rng.below(10) {
            0..=3 if !layout.fat.is_empty() => {
                let linked = links(bytes, &layout.fat);
                let index = if !linked.is_empty() && rng.below(4) != 0 {
                    linked[rng.below(linked.len())]
                } else {
                    rng.below(layout.fat.len())
                };
                let value = link(rng, layout.sectors as usize, &linked, index);
                put_u32(bytes, layout.fat[index], value);
                faults.push(format!("fat[{index}]={value:08X}"));
            },
            4 | 5 if !layout.minifat.is_empty() => {
                let linked = links(bytes, &layout.minifat);
                let index = if !linked.is_empty() && rng.below(4) != 0 {
                    linked[rng.below(linked.len())]
                } else {
                    rng.below(layout.minifat.len())
                };
                let value = link(rng, layout.minifat.len(), &linked, index);
                put_u32(bytes, layout.minifat[index], value);
                faults.push(format!("minifat[{index}]={value:08X}"));
            },
            _ if !streams.is_empty() => {
                let entry = streams[rng.below(streams.len())];
                let other = streams[rng.below(streams.len())];
                let start = u32_at(bytes, entry + 0x74).unwrap_or(0);
                let size = u32_at(bytes, entry + 0x78).unwrap_or(0);
                let other_start = u32_at(bytes, other + 0x74).unwrap_or(0);
                let other_size = u32_at(bytes, other + 0x78).unwrap_or(0);
                let sector = layout.sector_size as u32;
                let (new_start, new_size) = match rng.below(12) {
                    0 | 1 => (other_start, size),
                    2 | 3 => (other_start, other_size),
                    4 => (start, size.wrapping_add(1)),
                    5 => (start, size.wrapping_sub(1)),
                    6 => (start, size.wrapping_add(64)),
                    7 => (start, size.saturating_sub(64)),
                    8 => (start, size.wrapping_add(sector)),
                    9 => (start, [0, 4095, 4096][rng.below(3)]),
                    10 => ([ENDOFCHAIN, FREESECT][rng.below(2)], size),
                    _ => (rng.below(layout.sectors as usize + 2) as u32, size),
                };
                put_u32(bytes, entry + 0x74, new_start);
                put_u32(bytes, entry + 0x78, new_size);
                faults.push(format!("dir@{entry}=({new_start:08X},{new_size})"));
            },
            _ => faults.push("none".to_string()),
        }
    }
    faults
}

fn verdict<T>(result: Result<T, litchi_cfb::OleError>) -> Result<T, String> {
    result.map_err(|error| error.to_string())
}

/// Every stream of the reader read in list order, as one line per stream.
fn read_lines(ole: &mut OleFile<Cursor<&[u8]>>) -> Vec<String> {
    let mut lines = Vec::new();
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        lines.push(match verdict(ole.open_stream(&refs)) {
            Ok(data) => format!("{}:ok:{}:{}", path.join("/"), data.len(), sha(&data)),
            Err(message) => format!("{}:err:{message}", path.join("/")),
        });
    }
    lines
}

fn shared_lines(bytes: &[u8]) -> (String, Vec<String>) {
    let source: Arc<[u8]> = Arc::from(bytes.to_vec().into_boxed_slice());
    match verdict(SharedOleFile::open_owned(source, SourceVersion::new(767, 0))) {
        Ok(shared) => {
            let mut paths: Vec<String> = shared
                .directory_entries()
                .filter(|entry| entry.entry_type == 2)
                .map(|entry| entry.name.clone())
                .collect();
            paths.sort();
            let mut lines = Vec::new();
            // Root-level names only: nested streams are covered by the
            // cursor reader's lines.
            for name in paths {
                lines.push(match verdict(shared.open_stream(&[name.as_str()])) {
                    Ok(data) => format!("{name}:ok:{}:{}", data.len(), sha(&data)),
                    Err(message) => format!("{name}:err:{message}"),
                });
            }
            ("ok".to_string(), lines)
        },
        Err(message) => (format!("err:{message}"), Vec::new()),
    }
}

/// Prints one JSON line per case and a closing summary line.
pub(crate) fn run(input: &Path, cases: usize, seed: u64) -> Result<(), BoxError> {
    let clean = std::fs::read(input)?;
    let layout = Layout::locate(&clean)?;
    let mut rng = Rng::new(seed ^ u64::from_le_bytes(sha(&clean).as_bytes()[..8].try_into()?));
    let (mut opened, mut refused, mut reads_ok, mut reads_err) = (0usize, 0usize, 0usize, 0usize);
    for case in 0..=cases {
        let mut bytes = clean.clone();
        let faults = if case == 0 {
            vec!["clean".to_string()]
        } else {
            inject(&mut bytes, &layout, &mut rng)
        };
        let (open, lines) = match verdict(OleFile::open(Cursor::new(bytes.as_slice()))) {
            Ok(mut ole) => {
                opened += 1;
                ("ok".to_string(), read_lines(&mut ole))
            },
            Err(message) => {
                refused += 1;
                (format!("err:{message}"), Vec::new())
            },
        };
        reads_ok += lines.iter().filter(|line| line.contains(":ok:")).count();
        reads_err += lines.iter().filter(|line| line.contains(":err:")).count();
        // The distinct read errors, in first-seen order.
        let mut read_errors: Vec<&str> = Vec::new();
        for line in &lines {
            if let Some((_, message)) = line.split_once(":err:")
                && !read_errors.contains(&message)
            {
                read_errors.push(message);
            }
        }
        let (shared, shared_reads) = shared_lines(&bytes);
        println!(
            "{{\"case\":{case},\"faults\":{},\"open\":{},\"reads\":{},\"read_errors\":{},\"read_digest\":{},\"shared\":{},\"shared_read_digest\":{}}}",
            json_string(&faults.join(";")),
            json_string(&open),
            lines.len(),
            json_string(&read_errors.join(" | ")),
            json_string(&sha(lines.join("\n").as_bytes())),
            json_string(&shared),
            json_string(&sha(shared_reads.join("\n").as_bytes())),
        );
    }
    println!(
        "{{\"summary\":true,\"input\":{},\"input_sha256\":{},\"cases\":{cases},\"opened\":{opened},\"refused\":{refused},\"reads_ok\":{reads_ok},\"reads_err\":{reads_err}}}",
        json_string(&input.display().to_string()),
        json_string(&sha(&clean)),
    );
    Ok(())
}

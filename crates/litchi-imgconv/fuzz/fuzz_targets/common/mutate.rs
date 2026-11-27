//! Record-aware mutation for metafile targets.
//!
//! Byte-level mutation breaks the size chain of a metafile as often as it
//! changes a field, so most mutants die in the framing checks and never reach
//! playback. These mutators split an input into its records, edit whole
//! records or single fields, and write the records back with consistent sizes
//! (and, usually, a consistent header), so a mutant keeps parsing.

#![allow(dead_code)]

use std::ops::Range;

/// A small deterministic generator seeded by libFuzzer.
pub struct Rng(u64);

impl Rng {
    pub fn new(seed: u32) -> Self {
        Self((u64::from(seed) << 32) | 0x9e37_79b9)
    }

    pub fn next(&mut self) -> u64 {
        // xorshift64*
        self.0 ^= self.0 >> 12;
        self.0 ^= self.0 << 25;
        self.0 ^= self.0 >> 27;
        self.0.wrapping_mul(0x2545_f491_4f6c_dd1d)
    }

    pub fn below(&mut self, bound: usize) -> usize {
        if bound == 0 {
            0
        } else {
            (self.next() >> 11) as usize % bound
        }
    }

    pub fn chance(&mut self, numerator: usize, denominator: usize) -> bool {
        self.below(denominator) < numerator
    }
}

const INTERESTING: [u32; 20] = [
    0,
    1,
    2,
    3,
    4,
    8,
    16,
    0xff,
    0x100,
    0x7fff,
    0x8000,
    0xffff,
    0x1_0000,
    0x00ff_ffff,
    0x7fff_ffff,
    0x8000_0000,
    0xffff_8000,
    0xffff_fffe,
    0xffff_ffff,
    0x4000_0000,
];

/// How one metafile format frames its records.
pub trait Framing {
    /// Bytes per field edit: 4 for EMF and EMF+, 2 for WMF.
    const UNIT: usize;
    /// Header bytes at the start of every record that field edits avoid.
    const RECORD_HEADER: usize;

    /// Splits `data` into the bytes before the first record, the records and
    /// the bytes after the last one. `None` when `data` is not framed.
    fn split(data: &[u8]) -> Option<Split>;

    /// The bytes of a record that hold its size.
    fn size_field() -> Range<usize>;

    /// Rewrites the size field of `record` after its length changed.
    fn set_size(record: &mut [u8]);

    /// Brings the header (total size, record count) in line with `records`.
    fn fix_header(prefix: &mut [u8], records: &mut [Vec<u8>], tail_len: usize);
}

pub struct Split {
    pub prefix: Range<usize>,
    pub records: Vec<Range<usize>>,
    pub tail: Range<usize>,
}

fn u32_at(data: &[u8], offset: usize) -> Option<u32> {
    let bytes = data.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes(bytes.try_into().ok()?))
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    if let Some(slot) = data.get_mut(offset..offset + 4) {
        slot.copy_from_slice(&value.to_le_bytes());
    }
}

/// EMF: records are `Type`, `Size` (bytes, a multiple of four), payload.
pub struct Emf;

impl Framing for Emf {
    const UNIT: usize = 4;
    const RECORD_HEADER: usize = 8;

    fn split(data: &[u8]) -> Option<Split> {
        let mut records = Vec::new();
        let mut position = 0usize;
        while let (Some(_), Some(size)) = (u32_at(data, position), u32_at(data, position + 4)) {
            let size = size as usize;
            if size < 8 || size % 4 != 0 || position.checked_add(size)? > data.len() {
                break;
            }
            records.push(position..position + size);
            position += size;
        }
        if records.len() < 2 {
            return None;
        }
        Some(Split {
            prefix: 0..0,
            records,
            tail: position..data.len(),
        })
    }

    fn size_field() -> Range<usize> {
        4..8
    }

    fn set_size(record: &mut [u8]) {
        let size = record.len() as u32;
        put_u32(record, 4, size);
    }

    fn fix_header(_prefix: &mut [u8], records: &mut [Vec<u8>], tail_len: usize) {
        // `EMR_HEADER` keeps the file size at offset 48 and the record count
        // at offset 52.
        let total = records.iter().map(Vec::len).sum::<usize>() + tail_len;
        let count = records.len();
        if let Some(first) = records.first_mut()
            && u32_at(first, 0) == Some(1)
            && first.len() >= 56
        {
            put_u32(first, 48, total as u32);
            put_u32(first, 52, count as u32);
        }
    }
}

/// WMF: an optional placeable header and the 18-byte header, then records of
/// `Size` (16-bit words), `Function`, parameters.
pub struct Wmf;

impl Wmf {
    const PLACEABLE: usize = 22;
    const HEADER: usize = 18;
}

impl Framing for Wmf {
    const UNIT: usize = 2;
    const RECORD_HEADER: usize = 6;

    fn split(data: &[u8]) -> Option<Split> {
        let mut start = 0usize;
        if u32_at(data, 0) == Some(0x9ac6_cdd7) {
            start = Self::PLACEABLE;
        }
        start += Self::HEADER;
        if data.len() < start {
            return None;
        }
        let mut records = Vec::new();
        let mut position = start;
        while let Some(words) = u32_at(data, position) {
            let size = (words as usize).checked_mul(2)?;
            if size < 6 || position.checked_add(size)? > data.len() {
                break;
            }
            records.push(position..position + size);
            position += size;
        }
        if records.len() < 2 {
            return None;
        }
        Some(Split {
            prefix: 0..start,
            records,
            tail: position..data.len(),
        })
    }

    fn size_field() -> Range<usize> {
        0..4
    }

    fn set_size(record: &mut [u8]) {
        let words = (record.len() / 2) as u32;
        put_u32(record, 0, words);
    }

    fn fix_header(prefix: &mut [u8], records: &mut [Vec<u8>], tail_len: usize) {
        // The standard header holds the file size in words at offset 6 and
        // the largest record in words at offset 12.
        let header = if prefix.len() >= Self::PLACEABLE + Self::HEADER {
            Self::PLACEABLE
        } else {
            0
        };
        let body: usize = records.iter().map(Vec::len).sum::<usize>() + tail_len;
        let total_words = (Self::HEADER + body) / 2;
        let largest = records
            .iter()
            .map(|record| record.len() / 2)
            .max()
            .unwrap_or(0);
        put_u32(prefix, header + 6, total_words as u32);
        put_u32(prefix, header + 12, largest as u32);
    }
}

/// EMF+: records are `Type` and `Flags` (16 bits each), `Size` (bytes, a
/// multiple of four, header included), `DataSize`, payload.
pub struct EmfPlus;

impl Framing for EmfPlus {
    const UNIT: usize = 4;
    const RECORD_HEADER: usize = 12;

    fn split(data: &[u8]) -> Option<Split> {
        let mut records = Vec::new();
        let mut position = 0usize;
        while let (Some(_), Some(size)) = (u32_at(data, position), u32_at(data, position + 4)) {
            let size = size as usize;
            if size < 12 || size % 4 != 0 || position.checked_add(size)? > data.len() {
                break;
            }
            records.push(position..position + size);
            position += size;
        }
        if records.len() < 2 {
            return None;
        }
        Some(Split {
            prefix: 0..0,
            records,
            tail: position..data.len(),
        })
    }

    fn size_field() -> Range<usize> {
        4..8
    }

    fn set_size(record: &mut [u8]) {
        let size = record.len() as u32;
        put_u32(record, 4, size);
        put_u32(record, 8, size.saturating_sub(12));
    }

    fn fix_header(_prefix: &mut [u8], _records: &mut [Vec<u8>], _tail_len: usize) {}
}

/// Mutates `data[..size]` record by record. Returns the new length, or `None`
/// when the input is not framed (the caller then falls back to byte mutation).
pub fn mutate_records<F: Framing>(
    data: &mut [u8],
    size: usize,
    max_size: usize,
    seed: u32,
    byte_mutate: impl Fn(&mut [u8], usize, usize) -> usize,
) -> Option<usize> {
    let input = data.get(..size)?;
    let split = F::split(input)?;
    let mut rng = Rng::new(seed);
    let mut prefix = input[split.prefix.clone()].to_vec();
    let mut records: Vec<Vec<u8>> = split
        .records
        .iter()
        .map(|range| input[range.clone()].to_vec())
        .collect();
    let tail = input[split.tail.clone()].to_vec();

    for _ in 0..=rng.below(3) {
        apply_one::<F>(&mut records, &mut rng, &byte_mutate);
    }

    if rng.chance(3, 4) {
        F::fix_header(&mut prefix, &mut records, tail.len());
    }

    let total = prefix.len() + records.iter().map(Vec::len).sum::<usize>() + tail.len();
    if total > max_size || records.len() < 2 {
        return None;
    }
    let mut output = Vec::with_capacity(total);
    output.extend_from_slice(&prefix);
    for record in &records {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(&tail);
    data[..output.len()].copy_from_slice(&output);
    Some(output.len())
}

fn apply_one<F: Framing>(
    records: &mut Vec<Vec<u8>>,
    rng: &mut Rng,
    byte_mutate: &impl Fn(&mut [u8], usize, usize) -> usize,
) {
    let count = records.len();
    let index = rng.below(count);
    match rng.below(10) {
        0..=3 => set_field::<F>(&mut records[index], rng),
        4 => {
            let record = &mut records[index];
            if !record.is_empty() {
                let bit = rng.below(record.len() * 8);
                record[bit / 8] ^= 1 << (bit % 8);
            }
        },
        5 => {
            let copy = records[index].clone();
            records.insert(rng.below(count + 1), copy);
        },
        6 => {
            if count > 2 {
                records.remove(index);
            }
        },
        7 => {
            let other = rng.below(count);
            records.swap(index, other);
        },
        8 => {
            // Byte-level mutation of one record's payload, keeping its length.
            let record = &mut records[index];
            let start = F::RECORD_HEADER.min(record.len());
            let length = record.len() - start;
            if length > 0 {
                let new_length = byte_mutate(&mut record[start..], length, length);
                for byte in &mut record[start + new_length.min(length)..] {
                    *byte = 0;
                }
            }
        },
        _ => {
            // Grow or shrink the payload by one unit and fix the size field.
            let record = &mut records[index];
            if rng.chance(1, 2) || record.len() <= F::RECORD_HEADER + F::UNIT {
                for _ in 0..F::UNIT {
                    record.push(rng.below(256) as u8);
                }
            } else {
                let keep = record.len() - F::UNIT;
                record.truncate(keep);
            }
            F::set_size(record);
        },
    }
}

fn set_field<F: Framing>(record: &mut [u8], rng: &mut Rng) {
    // Mostly payload fields; sometimes the type, which retargets the record.
    let editable_from = if rng.chance(1, 8) {
        0
    } else {
        F::RECORD_HEADER
    };
    let slots = record.len().saturating_sub(editable_from) / F::UNIT;
    if slots == 0 {
        return;
    }
    let offset = editable_from + rng.below(slots) * F::UNIT;
    let value = if rng.chance(3, 4) {
        INTERESTING[rng.below(INTERESTING.len())]
    } else {
        rng.next() as u32
    };
    // The size field keeps the record framed; every other field is fair game.
    let size_field = F::size_field();
    let saved = record[size_field.clone()].to_vec();
    let bytes = value.to_le_bytes();
    record[offset..offset + F::UNIT].copy_from_slice(&bytes[..F::UNIT]);
    record[size_field].copy_from_slice(&saved);
}

/// Mutates a framed metafile record by record, and any other input (and a
/// share of framed ones, so headers still get byte-level edits) byte by byte.
pub fn mutate<F: Framing>(data: &mut [u8], size: usize, max_size: usize, seed: u32) -> usize {
    let bytes = |data: &mut [u8], size: usize, max_size: usize| {
        libfuzzer_sys::fuzzer_mutate(data, size, max_size)
    };
    if !Rng::new(seed ^ 0x5bd1_e995).chance(1, 8)
        && let Some(length) = mutate_records::<F>(data, size, max_size, seed, bytes)
    {
        return length;
    }
    libfuzzer_sys::fuzzer_mutate(data, size, max_size)
}

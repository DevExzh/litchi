#![allow(
    clippy::print_stdout,
    reason = "this measurement probe intentionally prints its JSON result"
)]

//! Scoped-family patch probe (record 0357a).
//!
//! Commits `N` independent single-operation `external_media` edits on one
//! source package (`set_path` of one movie to a distinct path each time) and
//! retains every patch, as a patch log or undo history would. It then
//! replays every patch onto the source and every inverse onto its result,
//! checking the replayed bytes outside the timed regions.
//!
//! Prints one JSON line: the patch count, the source size, the bytes the
//! retained patches hold (read from the patch's content-free `Debug` view,
//! so the same probe builds against the whole-artifact and the record-sized
//! patch), and the median commit, apply and inverse-apply times. Peak RSS is
//! measured outside (`/usr/bin/time -v`).
//!
//! Record 0357b adds `first_commit_median_us`: the same single-operation
//! commit on a freshly opened, never-edited package (one per iteration,
//! `FIRST` iterations, default 200), which pays any once-per-package digest
//! the steady-state commits amortize. The open is outside the timed region.
//!
//! Run with: `cargo run --release -p litchi-ppt --example scoped_patch_probe
//! -- test-data/unified-ops/0294/video.ppt 1000 200`

use std::time::Instant;

use litchi_ppt::{Package, external_media};

/// The bytes one patch retains, from its `Debug` fields: `payload_bytes`
/// for the record-sized patch, `before_bytes + after_bytes` for the
/// whole-artifact one.
fn retained(patch: &external_media::Patch) -> usize {
    let text = format!("{patch:?}");
    let field = |name: &str| -> Option<usize> {
        let start = text.find(&format!("{name}: "))? + name.len() + 2;
        let digits: String = text
            .get(start..)?
            .chars()
            .take_while(char::is_ascii_digit)
            .collect();
        digits.parse().ok()
    };
    field("payload_bytes")
        .or_else(|| Some(field("before_bytes")? + field("after_bytes")?))
        .unwrap_or(0)
}

fn median(samples: &mut [u128]) -> f64 {
    samples.sort_unstable();
    let middle = samples.len() / 2;
    if samples.len() % 2 == 1 {
        samples[middle] as f64 / 1000.0
    } else {
        (samples[middle - 1] + samples[middle]) as f64 / 2000.0
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path = args
        .next()
        .ok_or("usage: scoped_patch_probe FIXTURE [N] [FIRST]")?;
    let count: usize = args.next().map_or(Ok(1000), |value| value.parse())?;
    let first: usize = args.next().map_or(Ok(200), |value| value.parse())?;
    let fixture = std::fs::read(path)?;

    let mut first_ns = Vec::with_capacity(first);
    for index in 0..first {
        let fresh = Package::from_bytes(&fixture)?;
        let begin = Instant::now();
        let mut edit = fresh.edit_external_media()?;
        edit.set_path(1, Some(format!("first-{index:04}.avi")))?;
        let commit = edit.commit()?;
        first_ns.push(begin.elapsed().as_nanos());
        drop(commit);
    }

    let source = Package::from_bytes(&fixture)?;
    let source_bytes = source.to_bytes()?;

    let started = Instant::now();
    let mut patches = Vec::with_capacity(count);
    let mut commit_ns = Vec::with_capacity(count);
    let mut target_len = 0;
    for index in 0..count {
        let begin = Instant::now();
        let mut edit = source.edit_external_media()?;
        edit.set_path(1, Some(format!("clip-{index:04}.avi")))?;
        let commit = edit.commit()?;
        commit_ns.push(begin.elapsed().as_nanos());
        let (package, patch) = commit.into_parts();
        target_len = package.to_bytes()?.len();
        patches.push(patch);
    }
    let commit_total = started.elapsed();
    let retained_bytes: usize = patches.iter().map(retained).sum();

    let mut apply_ns = Vec::with_capacity(count);
    let mut inverse_ns = Vec::with_capacity(count);
    for patch in &patches {
        let begin = Instant::now();
        let replayed = source.apply_external_media(patch)?;
        apply_ns.push(begin.elapsed().as_nanos());
        let inverse = patch.inverse();
        let begin = Instant::now();
        let restored = replayed.package().apply_external_media(&inverse)?;
        inverse_ns.push(begin.elapsed().as_nanos());
        if restored.package().to_bytes()? != source_bytes {
            return Err("the inverse did not restore the source bytes".into());
        }
    }

    println!(
        "{{\"patches\":{count},\"source_len\":{},\"target_len\":{target_len},\
         \"retained_patch_bytes\":{retained_bytes},\"retained_per_patch\":{},\
         \"commit_total_ms\":{:.3},\"commit_median_us\":{:.3},\
         \"apply_median_us\":{:.3},\"inverse_apply_median_us\":{:.3},\
         \"first_commit_median_us\":{:.3}}}",
        source_bytes.len(),
        retained_bytes / count.max(1),
        commit_total.as_secs_f64() * 1000.0,
        median(&mut commit_ns),
        median(&mut apply_ns),
        median(&mut inverse_ns),
        median(&mut first_ns),
    );
    Ok(())
}

#![allow(
    clippy::print_stdout,
    reason = "this measurement probe intentionally prints its JSON result"
)]

//! Source-backed text-splice probe (record 0357a).
//!
//! `text_splice_probe make PATH` writes a 1,000-shape deck (10 slides of 100
//! text boxes, each `s00-shape-000`) built by `litchi_ppt::Builder`.
//!
//! `text_splice_probe MODE PATH [N]` then runs `N` iterations of: open the
//! deck from the file (`Package::open`, positional), stage one same-length
//! replacement of the text of slide 5 / shape 50, commit, and stream the
//! published package to a counting sink (`write_to`). `MODE` is
//!
//! - `text`: `Package::edit_text` (the source-backed splice family);
//! - `primary`: `Package::edit().set_shape_text` (the primary edit).
//!
//! Prints one JSON line with the median microseconds of each phase (open;
//! edit, stage and commit; the commit alone; `write_to`) and of the whole
//! iteration. Peak RSS is measured outside (`/usr/bin/time`).

use std::io::{self, Write};
use std::time::Instant;

use litchi_ppt::Package;
use litchi_ppt::package::edit::{Position, Target};

struct Counter(u64);

impl Write for Counter {
    fn write(&mut self, data: &[u8]) -> io::Result<usize> {
        self.0 += data.len() as u64;
        Ok(data.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
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

fn make(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut builder = litchi_ppt::Builder::new();
    for slide_index in 0..10 {
        let slide = builder.add_slide()?;
        for shape_index in 0..100_i32 {
            let x = 10 + (shape_index % 10) * 60;
            let y = 10 + (shape_index / 10) * 40;
            let text = format!("s{slide_index:02}-shape-{shape_index:03}");
            builder.add_textbox(slide, x, y, 55, 35, &text)?;
        }
    }
    std::fs::write(path, builder.build()?.to_bytes()?)?;
    Ok(())
}

fn stage_text(package: &Package, target: Target, text: &str) -> Result<(Package, u128), String> {
    let mut edit = package
        .edit_text(target)
        .map_err(|error| format!("{error}"))?;
    edit.set_text(text).map_err(|error| format!("{error}"))?;
    let staged = Instant::now();
    let package = edit
        .commit()
        .map_err(|error| format!("{error}"))?
        .into_package();
    Ok((package, staged.elapsed().as_nanos()))
}

fn stage_primary(package: &Package, target: Target, text: &str) -> Result<(Package, u128), String> {
    let mut edit = package.edit().map_err(|error| format!("{error}"))?;
    edit.set_shape_text(target, text)
        .map_err(|error| format!("{error}"))?;
    let staged = Instant::now();
    let package = edit
        .commit()
        .map_err(|error| format!("{error}"))?
        .into_package();
    Ok((package, staged.elapsed().as_nanos()))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let mode = args
        .next()
        .ok_or("usage: text_splice_probe make|text|primary PATH [N]")?;
    let path = args.next().ok_or("missing PATH")?;
    if mode == "make" {
        return make(&path);
    }
    let count: usize = args.next().map_or(Ok(50), |value| value.parse())?;
    let target = Target::new(Position::new(5), Position::new(50));
    let (mut open, mut stage, mut write, mut total) = (vec![], vec![], vec![], vec![]);
    let mut commit = vec![];
    let mut written = 0;
    for iteration in 0..count {
        let replacement = format!("x{:02}-shape-050", iteration % 100);
        let start = Instant::now();
        let package = Package::open(&path)?;
        let opened = Instant::now();
        let (published, commit_ns) = match mode.as_str() {
            "text" => stage_text(&package, target, &replacement)?,
            "primary" => stage_primary(&package, target, &replacement)?,
            _ => return Err(format!("mode {mode} is not available in this build").into()),
        };
        let committed = Instant::now();
        let mut sink = Counter(0);
        published.write_to(&mut sink)?;
        let finished = Instant::now();
        written = sink.0;
        open.push((opened - start).as_nanos());
        stage.push((committed - opened).as_nanos());
        commit.push(commit_ns);
        write.push((finished - committed).as_nanos());
        total.push((finished - start).as_nanos());
    }
    println!(
        "{{\"mode\":\"{mode}\",\"iterations\":{count},\"written_bytes\":{written},\"open_us\":{:.1},\"edit_commit_us\":{:.1},\"commit_us\":{:.1},\"write_to_us\":{:.1},\"total_us\":{:.1}}}",
        median(&mut open),
        median(&mut stage),
        median(&mut commit),
        median(&mut write),
        median(&mut total)
    );
    Ok(())
}

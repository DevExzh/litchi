//! Change 0637 probe: the eager PPTX slide catalog and the per-slide semantic
//! text passes.
//!
//! `tools/perf-baseline` has no selector that issues more than one catalog
//! query against a single borrowed `Presentation`, and none that iterates a
//! deck by index, so this probe supplies those caller shapes. Every operation
//! loads the package once, outside the measured loop; one iteration is one
//! operation.
//!
//! Subcommands:
//!   counts <file> <op> <iters>            run `iters` operations, print nothing
//!   time <file> <op> <iters> <samples>    print nanoseconds per operation, one line per sample
//!   oracle <dir>                          per-fixture catalog, lookup, text and refusal census
//!   build <slides> <boxes> <out>          author a deterministic deck
//!
//! Operations:
//!   open           `package.presentation()` only
//!   count1         one `slide_count()` on a fresh borrowed presentation
//!   refs1          one `slide_references()` on a fresh borrowed presentation
//!   refsN:<k>      `k` x `slide_references()` on ONE borrowed presentation
//!   countN:<k>     `k` x `slide_count()` on ONE borrowed presentation
//!   slide1         one `slide(0)`
//!   slideall       `for i in 0..n { slide(i) }` on one presentation (the by-index iteration)
//!   slides         one `slides()`
//!   findname       one `find_slide(Key::Name(<last slide name>))`
//!   session        `slide_count()` + `slide(mid)` + `slides()` on one presentation
//!   text           one `text()`
//!   writetext      one `write_text_to(<counting sink>)`
//!   slidetext      `slides()` then `text()` of every slide (the PPTX-4 three-pass region)

use std::env;
use std::error::Error;
use std::fmt::Write as _;
use std::fs;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::time::Instant;

use litchi_core::TextOutputOptions;
use litchi_pptx::Package;
use litchi_pptx::slide::Key;
use sha2::{Digest, Sha256};

type BoxError = Box<dyn Error + Send + Sync>;

fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hasher
        .finalize()
        .iter()
        .fold(String::with_capacity(64), |mut out, byte| {
            let _ = write!(out, "{byte:02x}");
            out
        })
}

/// A sink that retains nothing but its length and digest.
struct CountingSink {
    hasher: Sha256,
    len: usize,
}

impl CountingSink {
    fn new() -> Self {
        Self {
            hasher: Sha256::new(),
            len: 0,
        }
    }

    fn finish(self) -> (usize, String) {
        let digest = self
            .hasher
            .finalize()
            .iter()
            .fold(String::with_capacity(64), |mut out, byte| {
                let _ = write!(out, "{byte:02x}");
                out
            });
        (self.len, digest)
    }
}

impl std::io::Write for CountingSink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.hasher.update(buf);
        self.len += buf.len();
        Ok(buf.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum Op {
    Open,
    Count1,
    CountN(usize),
    Refs1,
    RefsN(usize),
    Slide1,
    SlideAll,
    Slides,
    FindName,
    Session,
    Text,
    WriteText,
    SlideText,
}

fn parse_op(value: &str) -> Result<Op, BoxError> {
    if let Some(rest) = value.strip_prefix("countN:") {
        return Ok(Op::CountN(rest.parse()?));
    }
    if let Some(rest) = value.strip_prefix("refsN:") {
        return Ok(Op::RefsN(rest.parse()?));
    }
    Ok(match value {
        "open" => Op::Open,
        "count1" => Op::Count1,
        "refs1" => Op::Refs1,
        "slide1" => Op::Slide1,
        "slideall" => Op::SlideAll,
        "slides" => Op::Slides,
        "findname" => Op::FindName,
        "session" => Op::Session,
        "text" => Op::Text,
        "writetext" => Op::WriteText,
        "slidetext" => Op::SlideText,
        other => return Err(format!("unknown operation '{other}'").into()),
    })
}

/// One operation. Returns a value derived from the result so nothing is
/// optimized away.
fn run_op(package: &Package, op: Op, name: &str) -> Result<usize, BoxError> {
    let presentation = package.presentation()?;
    Ok(match op {
        Op::Open => std::hint::black_box(&presentation).slide_size().is_ok() as usize,
        Op::Count1 => presentation.slide_count()?,
        Op::CountN(k) => {
            let mut total = 0usize;
            for _ in 0..k {
                total = total.wrapping_add(presentation.slide_count()?);
            }
            total
        },
        Op::Refs1 => presentation.slide_references()?.len(),
        Op::RefsN(k) => {
            let mut total = 0usize;
            for _ in 0..k {
                total = total.wrapping_add(presentation.slide_references()?.len());
            }
            total
        },
        Op::Slide1 => presentation.slide(0)?.is_some() as usize,
        Op::SlideAll => {
            let count = presentation.slide_count()?;
            let mut total = 0usize;
            for index in 0..count {
                total += presentation.slide(index)?.is_some() as usize;
            }
            total
        },
        Op::Slides => presentation.slides()?.len(),
        // A deck whose slides share a name refuses with `AmbiguousSlideName`
        // after the same complete scan, so the refusal is counted, not
        // propagated: the work under measurement is identical.
        Op::FindName => match presentation.find_slide(Key::Name(name)) {
            Ok(found) => found.is_some() as usize,
            Err(_ambiguous) => 1,
        },
        Op::Session => {
            let count = presentation.slide_count()?;
            let middle = presentation.slide(count / 2)?.is_some() as usize;
            count + middle + presentation.slides()?.len()
        },
        Op::Text => presentation.text()?.len(),
        Op::WriteText => {
            let mut sink = CountingSink::new();
            presentation.write_text_to(&mut sink, TextOutputOptions::default())?;
            sink.finish().0
        },
        Op::SlideText => {
            let mut total = 0usize;
            for slide in presentation.slides()? {
                total += slide.text()?.len();
            }
            total
        },
    })
}

/// The name `findname` looks up: the last slide's name, the worst case for the
/// linear by-name scan.
fn last_slide_name(package: &Package) -> Result<String, BoxError> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    match slides.last() {
        Some(slide) => Ok(slide.name()?),
        None => Ok(String::new()),
    }
}

fn counts(path: &str, op: &str, iters: usize) -> Result<(), BoxError> {
    let package = Package::from_bytes(&fs::read(path)?)?;
    let op = parse_op(op)?;
    let name = last_slide_name(&package)?;
    let mut total = 0usize;
    for _ in 0..iters {
        total = total.wrapping_add(run_op(&package, op, &name)?);
    }
    std::hint::black_box(total);
    Ok(())
}

fn time(path: &str, op: &str, iters: usize, samples: usize) -> Result<(), BoxError> {
    let package = Package::from_bytes(&fs::read(path)?)?;
    let parsed = parse_op(op)?;
    let name = last_slide_name(&package)?;
    // Warm up outside the reported samples.
    for _ in 0..3 {
        for _ in 0..iters {
            std::hint::black_box(run_op(&package, parsed, &name)?);
        }
    }
    let mut out = String::new();
    for _ in 0..samples {
        let started = Instant::now();
        for _ in 0..iters {
            std::hint::black_box(run_op(&package, parsed, &name)?);
        }
        let elapsed = started.elapsed().as_nanos();
        let _ = writeln!(out, "{}", elapsed / iters as u128);
    }
    print!("{out}");
    Ok(())
}

fn collect_pptx(dir: &Path, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            collect_pptx(&path, out)?;
        } else if path.extension().is_some_and(|ext| ext == "pptx") {
            out.push(path);
        }
    }
    Ok(())
}

/// Every catalog-dependent public projection of one fixture, as text.
fn oracle_one(path: &Path, out: &mut String) -> Result<(), BoxError> {
    let label = path
        .file_name()
        .map_or_else(|| path.display().to_string(), |n| n.to_string_lossy().into());
    let bytes = match fs::read(path) {
        Ok(bytes) => bytes,
        Err(error) => {
            let _ = writeln!(out, "{label}\tread\tERR\t{error}");
            return Ok(());
        },
    };
    let package = match Package::from_bytes(&bytes) {
        Ok(package) => package,
        Err(error) => {
            let _ = writeln!(out, "{label}\topen\tERR\t{error:?}");
            return Ok(());
        },
    };
    let presentation = match package.presentation() {
        Ok(presentation) => presentation,
        Err(error) => {
            let _ = writeln!(out, "{label}\tpresentation\tERR\t{error:?}");
            return Ok(());
        },
    };
    match presentation.slide_references() {
        Ok(references) => {
            let _ = writeln!(out, "{label}\tslide_references\tok\t{}", references.len());
            for (index, reference) in references.iter().enumerate() {
                let _ = writeln!(
                    out,
                    "{label}\treference[{index}]\tok\tid={} rid={}",
                    reference.id(),
                    reference.relationship_id()
                );
            }
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tslide_references\tERR\t{error:?}");
        },
    }
    let count = match presentation.slide_count() {
        Ok(count) => {
            let _ = writeln!(out, "{label}\tslide_count\tok\t{count}");
            Some(count)
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tslide_count\tERR\t{error:?}");
            None
        },
    };
    match presentation.slide_size() {
        Ok((cx, cy)) => {
            let _ = writeln!(out, "{label}\tslide_size\tok\t{cx}x{cy}");
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tslide_size\tERR\t{error:?}");
        },
    }
    // Slide list, by index, including one position past the end.
    let probe_upper = count.unwrap_or(0).saturating_add(1);
    let mut names: Vec<String> = Vec::new();
    for index in 0..probe_upper {
        match presentation.slide(index) {
            Ok(None) => {
                let _ = writeln!(out, "{label}\tslide[{index}]\tnone\t-");
            },
            Ok(Some(slide)) => {
                let partname = slide.part().part().partname().to_string();
                let name = match slide.name() {
                    Ok(name) => {
                        names.push(name.clone());
                        name
                    },
                    Err(error) => format!("ERR {error:?}"),
                };
                let text = match slide.text() {
                    Ok(text) => format!("len={} sha={}", text.len(), sha256_hex(text.as_bytes())),
                    Err(error) => format!("ERR {error:?}"),
                };
                let _ = writeln!(
                    out,
                    "{label}\tslide[{index}]\tok\tpart={partname} name={name} text={text}"
                );
            },
            Err(error) => {
                let _ = writeln!(out, "{label}\tslide[{index}]\tERR\t{error:?}");
            },
        }
    }
    // `slides()` as a whole, and the by-name lookup for every observed name
    // plus one name that cannot exist.
    match presentation.slides() {
        Ok(slides) => {
            let _ = writeln!(out, "{label}\tslides\tok\t{}", slides.len());
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tslides\tERR\t{error:?}");
        },
    }
    names.sort();
    names.dedup();
    names.push("litchi-0637-name-that-cannot-exist".to_owned());
    for name in &names {
        match presentation.find_slide(Key::Name(name.as_str())) {
            Ok(None) => {
                let _ = writeln!(out, "{label}\tfind[{name}]\tnone\t-");
            },
            Ok(Some(slide)) => {
                let _ = writeln!(
                    out,
                    "{label}\tfind[{name}]\tok\t{}",
                    slide.part().part().partname()
                );
            },
            Err(error) => {
                let _ = writeln!(out, "{label}\tfind[{name}]\tERR\t{error:?}");
            },
        }
    }
    match presentation.text() {
        Ok(text) => {
            let _ = writeln!(
                out,
                "{label}\ttext\tok\tlen={} sha={}",
                text.len(),
                sha256_hex(text.as_bytes())
            );
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\ttext\tERR\t{error:?}");
        },
    }
    let mut sink = CountingSink::new();
    match presentation.write_text_to(&mut sink, TextOutputOptions::default()) {
        Ok(report) => {
            let (len, digest) = sink.finish();
            let _ = writeln!(
                out,
                "{label}\twrite_text_to\tok\tlen={len} sha={digest} report={report:?}"
            );
        },
        Err(error) => {
            let (len, digest) = sink.finish();
            let _ = writeln!(
                out,
                "{label}\twrite_text_to\tERR\tlen={len} sha={digest} {error:?}"
            );
        },
    }
    match presentation.content_parts() {
        Ok(parts) => {
            let _ = writeln!(out, "{label}\tcontent_parts\tok\t{}", parts.len());
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tcontent_parts\tERR\t{error:?}");
        },
    }
    match presentation.hyperlinks() {
        Ok(links) => {
            let _ = writeln!(out, "{label}\thyperlinks\tok\t{}", links.len());
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\thyperlinks\tERR\t{error:?}");
        },
    }
    match presentation.slide_masters() {
        Ok(masters) => {
            let _ = writeln!(out, "{label}\tslide_masters\tok\t{}", masters.len());
        },
        Err(error) => {
            let _ = writeln!(out, "{label}\tslide_masters\tERR\t{error:?}");
        },
    }
    Ok(())
}

fn oracle(dir: &str) -> Result<(), BoxError> {
    let mut files = Vec::new();
    collect_pptx(Path::new(dir), &mut files)?;
    files.sort();
    let mut out = String::new();
    for path in &files {
        oracle_one(path, &mut out)?;
    }
    let _ = writeln!(out, "# fixtures={}", files.len());
    print!("{out}");
    std::io::stdout().flush()?;
    Ok(())
}

fn build(slides: usize, boxes: usize, out: &str) -> Result<(), BoxError> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for index in 0..slides {
            let slide = presentation.add_slide()?;
            slide.set_title(&format!("Slide {index:04}"));
            for box_index in 0..boxes {
                slide.add_text_box(
                    &format!(
                        "Slide {index:04} line {box_index} - a sentence of ordinary presentation body text that a real deck would carry."
                    ),
                    914_400,
                    1_828_800 + 457_200 * i64::try_from(box_index)?,
                    7_315_200,
                    457_200,
                );
            }
        }
    }
    let bytes = package.to_bytes()?;
    fs::write(out, &bytes)?;
    println!("{}\t{}\t{}", out, bytes.len(), sha256_hex(&bytes));
    Ok(())
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("counts") => counts(&args[2], &args[3], args[4].parse()?),
        Some("time") => time(&args[2], &args[3], args[4].parse()?, args[5].parse()?),
        Some("oracle") => oracle(&args[2]),
        Some("build") => build(args[2].parse()?, args[3].parse()?, &args[4]),
        _ => Err("usage: counts|time|oracle|build ...".into()),
    }
}

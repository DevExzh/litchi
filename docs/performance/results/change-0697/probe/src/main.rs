//! Isolated default-profile MCE attribution, not a presentation benchmark.
use litchi_ooxml_common::mce::process_ooxml;
use sha2::{Digest, Sha256};
use std::{borrow::Cow, env, error::Error, fs, hint::black_box, path::Path, time::Instant};

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 5 {
        return Err("usage: probe <sequence-file> <warmups> <samples> <batch>".into());
    }
    let sequence = Path::new(&args[1]);
    let warmups: usize = args[2].parse()?;
    let samples: usize = args[3].parse()?;
    let batch: usize = args[4].parse()?;
    if samples == 0 || samples > 100_000 || warmups > 10_000 || batch == 0 || batch > 10_000 {
        return Err("invalid bounded iteration count".into());
    }
    let base = sequence.parent().ok_or("sequence has no parent")?;
    let mut owners = Vec::new();
    let mut paths = Vec::new();
    let mut order = Vec::new();
    for line in fs::read_to_string(sequence)?.lines() {
        if line.is_empty() {
            continue;
        }
        let path = base.join(line);
        let index = if let Some(index) = paths.iter().position(|p| p == &path) {
            index
        } else {
            owners.push(fs::read(&path)?);
            paths.push(path);
            owners.len() - 1
        };
        order.push(index);
    }
    if order.is_empty() || order.len() > 1024 {
        return Err("invalid sequence".into());
    }
    for (index, input) in owners.iter().enumerate() {
        let output = process_ooxml(input)?;
        println!(
            "IDENTITY\tindex={index}\tinput_bytes={}\tinput_sha256={}\toutput_bytes={}\toutput_sha256={}\tborrowed={}",
            input.len(),
            digest(input),
            output.len(),
            digest(output.as_ref()),
            matches!(output, Cow::Borrowed(_))
        );
    }
    let run = || -> Result<(), litchi_ooxml_common::mce::Error> {
        for _ in 0..batch {
            for &index in &order {
                drop(black_box(process_ooxml(black_box(&owners[index]))?));
            }
        }
        Ok(())
    };
    for _ in 0..warmups {
        run()?;
    }
    let mut times = Vec::with_capacity(samples);
    for _ in 0..samples {
        let start = Instant::now();
        run()?;
        times.push(start.elapsed().as_nanos());
    }
    println!(
        "META\tprobe=0697\twarmups={warmups}\tsamples={samples}\tbatch={batch}\tcalls={}",
        order.len()
    );
    for (index, elapsed_ns) in times.into_iter().enumerate() {
        println!("SAMPLE\tindex={index}\telapsed_ns={elapsed_ns}");
    }
    Ok(())
}

fn digest(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

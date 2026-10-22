//! Clone/drop probe for the W5 OLE package owner.
//!
//! The owner migration changed a package clone from one owner-state `Arc`
//! increment to one increment per shared field.  This probe keeps an admitted
//! owner alive and measures only repeated clone/drop operations for small and
//! larger captured catalogs.  Opening and CSV formatting are outside the
//! timed loop, and no package bytes are rendered by the operation.

use std::error::Error as StdError;
use std::hint::black_box;
use std::io::Cursor;
use std::time::Instant;

use litchi_cfb::OleWriter;
use litchi_ole_common::object::{Limits as ObjectLimits, Targets};

#[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
use litchi_ole_common::object::Edit as ObjectEdit;
#[cfg(feature = "current-owner-api")]
use litchi_ole_common::object::{OpenOptions, Package as ObjectPackage};

#[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
type OwnerSource = litchi_ole_common::object::Snapshot;
#[cfg(feature = "current-owner-api")]
type OwnerSource = litchi_ole_common::object::Package;

const SMALL_SIBLINGS: usize = 1;
const LARGE_SIBLINGS: usize = 1_024;
const TARGET_NAME: &str = "ole_w5_clone_target.bin";
const TARGET_BYTES: &[u8] = b"ole-w5-clone-target";

fn main() -> Result<(), Box<dyn StdError>> {
    let (warmup, samples, batch) = arguments()?;
    let limits = limits(LARGE_SIBLINGS + 1);
    let small = archive(SMALL_SIBLINGS)?;
    let large = archive(LARGE_SIBLINGS)?;

    println!("route,size,sample,nanoseconds,batch,archive_bytes");
    for (size, bytes) in [("small", small), ("large", large)] {
        let source = open_source(&bytes, limits)?;
        verify_source(&source)?;
        for iteration in 0..warmup.checked_add(samples).ok_or("sample count overflow")? {
            let started = Instant::now();
            for _ in 0..batch {
                let clone = source.clone();
                black_box(&clone);
                drop(clone);
            }
            let elapsed = started.elapsed().as_nanos();
            if iteration >= warmup {
                println!(
                    "clone_drop,{size},{},{elapsed},{batch},{}",
                    iteration - warmup,
                    bytes.len()
                );
            }
        }
        black_box(source);
    }
    Ok(())
}

fn arguments() -> Result<(usize, usize, usize), Box<dyn StdError>> {
    let mut warmup = 10usize;
    let mut samples = 100usize;
    let mut batch = 100_000usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--warmup" => warmup = arguments.next().ok_or("--warmup needs a value")?.parse()?,
            "--samples" => samples = arguments.next().ok_or("--samples needs a value")?.parse()?,
            "--batch" => batch = arguments.next().ok_or("--batch needs a value")?.parse()?,
            "--help" | "-h" => {
                println!(
                    "usage: ole_package_clone_probe [--warmup N] [--samples N] [--batch N]"
                );
                return Ok((warmup, samples, batch));
            },
            other => return Err(format!("unknown argument {other:?}").into()),
        }
    }
    if samples == 0 || batch == 0 {
        return Err("--samples and --batch must be nonzero".into());
    }
    Ok((warmup, samples, batch))
}

fn limits(streams: usize) -> ObjectLimits {
    ObjectLimits {
        max_objects: 1,
        max_storage_depth: 1,
        max_streams_per_object: streams,
        max_streams: streams,
        max_stream_size: 1 << 20,
        max_object_size: 8 << 20,
        max_total_size: 8 << 20,
    }
}

fn archive(siblings: usize) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut writer = OleWriter::new();
    for index in 0..siblings {
        let name = format!("ole_w5_clone_sibling_{index:04}");
        let payload = format!("clone probe sibling {index:04}");
        writer.create_stream(&[name.as_str()], payload.as_bytes())?;
    }
    writer.create_stream(&[TARGET_NAME], TARGET_BYTES)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn open_source(bytes: &[u8], limits: ObjectLimits) -> Result<OwnerSource, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(ObjectEdit::open(bytes.to_vec(), Targets::default(), limits)?.staged())
    }
    #[cfg(feature = "current-owner-api")]
    {
        Ok(ObjectPackage::open_with_options(
            bytes.to_vec(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(limits),
        )?)
    }
}

fn verify_source(source: &OwnerSource) -> Result<(), Box<dyn StdError>> {
    if source.stream(&[TARGET_NAME.to_owned()]) != Some(TARGET_BYTES) {
        return Err("clone probe target stream differs from the admitted source".into());
    }
    Ok(())
}

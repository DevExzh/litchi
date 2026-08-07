//! Re-encode every IWA component in an iWork package and save the result.
//!
//! This is useful for validating format compatibility and as a starting point
//! for custom object-level physical transformations. Application semantics
//! remain the responsibility of the focused Pages, Numbers, and Keynote
//! adapters.

#![allow(
    clippy::print_stdout,
    reason = "This command-line example intentionally reports its component count."
)]

mod support;

use std::env;
use std::fs;
use std::path::Path;

use litchi_iwa_archive::package::{self, Catalog};
use litchi_iwa_archive::{Error, Limits};
use litchi_iwa_core::{Archive, SnappyStream};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: roundtrip_iwork <input> <output>")?;
    let output = arguments
        .next()
        .ok_or("usage: roundtrip_iwork <input> <output>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let limits = Limits::default();
    let source = support::read_package(Path::new(&input), limits)?;
    let catalog = Catalog::from_shared_bytes(source)?;
    let component_capacity = catalog.len();
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(component_capacity)
        .map_err(|_error| Error::Allocation {
            resource: "round-trip package entries",
            amount: component_capacity,
        })?;
    let mut archive_count = 0usize;
    for entry in catalog {
        let (name, mut data) = entry.into_parts();
        #[allow(
            clippy::case_sensitive_file_extension_comparisons,
            reason = "IWA member names are case-sensitive protocol names."
        )]
        let is_iwa = name.ends_with(".iwa");
        if is_iwa {
            let decompressed = SnappyStream::decompress(&data)?;
            let archive = Archive::parse(decompressed.as_bytes())?;
            data = SnappyStream::compress(&archive.to_bytes()?)?;
            archive_count = archive_count.checked_add(1).ok_or_else(|| {
                Error::InvalidBundle("round-trip IWA component count overflows usize".to_owned())
            })?;
        }
        entries.push((name, data));
    }

    let artifact = package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        limits,
    )?;
    fs::write(output, artifact)?;
    println!("re-encoded {archive_count} IWA components");
    Ok(())
}

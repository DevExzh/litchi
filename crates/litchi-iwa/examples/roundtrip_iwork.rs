//! Re-encode every IWA component in an iWork package and save the result.
//!
//! This is useful for validating format compatibility and as a starting point
//! for custom object-level edits.

use std::env;

use litchi_iwa::raw::package::IWorkPackage;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: roundtrip_iwork <input> <output>")?;
    let output = arguments
        .next()
        .ok_or("usage: roundtrip_iwork <input> <output>")?;

    let mut package = IWorkPackage::open(input)?;
    let archive_names: Vec<String> = package
        .entry_names()
        .filter(|name| name.ends_with(".iwa"))
        .map(ToOwned::to_owned)
        .collect();
    for name in &archive_names {
        package.update_archive(name, |_| Ok(()))?;
    }
    package.save(output)?;
    println!("re-encoded {} IWA components", archive_names.len());
    Ok(())
}

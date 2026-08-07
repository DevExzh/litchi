//! Rebuild every IWA component without changing its decompressed contents.
//!
//! This is a migration-verification utility: it exercises bounded ZIP ingress,
//! Snappy framing, archive-header decoding/encoding, and physical package
//! reassembly before a native iWork application opens the result.

use std::env;
use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::io;
use std::path::PathBuf;

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_core::{Archive, SnappyStream};

fn required_path(argument: Option<OsString>, label: &str) -> io::Result<PathBuf> {
    argument.map(PathBuf::from).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("missing {label}; usage: native_noop <input> <output>"),
        )
    })
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = env::args_os().skip(1);
    let input = required_path(arguments.next(), "input path")?;
    let output = required_path(arguments.next(), "output path")?;
    if arguments.next().is_some() {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "too many arguments; usage: native_noop <input> <output>",
        )
        .into());
    }

    let source = fs::read(&input)?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut rebuilt = Vec::<(String, Vec<u8>)>::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        if entry.is_opaque() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!("IWA component is opaque: {}", entry.name()),
            )
            .into());
        }
        let decompressed = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(decompressed.as_bytes())?;
        let encoded = archive.to_bytes()?;
        if encoded != decompressed.as_bytes() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!("IWA no-op changed decompressed component {}", entry.name()),
            )
            .into());
        }
        rebuilt.push((entry.name().to_owned(), SnappyStream::compress(&encoded)?));
    }

    let edits = rebuilt
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect::<Vec<_>>();
    let artifact = catalog.reassemble_to_bytes(&edits, Limits::default())?;
    fs::write(output, artifact)?;
    Ok(())
}

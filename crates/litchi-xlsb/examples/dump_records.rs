//! Print the validated record inventory of one decompressed XLSB binary part.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "this command-line example intentionally prints each record and keeps its small error type next to main"
)]

use std::error::Error;
use std::ffi::OsString;
use std::fs;

use litchi_xlsb::raw::Records;

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os().nth(1).ok_or_else(|| Usage {
        program: std::env::args_os()
            .next()
            .unwrap_or_else(|| OsString::from("dump_records")),
    })?;
    let bytes = fs::read(&path)?;
    for (index, record) in Records::new(&bytes).enumerate() {
        let record = record?;
        println!(
            "{index:08} offset={:#010x} kind={} len={}",
            record.offset(),
            record.kind(),
            record.len()
        );
    }
    Ok(())
}

#[derive(Debug)]
struct Usage {
    program: OsString,
}

impl std::fmt::Display for Usage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "usage: {} <decompressed-xlsb-part.bin>",
            self.program.to_string_lossy()
        )
    }
}

impl Error for Usage {}

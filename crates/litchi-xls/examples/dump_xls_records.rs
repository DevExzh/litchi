//! Debug utility: dump BIFF records from an XLS Workbook stream
//!
//! Usage:
//!
//! ```bash
//! cargo run --example dump_xls_records --features ole --no-default-features -- path/to/file.xls
//! ```

use litchi_cfb::OleFile;
use litchi_xls::records::{BiffVersion, RecordIter};
use std::env;
use std::fs::File;
use std::io::{BufReader, Cursor};

fn name_for_sid(sid: u16) -> &'static str {
    match sid {
        0x0809 => "BOF",
        0x000A => "EOF",
        0x0042 => "CODEPAGE",
        0x0022 => "DATE1904",
        0x003D => "WINDOW1",
        0x0085 => "BOUNDSHEET8",
        0x0018 => "NAME",
        0x00FC => "SST",
        0x003C => "CONTINUE",
        0x00E0 => "XF",
        0x0031 => "FONT",
        0x041E => "FORMAT",
        0x023E => "WINDOW2",
        0x0081 => "WSBOOL",
        0x0200 => "DIMENSIONS",
        0x0203 => "NUMBER",
        0x00FD => "LABELSST",
        0x0205 => "BOOLERR",
        _ => "?",
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .unwrap_or_else(|| "minimal.xls".to_string());
    eprintln!("Dumping BIFF records for {path} (Workbook stream)\n");

    let file = File::open(&path)?;
    let reader = BufReader::new(file);
    let mut ole = OleFile::open(reader)?;

    let workbook_stream = ole.open_stream(&["Workbook"])?;
    let len = workbook_stream.len();
    eprintln!("Workbook stream size: {len} bytes\n");

    let cursor = Cursor::new(&workbook_stream[..]);
    let iter = RecordIter::new(cursor)?;

    let mut offset: u64 = 0;
    let mut in_workbook = true;
    let mut sheet_index: usize = 0;

    for rec in iter {
        let rec = rec?;
        let sid = rec.header.record_type;
        let len = rec.header.data_len;
        let name = name_for_sid(sid);

        if sid == 0x0809 {
            // BOF: determine substream type
            let version_raw = u16::from_le_bytes([rec.data[0], rec.data[1]]);
            let dt = u16::from_le_bytes([rec.data[2], rec.data[3]]);
            let ver = BiffVersion::from_bof_version(version_raw).unwrap_or(BiffVersion::Biff8);
            if in_workbook {
                println!("{offset:06X}: BOF(Workbook) ver={ver:?} dt=0x{dt:04X} len={len}",);
            } else {
                println!(
                    "{offset:06X}: BOF(Sheet#{sheet_index}) ver={ver:?} dt=0x{dt:04X} len={len}",
                );
            }
        } else if sid == 0x000A {
            println!("{offset:06X}: EOF len={len}");
            if in_workbook {
                in_workbook = false;
            } else {
                sheet_index += 1;
            }
        } else {
            if in_workbook {
                println!("{offset:06X}: {name} (0x{sid:04X}) len={len}");
            } else {
                println!("{offset:06X}: [Sheet#{sheet_index}] {name} (0x{sid:04X}) len={len}",);
            }
        }

        offset += 4 + len as u64;
    }

    Ok(())
}

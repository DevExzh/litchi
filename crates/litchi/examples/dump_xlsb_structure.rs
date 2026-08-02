//! Print the validated BIFF12 record inventory of an XLSB package.

use std::error::Error;
use std::ffi::OsString;
use std::fs::File;

use litchi::ooxml::opc::OpcPackage;
use litchi_xlsb::raw::Records;

fn dump_records(label: &str, bytes: &[u8]) -> Result<(), litchi_xlsb::Error> {
    println!("=== {label} records ===");
    for (index, record) in Records::new(bytes).enumerate() {
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

fn main() -> Result<(), Box<dyn Error>> {
    let path = std::env::args_os().nth(1).ok_or_else(|| Usage {
        program: std::env::args_os()
            .next()
            .unwrap_or_else(|| OsString::from("dump_xlsb_structure")),
    })?;
    let package = OpcPackage::from_reader(File::open(&path)?)?;

    println!("=== Parts ===");
    for part in package.iter_parts() {
        println!(
            "part={} content_type={}",
            part.partname(),
            part.content_type()
        );
    }

    let main_part = package.main_document_part()?;
    dump_records("workbook", main_part.blob())?;

    for part in package.iter_parts() {
        let part_name = part.partname().as_str();
        let content_type = part.content_type();

        let is_worksheet = content_type == "application/vnd.ms-excel.worksheet"
            || (part_name.contains("/worksheets/") && part_name.ends_with(".bin"));
        let is_styles =
            content_type == "application/vnd.ms-excel.styles" || part_name.ends_with("/styles.bin");

        if is_worksheet || is_styles {
            dump_records(part_name, part.blob())?;
        }
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
            "usage: {} <workbook.xlsb>",
            self.program.to_string_lossy()
        )
    }
}

impl Error for Usage {}

use litchi::ooxml::opc::OpcPackage;
use litchi::ooxml::xlsb::XlsbRecordIter;
use std::env;
use std::fs::File;
use std::io::Cursor;

fn dump_records(label: &str, data: &[u8]) {
    println!("=== {} records ===", label);
    let cursor = Cursor::new(data);
    let iter = XlsbRecordIter::new(cursor);
    for (idx, rec_res) in iter.enumerate() {
        match rec_res {
            Ok(rec) => {
                println!(
                    "{:04} type=0x{:04X} len={}",
                    idx, rec.header.record_type, rec.header.data_len
                );
            },
            Err(e) => {
                eprintln!("Error at {}: {:?}", idx, e);
                break;
            },
        }
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() != 2 {
        eprintln!("Usage: dump_xlsb_structure <file.xlsb>");
        std::process::exit(1);
    }

    let path = &args[1];
    let file = File::open(path)?;
    let pkg = OpcPackage::from_reader(file)?;

    println!("=== Parts ===");
    for part in pkg.iter_parts() {
        println!(
            "part={} content_type={}",
            part.partname(),
            part.content_type()
        );
    }

    let main_part = pkg.main_document_part()?;
    dump_records("workbook", main_part.blob());

    for part in pkg.iter_parts() {
        let partname = part.partname().to_string();
        let ct = part.content_type();

        if ct == "application/vnd.ms-excel.worksheet"
            || (partname.contains("/worksheets/") && partname.ends_with(".bin"))
        {
            dump_records(&partname, part.blob());
        } else if ct == "application/vnd.ms-excel.styles" || partname.ends_with("/styles.bin") {
            dump_records("/xl/styles.bin", part.blob());
        }
    }

    Ok(())
}

use litchi_cfb::OleFile;
use litchi_ppt::RecordType;
use litchi_ppt::records::record::Record;
use std::env;
use std::fs::File;

fn dump_records(data: &[u8], indent: usize) {
    let mut off = 0;
    while off + 8 <= data.len() {
        match Record::parse(data, off) {
            Ok((rec, consumed)) => {
                if rec.record_type == RecordType::SlideListWithText {
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={}",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                } else if rec.record_type_raw == 1011 {
                    // SlidePersistAtom
                    if rec.data.len() >= 20 {
                        let ref_id = u32::from_le_bytes([
                            rec.data[0],
                            rec.data[1],
                            rec.data[2],
                            rec.data[3],
                        ]);
                        let flags = u32::from_le_bytes([
                            rec.data[4],
                            rec.data[5],
                            rec.data[6],
                            rec.data[7],
                        ]);
                        let nph = u32::from_le_bytes([
                            rec.data[8],
                            rec.data[9],
                            rec.data[10],
                            rec.data[11],
                        ]);
                        let sid = u32::from_le_bytes([
                            rec.data[12],
                            rec.data[13],
                            rec.data[14],
                            rec.data[15],
                        ]);
                        println!(
                            "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} refID={} flags=0x{:X} nph={} slideId={}",
                            rec.record_type_raw,
                            rec.version,
                            rec.instance,
                            rec.data_length,
                            ref_id,
                            flags,
                            nph,
                            sid,
                            indent = "",
                            width = indent
                        );
                    } else {
                        println!(
                            "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={} (short)",
                            rec.record_type_raw,
                            rec.version,
                            rec.instance,
                            rec.data_length,
                            rec.children.len(),
                            indent = "",
                            width = indent
                        );
                    }
                } else if rec.record_type_raw == 4003 {
                    // TxMasterStyleAtom (dump hex payload for analysis)
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={} (hex follows)",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                    let mut line = String::new();
                    for (i, b) in rec.data.iter().enumerate() {
                        if i % 16 == 0 {
                            if !line.is_empty() {
                                println!("{indent:>width$}  {}", line, indent = "", width = indent);
                            }
                            line.clear();
                        }
                        if !line.is_empty() {
                            line.push(' ');
                        }
                        line.push_str(&format!("{:02X}", b));
                    }
                    if !line.is_empty() {
                        println!("{indent:>width$}  {}", line, indent = "", width = indent);
                    }
                } else if rec.record_type_raw == 0x07F0 {
                    // ColorSchemeAtom (dump hex payload for analysis)
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={} (hex follows)",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                    let mut line = String::new();
                    for (i, b) in rec.data.iter().enumerate() {
                        if i % 16 == 0 {
                            if !line.is_empty() {
                                println!("{indent:>width$}  {}", line, indent = "", width = indent);
                            }
                            line.clear();
                        }
                        if !line.is_empty() {
                            line.push(' ');
                        }
                        line.push_str(&format!("{:02X}", b));
                    }
                    if !line.is_empty() {
                        println!("{indent:>width$}  {}", line, indent = "", width = indent);
                    }
                } else if rec.record_type_raw == 0x0FBA || rec.record_type_raw == 0x138B {
                    // Misc atoms used in MainMaster tail (CString and BinaryTagData) - dump hex
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={} (hex follows)",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                    let mut line = String::new();
                    for (i, b) in rec.data.iter().enumerate() {
                        if i % 16 == 0 {
                            if !line.is_empty() {
                                println!("{indent:>width$}  {}", line, indent = "", width = indent);
                            }
                            line.clear();
                        }
                        if !line.is_empty() {
                            line.push(' ');
                        }
                        line.push_str(&format!("{:02X}", b));
                    }
                    if !line.is_empty() {
                        println!("{indent:>width$}  {}", line, indent = "", width = indent);
                    }
                } else if [
                    0x0F9F, 0x0FA0, 0x0FA8, 0x0FA1, 0x0FA2, 0x0FAA, 0x0BC3, 0xF010,
                ]
                .contains(&rec.record_type_raw)
                {
                    // Text and Escher client atoms (Header, Bytes, Chars, Props, Ruler, Placeholder, Anchor)
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={} (hex follows)",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                    let mut line = String::new();
                    for (i, b) in rec.data.iter().enumerate() {
                        if i % 16 == 0 {
                            if !line.is_empty() {
                                println!("{indent:>width$}  {}", line, indent = "", width = indent);
                            }
                            line.clear();
                        }
                        if !line.is_empty() {
                            line.push(' ');
                        }
                        line.push_str(&format!("{:02X}", b));
                    }
                    if !line.is_empty() {
                        println!("{indent:>width$}  {}", line, indent = "", width = indent);
                    }
                } else {
                    println!(
                        "{indent:>width$}- type=0x{:04X} ver=0x{:X} inst={} len={} children={}",
                        rec.record_type_raw,
                        rec.version,
                        rec.instance,
                        rec.data_length,
                        rec.children.len(),
                        indent = "",
                        width = indent
                    );
                }
                // Recurse into container payload either when children are present
                // or when the record version indicates a container (0x0F)
                if !rec.children.is_empty() || rec.version == 0x0F {
                    dump_records(&rec.data, indent + 2);
                }
                if consumed == 0 {
                    break;
                }
                off += consumed;
            },
            Err(_) => {
                off += 1;
                if off + 8 > data.len() {
                    break;
                }
            },
        }
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("usage: dump_ppt <file.ppt>");
    let file = File::open(&path)?;
    let mut ole = OleFile::open(file)?;

    // Try primary PowerPoint Document stream
    let data = ole
        .open_stream(&["PowerPoint Document"])
        .or_else(|_| ole.open_stream(&["PP97_DUALSTORAGE", "PowerPoint Document"]))?;

    println!("Dumping PPT records for {} ({} bytes)", path, data.len());
    dump_records(&data, 0);
    Ok(())
}

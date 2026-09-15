//! Scratch probe (synthetic, survey-only): open one XLSX through the public
//! facade and read one cell on either the eager or the source-backed path.
//! Used under callgrind to count instructions per XML pass on a real-producer
//! worksheet versus a marker-stripped control.

use std::env;
use std::hint::black_box;

fn main() {
    let args: Vec<String> = env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("eager");
    let path = args.get(2).expect("path");
    let addr = args.get(3).map(String::as_str).unwrap_or("H680");
    match mode {
        "eager" => {
            let wb = litchi_xlsx::Workbook::open(path).expect("open");
            let sheet = wb.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("eager ok");
        },
        "source" => {
            let wb = litchi_xlsx::SourceBackedWorkbook::from_path(path).expect("open");
            let sheet = wb.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("source ok");
        },
        "mce" => {
            // `path` is a raw worksheet XML file; run the MCE codec once and
            // report input/output lengths and whether the output was borrowed.
            let bytes = std::fs::read(path).expect("read xml");
            let out = litchi_ooxml_common::mce::process_ooxml(&bytes).expect("mce");
            let borrowed = matches!(out, std::borrow::Cow::Borrowed(_));
            println!("mce in={} out={} borrowed={}", bytes.len(), out.len(), borrowed);
            if env::var_os("MCE_DUMP").is_some() {
                std::fs::write(format!("{path}.mce.out"), out.as_ref()).expect("dump");
            }
            black_box(out);
        },
        _ => panic!("mode"),
    }
}

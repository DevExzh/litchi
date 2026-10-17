//! Record 0327: the worksheet-feature read hot path of XLS and ODS, leaf
//! against unified, for callgrind (`perf` is blocked on the host).
//!
//! `sheet_features_bench <xls|ods> <setup|leaf|unified> <iterations>`
//! builds one workbook with every feature kind (outside the measured
//! loop), then lists the features `iterations` times: the leaf reads
//! `Package::sheet_features` (XLS) or `sheet_features::read` (ODS); the
//! unified mode resolves every feature node and reads its properties.
//! `setup` stops after building, so `mode - setup` is the loop's cost.

use litchi_core::FileFormat;
use litchi_ops::{File, NewNode};

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==";

fn build(format: &str) -> File {
    let (mut file, sheet) = if format == "xls" {
        (
            litchi_ops::create(FileFormat::Xls).expect("create"),
            "/Sheet1",
        )
    } else {
        (
            litchi_ops::open(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/odf/corpus/calc-two-sheets.ods"
            ))
            .expect("open"),
            "/sheet1",
        )
    };
    let nodes = [
        NewNode::new("hyperlink")
            .prop("ref", "A1")
            .prop("url", "https://example.test/"),
        NewNode::new("picture").prop("src", PNG),
        NewNode::new("shape").prop("text", "box"),
        NewNode::new("cf")
            .prop("type", "cellIs")
            .prop("ref", "A1:A9")
            .prop("value", "5"),
        NewNode::new("cf")
            .prop("type", "dataBar")
            .prop("ref", "B1:B9"),
        NewNode::new("dataValidation")
            .prop("type", "list")
            .prop("ref", "C1:C2")
            .prop("formula1", "a,b"),
        NewNode::new("autofilter").prop("range", "A1:C3"),
        NewNode::new("rowbreak").prop("row", "2"),
    ];
    for node in nodes {
        file = file.add(sheet, node).expect("add").file;
    }
    // A fresh snapshot, so no read cache is warm.
    litchi_ops::open_bytes(file.to_bytes().expect("bytes")).expect("reopen")
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let format = args.get(1).map_or("ods", String::as_str);
    let mode = args.get(2).map_or("unified", String::as_str);
    let iterations: usize = args
        .get(3)
        .and_then(|value| value.parse().ok())
        .unwrap_or(10);
    let file = build(format);
    let bytes = file.to_bytes().expect("bytes");
    let sheet = if format == "xls" { "Sheet1" } else { "sheet1" };
    let mut total = 0_usize;
    for _ in 0..iterations {
        match mode {
            "setup" => {},
            "leaf" if format == "xls" => {
                let package = litchi_xls::Package::from_bytes(&bytes).expect("package");
                let features = package
                    .sheet_features(0, litchi_xls::sheet_features::FeatureSet::ALL)
                    .expect("features");
                total += features.formats.len() + features.drawings.len();
            },
            "leaf" => {
                let package = litchi_ods::Package::from_bytes(&bytes).expect("package");
                let workbook = package.workbook().expect("workbook");
                let features =
                    litchi_ods::advanced::sheet_features::read(workbook.content_xml(), sheet)
                        .expect("features");
                total += features.formats.len() + features.drawings.len();
            },
            _ => {
                let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
                for kind in [
                    "cf",
                    "picture",
                    "shape",
                    "dataValidation",
                    "autofilter",
                    "rowbreak",
                ] {
                    for node in file.query(kind).expect("query") {
                        let node = node.expect("node");
                        total += node.view().format().expect("format").count();
                    }
                }
            },
        }
    }
    println!("{format} {mode} {iterations} {total}");
}

//! Count-only diagnostic; fixture creation, oracle, and JSON are outside regions.
#[path = "../../probe-src/src/allocation_metrics.rs"]
mod allocation_metrics;
#[path = "../../probe-src/src/counting_allocator.rs"]
mod counting_allocator;

use litchi_opc::xml_attributes::{BytesStartExt, CheckedAttributes};
use quick_xml::events::BytesStart;
use std::hint::black_box;
use std::io::Write;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args().nth(1).ok_or("missing output")?;
    let mut rows = Vec::new();
    allocation_metrics::enable();
    for count in [0, 1, 2, 3, 4, 5, 8, 9, 16, 17, 31, 32, 33, 64] {
        let mut content = String::from("e");
        for i in 0..count {
            content.push_str(&format!(" n{i}=\"v\""));
        }
        let tag = BytesStart::from_content(content, 1);
        for repeat in 0..3 {
            let region = allocation_metrics::begin();
            let mut seen = 0;
            for attribute in black_box(&tag).checked_attributes() {
                black_box(attribute?);
                seen += 1;
            }
            let sample = region.finish().ok_or("allocation region missing")?;
            assert_eq!(seen, count);
            rows.push(serde_json::json!({"attributes": count, "repeat": repeat, "seen": seen, "allocation": sample}));
        }
    }
    let report = serde_json::json!({
        "schema": "litchi.performance.0794.attribute-diagnostic.v1",
        "iterator_size_bytes": std::mem::size_of::<CheckedAttributes<'_>>(),
        "scope": "Construct, consume, and drop checked iterator; input and output allocation excluded; no timing measurement",
        "rows": rows,
    });
    let mut file = std::fs::OpenOptions::new().write(true).create_new(true).open(output)?;
    file.write_all(serde_json::to_string_pretty(&report)?.as_bytes())?;
    Ok(())
}

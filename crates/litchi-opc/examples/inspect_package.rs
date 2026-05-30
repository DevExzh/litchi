//! Inspect the contents of an OPC package (.docx, .xlsx, .pptx, ...).
//!
//! Iterates every part in the package, printing its partname, content type and
//! blob size, then dumps all package-level relationships.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-opc --example inspect_package
//! cargo run -p litchi-opc --example inspect_package -- path/to/file.docx
//! ```

use litchi_opc::OpcPackage;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path: PathBuf = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| {
            PathBuf::from("test-data/ooxml/docx/documentProperties.docx")
        });

    println!("Opening OPC package: {}", path.display());
    let pkg = OpcPackage::open(&path)?;

    println!("\nPackage contains {} parts:", pkg.part_count());
    println!("{:-<100}", "");
    println!(
        "{:<60} {:<30} {:>10}",
        "Partname", "Content-Type (truncated)", "Size"
    );
    println!("{:-<100}", "");

    // Collect into a Vec so we can sort for stable, human-friendly output.
    let mut parts: Vec<_> = pkg.iter_parts().collect();
    parts.sort_by(|a, b| a.partname().as_str().cmp(b.partname().as_str()));

    for part in parts {
        let ct = part.content_type();
        let ct_short = if ct.len() > 30 {
            format!("{}...", &ct[..27])
        } else {
            ct.to_string()
        };
        println!(
            "{:<60} {:<30} {:>10}",
            part.partname().as_str(),
            ct_short,
            part.blob().len()
        );
    }

    println!("\nPackage-level relationships ({}):", pkg.rels().len());
    println!("{:-<100}", "");
    let mut rels: Vec<_> = pkg.rels().iter().collect();
    rels.sort_by(|a, b| a.r_id().cmp(b.r_id()));
    for rel in rels {
        let mode = if rel.is_external() {
            "External"
        } else {
            "Internal"
        };
        println!(
            "  {} -> {} [{}]\n     type: {}",
            rel.r_id(),
            rel.target_ref(),
            mode,
            rel.reltype(),
        );
    }

    Ok(())
}

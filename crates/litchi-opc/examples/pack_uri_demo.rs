//! Demonstrate `PackURI` parsing, normalization, and resolution.
//!
//! Walks through several typical inputs (absolute partnames, relative
//! references with `..` and `.`, Excel/PowerPoint-style sibling references)
//! and prints the resulting normalized `PackURI` for each.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-opc --example pack_uri_demo
//! ```

use litchi_opc::PackURI;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== PackURI::new (absolute partnames) ===");
    let absolute_inputs = [
        "/word/document.xml",
        "/ppt/slides/slide1.xml",
        "/xl/worksheets/sheet1.xml",
        "/[Content_Types].xml",
        "/",
    ];
    for input in absolute_inputs {
        match PackURI::new(input) {
            Ok(uri) => {
                println!(
                    "  {:<32} -> uri={}, base={}, filename={}, ext={}, idx={:?}",
                    format!("{input:?}"),
                    uri.as_str(),
                    uri.base_uri(),
                    uri.filename(),
                    uri.ext(),
                    uri.idx(),
                );
            },
            Err(e) => println!("  {input:?} -> ERROR: {e}"),
        }
    }

    println!("\n=== PackURI::new error cases ===");
    for input in ["word/document.xml", "relative/path"] {
        match PackURI::new(input) {
            Ok(uri) => println!("  {input:?} -> ok: {uri}"),
            Err(e) => println!("  {input:?} -> ERROR: {e}"),
        }
    }

    println!("\n=== PackURI::from_rel_ref (resolve & normalize) ===");
    // (base_uri, relative_ref)
    let rel_inputs: &[(&str, &str)] = &[
        // Sibling reference inside same dir
        ("/word", "document.xml"),
        // Climb out of /word into /
        ("/word", "../docProps/core.xml"),
        // Multi-step climb
        ("/word/embeddings", "../media/image1.png"),
        // "." current-directory marker
        ("/word", "./styles.xml"),
        // Mixed "./" and "../"
        ("/ppt/slides", "../slideLayouts/slideLayout1.xml"),
        // Already at root
        ("/", "word/document.xml"),
        // Excessive "../" should clamp at root
        ("/word", "../../../oops/escaped.xml"),
    ];
    for (base, rel) in rel_inputs {
        match PackURI::from_rel_ref(base, rel) {
            Ok(uri) => println!(
                "  base={:<22} rel={:<40} -> {}",
                format!("{base:?}"),
                format!("{rel:?}"),
                uri,
            ),
            Err(e) => println!(
                "  base={:<22} rel={:<40} -> ERROR: {}",
                format!("{base:?}"),
                format!("{rel:?}"),
                e,
            ),
        }
    }

    println!("\n=== PackURI::relative_ref (inverse direction) ===");
    let relative_pairs: &[(&str, &str)] = &[
        ("/ppt/slideLayouts/slideLayout1.xml", "/ppt/slides"),
        ("/word/media/image1.png", "/word"),
        ("/docProps/core.xml", "/"),
    ];
    for (target, base) in relative_pairs {
        let uri = PackURI::new(*target)?;
        println!(
            "  target={:<42} base={:<14} -> {}",
            format!("{target:?}"),
            format!("{base:?}"),
            uri.relative_ref(base),
        );
    }

    println!("\n=== PackURI::rels_uri ===");
    for input in ["/word/document.xml", "/ppt/presentation.xml", "/"] {
        let uri = PackURI::new(input)?;
        match uri.rels_uri() {
            Ok(rels) => println!("  {:<28} -> {}", uri.as_str(), rels),
            Err(e) => println!("  {:<28} -> ERROR: {}", uri.as_str(), e),
        }
    }

    Ok(())
}

/// Example demonstrating how to extract content from XML parts in an OOXML document.
///
/// This example shows:
/// - Downcasting parts to XmlPart
/// - Efficient XML parsing and content extraction
/// - Using quick-xml for zero-copy parsing
use litchi::ooxml::OpcPackage;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <path-to-docx-file>", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    println!("Extracting XML content from: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the package
    let package = OpcPackage::open(file_path)?;

    // Access the main document part
    match package.main_document_part() {
        Ok(main_part) => {
            println!("\n📄 Main Document:");
            println!("   Partname: {}", main_part.partname());

            // Get the XML content as a string
            if let Ok(xml_str) = std::str::from_utf8(main_part.blob()) {
                println!("\n📝 XML Content (first 500 chars):");
                let preview = if xml_str.len() > 500 {
                    &xml_str[..500]
                } else {
                    xml_str
                };
                println!("{}", preview);
                if xml_str.len() > 500 {
                    println!("   ... (truncated, {} total bytes)", xml_str.len());
                }
            }

            // If it's an XML part, we can parse it
            println!("\n🔍 XML Structure Analysis:");
            println!("   Searching for common WordprocessingML elements...");

            // Use quick-xml to count elements
            use quick_xml::events::Event;
            use quick_xml::Reader;

            let mut reader = Reader::from_reader(main_part.blob());
            reader.config_mut().trim_text(true);

            let mut element_counts: std::collections::HashMap<String, usize> =
                std::collections::HashMap::new();
            let mut buf = Vec::new();

            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                        let name = String::from_utf8_lossy(e.local_name().as_ref()).to_string();
                        *element_counts.entry(name).or_insert(0) += 1;
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => {
                        eprintln!("   XML parse error: {}", e);
                        break;
                    }
                    _ => {}
                }
                buf.clear();
            }

            // Display element statistics
            println!("\n📊 Element Statistics:");
            let mut counts: Vec<_> = element_counts.iter().collect();
            counts.sort_by(|a, b| b.1.cmp(a.1));

            for (name, count) in counts.iter().take(10) {
                println!("   {}: {}", name, count);
            }
            if counts.len() > 10 {
                println!("   ... and {} more element types", counts.len() - 10);
            }
        }
        Err(e) => {
            eprintln!("Error accessing main document: {}", e);
        }
    }

    // Examine other XML parts
    println!("\n🗂️  Other XML Parts:");
    let xml_parts: Vec<_> = package
        .iter_parts()
        .filter(|part| {
            part.content_type().ends_with("+xml") || part.content_type().ends_with("/xml")
        })
        .collect();

    println!("   Found {} XML parts total", xml_parts.len());
    for part in xml_parts.iter().take(5) {
        println!("   - {}", part.partname());
        println!("     Type: {}", part.content_type());
    }
    if xml_parts.len() > 5 {
        println!("   ... and {} more", xml_parts.len() - 5);
    }

    println!("\n✅ Done!");
    Ok(())
}

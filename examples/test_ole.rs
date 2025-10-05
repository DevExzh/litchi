use litchi::ole::{is_ole_file, OleFile};
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== OLE File Parser Test ===\n");

    // Test 1: Check if file is OLE
    println!("1. Checking if test.doc is an OLE file...");
    let file_data = std::fs::read("test.doc")?;
    if is_ole_file(&file_data) {
        println!("   ✓ File is a valid OLE file\n");
    } else {
        println!("   ✗ File is NOT an OLE file\n");
        return Ok(());
    }

    // Test 2: Open and parse the file
    println!("2. Opening and parsing OLE file...");
    let file = File::open("test.doc")?;
    let mut ole = OleFile::open(file)?;
    println!("   ✓ Successfully opened and parsed\n");

    // Test 3: Get root entry name
    println!("3. Root entry information:");
    if let Some(root_name) = ole.get_root_name() {
        println!("   Root entry name: \"{}\"", root_name);
    }
    println!();

    // Test 4: List all streams
    println!("4. Listing all streams in the file:");
    let streams = ole.list_streams();
    println!("   Found {} stream(s):", streams.len());
    for (i, stream) in streams.iter().enumerate() {
        let path = stream.join("/");
        println!("   [{}] {}", i + 1, path);
    }
    println!();

    // Test 5: Try to read a specific stream
    println!("5. Attempting to read common Office streams:");

    // Try WordDocument stream (for .doc files)
    if ole.exists(&["WordDocument"]) {
        println!("   Found WordDocument stream");
        match ole.open_stream(&["WordDocument"]) {
            Ok(data) => {
                println!("   ✓ Successfully read WordDocument: {} bytes", data.len());
                // Show first few bytes
                let preview = &data[..data.len().min(16)];
                print!("   First bytes: ");
                for byte in preview {
                    print!("{:02X} ", byte);
                }
                println!();
            }
            Err(e) => println!("   ✗ Failed to read: {}", e),
        }
    } else {
        println!("   WordDocument stream not found (may not be a Word document)");
    }
    println!();

    // Test 6: Try to extract metadata
    println!("6. Extracting metadata:");
    match ole.get_metadata() {
        Ok(metadata) => {
            println!("   Metadata extracted:");
            if let Some(title) = metadata.title {
                println!("   - Title: {}", title);
            }
            if let Some(author) = metadata.author {
                println!("   - Author: {}", author);
            }
            if let Some(subject) = metadata.subject {
                println!("   - Subject: {}", subject);
            }
            if let Some(keywords) = metadata.keywords {
                println!("   - Keywords: {}", keywords);
            }
            if let Some(comments) = metadata.comments {
                println!("   - Comments: {}", comments);
            }
            if let Some(creating_app) = metadata.creating_application {
                println!("   - Creating Application: {}", creating_app);
            }
            if let Some(company) = metadata.company {
                println!("   - Company: {}", company);
            }
            if let Some(category) = metadata.category {
                println!("   - Category: {}", category);
            }
            if let Some(num_pages) = metadata.num_pages {
                println!("   - Number of Pages: {}", num_pages);
            }
            if let Some(num_words) = metadata.num_words {
                println!("   - Number of Words: {}", num_words);
            }
        }
        Err(e) => println!("   ✗ Failed to extract metadata: {}", e),
    }
    println!();

    println!("=== Test completed successfully! ===");

    Ok(())
}

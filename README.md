# Litchi

A high-performance Rust library for parsing Microsoft Office file formats, including legacy OLE2 formats (.doc, .xls, .ppt) and modern Office Open XML formats (.docx, .xlsx, .pptx).

## Features

### OLE2 (Legacy Office Formats)
- ✅ Parse OLE2 structured storage files
- ✅ Extract metadata and directory entries
- ✅ Read streams and storages
- ✅ Support for .doc, .xls, .ppt files

### OOXML (Modern Office Formats)
- ✅ Full Open Packaging Conventions (OPC) implementation
- ✅ Parse .docx, .xlsx, .pptx files
- ✅ Content type management
- ✅ Relationship resolution
- ✅ Efficient ZIP-based package reading
- ✅ Zero-copy XML parsing where possible

## Performance

Litchi is designed for maximum performance:

- **`memchr`** - Fast string searching using SIMD instructions
- **`atoi_simd`** - SIMD-accelerated integer parsing from byte slices
- **`fast-float2`** - Efficient floating-point number parsing
- **`quick-xml`** - Zero-copy XML parsing with minimal allocation
- **Borrowing over cloning** - Minimal memory allocations throughout

## Installation

Add this to your `Cargo.toml`:

```toml
[dependencies]
litchi = "0.0.1"
```

## Usage

### Parsing OOXML Documents

```rust
use litchi::ooxml::OpcPackage;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open a .docx file
    let package = OpcPackage::open("document.docx")?;
    
    // Get the main document part
    let main_part = package.main_document_part()?;
    println!("Content type: {}", main_part.content_type());
    println!("Size: {} bytes", main_part.blob().len());
    
    // List all parts
    for part in package.iter_parts() {
        println!("Part: {}", part.partname());
    }
    
    Ok(())
}
```

### Parsing OLE2 Documents

```rust
use litchi::ole::file::OleFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open a .doc file
    let mut ole = OleFile::open("document.doc")?;
    
    // List all streams and storages
    for entry in ole.list_dir() {
        println!("{} ({})", entry.name, entry.object_type);
    }
    
    // Read a stream
    let stream_data = ole.open_stream("WordDocument")?;
    println!("Stream size: {} bytes", stream_data.len());
    
    Ok(())
}
```

## Examples

Run the included examples:

```bash
# Parse a .docx file and display package information
cargo run --example parse_docx document.docx

# Extract XML content from a .docx file
cargo run --example extract_xml_content document.docx

# Parse a .doc file (OLE2)
cargo run --example test_ole document.doc
```

## Architecture

### OPC (Open Packaging Conventions)

The OOXML implementation follows the OPC specification and is structured around these core concepts:

- **Package**: The root container (ZIP archive)
- **Parts**: Individual content items within the package
- **Relationships**: Connections between parts
- **Content Types**: MIME types for parts

### Implementation

The codebase is organized as follows:

```
src/
├── ole/           # OLE2 format support
│   ├── file.rs    # OLE file reading
│   ├── metadata.rs # Directory and stream metadata
│   └── consts.rs  # OLE constants
└── ooxml/         # OOXML format support
    └── opc/       # Open Packaging Conventions
        ├── constants.rs   # Content types, namespaces, relationship types
        ├── packuri.rs     # Package URI handling
        ├── rel.rs         # Relationships management
        ├── part.rs        # Part implementations
        ├── phys_pkg.rs    # Physical package (ZIP) reading
        ├── pkgreader.rs   # Package reader with content type mapping
        └── package.rs     # Main OpcPackage API
```

## Design Philosophy

1. **Performance First**: Uses SIMD instructions and minimal allocations
2. **Zero-Copy**: Borrows data instead of cloning wherever possible
3. **Type Safety**: Leverages Rust's type system for correctness
4. **Standard Compliance**: Follows OPC and OLE2 specifications
5. **Ergonomic API**: Simple and intuitive interfaces

## Roadmap

- [ ] Document content extraction APIs
- [ ] Spreadsheet parsing
- [ ] Presentation parsing
- [ ] Document writing/modification
- [ ] Advanced XML element querying
- [ ] Streaming API for large files

## License

Licensed under the Apache License, Version 2.0.

## Acknowledgments

This implementation is inspired by the excellent [python-docx](https://github.com/python-openxml/python-docx) library, adapted for Rust with performance optimizations.


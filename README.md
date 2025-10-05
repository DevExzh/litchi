# Litchi

A high-performance Rust library for parsing Microsoft Office file formats, including legacy OLE2 formats (.doc, .xls, .ppt) and modern Office Open XML formats (.docx, .xlsx, .pptx).

## Features

### OLE2 (Legacy Office Formats)
- ✅ Parse OLE2 structured storage files
- ✅ Extract metadata and directory entries
- ✅ Read streams and storages
- ✅ Support for .doc, .xls, .ppt files

### DOC (Legacy Word Documents)
- ✅ Parse legacy Word documents (.doc)
- ✅ File Information Block (FIB) parsing
- ✅ Text extraction via piece table
- ✅ Support for both ANSI (Windows-1252) and Unicode (UTF-16LE) text
- ✅ Character formatting (bold, italic, underline, strikethrough, etc.)
- ✅ Font properties (size, color, highlighting)
- ✅ Text effects (superscript, subscript, small caps, all caps)
- ✅ Table structure parsing with properties
- ✅ Cell and row formatting
- ✅ Paragraph and run enumeration
- ✅ Document structure access

### OOXML (Modern Office Formats)
- ✅ Full Open Packaging Conventions (OPC) implementation
- ✅ Parse .docx, .xlsx, .pptx files
- ✅ Comprehensive Word document (.docx) API
  - **Full paragraph iteration** with run-level access
  - **Run formatting**: bold, italic, underline, font name, font size
  - **Table iteration**: rows, cells, and nested content
  - **Text extraction**: fast text content extraction
  - Document statistics (paragraph count, table count)
  - Access to document structure
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

### Parsing Legacy Word Documents (.doc)

```rust
use litchi::ole::doc::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open a .doc file
    let mut package = Package::open("document.doc")?;
    let document = package.document()?;
    
    // Extract all text
    let text = document.text()?;
    println!("Document text: {}", text);
    
    // Get document information
    let fib = document.fib();
    println!("Format version: 0x{:04X}", fib.version());
    println!("Encrypted: {}", fib.is_encrypted());
    
    // Iterate through paragraphs
    for para in document.paragraphs()? {
        println!("Paragraph: {}", para.text()?);
    }
    
    Ok(())
}
```

### Parsing Modern Word Documents (.docx)

```rust
use litchi::ooxml::docx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open a .docx file
    let package = Package::open("document.docx")?;
    let document = package.document()?;
    
    // Extract all text
    let text = document.text()?;
    println!("Document text: {}", text);
    
    // Get document statistics
    println!("Paragraphs: {}", document.paragraph_count()?);
    println!("Tables: {}", document.table_count()?);
    
    // Iterate through paragraphs and runs with formatting
    for para in document.paragraphs()? {
        println!("Paragraph: {}", para.text()?);
        
        for run in para.runs()? {
            println!("  Run: \"{}\"", run.text()?);
            println!("    Bold: {:?}", run.bold()?);
            println!("    Italic: {:?}", run.italic()?);
            println!("    Font: {:?}", run.font_name()?);
        }
    }
    
    // Iterate through tables
    for table in document.tables()? {
        println!("Table: {} rows × {} cols", 
                 table.row_count()?, table.column_count()?);
        
        for row in table.rows()? {
            for cell in row.cells()? {
                println!("Cell: {}", cell.text()?);
            }
        }
    }
    
    Ok(())
}
```

### Low-Level OOXML/OPC API

```rust
use litchi::ooxml::OpcPackage;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open any OOXML file (.docx, .xlsx, .pptx)
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
# Parse a legacy Word document (.doc)
cargo run --example parse_doc document.doc

# Comprehensive docx parsing with formatting and tables
cargo run --example docx_comprehensive document.docx

# Parse a .docx file using the high-level OOXML API
cargo run --example parse_docx_ooxml document.docx

# Parse a .docx file and display package information (low-level OPC API)
cargo run --example parse_docx document.docx

# Extract XML content from a .docx file
cargo run --example extract_xml_content document.docx

# Parse OLE2 file structure (low-level)
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
│   ├── consts.rs  # OLE constants
│   └── doc/       # Legacy Word document support
│       ├── package.rs     # DOC package wrapper
│       ├── document.rs    # Document API
│       ├── paragraph.rs   # Paragraph and Run structures
│       ├── table.rs       # Table structures
│       └── parts/         # Binary structure parsers
│           ├── fib.rs     # File Information Block
│           ├── text.rs    # Text extraction
│           ├── chp.rs     # Character Properties
│           └── tap.rs     # Table Properties
└── ooxml/         # OOXML format support
    ├── shared.rs  # Shared utilities (Length, RGBColor)
    ├── error.rs   # Error types
    ├── opc/       # Open Packaging Conventions (low-level)
    │   ├── constants.rs   # Content types, namespaces, relationship types
    │   ├── packuri.rs     # Package URI handling
    │   ├── rel.rs         # Relationships management
    │   ├── part.rs        # Part implementations
    │   ├── phys_pkg.rs    # Physical package (ZIP) reading
    │   ├── pkgreader.rs   # Package reader with content type mapping
    │   └── package.rs     # Main OpcPackage API
    ├── docx/      # Modern Word document support
    │   ├── package.rs     # Word package wrapper
    │   ├── document.rs    # Document API
    │   └── parts/         # Document-specific parts
    ├── xlsx/      # Excel spreadsheet support (placeholder)
    └── pptx/      # PowerPoint presentation support (placeholder)
```

## Design Philosophy

1. **Performance First**: Uses SIMD instructions and minimal allocations
2. **Zero-Copy**: Borrows data instead of cloning wherever possible
3. **Type Safety**: Leverages Rust's type system for correctness
4. **Standard Compliance**: Follows OPC and OLE2 specifications
5. **Ergonomic API**: Simple and intuitive interfaces

## Roadmap

### Completed
- [x] OPC (Open Packaging Conventions) implementation
- [x] Full Word document (.docx) reading API
  - [x] Paragraph iteration
  - [x] Run iteration with formatting (bold, italic, underline)
  - [x] Font properties (name, size)
  - [x] Table, row, and cell iteration
  - [x] Text extraction from paragraphs and cells
- [x] Document statistics (paragraphs, tables)
- [x] Legacy Word document (.doc) reading
  - [x] File Information Block (FIB) parsing
  - [x] Text extraction via piece table
  - [x] ANSI and Unicode text support
  - [x] Character formatting (bold, italic, underline, etc.)
  - [x] Font properties and text effects
  - [x] Table structure with properties
  - [x] Cell and row formatting
  - [x] Paragraph and run enumeration

### Planned
- [ ] Additional DOC features
  - [ ] Paragraph formatting (alignment, indentation, spacing)
  - [ ] Style definitions and style hierarchy
  - [ ] Embedded objects and images
  - [ ] Headers and footers parsing
- [ ] Additional DOCX formatting
  - [ ] Paragraph alignment and indentation
  - [ ] Styles and style hierarchy
  - [ ] Text color and highlighting
  - [ ] More underline styles (double, wavy, etc.)
- [ ] Excel spreadsheet (.xlsx) parsing
- [ ] Legacy Excel (.xls) parsing
- [ ] PowerPoint presentation (.pptx) parsing
- [ ] Legacy PowerPoint (.ppt) parsing
- [ ] Document writing/modification
- [ ] Advanced XML element querying
- [ ] Streaming API for large files

## License

Licensed under the Apache License, Version 2.0.

## Documentation

- **[OLE Implementation Guide](OLE_IMPLEMENTATION.md)** - Details on OLE2 file format support
- **[DOC Implementation Guide](DOC_IMPLEMENTATION.md)** - Details on legacy Word document support
- **[OOXML Implementation Guide](OOXML_IMPLEMENTATION.md)** - Details on OOXML format support
- **[OOXML API Guide](OOXML_API_GUIDE.md)** - Comprehensive API documentation for OOXML
- **[DOCX Reading Features](DOCX_READING_FEATURES.md)** - Guide to reading DOCX files

## Acknowledgments

This implementation is inspired by:
- [python-docx](https://github.com/python-openxml/python-docx) - Python library for DOCX files
- [Apache POI](https://poi.apache.org/) - Java library for Microsoft Office formats
- Microsoft's official file format specifications


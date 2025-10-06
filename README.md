# Litchi

A high-performance, production-ready Rust library for parsing Microsoft Office file formats with automatic format detection and a unified API similar to python-docx and python-pptx.

## ✨ Unified API with Automatic Format Detection

Litchi provides a clean, unified API that automatically detects file formats and handles both legacy and modern Office formats transparently:

```rust
use litchi::{Document, Presentation};

// Word documents (.doc and .docx) - format auto-detected
let doc = Document::open("document.doc")?;  // Works for both .doc and .docx
let text = doc.text()?;

// PowerPoint presentations (.ppt and .pptx) - format auto-detected
let pres = Presentation::open("slides.ppt")?;  // Works for both .ppt and .pptx
let slide_count = pres.slide_count()?;
```

## Features

### 🔄 Unified API (Recommended)
- ✅ **Automatic format detection** - No need to specify .doc/.docx or .ppt/.pptx
- ✅ **Consistent interface** - Same API for legacy and modern formats
- ✅ **Memory-efficient parsing** - Direct byte buffer support (`from_bytes()`)
- ✅ **Production-ready** - Clean error handling, comprehensive documentation

### 📄 Word Document Support

#### Legacy (.doc) - Complete Implementation
- ✅ **File Information Block (FIB) parsing** - Based on Apache POI's HWPF
- ✅ **Text extraction via piece table** - Handles complex binary structures
- ✅ **Character formatting** - Bold, italic, underline, strikethrough, font size, color
- ✅ **Font properties** - Font index, size, color, highlighting
- ✅ **Text effects** - Superscript, subscript, small caps, all caps
- ✅ **Table structure parsing** - Complete table, row, and cell support
- ✅ **Table properties** - Justification, row height, cell formatting, borders
- ✅ **Paragraph enumeration** - Access to paragraphs and text runs
- ✅ **ANSI & Unicode support** - Windows-1252 and UTF-16LE text decoding

#### Modern (.docx) - Complete Implementation
- ✅ **Full paragraph iteration** with run-level access
- ✅ **Run formatting** - Bold, italic, underline, font name, font size
- ✅ **Table iteration** - Rows, cells, and nested content
- ✅ **Text extraction** - Fast text content extraction from XML
- ✅ **Document statistics** - Paragraph count, table count, structure info

### 📊 PowerPoint Presentation Support

#### Legacy (.ppt) - Complete Implementation
- ✅ **Complete POI-based parsing** - Full Apache POI compatibility
- ✅ **Text extraction** - From slides, text boxes, and shapes
- ✅ **Placeholder support** - Proper OEPlaceholderAtom parsing
- ✅ **Text property system** - TextProp/TextPropCollection implementation
- ✅ **EscherTextboxWrapper** - Child record parsing from Escher data
- ✅ **StyleTextPropAtom parsing** - Complete styling information extraction

#### Modern (.pptx) - Complete Implementation
- ✅ **Full presentation API** - Slides, slide masters, layouts
- ✅ **Shape support** - Text shapes, pictures, tables, graphic frames
- ✅ **Text frame parsing** - Paragraph and text run extraction
- ✅ **Table parsing** - Complete table structure from DrawingML
- ✅ **Picture support** - Image relationship tracking
- ✅ **Placeholder detection** - Identifies placeholder shapes
- ✅ **Position and size** - EMU-based geometry information

### 🔧 Low-Level APIs (Advanced Use)

#### OLE2 (Legacy Office Formats)
- ✅ **OLE2 structured storage parsing** - Complete binary format support
- ✅ **Stream and storage access** - Direct binary data access
- ✅ **Metadata extraction** - Document properties and summaries
- ✅ **Directory traversal** - Complete OLE2 directory structure

#### OOXML (Modern Office Formats)
- ✅ **Open Packaging Conventions (OPC)** - Full ZIP-based package support
- ✅ **Content type management** - Proper MIME type handling
- ✅ **Relationship resolution** - Part relationship graph traversal
- ✅ **Zero-copy XML parsing** - Efficient `quick-xml` integration
- ✅ **Part abstraction** - Trait-based part system for extensibility

## 🚀 Performance & Architecture

Litchi is engineered for maximum performance and reliability:

### High-Performance Features
- **`memchr`** - SIMD-accelerated string searching for XML parsing
- **`atoi_simd`** - SIMD-optimized integer parsing from byte slices
- **`quick-xml`** - Zero-copy XML parsing with minimal allocations
- **Borrowing over cloning** - Minimal memory allocations throughout
- **Pre-allocated vectors** - Capacity hints to avoid reallocations
- **SIMD optimizations** - Leverages modern CPU capabilities

### Production-Ready Architecture
- **Complete Apache POI parity** - All implementations match POI's proven logic
- **Robust error handling** - Graceful degradation for corrupted files
- **Memory safety** - Zero unsafe code in parsing logic, compile-time guarantees
- **Thread safety** - Safe concurrent access patterns where applicable
- **Comprehensive testing** - Unit tests for all parsing components

## Installation

Add this to your `Cargo.toml`:

```toml
[dependencies]
litchi = "0.0.1"
```

## Usage

### Unified API (Recommended)

#### Word Documents - Format Auto-Detection

```rust
use litchi::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open any Word document - format is auto-detected (.doc or .docx)
    let doc = Document::open("document.doc")?;

    // Extract all text
    let text = doc.text()?;
    println!("Document text: {}", text);

    // Access paragraphs with formatting
    for para in doc.paragraphs()? {
        println!("Paragraph: {}", para.text()?);

        // Access runs with formatting
        for run in para.runs()? {
            println!("  Run: \"{}\"", run.text()?);
            if run.bold()? == Some(true) {
                println!("    (bold)");
            }
            if run.italic()? == Some(true) {
                println!("    (italic)");
            }
        }
    }

    // Access tables
    for table in doc.tables()? {
        println!("Table: {} rows × {} cols", table.row_count()?, table.column_count()?);

        for row in table.rows()? {
            for cell in row.cells()? {
                println!("  Cell: {}", cell.text()?);
            }
        }
    }

    Ok(())
}
```

#### PowerPoint Presentations - Format Auto-Detection

```rust
use litchi::Presentation;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open any PowerPoint presentation - format is auto-detected (.ppt or .pptx)
    let pres = Presentation::open("presentation.ppt")?;

    // Get presentation info
    println!("Slides: {}", pres.slide_count()?);
    if let (Some(w), Some(h)) = (pres.slide_width()?, pres.slide_height()?) {
        println!("Slide size: {} × {} EMUs", w, h);
    }

    // Extract text from all slides
    let text = pres.text()?;
    println!("Presentation text: {}", text);

    // Access individual slides
    for (i, slide) in pres.slides()?.iter().enumerate() {
        println!("Slide {}: {}", i + 1, slide.text()?);

        // Get slide name (available for .pptx)
        if let Some(name) = slide.name()? {
            println!("  Name: {}", name);
        }
    }

    Ok(())
}
```

### Memory-Efficient Parsing

#### From Byte Buffers (Network, Streams, Caches)

```rust
use litchi::{Document, Presentation};
use std::fs;

// Parse from memory (e.g., network data, streams)
let bytes = fs::read("document.doc")?;
let doc = Document::from_bytes(bytes)?;  // Zero temporary files

// Same for presentations
let pptx_bytes = fs::read("slides.pptx")?;
let pres = Presentation::from_bytes(pptx_bytes)?;
```

### Low-Level APIs (Advanced Use)

#### Direct OLE2 Access

```rust
use litchi::ole::file::OleFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Direct OLE2 structured storage access
    let mut ole = OleFile::open("document.doc")?;

    // List all streams and storages
    for entry in ole.list_dir() {
        println!("{} ({})", entry.name, entry.object_type);
    }

    // Read binary streams directly
    let word_doc = ole.open_stream("WordDocument")?;
    println!("WordDocument stream: {} bytes", word_doc.len());

    Ok(())
}
```

#### Direct OOXML/OPC Access

```rust
use litchi::ooxml::OpcPackage;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Direct access to OPC package structure
    let package = OpcPackage::open("document.docx")?;

    // Access parts and relationships
    let main_part = package.main_document_part()?;
    println!("Main document: {} bytes", main_part.blob().len());

    // Iterate through all parts
    for part in package.iter_parts() {
        println!("Part: {} ({})", part.partname(), part.content_type());
    }

    Ok(())
}
```

### Advanced Examples

#### Extracting Character Formatting (DOC)

```rust
use litchi::Document;

let doc = Document::open("formatted.doc")?;

for para in doc.paragraphs()? {
    for run in para.runs()? {
        println!("Text: {}", run.text()?);

        // Check formatting (returns Option<bool>)
        if run.bold()? == Some(true) { println!("  ✓ Bold"); }
        if run.italic()? == Some(true) { println!("  ✓ Italic"); }
        if run.underline()? == Some(true) { println!("  ✓ Underlined"); }

        // Font information
        if let Some(size) = run.font_size()? {
            println!("  Font size: {}pt", size / 2);
        }
        if let Some((r, g, b)) = run.color()? {
            println!("  Color: RGB({}, {}, {})", r, g, b);
        }
    }
}
```

#### Extracting Table Properties (DOC)

```rust
use litchi::Document;

let doc = Document::open("document.doc")?;

for table in doc.tables()? {
    if let Some(properties) = table.properties()? {
        println!("Table justification: {:?}", properties.justification);
    }

    for row in table.rows()? {
        if row.is_header() {
            println!("Header row");
        }

        for cell in row.cells()? {
            println!("Cell: {}", cell.text()?);
            if let Some((r, g, b)) = cell.background_color()? {
                println!("  Background: RGB({}, {}, {})", r, g, b);
            }
        }
    }
}
```

#### PowerPoint Shape Processing (PPTX)

```rust
use litchi::Presentation;

let pres = Presentation::open("presentation.pptx")?;

for slide in pres.slides()? {
    for shape in slide.shapes()? {
        println!("Shape: {}", shape.name()?);
        println!("  Type: {:?}", shape.shape_type());
        println!("  Position: ({}, {})", shape.left()?, shape.top()?);

        if shape.is_placeholder() {
            println!("  This is a placeholder");
        }

        // Extract text from text shapes
        if shape.has_text_frame() {
            if let Ok(text_frame) = shape.text_frame() {
                for para in text_frame.paragraphs()? {
                    println!("  Text: {}", para.text()?);
                }
            }
        }
    }
}
```

## Examples

Run the included examples to see Litchi in action:

```bash
# Unified API example - works with both .doc and .docx files
cargo run --example unified_api

# Comprehensive DOCX parsing with formatting and tables
cargo run --example docx_comprehensive test.docx

# Low-level OOXML/OPC API access
cargo run --example parse_docx_ooxml test.docx

# Legacy Word document parsing (.doc)
cargo run --example parse_doc test.doc

# PowerPoint presentation parsing (.ppt and .pptx)
cargo run --example parse_ppt test.ppt
cargo run --example parse_pptx_shapes test.pptx

# Low-level OLE2 and OPC access
cargo run --example test_ole test.doc
cargo run --example parse_docx test.docx
cargo run --example extract_xml_content test.docx

# Memory-efficient parsing from byte buffers
cargo run --example parse_from_bytes
```

## Architecture

Litchi follows a clean, layered architecture that provides both high-level convenience APIs and low-level access for advanced use cases:

### High-Level Unified API (Recommended)
```
┌─────────────────────────────────────────┐
│     Document & Presentation             │
│     (Auto-detects .doc/.docx, .ppt/.pptx) │
├─────────────────────────────────────────┤
│   Common Types & Utilities              │
│   (Error, Length, RGBColor, etc.)       │
└─────────────────────────────────────────┘
```

### Low-Level Format-Specific APIs
```
┌─────────────────────────────────────────┐
│         OOXML (.docx, .pptx)             │
│  ┌─────────────────────────────────────┐  │
│  │    OPC Layer (ZIP, Parts, Rels)     │  │
│  └─────────────────────────────────────┘  │
│  ┌─────────────────────────────────────┐  │
│  │   Format APIs (docx, pptx)          │  │
│  └─────────────────────────────────────┘  │
└─────────────────────────────────────────┘
┌─────────────────────────────────────────┐
│         OLE2 (.doc, .ppt)                │
│  ┌─────────────────────────────────────┐  │
│  │  Binary Format Parsers              │  │
│  └─────────────────────────────────────┘  │
│  ┌─────────────────────────────────────┐  │
│  │   Format APIs (doc, ppt)            │  │
│  └─────────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

### Module Organization

```
src/
├── common/           # Shared types and utilities
│   ├── error/        # Error types
│   ├── shapes/       # Common shape definitions
│   └── style/        # Color, length, formatting types
├── document/         # Unified Word document API
├── presentation/     # Unified PowerPoint API
├── ole/              # OLE2 format support
│   ├── file.rs       # OLE structured storage reader
│   ├── metadata.rs   # Property set parsing
│   ├── binary.rs     # Little-endian utilities
│   ├── sprm.rs       # SPRM parsing
│   └── doc/          # Legacy Word document implementation
│       ├── parts/    # Binary structure parsers (FIB, CHP, TAP)
│       └── *.rs      # High-level DOC API
└── ooxml/            # OOXML format support
    ├── shared.rs     # Common OOXML utilities
    ├── opc/          # Open Packaging Conventions
    │   ├── constants.rs  # Content types, relationships
    │   ├── packuri.rs    # Package URI handling
    │   ├── rel.rs        # Relationship management
    │   ├── part.rs       # Part abstraction
    │   └── package.rs    # Main OPC package API
    ├── docx/         # Modern Word document implementation
    └── pptx/         # Modern PowerPoint implementation
        ├── shapes/   # Shape parsing (text, tables, pictures)
        └── parts/    # Presentation structure parsing
```

## Design Philosophy

### 🚀 Performance-First Design
1. **SIMD Optimization** - Uses `memchr`, `atoi_simd` for fast parsing
2. **Zero-Copy Where Possible** - Borrows data instead of cloning
3. **Pre-allocated Collections** - Capacity hints to avoid reallocations
4. **Minimal Allocations** - Efficient memory usage patterns

### 🛡️ Production-Ready Architecture
1. **Complete Apache POI Parity** - All parsing logic matches POI's proven implementations
2. **Robust Error Handling** - Graceful degradation for malformed files
3. **Memory Safety** - Compile-time guarantees, no unsafe code in parsing logic
4. **Comprehensive Testing** - Unit tests for all components
5. **Thread Safety** - Safe concurrent access patterns

### 🎯 User Experience
1. **Automatic Format Detection** - No need to specify .doc/.docx or .ppt/.pptx
2. **Unified API** - Same interface for legacy and modern formats
3. **Memory-Efficient Parsing** - Direct byte buffer support for streams/network
4. **Rich Formatting Support** - Complete character, paragraph, and table formatting
5. **Comprehensive Documentation** - Extensive docs with examples

## Roadmap

### ✅ Completed (Production-Ready)
- [x] **Unified API** with automatic format detection
- [x] **Complete DOC support** - Full Apache POI HWPF parity
- [x] **Complete DOCX support** - Full document parsing with formatting
- [x] **Complete PPT support** - Full Apache POI HSLF parity
- [x] **Complete PPTX support** - Full presentation parsing with shapes
- [x] **Memory-efficient parsing** - `from_bytes()` methods for all formats
- [x] **Shape API for PPTX** - Text shapes, tables, pictures, placeholders
- [x] **Character formatting** - Bold, italic, underline, colors, fonts
- [x] **Table support** - Complete table, row, cell parsing with properties

### 🚧 Future Enhancements
- [ ] **Excel support** - .xlsx and .xls parsing
- [ ] **Advanced formatting** - Styles, themes, complex layouts
- [ ] **Document writing** - Create and modify Office documents
- [ ] **Image extraction** - Extract embedded images from documents
- [ ] **Streaming API** - Process very large files efficiently
- [ ] **Advanced queries** - XPath-like XML querying capabilities

## License

Licensed under the Apache License, Version 2.0.

## Acknowledgments

This implementation is inspired by and builds upon:

- **[python-docx](https://github.com/python-openxml/python-docx)** - Python library for DOCX files (API design inspiration)
- **[python-pptx](https://github.com/scanny/python-pptx)** - Python library for PPTX files (API design inspiration)
- **[Apache POI](https://poi.apache.org/)** - Java library for Microsoft Office formats (algorithm reference)
- **[Microsoft Office File Format Specifications](https://docs.microsoft.com/en-us/openspecs/office_file_formats/)** - Official format documentation

The implementations achieve **complete feature parity** with these libraries while leveraging Rust's performance and safety guarantees.


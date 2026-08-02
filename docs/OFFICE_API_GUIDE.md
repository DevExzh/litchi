# Office Document CRUD API Guide

This guide provides comprehensive documentation for creating, reading, and updating Office documents using Litchi.

## Table of Contents

- [Overview](#overview)
- [Word Documents (DOCX)](#word-documents-docx)
- [Excel Spreadsheets (XLSX)](#excel-spreadsheets-xlsx)
- [PowerPoint Presentations (PPTX)](#powerpoint-presentations-pptx)
- [Unified API](#unified-api)
- [Performance Tips](#performance-tips)
- [Error Handling](#error-handling)

## Overview

Litchi provides a comprehensive, high-performance API for working with Microsoft Office file formats. The API follows these principles:

- **Idiomatic Rust**: Using standard Rust patterns and conventions
- **Zero-copy where possible**: Minimizing memory allocations
- **Type-safe**: Leveraging Rust's type system for correctness
- **Production-ready**: Fast, safe, and well-tested

## Word Documents (DOCX)

### Creating Documents

```rust
use litchi::ooxml::docx::Package;

// Create a new document
let mut pkg = Package::new()?;
let mut doc = pkg.document_mut()?;

// Add headings
doc.add_heading("Document Title", 1)?;
doc.add_heading("Section 1", 2)?;

// Add paragraphs
doc.add_paragraph_with_text("This is a simple paragraph.");

// Add formatted text
let para = doc.add_paragraph();
para.add_run_with_text("Bold text ").bold(true);
para.add_run_with_text("and italic text").italic(true);

// Add tables
let table = doc.add_table(3, 4); // 3 rows, 4 columns
if let Some(cell) = table.cell(0, 0) {
    cell.set_text("Header 1");
}
if let Some(cell) = table.cell(0, 1) {
    cell.set_text("Header 2");
}

// Set metadata
pkg.properties_mut().title = Some("My Document".to_string());
pkg.properties_mut().creator = Some("Your Name".to_string());

// Save
pkg.save("output.docx")?;
```

### Reading Documents

```rust
use litchi::ooxml::docx::Package;

// Open document
let pkg = Package::open("document.docx")?;
let doc = pkg.document()?;

// Get statistics
println!("Paragraphs: {}", doc.paragraph_count()?);
println!("Tables: {}", doc.table_count()?);

// Extract all text
let text = doc.text()?;
println!("Text: {}", text);

// Iterate paragraphs
for para in doc.paragraphs()? {
    println!("Paragraph: {}", para.text()?);
    
    // Access runs
    for run in para.runs()? {
        println!("  Text: {}", run.text()?);
        println!("  Bold: {:?}", run.bold()?);
        println!("  Italic: {:?}", run.italic()?);
    }
}

// Access tables
for table in doc.tables()? {
    for row in table.rows()? {
        for cell in row.cells()? {
            print!("{}\t", cell.text()?);
        }
        println!();
    }
}

// Search
let matches = doc.search("important")?;
println!("Found in {} paragraphs", matches.len());
```

### Updating Documents

```rust
use litchi::ooxml::docx::Package;

// Open existing document
let mut pkg = Package::open("document.docx")?;
let mut doc = pkg.document_mut()?;

// Add new content
doc.add_heading("New Section", 2)?;
doc.add_paragraph_with_text("Additional content...");

// Update metadata
pkg.properties_mut().last_modified_by = Some("Editor Name".to_string());

// Save (can overwrite or save to new file)
pkg.save("updated.docx")?;
```

### Custom Properties and XML Data

Custom document properties use the shared, bounded `custom::{Props, Value}`
vocabulary. Names follow Office's case-insensitive identity rules, while the
original spelling is preserved. Insert is fallible so invalid names, values,
or resource limits cannot be bypassed.

```rust
use litchi::ooxml::custom::Value;
use litchi::ooxml::docx::Package;

let mut pkg = Package::open("document.docx")?;
pkg.custom_props_mut().insert("Project", "Litchi")?;
pkg.custom_props_mut().insert("Version", Value::I32(38))?;

if let Some(value) = pkg.custom_props().get("PROJECT") {
    println!("Project: {value:?}");
}

pkg.custom_props_mut().remove("Version");
pkg.save("updated.docx")?;
```

Custom XML payloads stay inert: Litchi validates their bounded XML and package
graph but never fetches schemas or executes XPath. Loaded payloads borrow shared
immutable OPC storage internally, so repeated relationships do not duplicate
the XML allocation.

```rust
use litchi::ooxml::custom_xml::Conformance;
use litchi::ooxml::docx::{NewStore, Package};

let mut pkg = Package::open("document.docx")?;
let item = pkg.add_custom_xml(NewStore {
    xml: br#"<customer xmlns="urn:customer" id="7"/>"#.to_vec(),
    content_type: "application/xml".to_string(),
    id: "{11111111-1111-4111-8111-111111111111}".to_string(),
    schemas: vec!["urn:customer".to_string()],
    conformance: Conformance::Transitional,
})?;

println!("Stored {} bytes in {}", item.xml().len(), item.part());
pkg.set_custom_xml(
    "{11111111-1111-4111-8111-111111111111}",
    br#"<customer xmlns="urn:customer" id="8"/>"#.to_vec(),
)?;
pkg.save("updated.docx")?;
```

### Advanced Document Operations

```rust
// Get specific paragraph
if let Some(para) = doc.paragraph(0)? {
    println!("First paragraph: {}", para.text()?);
}

// Get text range
let text = doc.text_range(5, 10)?; // Paragraphs 5-10

// Check for tables
if doc.has_tables()? {
    println!("Document contains tables");
}

// Case-insensitive search
let matches = doc.search_ignore_case("IMPORTANT")?;

// Access sections
let sections = doc.sections()?;
for section in sections.iter() {
    if let Some(width) = section.page_width() {
        println!("Page width: {} inches", width.to_inches());
    }
}

// Access styles
let styles = doc.styles()?;
if let Some(style) = styles.get_by_name("Heading 1")? {
    println!("Found style: {}", style.style_id());
}
```

### Document Inventories (Read-Only)

SmartArt diagrams and DrawingML text boxes are exposed as typed, inert
inventories in both transitional and Strict namespace dialects:

```rust
use litchi::ooxml::docx::Package;

let pkg = Package::open("document.docx")?;

// SmartArt (DrawingML diagram) inventory: data-model node trees plus
// layout/quick-style/colors part metadata.
for art in pkg.smart_arts()? {
    println!("Diagram: {:?} ({} parts)", art.diagram_type(), art.text().lines().count());
    println!("{}", art.text());
}

// Text boxes and WordArt anchored in the main document, including VML
// fallbacks collapsed from mc:AlternateContent.
for text_box in pkg.text_boxes()? {
    let name = text_box.name.as_deref().unwrap_or("<unnamed>");
    println!("Shape '{name}' (word art: {})", text_box.is_word_art());
    println!("{}", text_box.text());
}
```

### Authoring Rich Content

Text boxes, SmartArt diagrams, embedded OLE objects, and legacy VML shapes
can be authored through the mutable document and round-trip through the
read inventories above:

```rust
use litchi::ooxml::docx::{MutableTextBox, MutableOleObject, MutableSmartArt, MutableVmlShape, VmlShapeKind};
use litchi::drawing::diagram::{DiagramType, SmartArtBuilder};
use litchi::ooxml::docx::Package;

let mut pkg = Package::new()?;
let mut doc = pkg.document_mut()?;

// DrawingML text box with formatted runs and body properties.
let mut text_box = MutableTextBox::new("Intro Box", 2_743_200, 914_400)?; // 3" x 1"
text_box.add_run("Bold heading").bold = true;
text_box.add_paragraph_with_text("Second paragraph");
doc.add_text_box(text_box);

// SmartArt: definition parts and the dgm:relIds anchor are generated.
let diagram = SmartArtBuilder::new(DiagramType::Process)
    .add_items(["Plan", "Build", "Ship"])
    .build();
doc.add_smart_art(MutableSmartArt::new(diagram)?);

// Inert embedded OLE payload with a validated ProgID.
let ole = MutableOleObject::new("Package", vec![0x50, 0x4B, 0x03, 0x04])?;
doc.add_ole_object(ole);

// Legacy VML shape with fill and stroke colors.
let mut shape = MutableVmlShape::new(VmlShapeKind::RoundRectangle, 1_828_800, 914_400)?;
shape.set_fill_color("FFDDEE")?;
shape.set_stroke_color("336699")?;
doc.add_vml_shape(shape);

pkg.save("output.docx")?;
```

## Excel Spreadsheets (XLSX)

### Creating Workbooks

```rust
use litchi::ooxml::xlsx::Workbook;

// Create new workbook
let mut wb = Workbook::create()?;

// Access first worksheet
let mut ws = wb.worksheet_mut(0)?;

// Set cell values
ws.set_cell_value(1, 1, "Name");
ws.set_cell_value(1, 2, "Age");
ws.set_cell_value(2, 1, "Alice");
ws.set_cell_value(2, 2, 30);
ws.set_cell_value(3, 1, "Bob");
ws.set_cell_value(3, 2, 25);

// Set formulas
ws.set_cell_formula(4, 2, "=AVERAGE(B2:B3)");

// Add more worksheets
let mut ws2 = wb.add_worksheet("Summary");
ws2.set_cell_value(1, 1, "Total");

// Define named ranges
wb.define_name("DataRange", "Sheet1!$A$1:$B$3");

// Set freeze panes
ws.freeze_panes(2, 1)?; // Freeze first row

// Set metadata
wb.properties_mut().title = Some("My Workbook".to_string());

// Save
wb.save("output.xlsx")?;
```

### Reading Workbooks

```rust
use litchi::ooxml::xlsx::Workbook;

// Open workbook
let wb = Workbook::open("workbook.xlsx")?;

// Get worksheet info
println!("Worksheets: {}", wb.worksheet_count());
for name in wb.worksheet_names() {
    println!("  - {}", name);
}

// Access worksheet
let ws = wb.worksheet_by_name("Sheet1")?;
println!("Sheet: {}", ws.name());

// Get dimensions
if let Some((min_row, min_col, max_row, max_col)) = ws.used_range() {
    println!("Range: {}x{}", max_row - min_row + 1, max_col - min_col + 1);
}

// Access specific cell
let cell = ws.cell(1, 1)?;
if let Some(value) = cell.value() {
    println!("A1: {:?}", value);
}

// Iterate cells
for row in ws.rows()? {
    for cell in row.cells()? {
        if let Some(value) = cell.value() {
            print!("{:?}\t", value);
        }
    }
    println!();
}

// Search
let matches = ws.find_text("Total")?;
println!("Found in {} cells", matches.len());
```

### Updating Workbooks

```rust
use litchi::ooxml::xlsx::Workbook;

// Open existing workbook
let mut wb = Workbook::open("workbook.xlsx")?;

// Update cells
let mut ws = wb.worksheet_mut(0)?;
ws.set_cell_value(4, 1, "Charlie");
ws.set_cell_value(4, 2, 35);

// Add new worksheet
let mut new_ws = wb.add_worksheet("Q2 Data");
new_ws.set_cell_value(1, 1, "Revenue");

// Save
wb.save("updated.xlsx")?;
```

### Advanced Worksheet Operations

```rust
// Get column values
let column_a = ws.column_values(1)?;
for value in column_a {
    println!("{:?}", value);
}

// Get row values
let row_1 = ws.row_values(1)?;

// Get range as 2D array
let range = ws.range(1, 1, 5, 3)?; // A1:C5
for row in range {
    for cell in row {
        print!("{:?}\t", cell);
    }
    println!();
}

// Check if cell is empty
if ws.is_cell_empty(1, 1) {
    println!("Cell A1 is empty");
}

// Count non-empty cells
println!("Non-empty cells: {}", ws.non_empty_cell_count());
```

### Pivot Inventories (Read-Only)

Pivot charts bound to a worksheet's pivot tables are exposed as a typed,
validated graph; XLSB workbooks additionally expose their PivotCache
definition streams:

```rust
use litchi::ooxml::xlsx::Workbook;

let wb = Workbook::open("report.xlsx")?;

// Pivot charts: chart part, c:pivotSource metadata, per-series drop-zone
// visibility, and the resolved typed pivot table.
for chart in wb.pivot_charts_on_sheet("Sheet1")? {
    println!("Pivot chart '{}' -> table '{}'",
        chart.part_name, chart.pivot_table.name);
}

// XLSB PivotCache definitions (via litchi::ooxml::xlsb::Workbook):
// refresh metadata, sources, cache fields, shared items, grouping,
// OLAP hierarchies, and calculated members.
```

## PowerPoint Presentations (PPTX)

### Creating Presentations

```rust
use litchi::ooxml::pptx::Package;

// Create new presentation
let mut pkg = Package::new()?;
let mut pres = pkg.presentation_mut()?;

// Add slides
let slide1 = pres.add_slide()?;
slide1.set_title("Welcome");
slide1.add_text_box(
    "Presentation Title",
    914400,   // x: 1 inch (914400 EMUs)
    2743200,  // y: 3 inches
    7315200,  // width: 8 inches
    914400,   // height: 1 inch
);

let slide2 = pres.add_slide()?;
slide2.set_title("Agenda");
slide2.add_bullet_points(&[
    "Introduction",
    "Main Content",
    "Conclusion",
])?;

// Add image
let image_data = std::fs::read("logo.png")?;
slide2.add_image(&image_data, 914400, 914400, 1828800, 1828800)?;

// Set metadata
pkg.properties_mut().title = Some("My Presentation".to_string());

// Save
pkg.save("output.pptx")?;
```

### Reading Presentations

```rust
use litchi::ooxml::pptx::Package;

// Open presentation
let pkg = Package::open("presentation.pptx")?;
let pres = pkg.presentation()?;

// Get info
println!("Slides: {}", pres.slide_count()?);
if let Some(width) = pres.slide_width()? {
    println!("Width: {} inches", width as f64 / 914400.0);
}

// Iterate slides
for (idx, slide) in pres.slides()?.iter().enumerate() {
    println!("\nSlide {}:", idx + 1);
    println!("  Shapes: {}", slide.shape_count()?);
    println!("  Text: {}", slide.text()?);
    
    // Access shapes
    for shape in slide.shapes()? {
        if let Some(text) = shape.text()? {
            println!("  Shape text: {}", text);
        }
    }
}
```

### Updating Presentations

```rust
use litchi::ooxml::pptx::Package;

// Open existing presentation
let mut pkg = Package::open("presentation.pptx")?;
let mut pres = pkg.presentation_mut()?;

// Add new slide
let new_slide = pres.add_slide()?;
new_slide.set_title("Conclusion");
new_slide.add_text_box("Thank you!", 914400, 3657600, 7315200, 914400);

// Save
pkg.save("updated.pptx")?;
```

### Advanced Slide Operations

```rust
// Get specific slide shape
if let Some(shape) = slide.shape(0)? {
    println!("First shape: {:?}", shape);
}

// Check for tables
if slide.has_tables()? {
    println!("Slide contains tables");
}

// Check for pictures
if slide.has_pictures()? {
    println!("Slide contains pictures");
}

// Get text shapes only
for shape in slide.text_shapes()? {
    if let Some(text) = shape.text()? {
        println!("Text: {}", text);
    }
}

// Search in slide
let matches = slide.find_text("important")?;
println!("Found in {} shapes", matches.len());

// Check if empty
if slide.is_empty()? {
    println!("Slide is empty");
}
```

### Masters, Themes, and Package-Level Authoring

Slide masters, layouts, themes, embedded OLE objects, and ink annotations
are authored at the package level and re-validated against the read-side
graph rules:

```rust
use litchi::ooxml::pptx::{Package, SlideLayoutKind, ThemeColorScheme, ThemeColorSlot, ThemeColorValue, ThemeFontScheme, ThemeFontFace};

let mut pkg = Package::new()?;

// A new slide master with default text styles, plus a typed layout.
let master = pkg.add_slide_master()?;
let _layout = pkg.add_slide_layout(&master.part_name, SlideLayoutKind::TitleOnly, "Custom", &[])?;

// A theme with a typed 12-slot color scheme and major/minor fonts.
let colors = ThemeColorScheme::new("Brand")
    .with_color(ThemeColorSlot::Accent1, ThemeColorValue::Srgb("4472C4".to_string()));
let fonts = ThemeFontScheme::new("Brand", ThemeFontFace::new("Calibri Light"), ThemeFontFace::new("Calibri"));
let theme = pkg.add_theme("Brand", &colors, &fonts)?;
pkg.attach_theme_to_master(&master.part_name, &theme.part_name)?;

// Inert ink annotation (validated InkML stored verbatim).
// PackURI comes from the litchi-opc crate.
let inkml = br#"<ink xmlns="http://www.w3.org/2003/InkML"><traceGroup/></ink>"#;
let slide = litchi_opc::PackURI::new("/ppt/slides/slide1.xml")?;
pkg.add_ink_annotation(&slide, inkml)?;

pkg.save("branded.pptx")?;
```

## Unified API

For simpler operations, use the unified helper API:

```rust
use litchi::ooxml::api::helpers;

// Extract text from any Office format
let text = helpers::extract_text("document.docx")?;
println!("{}", text);

// Get metadata from any format
let props = helpers::get_properties("document.docx")?;
if let Some(title) = props.title {
    println!("Title: {}", title);
}
```

## Performance Tips

1. **Lazy Loading**: Content is loaded on-demand. Access only what you need.

2. **Zero-Copy**: Use references where possible:
```rust
// Good: Borrows
let text = doc.text()?;

// Avoid unnecessary clones
let paragraphs = doc.paragraphs()?; // Already returns Vec
```

3. **Batch Operations**: Group multiple writes together:
```rust
let mut ws = wb.worksheet_mut(0)?;
for i in 1..=1000 {
    ws.set_cell_value(i, 1, format!("Row {}", i));
}
wb.save("output.xlsx")?; // Single save at end
```

4. **Iterator Patterns**: Use iterators for memory efficiency:
```rust
// Good: Iterates without collecting all
for para in doc.paragraphs()? {
    process_paragraph(para)?;
}
```

5. **SIMD Acceleration**: The library uses SIMD for string operations automatically.

## Error Handling

All operations return `Result` types with descriptive errors:

```rust
use litchi::ooxml::error::OoxmlError;

match Package::open("document.docx") {
    Ok(pkg) => {
        // Process document
    }
    Err(OoxmlError::IoError(e)) => {
        eprintln!("IO error: {}", e);
    }
    Err(OoxmlError::InvalidFormat(msg)) => {
        eprintln!("Invalid format: {}", msg);
    }
    Err(OoxmlError::PartNotFound(part)) => {
        eprintln!("Missing part: {}", part);
    }
    Err(e) => {
        eprintln!("Error: {}", e);
    }
}
```

## Examples

See the `examples/` directory for complete working examples:

- `office_crud_demo.rs` - Full CRUD operations demo
- `read_office_files.rs` - Reading and analysis examples

Run examples with:
```bash
cargo run --example office_crud_demo
cargo run --example read_office_files
```

## API Reference

For detailed API documentation, run:
```bash
cargo doc --open
```

## Thread Safety

- Package types are `Send` but not `Sync`
- For concurrent access, use separate Package instances per thread
- Read-only operations can share references appropriately

## License

See the main project LICENSE file for details.

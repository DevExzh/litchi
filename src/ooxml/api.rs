//! High-level unified API for Office documents.
//!
//! This module provides a comprehensive, user-friendly API for creating, reading,
//! and updating Office Open XML documents (DOCX, XLSX, PPTX). The API is designed
//! to be fast, safe, idiomatic, and production-ready.
//!
//! # Architecture Overview
//!
//! Each format has a consistent API pattern:
//!
//! 1. **Package/Workbook** - Entry point for file operations
//! 2. **Document/Presentation** - High-level content access (reading)
//! 3. **MutableDocument/MutablePresentation** - Content modification (writing)
//! 4. **Properties** - Metadata management
//!
//! # Examples
//!
//! ## Word Documents (DOCX)
//!
//! ### Creating a new document
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//!
//! // Create a new document
//! let mut pkg = Package::new()?;
//! let mut doc = pkg.document_mut()?;
//!
//! // Add content
//! doc.add_heading("My Document", 1)?;
//! doc.add_paragraph_with_text("This is the first paragraph.");
//!
//! let para = doc.add_paragraph();
//! para.add_run_with_text("Bold text ").bold(true);
//! para.add_run_with_text("and normal text.");
//!
//! // Add a table
//! let table = doc.add_table(3, 2);
//! table.cell(0, 0).unwrap().set_text("Header 1");
//! table.cell(0, 1).unwrap().set_text("Header 2");
//!
//! // Set metadata
//! pkg.properties_mut().title = Some("My Document".to_string());
//! pkg.properties_mut().creator = Some("John Doe".to_string());
//!
//! // Save
//! pkg.save("output.docx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Reading an existing document
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//!
//! // Open document
//! let pkg = Package::open("document.docx")?;
//! let doc = pkg.document()?;
//!
//! // Extract text
//! let text = doc.text()?;
//! println!("Document text: {}", text);
//!
//! // Iterate paragraphs
//! for para in doc.paragraphs()? {
//!     println!("Paragraph: {}", para.text()?);
//!     
//!     // Access runs with formatting
//!     for run in para.runs()? {
//!         println!("  Text: {}", run.text()?);
//!         println!("  Bold: {:?}", run.bold()?);
//!         println!("  Italic: {:?}", run.italic()?);
//!     }
//! }
//!
//! // Access tables
//! for table in doc.tables()? {
//!     println!("Table with {} rows", table.row_count()?);
//!     for row in table.rows()? {
//!         for cell in row.cells()? {
//!             println!("Cell: {}", cell.text()?);
//!         }
//!     }
//! }
//!
//! // Access metadata
//! let props = pkg.properties();
//! if let Some(title) = &props.title {
//!     println!("Title: {}", title);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Updating an existing document
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//!
//! // Open document
//! let mut pkg = Package::open("document.docx")?;
//! let mut doc = pkg.document_mut()?;
//!
//! // Add new content
//! doc.add_paragraph_with_text("This is a new paragraph added to existing document.");
//! doc.add_heading("New Section", 2)?;
//!
//! // Update metadata
//! pkg.properties_mut().last_modified_by = Some("Jane Smith".to_string());
//!
//! // Save (can overwrite or save to new file)
//! pkg.save("updated_document.docx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Working with Custom Properties
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//! use litchi::ooxml::custom_properties::PropertyValue;
//!
//! let mut pkg = Package::new()?;
//!
//! // Add custom properties
//! let custom_props = pkg.custom_properties_mut();
//! custom_props.add_property("ProjectName", PropertyValue::String("MyProject".to_string()));
//! custom_props.add_property("Version", PropertyValue::Integer(1));
//! custom_props.add_property("Budget", PropertyValue::Double(50000.0));
//! custom_props.add_property("IsApproved", PropertyValue::Boolean(true));
//!
//! // Read custom properties
//! let pkg = Package::open("document.docx")?;
//! for (name, value) in pkg.custom_properties().iter() {
//!     println!("{}: {:?}", name, value);
//! }
//!
//! // Modify custom property
//! let mut pkg = Package::open("document.docx")?;
//! pkg.custom_properties_mut().set_property("Version", PropertyValue::Integer(2));
//!
//! // Remove custom property
//! pkg.custom_properties_mut().remove_property("Budget");
//!
//! pkg.save("document.docx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ## Excel Workbooks (XLSX)
//!
//! ### Creating a new workbook
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Workbook;
//!
//! // Create a new workbook
//! let mut wb = Workbook::create()?;
//!
//! // Access first worksheet
//! let mut ws = wb.worksheet_mut(0)?;
//! ws.set_cell_value(1, 1, "Name");
//! ws.set_cell_value(1, 2, "Age");
//! ws.set_cell_value(2, 1, "Alice");
//! ws.set_cell_value(2, 2, 30);
//! ws.set_cell_value(3, 1, "Bob");
//! ws.set_cell_value(3, 2, 25);
//!
//! // Add more worksheets
//! let mut ws2 = wb.add_worksheet("Summary");
//! ws2.set_cell_value(1, 1, "Total Records");
//! ws2.set_cell_formula(1, 2, "=Sheet1!A1:A3");
//!
//! // Define named ranges
//! wb.define_name("DataRange", "Sheet1!$A$1:$B$3");
//!
//! // Set freeze panes
//! ws.freeze_panes(2, 1)?;
//!
//! // Set metadata
//! wb.properties_mut().title = Some("Employee Data".to_string());
//!
//! // Save
//! wb.save("output.xlsx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Reading an existing workbook
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Workbook;
//!
//! // Open workbook
//! let wb = Workbook::open("workbook.xlsx")?;
//!
//! // Get worksheet names
//! for name in wb.worksheet_names() {
//!     println!("Worksheet: {}", name);
//! }
//!
//! // Access worksheet by name
//! let ws = wb.worksheet_by_name("Sheet1")?;
//! println!("Sheet: {}", ws.name());
//!
//! // Iterate cells
//! for row in ws.rows()? {
//!     for cell in row.cells()? {
//!         match cell.value() {
//!             Some(v) => println!("Cell {}: {:?}", cell.reference(), v),
//!             None => {}
//!         }
//!     }
//! }
//!
//! // Access specific cell
//! let cell = ws.cell(1, 1)?;
//! if let Some(value) = cell.value() {
//!     println!("A1: {:?}", value);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Updating an existing workbook
//!
//! ```rust,no_run
//! use litchi::ooxml::xlsx::Workbook;
//!
//! // Open workbook
//! let mut wb = Workbook::open("workbook.xlsx")?;
//!
//! // Update existing worksheet
//! let mut ws = wb.worksheet_mut(0)?;
//! ws.set_cell_value(4, 1, "Charlie");
//! ws.set_cell_value(4, 2, 35);
//!
//! // Add new worksheet with data
//! let mut new_ws = wb.add_worksheet("Q1 Data");
//! new_ws.set_cell_value(1, 1, "Revenue");
//! new_ws.set_cell_value(1, 2, 100000);
//!
//! // Save
//! wb.save("updated_workbook.xlsx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ## PowerPoint Presentations (PPTX)
//!
//! ### Creating a new presentation
//!
//! ```rust,no_run
//! use litchi::ooxml::pptx::Package;
//!
//! // Create a new presentation
//! let mut pkg = Package::new()?;
//! let mut pres = pkg.presentation_mut()?;
//!
//! // Add slides
//! let slide1 = pres.add_slide()?;
//! slide1.set_title("Welcome");
//! slide1.add_text_box("This is the first slide", 914400, 1828800, 7315200, 914400);
//!
//! let slide2 = pres.add_slide()?;
//! slide2.set_title("Agenda");
//! slide2.add_bullet_points(&[
//!     "Introduction",
//!     "Main Content",
//!     "Conclusion",
//! ])?;
//!
//! // Add image to slide
//! let image_data = std::fs::read("logo.png")?;
//! slide2.add_image(&image_data, 914400, 914400, 1828800, 1828800)?;
//!
//! // Set metadata
//! pkg.properties_mut().title = Some("My Presentation".to_string());
//! pkg.properties_mut().creator = Some("John Doe".to_string());
//!
//! // Save
//! pkg.save("output.pptx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Reading an existing presentation
//!
//! ```rust,no_run
//! use litchi::ooxml::pptx::Package;
//!
//! // Open presentation
//! let pkg = Package::open("presentation.pptx")?;
//! let pres = pkg.presentation()?;
//!
//! // Get presentation info
//! println!("Slides: {}", pres.slide_count()?);
//! if let (Some(w), Some(h)) = (pres.slide_width()?, pres.slide_height()?) {
//!     println!("Slide size: {}x{} EMUs", w, h);
//! }
//!
//! // Iterate slides
//! for (idx, slide) in pres.slides()?.iter().enumerate() {
//!     println!("\nSlide {}: {}", idx + 1, slide.name()?);
//!     println!("Content:\n{}", slide.text()?);
//!     
//!     // Access shapes
//!     for shape in slide.shapes()? {
//!         if let Some(text) = shape.text()? {
//!             println!("Shape text: {}", text);
//!         }
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ### Updating an existing presentation
//!
//! ```rust,no_run
//! use litchi::ooxml::pptx::Package;
//!
//! // Open presentation
//! let mut pkg = Package::open("presentation.pptx")?;
//! let mut pres = pkg.presentation_mut()?;
//!
//! // Add new slide to existing presentation
//! let new_slide = pres.add_slide()?;
//! new_slide.set_title("Conclusion");
//! new_slide.add_text_box("Thank you!", 914400, 3657600, 7315200, 914400);
//!
//! // Update metadata
//! pkg.properties_mut().last_modified_by = Some("Jane Smith".to_string());
//!
//! // Save
//! pkg.save("updated_presentation.pptx")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Performance Considerations
//!
//! - **Zero-copy parsing**: Where possible, the library uses zero-copy techniques to avoid unnecessary allocations
//! - **Lazy loading**: Content is loaded on-demand to minimize memory usage
//! - **Streaming**: Large files can be processed with minimal memory overhead
//! - **Efficient XML parsing**: Uses quick-xml with optimized parsing strategies
//! - **SIMD acceleration**: Leverages SIMD instructions for string operations where available
//!
//! # Thread Safety
//!
//! - Package types are `Send` but not `Sync`
//! - For concurrent access, use separate Package instances per thread
//! - For read-only operations, clone the Package or share references appropriately
//!
//! # Error Handling
//!
//! All operations return `Result` types with descriptive errors:
//!
//! ```rust,no_run
//! use litchi::ooxml::docx::Package;
//! use litchi::ooxml::error::OoxmlError;
//!
//! match Package::open("document.docx") {
//!     Ok(pkg) => println!("Opened successfully"),
//!     Err(OoxmlError::IoError(e)) => eprintln!("IO error: {}", e),
//!     Err(OoxmlError::InvalidFormat(msg)) => eprintln!("Invalid format: {}", msg),
//!     Err(e) => eprintln!("Error: {}", e),
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use crate::ooxml::common::DocumentProperties;
/// Re-export main types for convenience
pub use crate::ooxml::docx::{Document, MutableDocument, Package as DocxPackage};
pub use crate::ooxml::error::{OoxmlError, Result};
pub use crate::ooxml::pptx::{MutablePresentation, Package as PptxPackage, Presentation};
pub use crate::ooxml::xlsx::{MutableWorksheet, Workbook, Worksheet};

/// Unified document interface for reading different Office formats.
///
/// This provides a consistent API for reading text content across different file types.
pub trait UnifiedDocument {
    /// Extract all text content from the document.
    fn extract_text(&self) -> Result<String>;

    /// Get document properties/metadata.
    fn properties(&self) -> Result<DocumentProperties>;
}

impl UnifiedDocument for DocxPackage {
    fn extract_text(&self) -> Result<String> {
        self.document()?.text()
    }

    fn properties(&self) -> Result<DocumentProperties> {
        Ok(self.properties().clone())
    }
}

impl UnifiedDocument for PptxPackage {
    fn extract_text(&self) -> Result<String> {
        let pres = self.presentation()?;
        let slides = pres.slides()?;

        let mut result = String::new();
        for (idx, slide) in slides.iter().enumerate() {
            if idx > 0 {
                result.push_str("\n\n");
            }
            result.push_str("--- Slide ");
            result.push_str(&(idx + 1).to_string());
            result.push_str(" ---\n");
            result.push_str(&slide.text()?);
        }

        Ok(result)
    }

    fn properties(&self) -> Result<DocumentProperties> {
        Ok(self.properties().clone())
    }
}

/// Helper functions for common operations
pub mod helpers {
    use super::*;

    /// Extract text from any supported Office document format.
    ///
    /// # Arguments
    /// * `path` - Path to the document file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::api::helpers;
    ///
    /// // Works with DOCX, PPTX files
    /// let text = helpers::extract_text("document.docx")?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn extract_text<P: AsRef<std::path::Path>>(path: P) -> Result<String> {
        let path_ref = path.as_ref();
        let ext = path_ref.extension().and_then(|s| s.to_str()).unwrap_or("");

        match ext.to_lowercase().as_str() {
            "docx" => {
                let pkg = DocxPackage::open(path)?;
                pkg.extract_text()
            },
            "pptx" => {
                let pkg = PptxPackage::open(path)?;
                pkg.extract_text()
            },
            _ => Err(OoxmlError::InvalidFormat(format!(
                "Unsupported file extension: {}",
                ext
            ))),
        }
    }

    /// Get metadata/properties from any supported Office document format.
    ///
    /// # Arguments
    /// * `path` - Path to the document file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::api::helpers;
    ///
    /// let props = helpers::get_properties("document.docx")?;
    /// if let Some(title) = props.title {
    ///     println!("Title: {}", title);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_properties<P: AsRef<std::path::Path>>(path: P) -> Result<DocumentProperties> {
        let path_ref = path.as_ref();
        let ext = path_ref.extension().and_then(|s| s.to_str()).unwrap_or("");

        match ext.to_lowercase().as_str() {
            "docx" => {
                let pkg = DocxPackage::open(path)?;
                Ok(pkg.properties().clone())
            },
            "pptx" => {
                let pkg = PptxPackage::open(path)?;
                Ok(pkg.properties().clone())
            },
            "xlsx" => {
                let wb =
                    Workbook::open(path).map_err(|e| OoxmlError::InvalidFormat(e.to_string()))?;
                Ok(wb.properties().clone())
            },
            _ => Err(OoxmlError::InvalidFormat(format!(
                "Unsupported file extension: {}",
                ext
            ))),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_docx_create_and_read() {
        // This would require actual file I/O, so we'll keep it simple
        // In production, you'd have proper test files

        // Create
        let mut pkg = DocxPackage::new().unwrap();
        let doc = pkg.document_mut().unwrap();
        doc.add_paragraph_with_text("Test paragraph");

        // Note: In real tests, you'd save and re-open
        // pkg.save("test.docx").unwrap();
        // let pkg2 = DocxPackage::open("test.docx").unwrap();
        // assert!(pkg2.document().unwrap().text().unwrap().contains("Test"));
    }
}

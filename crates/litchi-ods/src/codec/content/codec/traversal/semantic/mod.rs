//! Semantic ODS content traversal facade.
//!
//! The public parser entry point lives on [`Parser`], while the streaming
//! sheet traversal is kept in its own owner module so the traversal layer can
//! grow without returning to a monolithic semantic file.

mod sheets;

use super::{Parser, Result, Sheet};
use sheets::SheetTraversal;

// Keep the facade intentionally small: `sheets` owns the streaming state
// machine and its namespace-aware event handling.
impl Parser {
    /// Parse all sheets from ODS content.xml.
    pub fn parse_sheets(xml_content: &str) -> Result<Vec<Sheet>> {
        <Self as SheetTraversal>::parse_sheets(xml_content)
    }
}

//! Numbers Sheet Structure
//!
//! Sheets in Numbers documents contain multiple tables and other content.

use super::table::NumbersTable;

/// Represents a sheet in a Numbers document
#[derive(Debug, Clone)]
pub struct NumbersSheet {
    /// Sheet name
    pub name: String,
    /// Sheet index (0-based)
    pub index: usize,
    /// Tables in this sheet
    pub tables: Vec<NumbersTable>,
}

impl NumbersSheet {
    /// Create a new sheet
    pub fn new(name: String, index: usize) -> Self {
        Self {
            name,
            index,
            tables: Vec::new(),
        }
    }

    /// Add a table to the sheet
    pub fn add_table(&mut self, table: NumbersTable) {
        self.tables.push(table);
    }

    /// Get a table by name
    pub fn get_table(&self, name: &str) -> Option<&NumbersTable> {
        self.tables.iter().find(|t| t.name == name)
    }

    /// Get a mutable reference to a table by name
    pub fn get_table_mut(&mut self, name: &str) -> Option<&mut NumbersTable> {
        self.tables.iter_mut().find(|t| t.name == name)
    }

    /// Get all table names
    pub fn table_names(&self) -> Vec<String> {
        self.tables.iter().map(|t| t.name.clone()).collect()
    }

    /// Check if sheet is empty
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// Get total number of tables
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Get total number of cells across all tables
    pub fn total_cell_count(&self) -> usize {
        self.tables
            .iter()
            .map(|t| t.row_count * t.column_count)
            .sum()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::iwa::numbers::cell::CellValue;

    #[test]
    fn test_sheet_creation() {
        let sheet = NumbersSheet::new("Sheet1".to_string(), 0);
        assert_eq!(sheet.name, "Sheet1");
        assert_eq!(sheet.index, 0);
        assert!(sheet.is_empty());
        assert_eq!(sheet.table_count(), 0);
    }

    #[test]
    fn test_sheet_add_table() {
        let mut sheet = NumbersSheet::new("Sheet1".to_string(), 0);
        
        let mut table = NumbersTable::new("Table1".to_string());
        table.set_cell(0, 0, CellValue::Number(1.0));
        
        sheet.add_table(table);
        
        assert_eq!(sheet.table_count(), 1);
        assert!(!sheet.is_empty());
    }

    #[test]
    fn test_sheet_get_table() {
        let mut sheet = NumbersSheet::new("Sheet1".to_string(), 0);
        
        let table1 = NumbersTable::new("Table1".to_string());
        let table2 = NumbersTable::new("Table2".to_string());
        
        sheet.add_table(table1);
        sheet.add_table(table2);
        
        assert!(sheet.get_table("Table1").is_some());
        assert!(sheet.get_table("Table2").is_some());
        assert!(sheet.get_table("Table3").is_none());
    }

    #[test]
    fn test_sheet_table_names() {
        let mut sheet = NumbersSheet::new("Sheet1".to_string(), 0);
        
        sheet.add_table(NumbersTable::new("Table1".to_string()));
        sheet.add_table(NumbersTable::new("Table2".to_string()));
        sheet.add_table(NumbersTable::new("Table3".to_string()));
        
        let names = sheet.table_names();
        assert_eq!(names.len(), 3);
        assert!(names.contains(&"Table1".to_string()));
        assert!(names.contains(&"Table2".to_string()));
        assert!(names.contains(&"Table3".to_string()));
    }
}


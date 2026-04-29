#[derive(Debug, Clone)]
pub struct Location {
    pub reference: String,
    pub first_header_row: u32,
    pub first_data_row: u32,
    pub first_data_col: u32,
    pub row_page_count: Option<u32>,
    pub col_page_count: Option<u32>,
}

impl Default for Location {
    fn default() -> Self {
        Self {
            reference: String::new(),
            first_header_row: 1,
            first_data_row: 1,
            first_data_col: 1,
            row_page_count: None,
            col_page_count: None,
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct PivotTableStyle {
    pub name: Option<String>,
    pub show_row_headers: Option<bool>,
    pub show_col_headers: Option<bool>,
    pub show_row_stripes: Option<bool>,
    pub show_col_stripes: Option<bool>,
    pub show_last_column: Option<bool>,
}

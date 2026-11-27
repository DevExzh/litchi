//! Shared worksheet-chart facade for XLSX, and the worksheet chart
//! lifecycle family ([`edit`]).

pub mod edit;

pub use litchi_spreadsheet_drawing::chart::{read, write};

//! Layered BIFF12/XLSB workbook codec facade.
//!
//! Record semantics, stream materialization, and package mutation live in
//! dedicated owners so the workbook facade remains stable and ergonomic.

mod codec;
mod package;
mod records;

#[cfg(test)]
mod tests;

//! Layered BIFF writer operations for the XLS writer facade.
//!
//! Each child module owns one semantic family of writer mutations while the
//! inherent `Writer` API remains unchanged at the public facade.

mod cells;
mod drawings;
mod lifecycle;
mod names;
mod pivots;
mod protection;
mod tables;
mod workbook;
mod worksheet;

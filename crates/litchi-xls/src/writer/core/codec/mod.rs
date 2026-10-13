//! Layered BIFF writer operations for the XLS writer facade.
//!
//! Each child module owns one semantic family of writer mutations while the
//! inherent `Builder` API remains unchanged at the public facade.

mod cells;
mod drawings;
mod lifecycle;
mod names;
mod pivots;
mod properties;
mod protection;
mod tables;
mod workbook;
mod worksheet;

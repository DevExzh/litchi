//! Inert reader for Microsoft `ChartEx` (cx:chartSpace) parts.

pub const CONTENT_TYPE: &str = "application/vnd.ms-office.chartex+xml";

mod limits;
mod package;
mod semantic;
mod xml;

#[cfg(test)]
mod tests;

//! Inert Word OLE-control metadata (`OcxInfo`/`RgxOcxInfo`).
//!
//! This context owns only the table records described by [MS-DOC] sections
//! 2.9.161 and 2.9.229. It does not resolve, activate, render, or otherwise
//! execute a control.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{parse_bytes, to_bytes};
pub use model::{Flags, OcxInfo, RgxOcxInfo, Story};
pub use package::{FIB_INDEX_PLC_OCX, parse};

//! Semantic XML codec facade.

mod parser;
mod support;
mod writer;

pub(in crate::web) use parser::*;
pub(in crate::web) use support::*;
pub(in crate::web) use writer::*;

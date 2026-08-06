//! Bounded BIFF8 primitives for worksheet-table payloads.
//!
//! This facade keeps the historical `codec::binary::*` helper paths intact
//! while the wire concerns live in focused primitive, string, formula, web,
//! and retained-future-record modules.

mod formulas;
mod future_records;
mod primitives;
mod strings;
mod web;

pub(super) use formulas::{append_formula, parse_formula, parse_list_formula_extra_end};
pub(in crate::list_object) use future_records::PendingFeature;
pub(in crate::list_object) use primitives::{
    append_frt, append_range, parse_range, record, u16_at, u32_at, validate_frt, validate_frt_any,
};
pub(in crate::list_object) use strings::{append_string, parse_string};
pub(super) use web::{append_web_info, parse_web_info};

//! Layered, namespace-aware parser for sparse `SpreadsheetML` worksheet data.
//!
//! The facade keeps worksheet consumers independent from the streaming codec;
//! raw records, semantic materialization, and validation rules live in focused
//! sibling modules so the hot path remains allocation-conscious.

mod codec;
pub(crate) mod edit;
mod model;
mod semantic;
mod validation;
mod x14ac;

#[cfg(test)]
mod tests;

pub(crate) use model::merge_successor;
pub(crate) use validation::{
    optional_bool, optional_u32, parse_a1, parse_one_based_row, required_u32,
};

use crate::cell::{Store, Text};
use crate::error::{Result, invalid};
use crate::layout::Defaults;
use litchi_ooxml_common::mce::process_ooxml;

pub(crate) fn parse<'a, F>(content: &[u8], strings: F) -> Result<Store>
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
{
    let needs_extension_capture = x14ac::may_contain_descent(content);
    let extensions = if needs_extension_capture {
        x14ac::capture(content)?
    } else {
        x14ac::Values::default()
    };
    let parsed = (|| {
        let processed = process_ooxml(content)?;
        let content = std::str::from_utf8(processed.as_ref())
            .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
        model::Parser::parse(content, strings, extensions)
    })();

    if parsed.is_err() && !needs_extension_capture {
        // The extension scan historically ran first. Repeat it only on a
        // rejected plain worksheet so its typed error and error precedence
        // remain unchanged without charging successful no-extension reads.
        x14ac::capture(content)?;
    }
    parsed
}

pub(crate) fn parse_defaults(content: &[u8]) -> Result<Option<Defaults>> {
    let mut descent = x14ac::capture_defaults(content)?;
    let processed = process_ooxml(content)?;
    let content = std::str::from_utf8(processed.as_ref())
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    codec::parse_processed_defaults(content, descent.take())
}

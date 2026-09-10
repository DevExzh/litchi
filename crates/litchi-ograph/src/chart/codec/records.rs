//! Layered MS-OGRAPH chart record codecs.

use super::super::model::Chart;
use crate::{Limits, Result};

mod cache;
mod encode;
mod links;
mod parse;
mod patch;
mod text;
mod validate;
mod wire;

pub(crate) use encode::encode;
pub(crate) use parse::parse;
pub(crate) use patch::patch;
pub(crate) use validate::valid_props;
pub(crate) use wire::{PLOT_AREA, SHT_PROPS};

/// Validates a chart from the semantic authoring boundary.
pub(crate) fn validate_chart(chart: &Chart, limits: Limits, require_topology: bool) -> Result<()> {
    validate::validate(chart, limits, require_topology)
}

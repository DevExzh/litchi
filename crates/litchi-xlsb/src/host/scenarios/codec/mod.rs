//! BIFF12 wire conversion for Scenario Manager records.

mod semantic;
mod wire;

pub(crate) use semantic::{parse_manager, write_manager};
pub(crate) use wire::{begin_manager, end_manager};

#[cfg(test)]
pub(crate) use wire::{begin_scenario, scenario_cell};

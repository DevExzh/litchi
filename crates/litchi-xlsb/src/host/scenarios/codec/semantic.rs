//! Semantic conversion between the scenario record sequence and its model.

use super::super::model::{
    Child, Entry, MAX_UNKNOWN_PAYLOAD, MAX_UNKNOWN_RECORDS, Manager, Scenario, UnknownRecord,
};
use super::wire;
use crate::package::error::{Error, Result};
use crate::raw::{Encoder, Record};

#[derive(Default)]
struct OpaqueBudget {
    count: usize,
    bytes: usize,
}

fn retain_unknown(record: &Record<'_>, budget: &mut OpaqueBudget) -> Result<UnknownRecord> {
    budget.count = budget.count.checked_add(1).ok_or(Error::Unrecognized {
        typ: "scenario unknown records",
        val: "record count overflow",
    })?;
    if budget.count > MAX_UNKNOWN_RECORDS {
        return Err(Error::Unrecognized {
            typ: "scenario unknown records",
            val: "record count exceeds",
        });
    }
    budget.bytes = budget
        .bytes
        .checked_add(record.payload().len())
        .ok_or(Error::Unrecognized {
            typ: "scenario unknown records",
            val: "payload byte count overflow",
        })?;
    if budget.bytes > MAX_UNKNOWN_PAYLOAD {
        return Err(Error::Unrecognized {
            typ: "scenario unknown records",
            val: "payload bytes exceed",
        });
    }
    UnknownRecord::new(record.kind().get(), record.payload())
}

pub(crate) fn parse_manager<'a>(
    header: &[u8],
    records: &[Record<'a>],
    limits: crate::raw::Limits,
) -> Result<Manager> {
    let header = wire::parse_manager_header(header, limits)?;
    let current = if header.current == u16::MAX {
        None
    } else {
        Some(usize::from(header.current))
    };
    let shown = if header.shown == u16::MAX {
        None
    } else {
        Some(usize::from(header.shown))
    };

    let mut scenarios = Vec::new();
    let mut unknown = Vec::new();
    let mut order = Vec::new();
    let mut budget = OpaqueBudget::default();
    let mut index = 0usize;
    while let Some(record) = records.get(index) {
        match record.kind() {
            kind if kind == wire::begin_scenario() => {
                let end = records[index + 1..]
                    .iter()
                    .position(|candidate| candidate.kind() == wire::end_scenario())
                    .map(|offset| index + offset + 1)
                    .ok_or(Error::UnexpectedEndOfStream("BrtEndSct"))?;
                let scenario = parse_scenario(
                    record.payload(),
                    &records[index + 1..end],
                    &mut budget,
                    limits,
                )?;
                let scenario_index = scenarios.len();
                scenarios.push(scenario);
                order.push(Entry::Scenario(scenario_index));
                index = end + 1;
            },
            kind if kind == wire::end_scenario() => {
                return Err(Error::UnexpectedRecord {
                    expected: wire::begin_scenario().get(),
                    found: kind.get(),
                });
            },
            kind if kind == wire::begin_manager() || kind == wire::end_manager() => {
                return Err(Error::Unrecognized {
                    typ: "Scenario Manager collection",
                    val: "unexpected record",
                });
            },
            _ => {
                let opaque = retain_unknown(record, &mut budget)?;
                let opaque_index = unknown.len();
                unknown.push(opaque);
                order.push(Entry::Unknown(opaque_index));
                index += 1;
            },
        }
    }

    Manager::from_wire(
        current,
        shown,
        header.result_ranges,
        scenarios,
        unknown,
        order,
    )
}

fn parse_scenario<'a>(
    header: &[u8],
    records: &[Record<'a>],
    budget: &mut OpaqueBudget,
    limits: crate::raw::Limits,
) -> Result<Scenario> {
    let header = wire::parse_scenario_header(header, limits)?;
    let mut changed_cells = Vec::new();
    let mut unknown = Vec::new();
    let mut order = Vec::new();
    for record in records {
        match record.kind() {
            kind if kind == wire::scenario_cell() => {
                let index = changed_cells.len();
                changed_cells.push(wire::parse_changed_cell(record.payload(), limits)?);
                order.push(Child::Changed(index));
            },
            kind if kind == wire::begin_scenario()
                || kind == wire::end_scenario()
                || kind == wire::begin_manager()
                || kind == wire::end_manager() =>
            {
                return Err(Error::Unrecognized {
                    typ: "Scenario collection",
                    val: "unexpected record",
                });
            },
            _ => {
                let opaque = retain_unknown(record, budget)?;
                let index = unknown.len();
                unknown.push(opaque);
                order.push(Child::Unknown(index));
            },
        }
    }
    if usize::from(header.count) != changed_cells.len() {
        return Err(Error::Unrecognized {
            typ: "BrtBeginSct cref",
            val: "declared, found BrtSlc records",
        });
    }
    Scenario::from_wire(
        header.name,
        header.locked,
        header.hidden,
        header.comment,
        header.user_name,
        changed_cells,
        unknown,
        order,
    )
}

pub(crate) fn write_manager(manager: &Manager, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    manager.validate()?;
    let mut bytes = Vec::new();
    let mut writer = Encoder::new(&mut bytes, limits);
    wire::write_record(
        &mut writer,
        wire::begin_manager(),
        &wire::write_manager_header(
            manager.current(),
            manager.shown(),
            manager.result_ranges(),
            limits,
        )?,
    )?;

    if manager.order.is_empty() {
        for scenario in manager.scenarios() {
            write_scenario(&mut writer, scenario, limits)?;
        }
        for record in manager.unknown_records() {
            write_unknown(&mut writer, record)?;
        }
    } else {
        for entry in &manager.order {
            match *entry {
                Entry::Scenario(index) => write_scenario(
                    &mut writer,
                    manager.scenarios().get(index).ok_or(Error::Unrecognized {
                        typ: "BrtBeginScenMan order",
                        val: "scenario index is invalid",
                    })?,
                    limits,
                )?,
                Entry::Unknown(index) => write_unknown(
                    &mut writer,
                    manager
                        .unknown_records()
                        .get(index)
                        .ok_or(Error::Unrecognized {
                            typ: "BrtBeginScenMan order",
                            val: "unknown-record index is invalid",
                        })?,
                )?,
            }
        }
    }
    wire::write_record(&mut writer, wire::end_manager(), &[])?;
    Ok(bytes)
}

fn write_scenario<W: std::io::Write>(
    writer: &mut Encoder<W>,
    scenario: &Scenario,
    limits: crate::raw::Limits,
) -> Result<()> {
    wire::write_record(
        writer,
        wire::begin_scenario(),
        &wire::write_scenario_header(
            scenario.changed_cells().len(),
            scenario.locked(),
            scenario.hidden(),
            scenario.name(),
            scenario.comment(),
            scenario.user_name(),
            limits,
        )?,
    )?;

    if scenario.order.is_empty() {
        for cell in scenario.changed_cells() {
            wire::write_record(
                writer,
                wire::scenario_cell(),
                &wire::write_changed_cell(cell, limits)?,
            )?;
        }
        for record in scenario.unknown_records() {
            write_unknown(writer, record)?;
        }
    } else {
        for entry in &scenario.order {
            match *entry {
                Child::Changed(index) => wire::write_record(
                    writer,
                    wire::scenario_cell(),
                    &wire::write_changed_cell(
                        scenario
                            .changed_cells()
                            .get(index)
                            .ok_or(Error::Unrecognized {
                                typ: "BrtBeginSct order",
                                val: "changed-cell index is invalid",
                            })?,
                        limits,
                    )?,
                )?,
                Child::Unknown(index) => write_unknown(
                    writer,
                    scenario
                        .unknown_records()
                        .get(index)
                        .ok_or(Error::Unrecognized {
                            typ: "BrtBeginSct order",
                            val: "unknown-record index is invalid",
                        })?,
                )?,
            }
        }
    }
    wire::write_record(writer, wire::end_scenario(), &[])
}

fn write_unknown<W: std::io::Write>(writer: &mut Encoder<W>, record: &UnknownRecord) -> Result<()> {
    let kind = crate::raw::Kind::new(record.kind()).map_err(|error| Error::from(error))?;
    wire::write_record(writer, kind, record.payload())
}

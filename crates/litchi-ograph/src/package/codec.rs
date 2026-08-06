use std::io::Cursor;

use litchi_cfb::OleFile;
use litchi_cfb::OleWriter;

use crate::Limits;
use crate::Result;
use crate::chart;

use super::snapshot::Snapshot;
use super::validation::{WorkbookLayout, check_limit};

pub(super) const WORKBOOK: &str = "Workbook";
pub(super) const COMP_OBJ: &str = "\u{1}CompObj";
pub(super) const OLE: &str = "\u{1}Ole";

pub(super) fn read_workbook(
    package_bytes: &[u8],
    layout: WorkbookLayout,
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut cfb = OleFile::open(Cursor::new(package_bytes))?;
    let bytes = cfb.open_stream(&[WORKBOOK])?;
    check_limit("Workbook bytes", bytes.len(), limits.max_workbook_bytes)?;
    layout.check(bytes.len())?;
    Ok(bytes)
}

pub(super) fn read_chart(source: &Snapshot) -> Result<chart::Chart> {
    let bytes = read_workbook(source.source_bytes(), source.workbook, source.limits)?;
    let workbook = super::semantic::Workbook::with_limits(bytes, source.limits)?;
    chart::Chart::parse(workbook.chart(), chart::Context::graph())
}

pub(super) fn replace_chart(source: &Snapshot, replacement: &[u8]) -> Result<Vec<u8>> {
    let workbook = read_workbook(source.source_bytes(), source.workbook, source.limits)?;
    let chart_range = source.workbook.chart_start..source.workbook.chart_end;
    let chart = workbook
        .get(chart_range.clone())
        .ok_or(crate::Error::UnsupportedMutation {
            operation: "package-chart-patch",
            reason: "validated chart range falls outside the Workbook stream",
        })?;
    if chart.len() != replacement.len() {
        return Err(crate::Error::UnsupportedMutation {
            operation: "package-chart-patch",
            reason: "typed chart replacement changes the Workbook envelope length",
        });
    }

    let mut workbook = workbook;
    workbook
        .get_mut(chart_range)
        .ok_or(crate::Error::UnsupportedMutation {
            operation: "package-chart-patch",
            reason: "validated chart range cannot be mutated",
        })?
        .copy_from_slice(replacement);

    let mut cfb = OleFile::open(Cursor::new(source.source_bytes()))?;
    let comp_obj = if source.topology.comp_obj_bytes().is_some() {
        Some(cfb.open_stream(&[COMP_OBJ])?)
    } else {
        None
    };
    let ole = if source.topology.ole_bytes().is_some() {
        Some(cfb.open_stream(&[OLE])?)
    } else {
        None
    };

    let sector_size = match source.source_bytes().get(0x1E..0x20) {
        Some([0x09, 0x00]) => 512,
        Some([0x0C, 0x00]) => 4096,
        _ => {
            return Err(crate::Error::UnsupportedMutation {
                operation: "package-chart-patch",
                reason: "source CFB sector size is not supported by the package writer",
            });
        },
    };
    let mut writer = OleWriter::with_sector_size(sector_size)?;
    writer.create_stream(&[WORKBOOK], &workbook)?;
    if let Some(comp_obj) = comp_obj {
        writer.create_stream(&[COMP_OBJ], &comp_obj)?;
    }
    if let Some(ole) = ole {
        writer.create_stream(&[OLE], &ole)?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let output = output.into_inner();
    check_limit(
        "package bytes",
        output.len(),
        source.limits.max_package_bytes,
    )?;
    Ok(output)
}

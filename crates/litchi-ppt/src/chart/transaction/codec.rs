//! OLE2 stream assembly for the chart transaction.

use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};
use litchi_ograph::{Limits, PackageRef};

use crate::package::Result;

pub(super) const COMP_OBJ: &str = "\u{1}CompObj";
pub(super) const OLE: &str = "\u{1}Ole";
pub(super) const WORKBOOK: &str = "Workbook";

pub(super) struct Parts {
    pub(super) workbook: Vec<u8>,
    pub(super) comp_obj: Option<Vec<u8>>,
    pub(super) ole: Option<Vec<u8>>,
    pub(super) chart_start: usize,
    pub(super) chart_end: usize,
}

/// Validate a standalone `OGraph` package and copy its logical root streams.
pub(super) fn read(bytes: &[u8], limits: Limits) -> Result<Parts> {
    let package = PackageRef::with_limits(bytes, limits)?;
    let workbook = package.workbook()?;
    let chart = workbook.chart();
    let chart_start = chart.offset();
    let chart_end = chart_start.checked_add(chart.as_bytes().len()).ok_or(
        litchi_ograph::Error::SizeOverflow {
            resource: "chart substream",
        },
    )?;
    let workbook_bytes = workbook.into_bytes();

    let mut cfb = OleFile::open(Cursor::new(bytes))?;
    let comp_obj = package
        .topology()
        .comp_obj_bytes()
        .is_some()
        .then(|| cfb.open_stream(&[COMP_OBJ]))
        .transpose()?;
    let ole = package
        .topology()
        .ole_bytes()
        .is_some()
        .then(|| cfb.open_stream(&[OLE]))
        .transpose()?;

    Ok(Parts {
        workbook: workbook_bytes,
        comp_obj,
        ole,
        chart_start,
        chart_end,
    })
}

/// Rebuild and revalidate the exact standalone `OGraph` root topology.
pub(super) fn write(
    workbook: &[u8],
    comp_obj: Option<&[u8]>,
    ole: Option<&[u8]>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.create_stream(&[WORKBOOK], workbook)?;
    if let Some(comp_obj_bytes) = comp_obj {
        writer.create_stream(&[COMP_OBJ], comp_obj_bytes)?;
    }
    if let Some(ole_bytes) = ole {
        writer.create_stream(&[OLE], ole_bytes)?;
    }

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let bytes = output.into_inner();
    PackageRef::with_limits(&bytes, limits)?;
    Ok(bytes)
}

use std::io::Cursor;

use litchi_cfb::OleFile;

use crate::Limits;
use crate::Result;

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

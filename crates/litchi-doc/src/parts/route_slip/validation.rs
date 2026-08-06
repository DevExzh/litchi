//! Validation shared by the route-slip semantic and package owners.

use super::model::{Metadata, Protection, Recipient};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};

pub(crate) const ROUTE_SLIP_FIB_INDEX: usize = 70;
pub(crate) const ROUTE_SLIP_POINTER_OFFSET: usize = 154 + ROUTE_SLIP_FIB_INDEX * 8;

pub(crate) fn metadata(value: &Metadata) -> Result<()> {
    value.validate()
}

pub(crate) fn metadata_option(value: &Option<Metadata>) -> Result<()> {
    value.as_ref().map_or(Ok(()), metadata)
}

pub(crate) fn recipient(value: &Recipient) -> Result<()> {
    value.validate()
}

pub(crate) fn editable(value: &Metadata) -> std::result::Result<(), Protection> {
    if value.protection == Protection::Off {
        Ok(())
    } else {
        Err(value.protection)
    }
}

pub(crate) fn package_fib(fib: &FileInformationBlock) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    if fib.is_encrypted() {
        return Err(PackageError::InvalidFormat(
            "encrypted DOC packages cannot be edited by the route-slip owner".into(),
        ));
    }
    if fib.table_pointer_count().is_none() {
        return Err(corrupted(
            "WordDocument FIB table-pointer array is truncated",
        ));
    }
    if fib.table_pointer_count().unwrap_or(0) <= ROUTE_SLIP_FIB_INDEX {
        return Err(corrupted(
            "WordDocument FIB does not expose fcRouteSlip/lcbRouteSlip",
        ));
    }
    Ok(())
}

pub(crate) fn route_pointer_location(fib: &FileInformationBlock) -> Result<usize> {
    package_fib(fib)?;
    let end = ROUTE_SLIP_POINTER_OFFSET
        .checked_add(8)
        .ok_or_else(|| corrupted("fcRouteSlip/lcbRouteSlip offset overflows"))?;
    if end > fib.raw_data().len() {
        return Err(corrupted(
            "WordDocument FIB does not contain fcRouteSlip/lcbRouteSlip",
        ));
    }
    Ok(ROUTE_SLIP_POINTER_OFFSET)
}

pub(crate) fn table_range(table: &[u8], offset: u32, length: u32) -> Result<&[u8]> {
    let start =
        usize::try_from(offset).map_err(|_| corrupted("fcRouteSlip offset exceeds usize"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("lcbRouteSlip length exceeds usize"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("fcRouteSlip/lcbRoute range overflows"))?;
    table
        .get(start..end)
        .ok_or_else(|| corrupted("fcRouteSlip/lcbRoute extends beyond the table stream"))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

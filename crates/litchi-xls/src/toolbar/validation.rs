//! Shared structural validation for the bounded XLS XCB model.

use crate::{Error, Result};

use super::model::APPLICATION_TOOLBAR_ID;
use super::{Control, Toolbar, ToolbarSet, Wrapper};

/// Bound allocations and recursive parsing for hostile XCB streams.
pub(super) const MAX_TOOLBARS: usize = 4096;
pub(super) const MAX_CONTROLS: usize = 1 << 15;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XCB {}", message.into()))
}

pub(super) fn unsupported(message: impl Into<String>) -> Error {
    Error::UnsupportedFeature(format!("XCB {}", message.into()))
}

pub(super) fn validate_toolbar_set(value: &ToolbarSet) -> Result<()> {
    if value.signature() != 0x01 {
        return Err(invalid(format!(
            "CTBS signature must be 0x01, got 0x{:02X}",
            value.signature()
        )));
    }
    if value.version() != 0x01 {
        return Err(invalid(format!(
            "CTBS version must be 0x01, got 0x{:02X}",
            value.version()
        )));
    }
    let count = usize::from(value.toolbar_count());
    if count == 0 {
        return Err(invalid("CTBS ctb must be greater than zero"));
    }
    if count > MAX_TOOLBARS {
        return Err(invalid("CTBS ctb exceeds the bounded limit"));
    }
    if value.view_count() != 0x0003 {
        return Err(invalid("CTBS ctbViews must be 0x0003"));
    }
    if value.active_view() > 0x0001 {
        return Err(invalid("CTBS ictbView must be 0x0000 or 0x0001"));
    }
    Ok(())
}

pub(super) fn validate_control(value: &Control) -> Result<()> {
    if !value.is_active_x() {
        return Err(unsupported(format!(
            "TBCData for control type 0x{:02X} is not decoded",
            value.header().control_type().raw()
        )));
    }
    Ok(())
}

pub(super) fn validate_toolbar(value: &Toolbar<'_>) -> Result<()> {
    let count = value.header().control_count();
    if count < 0 {
        return Err(invalid("TB cCL must not be negative"));
    }
    let count = usize::try_from(count).map_err(|_| invalid("TB cCL overflows usize"))?;
    if count > MAX_CONTROLS {
        return Err(invalid("TB cCL exceeds the bounded limit"));
    }
    if count != value.controls().len() {
        return Err(invalid(format!(
            "CTB control count {} does not match TB cCL {}",
            value.controls().len(),
            count
        )));
    }
    if value.application_id() != APPLICATION_TOOLBAR_ID {
        return Err(invalid(format!(
            "CTB ectbid must be 0x{:08X}, got 0x{:08X}",
            APPLICATION_TOOLBAR_ID,
            value.application_id()
        )));
    }
    for control in value.controls() {
        validate_control(control)?;
    }
    Ok(())
}

pub(super) fn validate_wrapper(value: &Wrapper<'_>) -> Result<()> {
    value.toolbar_set().validate()?;
    if usize::from(value.toolbar_set().toolbar_count()) != value.toolbars().len() {
        return Err(invalid(format!(
            "CTBWRAPPER ctb {} does not match toolbar count {}",
            value.toolbar_set().toolbar_count(),
            value.toolbars().len()
        )));
    }
    for toolbar in value.toolbars() {
        validate_toolbar(toolbar)?;
    }
    Ok(())
}

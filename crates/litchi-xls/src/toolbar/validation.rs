//! Shared structural validation for the bounded XLS XCB model.

use crate::{Error, Result};
use litchi_ole_common::toolbar::ControlHeader;

use super::model::APPLICATION_TOOLBAR_ID;
use super::{Control, Toolbar, ToolbarSet, Wrapper};

/// Bound allocations and recursive parsing for hostile XCB streams.
pub(super) const MAX_TOOLBARS: usize = 4096;
pub(super) const MAX_CONTROLS: usize = 1 << 15;
pub(super) const MAX_BOUNDARY_CANDIDATES: usize = 4096;

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

pub(super) fn command_allowed(header: &ControlHeader) -> bool {
    !matches!(
        header.control_id(),
        0x0001 | 0x06CC | 0x03D8 | 0x03EC | 0x1051
    ) && matches!(
        header.control_type().raw(),
        0x01 | 0x02 | 0x03 | 0x04 | 0x06 | 0x07 | 0x0A | 0x0C | 0x0D | 0x0E | 0x0F | 0x15
    )
}

fn supported_control_type(raw: u8) -> bool {
    matches!(
        raw,
        0x01 | 0x02
            | 0x03
            | 0x04
            | 0x06
            | 0x07
            | 0x09
            | 0x0A
            | 0x0C
            | 0x0D
            | 0x0E
            | 0x0F
            | 0x10
            | 0x12
            | 0x13
            | 0x14
            | 0x15
            | 0x16
    )
}

pub(super) fn validate_control(value: &Control<'_>) -> Result<()> {
    let raw = value.header().control_type().raw();
    if !supported_control_type(raw) {
        return Err(invalid(format!(
            "TBCHeader control type 0x{raw:02X} is not defined"
        )));
    }

    if value.is_active_x() {
        if value.command().is_some() || value.data().is_some() {
            return Err(invalid("ActiveX TBC must not contain TBCCmd or TBCData"));
        }
    } else {
        let data = value.data().ok_or_else(|| {
            unsupported(format!("TBCData is required for control type 0x{raw:02X}"))
        })?;
        if (data.general().flags().save_text() || data.general().flags().save_misc_ui_strings())
            && !value.header().specifics().save_ui_strings()
        {
            return Err(invalid(
                "TBCGeneralInfo UI fields require TBCSFlags fSaveUIStrings",
            ));
        }
        if matches!(raw, 0x07 | 0x0F | 0x12 | 0x13 | 0x15) && !data.specific().is_empty() {
            return Err(invalid(format!(
                "TBCData control-specific information is forbidden for type 0x{raw:02X}"
            )));
        }
        if value.command().is_some() && !command_allowed(value.header()) {
            return Err(invalid(format!(
                "TBCCmd is not permitted for control id 0x{:04X} and type 0x{:02X}",
                value.header().control_id(),
                raw
            )));
        }
    }
    Ok(())
}

pub(super) fn validate_toolbar(value: &Toolbar<'_>) -> Result<()> {
    let count = value.header().control_count();
    if count < 0 {
        return Err(invalid("TB cCL must not be negative"));
    }
    let count = usize::try_from(count).map_err(|_error| invalid("TB cCL overflows usize"))?;
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

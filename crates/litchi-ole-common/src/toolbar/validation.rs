//! Bounded structural validation for toolbar transaction state.

use super::{Body, Control, ControlHeader, ControlType, Error};

/// Maximum serialized size of one common toolbar-control transaction.
pub(crate) const MAX_CONTROL_BYTES: usize = 4 * 1024 * 1024;

/// Maximum opaque host-specific prefix retained before `TBCData`.
pub(crate) const MAX_PREFIX_BYTES: usize = 64 * 1024;

pub(crate) fn validate_authored(value: &Control<'_>) -> Result<(), Error> {
    value.header().validate()?;
    validate_edited(value)
}

pub(crate) fn validate_decoded(value: &Control<'_>) -> Result<(), Error> {
    validate_header_shape(value)?;
    validate_body(value)
}

pub(crate) fn validate_edited(value: &Control<'_>) -> Result<(), Error> {
    validate_header_shape(value)?;
    validate_body(value)
}

fn validate_header_shape(value: &Control<'_>) -> Result<(), Error> {
    if value.prefix().len() > MAX_PREFIX_BYTES {
        return Err(Error::invalid(
            "toolbar control prefix exceeds the bounded limit",
        ));
    }
    let header_len = value.header().to_bytes().len();
    let body_len = value.body().to_bytes().len();
    let total = header_len
        .checked_add(value.prefix().len())
        .and_then(|length| length.checked_add(body_len))
        .ok_or_else(|| Error::invalid("toolbar control size overflows usize"))?;
    if total > MAX_CONTROL_BYTES {
        return Err(Error::invalid("toolbar control exceeds the bounded size"));
    }
    validate_header(value.header())?;
    Ok(())
}

pub(crate) fn validate_header(value: &ControlHeader) -> Result<(), Error> {
    if value.priority() > 7 {
        return Err(Error::invalid("TBCHeader priority exceeds 7"));
    }
    if value.flags().save_dimensions() != value.dimensions().is_some() {
        return Err(Error::invalid("TBCHeader dimensions must match fSaveDxy"));
    }
    Ok(())
}

fn validate_body(value: &Control<'_>) -> Result<(), Error> {
    match value.body() {
        Body::Empty => {
            if !matches!(
                value.header().control_type(),
                ControlType::ActiveX | ControlType::Unknown(_)
            ) {
                return Err(Error::invalid(
                    "supported non-ActiveX toolbar controls require a body",
                ));
            }
        },
        Body::Data(data) => {
            if matches!(
                value.header().control_type(),
                ControlType::ActiveX | ControlType::Unknown(_)
            ) {
                return Err(Error::invalid(
                    "unsupported toolbar controls cannot use decoded TBCData",
                ));
            }
            if (data.general().flags().save_text() || data.general().flags().save_misc_ui_strings())
                && !value.header().specifics().save_ui_strings()
            {
                return Err(Error::invalid(
                    "TBCGeneralInfo UI fields require fSaveUIStrings",
                ));
            }
        },
        Body::Opaque(_) => {},
    }
    Ok(())
}

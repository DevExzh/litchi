//! Semantic invariants for typed DXF and `XFProps` values.

use super::{DifferentialFormat, HorizontalAlignment, Result, XfProperties, XfProperty, invalid};

pub(crate) fn validate_unit_interval(value: f64, field: &str) -> Result<()> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(invalid(format!(
            "gradient {field} coordinate must be between 0.0 and 1.0"
        )));
    }
    Ok(())
}

pub(crate) fn validate_properties(value: &XfProperties) -> Result<()> {
    const MAX_XF_PROPERTIES: usize = 2_048;
    if value.properties.len() > MAX_XF_PROPERTIES || value.properties.len() > u16::MAX as usize {
        return Err(invalid(format!(
            "XFProps count exceeds resource cap {MAX_XF_PROPERTIES}"
        )));
    }
    let has_pattern = value
        .properties
        .iter()
        .any(|property| matches!(property, XfProperty::FillPattern(_)));
    let has_gradient = value.properties.iter().any(|property| {
        matches!(
            property,
            XfProperty::Gradient(_) | XfProperty::GradientStop(_)
        )
    });
    if has_pattern && has_gradient {
        return Err(invalid(
            "XFProps cannot combine pattern-fill and gradient properties",
        ));
    }
    let mut preceding_gradient = false;
    let mut distributed = false;
    let mut horizontal_distributed = false;
    for property in &value.properties {
        match property {
            XfProperty::Gradient(_) => preceding_gradient = true,
            XfProperty::GradientStop(_) if !preceding_gradient => {
                return Err(invalid("gradient stop has no preceding gradient property"));
            },
            XfProperty::JustifyDistributed(true) => distributed = true,
            XfProperty::HorizontalAlignment(Some(HorizontalAlignment::Distributed)) => {
                horizontal_distributed = true;
            },
            _ => {},
        }
        property.data_bytes()?;
    }
    if distributed && !horizontal_distributed {
        return Err(invalid(
            "justify-distributed requires distributed horizontal alignment",
        ));
    }
    Ok(())
}

pub(crate) fn validate_format(value: &DifferentialFormat) -> Result<()> {
    validate_properties(&value.properties)?;
    if !value.new_border
        && value.properties.properties.iter().any(|property| {
            matches!(
                property,
                XfProperty::VerticalBorder(_) | XfProperty::HorizontalBorder(_)
            )
        })
    {
        return Err(invalid("internal border properties require fNewBorder"));
    }
    Ok(())
}

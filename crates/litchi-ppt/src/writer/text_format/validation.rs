//! Validation for the bounded PowerPoint text-format wire model.
//!
//! These checks are kept beside the codec so serialization has one source of
//! truth for the limits defined by [MS-PPT] section 2.9.

use super::semantic::TextColor;

pub(super) fn validate_indent_level(level: u16) -> std::io::Result<()> {
    if level > 4 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT paragraph indent level must be between 0 and 4",
        ));
    }
    Ok(())
}

pub(super) fn validate_bullet_size(size: i16) -> std::io::Result<()> {
    if !((25..=400).contains(&size) || (-4000..=-1).contains(&size)) {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT bullet size must be 25..=400 percent or -4000..=-1 points",
        ));
    }
    Ok(())
}

pub(super) fn validate_bullet_color(color: TextColor) -> std::io::Result<()> {
    if color.use_scheme && color.scheme_index > 7 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT bullet color-scheme index must be between 0 and 7",
        ));
    }
    Ok(())
}

pub(super) fn validate_font_size(size: u16) -> std::io::Result<()> {
    if !(1..=4000).contains(&size) {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT font size must be between 1 and 4000 points",
        ));
    }
    Ok(())
}

pub(super) fn validate_run_color(color: TextColor) -> std::io::Result<()> {
    if color.use_scheme && color.scheme_index > 7 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT color-scheme index must be between 0 and 7",
        ));
    }
    Ok(())
}

pub(super) fn validate_baseline_position(position: i16) -> std::io::Result<()> {
    if !(-100..=100).contains(&position) {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT baseline position must be between -100 and 100 percent",
        ));
    }
    Ok(())
}

pub(super) fn validate_pp9_run_id(id: u8) -> std::io::Result<()> {
    if id > 0x0F {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT pp9 run grouping identifier must be between 0 and 15",
        ));
    }
    Ok(())
}

pub(super) fn validate_style_mask(mask: u16) -> std::io::Result<()> {
    if mask & !0x3FB7 != 0 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "PPT character style specified mask contains reserved bits",
        ));
    }
    Ok(())
}

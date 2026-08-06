//! Terminal checks and asset attachment for sheet traversal.

use super::super::super::{Error, Result, Sheet};

pub(super) fn finish(
    xml_content: &str,
    mut sheets: Vec<Sheet>,
    sheet_dde_source_depth: Option<usize>,
    conditional_formats_depth: Option<usize>,
    sparkline_groups_depth: Option<usize>,
) -> Result<Vec<Sheet>> {
    if sheet_dde_source_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated office:dde-source".to_string(),
        ));
    }
    if conditional_formats_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated calcext:conditional-formats".to_string(),
        ));
    }
    if sparkline_groups_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated calcext:sparkline-groups".to_string(),
        ));
    }

    super::super::super::super::super::package::attach_content_assets(xml_content, &mut sheets)?;
    Ok(sheets)
}

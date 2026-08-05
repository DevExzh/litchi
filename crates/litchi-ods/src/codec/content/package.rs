//! Content-to-package boundary services.
//!
//! This owner does not assemble or rewrite an ODS archive. It only joins
//! package-visible content XML assets (sheet images and drawing shapes) to the
//! semantic sheets produced by the streaming traversal. Archive ownership
//! remains in `crate::package`, as required by the ODF family split.

use super::super::{Sheet, shape, sheet_image};
use litchi_core::{Error, Result};

pub(super) fn attach_content_assets(xml_content: &str, sheets: &mut [Sheet]) -> Result<()> {
    let images = crate::media::scan_content_images(xml_content)?;
    let mut sheet_indices = std::collections::HashMap::with_capacity(sheets.len());
    for (index, sheet) in sheets.iter().enumerate() {
        if sheet_indices.insert(sheet.name.clone(), index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate table name '{}' prevents sheet-image association",
                sheet.name
            )));
        }
    }
    for image in images {
        let Some(frame) = image.frame.as_ref().filter(|frame| frame.sheet_shape) else {
            continue;
        };
        let sheet_name = frame.sheet_name.as_deref().ok_or_else(|| {
            Error::InvalidFormat("sheet image has no containing table name".to_string())
        })?;
        let index = *sheet_indices.get(sheet_name).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "sheet image references unknown table '{sheet_name}'"
            ))
        })?;
        sheet_image::validate_sheet_image(&image)?;
        sheets[index].images.push(image);
    }
    for sheet in sheets.iter() {
        sheet_image::validate_sheet_images(&sheet.images)?;
    }

    let shape_tables = crate::odp::OdpParser::parse_sheet_shape_tables(xml_content)?;
    if shape_tables.len() != sheets.len() {
        return Err(Error::InvalidFormat(format!(
            "spreadsheet table structure changed during shape parsing: {} shape container(s) for {} table(s)",
            shape_tables.len(),
            sheets.len()
        )));
    }
    for (sheet, shapes) in sheets.iter_mut().zip(shape_tables) {
        for shape in shapes {
            if let Some(sheet_shape) = shape::sheet_shape_from_parsed(shape)? {
                sheet.shapes.push(sheet_shape);
            }
        }
        shape::validate_sheet_shapes(&sheet.shapes)?;
    }
    Ok(())
}

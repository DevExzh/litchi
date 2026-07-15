//! Typed Pages document page-layout CRUD.

use super::*;

const PAGE_WIDTH_FIELD: u32 = 30;
const PAGE_HEIGHT_FIELD: u32 = 31;
const LEFT_MARGIN_FIELD: u32 = 32;
const RIGHT_MARGIN_FIELD: u32 = 33;
const TOP_MARGIN_FIELD: u32 = 34;
const BOTTOM_MARGIN_FIELD: u32 = 35;
const HEADER_MARGIN_FIELD: u32 = 36;
const FOOTER_MARGIN_FIELD: u32 = 37;
const PAGE_SCALE_FIELD: u32 = 38;
const VERTICAL_BODY_LAYOUT_FIELD: u32 = 39;
const ORIENTATION_FIELD: u32 = 42;

impl PagesEditor {
    /// Read the page geometry fields from the Pages document root.
    pub fn page_layout(&self) -> Result<PagesPageLayout> {
        Ok(PagesPageLayout::from(&root_document(self.text.package())?))
    }

    /// Replace the page geometry fields transactionally.
    pub fn set_page_layout(&mut self, layout: PagesPageLayout) -> Result<()> {
        validate_page_layout(&layout)?;
        let mut staged = self.text.package().clone();
        staged.update_archive(DOCUMENT_ARCHIVE_NAME, |archive| {
            let object = archive.object_mut(DOCUMENT_OBJECT_ID).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages root object {DOCUMENT_OBJECT_ID} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
                .ok_or_else(|| {
                    Error::InvalidFormat("Pages root has no TP.DocumentArchive payload".to_owned())
                })?;
            let original = &object.messages[message_index].data;
            let document = DocumentArchive::decode(original.as_slice())?;
            let mut data = original.clone();
            for (field_number, current, replacement) in [
                (PAGE_WIDTH_FIELD, document.page_width, layout.page_width),
                (PAGE_HEIGHT_FIELD, document.page_height, layout.page_height),
                (LEFT_MARGIN_FIELD, document.left_margin, layout.left_margin),
                (
                    RIGHT_MARGIN_FIELD,
                    document.right_margin,
                    layout.right_margin,
                ),
                (TOP_MARGIN_FIELD, document.top_margin, layout.top_margin),
                (
                    BOTTOM_MARGIN_FIELD,
                    document.bottom_margin,
                    layout.bottom_margin,
                ),
                (
                    HEADER_MARGIN_FIELD,
                    document.header_margin,
                    layout.header_margin,
                ),
                (
                    FOOTER_MARGIN_FIELD,
                    document.footer_margin,
                    layout.footer_margin,
                ),
                (PAGE_SCALE_FIELD, document.page_scale, layout.page_scale),
            ] {
                data = patch_fixed32_field(
                    &data,
                    field_number,
                    current.is_some(),
                    replacement.map(f32::to_bits),
                )?;
            }
            data = patch_varint_field(
                &data,
                VERTICAL_BODY_LAYOUT_FIELD,
                document.lays_out_body_vertically.is_some(),
                layout.lays_out_body_vertically.map(u64::from),
            )?;
            data = patch_varint_field(
                &data,
                ORIENTATION_FIELD,
                document.orientation.is_some(),
                layout
                    .orientation
                    .map(PagesPageOrientation::as_raw)
                    .map(u64::from),
            )?;
            let verified = DocumentArchive::decode(data.as_slice())?;
            if PagesPageLayout::from(&verified) != layout {
                return Err(Error::InvalidFormat(
                    "Pages page-layout wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        *self = Self::from_package(staged)?;
        Ok(())
    }
}

fn validate_page_layout(layout: &PagesPageLayout) -> Result<()> {
    if layout
        .orientation
        .is_some_and(|orientation| !orientation.is_canonical())
    {
        return Err(Error::ParseError(
            "Pages unknown orientation must not alias a known orientation".to_owned(),
        ));
    }
    for (name, value) in [
        ("page width", layout.page_width),
        ("page height", layout.page_height),
        ("page scale", layout.page_scale),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value <= 0.0) {
            return Err(Error::ParseError(format!(
                "Pages {name} must be finite and greater than zero"
            )));
        }
    }
    for (name, value) in [
        ("left margin", layout.left_margin),
        ("right margin", layout.right_margin),
        ("top margin", layout.top_margin),
        ("bottom margin", layout.bottom_margin),
        ("header margin", layout.header_margin),
        ("footer margin", layout.footer_margin),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value < 0.0) {
            return Err(Error::ParseError(format!(
                "Pages {name} must be finite and non-negative"
            )));
        }
    }
    if let (Some(width), Some(left), Some(right)) =
        (layout.page_width, layout.left_margin, layout.right_margin)
        && left + right >= width
    {
        return Err(Error::ParseError(
            "Pages horizontal margins must leave positive body width".to_owned(),
        ));
    }
    if let (Some(height), Some(top), Some(bottom)) =
        (layout.page_height, layout.top_margin, layout.bottom_margin)
        && top + bottom >= height
    {
        return Err(Error::ParseError(
            "Pages vertical margins must leave positive body height".to_owned(),
        ));
    }
    Ok(())
}

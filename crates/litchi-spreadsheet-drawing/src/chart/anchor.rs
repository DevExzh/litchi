//! Worksheet cell-anchor vocabulary and drawing-part emission.

use super::model::Chart;
use crate::{Error, Result};

const WORKSHEET_ROW_COUNT: u32 = 1_048_576;
const WORKSHEET_COLUMN_COUNT: u32 = 16_384;

/// Chart anchor position in a worksheet.
///
/// Specifies the position and size of a chart using cell anchors and offsets.
#[derive(Debug, Clone)]
pub struct Anchor {
    /// Starting column (0-based)
    pub from_col: u32,
    /// Offset from the left edge of `from_col` (in EMUs)
    pub from_col_offset: i64,
    /// Starting row (0-based)
    pub from_row: u32,
    /// Offset from the top edge of `from_row` (in EMUs)
    pub from_row_offset: i64,
    /// Ending column (0-based)
    pub to_col: u32,
    /// Offset from the left edge of `to_col` (in EMUs)
    pub to_col_offset: i64,
    /// Ending row (0-based)
    pub to_row: u32,
    /// Offset from the top edge of `to_row` (in EMUs)
    pub to_row_offset: i64,
}

impl Anchor {
    /// Create a new chart anchor from cell positions.
    ///
    /// # Arguments
    ///
    /// * `from_col` - Starting column (0-based)
    /// * `from_row` - Starting row (0-based)
    /// * `to_col` - Ending column (0-based)
    /// * `to_row` - Ending row (0-based)
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Chart spanning from B2 to H15
    /// let anchor = Anchor::new(1, 1, 7, 14);
    /// ```
    #[must_use]
    pub fn new(from_col: u32, from_row: u32, to_col: u32, to_row: u32) -> Self {
        Self {
            from_col,
            from_col_offset: 0,
            from_row,
            from_row_offset: 0,
            to_col,
            to_col_offset: 0,
            to_row,
            to_row_offset: 0,
        }
    }

    /// Create a chart anchor with precise offsets.
    #[allow(
        clippy::too_many_arguments,
        reason = "The OOXML two-cell anchor has four coordinates and four offsets."
    )]
    #[must_use]
    pub fn with_offsets(
        from_col: u32,
        from_col_offset: i64,
        from_row: u32,
        from_row_offset: i64,
        to_col: u32,
        to_col_offset: i64,
        to_row: u32,
        to_row_offset: i64,
    ) -> Self {
        Self {
            from_col,
            from_col_offset,
            from_row,
            from_row_offset,
            to_col,
            to_col_offset,
            to_row,
            to_row_offset,
        }
    }
}

impl Default for Anchor {
    fn default() -> Self {
        Self::new(0, 0, 10, 15)
    }
}

/// # Errors
///
/// Returns an error when the anchor is descending, out of worksheet bounds, or has a negative offset.
pub fn validate(anchor: &Anchor) -> Result<()> {
    if anchor.to_row < anchor.from_row || anchor.to_col < anchor.from_col {
        return Err(Error::Invalid(
            "chart anchor cannot be descending".to_string(),
        ));
    }
    if anchor.to_row >= WORKSHEET_ROW_COUNT || anchor.to_col >= WORKSHEET_COLUMN_COUNT {
        return Err(Error::Invalid(
            "chart anchor exceeds worksheet bounds".to_string(),
        ));
    }
    if [
        anchor.from_col_offset,
        anchor.from_row_offset,
        anchor.to_col_offset,
        anchor.to_row_offset,
    ]
    .iter()
    .any(|offset| *offset < 0)
    {
        return Err(Error::Invalid(
            "chart anchor offsets cannot be negative".to_string(),
        ));
    }
    Ok(())
}

/// # Errors
///
/// Returns an error when a chart anchor is invalid, an identifier overflows, or XML formatting fails.
pub fn write_all(
    xml: &mut String,
    charts: &[Chart],
    object_id_offset: usize,
    relationship_id_offset: usize,
) -> Result<()> {
    use std::fmt::Write as _;

    for (index, chart) in charts.iter().enumerate() {
        validate(&chart.anchor)?;
        let object_id = object_id_offset
            .checked_add(index)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| Error::Invalid("chart object ID overflow".to_string()))?;
        let relationship_id = relationship_id_offset
            .checked_add(index)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| Error::Invalid("chart relationship ID overflow".to_string()))?;
        let anchor = &chart.anchor;
        xml.push_str("<xdr:twoCellAnchor>");
        write!(
            xml,
            "<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:from>",
            anchor.from_col,
            anchor.from_col_offset,
            anchor.from_row,
            anchor.from_row_offset
        )
        .map_err(|error| Error::Encoding(error.to_string()))?;
        write!(
            xml,
            "<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>",
            anchor.to_col,
            anchor.to_col_offset,
            anchor.to_row,
            anchor.to_row_offset
        )
        .map_err(|error| Error::Encoding(error.to_string()))?;
        write!(
            xml,
            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{object_id}" name="Chart {}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
            index + 1
        )
        .map_err(|error| Error::Encoding(error.to_string()))?;
        xml.push_str(
            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#,
        );
        write!(
            xml,
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{relationship_id}"/>"#
        )
        .map_err(|error| Error::Encoding(error.to_string()))?;
        xml.push_str(
            "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>",
        );
    }
    Ok(())
}

//! Dependency-free table appearance vocabulary.
//!
//! Native style inheritance, protobuf fields, and archive identifiers stay in
//! the concrete iWork format owners. This module contains only the compact
//! semantic value exchanged at that boundary.

/// Whether alternating body-row fills are enabled.
#[repr(u8)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Banding {
    /// Use one fill for adjacent body rows.
    #[default]
    Disabled,
    /// Apply the table style's alternating-row fill.
    Enabled,
}

/// Whether row heights are fixed or fit their cell contents.
#[repr(u8)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum RowSizing {
    /// Preserve explicit or style-provided row heights.
    #[default]
    Fixed,
    /// Automatically expand rows to fit their cell contents.
    FitCellContents,
}

/// Whether one family of table gridlines is drawn.
#[repr(u8)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum GridlineVisibility {
    /// Do not draw this gridline family.
    Hidden,
    /// Draw this gridline family using the table style's strokes.
    #[default]
    Visible,
}

/// Gridline visibility for each table region.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct Gridlines {
    /// Horizontal lines between body rows, excluding header-column lines.
    pub body_horizontal: GridlineVisibility,
    /// Horizontal lines between rows inside the header-column region.
    pub header_columns_horizontal: GridlineVisibility,
    /// Vertical lines between body columns, excluding header-row and footer lines.
    pub body_vertical: GridlineVisibility,
    /// Vertical lines between columns inside the header-row region.
    pub header_rows_vertical: GridlineVisibility,
    /// Vertical lines between columns inside the footer-row region.
    pub footer_rows_vertical: GridlineVisibility,
}

/// Effective appearance settings for one native table.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct Appearance {
    /// Alternating body-row fill behavior.
    pub row_banding: Banding,
    /// Automatic row-height behavior.
    pub row_sizing: RowSizing,
    /// Horizontal and vertical gridline visibility by table region.
    pub gridlines: Gridlines,
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing};

    #[test]
    fn appearance_is_compact_and_defaults_match_native_defaults() {
        assert_eq!(size_of::<Banding>(), 1);
        assert_eq!(size_of::<RowSizing>(), 1);
        assert_eq!(size_of::<GridlineVisibility>(), 1);
        assert_eq!(size_of::<Gridlines>(), 5);
        assert_eq!(size_of::<Appearance>(), 7);

        let appearance = Appearance::default();
        assert_eq!(appearance.row_banding, Banding::Disabled);
        assert_eq!(appearance.row_sizing, RowSizing::Fixed);
        assert_eq!(
            appearance.gridlines,
            Gridlines {
                body_horizontal: GridlineVisibility::Visible,
                header_columns_horizontal: GridlineVisibility::Visible,
                body_vertical: GridlineVisibility::Visible,
                header_rows_vertical: GridlineVisibility::Visible,
                footer_rows_vertical: GridlineVisibility::Visible,
            }
        );
    }
}

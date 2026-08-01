use super::shape::{
    XlsShapeAnchor, XlsShapeColor, XlsShapeFill, XlsShapeKind, XlsShapeLine, XlsShapeText,
    validate_shape_style,
};
use crate::{XlsError, XlsResult};

/// Rectangle expressed in a group's child coordinate space (MS-ODRAW OfficeArtFSPGR /
/// OfficeArtChildAnchor).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsGroupRect {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl XlsGroupRect {
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    pub(crate) fn validate(self) -> XlsResult<()> {
        if self.left >= self.right || self.top >= self.bottom {
            return Err(XlsError::InvalidData(
                "group rectangle must have left < right and top < bottom".to_string(),
            ));
        }
        Ok(())
    }
}

/// Writable primitive that lives inside a shape group and is anchored in the
/// group's child coordinate space.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsShapeGroupChild {
    pub kind: XlsShapeKind,
    /// Child anchor in the group coordinate space declared by the group.
    pub anchor: XlsGroupRect,
    /// Optional requested OBJ identifier. `None` assigns the first free canonical ID.
    pub object_id: Option<u16>,
    pub text: Option<XlsShapeText>,
    pub fill: XlsShapeFill,
    pub line: XlsShapeLine,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeGroupChild {
    pub fn new(kind: XlsShapeKind, anchor: XlsGroupRect) -> Self {
        Self {
            kind,
            anchor,
            object_id: None,
            text: None,
            fill: XlsShapeFill::Solid(XlsShapeColor::rgb(255, 255, 255)),
            line: XlsShapeLine::Solid {
                color: XlsShapeColor::rgb(0, 0, 0),
                width_emu: 12_700,
            },
            visible: true,
            locked: true,
        }
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
        self.anchor.validate()?;
        validate_shape_style(
            self.kind,
            self.object_id,
            self.fill,
            self.line,
            self.text.as_ref(),
        )
    }
}

/// Writable, macro-inert BIFF8 shape group (OfficeArt SpgrContainer).
///
/// The group itself is anchored to worksheet cells while every child is anchored
/// inside [`XlsShapeGroupWrite::coordinates`], the group coordinate space.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsShapeGroupWrite {
    /// Cell-relative anchor of the whole group.
    pub anchor: XlsShapeAnchor,
    /// Coordinate space that child anchors are expressed in (OfficeArtFSPGR).
    pub coordinates: XlsGroupRect,
    /// Optional requested OBJ identifier for the group. `None` assigns the first
    /// free canonical ID.
    pub object_id: Option<u16>,
    pub children: Vec<XlsShapeGroupChild>,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeGroupWrite {
    pub fn new(anchor: XlsShapeAnchor) -> Self {
        Self {
            anchor,
            coordinates: XlsGroupRect::new(0, 0, 1023, 255),
            object_id: None,
            children: Vec::new(),
            visible: true,
            locked: true,
        }
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
        self.anchor.validate()?;
        self.coordinates.validate()?;
        if matches!(self.object_id, Some(0 | u16::MAX)) {
            return Err(XlsError::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        if self.children.is_empty() {
            return Err(XlsError::InvalidData(
                "a shape group must contain at least one child shape".to_string(),
            ));
        }
        let mut requested = self.object_id.into_iter().collect::<Vec<_>>();
        for child in &self.children {
            child.validate()?;
            requested.extend(child.object_id);
        }
        requested.sort_unstable();
        if requested.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(XlsError::InvalidData(
                "shape group requests the same object ID more than once".to_string(),
            ));
        }
        Ok(())
    }

    /// Object IDs consumed by this group: one for the group plus one per child.
    pub(crate) fn object_count(&self) -> usize {
        1 + self.children.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn anchor() -> XlsShapeAnchor {
        XlsShapeAnchor {
            move_with_cells: true,
            size_with_cells: true,
            first_column: 0,
            first_column_offset: 0,
            first_row: 0,
            first_row_offset: 0,
            last_column: 5,
            last_column_offset: 0,
            last_row: 8,
            last_row_offset: 0,
        }
    }

    #[test]
    fn empty_group_is_rejected() {
        let group = XlsShapeGroupWrite::new(anchor());
        assert!(group.validate().is_err());
    }

    #[test]
    fn degenerate_rectangles_are_rejected() {
        assert!(XlsGroupRect::new(0, 0, 0, 10).validate().is_err());
        assert!(XlsGroupRect::new(0, 10, 10, 10).validate().is_err());
        assert!(XlsGroupRect::new(-5, -5, 5, 5).validate().is_ok());

        let mut group = XlsShapeGroupWrite::new(anchor());
        group.children.push(XlsShapeGroupChild::new(
            XlsShapeKind::Rectangle,
            XlsGroupRect::new(20, 20, 20, 40),
        ));
        assert!(group.validate().is_err());
    }

    #[test]
    fn duplicate_requested_ids_inside_one_group_are_rejected() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        group.object_id = Some(4);
        let mut child =
            XlsShapeGroupChild::new(XlsShapeKind::Ellipse, XlsGroupRect::new(0, 0, 100, 100));
        child.object_id = Some(4);
        group.children.push(child);
        assert!(group.validate().is_err());

        group.children[0].object_id = Some(5);
        assert!(group.validate().is_ok());
    }

    #[test]
    fn grouped_line_children_reject_fill_and_text() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        let mut line =
            XlsShapeGroupChild::new(XlsShapeKind::Line, XlsGroupRect::new(0, 0, 200, 100));
        line.text = Some(XlsShapeText::new("no text on lines"));
        group.children.push(line);
        assert!(group.validate().is_err());

        group.children[0].text = None;
        group.children[0].fill = XlsShapeFill::None;
        assert!(group.validate().is_ok());
    }

    #[test]
    fn group_object_count_includes_group_marker() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        for _ in 0..3 {
            group.children.push(XlsShapeGroupChild::new(
                XlsShapeKind::Rectangle,
                XlsGroupRect::new(0, 0, 10, 10),
            ));
        }
        assert_eq!(group.object_count(), 4);
    }
}

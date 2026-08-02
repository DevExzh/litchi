use super::shape::{
    Anchor, Rect, XlsShapeColor, XlsShapeFill, XlsShapeKind, XlsShapeLine, XlsShapeText,
    validate_shape_style,
};
use crate::{XlsError, XlsResult};

/// Writable primitive that lives inside a shape group and is anchored in the
/// group's child coordinate space.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsShapeGroupChild {
    pub kind: XlsShapeKind,
    /// Child anchor in the group coordinate space declared by the group.
    pub anchor: Rect,
    /// Optional requested OBJ identifier. `None` assigns the first free canonical ID.
    pub object_id: Option<u16>,
    pub text: Option<XlsShapeText>,
    pub fill: XlsShapeFill,
    pub line: XlsShapeLine,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeGroupChild {
    pub fn new(kind: XlsShapeKind, anchor: Rect) -> Self {
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
    pub anchor: Anchor,
    /// Coordinate space that child anchors are expressed in (OfficeArtFSPGR).
    pub coordinates: Rect,
    /// Optional requested OBJ identifier for the group. `None` assigns the first
    /// free canonical ID.
    pub object_id: Option<u16>,
    pub children: Vec<XlsShapeGroupChild>,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeGroupWrite {
    pub fn new(anchor: Anchor) -> Self {
        Self {
            anchor,
            coordinates: Rect::DEFAULT_GROUP,
            object_id: None,
            children: Vec::new(),
            visible: true,
            locked: true,
        }
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
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
        let requested_capacity = self
            .children
            .len()
            .checked_add(usize::from(self.object_id.is_some()))
            .ok_or(XlsError::Allocation(
                "computing the shape-group object-ID count",
            ))?;
        let mut requested = Vec::new();
        requested
            .try_reserve_exact(requested_capacity)
            .map_err(|_| {
                XlsError::Allocation("reserving shape group object-ID validation storage")
            })?;
        if let Some(object_id) = self.object_id {
            requested.push(object_id);
        }
        for child in &self.children {
            child.validate()?;
            if let Some(object_id) = child.object_id {
                requested.push(object_id);
            }
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
    pub(crate) fn object_count(&self) -> XlsResult<usize> {
        self.children
            .len()
            .checked_add(1)
            .ok_or(XlsError::Allocation(
                "computing the shape-group object count",
            ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::panic::catch_unwind;

    fn anchor() -> Anchor {
        Anchor::cells(0, 0, 8, 5, super::super::shape::Behavior::MoveAndSize).unwrap()
    }

    #[test]
    fn empty_group_is_rejected() {
        let group = XlsShapeGroupWrite::new(anchor());
        assert!(group.validate().is_err());
    }

    #[test]
    fn degenerate_rectangles_are_rejected() {
        assert!(Rect::new(0, 0, 0, 10).is_err());
        assert!(Rect::new(0, 10, 10, 10).is_err());
        assert!(Rect::new(-5, -5, 5, 5).is_ok());
    }

    #[test]
    fn duplicate_requested_ids_inside_one_group_are_rejected() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        group.object_id = Some(4);
        let mut child =
            XlsShapeGroupChild::new(XlsShapeKind::Ellipse, Rect::new(0, 0, 100, 100).unwrap());
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
            XlsShapeGroupChild::new(XlsShapeKind::Line, Rect::new(0, 0, 200, 100).unwrap());
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
                Rect::new(0, 0, 10, 10).unwrap(),
            ));
        }
        assert_eq!(group.object_count().unwrap(), 4);
    }

    #[test]
    fn requested_id_validation_does_not_unwind_for_large_groups() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        for object_id in 1..=1_024 {
            let mut child =
                XlsShapeGroupChild::new(XlsShapeKind::Rectangle, Rect::new(0, 0, 10, 10).unwrap());
            child.object_id = Some(object_id);
            group.children.push(child);
        }
        let outcome = catch_unwind(|| group.validate());
        assert!(matches!(outcome, Ok(Ok(()))));
    }
}

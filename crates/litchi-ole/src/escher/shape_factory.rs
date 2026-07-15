//! High-performance shape factory for creating shapes from Escher records.
//!
//! # Performance
//!
//! - Zero-copy shape data access
//! - Iterator-based shape enumeration
//! - Pattern matching for shape type detection

use super::container::EscherContainer;
use super::record::Result;
use super::shape::EscherShape;
use super::types::EscherRecordType;

/// Factory for creating shapes from Escher records.
pub struct EscherShapeFactory;

impl EscherShapeFactory {
    /// Extract all shapes from an Escher/Drawing data.
    ///
    /// # Performance
    ///
    /// - Depth-first traversal
    /// - Pre-allocated results vector
    /// - Short-circuits on errors
    pub fn extract_shapes_from_drawing(data: &[u8]) -> Result<Vec<EscherShape<'_>>> {
        let parser = super::parser::EscherParser::new(data);

        let mut shapes = Vec::new();

        if let Some(root_result) = parser.root_container() {
            let root = root_result?;
            if root.record().record_type == EscherRecordType::SpgrContainer {
                Self::extract_shapes_from_root_group(&root, &mut shapes);
            } else {
                Self::extract_shapes_from_container(&root, &mut shapes);
            }
        }

        Ok(shapes)
    }

    /// Recursively extract shapes from a container.
    ///
    /// # Implementation Notes
    ///
    /// Follows Apache POI's logic:
    /// - SpContainer: Single shape, add to results
    /// - SpgrContainer: Group shape, process with skip-first logic
    fn extract_shapes_from_container<'data>(
        container: &EscherContainer<'data>,
        shapes: &mut Vec<EscherShape<'data>>,
    ) {
        // A drawing container can also hold a background SpContainer. Public
        // sheet shapes are the user children of its root SpgrContainer only.
        if container.record().record_type == EscherRecordType::DgContainer
            && let Some(root_group) = container.find_child(EscherRecordType::SpgrContainer)
        {
            Self::extract_shapes_from_root_group(&EscherContainer::new(root_group), shapes);
            return;
        }

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    let sp_container = EscherContainer::new(child);
                    let shape = EscherShape::from_container(sp_container);
                    shapes.push(shape);
                },
                EscherRecordType::SpgrContainer => {
                    let group_container = EscherContainer::new(child);
                    shapes.push(EscherShape::from_container(group_container));
                },
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    Self::extract_shapes_from_container(&child_container, shapes);
                },
                _ => {},
            }
        }
    }

    /// Extract user shapes from the root SpgrContainer.
    ///
    /// Based on Apache POI's HSLFGroupShape.getShapes():
    /// - The first SpContainer in SpgrContainer is the group shape itself
    /// - Remaining SpContainer children are the actual child shapes
    fn extract_shapes_from_root_group<'data>(
        container: &EscherContainer<'data>,
        shapes: &mut Vec<EscherShape<'data>>,
    ) {
        let mut is_first = true;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    if is_first {
                        is_first = false;
                    } else {
                        let sp_container = EscherContainer::new(child);
                        let child_shape = EscherShape::from_container(sp_container);
                        shapes.push(child_shape);
                    }
                },
                EscherRecordType::SpgrContainer => {
                    let nested_group = EscherContainer::new(child);
                    shapes.push(EscherShape::from_container(nested_group));
                },
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    Self::extract_shapes_from_container(&child_container, shapes);
                },
                _ => {},
            }
        }
    }

    /// Count shapes in drawing data (without full parsing).
    ///
    /// # Performance
    ///
    /// - Counts only shapes exposed by extraction
    /// - No shape object allocation
    /// - Early termination on errors
    pub fn count_shapes_in_drawing(data: &[u8]) -> usize {
        let parser = super::parser::EscherParser::new(data);

        if let Some(root_result) = parser.root_container()
            && let Ok(root) = root_result
        {
            return if root.record().record_type == EscherRecordType::SpgrContainer {
                Self::count_shapes_in_root_group(&root)
            } else {
                Self::count_shapes_in_container(&root)
            };
        }

        0
    }

    /// Recursively count shapes in a container.
    fn count_shapes_in_container(container: &EscherContainer<'_>) -> usize {
        if container.record().record_type == EscherRecordType::DgContainer
            && let Some(root_group) = container.find_child(EscherRecordType::SpgrContainer)
        {
            return Self::count_shapes_in_root_group(&EscherContainer::new(root_group));
        }

        let mut count = 0;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    count += 1;
                },
                EscherRecordType::SpgrContainer => {
                    count += 1;
                },
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    count += Self::count_shapes_in_container(&child_container);
                },
                _ => {},
            }
        }

        count
    }

    /// Count the top-level user shapes in a drawing group without allocating.
    fn count_shapes_in_root_group(container: &EscherContainer<'_>) -> usize {
        let mut count = 0;
        let mut is_first = true;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer if is_first => {
                    is_first = false;
                },
                EscherRecordType::SpContainer | EscherRecordType::SpgrContainer => {
                    count += 1;
                },
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    count += Self::count_shapes_in_container(&child_container);
                },
                _ => {},
            }
        }

        count
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::escher::writer::{
        ShapeBuilder, record_type, write_atom, write_child_anchor, write_client_anchor,
        write_container, write_spgr,
    };

    fn shape_container(shape_type: u16, shape_id: u32, child_anchor: bool) -> Vec<u8> {
        let mut children = Vec::new();
        ShapeBuilder::new(shape_type, shape_id)
            .write(&mut children)
            .unwrap();
        if child_anchor {
            write_child_anchor(&mut children, 10, 20, 110, 70).unwrap();
        } else {
            write_client_anchor(&mut children, 10, 20, 110, 70).unwrap();
        }

        let mut container = Vec::new();
        write_container(&mut container, 0, record_type::SP_CONTAINER, &children).unwrap();
        container
    }

    fn grouped_drawing() -> Vec<u8> {
        let patriarch = shape_container(0, 1, false);
        let rectangle = shape_container(1, 2, false);

        let mut group_header_children = Vec::new();
        write_spgr(&mut group_header_children, 0, 0, 1000, 500).unwrap();
        ShapeBuilder::new(0, 3)
            .write(&mut group_header_children)
            .unwrap();
        write_child_anchor(&mut group_header_children, 100, 200, 500, 400).unwrap();
        let mut group_header = Vec::new();
        write_container(
            &mut group_header,
            0,
            record_type::SP_CONTAINER,
            &group_header_children,
        )
        .unwrap();

        let ellipse = shape_container(3, 4, true);
        let mut nested_group_children = group_header;
        nested_group_children.extend_from_slice(&ellipse);
        let mut nested_group = Vec::new();
        write_container(
            &mut nested_group,
            0,
            record_type::SPGR_CONTAINER,
            &nested_group_children,
        )
        .unwrap();

        let mut root_group_children = patriarch;
        root_group_children.extend_from_slice(&rectangle);
        root_group_children.extend_from_slice(&nested_group);
        let mut root_group = Vec::new();
        write_container(
            &mut root_group,
            0,
            record_type::SPGR_CONTAINER,
            &root_group_children,
        )
        .unwrap();

        let mut drawing = Vec::new();
        let background = shape_container(1, 5, false);
        let mut drawing_children = root_group;
        drawing_children.extend_from_slice(&background);
        write_container(
            &mut drawing,
            0,
            record_type::DG_CONTAINER,
            &drawing_children,
        )
        .unwrap();
        drawing
    }

    #[test]
    fn root_patriarch_is_hidden_and_nested_group_is_preserved() {
        let drawing = grouped_drawing();
        let shapes = EscherShapeFactory::extract_shapes_from_drawing(&drawing).unwrap();

        assert_eq!(shapes.len(), 2);
        assert_eq!(shapes[0].shape_id(), Some(2));
        assert_eq!(
            shapes[0].shape_type(),
            super::super::shape::EscherShapeType::Rectangle
        );

        let group = &shapes[1];
        assert_eq!(
            group.shape_type(),
            super::super::shape::EscherShapeType::Group
        );
        assert_eq!(group.shape_id(), Some(3));
        assert_eq!(group.anchor().map(|anchor| anchor.left), Some(100));
        assert_eq!(group.anchor().map(|anchor| anchor.top), Some(200));
        assert_eq!(group.children.len(), 1);
        assert_eq!(group.children[0].shape_id(), Some(4));
        assert_eq!(
            group.children[0].shape_type(),
            super::super::shape::EscherShapeType::Ellipse
        );
    }

    #[test]
    fn exposes_inert_animation_info_from_shape_client_data() {
        use crate::ppt::animation::{
            AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
            write_animation_info,
        };

        let atom = LegacyAnimationAtom {
            build_type: LegacyAnimationBuild::OneBuild,
            effect: LegacyAnimationEffect::Wipe,
            effect_direction: 2,
            order_id: 4,
            ..LegacyAnimationAtom::default()
        };
        let mut info = AnimationInfo::new();
        info.legacy_atom = Some(atom.clone());
        let (animation, _) = write_animation_info(&info).unwrap();

        let mut children = Vec::new();
        ShapeBuilder::new(1, 77).write(&mut children).unwrap();
        write_atom(&mut children, 0, 0, record_type::CLIENT_DATA, &animation).unwrap();
        let mut bytes = Vec::new();
        write_container(&mut bytes, 0, record_type::SP_CONTAINER, &children).unwrap();
        let (record, _) = super::super::record::EscherRecord::parse(&bytes, 0).unwrap();
        let shape = super::super::shape::EscherShape::from_container(EscherContainer::new(record));
        let parsed = shape.animation_info().unwrap().unwrap();
        assert_eq!(shape.shape_id(), Some(77));
        assert_eq!(parsed.legacy_atom, Some(atom));
    }

    #[test]
    fn shape_count_matches_extracted_top_level_shapes() {
        let drawing = grouped_drawing();
        let shapes = EscherShapeFactory::extract_shapes_from_drawing(&drawing).unwrap();

        assert_eq!(
            EscherShapeFactory::count_shapes_in_drawing(&drawing),
            shapes.len()
        );
    }
}

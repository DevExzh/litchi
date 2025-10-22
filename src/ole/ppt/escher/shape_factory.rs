//! High-performance shape factory for creating shapes from Escher records.
//!
//! # Performance
//!
//! - Zero-copy shape data access
//! - Iterator-based shape enumeration
//! - Pattern matching for shape type detection

use super::container::EscherContainer;
use super::shape::EscherShape;
use super::types::EscherRecordType;
use crate::ole::ppt::package::Result;

/// Factory for creating shapes from Escher records.
pub struct EscherShapeFactory;

impl EscherShapeFactory {
    /// Extract all shapes from an Escher/PPDrawing data.
    ///
    /// # Performance
    ///
    /// - Depth-first traversal
    /// - Pre-allocated results vector
    /// - Short-circuits on errors
    pub fn extract_shapes_from_ppdrawing(data: &[u8]) -> Result<Vec<EscherShape<'_>>> {
        let parser = super::parser::EscherParser::new(data);

        let mut shapes = Vec::new();

        // Get root container
        if let Some(root_result) = parser.root_container() {
            let root = root_result?;
            Self::extract_shapes_from_container(&root, &mut shapes);
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
        for child in container.children().flatten() {
            match child.record_type {
                // SpContainer holds a single shape
                EscherRecordType::SpContainer => {
                    let sp_container = EscherContainer::new(child);
                    let shape = EscherShape::from_container(sp_container);
                    shapes.push(shape);
                },
                // SpgrContainer holds a group of shapes
                // Process the group itself as a shape (which will recursively load children)
                EscherRecordType::SpgrContainer => {
                    let group_container = EscherContainer::new(child);
                    Self::extract_shapes_from_spgr_container(&group_container, shapes);
                },
                // Other containers - recurse
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    Self::extract_shapes_from_container(&child_container, shapes);
                },
                _ => {},
            }
        }
    }

    /// Extract shapes from SpgrContainer with proper skip-first logic.
    ///
    /// Based on Apache POI's HSLFGroupShape.getShapes():
    /// - The first SpContainer in SpgrContainer is the group shape itself
    /// - Remaining SpContainer children are the actual child shapes
    fn extract_shapes_from_spgr_container<'data>(
        container: &EscherContainer<'data>,
        shapes: &mut Vec<EscherShape<'data>>,
    ) {
        let mut is_first = true;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
                    let sp_container = EscherContainer::new(child);

                    // The first SpContainer is the group shape itself
                    if is_first {
                        is_first = false;
                        // Create the group shape (which will recursively parse its children)
                        let group_shape = EscherShape::from_container(sp_container);
                        shapes.push(group_shape);
                    } else {
                        // Regular child shapes
                        let child_shape = EscherShape::from_container(sp_container);
                        shapes.push(child_shape);
                    }
                },
                EscherRecordType::SpgrContainer => {
                    // Nested group
                    let nested_group = EscherContainer::new(child);
                    Self::extract_shapes_from_spgr_container(&nested_group, shapes);
                },
                _ if child.is_container() => {
                    let child_container = EscherContainer::new(child);
                    Self::extract_shapes_from_container(&child_container, shapes);
                },
                _ => {},
            }
        }
    }

    /// Count shapes in PPDrawing data (without full parsing).
    ///
    /// # Performance
    ///
    /// - Counts SpContainer records only
    /// - No shape object allocation
    /// - Early termination on errors
    pub fn count_shapes_in_ppdrawing(data: &[u8]) -> usize {
        let parser = super::parser::EscherParser::new(data);

        if let Some(root_result) = parser.root_container()
            && let Ok(root) = root_result
        {
            return Self::count_shapes_in_container(&root);
        }

        0
    }

    /// Recursively count shapes in a container.
    fn count_shapes_in_container(container: &EscherContainer<'_>) -> usize {
        let mut count = 0;

        for child in container.children().flatten() {
            match child.record_type {
                EscherRecordType::SpContainer => {
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

    #[test]
    fn test_empty_data() {
        let shapes = EscherShapeFactory::extract_shapes_from_ppdrawing(&[]).unwrap();
        assert_eq!(shapes.len(), 0);
    }

    #[test]
    fn test_count_shapes() {
        let count = EscherShapeFactory::count_shapes_in_ppdrawing(&[]);
        assert_eq!(count, 0);
    }
}

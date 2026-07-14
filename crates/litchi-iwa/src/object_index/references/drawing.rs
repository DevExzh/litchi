//! Application-specific object-reference extraction.

use super::super::ObjectIndex;
use crate::archive::RawMessage;

impl ObjectIndex {
    pub(super) fn extract_drawing_references(&mut self, object_id: u64, raw_msg: &RawMessage) {
        use prost::Message;

        match raw_msg.type_ {
            // TSD (Drawing/Shape) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSDArchives.proto and libetonyek's reference extraction
            3002 => {
                // TSD.DrawableArchive - base type for all drawables
                if let Ok(drawable) = crate::protobuf::tsd::DrawableArchive::decode(&*raw_msg.data)
                {
                    // Extract parent reference (drawable hierarchy)
                    if let Some(ref parent) = drawable.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Note: geometry is not a reference, just position/size data
                    // exterior_text_wrap is configuration, not a reference
                }
            },
            3003 => {
                // TSD.ContainerArchive - container for grouped objects
                if let Ok(container) =
                    crate::protobuf::tsd::ContainerArchive::decode(&*raw_msg.data)
                {
                    // Extract parent reference
                    if let Some(ref parent) = container.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Extract all child references
                    for child in &container.children {
                        self.extract_reference(object_id, child);
                    }
                }
            },
            3004 => {
                // TSD.ShapeArchive - shapes (rectangles, circles, polygons, etc.)
                if let Ok(shape) = crate::protobuf::tsd::ShapeArchive::decode(&*raw_msg.data) {
                    // ShapeArchive embeds DrawableArchive in 'super' field (required)
                    // Extract parent from the super DrawableArchive
                    if let Some(ref parent) = shape.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = shape.style {
                        self.extract_reference(object_id, style);
                    }
                    // Note: pathsource, head_line_end, tail_line_end are not references
                    // but embedded data structures
                }
            },
            3005 => {
                // TSD.ImageArchive - images
                if let Ok(image) = crate::protobuf::tsd::ImageArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = image.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = image.style {
                        self.extract_reference(object_id, style);
                    }
                    // Note: data field is a DataReference, not an object Reference
                    // database_originalData is also for media assets
                }
            },
            3006 => {
                // TSD.MaskArchive - image masks
                if let Ok(mask) = crate::protobuf::tsd::MaskArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = mask.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Note: pathsource is embedded data, not a reference
                }
            },
            3007 => {
                // TSD.MovieArchive - video objects
                if let Ok(movie) = crate::protobuf::tsd::MovieArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = movie.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = movie.style {
                        self.extract_reference(object_id, style);
                    }
                    // Note: movieData is a DataReference, not an object Reference
                }
            },
            3008 => {
                // TSD.GroupArchive - grouped shapes/objects
                if let Ok(group) = crate::protobuf::tsd::GroupArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = group.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    // Extract all child references (objects in the group)
                    for child in &group.children {
                        self.extract_reference(object_id, child);
                    }
                }
            },
            3009 => {
                // TSD.ConnectionLineArchive - connector lines between shapes
                if let Ok(conn_line) =
                    crate::protobuf::tsd::ConnectionLineArchive::decode(&*raw_msg.data)
                {
                    // Extract parent and style from super ShapeArchive (required field)
                    // ConnectionLineArchive.super_ is ShapeArchive
                    // ShapeArchive.super_ is DrawableArchive
                    if let Some(ref parent) = conn_line.super_.super_.parent {
                        self.extract_reference(object_id, parent);
                    }
                    if let Some(ref style) = conn_line.super_.style {
                        self.extract_reference(object_id, style);
                    }
                    // Extract connection endpoints
                    if let Some(ref connected_from) = conn_line.connected_from {
                        self.extract_reference(object_id, connected_from);
                    }
                    if let Some(ref connected_to) = conn_line.connected_to {
                        self.extract_reference(object_id, connected_to);
                    }
                }
            },

            _ => {},
        }
    }
}

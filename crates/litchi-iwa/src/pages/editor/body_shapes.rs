//! Body-anchored ordinary shape CRUD for Pages documents.

use std::{collections::HashMap, ops::Range};

use super::*;
use crate::package_metadata::{
    add_component_external_reference, component_identifier_for_entry,
    remove_component_external_references_to_object,
};
use crate::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, Endpoints,
    LineSegment, LineStyle, RgbaColor, Shadow, ShapeFill, ShapeImageFill, ShapeImageFillTechnique,
    ShapePathKind, Stroke, flip_drawable_geometry, line_geometry, line_path_source,
    line_segments_match, reset_shape_effects, reset_shape_fill, reset_shape_shadow,
    reset_shape_stroke, reset_shape_text_layout, set_shape_effects, set_shape_fill,
    set_shape_geometry, set_shape_image_fill_data, set_shape_line_endpoints,
    set_shape_line_segment, set_shape_preset, set_shape_shadow, set_shape_stroke,
    set_shape_text_layout, shape_effects, shape_fill, shape_line_endpoints, shape_path_source,
    shape_shadow, shape_stroke, shape_text_layout,
};
use crate::text::layout::Layout;
use litchi_iwa_common::shape::effects::Effects;
use litchi_iwa_common::shape::path::Preset;

use super::text_box_create::{
    BodyTextShapeObjectIds, BodyTextShapeRole, body_text_shape_objects, body_text_storage,
};

mod caption;
mod graph;

use caption::*;
use graph::*;

/// Structural path family used by an ordinary Pages body shape.
pub type PagesBodyShapeKind = ShapePathKind;

/// One ordinary, non-text-box shape anchored to the Pages body text flow.
#[derive(Debug, Clone, PartialEq)]
pub struct PagesBodyShapeInfo {
    pub drawable_object_id: u64,
    /// UTF-16 index of the object-replacement character in the body text.
    pub anchor_character_index: u32,
    pub kind: PagesBodyShapeKind,
    /// Source-buildable preset and its native controls, when recognized.
    pub preset: Option<Preset>,
    /// Document-space endpoints when this shape is a native straight line.
    pub line_segment: Option<LineSegment>,
    /// Directed start/end decorations when this shape is a native straight line.
    pub line_endpoints: Option<Endpoints>,
    pub storage: TextStorageInfo,
    pub geometry: DrawableGeometry,
    pub properties: DrawableProperties,
}

/// Result of removing an ordinary body shape and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedPagesBodyShape {
    pub shape: PagesBodyShapeInfo,
}

impl PagesEditor {
    /// List ordinary non-text-box shapes anchored to the document body.
    pub fn body_shapes(&self) -> Result<Vec<PagesBodyShapeInfo>> {
        body_shape_infos(self)
    }

    /// Add a text-bearing rectangle at one UTF-16 position in the body.
    ///
    /// The shape, path, writable storage, stand-ins, body attachment, z-order,
    /// UUIDs, and style relationship are constructed directly from typed
    /// values. No source drawable or package template is copied.
    pub fn add_body_rectangle(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<PagesBodyShapeInfo> {
        self.add_body_shape(
            anchor_character_index,
            text,
            position,
            size,
            Preset::Rectangle,
        )
    }

    /// Add a typed preset shape at one UTF-16 position in the body.
    ///
    /// The path, writable storage, stand-ins, body attachment, z-order, UUIDs,
    /// and style relationship are constructed directly from typed values. No
    /// source drawable or package template is copied.
    pub fn add_body_shape(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: Preset,
    ) -> Result<PagesBodyShapeInfo> {
        let geometry = new_shape_geometry(position, size)?;
        self.add_body_shape_path(
            anchor_character_index,
            text,
            geometry,
            shape_path_source(preset, size)?,
            Some(preset),
            None,
        )
    }

    /// Add a typed preset shape with an explicit standard fill.
    pub fn add_body_shape_with_fill(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: Preset,
        fill: ShapeFill,
    ) -> Result<PagesBodyShapeInfo> {
        let created = self.add_body_shape(anchor_character_index, text, position, size, preset)?;
        self.set_body_shape_fill(created.drawable_object_id, &fill)?;
        Ok(body_shape_graph(self, created.drawable_object_id)?.info)
    }

    /// Add a native straight line between two body-document points.
    ///
    /// The line path, empty writable storage, stand-ins, body attachment,
    /// z-order, UUIDs, and style relationship are constructed directly from
    /// typed values. No source drawable or package template is copied.
    pub fn add_body_line(
        &mut self,
        anchor_character_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<PagesBodyShapeInfo> {
        let segment = LineSegment::new(start, end)?;
        self.add_body_shape_path(
            anchor_character_index,
            "",
            line_geometry(segment),
            line_path_source(segment),
            None,
            Some(segment),
        )
    }

    /// Add a native straight line with independently typed start and end decorations.
    pub fn add_body_line_with_endpoints(
        &mut self,
        anchor_character_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
        endpoints: Endpoints,
    ) -> Result<PagesBodyShapeInfo> {
        let created = self.add_body_line(anchor_character_index, start, end)?;
        self.set_body_line_endpoints(created.drawable_object_id, endpoints)?;
        Ok(body_shape_graph(self, created.drawable_object_id)?.info)
    }

    /// Add a native straight line with a typed stroke and endpoint appearance.
    pub fn add_body_line_with_style(
        &mut self,
        anchor_character_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
        style: LineStyle,
    ) -> Result<PagesBodyShapeInfo> {
        let created = self.add_body_line(anchor_character_index, start, end)?;
        self.set_body_shape_stroke(created.drawable_object_id, style.stroke)?;
        self.set_body_line_endpoints(created.drawable_object_id, style.endpoints)?;
        Ok(body_shape_graph(self, created.drawable_object_id)?.info)
    }

    fn add_body_shape_path(
        &mut self,
        anchor_character_index: usize,
        text: &str,
        geometry: DrawableGeometry,
        path_source: tsd::PathSourceArchive,
        expected_preset: Option<Preset>,
        expected_line: Option<LineSegment>,
    ) -> Result<PagesBodyShapeInfo> {
        let root = root_document(self.package())?;
        let body: StorageArchive = decode_typed_package_object(
            self.package(),
            self.body_storage_id,
            self.body_storage()?.message_type,
            "TSWP.StorageArchive",
        )?;
        let style_id = shape_style_id(self.package(), &root)?;
        let storage = body_text_storage(text, &body);
        let first_identifier = next_object_identifier(self.package())?;
        let (creates_z_order, z_order_id) = if let Some(z_order) = &root.drawables_zorder {
            (false, z_order.identifier)
        } else {
            (true, first_identifier)
        };
        let graph_first_identifier = first_identifier
            .checked_add(u64::from(creates_z_order))
            .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
        let ids = BodyTextShapeObjectIds::allocate(graph_first_identifier)?;
        let archive_name = find_object_archive(self.package(), self.body_storage_id)?;
        let objects = body_text_shape_objects(
            ids,
            self.body_storage_id,
            style_id,
            geometry,
            storage,
            root.left_margin.unwrap_or_default(),
            path_source,
            BodyTextShapeRole::Shape,
        )?;

        let mut staged = self.package().clone();
        if creates_z_order {
            text_box_create::create_drawable_z_order(&mut staged, &archive_name, z_order_id)?;
        }
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            ids.attachment,
        )?;
        patch_pages_zorder(&mut staged, None, Some(ids.drawable))?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &ids.uuid_objects())?;
        let style_archive = find_object_archive(&staged, style_id)?;
        if let Some(stylesheet_component) = component_identifier_for_entry(&staged, &style_archive)?
            && stylesheet_component != DOCUMENT_OBJECT_ID
        {
            add_component_external_reference(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                stylesheet_component,
                style_id,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .body_shapes()?
            .into_iter()
            .find(|shape| shape.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages shape creation failed validation".to_owned())
            })?;
        let created_graph = body_shape_graph(&verified, ids.drawable)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        let line_matches = match (created.line_segment, expected_line) {
            (Some(actual), Some(expected)) => line_segments_match(actual, expected),
            (None, None) => true,
            _ => false,
        };
        if created.anchor_character_index != expected_anchor
            || created.preset != expected_preset
            || !line_matches
            || created.storage.object_id != ids.storage
            || created.storage.text != text
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
        {
            return Err(Error::InvalidFormat(
                "Pages shape creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the document-space endpoints of one native straight line.
    pub fn body_line_segment(&self, drawable_object_id: u64) -> Result<LineSegment> {
        body_shape_graph(self, drawable_object_id)?
            .info
            .line_segment
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Pages drawable {drawable_object_id} is not a native straight line"
                ))
            })
    }

    /// Read the directed start and end decorations of one native straight line.
    pub fn body_line_endpoints(&self, drawable_object_id: u64) -> Result<Endpoints> {
        let source = body_shape_graph(self, drawable_object_id)?;
        source.info.line_endpoints.ok_or_else(|| {
            Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a native straight line"
            ))
        })
    }

    /// Replace both decorations transactionally, using copy-on-write when the style is shared.
    pub fn set_body_line_endpoints(
        &mut self,
        drawable_object_id: u64,
        endpoints: Endpoints,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a native straight line"
            )));
        }
        let mut staged = self.package().clone();
        set_shape_line_endpoints(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            endpoints,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_line_endpoints(drawable_object_id)? != endpoints {
            return Err(Error::InvalidFormat(
                "Pages line endpoint-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete explicit endpoint decorations and restore undecorated style inheritance.
    ///
    /// Returns `true` when decorations were reset and `false` when the line was
    /// already undecorated.
    pub fn reset_body_line_endpoints(&mut self, drawable_object_id: u64) -> Result<bool> {
        if self.body_line_endpoints(drawable_object_id)? == Endpoints::default() {
            return Ok(false);
        }
        self.set_body_line_endpoints(drawable_object_id, Endpoints::default())?;
        Ok(true)
    }

    /// Read the effective standard stroke of one ordinary body shape.
    pub fn body_shape_stroke(&self, drawable_object_id: u64) -> Result<Option<Stroke>> {
        let source = body_shape_graph(self, drawable_object_id)?;
        shape_stroke(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's stroke transactionally, using copy-on-write for shared styles.
    pub fn set_body_shape_stroke(&mut self, drawable_object_id: u64, stroke: Stroke) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_stroke(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            stroke,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_stroke(drawable_object_id)? != Some(stroke) {
            return Err(Error::InvalidFormat(
                "Pages shape stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct stroke override and restore the inherited appearance.
    pub fn reset_body_shape_stroke(&mut self, drawable_object_id: u64) -> Result<bool> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_stroke(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective standard fill of one ordinary body shape.
    pub fn body_shape_fill(&self, drawable_object_id: u64) -> Result<ShapeFill> {
        let source = body_shape_graph(self, drawable_object_id)?;
        shape_fill(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's fill transactionally, using copy-on-write for shared styles.
    pub fn set_body_shape_fill(&mut self, drawable_object_id: u64, fill: &ShapeFill) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_fill(&mut staged, &source.archive_name, drawable_object_id, fill)?;
        let verified = Self::from_package(staged)?;
        if &verified.body_shape_fill(drawable_object_id)? != fill {
            return Err(Error::InvalidFormat(
                "Pages shape fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Embed image bytes and use them as a simple or tinted native shape fill.
    pub fn set_body_shape_image_fill(
        &mut self,
        drawable_object_id: u64,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let fill_size = source.info.geometry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages shape {drawable_object_id} has no image-fill dimensions"
            ))
        })?;
        let mut staged = self.package().clone();
        let image = set_shape_image_fill_data(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            preferred_filename,
            data,
            technique,
            fill_size,
            tint,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_fill(drawable_object_id)? != ShapeFill::Image(image.clone()) {
            return Err(Error::InvalidFormat(
                "Pages shape image-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(image)
    }

    /// Remove a direct fill override and restore the inherited appearance.
    pub fn reset_body_shape_fill(&mut self, drawable_object_id: u64) -> Result<bool> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_fill(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective whole-object opacity and reflection settings.
    pub fn body_shape_effects(&self, drawable_object_id: u64) -> Result<Effects> {
        let source = body_shape_graph(self, drawable_object_id)?;
        shape_effects(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace opacity and reflection atomically while preserving other style properties.
    pub fn set_body_shape_effects(
        &mut self,
        drawable_object_id: u64,
        effects: Effects,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_effects(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            effects,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_effects(drawable_object_id)? != effects {
            return Err(Error::InvalidFormat(
                "Pages shape effect update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct effect overrides and restore inherited values.
    pub fn reset_body_shape_effects(&mut self, drawable_object_id: u64) -> Result<bool> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_effects(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective drop, contact, curved, or disabled shadow state.
    pub fn body_shape_shadow(&self, drawable_object_id: u64) -> Result<Shadow> {
        let source = body_shape_graph(self, drawable_object_id)?;
        shape_shadow(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace the shadow while preserving fill, stroke, opacity, and reflection.
    pub fn set_body_shape_shadow(&mut self, drawable_object_id: u64, shadow: Shadow) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let staged = set_shape_shadow(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
            shadow,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_shadow(drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Pages shape shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct shadow override and restore inherited shadow state.
    pub fn reset_body_shape_shadow(&mut self, drawable_object_id: u64) -> Result<bool> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let (staged, changed) = reset_shape_shadow(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective vertical alignment, edge insets, and autosizing.
    pub fn body_shape_text_layout(&self, drawable_object_id: u64) -> Result<Layout> {
        let source = body_shape_graph(self, drawable_object_id)?;
        shape_text_layout(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace frame-level text layout while preserving drawing style and columns.
    pub fn set_body_shape_text_layout(
        &mut self,
        drawable_object_id: u64,
        layout: Layout,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_text_layout(drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Pages shape text-layout update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct frame-level text-layout overrides and restore inherited values.
    pub fn reset_body_shape_text_layout(&mut self, drawable_object_id: u64) -> Result<bool> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_layout(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Move or resize one native straight line by replacing its endpoints.
    pub fn set_body_line_segment(
        &mut self,
        drawable_object_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Pages drawable {drawable_object_id} is not a native straight line"
            )));
        }
        let replacement = LineSegment::new(start, end)?;
        let mut staged = self.package().clone();
        set_shape_line_segment(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            replacement,
        )?;
        let verified = Self::from_package(staged)?;
        let actual = verified.body_line_segment(drawable_object_id)?;
        if !line_segments_match(actual, replacement) {
            return Err(Error::InvalidFormat(
                "Pages line endpoint update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the recognized preset and native controls for one body shape.
    pub fn body_shape_preset(&self, drawable_object_id: u64) -> Result<Option<Preset>> {
        Ok(body_shape_graph(self, drawable_object_id)?.info.preset)
    }

    /// Replace a body shape's preset path while retaining its text and style.
    pub fn set_body_shape_preset(&mut self, drawable_object_id: u64, preset: Preset) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_preset(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            preset,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_preset(drawable_object_id)? != Some(preset) {
            return Err(Error::InvalidFormat(
                "Pages shape preset update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Apply one native Arrange Flip operation to an ordinary body shape.
    ///
    /// Returns the updated geometry after applying the same transform as the
    /// Pages Flip Horizontally or Flip Vertically inspector button.
    pub fn flip_body_shape(
        &mut self,
        drawable_object_id: u64,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let geometry = flip_drawable_geometry(source.info.geometry, axis)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages shape flip update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Read typed geometry for one ordinary body shape.
    pub fn body_shape_geometry(&self, drawable_object_id: u64) -> Result<DrawableGeometry> {
        Ok(body_shape_graph(self, drawable_object_id)?.info.geometry)
    }

    /// Update body-shape geometry while preserving unknown shape fields.
    pub fn set_body_shape_geometry(
        &mut self,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_geometry(drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Pages shape geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary body shape.
    pub fn body_shape_properties(&self, drawable_object_id: u64) -> Result<DrawableProperties> {
        Ok(body_shape_graph(self, drawable_object_id)?.info.properties)
    }

    /// Update shared drawable properties while preserving unknown fields.
    pub fn set_body_shape_properties(
        &mut self,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.body_shape_properties(drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Pages shape property update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a UTF-16 range in an ordinary body shape's text.
    pub fn replace_body_shape_text(
        &mut self,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package().clone());
        text.replace_text(source.info.storage.object_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        body_shape_graph(&verified, drawable_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Replace all text in an ordinary body shape.
    pub fn set_body_shape_text(
        &mut self,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let source = body_shape_graph(self, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package().clone());
        text.set_text(source.info.storage.object_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = body_shape_graph(&verified, drawable_object_id)?;
        if updated.info.storage.text != replacement {
            return Err(Error::InvalidFormat(
                "Pages shape text update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear an ordinary body shape without deleting it.
    pub fn clear_body_shape_text(&mut self, drawable_object_id: u64) -> Result<()> {
        self.set_body_shape_text(drawable_object_id, "")
    }

    /// Duplicate an ordinary body shape at a UTF-16 body position.
    ///
    /// The shape, rich-text storage, body attachment, and stand-in objects
    /// receive fresh identifiers while retaining the source's styling and
    /// unknown protobuf fields. The clone is offset using Pages' native
    /// duplicate placement and its writable storage remains independent.
    pub fn duplicate_body_shape(
        &mut self,
        source_drawable_object_id: u64,
        anchor_character_index: usize,
    ) -> Result<PagesBodyShapeInfo> {
        let source = body_shape_graph(self, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Pages shape graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages shape object {identifier} is missing"))
                })?;
                clone_pages_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = remap[&source_drawable_object_id];
        let new_storage_id = remap[&source.info.storage.object_id];
        let new_attachment_id = remap[&source.attachment_id];
        offset_pages_body_drawable_clone(
            &mut staged,
            new_drawable_id,
            new_attachment_id,
            BODY_DRAWABLE_DUPLICATE_OFFSET,
        )?;
        let mut text_editor = IWorkTextEditor::from_package(staged);
        text_editor.replace_text(
            self.body_storage_id,
            anchor_character_index..anchor_character_index,
            "\u{fffc}",
        )?;
        staged = text_editor.into_package();
        add_body_drawable_attachment(
            &mut staged,
            self.body_storage_id,
            anchor_character_index,
            new_attachment_id,
        )?;
        patch_pages_zorder(&mut staged, None, Some(new_drawable_id))?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Pages shape graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &new_uuid_object_ids)?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .body_shapes()?
            .into_iter()
            .find(|shape| shape.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Pages shape duplication failed validation".to_owned())
            })?;
        let created_graph = body_shape_graph(&verified, new_drawable_id)?;
        let expected_anchor = u32::try_from(anchor_character_index)
            .map_err(|_| Error::ParseError("Pages body attachment index exceeds u32".to_owned()))?;
        if created.anchor_character_index != expected_anchor
            || created.storage.object_id != new_storage_id
            || created.storage.text != source.info.storage.text
            || created.kind != source.info.kind
            || created.preset != source.info.preset
            || created.line_segment != source.info.line_segment
            || created.line_endpoints != source.info.line_endpoints
            || created.geometry.size != source.info.geometry.size
            || created.geometry.flags != source.info.geometry.flags
            || created.geometry.angle != source.info.geometry.angle
            || created.properties != source.info.properties
            || created_graph.object_ids.len() != source.object_ids.len()
        {
            return Err(Error::InvalidFormat(
                "Pages shape duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Remove an ordinary body shape and its private text-bearing graph.
    pub fn remove_body_shape(&mut self, drawable_object_id: u64) -> Result<RemovedPagesBodyShape> {
        let graph = body_shape_graph(self, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(litchi_iwa_common::comment::DrawableId::from_raw(
            drawable_object_id,
        )?)?;
        let mut text_editor = IWorkTextEditor::from_package(comments.into_package());
        let anchor = graph.info.anchor_character_index as usize;
        text_editor.replace_text(self.body_storage_id, anchor..anchor + 1, "")?;
        let mut staged = text_editor.into_package();
        patch_pages_zorder(&mut staged, Some(drawable_object_id), None)?;
        for identifier in &graph.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                DOCUMENT_OBJECT_ID,
                *identifier,
            )?;
        }
        staged.update_archive(&graph.archive_name, |archive| {
            for identifier in &graph.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Pages shape object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        for identifier in &graph.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Pages shape object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &graph.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &graph.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .body_shapes()?
            .iter()
            .any(|shape| shape.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Pages shape deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedPagesBodyShape { shape: graph.info })
    }
}

#[cfg(test)]
mod tests {
    use std::{fs, path::PathBuf};

    use super::*;
    use crate::shapes::{Appearance, BlurRadius, Drop, Endpoint, Offset, Pattern, RgbColorSpace,
        RgbaColor, Width};
    use crate::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
    use litchi_iwa_common::shape::effects::{Effects, Opacity as EffectsOpacity, Reflection,
        ReflectionOpacity};
    use litchi_iwa_common::shape::fill::{Angle, Gradient};
    use litchi_iwa_common::shape::shadow::{Angle as ShadowAngle, Opacity as ShadowOpacity};
    use litchi_iwa_common::shape::path::CornerRadius;

    const POSITION: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 300.0,
        height: 150.0,
    };
    const LINE_START: DrawablePoint = DrawablePoint { x: 180.0, y: 240.0 };
    const LINE_END: DrawablePoint = DrawablePoint { x: 480.0, y: 390.0 };
    const UPDATED_LINE_START: DrawablePoint = DrawablePoint { x: 96.0, y: 180.0 };
    const UPDATED_LINE_END: DrawablePoint = DrawablePoint { x: 456.0, y: 180.0 };

    fn fixture(relative: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        fs::read(root.join(relative)).unwrap()
    }

    #[test]
    fn scratch_document_supports_rectangle_crud_without_a_source_drawable() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline_body = editor.body_text().unwrap();
        assert!(editor.body_shapes().unwrap().is_empty());

        let created = editor
            .add_body_rectangle(4, "Built from typed objects", POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, PagesBodyShapeKind::Rectangle);
        assert_eq!(created.preset, Some(Preset::Rectangle));
        assert_eq!(created.storage.text, "Built from typed objects");
        let horizontally_flipped = editor
            .flip_body_shape(created.drawable_object_id, DrawableFlipAxis::Horizontal)
            .unwrap();
        assert_eq!(
            editor
                .body_shape_geometry(created.drawable_object_id)
                .unwrap(),
            horizontally_flipped
        );
        assert_ne!(horizontally_flipped.flags, created.geometry.flags);
        let vertically_flipped = editor
            .flip_body_shape(created.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();
        assert_eq!(
            editor
                .body_shape_geometry(created.drawable_object_id)
                .unwrap(),
            vertically_flipped
        );
        assert_ne!(vertically_flipped.angle, created.geometry.angle);
        editor
            .replace_body_shape_text(created.drawable_object_id, 0..5, "Made")
            .unwrap();
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 96.0, y: 180.0 }),
            size: Some(DrawableSize {
                width: 360.0,
                height: 180.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(12.0),
        };
        editor
            .set_body_shape_geometry(created.drawable_object_id, geometry)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/pages-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Generated rectangle".to_owned()),
        };
        editor
            .set_body_shape_properties(created.drawable_object_id, properties.clone())
            .unwrap();
        let rectangle_stroke = Stroke::new(
            RgbaColor::new(0.2, 0.7, 0.3, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(2.0).unwrap(),
            Pattern::Solid,
        );
        editor
            .set_body_shape_stroke(created.drawable_object_id, rectangle_stroke)
            .unwrap();

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let shape = &reopened.body_shapes().unwrap()[0];
        assert_eq!(shape.storage.text, "Made from typed objects");
        assert_eq!(shape.geometry, geometry);
        assert_eq!(shape.properties, properties);
        assert_eq!(
            reopened
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(rectangle_stroke)
        );
        assert!(
            editor
                .reset_body_shape_stroke(created.drawable_object_id)
                .unwrap()
        );

        let mut package = editor.into_package();
        let root = root_document(&package).unwrap();
        let style_id = shape_style_id(&package, &root).unwrap();
        let style_archive = find_object_archive(&package, style_id).unwrap();
        let style_component = component_identifier_for_entry(&package, &style_archive)
            .unwrap()
            .unwrap();
        add_component_external_reference(
            &mut package,
            style_component,
            DOCUMENT_OBJECT_ID,
            created.drawable_object_id,
        )
        .unwrap();
        editor = PagesEditor::from_package(package).unwrap();

        let removed = editor
            .remove_body_shape(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.body_text().unwrap(), baseline_body);
        assert!(editor.body_shapes().unwrap().is_empty());

        let rectangle = editor
            .add_body_rectangle(4, "Not a line", POSITION, SIZE)
            .unwrap();
        let before_cross_type = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_body_shape(u64::MAX, DrawableFlipAxis::Horizontal)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
        assert!(
            editor
                .set_body_line_segment(rectangle.drawable_object_id, LINE_START, LINE_END)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
    }

    #[test]
    fn scratch_document_supports_native_shape_duplication() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let created = editor
            .add_body_shape(4, "Source shape", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.2, 0.6, 0.9, 1.0, RgbColorSpace::Srgb).unwrap());
        editor
            .set_body_shape_fill(created.drawable_object_id, &fill)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/source-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(false),
            accessibility_description: Some("Source shape".to_owned()),
        };
        editor
            .set_body_shape_properties(created.drawable_object_id, properties.clone())
            .unwrap();
        editor
            .flip_body_shape(created.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();
        let source = editor.body_shapes().unwrap().into_iter().next().unwrap();
        let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();

        let duplicate = editor
            .duplicate_body_shape(source.drawable_object_id, duplicate_anchor)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        assert_ne!(duplicate.storage.object_id, source.storage.object_id);
        assert_eq!(duplicate.anchor_character_index, duplicate_anchor as u32);
        assert_eq!(duplicate.storage.text, source.storage.text);
        assert_eq!(duplicate.kind, source.kind);
        assert_eq!(duplicate.preset, source.preset);
        assert_eq!(duplicate.properties, properties);
        assert_eq!(
            duplicate.geometry.position,
            source.geometry.position.map(|position| DrawablePoint {
                x: position.x + BODY_DRAWABLE_DUPLICATE_OFFSET,
                y: position.y + BODY_DRAWABLE_DUPLICATE_OFFSET,
            })
        );
        assert_eq!(
            editor
                .body_shape_fill(duplicate.drawable_object_id)
                .unwrap(),
            fill
        );

        editor
            .set_body_shape_text(duplicate.drawable_object_id, "Independent copy")
            .unwrap();
        assert_eq!(
            editor
                .body_shapes()
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == source.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Source shape"
        );
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shapes()
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Independent copy"
        );
        assert_eq!(reopened.body_shapes().unwrap().len(), 2);

        let removed = editor
            .remove_body_shape(duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.storage.text, "Independent copy");
        assert_eq!(editor.body_shapes().unwrap().len(), 1);
    }

    #[test]
    fn scratch_document_supports_straight_line_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline_body = editor.body_text().unwrap();
        let created = editor.add_body_line(4, LINE_START, LINE_END).unwrap();
        assert_eq!(created.kind, PagesBodyShapeKind::Line);
        assert_eq!(created.preset, None);
        assert_eq!(created.storage.text, "");
        assert!(line_segments_match(
            created.line_segment.unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(line_segments_match(
            reopened
                .body_line_segment(created.drawable_object_id)
                .unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        editor
            .set_body_line_segment(
                created.drawable_object_id,
                UPDATED_LINE_START,
                UPDATED_LINE_END,
            )
            .unwrap();
        assert!(line_segments_match(
            editor
                .body_line_segment(created.drawable_object_id)
                .unwrap(),
            LineSegment::new(UPDATED_LINE_START, UPDATED_LINE_END).unwrap()
        ));

        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_line_segment(
                    created.drawable_object_id,
                    UPDATED_LINE_START,
                    UPDATED_LINE_START,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        let removed = editor
            .remove_body_shape(created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.kind, PagesBodyShapeKind::Line);
        assert_eq!(editor.body_text().unwrap(), baseline_body);
        assert!(editor.body_shapes().unwrap().is_empty());
    }

    #[test]
    fn scratch_document_supports_typed_shape_stroke_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let stroke = Stroke::new(
            RgbaColor::new(0.8, 0.1, 0.2, 0.9, RgbColorSpace::DisplayP3).unwrap(),
            Width::new(3.5).unwrap(),
            Pattern::MediumDash,
        );
        let endpoints = Endpoints::new(Endpoint::OpenCircle, Endpoint::FilledArrow);
        let created = editor
            .add_body_line_with_style(
                4,
                LINE_START,
                LINE_END,
                LineStyle::new(stroke).with_endpoints(endpoints),
            )
            .unwrap();
        assert_eq!(
            editor
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(stroke)
        );
        assert_eq!(
            editor
                .body_line_endpoints(created.drawable_object_id)
                .unwrap(),
            endpoints
        );
        let object_count = editor
            .package()
            .iwa_entry_names()
            .map(|name| editor.package().archive(name).unwrap().objects.len())
            .sum::<usize>();

        let replacement = Stroke::new(
            RgbaColor::new(0.1, 0.3, 0.9, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(2.25).unwrap(),
            Pattern::LongDash,
        );
        editor
            .set_body_shape_stroke(created.drawable_object_id, replacement)
            .unwrap();
        assert_eq!(
            editor
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(replacement)
        );
        assert_eq!(
            editor
                .body_line_endpoints(created.drawable_object_id)
                .unwrap(),
            endpoints
        );
        assert_eq!(
            editor
                .package()
                .iwa_entry_names()
                .map(|name| editor.package().archive(name).unwrap().objects.len())
                .sum::<usize>(),
            object_count
        );
        assert!(
            editor
                .reset_body_line_endpoints(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            editor
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(replacement)
        );
        editor
            .set_body_line_endpoints(created.drawable_object_id, endpoints)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(replacement)
        );
        assert!(
            reopened
                .reset_body_shape_stroke(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            None
        );
        assert_eq!(
            reopened
                .body_line_endpoints(created.drawable_object_id)
                .unwrap(),
            endpoints
        );
        assert!(
            !reopened
                .reset_body_shape_stroke(created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_document_supports_composable_shape_fill_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let fill = ShapeFill::Solid(
            RgbaColor::new(0.85, 0.2, 0.15, 0.9, RgbColorSpace::DisplayP3).unwrap(),
        );
        let created = editor
            .add_body_shape(4, "Filled", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited_fill = editor.body_shape_fill(created.drawable_object_id).unwrap();
        editor
            .set_body_shape_fill(created.drawable_object_id, &fill)
            .unwrap();
        assert_eq!(
            editor.body_shape_fill(created.drawable_object_id).unwrap(),
            fill
        );

        let stroke = Stroke::new(RgbaColor::black(), Width::new(2.0).unwrap(), Pattern::Solid);
        editor
            .set_body_shape_stroke(created.drawable_object_id, stroke)
            .unwrap();
        let object_count = editor
            .package()
            .iwa_entry_names()
            .map(|name| editor.package().archive(name).unwrap().objects.len())
            .sum::<usize>();
        let replacement = ShapeFill::Gradient(Gradient::linear(
            RgbaColor::new(0.1, 0.45, 0.9, 1.0, RgbColorSpace::Srgb).unwrap(),
            RgbaColor::new(0.8, 0.15, 0.55, 1.0, RgbColorSpace::DisplayP3).unwrap(),
            Angle::from_degrees(45.0).unwrap(),
        ));
        editor
            .set_body_shape_fill(created.drawable_object_id, &replacement)
            .unwrap();
        assert_eq!(
            editor
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(stroke)
        );
        assert_eq!(
            editor
                .package()
                .iwa_entry_names()
                .map(|name| editor.package().archive(name).unwrap().objects.len())
                .sum::<usize>(),
            object_count
        );

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_fill(created.drawable_object_id)
                .unwrap(),
            replacement
        );
        assert!(
            reopened
                .reset_body_shape_fill(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_fill(created.drawable_object_id)
                .unwrap(),
            inherited_fill
        );
        assert_eq!(
            reopened
                .body_shape_stroke(created.drawable_object_id)
                .unwrap(),
            Some(stroke)
        );
        assert!(
            !reopened
                .reset_body_shape_fill(created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_document_supports_embedded_shape_image_fill_crud() {
        let image_bytes = fixture("test-data/images/png/lena.png");
        let replacement_bytes = fixture("crates/soapberry-zip/assets/gophercolor16x16.png");
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let created = editor
            .add_body_shape(4, "Image", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor.body_shape_fill(created.drawable_object_id).unwrap();
        let image = editor
            .set_body_shape_image_fill(
                created.drawable_object_id,
                "lena.png",
                &image_bytes,
                ShapeImageFillTechnique::ScaleToFill,
                None,
            )
            .unwrap();
        let identifier = image.data_identifier().unwrap();
        assert_eq!(image.fill_size(), SIZE);
        assert_eq!(
            editor
                .extract_media(crate::MediaAssetId::try_from(identifier.get()).unwrap())
                .unwrap(),
            image_bytes
        );
        assert_eq!(editor.media_assets().unwrap().len(), 1);

        let advanced = image
            .with_technique(ShapeImageFillTechnique::Tile)
            .with_tint(RgbaColor::new(0.2, 0.4, 0.8, 0.5, RgbColorSpace::Srgb).unwrap());
        editor
            .set_body_shape_fill(
                created.drawable_object_id,
                &ShapeFill::Image(advanced.clone()),
            )
            .unwrap();
        assert_eq!(
            editor.body_shape_fill(created.drawable_object_id).unwrap(),
            ShapeFill::Image(advanced)
        );
        assert_eq!(
            editor
                .replace_media(
                    crate::MediaAssetId::try_from(identifier.get()).unwrap(),
                    &replacement_bytes,
                )
                .unwrap(),
            image_bytes
        );

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .reset_body_shape_fill(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_fill(created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(reopened.media_assets().unwrap().is_empty());
    }

    #[test]
    fn scratch_document_supports_composable_shape_effect_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.15, 0.45, 0.85, 1.0, RgbColorSpace::Srgb).unwrap());
        let created = editor
            .add_body_shape(4, "Effects", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        editor
            .set_body_shape_fill(created.drawable_object_id, &fill)
            .unwrap();
        let inherited = editor
            .body_shape_effects(created.drawable_object_id)
            .unwrap();
        let effects = Effects::new(
            EffectsOpacity::new(0.72).unwrap(),
            Reflection::Enabled(ReflectionOpacity::new(0.35).unwrap()),
        );
        editor
            .set_body_shape_effects(created.drawable_object_id, effects)
            .unwrap();
        assert_eq!(
            editor
                .body_shape_effects(created.drawable_object_id)
                .unwrap(),
            effects
        );
        assert_eq!(
            editor.body_shape_fill(created.drawable_object_id).unwrap(),
            fill
        );

        let replacement = Effects::new(EffectsOpacity::new(0.48).unwrap(), Reflection::Disabled);
        editor
            .set_body_shape_effects(created.drawable_object_id, replacement)
            .unwrap();
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_effects(created.drawable_object_id)
                .unwrap(),
            replacement
        );
        assert!(
            reopened
                .reset_body_shape_effects(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_effects(created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert_eq!(
            reopened
                .body_shape_fill(created.drawable_object_id)
                .unwrap(),
            fill
        );
        assert!(
            !reopened
                .reset_body_shape_effects(created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_document_supports_drop_shadow_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let created = editor
            .add_body_shape(4, "Shadow", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .body_shape_shadow(created.drawable_object_id)
            .unwrap();
        let effects = editor
            .body_shape_effects(created.drawable_object_id)
            .unwrap();
        let shadow = Shadow::Drop(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(7).unwrap(),
                Offset::from_points(11.0).unwrap(),
                ShadowOpacity::new(0.42).unwrap(),
            ),
            ShadowAngle::from_degrees(135.0).unwrap(),
        ));
        editor
            .set_body_shape_shadow(created.drawable_object_id, shadow)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_shadow(created.drawable_object_id)
                .unwrap(),
            shadow
        );
        assert_eq!(
            reopened
                .body_shape_effects(created.drawable_object_id)
                .unwrap(),
            effects
        );
        reopened
            .set_body_shape_shadow(created.drawable_object_id, Shadow::Disabled)
            .unwrap();
        assert_eq!(
            reopened
                .body_shape_shadow(created.drawable_object_id)
                .unwrap(),
            Shadow::Disabled
        );
        assert!(
            reopened
                .reset_body_shape_shadow(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_shadow(created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(
            !reopened
                .reset_body_shape_shadow(created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_document_supports_shape_text_layout_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let created = editor
            .add_body_shape(4, "Layout", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .body_shape_text_layout(created.drawable_object_id)
            .unwrap();
        let shadow = Shadow::Drop(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(7).unwrap(),
                Offset::from_points(11.0).unwrap(),
                ShadowOpacity::new(0.42).unwrap(),
            ),
            ShadowAngle::from_degrees(135.0).unwrap(),
        ));
        editor
            .set_body_shape_shadow(created.drawable_object_id, shadow)
            .unwrap();
        let layout = Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(12.0).unwrap()),
            AutoSize::Fixed,
        );
        editor
            .set_body_shape_text_layout(created.drawable_object_id, layout)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_text_layout(created.drawable_object_id)
                .unwrap(),
            layout
        );
        assert_eq!(
            reopened
                .body_shape_shadow(created.drawable_object_id)
                .unwrap(),
            shadow
        );
        assert!(
            reopened
                .reset_body_shape_text_layout(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_shape_text_layout(created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert_eq!(
            reopened
                .body_shape_shadow(created.drawable_object_id)
                .unwrap(),
            shadow
        );
        assert!(
            !reopened
                .reset_body_shape_text_layout(created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_document_supports_typed_line_endpoint_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let created = editor
            .add_body_line_with_endpoints(
                4,
                LINE_START,
                LINE_END,
                Endpoints::new(Endpoint::OpenCircle, Endpoint::FilledArrow),
            )
            .unwrap();
        assert_eq!(
            created.line_endpoints,
            Some(Endpoints::new(Endpoint::OpenCircle, Endpoint::FilledArrow))
        );
        let object_count_after_create = editor
            .package()
            .iwa_entry_names()
            .map(|name| editor.package().archive(name).unwrap().objects.len())
            .sum::<usize>();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let replacement = Endpoints::new(Endpoint::Line, Endpoint::SimpleArrow);
        reopened
            .set_body_line_endpoints(created.drawable_object_id, replacement)
            .unwrap();
        assert_eq!(
            reopened
                .body_line_endpoints(created.drawable_object_id)
                .unwrap(),
            replacement
        );
        let bytes_after_update = reopened.to_bytes().unwrap();
        reopened
            .set_body_line_endpoints(created.drawable_object_id, replacement)
            .unwrap();
        assert_eq!(reopened.to_bytes().unwrap(), bytes_after_update);
        let object_count_after_update = reopened
            .package()
            .iwa_entry_names()
            .map(|name| reopened.package().archive(name).unwrap().objects.len())
            .sum::<usize>();
        assert_eq!(object_count_after_update, object_count_after_create);
        assert!(
            reopened
                .reset_body_line_endpoints(created.drawable_object_id)
                .unwrap()
        );
        assert!(
            !reopened
                .reset_body_line_endpoints(created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .body_line_endpoints(created.drawable_object_id)
                .unwrap(),
            Endpoints::default()
        );
        let object_count_after_reset = reopened
            .package()
            .iwa_entry_names()
            .map(|name| reopened.package().archive(name).unwrap().objects.len())
            .sum::<usize>();
        assert_eq!(object_count_after_reset + 1, object_count_after_create);
    }

    #[test]
    fn scratch_document_supports_typed_preset_shape_crud() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline_body = editor.body_text().unwrap();
        let created = editor
            .add_body_shape(4, "Rounded", POSITION, SIZE, Preset::ROUNDED_RECTANGLE)
            .unwrap();
        assert_eq!(created.kind, PagesBodyShapeKind::RoundedRectangle);
        assert_eq!(created.preset, Some(Preset::ROUNDED_RECTANGLE));

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body_shape_preset(created.drawable_object_id)
                .unwrap(),
            Some(Preset::ROUNDED_RECTANGLE)
        );

        for (preset, kind) in [
            (Preset::Ellipse, PagesBodyShapeKind::Ellipse),
            (Preset::LeftArrow, PagesBodyShapeKind::LeftArrow),
            (Preset::RightArrow, PagesBodyShapeKind::RightArrow),
            (Preset::DoubleArrow, PagesBodyShapeKind::DoubleArrow),
            (Preset::PENTAGON, PagesBodyShapeKind::RegularPolygon),
            (Preset::STAR, PagesBodyShapeKind::Star),
        ] {
            editor
                .set_body_shape_preset(created.drawable_object_id, preset)
                .unwrap();
            let shape = &editor.body_shapes().unwrap()[0];
            assert_eq!(shape.kind, kind);
            assert_eq!(shape.preset, Some(preset));
            assert_eq!(shape.storage.text, "Rounded");
            assert_eq!(shape.anchor_character_index, 4);
        }

        editor
            .remove_body_shape(created.drawable_object_id)
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), baseline_body);
        assert!(editor.body_shapes().unwrap().is_empty());
    }

    #[test]
    fn invalid_creation_and_cross_type_updates_are_transactional() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .add_body_rectangle(
                    4,
                    "invalid",
                    POSITION,
                    DrawableSize {
                        width: 0.0,
                        height: SIZE.height,
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        assert!(
            editor
                .add_body_shape(
                    4,
                    "invalid radius",
                    POSITION,
                    SIZE,
                    Preset::RoundedRectangle {
                        corner_radius: CornerRadius::new(SIZE.height).unwrap(),
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let text_box = editor
            .add_text_box(4, "not a shape", POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_body_shape_text(text_box.drawable_object_id, "wrong type")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_body_shape_preset(text_box.drawable_object_id, Preset::Ellipse)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}

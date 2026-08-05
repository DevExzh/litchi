//! Ordinary text-bearing shape CRUD for Keynote slides.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use super::*;
use crate::image_caption::DrawableCaptionKind;
use crate::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize,
    LineEndpoints, LineSegment, LineStyle, RgbaColor, ShapeEffects, ShapeFill, ShapeImageFill,
    ShapeImageFillTechnique, ShapePathKind, ShapePreset, ShapeShadow, ShapeStroke,
    flip_drawable_geometry, line_geometry, line_path_source, line_segments_match,
    reset_shape_effects, reset_shape_fill, reset_shape_shadow, reset_shape_stroke,
    reset_shape_text_layout, set_shape_effects, set_shape_fill, set_shape_geometry,
    set_shape_image_fill_data, set_shape_line_endpoints, set_shape_line_segment, set_shape_preset,
    set_shape_shadow, set_shape_stroke, set_shape_text_layout, shape_effects, shape_fill,
    shape_line_endpoints, shape_line_segment, shape_path_kind, shape_path_source, shape_preset,
    shape_shadow, shape_stroke, shape_text_layout,
};
use crate::text::TextStorageInfo;
use crate::text::layout::Layout;

use super::text_box_create::{
    TextBoxObjectIds, slide_text_storage_template, text_box_context, text_box_objects,
    text_box_storage, text_box_theme_styles,
};

mod caption;

use caption::*;

const SHAPE_MESSAGE_TYPE: u32 = 2_011;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;

/// Structural path family used by an ordinary Keynote shape.
pub type KeynoteSlideShapeKind = ShapePathKind;

/// One ordinary, non-text-box shape owned directly by a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteSlideShapeInfo {
    pub slide_index: usize,
    pub drawable_object_id: u64,
    pub kind: KeynoteSlideShapeKind,
    /// Source-buildable preset and its native controls, when recognized.
    pub preset: Option<ShapePreset>,
    /// Slide-space endpoints when this shape is a native straight line.
    pub line_segment: Option<LineSegment>,
    /// Directed start/end decorations when this shape is a native straight line.
    pub line_endpoints: Option<LineEndpoints>,
    pub storage: TextStorageInfo,
    pub geometry: DrawableGeometry,
    pub properties: DrawableProperties,
}

/// Result of removing an ordinary shape and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedKeynoteSlideShape {
    pub shape: KeynoteSlideShapeInfo,
}

struct SlideShapeGraph {
    slide_id: u64,
    component_id: u64,
    archive_name: String,
    info: KeynoteSlideShapeInfo,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

impl KeynoteEditor {
    /// List ordinary non-text-box shapes owned directly by one slide.
    pub fn slide_shapes(&self, slide_index: usize) -> Result<Vec<KeynoteSlideShapeInfo>> {
        shape_infos(self, slide_index)
    }

    /// Add a rectangular shape with independent writable text to one slide.
    ///
    /// The shape, storage, title/caption stand-ins, ownership, z-order, UUID
    /// metadata, and style relationship are built directly from typed values.
    /// No source drawable or package template is copied.
    pub fn add_slide_rectangle(
        &mut self,
        slide_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<KeynoteSlideShapeInfo> {
        self.add_slide_shape(slide_index, text, position, size, ShapePreset::Rectangle)
    }

    /// Add a typed preset shape with independent writable text to one slide.
    ///
    /// The path, storage, title/caption stand-ins, ownership, z-order, UUID
    /// metadata, and style relationship are built directly from typed values.
    /// No source drawable or package template is copied.
    pub fn add_slide_shape(
        &mut self,
        slide_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: ShapePreset,
    ) -> Result<KeynoteSlideShapeInfo> {
        let geometry = new_shape_geometry(position, size)?;
        self.add_slide_shape_path(
            slide_index,
            text,
            geometry,
            shape_path_source(preset, size)?,
            Some(preset),
            None,
        )
    }

    /// Add a typed preset shape with an explicit standard fill.
    pub fn add_slide_shape_with_fill(
        &mut self,
        slide_index: usize,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: ShapePreset,
        fill: ShapeFill,
    ) -> Result<KeynoteSlideShapeInfo> {
        let created = self.add_slide_shape(slide_index, text, position, size, preset)?;
        self.set_slide_shape_fill(slide_index, created.drawable_object_id, &fill)?;
        Ok(shape_graph(self, slide_index, created.drawable_object_id)?.info)
    }

    /// Add a native straight line between two slide-space points.
    ///
    /// The line path, empty writable storage, stand-ins, ownership, z-order,
    /// UUID metadata, and style relationship are built directly from typed
    /// values. No source drawable or package template is copied.
    pub fn add_slide_line(
        &mut self,
        slide_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<KeynoteSlideShapeInfo> {
        let segment = LineSegment::new(start, end)?;
        self.add_slide_shape_path(
            slide_index,
            "",
            line_geometry(segment),
            line_path_source(segment),
            None,
            Some(segment),
        )
    }

    /// Add a native straight line with independently typed start and end decorations.
    pub fn add_slide_line_with_endpoints(
        &mut self,
        slide_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
        endpoints: LineEndpoints,
    ) -> Result<KeynoteSlideShapeInfo> {
        let created = self.add_slide_line(slide_index, start, end)?;
        self.set_slide_line_endpoints(slide_index, created.drawable_object_id, endpoints)?;
        Ok(shape_graph(self, slide_index, created.drawable_object_id)?.info)
    }

    /// Add a native straight line with a typed stroke and endpoint appearance.
    pub fn add_slide_line_with_style(
        &mut self,
        slide_index: usize,
        start: DrawablePoint,
        end: DrawablePoint,
        style: LineStyle,
    ) -> Result<KeynoteSlideShapeInfo> {
        let created = self.add_slide_line(slide_index, start, end)?;
        self.set_slide_shape_stroke(slide_index, created.drawable_object_id, style.stroke)?;
        self.set_slide_line_endpoints(slide_index, created.drawable_object_id, style.endpoints)?;
        Ok(shape_graph(self, slide_index, created.drawable_object_id)?.info)
    }

    fn add_slide_shape_path(
        &mut self,
        slide_index: usize,
        text: &str,
        geometry: DrawableGeometry,
        path_source: tsd::PathSourceArchive,
        expected_preset: Option<ShapePreset>,
        expected_line: Option<LineSegment>,
    ) -> Result<KeynoteSlideShapeInfo> {
        let graph = ObjectGraph::read(self.package())?;
        let context = text_box_context(&graph, slide_index)?;
        let styles = text_box_theme_styles(&graph, context.theme_id, context.stylesheet_id)?;
        let base_storage = slide_text_storage_template(&graph, &context.slide)?;
        let storage = text_box_storage(
            text,
            base_storage.as_ref(),
            &styles,
            context.language.as_deref(),
        );
        let ids = TextBoxObjectIds::allocate(next_object_identifier(self.package())?)?;
        let archive_name = graph.archive_name(context.slide_id)?.to_owned();
        let component_id = component_identifier_for_entry(self.package(), &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide component {archive_name} is not registered"
                ))
            })?;
        let stylesheet_archive = graph.archive_name(context.stylesheet_id)?;
        let stylesheet_component_id =
            component_identifier_for_entry(self.package(), stylesheet_archive)?.ok_or_else(
                || {
                    Error::InvalidFormat(format!(
                        "Keynote stylesheet component {stylesheet_archive} is not registered"
                    ))
                },
            )?;
        let objects = text_box_objects(
            ids,
            context.slide_id,
            styles.shape,
            geometry,
            storage,
            path_source,
            false,
        )?;

        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_slide_drawable_references(
            &mut staged,
            &archive_name,
            context.slide_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, component_id, &ids.all())?;
        add_component_external_reference(
            &mut staged,
            component_id,
            stylesheet_component_id,
            styles.shape,
        )?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .slide_shapes(slide_index)?
            .into_iter()
            .find(|shape| shape.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote shape creation failed validation".to_owned())
            })?;
        let created_graph = shape_graph(&verified, slide_index, ids.drawable)?;
        let line_matches = match (created.line_segment, expected_line) {
            (Some(actual), Some(expected)) => line_segments_match(actual, expected),
            (None, None) => true,
            _ => false,
        };
        if created.preset != expected_preset
            || !line_matches
            || created.storage.object_id != ids.storage
            || created.storage.text != text
            || created.geometry != geometry
            || created_graph.object_ids != ids.all()
        {
            return Err(Error::InvalidFormat(
                "Keynote shape creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the slide-space endpoints of one native straight line.
    pub fn slide_line_segment(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<LineSegment> {
        shape_graph(self, slide_index, drawable_object_id)?
            .info
            .line_segment
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote drawable {drawable_object_id} is not a native straight line"
                ))
            })
    }

    /// Read the directed start and end decorations of one native straight line.
    pub fn slide_line_endpoints(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<LineEndpoints> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        source.info.line_endpoints.ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote drawable {drawable_object_id} is not a native straight line"
            ))
        })
    }

    /// Replace both decorations transactionally, using copy-on-write when the style is shared.
    pub fn set_slide_line_endpoints(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        endpoints: LineEndpoints,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Keynote drawable {drawable_object_id} is not a native straight line"
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
        if verified.slide_line_endpoints(slide_index, drawable_object_id)? != endpoints {
            return Err(Error::InvalidFormat(
                "Keynote line endpoint-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete explicit endpoint decorations and restore undecorated style inheritance.
    ///
    /// Returns `true` when decorations were reset and `false` when the line was
    /// already undecorated.
    pub fn reset_slide_line_endpoints(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        if self.slide_line_endpoints(slide_index, drawable_object_id)? == LineEndpoints::default() {
            return Ok(false);
        }
        self.set_slide_line_endpoints(slide_index, drawable_object_id, LineEndpoints::default())?;
        Ok(true)
    }

    /// Read the effective standard stroke of one ordinary slide shape.
    pub fn slide_shape_stroke(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<ShapeStroke>> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        shape_stroke(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's stroke transactionally, using copy-on-write for shared styles.
    pub fn set_slide_shape_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        stroke: ShapeStroke,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_stroke(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            stroke,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_stroke(slide_index, drawable_object_id)? != Some(stroke) {
            return Err(Error::InvalidFormat(
                "Keynote shape stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct stroke override and restore the inherited appearance.
    pub fn reset_slide_shape_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_stroke(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective standard fill of one ordinary slide shape.
    pub fn slide_shape_fill(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ShapeFill> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        shape_fill(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's fill transactionally, using copy-on-write for shared styles.
    pub fn set_slide_shape_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        fill: &ShapeFill,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_fill(&mut staged, &source.archive_name, drawable_object_id, fill)?;
        let verified = Self::from_package(staged)?;
        if &verified.slide_shape_fill(slide_index, drawable_object_id)? != fill {
            return Err(Error::InvalidFormat(
                "Keynote shape fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Embed image bytes and use them as a simple or tinted native shape fill.
    pub fn set_slide_shape_image_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let fill_size = source.info.geometry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} has no image-fill dimensions"
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
        if verified.slide_shape_fill(slide_index, drawable_object_id)?
            != ShapeFill::Image(image.clone())
        {
            return Err(Error::InvalidFormat(
                "Keynote shape image-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(image)
    }

    /// Remove a direct fill override and restore the inherited appearance.
    pub fn reset_slide_shape_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_fill(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective whole-object opacity and reflection settings.
    pub fn slide_shape_effects(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ShapeEffects> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        shape_effects(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace opacity and reflection atomically while preserving other style properties.
    pub fn set_slide_shape_effects(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        effects: ShapeEffects,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_effects(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            effects,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_effects(slide_index, drawable_object_id)? != effects {
            return Err(Error::InvalidFormat(
                "Keynote shape effect update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct effect overrides and restore inherited values.
    pub fn reset_slide_shape_effects(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        let changed = reset_shape_effects(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective drop, contact, curved, or disabled shadow state.
    pub fn slide_shape_shadow(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ShapeShadow> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        shape_shadow(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace the shadow while preserving all other shape style properties.
    pub fn set_slide_shape_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shadow: ShapeShadow,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let staged = set_shape_shadow(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
            shadow,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_shadow(slide_index, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Keynote shape shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct shadow override and restore inherited shadow state.
    pub fn reset_slide_shape_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
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
    pub fn slide_shape_text_layout(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Layout> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        shape_text_layout(self.package(), &source.archive_name, drawable_object_id)
    }

    /// Replace frame-level text layout while preserving drawing style and columns.
    pub fn set_slide_shape_text_layout(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        layout: Layout,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package().clone(),
            &source.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_text_layout(slide_index, drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Keynote shape text-layout update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct frame-level text-layout overrides and restore inherited values.
    pub fn reset_slide_shape_text_layout(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
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
    pub fn set_slide_line_segment(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Keynote drawable {drawable_object_id} is not a native straight line"
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
        let actual = verified.slide_line_segment(slide_index, drawable_object_id)?;
        if !line_segments_match(actual, replacement) {
            return Err(Error::InvalidFormat(
                "Keynote line endpoint update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the recognized preset and native controls for one slide shape.
    pub fn slide_shape_preset(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<Option<ShapePreset>> {
        Ok(shape_graph(self, slide_index, drawable_object_id)?
            .info
            .preset)
    }

    /// Replace a slide shape's preset path while retaining its text and style.
    pub fn set_slide_shape_preset(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        preset: ShapePreset,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_preset(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            preset,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_preset(slide_index, drawable_object_id)? != Some(preset) {
            return Err(Error::InvalidFormat(
                "Keynote shape preset update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Apply one native Arrange Flip operation to an ordinary slide shape.
    ///
    /// Returns the updated geometry after applying the same transform as the
    /// Keynote Flip Horizontally or Flip Vertically inspector button.
    pub fn flip_slide_shape(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let geometry = flip_drawable_geometry(source.info.geometry, axis)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_geometry(slide_index, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Keynote shape flip update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Read typed geometry for one ordinary slide shape.
    pub fn slide_shape_geometry(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(shape_graph(self, slide_index, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update geometry while preserving unknown shape and path fields.
    pub fn set_slide_shape_geometry(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_geometry(slide_index, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Keynote shape geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary slide shape.
    pub fn slide_shape_properties(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(shape_graph(self, slide_index, drawable_object_id)?
            .info
            .properties)
    }

    /// Update shared drawable properties while preserving unknown fields.
    pub fn set_slide_shape_properties(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_shape_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.slide_shape_properties(slide_index, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Keynote shape property update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a UTF-16 range in an ordinary shape's writable text.
    pub fn replace_slide_shape_text(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        shape_graph(self, slide_index, drawable_object_id)?;
        self.replace_slide_text_storage(slide_index, drawable_object_id, range, replacement)?;
        shape_graph(self, slide_index, drawable_object_id)?;
        Ok(())
    }

    /// Replace all text in an ordinary slide shape.
    pub fn set_slide_shape_text(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        shape_graph(self, slide_index, drawable_object_id)?;
        self.set_slide_text_storage(slide_index, drawable_object_id, replacement)?;
        shape_graph(self, slide_index, drawable_object_id)?;
        Ok(())
    }

    /// Clear an ordinary slide shape without deleting it.
    pub fn clear_slide_shape_text(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<()> {
        self.set_slide_shape_text(slide_index, drawable_object_id, "")
    }

    /// Duplicate an ordinary slide shape with independent rich-text storage.
    ///
    /// The shape and its stand-in objects receive fresh identifiers and UUIDs
    /// while retaining the source's styling and unknown protobuf fields. The
    /// clone is added to the same slide and offset using Keynote's native
    /// duplicate placement so both objects remain independently selectable.
    pub fn duplicate_slide_shape(
        &mut self,
        slide_index: usize,
        source_drawable_object_id: u64,
    ) -> Result<KeynoteSlideShapeInfo> {
        let source = shape_graph(self, slide_index, source_drawable_object_id)?;
        let mut staged = self.package().clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len());
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Keynote shape graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Keynote shape object {identifier} is missing"))
                })?;
                clone_slide_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = remap[&source_drawable_object_id];
        let new_storage_id = remap[&source.info.storage.object_id];
        offset_keynote_drawable_clone(
            &mut staged,
            &source.archive_name,
            new_drawable_id,
            DRAWABLE_DUPLICATE_OFFSET,
        )?;
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = remap.values().copied().max().ok_or_else(|| {
            Error::InvalidFormat("Keynote shape graph has no object identifiers".to_owned())
        })?;
        set_package_last_object_identifier(&mut staged, last_identifier)?;
        let new_uuid_object_ids = source
            .uuid_object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        add_component_object_uuids(&mut staged, source.component_id, &new_uuid_object_ids)?;

        let verified = Self::from_package(staged)?;
        let created = verified
            .slide_shapes(slide_index)?
            .into_iter()
            .find(|shape| shape.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote shape duplication failed validation".to_owned())
            })?;
        let created_graph = shape_graph(&verified, slide_index, new_drawable_id)?;
        if created.storage.object_id != new_storage_id
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
                "Keynote shape duplication produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Remove an ordinary shape and its private text-bearing object graph.
    pub fn remove_slide_shape(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideShape> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(crate::comments::DrawableObjectId::from_object_id(
            drawable_object_id,
        )?)?;
        let mut staged = comments.into_package();
        patch_slide_drawable_references(
            &mut staged,
            &source.archive_name,
            source.slide_id,
            Some(drawable_object_id),
            None,
        )?;
        for identifier in &source.object_ids {
            remove_object(&mut staged, &source.archive_name, *identifier)?;
        }
        for identifier in &source.object_ids {
            if package_references_object(&staged, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Keynote shape object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slide_shapes(slide_index)?
            .iter()
            .any(|shape| shape.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Keynote shape deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedKeynoteSlideShape { shape: source.info })
    }
}

fn new_shape_geometry(position: DrawablePoint, size: DrawableSize) -> Result<DrawableGeometry> {
    if !position.x.is_finite()
        || !position.y.is_finite()
        || !size.width.is_finite()
        || !size.height.is_finite()
        || size.width <= 0.0
        || size.height <= 0.0
    {
        return Err(Error::ParseError(
            "Keynote shape position must be finite and size must be finite and positive".to_owned(),
        ));
    }
    DrawableGeometry {
        position: Some(position),
        size: Some(size),
        flags: Some(DEFAULT_DRAWABLE_FLAGS),
        angle: Some(DEFAULT_ROTATION_DEGREES),
    }
    .validate()
}

fn shape_infos(editor: &KeynoteEditor, slide_index: usize) -> Result<Vec<KeynoteSlideShapeInfo>> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let native: kn::SlideArchive = graph.decode_type(slide.slide_id, 5, "KN.SlideArchive")?;
    let mut result = Vec::new();
    for reference in &native.owned_drawables {
        let Some(messages) = graph.objects.get(&reference.identifier) else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide drawable {} is missing",
                reference.identifier
            )));
        };
        if !messages
            .iter()
            .any(|message| message.type_ == SHAPE_MESSAGE_TYPE)
        {
            continue;
        }
        let shape: tswp::ShapeInfoArchive = graph.decode_type(
            reference.identifier,
            SHAPE_MESSAGE_TYPE,
            "TSWP.ShapeInfoArchive",
        )?;
        if shape.is_text_box == Some(false) {
            result.push(shape_info(
                editor,
                &graph,
                slide_index,
                reference.identifier,
                &shape,
            )?);
        }
    }
    Ok(result)
}

fn shape_graph(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<SlideShapeGraph> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let native: kn::SlideArchive = graph.decode_type(slide.slide_id, 5, "KN.SlideArchive")?;
    for (name, references) in [
        ("owned_drawables", &native.owned_drawables),
        ("drawables_z_order", &native.drawables_z_order),
    ] {
        let matches = references
            .iter()
            .filter(|reference| reference.identifier == drawable_object_id)
            .count();
        if matches != 1 {
            return Err(Error::ParseError(format!(
                "Keynote slide {slide_index} {name} does not own shape {drawable_object_id} exactly once"
            )));
        }
    }
    let shape: tswp::ShapeInfoArchive = graph.decode_type(
        drawable_object_id,
        SHAPE_MESSAGE_TYPE,
        "TSWP.ShapeInfoArchive",
    )?;
    if shape.is_text_box != Some(false) {
        return Err(Error::ParseError(format!(
            "Keynote drawable {drawable_object_id} is not an ordinary shape"
        )));
    }
    let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
    if graph.archive_name(drawable_object_id)? != archive_name {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} is outside slide component {archive_name}"
        )));
    }
    let info = shape_info(editor, &graph, slide_index, drawable_object_id, &shape)?;
    let required = required_shape_objects(&graph, drawable_object_id, &shape)?;
    let archive = editor.package().archive(&archive_name)?;
    let object_ids = slide_create::graph::private_clone_object_ids(
        &archive,
        [drawable_object_id],
        "slide shape",
    )?;
    let actual = object_ids.iter().copied().collect::<HashSet<_>>();
    if actual != required || object_ids.len() != required.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} does not have an isolated private graph"
        )));
    }
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    Ok(SlideShapeGraph {
        slide_id: slide.slide_id,
        component_id,
        archive_name,
        info,
        object_ids,
        uuid_object_ids,
    })
}

#[allow(deprecated)]
fn shape_info(
    editor: &KeynoteEditor,
    graph: &ObjectGraph,
    slide_index: usize,
    drawable_object_id: u64,
    shape: &tswp::ShapeInfoArchive,
) -> Result<KeynoteSlideShapeInfo> {
    let storage_id = shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} has no writable storage"
            ))
        })?;
    if shape
        .deprecated_storage
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(storage_id)
        || graph.archive_name(storage_id)? != graph.archive_name(drawable_object_id)?
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} has inconsistent storage ownership"
        )));
    }
    let archive_name = graph.archive_name(drawable_object_id)?;
    let line_segment = shape_line_segment(shape)?;
    let line_endpoints = line_segment
        .map(|_| shape_line_endpoints(editor.package(), archive_name, drawable_object_id))
        .transpose()?;
    Ok(KeynoteSlideShapeInfo {
        slide_index,
        drawable_object_id,
        kind: shape_path_kind(shape)?,
        preset: shape_preset(shape)?,
        line_segment,
        line_endpoints,
        storage: editor.text.storage(storage_id)?,
        geometry: shape_geometry(
            editor.package(),
            graph.archive_name(drawable_object_id)?,
            drawable_object_id,
        )?,
        properties: shape_properties(
            editor.package(),
            graph.archive_name(drawable_object_id)?,
            drawable_object_id,
        )?,
    })
}

fn required_shape_objects(
    graph: &ObjectGraph,
    drawable_object_id: u64,
    shape: &tswp::ShapeInfoArchive,
) -> Result<HashSet<u64>> {
    let required_reference = |reference: Option<&tsp::Reference>, label: &str| {
        reference
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote shape {drawable_object_id} has no {label} object"
                ))
            })
    };
    let caption = caption_slot_from_reference(
        graph,
        drawable_object_id,
        shape.super_.super_.caption,
        DrawableCaptionKind::Caption,
    )?;
    let title = caption_slot_from_reference(
        graph,
        drawable_object_id,
        shape.super_.super_.title,
        DrawableCaptionKind::Title,
    )?;
    let storage_id = required_reference(shape.owned_storage.as_ref(), "storage")?;
    let storage = graph.objects.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} storage {storage_id} is missing"
        ))
    })?;
    if storage
        .iter()
        .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} storage {storage_id} must have exactly one writable payload"
        )));
    }
    let mut identifiers = vec![drawable_object_id];
    identifiers.extend(caption.object_ids);
    identifiers.extend(title.object_ids);
    identifiers.push(storage_id);
    let required = identifiers.iter().copied().collect::<HashSet<_>>();
    if required.len() != identifiers.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} aliases private objects"
        )));
    }
    Ok(required)
}

#[cfg(test)]
mod tests {
    use std::{fs, path::PathBuf};

    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{
        LineEndpoint, RgbColorSpace, RgbaColor, ShapeCornerRadius, ShapeCurvedShadow,
        ShapeGradient, ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity,
        ShapeGradientStop, ShapeGradientStopMidpoint, ShapeGradientStopPosition, ShapeOpacity,
        ShapePolygonSides, ShapeReflection, ShapeReflectionOpacity, ShapeShadowAngle,
        ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowCurve, ShapeShadowOffset,
        ShapeShadowOpacity, ShapeStarInnerRatio, ShapeStarPoints, StrokePattern, StrokeWidth,
    };
    use crate::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};

    const POSITION: DrawablePoint = DrawablePoint { x: 320.0, y: 240.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 480.0,
        height: 240.0,
    };
    const LINE_START: DrawablePoint = DrawablePoint { x: 720.0, y: 660.0 };
    const LINE_END: DrawablePoint = DrawablePoint {
        x: 1_200.0,
        y: 900.0,
    };
    const UPDATED_LINE_START: DrawablePoint = DrawablePoint { x: 96.0, y: 108.0 };
    const UPDATED_LINE_END: DrawablePoint = DrawablePoint { x: 456.0, y: 108.0 };

    fn fixture(relative: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        fs::read(root.join(relative)).unwrap()
    }

    #[test]
    fn scratch_presentation_supports_rectangle_crud_without_a_source_drawable() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch shape")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        assert!(editor.slide_shapes(0).unwrap().is_empty());
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_slide_rectangle(0, "Built from typed objects", POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, KeynoteSlideShapeKind::Rectangle);
        assert_eq!(created.preset, Some(ShapePreset::Rectangle));
        assert_eq!(created.storage.text, "Built from typed objects");
        assert_eq!(
            editor
                .slide_text_storages(0)
                .unwrap()
                .into_iter()
                .find(|text| text.drawable_object_id == created.drawable_object_id)
                .unwrap()
                .role,
            KeynoteSlideTextRole::Shape
        );
        let horizontally_flipped = editor
            .flip_slide_shape(0, created.drawable_object_id, DrawableFlipAxis::Horizontal)
            .unwrap();
        assert_eq!(
            editor
                .slide_shape_geometry(0, created.drawable_object_id)
                .unwrap(),
            horizontally_flipped
        );
        assert_ne!(horizontally_flipped.flags, created.geometry.flags);
        let vertically_flipped = editor
            .flip_slide_shape(0, created.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();
        assert_eq!(
            editor
                .slide_shape_geometry(0, created.drawable_object_id)
                .unwrap(),
            vertically_flipped
        );
        assert_ne!(vertically_flipped.angle, created.geometry.angle);

        editor
            .replace_slide_shape_text(0, created.drawable_object_id, 0..5, "Made")
            .unwrap();
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 96.0, y: 108.0 }),
            size: Some(DrawableSize {
                width: 360.0,
                height: 180.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(12.0),
        };
        editor
            .set_slide_shape_geometry(0, created.drawable_object_id, geometry)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Generated rectangle".to_owned()),
        };
        editor
            .set_slide_shape_properties(0, created.drawable_object_id, properties.clone())
            .unwrap();

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let shape = &reopened.slide_shapes(0).unwrap()[0];
        assert_eq!(shape.storage.text, "Made from typed objects");
        assert_eq!(shape.geometry, geometry);
        assert_eq!(shape.properties, properties);

        let removed = editor
            .remove_slide_shape(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_presentation_supports_native_shape_duplication() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Shape clone")
            .subtitle("Independent rich text")
            .build()
            .unwrap();
        let created = editor
            .add_slide_shape(0, "Source shape", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.8, 0.35, 0.2, 1.0, RgbColorSpace::Srgb).unwrap());
        editor
            .set_slide_shape_fill(0, created.drawable_object_id, &fill)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/keynote-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(false),
            accessibility_description: Some("Source shape".to_owned()),
        };
        editor
            .set_slide_shape_properties(0, created.drawable_object_id, properties.clone())
            .unwrap();
        editor
            .flip_slide_shape(0, created.drawable_object_id, DrawableFlipAxis::Vertical)
            .unwrap();
        let source = editor.slide_shapes(0).unwrap().into_iter().next().unwrap();

        let duplicate = editor
            .duplicate_slide_shape(0, source.drawable_object_id)
            .unwrap();
        assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
        assert_ne!(duplicate.storage.object_id, source.storage.object_id);
        assert_eq!(duplicate.storage.text, source.storage.text);
        assert_eq!(duplicate.kind, source.kind);
        assert_eq!(duplicate.preset, source.preset);
        assert_eq!(duplicate.properties, properties);
        assert_eq!(
            duplicate.geometry.position,
            source.geometry.position.map(|position| DrawablePoint {
                x: position.x + DRAWABLE_DUPLICATE_OFFSET,
                y: position.y + DRAWABLE_DUPLICATE_OFFSET,
            })
        );
        assert_eq!(
            editor
                .slide_shape_fill(0, duplicate.drawable_object_id)
                .unwrap(),
            fill
        );

        editor
            .set_slide_shape_text(0, duplicate.drawable_object_id, "Independent copy")
            .unwrap();
        assert_eq!(
            editor
                .slide_shapes(0)
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == source.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Source shape"
        );
        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shapes(0)
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Independent copy"
        );
        assert_eq!(reopened.slide_shapes(0).unwrap().len(), 2);

        let removed = editor
            .remove_slide_shape(0, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.storage.text, "Independent copy");
        assert_eq!(editor.slide_shapes(0).unwrap().len(), 1);
    }

    #[test]
    fn scratch_presentation_supports_straight_line_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Scratch line")
            .subtitle("No embedded package")
            .build()
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        let created = editor.add_slide_line(0, LINE_START, LINE_END).unwrap();
        assert_eq!(created.kind, KeynoteSlideShapeKind::Line);
        assert_eq!(created.preset, None);
        assert_eq!(created.storage.text, "");
        assert!(line_segments_match(
            created.line_segment.unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(line_segments_match(
            reopened
                .slide_line_segment(0, created.drawable_object_id)
                .unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        editor
            .set_slide_line_segment(
                0,
                created.drawable_object_id,
                UPDATED_LINE_START,
                UPDATED_LINE_END,
            )
            .unwrap();
        assert!(line_segments_match(
            editor
                .slide_line_segment(0, created.drawable_object_id)
                .unwrap(),
            LineSegment::new(UPDATED_LINE_START, UPDATED_LINE_END).unwrap()
        ));

        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_line_segment(
                    0,
                    created.drawable_object_id,
                    UPDATED_LINE_START,
                    UPDATED_LINE_START,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        let removed = editor
            .remove_slide_shape(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.kind, KeynoteSlideShapeKind::Line);
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let rectangle = editor
            .add_slide_rectangle(0, "Not a line", POSITION, SIZE)
            .unwrap();
        let before_cross_type = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_slide_shape(0, u64::MAX, DrawableFlipAxis::Horizontal)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
        assert!(
            editor
                .set_slide_line_segment(0, rectangle.drawable_object_id, LINE_START, LINE_END,)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
    }

    #[test]
    fn scratch_presentation_supports_typed_shape_stroke_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Stroke")
            .subtitle("Typed native style")
            .build()
            .unwrap();
        let stroke = ShapeStroke::new(
            RgbaColor::new(0.95, 0.45, 0.05, 0.85, RgbColorSpace::DisplayP3).unwrap(),
            StrokeWidth::new(5.0).unwrap(),
            StrokePattern::ShortDash,
        );
        let endpoints = LineEndpoints::new(LineEndpoint::FilledDiamond, LineEndpoint::SimpleArrow);
        let created = editor
            .add_slide_line_with_style(
                0,
                LINE_START,
                LINE_END,
                LineStyle::new(stroke).with_endpoints(endpoints),
            )
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_stroke(0, created.drawable_object_id)
                .unwrap(),
            Some(stroke)
        );
        assert_eq!(
            reopened
                .slide_line_endpoints(0, created.drawable_object_id)
                .unwrap(),
            endpoints
        );
        assert!(
            reopened
                .reset_slide_shape_stroke(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_stroke(0, created.drawable_object_id)
                .unwrap(),
            None
        );
        assert_eq!(
            reopened
                .slide_line_endpoints(0, created.drawable_object_id)
                .unwrap(),
            endpoints
        );
    }

    #[test]
    fn scratch_presentation_supports_typed_shape_fill_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Fill")
            .subtitle("Typed native style")
            .build()
            .unwrap();
        let orange = RgbaColor::new(0.95, 0.45, 0.05, 1.0, RgbColorSpace::DisplayP3).unwrap();
        let purple = RgbaColor::new(0.45, 0.1, 0.8, 1.0, RgbColorSpace::DisplayP3).unwrap();
        let fill = ShapeFill::Gradient(
            ShapeGradient::advanced(
                ShapeGradientKind::Radial,
                vec![
                    ShapeGradientStop::new(
                        orange,
                        ShapeGradientStopPosition::START,
                        ShapeGradientStopMidpoint::new(0.35).unwrap(),
                    ),
                    ShapeGradientStop::new(
                        purple,
                        ShapeGradientStopPosition::END,
                        ShapeGradientStopMidpoint::CENTER,
                    ),
                ],
                ShapeGradientOpacity::new(0.85).unwrap(),
                ShapeGradientAngle::from_degrees(315.0).unwrap(),
            )
            .unwrap(),
        );
        let created = editor
            .add_slide_shape(0, "Filled", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let inherited_fill = editor
            .slide_shape_fill(0, created.drawable_object_id)
            .unwrap();
        editor
            .set_slide_shape_fill(0, created.drawable_object_id, &fill)
            .unwrap();
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_fill(0, created.drawable_object_id)
                .unwrap(),
            fill
        );
        assert!(
            reopened
                .reset_slide_shape_fill(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_fill(0, created.drawable_object_id)
                .unwrap(),
            inherited_fill
        );
    }

    #[test]
    fn scratch_presentation_supports_embedded_shape_image_fill_crud() {
        let bytes = fixture("test-data/images/png/lena.png");
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Image fill")
            .subtitle("Embedded from scratch")
            .build()
            .unwrap();
        let created = editor
            .add_slide_shape(0, "Image", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let inherited = editor
            .slide_shape_fill(0, created.drawable_object_id)
            .unwrap();
        let tint = RgbaColor::new(0.1, 0.3, 0.75, 0.5, RgbColorSpace::Srgb).unwrap();
        let image = editor
            .set_slide_shape_image_fill(
                0,
                created.drawable_object_id,
                "lena.png",
                &bytes,
                ShapeImageFillTechnique::Tile,
                Some(tint),
            )
            .unwrap();
        assert_eq!(image.tint(), Some(tint));
        assert_eq!(
            editor
                .extract_media(image.data_identifier().unwrap().get())
                .unwrap(),
            bytes
        );

        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_fill(0, created.drawable_object_id)
                .unwrap(),
            ShapeFill::Image(image)
        );
        assert!(
            reopened
                .reset_slide_shape_fill(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_fill(0, created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(reopened.media_assets().unwrap().is_empty());
    }

    #[test]
    fn scratch_presentation_supports_shape_effect_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Effects")
            .subtitle("Typed object opacity and reflection")
            .build()
            .unwrap();
        let created = editor
            .add_slide_shape(0, "Effects", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let inherited = editor
            .slide_shape_effects(0, created.drawable_object_id)
            .unwrap();
        let effects = ShapeEffects::new(
            ShapeOpacity::new(0.61).unwrap(),
            ShapeReflection::Enabled(ShapeReflectionOpacity::new(0.2).unwrap()),
        );
        editor
            .set_slide_shape_effects(0, created.drawable_object_id, effects)
            .unwrap();

        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_effects(0, created.drawable_object_id)
                .unwrap(),
            effects
        );
        assert!(
            reopened
                .reset_slide_shape_effects(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_effects(0, created.drawable_object_id)
                .unwrap(),
            inherited
        );
    }

    #[test]
    fn scratch_presentation_supports_curved_shadow_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Shadows")
            .subtitle("Typed curved shadow")
            .build()
            .unwrap();
        let created = editor
            .add_slide_shape(0, "Shadow", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let inherited = editor
            .slide_shape_shadow(0, created.drawable_object_id)
            .unwrap();
        let shadow = ShapeShadow::Curved(ShapeCurvedShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::black(),
                ShapeShadowBlurRadius::from_points(15).unwrap(),
                ShapeShadowOffset::from_points(4.0).unwrap(),
                ShapeShadowOpacity::new(0.73).unwrap(),
            ),
            ShapeShadowAngle::from_degrees(310.0).unwrap(),
            ShapeShadowCurve::new(0.2).unwrap(),
        ));
        editor
            .set_slide_shape_shadow(0, created.drawable_object_id, shadow)
            .unwrap();

        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_shadow(0, created.drawable_object_id)
                .unwrap(),
            shadow
        );
        assert!(
            reopened
                .reset_slide_shape_shadow(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_shadow(0, created.drawable_object_id)
                .unwrap(),
            inherited
        );
    }

    #[test]
    fn scratch_presentation_supports_shape_text_layout_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Layout")
            .subtitle("Typed frame layout")
            .build()
            .unwrap();
        let created = editor
            .add_slide_shape(0, "Layout", POSITION, SIZE, ShapePreset::Rectangle)
            .unwrap();
        let inherited = editor
            .slide_shape_text_layout(0, created.drawable_object_id)
            .unwrap();
        let layout = Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(14.0).unwrap()),
            AutoSize::ShrinkToFit,
        );
        editor
            .set_slide_shape_text_layout(0, created.drawable_object_id, layout)
            .unwrap();

        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_text_layout(0, created.drawable_object_id)
                .unwrap(),
            layout
        );
        assert!(
            reopened
                .reset_slide_shape_text_layout(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_shape_text_layout(0, created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(
            !reopened
                .reset_slide_shape_text_layout(0, created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_presentation_supports_typed_line_endpoint_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Endpoint styles")
            .subtitle("Built from native line-end paths")
            .build()
            .unwrap();
        let created = editor
            .add_slide_line_with_endpoints(
                0,
                LINE_START,
                LINE_END,
                LineEndpoints::new(LineEndpoint::OpenSquare, LineEndpoint::FilledDiamond),
            )
            .unwrap();
        assert_eq!(
            created.line_endpoints,
            Some(LineEndpoints::new(
                LineEndpoint::OpenSquare,
                LineEndpoint::FilledDiamond
            ))
        );

        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let replacement = LineEndpoints::new(LineEndpoint::SimpleArrow, LineEndpoint::OpenCircle);
        reopened
            .set_slide_line_endpoints(0, created.drawable_object_id, replacement)
            .unwrap();
        assert_eq!(
            reopened
                .slide_line_endpoints(0, created.drawable_object_id)
                .unwrap(),
            replacement
        );
        assert!(
            reopened
                .reset_slide_line_endpoints(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_line_endpoints(0, created.drawable_object_id)
                .unwrap(),
            LineEndpoints::default()
        );
    }

    #[test]
    fn scratch_presentation_supports_typed_preset_shape_crud() {
        let mut editor = KeynoteDocumentBuilder::new()
            .title("Preset shapes")
            .subtitle("Built from typed paths")
            .build()
            .unwrap();
        let baseline = editor.to_bytes().unwrap();
        let created = editor
            .add_slide_shape(0, "Rounded", POSITION, SIZE, ShapePreset::ROUNDED_RECTANGLE)
            .unwrap();
        assert_eq!(created.kind, KeynoteSlideShapeKind::RoundedRectangle);
        assert_eq!(created.preset, Some(ShapePreset::ROUNDED_RECTANGLE));

        let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_shape_preset(0, created.drawable_object_id)
                .unwrap(),
            Some(ShapePreset::ROUNDED_RECTANGLE)
        );

        for (preset, kind) in [
            (ShapePreset::Ellipse, KeynoteSlideShapeKind::Ellipse),
            (ShapePreset::LeftArrow, KeynoteSlideShapeKind::LeftArrow),
            (ShapePreset::RightArrow, KeynoteSlideShapeKind::RightArrow),
            (ShapePreset::DoubleArrow, KeynoteSlideShapeKind::DoubleArrow),
            (ShapePreset::PENTAGON, KeynoteSlideShapeKind::RegularPolygon),
            (ShapePreset::STAR, KeynoteSlideShapeKind::Star),
        ] {
            editor
                .set_slide_shape_preset(0, created.drawable_object_id, preset)
                .unwrap();
            let shape = &editor.slide_shapes(0).unwrap()[0];
            assert_eq!(shape.kind, kind);
            assert_eq!(shape.preset, Some(preset));
            assert_eq!(shape.storage.text, "Rounded");
        }

        editor
            .remove_slide_shape(0, created.drawable_object_id)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn invalid_rectangle_creation_and_cross_type_updates_are_transactional() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .add_slide_rectangle(
                    0,
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
        assert!(ShapeCornerRadius::new(f32::NAN).is_err());
        assert!(ShapePolygonSides::new(2).is_err());
        assert!(ShapeStarPoints::new(2).is_err());
        assert!(ShapeStarInnerRatio::new(1.0).is_err());
        assert!(
            editor
                .add_slide_shape(
                    0,
                    "invalid radius",
                    POSITION,
                    SIZE,
                    ShapePreset::RoundedRectangle {
                        corner_radius: ShapeCornerRadius::new(SIZE.height).unwrap(),
                    },
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
        let text_box = editor
            .add_slide_text_box(0, "not a shape", POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_slide_shape_text(0, text_box.drawable_object_id, "wrong type")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_slide_shape_preset(0, text_box.drawable_object_id, ShapePreset::Ellipse)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}

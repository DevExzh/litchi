//! Ordinary text-bearing shape CRUD for Numbers sheets.

use std::collections::{HashMap, HashSet};
use std::ops::Range;

use super::*;
use crate::image_caption::DrawableCaptionKind;
use crate::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, Endpoints,
    LineSegment, LineStyle, RgbaColor, Shadow, ShapeFill, ShapeImageFill, ShapeImageFillTechnique,
    Stroke, flip_drawable_geometry, line_geometry, line_path_source,
    line_segments_match, reset_shape_effects, reset_shape_fill, reset_shape_shadow,
    reset_shape_stroke, reset_shape_text_layout, set_shape_effects, set_shape_fill,
    set_shape_geometry, set_shape_image_fill_data, set_shape_line_endpoints,
    set_shape_line_segment, set_shape_preset, set_shape_shadow, set_shape_stroke,
    set_shape_text_layout, shape_effects, shape_fill, shape_line_endpoints, shape_line_segment,
    shape_path_kind, shape_path_source, shape_preset, shape_shadow, shape_stroke,
    shape_text_layout,
};
use crate::text::layout::Layout;
use litchi_iwa_common::shape::effects::Effects;
use litchi_iwa_common::shape::path::{Kind, Preset};

use super::text_box_create::{
    TextBoxObjectIds, text_box_objects, text_box_storage, text_box_theme_styles,
};

mod caption;

use caption::*;

const DEFAULT_DRAWABLE_FLAGS: u32 = 3;
const DEFAULT_ROTATION_DEGREES: f32 = 0.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ShapeClonePlacement {
    Offset,
    Preserve,
}

/// One ordinary, non-text-box shape owned directly by a Numbers sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct NumbersSheetShapeInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub kind: Kind,
    /// Source-buildable preset and its native controls, when recognized.
    pub preset: Option<Preset>,
    /// Sheet-space endpoints when this shape is a native straight line.
    pub line_segment: Option<LineSegment>,
    /// Directed start/end decorations when this shape is a native straight line.
    pub line_endpoints: Option<Endpoints>,
    pub storage: TextStorageInfo,
    pub geometry: DrawableGeometry,
    pub properties: DrawableProperties,
}

/// Result of removing an ordinary Numbers shape and its private object graph.
#[derive(Debug, Clone, PartialEq)]
pub struct RemovedNumbersSheetShape {
    pub shape: NumbersSheetShapeInfo,
}

struct SheetShapeGraph {
    sheet_id: u64,
    archive_name: String,
    component_id: u64,
    info: NumbersSheetShapeInfo,
    object_ids: Vec<u64>,
    uuid_object_ids: Vec<u64>,
}

impl NumbersEditor {
    /// List ordinary non-text-box shapes owned by one reachable sheet.
    pub fn sheet_shapes(&self, sheet_id: u64) -> Result<Vec<NumbersSheetShapeInfo>> {
        shape_infos(self, sheet_id)
    }

    /// Add a rectangular shape with independent writable text to one sheet.
    ///
    /// The shape, storage, title/caption stand-ins, sheet ownership, UUIDs,
    /// style relationship, and package high-water mark are built directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_rectangle(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
    ) -> Result<NumbersSheetShapeInfo> {
        self.add_sheet_shape(sheet_id, text, position, size, Preset::Rectangle)
    }

    /// Add a typed preset shape with independent writable text to one sheet.
    ///
    /// The path, storage, title/caption stand-ins, sheet ownership, UUIDs,
    /// style relationship, and package high-water mark are built directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_shape(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: Preset,
    ) -> Result<NumbersSheetShapeInfo> {
        let geometry = new_shape_geometry(position, size)?;
        self.add_sheet_shape_path(
            sheet_id,
            text,
            geometry,
            shape_path_source(preset, size)?,
            Some(preset),
            None,
        )
    }

    /// Add a typed preset shape with an explicit standard fill.
    pub fn add_sheet_shape_with_fill(
        &mut self,
        sheet_id: u64,
        text: &str,
        position: DrawablePoint,
        size: DrawableSize,
        preset: Preset,
        fill: ShapeFill,
    ) -> Result<NumbersSheetShapeInfo> {
        let created = self.add_sheet_shape(sheet_id, text, position, size, preset)?;
        self.set_sheet_shape_fill(sheet_id, created.drawable_object_id, &fill)?;
        Ok(shape_graph(self, sheet_id, created.drawable_object_id)?.info)
    }

    /// Add a native straight line between two sheet-space points.
    ///
    /// The line path, empty writable storage, stand-ins, ownership, UUIDs,
    /// style relationship, and package high-water mark are built directly
    /// from typed values. No source drawable or package template is copied.
    pub fn add_sheet_line(
        &mut self,
        sheet_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<NumbersSheetShapeInfo> {
        let segment = LineSegment::new(start, end)?;
        self.add_sheet_shape_path(
            sheet_id,
            "",
            line_geometry(segment),
            line_path_source(segment),
            None,
            Some(segment),
        )
    }

    /// Add a native straight line with independently typed start and end decorations.
    pub fn add_sheet_line_with_endpoints(
        &mut self,
        sheet_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
        endpoints: Endpoints,
    ) -> Result<NumbersSheetShapeInfo> {
        let created = self.add_sheet_line(sheet_id, start, end)?;
        self.set_sheet_line_endpoints(sheet_id, created.drawable_object_id, endpoints)?;
        Ok(shape_graph(self, sheet_id, created.drawable_object_id)?.info)
    }

    /// Add a native straight line with a typed stroke and endpoint appearance.
    pub fn add_sheet_line_with_style(
        &mut self,
        sheet_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
        style: LineStyle,
    ) -> Result<NumbersSheetShapeInfo> {
        let created = self.add_sheet_line(sheet_id, start, end)?;
        self.set_sheet_shape_stroke(sheet_id, created.drawable_object_id, style.stroke)?;
        self.set_sheet_line_endpoints(sheet_id, created.drawable_object_id, style.endpoints)?;
        Ok(shape_graph(self, sheet_id, created.drawable_object_id)?.info)
    }

    fn add_sheet_shape_path(
        &mut self,
        sheet_id: u64,
        text: &str,
        geometry: DrawableGeometry,
        path_source: tsd::PathSourceArchive,
        expected_preset: Option<Preset>,
        expected_line: Option<LineSegment>,
    ) -> Result<NumbersSheetShapeInfo> {
        let (archive_name, _, _) = numbers_sheet(&self.package, sheet_id)?;
        let document = numbers_document(&self.package)?;
        let styles = text_box_theme_styles(
            &self.package,
            document.theme.identifier,
            document.stylesheet.identifier,
        )?;
        let storage = text_box_storage(text, &styles, document.super_.document_language.as_deref());
        let ids = TextBoxObjectIds::allocate(next_object_identifier(&self.package)?)?;
        let objects = text_box_objects(
            ids,
            sheet_id,
            styles.shape,
            geometry,
            storage,
            path_source,
            false,
        )?;
        let component_id = component_identifier_for_entry(&self.package, &archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet component {archive_name} is not registered"
                ))
            })?;
        let style_archive = object_locations(&self.package)?
            .get(&styles.shape)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers shape style {} is missing", styles.shape))
            })?;
        let style_component_id = component_identifier_for_entry(&self.package, &style_archive)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers stylesheet component {style_archive} is not registered"
                ))
            })?;

        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &archive_name,
            sheet_id,
            None,
            Some(ids.drawable),
        )?;
        add_component_object_uuids(&mut staged, component_id, &ids.all())?;
        if component_id != style_component_id {
            add_component_external_reference(
                &mut staged,
                component_id,
                style_component_id,
                styles.shape,
            )?;
        }
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = verified
            .sheet_shapes(sheet_id)?
            .into_iter()
            .find(|shape| shape.drawable_object_id == ids.drawable)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers shape creation failed validation".to_owned())
            })?;
        let created_graph = shape_graph(&verified, sheet_id, ids.drawable)?;
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
                "Numbers shape creation produced an inconsistent graph".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Read the sheet-space endpoints of one native straight line.
    pub fn sheet_line_segment(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<LineSegment> {
        shape_graph(self, sheet_id, drawable_object_id)?
            .info
            .line_segment
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers drawable {drawable_object_id} is not a native straight line"
                ))
            })
    }

    /// Read the directed start and end decorations of one native straight line.
    pub fn sheet_line_endpoints(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Endpoints> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        source.info.line_endpoints.ok_or_else(|| {
            Error::ParseError(format!(
                "Numbers drawable {drawable_object_id} is not a native straight line"
            ))
        })
    }

    /// Replace both decorations transactionally, using copy-on-write when the style is shared.
    pub fn set_sheet_line_endpoints(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        endpoints: Endpoints,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Numbers drawable {drawable_object_id} is not a native straight line"
            )));
        }
        let mut staged = self.package.clone();
        set_shape_line_endpoints(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            endpoints,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_line_endpoints(sheet_id, drawable_object_id)? != endpoints {
            return Err(Error::InvalidFormat(
                "Numbers line endpoint-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Delete explicit endpoint decorations and restore undecorated style inheritance.
    ///
    /// Returns `true` when decorations were reset and `false` when the line was
    /// already undecorated.
    pub fn reset_sheet_line_endpoints(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        if self.sheet_line_endpoints(sheet_id, drawable_object_id)? == Endpoints::default() {
            return Ok(false);
        }
        self.set_sheet_line_endpoints(sheet_id, drawable_object_id, Endpoints::default())?;
        Ok(true)
    }

    /// Read the effective standard stroke of one ordinary sheet shape.
    pub fn sheet_shape_stroke(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<Stroke>> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        shape_stroke(&self.package, &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's stroke transactionally, using copy-on-write for shared styles.
    pub fn set_sheet_shape_stroke(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        stroke: Stroke,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_stroke(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            stroke,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_stroke(sheet_id, drawable_object_id)? != Some(stroke) {
            return Err(Error::InvalidFormat(
                "Numbers shape stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct stroke override and restore the inherited appearance.
    pub fn reset_sheet_shape_stroke(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        let changed = reset_shape_stroke(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective standard fill of one ordinary sheet shape.
    pub fn sheet_shape_fill(&self, sheet_id: u64, drawable_object_id: u64) -> Result<ShapeFill> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        shape_fill(&self.package, &source.archive_name, drawable_object_id)
    }

    /// Replace one shape's fill transactionally, using copy-on-write for shared styles.
    pub fn set_sheet_shape_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        fill: &ShapeFill,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_fill(&mut staged, &source.archive_name, drawable_object_id, fill)?;
        let verified = Self::from_package(staged)?;
        if &verified.sheet_shape_fill(sheet_id, drawable_object_id)? != fill {
            return Err(Error::InvalidFormat(
                "Numbers shape fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Embed image bytes and use them as a simple or tinted native shape fill.
    pub fn set_sheet_shape_image_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let fill_size = source.info.geometry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers shape {drawable_object_id} has no image-fill dimensions"
            ))
        })?;
        let mut staged = self.package.clone();
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
        if verified.sheet_shape_fill(sheet_id, drawable_object_id)?
            != ShapeFill::Image(image.clone())
        {
            return Err(Error::InvalidFormat(
                "Numbers shape image-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(image)
    }

    /// Remove a direct fill override and restore the inherited appearance.
    pub fn reset_sheet_shape_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        let changed = reset_shape_fill(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective whole-object opacity and reflection settings.
    pub fn sheet_shape_effects(&self, sheet_id: u64, drawable_object_id: u64) -> Result<Effects> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        shape_effects(&self.package, &source.archive_name, drawable_object_id)
    }

    /// Replace opacity and reflection atomically while preserving other style properties.
    pub fn set_sheet_shape_effects(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        effects: Effects,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_effects(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            effects,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_effects(sheet_id, drawable_object_id)? != effects {
            return Err(Error::InvalidFormat(
                "Numbers shape effect update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct effect overrides and restore inherited values.
    pub fn reset_sheet_shape_effects(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        let changed = reset_shape_effects(&mut staged, &source.archive_name, drawable_object_id)?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read the effective drop, contact, curved, or disabled shadow state.
    pub fn sheet_shape_shadow(&self, sheet_id: u64, drawable_object_id: u64) -> Result<Shadow> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        shape_shadow(&self.package, &source.archive_name, drawable_object_id)
    }

    /// Replace the shadow while preserving all other shape style properties.
    pub fn set_sheet_shape_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        shadow: Shadow,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let staged = set_shape_shadow(
            self.package.clone(),
            &source.archive_name,
            drawable_object_id,
            shadow,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_shadow(sheet_id, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Numbers shape shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove a direct shadow override and restore inherited shadow state.
    pub fn reset_sheet_shape_shadow(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_shadow(
            self.package.clone(),
            &source.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Read effective vertical alignment, edge insets, and autosizing.
    pub fn sheet_shape_text_layout(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Layout> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        shape_text_layout(&self.package, &source.archive_name, drawable_object_id)
    }

    /// Replace frame-level text layout while preserving drawing style and columns.
    pub fn set_sheet_shape_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        layout: Layout,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let staged = set_shape_text_layout(
            self.package.clone(),
            &source.archive_name,
            drawable_object_id,
            layout,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_text_layout(sheet_id, drawable_object_id)? != layout {
            return Err(Error::InvalidFormat(
                "Numbers shape text-layout update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Remove direct frame-level text-layout overrides and restore inherited values.
    pub fn reset_sheet_shape_text_layout(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let (staged, changed) = reset_shape_text_layout(
            self.package.clone(),
            &source.archive_name,
            drawable_object_id,
        )?;
        if changed {
            *self = Self::from_package(staged)?;
        }
        Ok(changed)
    }

    /// Move or resize one native straight line by replacing its endpoints.
    pub fn set_sheet_line_segment(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        start: DrawablePoint,
        end: DrawablePoint,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        if source.info.line_segment.is_none() {
            return Err(Error::ParseError(format!(
                "Numbers drawable {drawable_object_id} is not a native straight line"
            )));
        }
        let replacement = LineSegment::new(start, end)?;
        let mut staged = self.package.clone();
        set_shape_line_segment(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            replacement,
        )?;
        let verified = Self::from_package(staged)?;
        let actual = verified.sheet_line_segment(sheet_id, drawable_object_id)?;
        if !line_segments_match(actual, replacement) {
            return Err(Error::InvalidFormat(
                "Numbers line endpoint update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the recognized preset and native controls for one sheet shape.
    pub fn sheet_shape_preset(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Option<Preset>> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?.info.preset)
    }

    /// Replace a sheet shape's preset path while retaining its text and style.
    pub fn set_sheet_shape_preset(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        preset: Preset,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_preset(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            preset,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_preset(sheet_id, drawable_object_id)? != Some(preset) {
            return Err(Error::InvalidFormat(
                "Numbers shape preset update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Apply one native Arrange Flip operation to an ordinary sheet shape.
    ///
    /// Returns the updated geometry after applying the same transform as the
    /// Numbers Flip Horizontally or Flip Vertically inspector button.
    pub fn flip_sheet_shape(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        axis: DrawableFlipAxis,
    ) -> Result<DrawableGeometry> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let geometry = flip_drawable_geometry(source.info.geometry, axis)?;
        let mut staged = self.package.clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers shape flip update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(geometry)
    }

    /// Read typed geometry for one ordinary sheet shape.
    pub fn sheet_shape_geometry(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableGeometry> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?
            .info
            .geometry)
    }

    /// Update shape geometry while preserving unknown path fields.
    pub fn set_sheet_shape_geometry(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        geometry: DrawableGeometry,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_geometry(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            geometry,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_geometry(sheet_id, drawable_object_id)? != geometry {
            return Err(Error::InvalidFormat(
                "Numbers shape geometry update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read shared drawable properties for one ordinary sheet shape.
    pub fn sheet_shape_properties(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<DrawableProperties> {
        Ok(shape_graph(self, sheet_id, drawable_object_id)?
            .info
            .properties)
    }

    /// Update shared drawable properties while preserving unknown fields.
    pub fn set_sheet_shape_properties(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        properties: DrawableProperties,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_shape_properties(
            &mut staged,
            &source.archive_name,
            drawable_object_id,
            &properties,
        )?;
        let verified = Self::from_package(staged)?;
        if verified.sheet_shape_properties(sheet_id, drawable_object_id)? != properties {
            return Err(Error::InvalidFormat(
                "Numbers shape property update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Replace a UTF-16 range in an ordinary sheet shape's text.
    pub fn replace_sheet_shape_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.replace_text(source.info.storage.object_id, range, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        shape_graph(&verified, sheet_id, drawable_object_id)?;
        *self = verified;
        Ok(())
    }

    /// Replace all text in an ordinary sheet shape.
    pub fn set_sheet_shape_text(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        replacement: &str,
    ) -> Result<()> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut text = IWorkTextEditor::from_package(self.package.clone());
        text.set_text(source.info.storage.object_id, replacement)?;
        let verified = Self::from_package(text.into_package())?;
        let updated = shape_graph(&verified, sheet_id, drawable_object_id)?;
        if updated.info.storage.text != replacement {
            return Err(Error::InvalidFormat(
                "Numbers shape text update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Clear an ordinary sheet shape without deleting it.
    pub fn clear_sheet_shape_text(&mut self, sheet_id: u64, drawable_object_id: u64) -> Result<()> {
        self.set_sheet_shape_text(sheet_id, drawable_object_id, "")
    }

    /// Duplicate an ordinary sheet shape with independent rich-text storage.
    ///
    /// The shape and its stand-in objects receive fresh identifiers and UUIDs
    /// while retaining the source's styling and unknown protobuf fields. The
    /// clone is added to the same sheet and offset using Numbers' native
    /// duplicate placement so both objects remain independently selectable.
    pub fn duplicate_sheet_shape(
        &mut self,
        sheet_id: u64,
        source_drawable_object_id: u64,
    ) -> Result<NumbersSheetShapeInfo> {
        self.clone_sheet_shape(
            sheet_id,
            source_drawable_object_id,
            sheet_id,
            ShapeClonePlacement::Offset,
        )
    }

    pub(super) fn duplicate_sheet_shape_to_sheet(
        &mut self,
        source_sheet_id: u64,
        source_drawable_object_id: u64,
        target_sheet_id: u64,
    ) -> Result<NumbersSheetShapeInfo> {
        self.clone_sheet_shape(
            source_sheet_id,
            source_drawable_object_id,
            target_sheet_id,
            ShapeClonePlacement::Preserve,
        )
    }

    fn clone_sheet_shape(
        &mut self,
        source_sheet_id: u64,
        source_drawable_object_id: u64,
        target_sheet_id: u64,
        placement: ShapeClonePlacement,
    ) -> Result<NumbersSheetShapeInfo> {
        let source = shape_graph(self, source_sheet_id, source_drawable_object_id)?;
        let (target_archive_name, _, _) = numbers_sheet(&self.package, target_sheet_id)?;
        if source.archive_name != target_archive_name {
            return Err(Error::ParseError(format!(
                "Cannot clone Numbers shape {source_drawable_object_id} across archive components"
            )));
        }
        let mut staged = self.package.clone();
        let first_identifier = next_object_identifier(&staged)?;
        let mut remap = HashMap::with_capacity(source.object_ids.len() + 1);
        remap.insert(source_sheet_id, target_sheet_id);
        for (offset, identifier) in source.object_ids.iter().copied().enumerate() {
            let offset = u64::try_from(offset)
                .map_err(|_| Error::ParseError("Numbers shape graph is too large".to_owned()))?;
            let replacement = first_identifier
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
            remap.insert(identifier, replacement);
        }

        for identifier in &source.object_ids {
            let cloned = {
                let archive = staged.archive(&source.archive_name)?;
                let source_object = archive.object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers shape object {identifier} is missing"))
                })?;
                clone_numbers_drawable_graph_object(source_object, &remap)?
            };
            staged.update_archive(&source.archive_name, |archive| {
                Ok(archive.insert_object(cloned)?)
            })?;
        }

        let new_drawable_id = remap[&source_drawable_object_id];
        let new_storage_id = remap[&source.info.storage.object_id];
        if placement == ShapeClonePlacement::Offset {
            offset_numbers_drawable_clone(
                &mut staged,
                &source.archive_name,
                new_drawable_id,
                DRAWABLE_DUPLICATE_OFFSET,
            )?;
        }
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            target_sheet_id,
            None,
            Some(new_drawable_id),
        )?;
        let last_identifier = source
            .object_ids
            .iter()
            .map(|identifier| remap[identifier])
            .max()
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers shape graph has no object identifiers".to_owned())
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
            .sheet_shapes(target_sheet_id)?
            .into_iter()
            .find(|shape| shape.drawable_object_id == new_drawable_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers shape duplication failed validation".to_owned())
            })?;
        let created_graph = shape_graph(&verified, target_sheet_id, new_drawable_id)?;
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
            || (placement == ShapeClonePlacement::Preserve
                && created.geometry != source.info.geometry)
        {
            return Err(Error::InvalidFormat(
                "Numbers shape duplication produced an inconsistent graph".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(created)
    }

    /// Remove an ordinary shape and its private text-bearing graph.
    pub fn remove_sheet_shape(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<RemovedNumbersSheetShape> {
        let source = shape_graph(self, sheet_id, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package.clone())?;
        comments.clear_comment(litchi_iwa_common::comment::DrawableId::from_raw(
            drawable_object_id,
        )?)?;
        let mut staged = comments.into_package();
        patch_numbers_sheet_drawable_reference(
            &mut staged,
            &source.archive_name,
            source.sheet_id,
            Some(drawable_object_id),
            None,
        )?;
        for identifier in &source.object_ids {
            remove_component_external_references_to_object(
                &mut staged,
                source.component_id,
                *identifier,
            )?;
        }
        staged.update_archive(&source.archive_name, |archive| {
            for identifier in &source.object_ids {
                archive.remove_object(*identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!("Numbers shape object {identifier} is missing"))
                })?;
            }
            Ok(())
        })?;
        let locations = object_locations(&staged)?;
        for identifier in &source.object_ids {
            if package_references_object(&staged, &locations, *identifier)? {
                return Err(Error::InvalidFormat(format!(
                    "Numbers shape object {identifier} remains referenced after deletion"
                )));
            }
        }
        remove_component_object_uuids(&mut staged, source.component_id, &source.uuid_object_ids)?;
        release_package_identifier_suffix(&mut staged, &source.object_ids)?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .sheet_shapes(sheet_id)?
            .iter()
            .any(|shape| shape.drawable_object_id == drawable_object_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers shape deletion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(RemovedNumbersSheetShape { shape: source.info })
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
            "Numbers shape position must be finite and size must be finite and positive".to_owned(),
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

fn shape_infos(editor: &NumbersEditor, sheet_id: u64) -> Result<Vec<NumbersSheetShapeInfo>> {
    let (_, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    let locations = object_locations(editor.package())?;
    let mut shapes = Vec::new();
    for reference in sheet.drawable_infos {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let archive = editor.package().archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} drawable {} is missing",
                reference.identifier
            ))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.is_empty() {
            continue;
        }
        let [message] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers drawable {} has multiple shape payloads",
                reference.identifier
            )));
        };
        let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
        if shape.is_text_box == Some(false) {
            shapes.push(shape_info(
                editor,
                &locations,
                sheet_id,
                reference.identifier,
                &shape,
            )?);
        }
    }
    Ok(shapes)
}

#[allow(deprecated)]
fn shape_graph(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<SheetShapeGraph> {
    let (archive_name, _, sheet) = numbers_sheet(editor.package(), sheet_id)?;
    if sheet
        .drawable_infos
        .iter()
        .filter(|reference| reference.identifier == drawable_object_id)
        .count()
        != 1
    {
        return Err(Error::ParseError(format!(
            "Numbers sheet {sheet_id} does not own shape {drawable_object_id} exactly once"
        )));
    }
    let locations = object_locations(editor.package())?;
    if locations.get(&drawable_object_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} is outside sheet component {archive_name}"
        )));
    }
    let archive = editor.package().archive(&archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers shape {drawable_object_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary shape"
        )));
    };
    let shape = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
    if shape.is_text_box != Some(false) {
        return Err(Error::ParseError(format!(
            "Numbers drawable {drawable_object_id} is not an ordinary shape"
        )));
    }
    if shape
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(sheet_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} does not name sheet {sheet_id} as its parent"
        )));
    }
    let caption = caption_slot_from_reference(
        editor.package(),
        drawable_object_id,
        shape.super_.super_.caption,
        DrawableCaptionKind::Caption,
    )?;
    let title = caption_slot_from_reference(
        editor.package(),
        drawable_object_id,
        shape.super_.super_.title,
        DrawableCaptionKind::Title,
    )?;
    let storage_id =
        required_reference(drawable_object_id, shape.owned_storage.as_ref(), "storage")?;
    if shape
        .deprecated_storage
        .as_ref()
        .map(|reference| reference.identifier)
        != Some(storage_id)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has inconsistent storage ownership"
        )));
    }
    let mut object_ids = vec![drawable_object_id];
    object_ids.extend(caption.object_ids);
    object_ids.extend(title.object_ids);
    object_ids.push(storage_id);
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} aliases private objects"
        )));
    }
    if locations.get(&storage_id).map(String::as_str) != Some(archive_name.as_str()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape storage {storage_id} is outside {archive_name}"
        )));
    }
    let storage = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers shape storage {storage_id} is missing"))
    })?;
    if storage
        .messages
        .iter()
        .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape storage {storage_id} must have exactly one expected payload"
        )));
    }
    validate_shape_ownership(editor.package(), drawable_object_id, storage_id)?;
    let component_id = component_identifier_for_entry(editor.package(), &archive_name)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet component {archive_name} is not registered"
            ))
        })?;
    let registered =
        component_uuid_identifiers(editor.package(), component_id)?.unwrap_or_default();
    let uuid_object_ids = object_ids
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect::<Vec<_>>();
    Ok(SheetShapeGraph {
        sheet_id,
        archive_name,
        component_id,
        info: shape_info(editor, &locations, sheet_id, drawable_object_id, &shape)?,
        object_ids,
        uuid_object_ids,
    })
}

fn required_reference(
    drawable_object_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers shape {drawable_object_id} has no {label} object"
            ))
        })
}

fn validate_shape_ownership(
    package: &IWorkPackage,
    drawable_object_id: u64,
    storage_id: u64,
) -> Result<()> {
    let document = numbers_document(package)?;
    let locations = object_locations(package)?;
    let mut drawable_owners = 0usize;
    let mut storage_owners = 0usize;
    for sheet in document.sheets {
        let archive_name = locations.get(&sheet.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", sheet.identifier))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(sheet.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", sheet.identifier))
        })?;
        let (_, native) = decode_sheet(object)?;
        for drawable in native.drawable_infos {
            if drawable.identifier == drawable_object_id {
                drawable_owners += 1;
            }
            let Some(drawable_archive_name) = locations.get(&drawable.identifier) else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers drawable {} is missing",
                    drawable.identifier
                )));
            };
            let drawable_archive = package.archive(drawable_archive_name)?;
            let drawable_object =
                drawable_archive
                    .object(drawable.identifier)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers drawable {} is missing",
                            drawable.identifier
                        ))
                    })?;
            for message in drawable_object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            {
                let candidate = tswp::ShapeInfoArchive::decode(message.data.as_slice())?;
                if candidate
                    .owned_storage
                    .as_ref()
                    .map(|reference| reference.identifier)
                    == Some(storage_id)
                {
                    storage_owners += 1;
                }
            }
        }
    }
    if drawable_owners != 1 || storage_owners != 1 {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has {drawable_owners} sheet owners and storage {storage_id} has {storage_owners} drawable owners"
        )));
    }
    Ok(())
}

fn shape_info(
    editor: &NumbersEditor,
    locations: &HashMap<u64, String>,
    sheet_id: u64,
    drawable_object_id: u64,
    shape: &tswp::ShapeInfoArchive,
) -> Result<NumbersSheetShapeInfo> {
    let storage_id = shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers shape {drawable_object_id} has no writable storage"
            ))
        })?;
    let archive_name = locations.get(&drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers shape {drawable_object_id} is missing"))
    })?;
    if locations.get(&storage_id) != Some(archive_name) {
        return Err(Error::InvalidFormat(format!(
            "Numbers shape {drawable_object_id} has storage outside its component"
        )));
    }
    let text = IWorkTextEditor::from_package(editor.package.clone());
    let line_segment = shape_line_segment(shape)?;
    let line_endpoints = line_segment
        .map(|_| shape_line_endpoints(editor.package(), archive_name, drawable_object_id))
        .transpose()?;
    Ok(NumbersSheetShapeInfo {
        sheet_id,
        drawable_object_id,
        kind: shape_path_kind(shape)?,
        preset: shape_preset(shape)?,
        line_segment,
        line_endpoints,
        storage: text.storage(storage_id)?,
        geometry: shape_geometry(editor.package(), archive_name, drawable_object_id)?,
        properties: shape_properties(editor.package(), archive_name, drawable_object_id)?,
    })
}

#[cfg(test)]
mod tests {
    use std::{fs, path::PathBuf};

    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{
        Appearance, BlurRadius, Contact, Endpoint, Offset, Pattern, RgbColorSpace, RgbaColor,
        Width,
    };
    use crate::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
    use litchi_iwa_common::shape::effects::{Effects, Opacity as EffectsOpacity, Reflection,
        ReflectionOpacity};
    use litchi_iwa_common::shape::fill::{Angle, Gradient};
    use litchi_iwa_common::shape::shadow::{Opacity as ShadowOpacity, Perspective};
    use litchi_iwa_common::shape::path::CornerRadius;

    const POSITION: DrawablePoint = DrawablePoint { x: 420.0, y: 300.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 300.0,
        height: 150.0,
    };
    const LINE_START: DrawablePoint = DrawablePoint { x: 420.0, y: 300.0 };
    const LINE_END: DrawablePoint = DrawablePoint { x: 720.0, y: 450.0 };
    const UPDATED_LINE_START: DrawablePoint = DrawablePoint { x: 72.0, y: 180.0 };
    const UPDATED_LINE_END: DrawablePoint = DrawablePoint { x: 432.0, y: 180.0 };

    fn fixture(relative: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        fs::read(root.join(relative)).unwrap()
    }

    #[test]
    fn scratch_spreadsheet_supports_rectangle_crud_without_a_source_drawable() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Scratch Shape")
            .table_name("Scratch Table")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        assert!(editor.sheet_shapes(sheet_id).unwrap().is_empty());
        let baseline = editor.to_bytes().unwrap();

        let created = editor
            .add_sheet_rectangle(sheet_id, "Built from typed objects", POSITION, SIZE)
            .unwrap();
        assert_eq!(created.kind, Kind::Rectangle);
        assert_eq!(created.preset, Some(Preset::Rectangle));
        assert_eq!(created.storage.text, "Built from typed objects");
        let horizontally_flipped = editor
            .flip_sheet_shape(
                sheet_id,
                created.drawable_object_id,
                DrawableFlipAxis::Horizontal,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_shape_geometry(sheet_id, created.drawable_object_id)
                .unwrap(),
            horizontally_flipped
        );
        assert_ne!(horizontally_flipped.flags, created.geometry.flags);
        let vertically_flipped = editor
            .flip_sheet_shape(
                sheet_id,
                created.drawable_object_id,
                DrawableFlipAxis::Vertical,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_shape_geometry(sheet_id, created.drawable_object_id)
                .unwrap(),
            vertically_flipped
        );
        assert_ne!(vertically_flipped.angle, created.geometry.angle);
        editor
            .replace_sheet_shape_text(sheet_id, created.drawable_object_id, 0..5, "Made")
            .unwrap();
        let geometry = DrawableGeometry {
            position: Some(DrawablePoint { x: 72.0, y: 180.0 }),
            size: Some(DrawableSize {
                width: 360.0,
                height: 180.0,
            }),
            flags: Some(DEFAULT_DRAWABLE_FLAGS),
            angle: Some(12.0),
        };
        editor
            .set_sheet_shape_geometry(sheet_id, created.drawable_object_id, geometry)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/numbers-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Generated rectangle".to_owned()),
        };
        editor
            .set_sheet_shape_properties(sheet_id, created.drawable_object_id, properties.clone())
            .unwrap();

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let shape = &reopened.sheet_shapes(sheet_id).unwrap()[0];
        assert_eq!(shape.storage.text, "Made from typed objects");
        assert_eq!(shape.geometry, geometry);
        assert_eq!(shape.properties, properties);

        let removed = editor
            .remove_sheet_shape(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.drawable_object_id, created.drawable_object_id);
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_spreadsheet_supports_native_shape_duplication() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Shape clone")
            .table_name("Source data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_shape(sheet_id, "Source shape", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.35, 0.7, 0.25, 1.0, RgbColorSpace::Srgb).unwrap());
        editor
            .set_sheet_shape_fill(sheet_id, created.drawable_object_id, &fill)
            .unwrap();
        let properties = DrawableProperties {
            hyperlink_url: Some("https://example.com/numbers-shape".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(false),
            accessibility_description: Some("Source shape".to_owned()),
        };
        editor
            .set_sheet_shape_properties(sheet_id, created.drawable_object_id, properties.clone())
            .unwrap();
        editor
            .flip_sheet_shape(
                sheet_id,
                created.drawable_object_id,
                DrawableFlipAxis::Vertical,
            )
            .unwrap();
        let source = editor
            .sheet_shapes(sheet_id)
            .unwrap()
            .into_iter()
            .next()
            .unwrap();

        let duplicate = editor
            .duplicate_sheet_shape(sheet_id, source.drawable_object_id)
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
                .sheet_shape_fill(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            fill
        );

        editor
            .set_sheet_shape_text(sheet_id, duplicate.drawable_object_id, "Independent copy")
            .unwrap();
        assert_eq!(
            editor
                .sheet_shapes(sheet_id)
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == source.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Source shape"
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shapes(sheet_id)
                .unwrap()
                .into_iter()
                .find(|shape| shape.drawable_object_id == duplicate.drawable_object_id)
                .unwrap()
                .storage
                .text,
            "Independent copy"
        );
        assert_eq!(reopened.sheet_shapes(sheet_id).unwrap().len(), 2);

        let removed = editor
            .remove_sheet_shape(sheet_id, duplicate.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.storage.text, "Independent copy");
        assert_eq!(editor.sheet_shapes(sheet_id).unwrap().len(), 1);
    }

    #[test]
    fn scratch_spreadsheet_supports_straight_line_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Scratch Line")
            .table_name("Source Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        let created = editor
            .add_sheet_line(sheet_id, LINE_START, LINE_END)
            .unwrap();
        assert_eq!(created.kind, Kind::Line);
        assert_eq!(created.preset, None);
        assert_eq!(created.storage.text, "");
        assert!(line_segments_match(
            created.line_segment.unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(line_segments_match(
            reopened
                .sheet_line_segment(sheet_id, created.drawable_object_id)
                .unwrap(),
            LineSegment::new(LINE_START, LINE_END).unwrap()
        ));

        editor
            .set_sheet_line_segment(
                sheet_id,
                created.drawable_object_id,
                UPDATED_LINE_START,
                UPDATED_LINE_END,
            )
            .unwrap();
        assert!(line_segments_match(
            editor
                .sheet_line_segment(sheet_id, created.drawable_object_id)
                .unwrap(),
            LineSegment::new(UPDATED_LINE_START, UPDATED_LINE_END).unwrap()
        ));

        let before_invalid = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_line_segment(
                    sheet_id,
                    created.drawable_object_id,
                    UPDATED_LINE_START,
                    UPDATED_LINE_START,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_invalid);

        let removed = editor
            .remove_sheet_shape(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(removed.shape.kind, Kind::Line);
        assert_eq!(editor.to_bytes().unwrap(), baseline);

        let rectangle = editor
            .add_sheet_rectangle(sheet_id, "Not a line", POSITION, SIZE)
            .unwrap();
        let before_cross_type = editor.to_bytes().unwrap();
        assert!(
            editor
                .flip_sheet_shape(sheet_id, u64::MAX, DrawableFlipAxis::Horizontal)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
        assert!(
            editor
                .set_sheet_line_segment(
                    sheet_id,
                    rectangle.drawable_object_id,
                    LINE_START,
                    LINE_END,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
    }

    #[test]
    fn scratch_spreadsheet_supports_typed_shape_stroke_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Stroke")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let stroke = Stroke::new(
            RgbaColor::new(0.15, 0.55, 0.85, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(4.0).unwrap(),
            Pattern::RoundedDash,
        );
        let endpoints = Endpoints::new(Endpoint::FilledCircle, Endpoint::OpenArrow);
        let created = editor
            .add_sheet_line_with_style(
                sheet_id,
                LINE_START,
                LINE_END,
                LineStyle::new(stroke).with_endpoints(endpoints),
            )
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_stroke(sheet_id, created.drawable_object_id)
                .unwrap(),
            Some(stroke)
        );
        assert_eq!(
            reopened
                .sheet_line_endpoints(sheet_id, created.drawable_object_id)
                .unwrap(),
            endpoints
        );
        assert!(
            reopened
                .reset_sheet_shape_stroke(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_stroke(sheet_id, created.drawable_object_id)
                .unwrap(),
            None
        );
        assert_eq!(
            reopened
                .sheet_line_endpoints(sheet_id, created.drawable_object_id)
                .unwrap(),
            endpoints
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_typed_shape_fill_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Fill")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let fill = ShapeFill::Gradient(Gradient::linear(
            RgbaColor::new(0.1, 0.55, 0.85, 1.0, RgbColorSpace::Srgb).unwrap(),
            RgbaColor::new(0.05, 0.15, 0.5, 1.0, RgbColorSpace::Srgb).unwrap(),
            Angle::from_degrees(0.0).unwrap(),
        ));
        let created = editor
            .add_sheet_shape(sheet_id, "Filled", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited_fill = editor
            .sheet_shape_fill(sheet_id, created.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_shape_fill(sheet_id, created.drawable_object_id, &fill)
            .unwrap();
        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap(),
            fill
        );
        assert!(
            reopened
                .reset_sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap(),
            inherited_fill
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_embedded_shape_image_fill_crud() {
        let bytes = fixture("test-data/images/png/lena.png");
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Image Fill")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_shape(sheet_id, "Image", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .sheet_shape_fill(sheet_id, created.drawable_object_id)
            .unwrap();
        let image = editor
            .set_sheet_shape_image_fill(
                sheet_id,
                created.drawable_object_id,
                "lena.png",
                &bytes,
                ShapeImageFillTechnique::ScaleToFit,
                None,
            )
            .unwrap();
        assert_eq!(
            editor
                .extract_media(image.data_identifier().unwrap().get())
                .unwrap(),
            bytes
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap(),
            ShapeFill::Image(image)
        );
        assert!(
            reopened
                .reset_sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_fill(sheet_id, created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(reopened.media_assets().unwrap().is_empty());
    }

    #[test]
    fn scratch_spreadsheet_supports_shape_effect_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Effects")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_shape(sheet_id, "Effects", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .sheet_shape_effects(sheet_id, created.drawable_object_id)
            .unwrap();
        let effects = Effects::new(
            EffectsOpacity::new(0.84).unwrap(),
            Reflection::Enabled(ReflectionOpacity::new(0.65).unwrap()),
        );
        editor
            .set_sheet_shape_effects(sheet_id, created.drawable_object_id, effects)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_effects(sheet_id, created.drawable_object_id)
                .unwrap(),
            effects
        );
        assert!(
            reopened
                .reset_sheet_shape_effects(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_effects(sheet_id, created.drawable_object_id)
                .unwrap(),
            inherited
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_contact_shadow_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Shadows")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_shape(sheet_id, "Shadow", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .sheet_shape_shadow(sheet_id, created.drawable_object_id)
            .unwrap();
        let shadow = Shadow::Contact(Contact::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(18).unwrap(),
                Offset::from_points(6.0).unwrap(),
                ShadowOpacity::new(0.58).unwrap(),
            ),
            Perspective::from_degrees(23.0).unwrap(),
        ));
        editor
            .set_sheet_shape_shadow(sheet_id, created.drawable_object_id, shadow)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_shadow(sheet_id, created.drawable_object_id)
                .unwrap(),
            shadow
        );
        assert!(
            reopened
                .reset_sheet_shape_shadow(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_shadow(sheet_id, created.drawable_object_id)
                .unwrap(),
            inherited
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_shape_text_layout_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Layout")
            .table_name("Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_shape(sheet_id, "Layout", POSITION, SIZE, Preset::Rectangle)
            .unwrap();
        let inherited = editor
            .sheet_shape_text_layout(sheet_id, created.drawable_object_id)
            .unwrap();
        let layout = Layout::new(
            VerticalAlignment::Bottom,
            Insets::uniform(Inset::from_points(9.0).unwrap()),
            AutoSize::Fixed,
        );
        editor
            .set_sheet_shape_text_layout(sheet_id, created.drawable_object_id, layout)
            .unwrap();

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_text_layout(sheet_id, created.drawable_object_id)
                .unwrap(),
            layout
        );
        assert!(
            reopened
                .reset_sheet_shape_text_layout(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_shape_text_layout(sheet_id, created.drawable_object_id)
                .unwrap(),
            inherited
        );
        assert!(
            !reopened
                .reset_sheet_shape_text_layout(sheet_id, created.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_typed_line_endpoint_crud() {
        use crate::shapes::{Endpoint, Endpoints};

        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Endpoint Styles")
            .table_name("Source Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_line_with_endpoints(
                sheet_id,
                LINE_START,
                LINE_END,
                Endpoints::new(Endpoint::FilledSquare, Endpoint::OpenArrow),
            )
            .unwrap();
        assert_eq!(
            created.line_endpoints,
            Some(Endpoints::new(Endpoint::FilledSquare, Endpoint::OpenArrow))
        );

        let mut reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let replacement = Endpoints::new(Endpoint::InvertedArrow, Endpoint::FilledCircle);
        reopened
            .set_sheet_line_endpoints(sheet_id, created.drawable_object_id, replacement)
            .unwrap();
        assert_eq!(
            reopened
                .sheet_line_endpoints(sheet_id, created.drawable_object_id)
                .unwrap(),
            replacement
        );
        assert!(
            reopened
                .reset_sheet_line_endpoints(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_line_endpoints(sheet_id, created.drawable_object_id)
                .unwrap(),
            Endpoints::default()
        );
    }

    #[test]
    fn scratch_spreadsheet_supports_typed_preset_shape_crud() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Preset Shapes")
            .table_name("Source Data")
            .build()
            .unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        let created = editor
            .add_sheet_shape(
                sheet_id,
                "Rounded",
                POSITION,
                SIZE,
                Preset::ROUNDED_RECTANGLE,
            )
            .unwrap();
        assert_eq!(created.kind, Kind::RoundedRectangle);
        assert_eq!(created.preset, Some(Preset::ROUNDED_RECTANGLE));

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_shape_preset(sheet_id, created.drawable_object_id)
                .unwrap(),
            Some(Preset::ROUNDED_RECTANGLE)
        );

        for (preset, kind) in [
            (Preset::Ellipse, Kind::Ellipse),
            (Preset::LeftArrow, Kind::LeftArrow),
            (Preset::RightArrow, Kind::RightArrow),
            (Preset::DoubleArrow, Kind::DoubleArrow),
            (Preset::PENTAGON, Kind::RegularPolygon),
            (Preset::STAR, Kind::Star),
        ] {
            editor
                .set_sheet_shape_preset(sheet_id, created.drawable_object_id, preset)
                .unwrap();
            let shape = &editor.sheet_shapes(sheet_id).unwrap()[0];
            assert_eq!(shape.kind, kind);
            assert_eq!(shape.preset, Some(preset));
            assert_eq!(shape.storage.text, "Rounded");
        }

        editor
            .remove_sheet_shape(sheet_id, created.drawable_object_id)
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn invalid_rectangle_creation_and_cross_type_updates_are_transactional() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .add_sheet_rectangle(
                    sheet_id,
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
                .add_sheet_shape(
                    sheet_id,
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
            .add_sheet_text_box(sheet_id, "not a shape", POSITION, SIZE)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_sheet_shape_text(sheet_id, text_box.drawable_object_id, "wrong type")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
        assert!(
            editor
                .set_sheet_shape_preset(sheet_id, text_box.drawable_object_id, Preset::Ellipse,)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}

//! Ordinary text-bearing shape CRUD for Keynote slides.

use std::collections::HashSet;
use std::ops::Range;

use super::*;
use crate::shapes::{
    DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize, LineSegment, ShapePathKind,
    ShapePreset, line_geometry, line_path_source, line_segments_match, set_shape_geometry,
    set_shape_line_segment, set_shape_preset, shape_line_segment, shape_path_kind,
    shape_path_source, shape_preset,
};
use crate::text::TextStorageInfo;

use super::text_box_create::{
    TextBoxObjectIds, slide_text_storage_template, text_box_context, text_box_objects,
    text_box_storage, text_box_theme_styles,
};

const SHAPE_MESSAGE_TYPE: u32 = 2_011;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;
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
        let storage = text_box_storage(text, base_storage.as_ref(), &styles);
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

    /// Remove an ordinary shape and its private text-bearing object graph.
    pub fn remove_slide_shape(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<RemovedKeynoteSlideShape> {
        let source = shape_graph(self, slide_index, drawable_object_id)?;
        let mut comments = IWorkDrawableCommentEditor::from_package(self.package().clone())?;
        comments.clear_comment(drawable_object_id)?;
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
    let required = required_shape_objects(drawable_object_id, &shape)?;
    let archive = editor.package().archive(&archive_name)?;
    validate_shape_private_objects(&archive, drawable_object_id, &shape)?;
    let object_ids = slide_create::graph::private_clone_object_ids(
        &archive,
        [drawable_object_id],
        "slide shape",
    )?;
    let actual = object_ids.iter().copied().collect::<HashSet<_>>();
    if actual != required || object_ids.len() != required.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} does not have an isolated four-object graph"
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
    if !registered.is_empty() && uuid_object_ids.len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide UUID map does not cover shape {drawable_object_id}"
        )));
    }
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
    Ok(KeynoteSlideShapeInfo {
        slide_index,
        drawable_object_id,
        kind: shape_path_kind(shape)?,
        preset: shape_preset(shape)?,
        line_segment: shape_line_segment(shape)?,
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
    let identifiers = [
        drawable_object_id,
        required_reference(shape.super_.super_.caption.as_ref(), "caption stand-in")?,
        required_reference(shape.super_.super_.title.as_ref(), "title stand-in")?,
        required_reference(shape.owned_storage.as_ref(), "storage")?,
    ];
    let required = identifiers.into_iter().collect::<HashSet<_>>();
    if required.len() != identifiers.len() {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} aliases private objects"
        )));
    }
    Ok(required)
}

fn validate_shape_private_objects(
    archive: &Archive,
    drawable_object_id: u64,
    shape: &tswp::ShapeInfoArchive,
) -> Result<()> {
    for (reference, label) in [
        (shape.super_.super_.caption.as_ref(), "caption"),
        (shape.super_.super_.title.as_ref(), "title"),
    ] {
        let identifier = reference
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote shape {drawable_object_id} has no {label} stand-in"
                ))
            })?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} {label} stand-in {identifier} is missing"
            ))
        })?;
        if object
            .messages
            .iter()
            .filter(|message| message.type_ == STANDIN_CAPTION_MESSAGE_TYPE)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} {label} object {identifier} must have exactly one stand-in payload"
            )));
        }
    }
    let storage_id = shape
        .owned_storage
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote shape {drawable_object_id} has no writable storage"
            ))
        })?;
    let storage = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} storage {storage_id} is missing"
        ))
    })?;
    if storage
        .messages
        .iter()
        .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote shape {drawable_object_id} storage {storage_id} must have exactly one writable payload"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::shapes::{
        ShapeCornerRadius, ShapePolygonSides, ShapeStarInnerRatio, ShapeStarPoints,
    };

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
                .set_slide_line_segment(0, rectangle.drawable_object_id, LINE_START, LINE_END,)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before_cross_type);
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

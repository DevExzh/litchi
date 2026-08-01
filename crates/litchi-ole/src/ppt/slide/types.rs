/// High-performance Slide implementation with lazy shape loading and zero-copy design.
use super::super::package::{PptError, Result};
use super::super::records::PptRecord;
use super::super::shapes::ShapeEnum;
use super::factory::SlideData;
use super::notes::{NoteDescriptor, SpeakerNotes};
use crate::consts::PptRecordType;
use crate::ppt::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::ppt::odraw::{FrameKind, ShapeExt as _};
use crate::ppt::slide_extension::PowerPoint12SlideExtension;
use crate::ppt::slide_round_trip::PowerPoint12SlideRoundTripMetadata;
use crate::ppt::slide_sync::PowerPointSlideSyncInfo;
use crate::ppt::transition::{TransitionInfo, parse_transition};
use once_cell::unsync::OnceCell;

/// A slide in a PowerPoint presentation with lazy-loaded shapes.
///
/// # Performance
///
/// - Shapes are parsed only when first accessed
/// - Uses `OnceCell` for one-time initialization
/// - Zero-copy text extraction where possible
pub struct Slide<'doc> {
    /// Slide persist ID
    persist_id: u32,
    /// Stable SlideId from the live SlidePersistAtom.
    slide_id: u32,
    slide_list_text: String,
    outline_text_interactions: Vec<crate::ppt::PowerPointTextBodyInteractions>,
    outline_text_refs: Vec<crate::ppt::PowerPointOutlineTextRef>,
    /// Slide number (1-based for display)
    slide_number: usize,
    /// Slide record
    record: PptRecord,
    /// Reference to document data for lazy shape parsing (reserved for future use)
    #[allow(dead_code)]
    doc_data: &'doc [u8],
    /// Lazily-loaded shapes (use 'static since they store owned data)
    shapes: OnceCell<Vec<ShapeEnum<'static>>>,
    /// Cached text content
    text_cache: OnceCell<String>,
    /// Lazily parsed, inert shape animation metadata.
    animations: OnceCell<Vec<ShapeAnimation>>,
    /// Lazily parsed PowerPoint 2002 slide animation extension.
    animation_extension: OnceCell<Option<SlideAnimationExtension>>,
    /// Lazily parsed PowerPoint 12 slide/master round-trip metadata.
    powerpoint12_extension: OnceCell<PowerPoint12SlideExtension>,
    /// Lazily parsed, inert slide-library synchronization metadata.
    sync_info: OnceCell<Option<PowerPointSlideSyncInfo>>,
    /// Lazily parsed direct PowerPoint 12 slide round-trip metadata.
    round_trip_metadata: OnceCell<PowerPoint12SlideRoundTripMetadata>,
    notes_descriptor: std::result::Result<Option<NoteDescriptor>, String>,
    speaker_notes: OnceCell<Option<SpeakerNotes>>,
}

impl<'doc> Slide<'doc> {
    /// Create a slide from parsed slide data.
    pub fn from_slide_data(data: SlideData<'doc>, slide_number: usize) -> Self {
        let doc_data_ref = data.doc_data();
        Self {
            persist_id: data.persist_id,
            slide_id: data.slide_id,
            slide_list_text: data.slide_list_text,
            outline_text_interactions: data.outline_text_interactions,
            outline_text_refs: data.outline_text_refs,
            slide_number,
            doc_data: doc_data_ref,
            record: data.record,
            shapes: OnceCell::new(),
            text_cache: OnceCell::new(),
            animations: OnceCell::new(),
            animation_extension: OnceCell::new(),
            powerpoint12_extension: OnceCell::new(),
            sync_info: OnceCell::new(),
            round_trip_metadata: OnceCell::new(),
            notes_descriptor: data.note_descriptor,
            speaker_notes: OnceCell::new(),
        }
    }

    /// Get the slide number (1-based).
    #[inline]
    pub fn slide_number(&self) -> usize {
        self.slide_number
    }

    /// Get the persist ID.
    #[inline]
    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    /// Get the stable presentation SlideId.
    #[inline]
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Get shapes on this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Shapes are parsed only on first call
    /// - Subsequent calls return cached reference
    /// - Zero allocation after first parse
    pub fn shapes(&self) -> Result<&[ShapeEnum<'static>]> {
        self.shapes
            .get_or_try_init(|| self.parse_shapes())
            .map(|v| v.as_slice())
    }

    /// Get the number of shapes (triggers parsing if not yet loaded).
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Return every shape that has a click or mouse-over interaction.
    pub fn shape_interactions(&self) -> Result<Vec<crate::ppt::PowerPointShapeInteractionEntry>> {
        self.shape_interactions_with_limits(crate::ppt::PowerPointInteractionLimits::default())
    }

    /// Return shape interactions with caller-supplied record and name limits.
    pub fn shape_interactions_with_limits(
        &self,
        limits: crate::ppt::PowerPointInteractionLimits,
    ) -> Result<Vec<crate::ppt::PowerPointShapeInteractionEntry>> {
        let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut result = Vec::new();
        let mut pending = escher_shapes.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            let interactions = shape.interactions_with_limits(limits)?;
            if !interactions.is_empty() {
                result.push(crate::ppt::PowerPointShapeInteractionEntry {
                    shape_id: shape.id(),
                    interactions,
                });
            }
            pending.extend(shape.children().iter().rev());
        }
        Ok(result)
    }

    /// Return every shape that has a range-anchored text action.
    pub fn shape_text_interactions(
        &self,
    ) -> Result<Vec<crate::ppt::PowerPointShapeTextInteractionEntry>> {
        self.shape_text_interactions_with_limits(
            crate::ppt::PowerPointTextInteractionLimits::default(),
        )
    }

    /// Return shape text actions with caller-supplied resource limits.
    pub fn shape_text_interactions_with_limits(
        &self,
        limits: crate::ppt::PowerPointTextInteractionLimits,
    ) -> Result<Vec<crate::ppt::PowerPointShapeTextInteractionEntry>> {
        let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut result = Vec::new();
        let mut pending = escher_shapes.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            let interactions = shape.text_interactions_with_limits(limits)?;
            if !interactions.is_empty() {
                result.push(crate::ppt::PowerPointShapeTextInteractionEntry {
                    shape_id: shape.id(),
                    interactions,
                });
            }
            pending.extend(shape.children().iter().rev());
        }
        Ok(result)
    }

    /// Range-anchored actions stored with outline/placeholder text.
    pub fn outline_text_interactions(&self) -> &[crate::ppt::PowerPointTextBodyInteractions] {
        &self.outline_text_interactions
    }

    /// Validated outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78)
    /// tying this slide's shapes to outline text bodies.
    pub fn outline_text_refs(&self) -> &[crate::ppt::PowerPointOutlineTextRef] {
        &self.outline_text_refs
    }

    /// Return every shape-scoped programmable-tag container on this slide.
    pub fn shape_programmable_tags(
        &self,
    ) -> Result<Vec<crate::ppt::PowerPointShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(
            crate::ppt::PowerPointShapeProgrammableTagLimits::default(),
        )
    }

    /// Return shape programmable tags with caller-supplied resource limits.
    pub fn shape_programmable_tags_with_limits(
        &self,
        limits: crate::ppt::PowerPointShapeProgrammableTagLimits,
    ) -> Result<Vec<crate::ppt::PowerPointShapeProgrammableTagsEntry>> {
        let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(programmable_tags) = shape.programmable_tags_with_limits(limits)? {
                result.push(crate::ppt::PowerPointShapeProgrammableTagsEntry {
                    shape_id: shape.id(),
                    programmable_tags,
                });
            }
        }
        Ok(result)
    }

    /// Return this slide's typed slide-level programmable tags (MS-PPT 2.5.19),
    /// when the slide carries a `SlideProgTagsContainer`.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ppt::PowerPointProgTags::slide_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    pub fn programmable_tags(&self) -> Result<Option<crate::ppt::PowerPointProgTags>> {
        self.programmable_tags_with_limits(crate::ppt::PowerPointProgTagLimits::default())
    }

    /// Return slide-level programmable tags with caller-supplied resource limits.
    pub fn programmable_tags_with_limits(
        &self,
        limits: crate::ppt::PowerPointProgTagLimits,
    ) -> Result<Option<crate::ppt::PowerPointProgTags>> {
        crate::ppt::PowerPointProgTags::parse_slide(&self.record, limits)
    }

    /// Return every typed shape-flag projection on this slide.
    pub fn shape_flags(&self) -> Result<Vec<crate::ppt::PowerPointShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ppt::PowerPointShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    pub fn shape_flags_with_limits(
        &self,
        limits: crate::ppt::PowerPointShapeFlagLimits,
    ) -> Result<Vec<crate::ppt::PowerPointShapeFlagEntry>> {
        let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(projection) = shape.ppt_flags_with(limits)? {
                result.push(crate::ppt::PowerPointShapeFlagEntry {
                    shape_id: shape.id(),
                    projection,
                });
            }
        }
        Ok(result)
    }

    /// Return context-validated placeholders on this presentation slide.
    pub fn placeholder_atoms(&self) -> Result<Vec<crate::ppt::PowerPointPlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::ppt::PowerPointPlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    pub fn placeholder_atoms_with_limits(
        &self,
        limits: crate::ppt::PowerPointPlaceholderLimits,
    ) -> Result<Vec<crate::ppt::PowerPointPlaceholderEntry>> {
        let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        else {
            return Ok(Vec::new());
        };
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut positions = std::collections::HashSet::new();
        let mut result = Vec::new();
        for shape in &escher_shapes {
            if let Some(placeholder) = shape.placeholder_atom_with_limits(
                crate::ppt::PowerPointPlaceholderContext::PresentationSlide,
                limits,
            )? {
                if placeholder.position != -1 && !positions.insert(placeholder.position) {
                    return Err(PptError::Corrupted(
                        "Presentation slide contains duplicate placeholder positions".to_string(),
                    ));
                }
                result.push(crate::ppt::PowerPointPlaceholderEntry {
                    shape_id: shape.id(),
                    placeholder,
                });
            }
        }
        Ok(result)
    }

    /// Return this slide's speaker-notes page, if one exists.
    pub fn speaker_notes(&self) -> Result<Option<&SpeakerNotes>> {
        self.speaker_notes
            .get_or_try_init(|| match &self.notes_descriptor {
                Ok(None) => Ok(None),
                Ok(Some(descriptor)) => SpeakerNotes::parse(*descriptor, self.doc_data).map(Some),
                Err(error) => Err(PptError::Corrupted(error.clone())),
            })
            .map(Option::as_ref)
    }

    /// Return inert PowerPoint 97 animation metadata keyed by shape ID.
    pub fn animations(&self) -> Result<&[ShapeAnimation]> {
        self.animations
            .get_or_try_init(|| self.parse_animations())
            .map(Vec::as_slice)
    }

    fn parse_animations(&self) -> Result<Vec<ShapeAnimation>> {
        let Some(ppdrawing) = self.record.find_child(PptRecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;
        let mut animations = Vec::new();
        let mut pending = shapes.iter().rev().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            if let Some(animation) = shape.animation()? {
                animations.push(ShapeAnimation {
                    shape_id: shape.id(),
                    animation,
                });
            }
            pending.extend(shape.children().iter().rev());
        }
        Ok(animations)
    }

    /// Return inert PowerPoint 2002 timing and build metadata from `___PPT10`.
    pub fn animation_extension(&self) -> Result<Option<&SlideAnimationExtension>> {
        self.animation_extension
            .get_or_try_init(|| self.parse_animation_extension())
            .map(Option::as_ref)
    }

    fn parse_animation_extension(&self) -> Result<Option<SlideAnimationExtension>> {
        for prog_tags in self.record.find_children(PptRecordType::ProgTags) {
            for prog_binary_tag in prog_tags.find_children(PptRecordType::ProgBinaryTag) {
                let Some(tag_name) = prog_binary_tag.find_child(PptRecordType::CString) else {
                    continue;
                };
                if !Self::is_ppt10_tag_name(tag_name) {
                    continue;
                }
                let data = prog_binary_tag
                    .find_child(PptRecordType::BinaryTagData)
                    .ok_or_else(|| {
                        PptError::Corrupted(
                            "___PPT10 programmable tag is missing BinaryTagData".to_string(),
                        )
                    })?;
                return crate::ppt::animation::parse_slide_animation_extension(&data.data)
                    .map(Some);
            }
        }
        Ok(None)
    }

    fn is_ppt10_tag_name(record: &PptRecord) -> bool {
        const PPT10: [u16; 8] = [0x5F, 0x5F, 0x5F, 0x50, 0x50, 0x54, 0x31, 0x30];
        record.version == 0
            && record.instance == 0
            && record.data.len() == 16
            && record
                .data
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .eq(PPT10)
    }

    /// Return PowerPoint 12 slide/master round-trip metadata from `___PPT12`.
    pub fn powerpoint12_extension(&self) -> Result<&PowerPoint12SlideExtension> {
        self.powerpoint12_extension
            .get_or_try_init(|| PowerPoint12SlideExtension::parse(&self.record))
    }

    /// Return inert PowerPoint 12 slide-library synchronization metadata.
    pub fn sync_info(&self) -> Result<Option<&PowerPointSlideSyncInfo>> {
        self.sync_info
            .get_or_try_init(|| PowerPointSlideSyncInfo::parse(&self.record))
            .map(Option::as_ref)
    }

    /// Return inert PowerPoint 12 metadata stored directly on this slide.
    pub fn powerpoint12_round_trip_metadata(&self) -> Result<&PowerPoint12SlideRoundTripMetadata> {
        self.round_trip_metadata
            .get_or_try_init(|| PowerPoint12SlideRoundTripMetadata::parse(&self.record))
    }

    /// Extract all text from this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Text is extracted and cached on first call
    /// - Includes text from:
    ///   * Direct text records in the slide
    ///   * Shapes (via PPDrawing/Escher)
    pub fn text(&self) -> Result<&str> {
        self.text_cache
            .get_or_try_init(|| {
                let text = self.extract_all_text()?;
                if text.is_empty() {
                    Ok(self.slide_list_text.clone())
                } else {
                    Ok(text)
                }
            })
            .map(|s| s.as_str())
    }

    /// Parse shapes from PPDrawing record.
    ///
    /// # Performance
    ///
    /// - Shapes store owned data to allow caching
    /// - Lazy: Only called when shapes() is accessed
    /// - Uses Escher parser for efficient traversal
    ///
    /// Note: Shapes use 'static lifetime since they're stored in the Slide
    /// and need to outlive the parsing function scope.
    fn parse_shapes(&self) -> Result<Vec<ShapeEnum<'static>>> {
        // Find PPDrawing record
        let ppdrawing = match self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        {
            Some(record) => record,
            None => return Ok(Vec::new()),
        };

        // Extract Escher shapes from PPDrawing data
        let escher_shapes = crate::ppt::odraw::parse(&ppdrawing.data)?;

        // Convert Escher shapes to ShapeEnum with full property extraction
        let mut shapes = Vec::with_capacity(escher_shapes.len());
        for escher_shape in &escher_shapes {
            if let Some(shape) = Self::convert_odraw_to_shape_enum(escher_shape)? {
                shapes.push(shape);
            }
        }

        Ok(shapes)
    }

    /// Convert an EscherShape to ShapeEnum with full property extraction.
    ///
    /// # Performance
    ///
    /// - Direct property access (no allocations)
    /// - Pattern matching for type dispatch
    pub(crate) fn convert_odraw_to_shape_enum(
        odraw_shape: &litchi_odraw::shape::Shape<'_>,
    ) -> Result<Option<ShapeEnum<'static>>> {
        Self::convert_odraw_shape(odraw_shape, 0)
    }

    fn convert_odraw_shape(
        escher_shape: &litchi_odraw::shape::Shape<'_>,
        depth: usize,
    ) -> Result<Option<ShapeEnum<'static>>> {
        const MAX_SHAPE_DEPTH: usize = 256;
        if depth >= MAX_SHAPE_DEPTH {
            return Err(PptError::Corrupted(
                "OfficeArt shape tree exceeds the PPT nesting limit".to_string(),
            ));
        }

        use super::super::shapes::*;
        use crate::ppt::slide_extension::PowerPointHeaderFooterPlaceholder;
        use litchi_odraw::shape::Kind;

        let shape_id = escher_shape.id();
        let anchor = crate::ppt::odraw::anchor(escher_shape)?;
        let powerpoint12_shape_metadata = escher_shape.powerpoint12_shape_metadata()?;

        if let Some(placeholder_info) = escher_shape.placeholder()? {
            let mut properties = shape::ShapeProperties {
                id: shape_id,
                shape_type: shape::ShapeType::Placeholder,
                powerpoint12_shape_metadata,
                ..Default::default()
            };
            if let Some(a) = anchor {
                properties.x = a.left();
                properties.y = a.top();
                properties.width = a.width();
                properties.height = a.height();
            }

            return Ok(Some(ShapeEnum::Placeholder(Placeholder::from_parsed(
                properties,
                PlaceholderType::from(placeholder_info.kind),
                PlaceholderSize::from(placeholder_info.size),
                placeholder_info.position,
                escher_shape.text()?,
            ))));
        }

        if let Some(header_footer) =
            powerpoint12_shape_metadata.and_then(|metadata| metadata.header_footer)
        {
            let placeholder_type = match header_footer {
                PowerPointHeaderFooterPlaceholder::Date => PlaceholderType::DateAndTime,
                PowerPointHeaderFooterPlaceholder::SlideNumber => PlaceholderType::SlideNumber,
                PowerPointHeaderFooterPlaceholder::Footer => PlaceholderType::Footer,
                PowerPointHeaderFooterPlaceholder::Header => PlaceholderType::Header,
            };
            let mut properties = shape::ShapeProperties {
                id: shape_id,
                shape_type: shape::ShapeType::Placeholder,
                powerpoint12_shape_metadata,
                ..Default::default()
            };
            if let Some(a) = anchor {
                properties.x = a.left();
                properties.y = a.top();
                properties.width = a.width();
                properties.height = a.height();
            }
            return Ok(Some(ShapeEnum::Placeholder(Placeholder::from_parsed(
                properties,
                placeholder_type,
                PlaceholderSize::Half,
                None,
                escher_shape.text()?,
            ))));
        }

        match escher_shape.kind() {
            Kind::TextBox => {
                // Create TextBox with proper properties
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: shape::ShapeType::TextBox,
                    powerpoint12_shape_metadata,
                    ..Default::default()
                };

                // Set coordinates if anchor exists
                if let Some(a) = anchor {
                    properties.x = a.left();
                    properties.y = a.top();
                    properties.width = a.width();
                    properties.height = a.height();
                }

                Ok(Some(ShapeEnum::TextBox(TextBox::from_odraw(
                    properties,
                    escher_shape,
                )?)))
            },

            Kind::Picture => {
                // Create PictureShape
                let mut picture = crate::ppt::shapes::PictureShape::new(shape_id);

                picture.set_frame_kind(match escher_shape.frame_kind()? {
                    FrameKind::Object => PictureFrameKind::OleObject,
                    FrameKind::Media => PictureFrameKind::Media,
                    FrameKind::Picture => PictureFrameKind::Picture,
                });

                if let Some(external_object_id) = escher_shape.external_object_id()? {
                    picture.set_external_object_id(external_object_id);
                }

                if let Some(a) = anchor {
                    picture.set_bounds(a.left(), a.top(), a.width(), a.height());
                }
                picture.properties_mut().powerpoint12_shape_metadata = powerpoint12_shape_metadata;

                // Extract the one-based BLIP store index from the pib property.
                use litchi_odraw::prop::Id;
                if let Some(blip_id) = escher_shape
                    .props()
                    .get_int(Id::BlipToDisplay)
                    .and_then(|value| u32::try_from(value).ok())
                    .filter(|value| *value != 0)
                {
                    picture.set_blip_id(blip_id);
                }

                Ok(Some(ShapeEnum::Picture(picture)))
            },

            Kind::Line | Kind::Connector => {
                // Create LineShape
                if let Some(a) = anchor {
                    let mut line = if escher_shape.kind() == Kind::Connector {
                        shape_enum::LineShape::connector(
                            shape_id,
                            a.left(),
                            a.top(),
                            a.right(),
                            a.bottom(),
                        )
                    } else {
                        shape_enum::LineShape::new(
                            shape_id,
                            a.left(),
                            a.top(),
                            a.right(),
                            a.bottom(),
                        )
                    };

                    // Extract line properties
                    use litchi_odraw::prop::Id;
                    if let Some(width) = escher_shape.props().get_int(Id::LineWidth) {
                        line.set_width(width);
                    }
                    if let Some(color) = escher_shape.props().get_color(Id::LineColor) {
                        line.set_color(color.raw());
                    }
                    line.set_powerpoint12_shape_metadata(powerpoint12_shape_metadata);

                    Ok(Some(ShapeEnum::Line(line)))
                } else {
                    Ok(None)
                }
            },

            Kind::Group => {
                // Create GroupShape and parse children recursively
                let mut group = shape_enum::GroupShape::new(shape_id);

                if let Some(a) = anchor {
                    group.set_bounds(a.left(), a.top(), a.width(), a.height());
                }
                group.set_powerpoint12_shape_metadata(powerpoint12_shape_metadata);

                // Recursively parse child shapes
                // This follows Apache POI's approach: iterate child shapes and convert them
                for child_escher in escher_shape.children() {
                    if let Some(child_shape) = Self::convert_odraw_shape(child_escher, depth + 1)? {
                        group.add_child(child_shape);
                    }
                }

                Ok(Some(ShapeEnum::Group(group)))
            },

            Kind::Table => {
                use std::collections::BTreeSet;

                let mut cells = Vec::new();
                for child in escher_shape.children().iter().filter(|child| {
                    matches!(
                        child.kind(),
                        Kind::Rectangle | Kind::TextBox | Kind::AutoShape
                    )
                }) {
                    if let Some(anchor) = crate::ppt::odraw::anchor(child)?
                        && anchor.width() > 0
                        && anchor.height() > 0
                    {
                        cells.push((child, anchor));
                    }
                }

                let columns: BTreeSet<i32> =
                    cells.iter().map(|(_, anchor)| anchor.left()).collect();
                let rows: BTreeSet<i32> = cells.iter().map(|(_, anchor)| anchor.top()).collect();
                let column_positions: Vec<_> = columns.into_iter().collect();
                let row_positions: Vec<_> = rows.into_iter().collect();

                let mut table = shape_enum::TableShape::new(
                    shape_id,
                    row_positions.len(),
                    column_positions.len(),
                );
                if let Some(a) = anchor {
                    table.set_bounds(a.left(), a.top(), a.width(), a.height());
                }
                table.set_powerpoint12_shape_metadata(powerpoint12_shape_metadata);

                for (cell, cell_anchor) in cells {
                    let Ok(row) = row_positions.binary_search(&cell_anchor.top()) else {
                        continue;
                    };
                    let Ok(column) = column_positions.binary_search(&cell_anchor.left()) else {
                        continue;
                    };
                    table.set_cell_text(row, column, cell.text()?.unwrap_or_default());
                }

                Ok(Some(ShapeEnum::Table(table)))
            },

            Kind::Rectangle | Kind::Ellipse | Kind::Callout | Kind::Polygon | Kind::AutoShape => {
                // Create AutoShape
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: shape::ShapeType::AutoShape,
                    powerpoint12_shape_metadata,
                    ..Default::default()
                };

                if let Some(a) = anchor {
                    properties.x = a.left();
                    properties.y = a.top();
                    properties.width = a.width();
                    properties.height = a.height();
                }

                let mut autoshape = AutoShape::from_odraw(
                    properties,
                    escher_shape.native_kind().raw(),
                    escher_shape.props(),
                );
                if let Some(text) = escher_shape.text()?.filter(|text| !text.is_empty()) {
                    autoshape.set_text(text);
                }
                Ok(Some(ShapeEnum::AutoShape(autoshape)))
            },

            // Unknown or unsupported shape types
            _ => Ok(None),
        }
    }

    /// Extract all text from slide and its shapes.
    fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // 1. Extract text from direct slide records (TextCharsAtom, etc.)
        // Note: record.extract_text() already recursively processes all children
        let record_text = self.record.extract_text()?;
        let trimmed = record_text.trim();
        if !trimmed.is_empty() {
            text_parts.push(trimmed.to_string());
        }

        // 2. Extract text from Escher/PPDrawing (shapes, text boxes)
        // This is separate from regular record text extraction
        if let Some(ppdrawing) = self
            .record
            .find_child(crate::consts::PptRecordType::PPDrawing)
        {
            let escher_text = crate::ppt::odraw::text_from_drawing(&ppdrawing.data)?;
            let trimmed = escher_text.trim();
            if !trimmed.is_empty() {
                text_parts.push(trimmed.to_string());
            }
        }

        Ok(if text_parts.is_empty() {
            String::new()
        } else {
            text_parts.join("\n")
        })
    }

    /// Check if this slide has a PPDrawing record (shapes).
    #[inline]
    pub fn has_drawing(&self) -> bool {
        self.record
            .find_child(crate::consts::PptRecordType::PPDrawing)
            .is_some()
    }

    /// Get raw slide record for advanced use cases.
    #[inline]
    pub fn record(&self) -> &PptRecord {
        &self.record
    }

    /// Parse comments from this slide's BinaryTagData.
    ///
    /// Comments are stored inside `ProgTags/ProgBinaryTag/BinaryTagData`
    /// as `Comment2000` (type=12000) containers.
    ///
    /// # Returns
    ///
    /// A vector of parsed comments (author, text, initials, position, date).
    /// Returns an empty vector if no comments are found.
    ///
    /// # Errors
    ///
    /// Returns an error when the PowerPoint 10 programmable tag or a comment record is malformed.
    pub fn comments(&self) -> Result<Vec<ParsedComment>> {
        crate::ppt::comments::parse_slide_comments(&self.record)
    }

    /// Get the slide transition from the `SSSlideInfoAtom` record.
    ///
    /// The transition describes the visual effect (type, direction, speed),
    /// the advance mode (on click, automatic, or both), and an optional
    /// sound played when the slide is shown.
    ///
    /// # Returns
    ///
    /// `Ok(None)` when the slide has no `SSSlideInfoAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error when the `SSSlideInfoAtom` record is truncated.
    pub fn transition(&self) -> Result<Option<TransitionInfo>> {
        match self.record.find_child(PptRecordType::SSSlideInfoAtom) {
            Some(info) => Ok(Some(parse_transition(info)?)),
            None => Ok(None),
        }
    }

    /// Get the slide timing from the SSSlideInfoAtom record.
    ///
    /// Returns `None` if the slide has no timing record.
    pub fn timing(&self) -> Option<ParsedSlideTiming> {
        // SSSlideInfoAtom (type=1017) is a direct child of the Slide container
        let info = self.record.find_child(PptRecordType::SSSlideInfoAtom)?;

        if info.data.len() < 16 {
            return None;
        }

        let d = &info.data;
        let slide_time_ms = u32::from_le_bytes([d[0], d[1], d[2], d[3]]);
        let _sound_id_ref = u32::from_le_bytes([d[4], d[5], d[6], d[7]]);
        let _effect_direction = d[8];
        let _effect_type = d[9];
        let flags = u16::from_le_bytes([d[10], d[11]]);
        let _speed = d[12];

        Some(ParsedSlideTiming {
            advance_time_ms: slide_time_ms,
            advance_on_click: (flags & (1 << 0)) != 0,
            auto_advance: (flags & (1 << 10)) != 0,
            hidden: (flags & (1 << 2)) != 0,
        })
    }
}

/// A parsed comment from a PPT slide.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParsedComment {
    /// Nonnegative comment index.
    pub index: i32,
    /// Author name.
    pub author: String,
    /// Comment text.
    pub text: String,
    /// Author initials.
    pub initials: String,
    /// Year.
    pub year: u16,
    /// Month (1-12).
    pub month: u16,
    /// Day of week (`0` is Sunday).
    pub day_of_week: u16,
    /// Day (1-31).
    pub day: u16,
    /// Hour (0-23).
    pub hour: u16,
    /// Minute (0-59).
    pub minute: u16,
    /// Second (0-59).
    pub second: u16,
    /// Millisecond (0-999).
    pub millisecond: u16,
    /// X position in master units (576/inch).
    pub x: i32,
    /// Y position in master units.
    pub y: i32,
}

/// Parsed per-slide timing information.
#[derive(Debug, Clone)]
pub struct ParsedSlideTiming {
    /// Auto-advance time in milliseconds (0 = no auto-advance).
    pub advance_time_ms: u32,
    /// Whether the slide advances on mouse click.
    pub advance_on_click: bool,
    /// Whether auto-advance is enabled.
    pub auto_advance: bool,
    /// Whether the slide is hidden.
    pub hidden: bool,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::consts::PptRecordType;
    use crate::ppt::records::PptRecord;
    use crate::ppt::slide::SlideData;

    const ROOT_SHAPE_FLAGS: u32 = 0x0A00;
    const CHILD_SHAPE_FLAGS: u32 = 0x0A02;

    fn drawing(children: &[u8]) -> Vec<u8> {
        use crate::escher::writer::{record_type, write_atom, write_container};

        let mut body = Vec::new();
        write_atom(&mut body, 0, 0, record_type::DG, &[0; 8]).unwrap();
        body.extend_from_slice(children);
        let mut drawing = Vec::new();
        write_container(&mut drawing, 0, record_type::DG_CONTAINER, &body).unwrap();
        drawing
    }

    fn create_frame_escher_drawing(
        blip_id: u32,
        interactive_action: Option<u8>,
        external_object_id: Option<u32>,
    ) -> Vec<u8> {
        use crate::escher::writer::{
            PropertyBuilder, ShapeBuilder, record_type, write_atom, write_container,
        };

        let mut shape_children = Vec::new();
        ShapeBuilder::new(75, 42)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();
        let mut properties = PropertyBuilder::new();
        properties.add_simple(0x4104, blip_id as i32);
        properties.write(&mut shape_children).unwrap();
        write_client_anchor(&mut shape_children, 10, 20, 210, 120).unwrap();

        let mut client_data_children = Vec::new();
        if let Some(external_object_id) = external_object_id {
            write_atom(
                &mut client_data_children,
                0,
                0,
                3009,
                &external_object_id.to_le_bytes(),
            )
            .unwrap();
        }
        if let Some(action) = interactive_action {
            let mut interactive_atom = [0u8; 16];
            interactive_atom[8] = action;
            let mut interactive_children = Vec::new();
            write_atom(&mut interactive_children, 0, 0, 4083, &interactive_atom).unwrap();
            write_container(&mut client_data_children, 0, 4082, &interactive_children).unwrap();
        }
        if !client_data_children.is_empty() {
            write_container(
                &mut shape_children,
                0,
                record_type::CLIENT_DATA,
                &client_data_children,
            )
            .unwrap();
        }

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();

        drawing(&shape_container)
    }

    fn create_picture_escher_drawing(blip_id: u32) -> Vec<u8> {
        create_frame_escher_drawing(blip_id, None, None)
    }

    fn create_autoshape_escher_drawing() -> Vec<u8> {
        use crate::escher::writer::{
            PropertyBuilder, ShapeBuilder, record_type, write_atom, write_container,
        };

        let mut shape_children = Vec::new();
        ShapeBuilder::new(13, 44)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();
        let mut properties = PropertyBuilder::new();
        properties.add_simple(0x0147, 32_768);
        properties.add_simple(0x0149, -123);
        properties.write(&mut shape_children).unwrap();
        write_client_anchor(&mut shape_children, 11, 22, 211, 122).unwrap();

        let utf16: Vec<u8> = "Arrow label"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut embedded_text = Vec::new();
        write_atom(&mut embedded_text, 0, 0, 4000, &utf16).unwrap();
        write_container(
            &mut shape_children,
            0,
            record_type::CLIENT_TEXTBOX,
            &embedded_text,
        )
        .unwrap();

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();

        drawing(&shape_container)
    }

    fn create_freeform_escher_drawing() -> Vec<u8> {
        use crate::escher::writer::{PropertyBuilder, ShapeBuilder, record_type, write_container};

        let mut shape_children = Vec::new();
        ShapeBuilder::new(0, 45)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();

        let mut vertices = Vec::new();
        vertices.extend_from_slice(&2u16.to_le_bytes());
        vertices.extend_from_slice(&2u16.to_le_bytes());
        vertices.extend_from_slice(&8u16.to_le_bytes());
        for (x, y) in [(0i32, 0i32), (21600, 21600)] {
            vertices.extend_from_slice(&x.to_le_bytes());
            vertices.extend_from_slice(&y.to_le_bytes());
        }
        let mut properties = PropertyBuilder::new();
        properties.add_simple(0x0140, 0);
        properties.add_simple(0x0141, 0);
        properties.add_simple(0x0142, 21600);
        properties.add_simple(0x0143, 21600);
        properties.add_simple(0x0144, 4);
        properties.add_complex(0x0145, &vertices);
        let segments = [
            2, 0, 2, 0, 2, 0, // IMsoArray header
            0x00, 0x40, // moveTo
            0x00, 0x80, // end
        ];
        properties.add_complex(0x0146, &segments);
        properties.write(&mut shape_children).unwrap();
        write_client_anchor(&mut shape_children, 5, 6, 105, 206).unwrap();

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();
        drawing(&shape_container)
    }

    fn create_animated_escher_drawing() -> Vec<u8> {
        use crate::escher::writer::{ShapeBuilder, record_type, write_container};
        use crate::ppt::animation::{
            AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
            write_animation_info,
        };

        let atom = LegacyAnimationAtom {
            build_type: LegacyAnimationBuild::OneBuild,
            effect: LegacyAnimationEffect::Fade,
            order_id: 2,
            ..LegacyAnimationAtom::default()
        };
        let mut info = AnimationInfo::new();
        info.legacy_atom = Some(atom);
        let (animation, _) = write_animation_info(&info).unwrap();

        let mut shape_children = Vec::new();
        ShapeBuilder::new(1, 88)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();
        write_client_anchor(&mut shape_children, 10, 20, 210, 120).unwrap();
        write_container(&mut shape_children, 0, record_type::CLIENT_DATA, &animation).unwrap();

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();
        drawing(&shape_container)
    }

    fn create_placeholder_escher_drawing(round_trip_records: &[u8]) -> Vec<u8> {
        use crate::escher::writer::{ShapeBuilder, record_type, write_atom, write_container};

        let mut shape_children = Vec::new();
        ShapeBuilder::new(202, 43)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();
        write_client_anchor(&mut shape_children, 15, 25, 315, 125).unwrap();

        let utf16: Vec<u8> = "Slide title"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut embedded_text = Vec::new();
        write_atom(&mut embedded_text, 0, 0, 4000, &utf16).unwrap();
        write_container(
            &mut shape_children,
            0,
            record_type::CLIENT_TEXTBOX,
            &embedded_text,
        )
        .unwrap();

        let mut placeholder_data = Vec::new();
        placeholder_data.extend_from_slice(&3u32.to_le_bytes());
        placeholder_data.push(13); // native slide title placeholder
        placeholder_data.push(2); // quarter size
        placeholder_data.extend_from_slice(&0u16.to_le_bytes());
        let mut client_data_children = Vec::new();
        write_atom(&mut client_data_children, 0, 0, 3011, &placeholder_data).unwrap();
        client_data_children.extend_from_slice(round_trip_records);
        write_container(
            &mut shape_children,
            0,
            record_type::CLIENT_DATA,
            &client_data_children,
        )
        .unwrap();

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();

        drawing(&shape_container)
    }

    fn create_round_trip_placeholder_escher_drawing(
        shape_type: u16,
        round_trip_records: &[u8],
    ) -> Vec<u8> {
        use crate::escher::writer::{ShapeBuilder, record_type, write_container};

        let mut shape_children = Vec::new();
        ShapeBuilder::new(shape_type, 46)
            .with_flags(ROOT_SHAPE_FLAGS)
            .write(&mut shape_children)
            .unwrap();
        write_client_anchor(&mut shape_children, 20, 30, 220, 130).unwrap();
        write_container(
            &mut shape_children,
            0,
            record_type::CLIENT_DATA,
            round_trip_records,
        )
        .unwrap();

        let mut shape_container = Vec::new();
        write_container(
            &mut shape_container,
            0,
            record_type::SP_CONTAINER,
            &shape_children,
        )
        .unwrap();
        drawing(&shape_container)
    }

    fn create_table_escher_drawing() -> Vec<u8> {
        use crate::escher::writer::{
            ShapeBuilder, record_type, write_atom, write_child_anchor, write_container, write_spgr,
        };

        fn shape_container(children: &[u8]) -> Vec<u8> {
            let mut container = Vec::new();
            write_container(&mut container, 0, record_type::SP_CONTAINER, children).unwrap();
            container
        }

        fn table_cell(shape_id: u32, text: &str, left: i32, top: i32) -> Vec<u8> {
            let mut children = Vec::new();
            ShapeBuilder::new(1, shape_id)
                .with_flags(CHILD_SHAPE_FLAGS)
                .write(&mut children)
                .unwrap();
            write_child_anchor(&mut children, left, top, left + 100, top + 50).unwrap();

            let utf16: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();
            let mut embedded_text = Vec::new();
            write_atom(&mut embedded_text, 0, 0, 4000, &utf16).unwrap();
            write_container(
                &mut children,
                0,
                record_type::CLIENT_TEXTBOX,
                &embedded_text,
            )
            .unwrap();
            shape_container(&children)
        }

        let mut patriarch_children = Vec::new();
        write_spgr(&mut patriarch_children, 0, 0, 0, 0).unwrap();
        ShapeBuilder::new(0, 1)
            .with_flags(0x0005)
            .write(&mut patriarch_children)
            .unwrap();
        let patriarch = shape_container(&patriarch_children);

        let mut table_header_children = Vec::new();
        write_spgr(&mut table_header_children, 0, 0, 200, 100).unwrap();
        ShapeBuilder::new(0, 10)
            .with_flags(0x0201)
            .write(&mut table_header_children)
            .unwrap();
        let mut table_properties = Vec::new();
        table_properties.extend_from_slice(&0x039Fu16.to_le_bytes());
        table_properties.extend_from_slice(&1i32.to_le_bytes());
        write_atom(&mut table_header_children, 3, 1, 0xF122, &table_properties).unwrap();
        write_client_anchor(&mut table_header_children, 20, 30, 220, 130).unwrap();
        let table_header = shape_container(&table_header_children);

        let mut table_children = table_header;
        for (shape_id, text, left, top) in [
            (11, "A1", 0, 0),
            (12, "B1", 100, 0),
            (13, "A2", 0, 50),
            (14, "B2", 100, 50),
        ] {
            table_children.extend_from_slice(&table_cell(shape_id, text, left, top));
        }
        let mut table_group = Vec::new();
        write_container(
            &mut table_group,
            0,
            record_type::SPGR_CONTAINER,
            &table_children,
        )
        .unwrap();

        let mut root_group_children = patriarch;
        root_group_children.extend_from_slice(&table_group);
        let mut root_group = Vec::new();
        write_container(
            &mut root_group,
            0,
            record_type::SPGR_CONTAINER,
            &root_group_children,
        )
        .unwrap();

        drawing(&root_group)
    }

    // Helper function to create a test record
    fn create_test_record(
        record_type: PptRecordType,
        data: Vec<u8>,
        children: Vec<PptRecord>,
    ) -> PptRecord {
        PptRecord {
            record_type,
            record_type_raw: record_type as u16,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children,
        }
    }

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn write_client_anchor(
        data: &mut Vec<u8>,
        left: i32,
        top: i32,
        right: i32,
        bottom: i32,
    ) -> std::io::Result<()> {
        crate::ppt::PowerPointClientAnchor::rect(left, top, right, bottom)
            .map_err(|error| {
                std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string())
            })?
            .write_to(data)
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = format!("___PPT{version}")
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        create_test_record(PptRecordType::ProgTags, tag, Vec::new())
    }

    // Helper function to create a basic slide record without children
    fn create_basic_slide_record() -> PptRecord {
        create_test_record(PptRecordType::Slide, vec![0u8; 8], Vec::new())
    }

    // Helper function to create a slide with PPDrawing
    fn create_slide_with_drawing() -> PptRecord {
        let dg = record_bytes(0, 0, 0xf008, &[0; 8]);
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            record_bytes(0x0f, 0, 0xf002, &dg),
            Vec::new(),
        );
        create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![ppdrawing])
    }

    // Helper function to create a slide with text
    fn create_slide_with_text() -> PptRecord {
        // Create a TextCharsAtom with "Test" in UTF-16 LE
        let text_data = vec![
            0x54, 0x00, // 'T'
            0x65, 0x00, // 'e'
            0x73, 0x00, // 's'
            0x74, 0x00, // 't'
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());
        create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text_atom])
    }

    // Helper function to create SlideData
    fn create_slide_data<'doc>(
        record: PptRecord,
        persist_id: u32,
        doc_data: &'doc [u8],
    ) -> SlideData<'doc> {
        SlideData::new_for_test(persist_id, 0, record, doc_data)
    }

    #[test]
    fn test_slide_creation() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.slide_number(), 1);
        assert_eq!(slide.persist_id(), 256);
    }

    #[test]
    fn test_slide_number_accessor() {
        let doc_data = vec![0u8; 512];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 100, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 5);

        assert_eq!(slide.slide_number(), 5);
    }

    #[test]
    fn exposes_inert_slide_library_synchronization_metadata() {
        let server: Vec<u8> = "server-id"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let url: Vec<u8> = "http://example.com/library"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let server = record_bytes(0, 0, 4026, &server);
        let url = record_bytes(0, 1, 4026, &url);
        let mut times = Vec::new();
        for fields in [
            [2026u16, 7, 4, 16, 12, 30, 45, 500],
            [2025u16, 1, 3, 2, 8, 0, 0, 0],
        ] {
            times.extend(fields.into_iter().flat_map(u16::to_le_bytes));
        }
        let atom = record_bytes(0, 0, 0x3715, &times);
        let container = record_bytes(0x0f, 0, 0x3714, &[server, url, atom].concat());
        let sync = PptRecord::parse(&container, 0).unwrap().0;
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![sync]);
        let doc_data = vec![0u8; 32];
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let sync = slide.sync_info().unwrap().unwrap();
        assert_eq!(sync.server_slide_id, "server-id");
        assert_eq!(sync.slide_library_url, "http://example.com/library");
        assert_eq!(sync.server_modified.year, 2026);
        assert_eq!(sync.client_inserted.year, 2025);
        assert!(std::ptr::eq(sync, slide.sync_info().unwrap().unwrap()));
    }

    #[test]
    fn exposes_direct_powerpoint12_slide_master_references() {
        let composite = PptRecord {
            version: 0,
            instance: 0,
            record_type: PptRecordType::RoundTripCompositeMasterId12Atom,
            record_type_raw: 0x041d,
            data_length: 4,
            data: 17u32.to_le_bytes().to_vec(),
            children: Vec::new(),
        };
        let mut content_data = Vec::new();
        content_data.extend_from_slice(&23u32.to_le_bytes());
        content_data.extend_from_slice(&5u16.to_le_bytes());
        content_data.extend_from_slice(&9u16.to_le_bytes());
        let content = PptRecord {
            version: 0,
            instance: 7,
            record_type: PptRecordType::RoundTripContentMasterId12Atom,
            record_type_raw: 0x0422,
            data_length: 8,
            data: content_data,
            children: Vec::new(),
        };
        let slide_record =
            create_test_record(PptRecordType::Slide, Vec::new(), vec![composite, content]);
        let doc_data = vec![0u8; 32];
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let metadata = slide.powerpoint12_round_trip_metadata().unwrap();
        assert_eq!(metadata.composite_master_id, Some(17));
        let content = metadata.content_master.unwrap();
        assert_eq!(content.record_instance, 7);
        assert_eq!(content.main_master_id, 23);
        assert_eq!(content.layout_instance_id, 5);
        assert_eq!(content.unused, 9);
    }

    #[test]
    fn test_persist_id_accessor() {
        let doc_data = vec![0u8; 512];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 999, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.persist_id(), 999);
    }

    #[test]
    fn test_has_drawing_without_ppdrawing() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert!(!slide.has_drawing());
    }

    #[test]
    fn test_has_drawing_with_ppdrawing() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert!(slide.has_drawing());
    }

    #[test]
    fn test_record_accessor() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let rec = slide.record();
        assert_eq!(rec.record_type, PptRecordType::Slide);
    }

    #[test]
    fn test_shapes_empty_slide() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);
    }

    #[test]
    fn test_shapes_lazy_loading() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // First call should initialize
        let shapes1 = slide.shapes().unwrap();
        // Second call should return cached value
        let shapes2 = slide.shapes().unwrap();

        // Both should return the same reference
        assert_eq!(shapes1.len(), shapes2.len());
    }

    #[test]
    fn test_shape_count_empty() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.shape_count().unwrap(), 0);
    }

    #[test]
    fn test_text_extraction_empty_slide() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "");
    }

    #[test]
    fn test_text_extraction_with_text_chars_atom() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_text();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Test");
    }

    #[test]
    fn test_text_lazy_loading() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_text();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // First call should extract text
        let text1 = slide.text().unwrap();
        // Second call should return cached value
        let text2 = slide.text().unwrap();

        assert_eq!(text1, text2);
        assert_eq!(text1, "Test");
    }

    #[test]
    fn test_text_extraction_with_nested_records() {
        let doc_data = vec![0u8; 1024];

        // Create nested structure: Slide -> SlideContainer -> TextCharsAtom
        let text_data = vec![
            0x41, 0x00, // 'A'
            0x42, 0x00, // 'B'
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let container = create_test_record(PptRecordType::SlideAtom, vec![0u8; 8], vec![text_atom]);

        let slide_record = create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![container]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "AB");
    }

    #[test]
    fn test_text_extraction_multiple_text_atoms() {
        let doc_data = vec![0u8; 1024];

        // Create multiple TextCharsAtom records
        let text1_data = vec![
            0x48, 0x00, // 'H'
            0x69, 0x00, // 'i'
        ];
        let text1 = create_test_record(PptRecordType::TextCharsAtom, text1_data, Vec::new());

        let text2_data = vec![
            0x42, 0x00, // 'B'
            0x79, 0x00, // 'y'
            0x65, 0x00, // 'e'
        ];
        let text2 = create_test_record(PptRecordType::TextCharsAtom, text2_data, Vec::new());

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text1, text2]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Both text atoms should be extracted and joined
        assert!(text.contains("Hi"));
        assert!(text.contains("Bye"));
    }

    #[test]
    fn test_slide_with_different_text_atom_types() {
        let doc_data = vec![0u8; 1024];

        // Create TextBytesAtom (ASCII/ANSI encoding)
        let text_bytes = vec![0x54, 0x65, 0x78, 0x74]; // "Text" in ASCII
        let text_bytes_atom =
            create_test_record(PptRecordType::TextBytesAtom, text_bytes, Vec::new());

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text_bytes_atom]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Text");
    }

    #[test]
    fn test_multiple_slide_numbers() {
        let doc_data = vec![0u8; 1024];

        let records: Vec<_> = (0..5).map(|_| create_basic_slide_record()).collect();

        let slides: Vec<_> = records
            .into_iter()
            .enumerate()
            .map(|(i, record)| {
                let slide_data = create_slide_data(record, 100 + i as u32, &doc_data);
                Slide::from_slide_data(slide_data, i + 1)
            })
            .collect();

        // Verify slide numbers are correctly assigned
        for (i, slide) in slides.iter().enumerate() {
            assert_eq!(slide.slide_number(), i + 1);
            assert_eq!(slide.persist_id(), 100 + i as u32);
        }
    }

    #[test]
    fn test_convert_escher_to_shape_enum_with_unknown_type() {
        // This tests that unknown shape types are filtered out
        // We can't easily construct EscherShape objects in tests without
        // implementing complex test data, but we can test the None path
        // through indirect testing via shapes()

        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // Should return empty vec for slide without PPDrawing
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);
    }

    #[test]
    fn referenced_picture_frame_is_exposed_as_picture_shape() {
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_picture_escher_drawing(7),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        let picture = shapes[0].as_picture().expect("picture frame");
        assert_eq!(picture.properties.id, 42);
        assert_eq!(picture.blip_id(), Some(7));
        assert_eq!(picture.properties.x, 10);
        assert_eq!(picture.properties.y, 20);
        assert_eq!(picture.properties.width, 200);
        assert_eq!(picture.properties.height, 100);
    }

    #[test]
    fn autoshape_preserves_native_type_and_sparse_adjustments() {
        use crate::ppt::shapes::{Shape, autoshape::AutoShapeType};

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_autoshape_escher_drawing(),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        let autoshape = shapes[0].as_autoshape().expect("auto shape");
        assert_eq!(autoshape.id(), 44);
        assert_eq!(autoshape.auto_shape_type(), AutoShapeType::Arrow);
        assert_eq!(autoshape.adjustments(), &[32_768, 0, -123]);
        assert_eq!(autoshape.bounds(), (11, 22, 200, 100));
        assert_eq!(autoshape.text(), "Arrow label");
        assert!(autoshape.has_text());
        assert_eq!(Shape::text(autoshape).unwrap(), "Arrow label");
    }

    #[test]
    fn non_primitive_shape_with_vertices_is_exposed_as_freeform_autoshape() {
        use crate::ppt::shapes::{
            Shape,
            autoshape::AutoShapeType,
            geometry::{GeometryRect, ShapePathType},
        };

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_freeform_escher_drawing(),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        let freeform = shapes[0].as_autoshape().expect("freeform auto shape");
        assert_eq!(freeform.auto_shape_type(), AutoShapeType::Custom(0));
        assert_eq!(freeform.properties().id, 45);
        assert_eq!(freeform.bounds(), (5, 6, 100, 200));
        let geometry = freeform.geometry().expect("freeform geometry");
        assert_eq!(
            geometry.coordinate_space(),
            Some(GeometryRect::new(0, 0, 21600, 21600))
        );
        assert_eq!(geometry.path_type(), Some(ShapePathType::Complex));
        assert_eq!(geometry.vertices(), &[(0, 0), (21600, 21600)]);
        assert_eq!(geometry.segment_info(), &[0x4000, 0x8000]);
    }

    #[test]
    fn ole_frame_is_distinguished_from_an_ordinary_picture() {
        use crate::ppt::shapes::{PictureFrameKind, ShapeType};

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_frame_escher_drawing(8, Some(5), Some(77)),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes[0].shape_type(), ShapeType::Object);
        let object = shapes[0].as_object_frame().expect("OLE frame");
        assert_eq!(object.frame_kind(), PictureFrameKind::OleObject);
        assert_eq!(object.external_object_id(), Some(77));
        assert_eq!(object.blip_id(), Some(8));
        assert_eq!(object.properties.x, 10);
        assert_eq!(object.properties.width, 200);
    }

    #[test]
    fn media_frame_preserves_preview_and_external_object_references() {
        use crate::ppt::shapes::{PictureFrameKind, ShapeType};

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_frame_escher_drawing(9, Some(6), Some(88)),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes[0].shape_type(), ShapeType::Media);
        let media = shapes[0].as_media_frame().expect("media frame");
        assert_eq!(media.frame_kind(), PictureFrameKind::Media);
        assert_eq!(media.external_object_id(), Some(88));
        assert_eq!(media.blip_id(), Some(9));
        assert_eq!(media.properties.y, 20);
        assert_eq!(media.properties.height, 100);
    }

    #[test]
    fn external_object_reference_alone_marks_an_ole_frame() {
        use crate::ppt::shapes::{PictureFrameKind, ShapeType};

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_frame_escher_drawing(10, None, Some(99)),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes[0].shape_type(), ShapeType::Object);
        let object = shapes[0].as_object_frame().expect("OLE frame");
        assert_eq!(object.frame_kind(), PictureFrameKind::OleObject);
        assert_eq!(object.external_object_id(), Some(99));
    }

    #[test]
    fn placeholder_client_data_is_exposed_with_text_and_geometry() {
        use crate::ppt::shapes::{PlaceholderSize, PlaceholderType, Shape};

        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_placeholder_escher_drawing(&[]),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        let placeholder = shapes[0].as_placeholder().expect("title placeholder");
        assert_eq!(placeholder.id(), 43);
        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
        assert_eq!(placeholder.placeholder_size(), PlaceholderSize::Quarter);
        assert_eq!(placeholder.index(), Some(3));
        assert_eq!(placeholder.bounds(), (15, 25, 300, 100));
        assert_eq!(shapes[0].text().unwrap(), "Slide title");
        assert!(placeholder.has_text());
    }

    #[test]
    fn powerpoint12_header_footer_placeholder_is_exposed_with_new_identity() {
        use crate::ppt::shapes::{PlaceholderSize, PlaceholderType};
        use crate::ppt::{
            PowerPoint12PlaceholderMetadata, PowerPointHeaderFooterPlaceholder,
            PowerPointNewPlaceholder,
        };

        let records = [
            record_bytes(0, 0, 0x0420, &[10]),
            record_bytes(0, 0, 0x0bdd, &[26]),
        ]
        .concat();
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_round_trip_placeholder_escher_drawing(202, &records),
            Vec::new(),
        );
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        let placeholder = shapes[0].as_placeholder().expect("header placeholder");
        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Header);
        assert_eq!(placeholder.placeholder_size(), PlaceholderSize::Half);
        assert_eq!(placeholder.index(), None);
        assert_eq!(
            shapes[0].powerpoint12_shape_metadata(),
            Some(&PowerPoint12PlaceholderMetadata {
                header_footer: Some(PowerPointHeaderFooterPlaceholder::Header),
                new_placeholder: Some(PowerPointNewPlaceholder::Picture),
                ..PowerPoint12PlaceholderMetadata::default()
            })
        );
    }

    #[test]
    fn legacy_placeholder_identity_precedes_powerpoint12_round_trip_identity() {
        use crate::ppt::PowerPointHeaderFooterPlaceholder;
        use crate::ppt::shapes::PlaceholderType;

        let footer = record_bytes(0, 0, 0x0420, &[9]);
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_placeholder_escher_drawing(&footer),
            Vec::new(),
        );
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        let placeholder = shapes[0].as_placeholder().expect("title placeholder");
        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
        assert_eq!(
            shapes[0]
                .powerpoint12_shape_metadata()
                .and_then(|metadata| metadata.header_footer),
            Some(PowerPointHeaderFooterPlaceholder::Footer)
        );
    }

    #[test]
    fn new_placeholder_identity_is_inert_on_non_placeholder_shapes() {
        use crate::ppt::PowerPointNewPlaceholder;

        let picture = record_bytes(0, 0, 0x0bdd, &[26]);
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_round_trip_placeholder_escher_drawing(1, &picture),
            Vec::new(),
        );
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert!(shapes[0].as_autoshape().is_some());
        assert_eq!(
            shapes[0]
                .powerpoint12_shape_metadata()
                .and_then(|metadata| metadata.new_placeholder),
            Some(PowerPointNewPlaceholder::Picture)
        );
    }

    #[test]
    fn rejects_malformed_or_duplicate_powerpoint12_placeholder_atoms() {
        let mut truncated = record_bytes(0, 0, 0x0420, &[7]);
        truncated[4..8].copy_from_slice(&2u32.to_le_bytes());
        let duplicate_hf = [
            record_bytes(0, 0, 0x0420, &[7]),
            record_bytes(0, 0, 0x0420, &[8]),
        ]
        .concat();
        let duplicate_new = [
            record_bytes(0, 0, 0x0bdd, &[25]),
            record_bytes(0, 0, 0x0bdd, &[26]),
        ]
        .concat();

        for malformed in [
            record_bytes(1, 0, 0x0420, &[7]),
            record_bytes(0, 1, 0x0420, &[7]),
            record_bytes(0, 0, 0x0420, &[]),
            record_bytes(0, 0, 0x0420, &[6]),
            record_bytes(0, 0, 0x0420, &[11]),
            record_bytes(0, 0, 0x0bdd, &[24]),
            record_bytes(0, 0, 0x0bdd, &[27]),
            truncated,
            duplicate_hf,
            duplicate_new,
        ] {
            let doc_data = vec![0u8; 32];
            let ppdrawing = create_test_record(
                PptRecordType::PPDrawing,
                create_round_trip_placeholder_escher_drawing(202, &malformed),
                Vec::new(),
            );
            let slide_record =
                create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
            let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);
            assert!(slide.shapes().is_err());
        }
    }

    #[test]
    fn accepts_every_powerpoint12_placeholder_identity() {
        use crate::ppt::{PowerPointHeaderFooterPlaceholder, PowerPointNewPlaceholder};

        for (id, expected) in [
            (7, PowerPointHeaderFooterPlaceholder::Date),
            (8, PowerPointHeaderFooterPlaceholder::SlideNumber),
            (9, PowerPointHeaderFooterPlaceholder::Footer),
            (10, PowerPointHeaderFooterPlaceholder::Header),
        ] {
            let atom = record_bytes(0, 0, 0x0420, &[id]);
            let drawing = create_round_trip_placeholder_escher_drawing(202, &atom);
            let shapes = litchi_odraw::shape::parse(&drawing).unwrap();
            assert_eq!(
                shapes[0]
                    .powerpoint12_shape_metadata()
                    .unwrap()
                    .and_then(|metadata| metadata.header_footer),
                Some(expected)
            );
        }

        for (id, expected) in [
            (25, PowerPointNewPlaceholder::VerticalObject),
            (26, PowerPointNewPlaceholder::Picture),
        ] {
            let atom = record_bytes(0, 0, 0x0bdd, &[id]);
            let drawing = create_round_trip_placeholder_escher_drawing(1, &atom);
            let shapes = litchi_odraw::shape::parse(&drawing).unwrap();
            assert_eq!(
                shapes[0]
                    .powerpoint12_shape_metadata()
                    .unwrap()
                    .and_then(|metadata| metadata.new_placeholder),
                Some(expected)
            );
        }
    }

    #[test]
    fn exposes_powerpoint12_shape_id_and_custom_layout_checksums() {
        use crate::ppt::PowerPointShapeChecksums;

        let mut checksums = Vec::new();
        checksums.extend_from_slice(&0u32.to_le_bytes());
        checksums.extend_from_slice(&u32::MAX.to_le_bytes());
        let records = [
            record_bytes(0, 0, 0x041f, &u32::MAX.to_le_bytes()),
            record_bytes(0, 0, 0x0426, &checksums),
        ]
        .concat();
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_round_trip_placeholder_escher_drawing(1, &records),
            Vec::new(),
        );
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);
        let metadata = slide.shapes().unwrap()[0]
            .powerpoint12_shape_metadata()
            .unwrap();

        assert_eq!(metadata.shape_id, Some(u32::MAX));
        assert_eq!(
            metadata.custom_layout_checksums,
            Some(PowerPointShapeChecksums {
                shape: 0,
                text: u32::MAX,
            })
        );
    }

    #[test]
    fn rejects_malformed_or_duplicate_powerpoint12_shape_round_trip_atoms() {
        let duplicate_id = [
            record_bytes(0, 0, 0x041f, &1u32.to_le_bytes()),
            record_bytes(0, 0, 0x041f, &2u32.to_le_bytes()),
        ]
        .concat();
        let checksum = [0u8; 8];
        let duplicate_checksums = [
            record_bytes(0, 0, 0x0426, &checksum),
            record_bytes(0, 0, 0x0426, &checksum),
        ]
        .concat();
        let mut truncated_id = record_bytes(0, 0, 0x041f, &[0; 3]);
        truncated_id[4..8].copy_from_slice(&4u32.to_le_bytes());
        let mut truncated_checksums = record_bytes(0, 0, 0x0426, &[0; 7]);
        truncated_checksums[4..8].copy_from_slice(&8u32.to_le_bytes());

        for malformed in [
            record_bytes(1, 0, 0x041f, &0u32.to_le_bytes()),
            record_bytes(0, 1, 0x041f, &0u32.to_le_bytes()),
            record_bytes(0, 0, 0x041f, &[0; 3]),
            record_bytes(0, 0, 0x041f, &[0; 5]),
            record_bytes(1, 0, 0x0426, &checksum),
            record_bytes(0, 1, 0x0426, &checksum),
            record_bytes(0, 0, 0x0426, &[0; 7]),
            record_bytes(0, 0, 0x0426, &[0; 9]),
            truncated_id,
            truncated_checksums,
            duplicate_id,
            duplicate_checksums,
        ] {
            let drawing = create_round_trip_placeholder_escher_drawing(1, &malformed);
            let shapes = litchi_odraw::shape::parse(&drawing).unwrap();
            assert!(shapes[0].powerpoint12_shape_metadata().is_err());
        }
    }

    #[test]
    fn table_group_is_exposed_with_grid_and_text() {
        let doc_data = vec![0u8; 32];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_table_escher_drawing(),
            Vec::new(),
        );
        let record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 1);
        let table = shapes[0].as_table().expect("table group");
        assert_eq!(table.id(), 10);
        assert_eq!(table.rows(), 2);
        assert_eq!(table.columns(), 2);
        assert_eq!(table.cell(0, 0), Some("A1"));
        assert_eq!(table.cell(0, 1), Some("B1"));
        assert_eq!(table.cell(1, 0), Some("A2"));
        assert_eq!(table.cell(1, 1), Some("B2"));
        assert_eq!((table.left(), table.top()), (20, 30));
        assert_eq!((table.width(), table.height()), (200, 100));
    }

    #[test]
    fn test_extract_text_recursive_depth() {
        let doc_data = vec![0u8; 1024];

        // Create deeply nested structure
        let text_data = vec![0x58, 0x00]; // 'X'
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let level3 = create_test_record(PptRecordType::SlideAtom, vec![], vec![text_atom]);

        let level2 = create_test_record(PptRecordType::SlideAtom, vec![], vec![level3]);

        let level1 = create_test_record(PptRecordType::Slide, vec![], vec![level2]);

        let slide_data = create_slide_data(level1, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "X");
    }

    #[test]
    fn test_slide_with_whitespace_only_text() {
        let doc_data = vec![0u8; 1024];

        // Create TextCharsAtom with only whitespace
        let text_data = vec![
            0x20, 0x00, // space
            0x20, 0x00, // space
            0x09, 0x00, // tab
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let slide_record = create_test_record(PptRecordType::Slide, vec![], vec![text_atom]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Whitespace-only text should be filtered out
        assert_eq!(text, "");
    }

    #[test]
    fn test_slide_zero_based_vs_one_based_numbering() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();

        // Test that slide_number is 1-based (display number)
        let slide_data = create_slide_data(record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.slide_number(), 1); // 1-based for user display
    }

    #[test]
    fn test_shape_count_matches_shapes_len() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let shape_count = slide.shape_count().unwrap();
        let shapes_len = slide.shapes().unwrap().len();

        assert_eq!(shape_count, shapes_len);
    }

    #[test]
    fn test_text_and_shapes_independent_caching() {
        let doc_data = vec![0u8; 1024];

        // Create slide with both text and PPDrawing
        let text_data = vec![0x41, 0x00]; // 'A'
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let dg = record_bytes(0, 0, 0xf008, &[0; 8]);
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            record_bytes(0x0f, 0, 0xf002, &dg),
            Vec::new(),
        );

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![], vec![text_atom, ppdrawing]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        // Access text first
        let text = slide.text().unwrap();
        assert_eq!(text, "A");

        // Then access shapes - should work independently
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);

        // Access again to verify both caches work
        let text2 = slide.text().unwrap();
        let shapes2 = slide.shapes().unwrap();

        assert_eq!(text, text2);
        assert_eq!(shapes.len(), shapes2.len());
    }

    #[test]
    fn test_slide_with_cstring_record() {
        let doc_data = vec![0u8; 1024];

        // CString records contain UTF-16LE text.
        let cstring_data = "Hi😀".encode_utf16().flat_map(u16::to_le_bytes).collect();
        let cstring = create_test_record(PptRecordType::CString, cstring_data, Vec::new());

        let slide_record = create_test_record(PptRecordType::Slide, vec![], vec![cstring]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Hi😀");
    }

    #[test]
    fn test_large_persist_id() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();

        // Test with large persist ID
        let large_id = u32::MAX - 1;
        let slide_data = create_slide_data(record, large_id, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.persist_id(), large_id);
    }

    #[test]
    fn test_slide_with_empty_data() {
        let doc_data = vec![0u8; 0]; // Empty document data
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // Should still work with basic accessors
        assert_eq!(slide.slide_number(), 1);
        assert_eq!(slide.persist_id(), 256);
        assert!(!slide.has_drawing());
    }

    #[test]
    fn exposes_inert_shape_animations_from_the_slide() {
        let doc_data = vec![0u8; 1024];
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            create_animated_escher_drawing(),
            Vec::new(),
        );
        let slide_record = create_test_record(PptRecordType::Slide, Vec::new(), vec![ppdrawing]);
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        let animations = slide.animations().unwrap();
        assert_eq!(animations.len(), 1);
        assert_eq!(animations[0].shape_id, 88);
        let atom = animations[0].animation.legacy_atom.as_ref().unwrap();
        assert_eq!(
            atom.effect,
            crate::ppt::animation::LegacyAnimationEffect::Fade
        );
        assert_eq!(atom.order_id, 2);
    }

    #[test]
    fn exposes_powerpoint_2002_animation_extension_from_programmable_tags() {
        use crate::escher::writer::{write_atom, write_container};
        use crate::ppt::animation::{
            BuildList, ExtendedTimeNode, TimeNodeAtom, TimeNodeKind, write_build_list,
            write_extended_time_node,
        };
        use crate::ppt::writer::comments::{SlideComment, build_slide_comments};

        let timing = ExtendedTimeNode {
            atom: TimeNodeAtom {
                node_type: Some(TimeNodeKind::Sequential),
                duration_ms: Some(2_000),
                ..TimeNodeAtom::default()
            },
            ..ExtendedTimeNode::default()
        };
        let comment = SlideComment::new("Ada Lovelace", "Animate this", 12, 34);
        let mut extension_data = build_slide_comments(&[comment]).unwrap();
        extension_data.extend(write_extended_time_node(&timing).unwrap());
        extension_data.extend(write_build_list(&BuildList::new()).unwrap());

        let tag_name: Vec<u8> = "___PPT10"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut prog_binary_children = Vec::new();
        write_atom(&mut prog_binary_children, 0, 0, 4026, &tag_name).unwrap();
        write_atom(
            &mut prog_binary_children,
            0,
            0,
            PptRecordType::BinaryTagData.as_u16(),
            &extension_data,
        )
        .unwrap();
        let mut prog_tags_children = Vec::new();
        write_container(
            &mut prog_tags_children,
            0,
            PptRecordType::ProgBinaryTag.as_u16(),
            &prog_binary_children,
        )
        .unwrap();
        let mut slide_children = Vec::new();
        write_container(
            &mut slide_children,
            0,
            PptRecordType::ProgTags.as_u16(),
            &prog_tags_children,
        )
        .unwrap();
        let mut bytes = Vec::new();
        write_container(
            &mut bytes,
            0,
            PptRecordType::Slide.as_u16(),
            &slide_children,
        )
        .unwrap();

        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        let doc_data = Vec::new();
        let slide = Slide::from_slide_data(create_slide_data(record, 256, &doc_data), 1);
        let extension = slide.animation_extension().unwrap().unwrap();
        assert_eq!(extension.time_node, Some(timing));
        assert_eq!(extension.build_list, Some(BuildList::new()));
        let comments = slide.comments().unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].author, "Ada Lovelace");
        assert_eq!(comments[0].text, "Animate this");
        assert_eq!(slide.text().unwrap(), "");
    }

    #[test]
    fn ignores_powerpoint10_slide_settings_in_other_tag_versions() {
        let flags = record_bytes(
            0,
            0,
            PptRecordType::SlideFlags10Atom.as_u16(),
            &[3, 0, 0, 0],
        );
        let slide_record = create_test_record(
            PptRecordType::Slide,
            Vec::new(),
            vec![prog_tags_record(9, &flags)],
        );
        let doc_data = Vec::new();
        let slide = Slide::from_slide_data(create_slide_data(slide_record, 256, &doc_data), 1);

        assert_eq!(slide.animation_extension().unwrap(), None);
    }

    #[test]
    fn truncated_comment_atoms_are_rejected_without_panicking() {
        let mut child = Vec::new();
        child.extend(0u16.to_le_bytes());
        child.extend(PptRecordType::Comment2000Atom.as_u16().to_le_bytes());
        child.extend(28u32.to_le_bytes());
        child.push(0);

        let mut data = Vec::new();
        data.extend(0x000Fu16.to_le_bytes());
        data.extend(PptRecordType::Comment2000.as_u16().to_le_bytes());
        data.extend(u32::try_from(child.len()).unwrap().to_le_bytes());
        data.extend(child);

        let extension = prog_tags_record(10, &data);
        let slide = create_test_record(PptRecordType::Slide, Vec::new(), vec![extension]);
        assert!(crate::ppt::comments::parse_slide_comments(&slide).is_err());
    }

    #[test]
    fn test_slide_text_extraction_preserves_order() {
        let doc_data = vec![0u8; 1024];

        // Create multiple text atoms in specific order
        let text1 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x31, 0x00], // '1'
            Vec::new(),
        );

        let text2 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x32, 0x00], // '2'
            Vec::new(),
        );

        let text3 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x33, 0x00], // '3'
            Vec::new(),
        );

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![], vec![text1, text2, text3]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Text should be extracted in order and joined with newlines
        assert!(text.contains('1'));
        assert!(text.contains('2'));
        assert!(text.contains('3'));
        // Verify order is preserved
        let pos1 = text.find('1').unwrap();
        let pos2 = text.find('2').unwrap();
        let pos3 = text.find('3').unwrap();
        assert!(pos1 < pos2);
        assert!(pos2 < pos3);
    }
}

//! RTF story event and legacy drawing output.

#![allow(
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "serialization helpers deliberately rebind a working value as the output is assembled"
)]
use super::super::{
    Annotation, BodyEvent, BodyEventKind, BodyStoryEvent, BookmarkTable, CustomXmlTag,
    DrawingStoryTextMode, EditableRegion, EmbeddedObject, Field, FieldOwner, FormField,
    GeneratedListMarker, LegacyCalloutAttachment, LegacyCalloutType, LegacyDrawing,
    LegacyDrawingArrow, LegacyDrawingArrowFill, LegacyDrawingColor, LegacyDrawingGeometry,
    LegacyDrawingLineStyle, LegacyDrawingPoint, LegacyDrawingPrimitive, LegacyDrawingProperties,
    LegacyHorizontalAnchor, LegacyTextBox, LegacyTextDirection, LegacyVerticalAnchor,
    MAX_EMBEDDED_OBJECTS, MAX_LEGACY_DRAWINGS, MAX_PICTURE_COMPATIBILITY_RECORDS, MathZone,
    NavigationEntry, Note, ObjectKind, ObjectResultKind, Picture, PictureCompatibilityKind,
    PictureCompatibilityRecord, PictureShapeProperties, ProtectionRange, Revision, RevisionType,
    RtfWriter, Section, Shape, ShapeGroup, ShapeGroupInfo, ShapeType, SoftBreakKind, StoryDrawing,
    StyleBlock, Write, field, form_field, invalid_story_reference, io, navigation_entry, section,
    take_story_item,
};

impl<W: Write> RtfWriter<W> {
    pub(in super::super) fn write_annotation_value(
        &mut self,
        control: &str,
        value: Option<&str>,
    ) -> io::Result<()> {
        let Some(value) = value else {
            return Ok(());
        };
        self.write_str("{\\*")?;
        self.write_control_word(control, None)?;
        self.write_str(" ")?;
        self.write_destination_text(value)?;
        self.write_str("}")
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "the writer threads the full block and markup context through the pipeline"
    )]
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "remaining variants share the same fallback by design"
    )]
    pub(in super::super::super) fn write_blocks_with_markup(
        &mut self,
        blocks: &[StyleBlock<'_>],
        body_boundaries: &[crate::story::Boundary],
        bookmarks: &BookmarkTable<'_>,
        custom_xml_tags: &[CustomXmlTag<'_>],
        math_zones: &[MathZone<'_>],
        protection_ranges: &[ProtectionRange<'_>],
        editable_regions: &[EditableRegion<'_>],
        annotations: &[Annotation<'_>],
        notes: &[Note<'_>],
        revisions: &[Revision<'_>],
        navigation_entries: &[NavigationEntry<'_>],
        generated_list_markers: &[GeneratedListMarker<'_>],
        shapes: &[Shape<'_>],
        shape_groups: &[ShapeGroup<'_>],
        drawing_order: &[StoryDrawing],
        picture_compatibility_records: &[PictureCompatibilityRecord],
        pictures: &[Picture<'_>],
        objects: &[EmbeddedObject<'_>],
        legacy_text_boxes: &[LegacyTextBox<'_>],
        legacy_drawings: &[LegacyDrawing<'_>],
        form_fields: &[FormField<'_>],
        fields: &[Field<'_>],
        sections: &[Section<'_>],
        body_story_events: &[BodyStoryEvent],
        opaque_nodes: &[crate::opaque::Node],
    ) -> io::Result<()> {
        if bookmarks.bookmarks().is_empty()
            && custom_xml_tags.is_empty()
            && math_zones.is_empty()
            && protection_ranges.is_empty()
            && editable_regions.is_empty()
            && annotations.is_empty()
            && notes.is_empty()
            && revisions.is_empty()
            && navigation_entries.is_empty()
            && generated_list_markers.is_empty()
            && shapes.iter().all(|shape| shape.is_background)
            && shape_groups.is_empty()
            && picture_compatibility_records.is_empty()
            && objects.is_empty()
            && legacy_text_boxes.is_empty()
            && legacy_drawings.is_empty()
            && form_fields.is_empty()
            && fields
                .iter()
                .all(|field| !matches!(field.owner, FieldOwner::Body))
            && body_story_events.is_empty()
            && opaque_nodes.is_empty()
        {
            let mut boundary = 0usize;
            let mut body_position = 0usize;
            for block in blocks {
                self.write_style_block_with_boundaries(
                    block,
                    body_position,
                    body_boundaries,
                    &mut boundary,
                )?;
                body_position = body_position.checked_add(block.text.len()).ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF body text size overflow")
                })?;
            }
            if boundary != body_boundaries.len() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF body text boundaries are incomplete",
                ));
            }
            return Ok(());
        }

        let body: String = blocks.iter().map(|block| block.text.as_ref()).collect();
        let event_count = bookmarks
            .bookmarks()
            .len()
            .saturating_add(annotations.len())
            .saturating_add(notes.len())
            .saturating_add(revisions.len())
            .saturating_mul(2);
        let event_count = event_count.saturating_add(navigation_entries.len());
        let event_count = event_count.saturating_add(custom_xml_tags.len().saturating_mul(2));
        let event_count = event_count.saturating_add(math_zones.len());
        let event_count = event_count.saturating_add(protection_ranges.len().saturating_mul(2));
        let event_count = event_count.saturating_add(editable_regions.len().saturating_mul(2));
        let event_count = event_count.saturating_add(shapes.len());
        let event_count = event_count.saturating_add(shape_groups.len());
        let event_count = event_count.saturating_add(picture_compatibility_records.len());
        let event_count = event_count.saturating_add(objects.len());
        let event_count = event_count.saturating_add(legacy_drawings.len());
        let event_count = event_count.saturating_add(form_fields.len().saturating_mul(2));
        let event_count = event_count.saturating_add(body_story_events.len());
        let event_count = event_count.saturating_add(opaque_nodes.len());
        let mut events = Vec::new();
        events.try_reserve(event_count).map_err(|_err| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body event table exceeds available memory",
            )
        })?;
        if notes.len() > section::MAX_NOTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF note count exceeds the safety limit",
            ));
        }
        let mut previous_note_position = None;
        let mut note_text_bytes = 0usize;
        for note in notes {
            note.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(note.position..note.position).is_none()
                || previous_note_position.is_some_and(|position| position > note.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF notes are outside or out of main-story order",
                ));
            }
            note_text_bytes = note_text_bytes
                .checked_add(note.text_bytes().ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF note text size overflow")
                })?)
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF note text size overflow")
                })?;
            if note_text_bytes > section::MAX_NOTE_TEXT_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF note aggregate text exceeds the safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: note.position,
                order: 1,
                kind: BodyEventKind::Note(note),
            });
            previous_note_position = Some(note.position);
        }
        let expected_drawings = shapes
            .iter()
            .filter(|shape| !shape.is_background)
            .count()
            .saturating_add(shape_groups.len());
        if drawing_order.len() != expected_drawings {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body drawing order is incomplete",
            ));
        }
        if fields.len() > field::MAX_GENERIC_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generic field count exceeds the safety limit",
            ));
        }
        if objects.len() > MAX_EMBEDDED_OBJECTS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF embedded object count exceeds the safety limit",
            ));
        }
        let mut previous_object_position = None;
        for object in objects {
            object
                .validate(&body, pictures.len())
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if previous_object_position.is_some_and(|position| position > object.position) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF embedded objects are not ordered by body position",
                ));
            }
            for picture_index in &object.result_picture_indices {
                pictures
                    .get(*picture_index)
                    .ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF embedded object references a missing result picture",
                        )
                    })?
                    .validate()
                    .map_err(|error| {
                        io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                    })?;
            }
            events.push(BodyEvent {
                offset: object.position,
                order: 1,
                kind: BodyEventKind::Object(object, pictures),
            });
            previous_object_position = Some(object.position);
        }
        if picture_compatibility_records.len() > MAX_PICTURE_COMPATIBILITY_RECORDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF picture-compatibility record count exceeds the safety limit",
            ));
        }
        let mut previous_picture_record = None;
        for record in picture_compatibility_records {
            record
                .validate(&body, pictures.len())
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if previous_picture_record.is_some_and(|previous: &PictureCompatibilityRecord| {
                previous.position > record.position
                    || (previous.position == record.position && previous.kind == record.kind)
            }) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF picture-compatibility records are duplicated or out of body order",
                ));
            }
            let picture = pictures.get(record.picture_index).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF picture-compatibility record references a missing picture",
                )
            })?;
            picture
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            events.push(BodyEvent {
                offset: record.position,
                order: 1,
                kind: BodyEventKind::PictureCompatibility(record, picture),
            });
            previous_picture_record = Some(record);
        }
        if form_fields.len() > form_field::MAX_FORM_FIELDS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF form-field count exceeds the safety limit",
            ));
        }
        let mut form_field_bytes = 0usize;
        let mut form_field_ranges: Vec<&FormField<'_>> = form_fields.iter().collect();
        form_field_ranges.sort_by_key(|field| (field.position, field.range_end));
        let mut previous_form_end = 0usize;
        for field in form_field_ranges {
            field
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let result = body.get(field.position..field.range_end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field range is outside body text or splits a character",
                )
            })?;
            if result != field.result_text {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result does not match its visible body range",
                ));
            }
            if field.position != field.range_end && field.position < previous_form_end {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field result ranges cannot overlap",
                ));
            }
            if field.position != field.range_end {
                previous_form_end = field.range_end;
            }
            form_field_bytes = form_field_bytes
                .checked_add(field.text_bytes().ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF form-field aggregate size overflow",
                    )
                })?;
            if form_field_bytes > form_field::MAX_FORM_FIELD_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF form-field aggregate text exceeds the safety limit",
                ));
            }
            let empty = field.position == field.range_end;
            events.push(BodyEvent {
                offset: field.position,
                order: 1,
                kind: BodyEventKind::FormFieldStart(field),
            });
            events.push(BodyEvent {
                offset: field.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::FormFieldEnd,
            });
        }
        if generated_list_markers.len() > crate::generated_list_marker::MAX_GENERATED_LIST_MARKERS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF generated list-marker count exceeds the safety limit",
            ));
        }
        let mut generated_marker_bytes = 0usize;
        let mut previous_generated_marker: Option<&GeneratedListMarker<'_>> = None;
        for marker in generated_list_markers {
            marker
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(marker.position..marker.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list-marker position is not a UTF-8 body boundary",
                ));
            }
            if previous_generated_marker.is_some_and(|previous| {
                previous.position > marker.position
                    || (previous.position == marker.position && previous.kind == marker.kind)
            }) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list markers are duplicated or out of body order",
                ));
            }
            generated_marker_bytes = generated_marker_bytes
                .checked_add(marker.text.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF generated list-marker text size overflow",
                    )
                })?;
            if generated_marker_bytes
                > crate::generated_list_marker::MAX_GENERATED_LIST_MARKER_TOTAL_BYTES
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF generated list-marker text exceeds the aggregate safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: marker.position,
                order: 1,
                kind: BodyEventKind::GeneratedListMarker(marker),
            });
            previous_generated_marker = Some(marker);
        }

        if legacy_text_boxes.len() > crate::legacy_text_box::MAX_LEGACY_TEXT_BOXES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF legacy text-box count exceeds the safety limit",
            ));
        }
        let mut legacy_text_box_bytes = 0usize;
        let mut previous_legacy_text_box_position = None;
        for text_box in legacy_text_boxes {
            text_box
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(text_box.position..text_box.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text-box position is not a UTF-8 body boundary",
                ));
            }
            if previous_legacy_text_box_position
                .is_some_and(|position| position > text_box.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text boxes are not ordered by body position",
                ));
            }
            legacy_text_box_bytes = legacy_text_box_bytes
                .checked_add(text_box.text.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF legacy text-box text size overflow",
                    )
                })?;
            if legacy_text_box_bytes > crate::legacy_text_box::MAX_LEGACY_TEXT_BOX_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy text-box text exceeds the aggregate safety limit",
                ));
            }
            events.push(BodyEvent {
                offset: text_box.position,
                order: 1,
                kind: BodyEventKind::LegacyTextBox(text_box),
            });
            previous_legacy_text_box_position = Some(text_box.position);
        }

        if legacy_drawings.len() > MAX_LEGACY_DRAWINGS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF legacy drawing count exceeds the safety limit",
            ));
        }
        let mut previous_legacy_drawing_position = None;
        for drawing in legacy_drawings {
            drawing
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(drawing.position..drawing.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy drawing position is not a UTF-8 body boundary",
                ));
            }
            if previous_legacy_drawing_position.is_some_and(|position| position > drawing.position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF legacy drawings are not ordered by body position",
                ));
            }
            events.push(BodyEvent {
                offset: drawing.position,
                order: 1,
                kind: BodyEventKind::LegacyDrawing(drawing),
            });
            previous_legacy_drawing_position = Some(drawing.position);
        }

        if navigation_entries.len() > navigation_entry::MAX_NAVIGATION_ENTRIES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF navigation-entry count limit exceeded",
            ));
        }
        let mut navigation_text_bytes = 0usize;
        let body_navigation: Vec<bool> = (0..navigation_entries.len())
            .map(|index| {
                body_story_events.iter().any(|event| {
                matches!(event, BodyStoryEvent::NavigationEntry(value) if *value == index)
            })
            })
            .collect();
        let body_revisions: Vec<bool> = (0..revisions.len())
            .map(|index| {
                body_story_events.iter().any(|event| {
                    matches!(
                        event,
                        BodyStoryEvent::RevisionStart(value)
                            | BodyStoryEvent::RevisionEnd(value)
                            | BodyStoryEvent::RevisionDeletion(value)
                            if *value == index
                    )
                })
            })
            .collect();
        for (entry, is_body_entry) in navigation_entries.iter().zip(&body_navigation) {
            entry
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if *is_body_entry && body.get(entry.position()..entry.position()).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry position is outside body text or splits a character",
                ));
            }
            navigation_text_bytes = navigation_text_bytes
                .checked_add(entry.text_bytes().ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "navigation-entry size overflow",
                    )
                })?)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "navigation-entry size overflow",
                    )
                })?;
            if navigation_text_bytes > navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF navigation-entry aggregate text limit exceeded",
                ));
            }
            if *is_body_entry {
                events.push(BodyEvent {
                    offset: entry.position(),
                    order: 1,
                    kind: BodyEventKind::NavigationEntry(entry),
                });
            }
        }
        for bookmark in bookmarks.bookmarks() {
            let end = bookmark
                .position
                .checked_add(bookmark.content.len())
                .ok_or_else(|| {
                    io::Error::new(io::ErrorKind::InvalidInput, "RTF bookmark range overflow")
                })?;
            let content = body.get(bookmark.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark range is outside body text or splits a character",
                )
            })?;
            if content != bookmark.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF bookmark content does not match its body range",
                ));
            }
            let empty = bookmark.content.is_empty();
            events.push(BodyEvent {
                offset: bookmark.position,
                order: 1,
                kind: BodyEventKind::BookmarkStart(bookmark),
            });
            events.push(BodyEvent {
                offset: end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::BookmarkEnd(bookmark),
            });
        }
        if custom_xml_tags.len() > crate::custom_xml::MAX_CUSTOM_XML_TAGS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF custom XML tag count exceeds the safety limit",
            ));
        }
        for tag in custom_xml_tags {
            tag.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = tag.position.checked_add(tag.content.len()).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tag range overflow",
                )
            })?;
            let content = body.get(tag.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tag range is outside body text or splits a character",
                )
            })?;
            if content != tag.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tag content does not match its body range",
                ));
            }
        }
        {
            let mut xml_stack: Vec<usize> = Vec::new();
            for event in body_story_events {
                match *event {
                    BodyStoryEvent::CustomXmlOpen(index) => {
                        if index >= custom_xml_tags.len() {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML story event references a missing tag",
                            ));
                        }
                        xml_stack.push(index);
                    },
                    BodyStoryEvent::CustomXmlClose(index) => {
                        let expected = xml_stack.pop();
                        if expected != Some(index) {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML tags are not properly nested",
                            ));
                        }
                    },
                    _ => {},
                }
            }
            if !xml_stack.is_empty() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF custom XML tags are not properly nested",
                ));
            }
        }
        if math_zones.len() > crate::math::MAX_MATH_ZONES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF math zone count exceeds the safety limit",
            ));
        }
        for zone in math_zones {
            zone.validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(zone.position..zone.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF math zone anchor is outside body text or splits a character",
                ));
            }
        }
        if protection_ranges.len() > crate::protection_range::MAX_PROTECTION_RANGES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF protection-range count exceeds the safety limit",
            ));
        }
        for range in protection_ranges {
            range
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = range
                .position
                .checked_add(range.content.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF protection-range range overflow",
                    )
                })?;
            let content = body.get(range.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF protection range is outside body text or splits a character",
                )
            })?;
            if content != range.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF protection-range content does not match its body range",
                ));
            }
        }
        if editable_regions.len() > crate::editable_region::MAX_EDITABLE_REGIONS {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF editable-region count exceeds the safety limit",
            ));
        }
        for region in editable_regions {
            region
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            let end = region
                .position
                .checked_add(region.content.len())
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF editable-region range overflow",
                    )
                })?;
            let content = body.get(region.position..end).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable region is outside body text or splits a character",
                )
            })?;
            if content != region.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable-region content does not match its body range",
                ));
            }
        }
        {
            let mut region_stack: Vec<usize> = Vec::new();
            for event in body_story_events {
                match *event {
                    BodyStoryEvent::EditableRegionStart(index) => {
                        if index >= editable_regions.len() {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF editable-region story event references a missing region",
                            ));
                        }
                        region_stack.push(index);
                    },
                    BodyStoryEvent::EditableRegionEnd(index) => {
                        let expected = region_stack.pop();
                        if expected != Some(index) {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF editable regions are not properly nested",
                            ));
                        }
                    },
                    _ => {},
                }
            }
            if !region_stack.is_empty() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF editable regions are not properly nested",
                ));
            }
        }
        for annotation in annotations {
            annotation
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
            if annotation.range_end < annotation.position
                || body
                    .get(annotation.position..annotation.range_end)
                    .is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF annotation range is outside body text or splits a character",
                ));
            }
            let empty = annotation.position == annotation.range_end;
            events.push(BodyEvent {
                offset: annotation.position,
                order: 1,
                kind: BodyEventKind::AnnotationStart(annotation),
            });
            events.push(BodyEvent {
                offset: annotation.range_end,
                order: if empty { 2 } else { 0 },
                kind: BodyEventKind::AnnotationEnd(annotation),
            });
        }
        let mut revision_ranges: Vec<&Revision<'_>> = revisions
            .iter()
            .zip(&body_revisions)
            .filter_map(|(revision, is_body_revision)| {
                (*is_body_revision && revision.revision_type == RevisionType::Insertion)
                    .then_some(revision)
            })
            .collect();
        revision_ranges.sort_by_key(|revision| (revision.position, revision.range_end));
        let mut previous_end = 0usize;
        for revision in revision_ranges {
            if revision.range_end <= revision.position
                || revision.position < previous_end
                || body.get(revision.position..revision.range_end).is_none()
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision ranges overlap, leave the body, or split a character",
                ));
            }
            let content = body
                .get(revision.position..revision.range_end)
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF revision range is outside body text or splits a character",
                    )
                })?;
            if content != revision.content {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF revision content does not match its body range",
                ));
            }
            previous_end = revision.range_end;
            events.push(BodyEvent {
                offset: revision.position,
                order: 1,
                kind: BodyEventKind::RevisionStart(revision),
            });
            events.push(BodyEvent {
                offset: revision.range_end,
                order: 0,
                kind: BodyEventKind::RevisionEnd,
            });
        }
        for revision in
            revisions
                .iter()
                .zip(&body_revisions)
                .filter_map(|(revision, is_body_revision)| {
                    (*is_body_revision && revision.revision_type == RevisionType::Deletion)
                        .then_some(revision)
                })
        {
            revision
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            if body.get(..revision.position).is_none() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF deletion position is outside body text or splits a character",
                ));
            }
            events.push(BodyEvent {
                offset: revision.position,
                order: 0,
                kind: BodyEventKind::RevisionDeletion(revision),
            });
        }
        events.clear();
        let bookmark_items = bookmarks.bookmarks();
        let mut saw_shapes = vec![false; shapes.len()];
        let mut saw_groups = vec![false; shape_groups.len()];
        let mut saw_fields = vec![false; fields.len()];
        let mut saw_bookmark_starts = vec![false; bookmark_items.len()];
        let mut saw_bookmark_ends = vec![false; bookmark_items.len()];
        let mut saw_custom_xml_opens = vec![false; custom_xml_tags.len()];
        let mut saw_custom_xml_closes = vec![false; custom_xml_tags.len()];
        let mut saw_math_zones = vec![false; math_zones.len()];
        let mut saw_protection_starts = vec![false; protection_ranges.len()];
        let mut saw_protection_ends = vec![false; protection_ranges.len()];
        let mut saw_editable_starts = vec![false; editable_regions.len()];
        let mut saw_editable_ends = vec![false; editable_regions.len()];
        let mut saw_annotation_starts = vec![false; annotations.len()];
        let mut saw_annotation_ends = vec![false; annotations.len()];
        let mut saw_notes = vec![false; notes.len()];
        let mut saw_objects = vec![false; objects.len()];
        let mut saw_picture_records = vec![false; picture_compatibility_records.len()];
        let mut saw_form_starts = vec![false; form_fields.len()];
        let mut saw_form_ends = vec![false; form_fields.len()];
        let mut saw_revision_starts = vec![false; revisions.len()];
        let mut saw_revision_ends = vec![false; revisions.len()];
        let mut saw_revision_deletions = vec![false; revisions.len()];
        let mut saw_generated_markers = vec![false; generated_list_markers.len()];
        let mut saw_legacy_text_boxes = vec![false; legacy_text_boxes.len()];
        let mut saw_legacy_drawings = vec![false; legacy_drawings.len()];
        let mut saw_navigation_entries = vec![false; navigation_entries.len()];
        let mut ordered_drawings = Vec::with_capacity(expected_drawings);
        let mut previous_story_position = None;
        let first_section_is_boundary_scoped = body_story_events.iter().any(|event| {
            matches!(
                event,
                BodyStoryEvent::SectionBreak(section_break)
                    if section_break.next_section == Some(0)
            )
        });
        let mut next_section_index = if first_section_is_boundary_scoped {
            0
        } else {
            usize::from(!sections.is_empty())
        };
        for story_event in body_story_events {
            let (position, kind) = match *story_event {
                BodyStoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    let shape = take_story_item(shapes, &mut saw_shapes, index)?;
                    if shape.is_background {
                        return Err(invalid_story_reference());
                    }
                    ordered_drawings.push(StoryDrawing::Shape(index));
                    (shape.position, BodyEventKind::Shape(shape))
                },
                BodyStoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    let group = take_story_item(shape_groups, &mut saw_groups, index)?;
                    ordered_drawings.push(StoryDrawing::ShapeGroup(index));
                    (group.position, BodyEventKind::ShapeGroup(group))
                },
                BodyStoryEvent::Field(index) => {
                    let field = take_story_item(fields, &mut saw_fields, index)?;
                    if !matches!(field.owner, FieldOwner::Body) {
                        return Err(invalid_story_reference());
                    }
                    (field.position, BodyEventKind::GenericField(field))
                },
                BodyStoryEvent::PageBreak(page_break) => {
                    (page_break.position, BodyEventKind::PageBreak)
                },
                BodyStoryEvent::SoftBreak(soft_break) => {
                    (soft_break.position, BodyEventKind::SoftBreak(soft_break))
                },
                BodyStoryEvent::ColumnBreak(column_break) => {
                    (column_break.position, BodyEventKind::ColumnBreak)
                },
                BodyStoryEvent::SectionBreak(section_break) => {
                    let section = match section_break.next_section {
                        None => None,
                        Some(index) if index == next_section_index => {
                            let section = sections.get(index).ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF section boundary references a missing section",
                                )
                            })?;
                            next_section_index =
                                next_section_index.checked_add(1).ok_or_else(|| {
                                    io::Error::new(
                                        io::ErrorKind::InvalidInput,
                                        "RTF section boundary index overflow",
                                    )
                                })?;
                            Some(section)
                        },
                        Some(_) => {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF section boundary has an invalid or out-of-order section reference",
                            ));
                        },
                    };
                    (section_break.position, BodyEventKind::SectionBreak(section))
                },
                BodyStoryEvent::BookmarkStart(index) => {
                    let bookmark =
                        take_story_item(bookmark_items, &mut saw_bookmark_starts, index)?;
                    (bookmark.position, BodyEventKind::BookmarkStart(bookmark))
                },
                BodyStoryEvent::BookmarkEnd(index) => {
                    let bookmark = take_story_item(bookmark_items, &mut saw_bookmark_ends, index)?;
                    (
                        bookmark
                            .position
                            .checked_add(bookmark.content.len())
                            .ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF bookmark range overflow",
                                )
                            })?,
                        BodyEventKind::BookmarkEnd(bookmark),
                    )
                },
                BodyStoryEvent::CustomXmlOpen(index) => {
                    let tag = take_story_item(custom_xml_tags, &mut saw_custom_xml_opens, index)?;
                    (tag.position, BodyEventKind::CustomXmlOpen(tag))
                },
                BodyStoryEvent::CustomXmlClose(index) => {
                    let tag = take_story_item(custom_xml_tags, &mut saw_custom_xml_closes, index)?;
                    (
                        tag.position.checked_add(tag.content.len()).ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF custom XML tag range overflow",
                            )
                        })?,
                        BodyEventKind::CustomXmlClose(tag),
                    )
                },
                BodyStoryEvent::MathZone(index) => {
                    let zone = take_story_item(math_zones, &mut saw_math_zones, index)?;
                    (zone.position, BodyEventKind::MathZone(zone))
                },
                BodyStoryEvent::ProtectionRangeStart(index) => {
                    let range =
                        take_story_item(protection_ranges, &mut saw_protection_starts, index)?;
                    (range.position, BodyEventKind::ProtectionRangeStart(range))
                },
                BodyStoryEvent::ProtectionRangeEnd(index) => {
                    let range =
                        take_story_item(protection_ranges, &mut saw_protection_ends, index)?;
                    (
                        range
                            .position
                            .checked_add(range.content.len())
                            .ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF protection-range range overflow",
                                )
                            })?,
                        BodyEventKind::ProtectionRangeEnd(range),
                    )
                },
                BodyStoryEvent::EditableRegionStart(index) => {
                    let region =
                        take_story_item(editable_regions, &mut saw_editable_starts, index)?;
                    (region.position, BodyEventKind::EditableRegionStart(region))
                },
                BodyStoryEvent::EditableRegionEnd(index) => {
                    let region = take_story_item(editable_regions, &mut saw_editable_ends, index)?;
                    (
                        region
                            .position
                            .checked_add(region.content.len())
                            .ok_or_else(|| {
                                io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    "RTF editable-region range overflow",
                                )
                            })?,
                        BodyEventKind::EditableRegionEnd(region),
                    )
                },
                BodyStoryEvent::AnnotationStart(index) => {
                    let annotation =
                        take_story_item(annotations, &mut saw_annotation_starts, index)?;
                    (
                        annotation.position,
                        BodyEventKind::AnnotationStart(annotation),
                    )
                },
                BodyStoryEvent::AnnotationEnd(index) => {
                    let annotation = take_story_item(annotations, &mut saw_annotation_ends, index)?;
                    (
                        annotation.range_end,
                        BodyEventKind::AnnotationEnd(annotation),
                    )
                },
                BodyStoryEvent::Note(index) => {
                    let note = take_story_item(notes, &mut saw_notes, index)?;
                    (note.position, BodyEventKind::Note(note))
                },
                BodyStoryEvent::Object(index) => {
                    let object = take_story_item(objects, &mut saw_objects, index)?;
                    (object.position, BodyEventKind::Object(object, pictures))
                },
                BodyStoryEvent::PictureCompatibility(index) => {
                    let record = take_story_item(
                        picture_compatibility_records,
                        &mut saw_picture_records,
                        index,
                    )?;
                    let picture = pictures.get(record.picture_index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF picture-compatibility record references a missing picture",
                        )
                    })?;
                    (
                        record.position,
                        BodyEventKind::PictureCompatibility(record, picture),
                    )
                },
                BodyStoryEvent::FormFieldStart(index) => {
                    let field = take_story_item(form_fields, &mut saw_form_starts, index)?;
                    (field.position, BodyEventKind::FormFieldStart(field))
                },
                BodyStoryEvent::FormFieldEnd(index) => {
                    let field = take_story_item(form_fields, &mut saw_form_ends, index)?;
                    (field.range_end, BodyEventKind::FormFieldEnd)
                },
                BodyStoryEvent::RevisionStart(index) => {
                    let revision = take_story_item(revisions, &mut saw_revision_starts, index)?;
                    if revision.revision_type != RevisionType::Insertion {
                        return Err(invalid_story_reference());
                    }
                    (revision.position, BodyEventKind::RevisionStart(revision))
                },
                BodyStoryEvent::RevisionEnd(index) => {
                    let revision = take_story_item(revisions, &mut saw_revision_ends, index)?;
                    if revision.revision_type != RevisionType::Insertion {
                        return Err(invalid_story_reference());
                    }
                    (revision.range_end, BodyEventKind::RevisionEnd)
                },
                BodyStoryEvent::RevisionDeletion(index) => {
                    let revision = take_story_item(revisions, &mut saw_revision_deletions, index)?;
                    if revision.revision_type != RevisionType::Deletion {
                        return Err(invalid_story_reference());
                    }
                    (revision.position, BodyEventKind::RevisionDeletion(revision))
                },
                BodyStoryEvent::GeneratedListMarker(index) => {
                    let marker =
                        take_story_item(generated_list_markers, &mut saw_generated_markers, index)?;
                    (marker.position, BodyEventKind::GeneratedListMarker(marker))
                },
                BodyStoryEvent::LegacyTextBox(index) => {
                    let text_box =
                        take_story_item(legacy_text_boxes, &mut saw_legacy_text_boxes, index)?;
                    (text_box.position, BodyEventKind::LegacyTextBox(text_box))
                },
                BodyStoryEvent::LegacyDrawing(index) => {
                    let drawing =
                        take_story_item(legacy_drawings, &mut saw_legacy_drawings, index)?;
                    (drawing.position, BodyEventKind::LegacyDrawing(drawing))
                },
                BodyStoryEvent::NavigationEntry(index) => {
                    let entry =
                        take_story_item(navigation_entries, &mut saw_navigation_entries, index)?;
                    (entry.position(), BodyEventKind::NavigationEntry(entry))
                },
            };
            if body.get(position..position).is_none()
                || previous_story_position.is_some_and(|previous| previous > position)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF body story events are outside or out of story order",
                ));
            }
            events.push(BodyEvent {
                offset: position,
                order: 1,
                kind,
            });
            previous_story_position = Some(position);
        }
        let complete = ordered_drawings == drawing_order
            && saw_shapes
                .iter()
                .zip(shapes)
                .all(|(seen, shape)| shape.is_background || *seen)
            && saw_groups.iter().all(|seen| *seen)
            && saw_fields
                .iter()
                .zip(fields)
                .all(|(seen, field)| !matches!(field.owner, FieldOwner::Body) || *seen)
            && saw_bookmark_starts.iter().all(|seen| *seen)
            && saw_bookmark_ends.iter().all(|seen| *seen)
            && saw_custom_xml_opens.iter().all(|seen| *seen)
            && saw_custom_xml_closes.iter().all(|seen| *seen)
            && saw_math_zones.iter().all(|seen| *seen)
            && saw_protection_starts.iter().all(|seen| *seen)
            && saw_protection_ends.iter().all(|seen| *seen)
            && saw_editable_starts.iter().all(|seen| *seen)
            && saw_editable_ends.iter().all(|seen| *seen)
            && saw_annotation_starts.iter().all(|seen| *seen)
            && saw_annotation_ends.iter().all(|seen| *seen)
            && saw_notes.iter().all(|seen| *seen)
            && saw_objects.iter().all(|seen| *seen)
            && saw_picture_records.iter().all(|seen| *seen)
            && saw_form_starts.iter().all(|seen| *seen)
            && saw_form_ends.iter().all(|seen| *seen)
            && revisions
                .iter()
                .zip(&body_revisions)
                .zip(&saw_revision_starts)
                .zip(&saw_revision_ends)
                .zip(&saw_revision_deletions)
                .all(|((((revision, is_body), start), end), deletion)| {
                    match revision.revision_type {
                        _ if !*is_body => true,
                        RevisionType::Insertion => *start && *end && !*deletion,
                        RevisionType::Deletion => *deletion && !*start && !*end,
                        RevisionType::FormatChange
                        | RevisionType::MovedFrom
                        | RevisionType::MovedTo => false,
                    }
                })
            && saw_generated_markers.iter().all(|seen| *seen)
            && saw_legacy_text_boxes.iter().all(|seen| *seen)
            && saw_legacy_drawings.iter().all(|seen| *seen)
            && next_section_index == sections.len()
            && saw_navigation_entries
                .iter()
                .zip(&body_navigation)
                .all(|(seen, is_body)| !*is_body || *seen);
        if !complete {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body story order is incomplete or changes drawing order",
            ));
        }
        for node in opaque_nodes {
            if let crate::opaque::Anchor::Body(offset) = node.anchor() {
                if body.get(offset..offset).is_none() {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF opaque node is outside body text or splits a character",
                    ));
                }
                events.push(BodyEvent {
                    offset,
                    order: 1,
                    kind: BodyEventKind::Opaque(node),
                });
            }
        }
        events.sort_by_key(|event| (event.offset, event.order));

        let mut event_index = 0usize;
        let mut boundary_index = 0usize;
        let mut body_offset = 0usize;
        for block in blocks {
            let block_end = body_offset.checked_add(block.text.len()).ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "RTF body text size overflow")
            })?;
            let mut local_offset = 0usize;
            while let Some(event) = events
                .get(event_index)
                .copied()
                .filter(|event| event.offset <= block_end)
            {
                let event_offset = event.offset;
                if event_offset < body_offset {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF bookmark events are not ordered",
                    ));
                }
                let local_end = event_offset.checked_sub(body_offset).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF body event precedes its text block",
                    )
                })?;
                if local_end > local_offset {
                    self.write_style_block_fragment(
                        block,
                        local_offset,
                        local_end,
                        body_offset,
                        body_boundaries,
                        &mut boundary_index,
                    )?;
                    local_offset = local_end;
                }
                while let Some(event) = events
                    .get(event_index)
                    .copied()
                    .filter(|event| event.offset == event_offset)
                {
                    self.write_body_event(event, fields)?;
                    event_index += 1;
                }
            }
            if local_offset < block.text.len() {
                self.write_style_block_fragment(
                    block,
                    local_offset,
                    block.text.len(),
                    body_offset,
                    body_boundaries,
                    &mut boundary_index,
                )?;
            }
            body_offset = block_end;
        }
        while let Some(event) = events
            .get(event_index)
            .copied()
            .filter(|event| event.offset == body_offset)
        {
            self.write_body_event(event, fields)?;
            event_index += 1;
        }
        if event_index != events.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF bookmark range extends beyond body text",
            ));
        }
        if boundary_index != body_boundaries.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF body text boundaries are incomplete",
            ));
        }
        Ok(())
    }

    pub(in super::super) fn write_body_event(
        &mut self,
        event: BodyEvent<'_, '_>,
        fields: &[Field<'_>],
    ) -> io::Result<()> {
        match event.kind {
            BodyEventKind::Shape(shape) => self.write_root_shape(shape),
            BodyEventKind::ShapeGroup(group) => self.write_shape_group(group, true),
            BodyEventKind::Object(object, pictures) => self.write_object(object, pictures),
            BodyEventKind::PictureCompatibility(record, picture) => {
                self.write_picture_compatibility(record.kind, picture)
            },
            BodyEventKind::GeneratedListMarker(marker) => self.write_generated_list_marker(marker),
            BodyEventKind::LegacyTextBox(text_box) => self.write_legacy_text_box(text_box),
            BodyEventKind::LegacyDrawing(drawing) => self.write_legacy_drawing(drawing),
            BodyEventKind::NavigationEntry(entry) => self.write_navigation_entry(entry),
            BodyEventKind::BookmarkStart(bookmark) => self.write_bookmark_start(bookmark),
            BodyEventKind::BookmarkEnd(bookmark) => self.write_bookmark_end(bookmark.name.as_ref()),
            BodyEventKind::CustomXmlOpen(tag) => self.write_custom_xml_open(tag),
            BodyEventKind::CustomXmlClose(tag) => self.write_custom_xml_close(tag),
            BodyEventKind::MathZone(zone) => self.write_math_zone(zone),
            BodyEventKind::ProtectionRangeStart(range) => {
                self.write_protection_range_marker("protstart", range)
            },
            BodyEventKind::ProtectionRangeEnd(range) => {
                self.write_protection_range_marker("protend", range)
            },
            BodyEventKind::EditableRegionStart(region) => {
                region.validate().map_err(|error| {
                    io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                })?;
                self.write_str("\\ebcstart ")
            },
            BodyEventKind::EditableRegionEnd(region) => {
                region.validate().map_err(|error| {
                    io::Error::new(io::ErrorKind::InvalidInput, error.to_string())
                })?;
                self.write_str("\\ebcend ")
            },
            BodyEventKind::AnnotationStart(annotation) => self.write_annotation_start(annotation),
            BodyEventKind::AnnotationEnd(annotation) => self.write_annotation_end(annotation),
            BodyEventKind::Note(note) => self.write_note_with_fields(note, fields),
            BodyEventKind::RevisionStart(revision) => self.write_revision_start(revision),
            BodyEventKind::RevisionEnd => self.write_str("}"),
            BodyEventKind::RevisionDeletion(revision) => self.write_revision(revision),
            BodyEventKind::FormFieldStart(field) => self.write_form_field_start(field),
            BodyEventKind::FormFieldEnd => self.write_str("}}"),
            BodyEventKind::GenericField(field) => self.write_field_with_fields(field, fields, 0),
            BodyEventKind::PageBreak => self.write_str("\\page "),
            BodyEventKind::SoftBreak(soft_break) => match soft_break.kind {
                SoftBreakKind::Page => self.write_str("\\softpage "),
                SoftBreakKind::Column => self.write_str("\\softcol "),
                SoftBreakKind::Line => self.write_str("\\softline "),
                SoftBreakKind::LineHeight(height) => {
                    self.write_control_word("softlheight", Some(height))?;
                    self.write_str(" ")
                },
            },
            BodyEventKind::ColumnBreak => self.write_str("\\column "),
            BodyEventKind::SectionBreak(section) => {
                self.write_control_word("sect", None)?;
                if let Some(section) = section {
                    self.write_section_with_fields(section, fields)?;
                }
                Ok(())
            },
            BodyEventKind::Opaque(node) => self.writer.write_all(node.source()),
        }
    }

    pub(in super::super) fn write_root_shape(&mut self, shape: &Shape<'_>) -> io::Result<()> {
        shape
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        if !shape.instruction_present {
            self.write_str("{\\shp")?;
            let result = shape.result.as_ref().ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "fallback-only RTF shape has no shape result",
                )
            })?;
            self.write_shape_result(result)?;
            return self.write_str("}");
        }
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape bottom edge overflows",
                )
            })?;
        self.write_str("{\\shp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if !shape
            .properties
            .iter()
            .any(|property| property.name == "shapeType")
        {
            let shape_type = match shape.shape_type {
                ShapeType::Rectangle => Some(1),
                ShapeType::RoundRectangle => Some(2),
                ShapeType::Ellipse => Some(3),
                ShapeType::Arc => Some(19),
                ShapeType::Line => Some(20),
                ShapeType::PictureFrame => Some(75),
                ShapeType::TextBox => Some(202),
                ShapeType::Group => Some(0),
                ShapeType::Custom(value) => Some(value),
                ShapeType::Polygon | ShapeType::Unknown => None,
            };
            if let Some(value) = shape_type {
                self.write_shape_scalar_property("shapeType", &value.to_string())?;
            }
        }
        for property in &shape.properties {
            self.write_shape_property(property)?;
        }
        if shape.text_destination_present
            || !shape.text.is_empty()
            || !shape.text_shapes.is_empty()
            || !shape.text_shape_groups.is_empty()
            || !shape.text_story_events.is_empty()
        {
            self.write_shape_text(shape)?;
        }
        self.write_str("}")?;
        if let Some(result) = &shape.result {
            self.write_shape_result(result)?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_shape_info(&mut self, info: &[ShapeGroupInfo]) -> io::Result<()> {
        for info in info {
            match *info {
                ShapeGroupInfo::ShapeId(value) => self.write_control_word("shplid", Some(value))?,
                ShapeGroupInfo::InHeader(value) => {
                    self.write_control_word("shpfhdr", Some(i32::from(value)))?;
                },
                ShapeGroupInfo::HorizontalPage => self.write_control_word("shpbxpage", None)?,
                ShapeGroupInfo::HorizontalMargin => self.write_control_word("shpbxmargin", None)?,
                ShapeGroupInfo::HorizontalColumn => self.write_control_word("shpbxcolumn", None)?,
                ShapeGroupInfo::IgnoreHorizontal => self.write_control_word("shpbxignore", None)?,
                ShapeGroupInfo::VerticalPage => self.write_control_word("shpbypage", None)?,
                ShapeGroupInfo::VerticalMargin => self.write_control_word("shpbymargin", None)?,
                ShapeGroupInfo::VerticalParagraph => self.write_control_word("shpbypara", None)?,
                ShapeGroupInfo::IgnoreVertical => self.write_control_word("shpbyignore", None)?,
                ShapeGroupInfo::Wrap(value) => self.write_control_word("shpwr", Some(value))?,
                ShapeGroupInfo::WrapSide(value) => {
                    self.write_control_word("shpwrk", Some(value))?;
                },
                ShapeGroupInfo::BelowText(value) => {
                    self.write_control_word("shpfblwtxt", Some(i32::from(value)))?;
                },
                ShapeGroupInfo::LockAnchor => self.write_control_word("shplockanchor", None)?,
            }
        }
        Ok(())
    }

    pub(in super::super) fn write_object(
        &mut self,
        object: &EmbeddedObject<'_>,
        pictures: &[Picture<'_>],
    ) -> io::Result<()> {
        self.write_str("{\\object")?;
        self.write_str(match object.kind {
            ObjectKind::Embedded => "\\objemb",
            ObjectKind::Link => "\\objlink",
            ObjectKind::AutoLink => "\\objautlink",
            ObjectKind::Html => "\\objhtml",
            ObjectKind::Subscriber => "\\objsub",
            ObjectKind::Publisher => "\\objpub",
            ObjectKind::InstallableCommand => "\\objicemb",
            ObjectKind::OleControl => "\\objocx",
            ObjectKind::Unknown => "",
        })?;
        if object.link_self {
            self.write_str("\\linkself")?;
        }
        if object.locked {
            self.write_str("\\objlock")?;
        }
        if object.update_requested {
            self.write_str("\\objupdate")?;
        }
        if !object.class_name.is_empty() {
            self.write_str("{\\*\\objclass ")?;
            self.write_destination_text(object.class_name.as_ref())?;
            self.write_str("}")?;
        }
        if !object.name.is_empty() {
            self.write_str("{\\*\\objname ")?;
            self.write_destination_text(object.name.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(alias) = &object.alias {
            self.write_str("{\\*\\objalias ")?;
            self.write_destination_text(alias.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(section) = &object.section {
            self.write_str("{\\*\\objsect ")?;
            self.write_destination_text(section.as_ref())?;
            self.write_str("}")?;
        }
        if let Some(time) = object.time {
            self.write_str("{\\*\\objtime ")?;
            for (name, value) in [
                ("yr", time.year),
                ("mo", time.month),
                ("dy", time.day),
                ("hr", time.hour),
                ("min", time.minute),
                ("sec", time.second),
            ] {
                if let Some(value) = value {
                    self.write_control_word(name, Some(value))?;
                }
            }
            self.write_str("}")?;
        }
        if object.set_size {
            self.write_str("\\objsetsize")?;
        }
        self.write_optional_object_value("objalign", object.alignment)?;
        self.write_optional_object_value("objtransy", object.translation_y)?;
        if object.height != 0 {
            self.write_control_word("objh", Some(object.height))?;
        }
        if object.width != 0 {
            self.write_control_word("objw", Some(object.width))?;
        }
        self.write_optional_object_value("objcropt", object.crop_top)?;
        self.write_optional_object_value("objcropb", object.crop_bottom)?;
        self.write_optional_object_value("objcropl", object.crop_left)?;
        self.write_optional_object_value("objcropr", object.crop_right)?;
        self.write_optional_object_value("objscalex", object.scale_x)?;
        self.write_optional_object_value("objscaley", object.scale_y)?;
        if object.merge_result {
            self.write_str("\\rsltmerge")?;
        }
        if let Some(kind) = object.result_kind {
            self.write_str(match kind {
                ObjectResultKind::Rtf => "\\rsltrtf",
                ObjectResultKind::Text => "\\rslttxt",
                ObjectResultKind::Picture => "\\rsltpict",
                ObjectResultKind::Bitmap => "\\rsltbmp",
                ObjectResultKind::Html => "\\rslthtml",
            })?;
        }
        if !object.class_id.is_empty() {
            self.write_str("{\\*\\oleclsid ")?;
            self.write_destination_text(object.class_id.as_ref())?;
            self.write_str("}")?;
        }
        self.write_str("{\\*\\objdata ")?;
        for byte in &object.data {
            write!(self.writer, "{byte:02x}")?;
        }
        self.write_str("}")?;
        if !object.result_text.is_empty() || !object.result_picture_indices.is_empty() {
            self.write_str("{\\result ")?;
            self.write_destination_text(object.result_text.as_ref())?;
            for index in &object.result_picture_indices {
                let picture = pictures.get(*index).ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "RTF embedded object references a missing result picture",
                    )
                })?;
                self.write_picture(picture)?;
            }
            self.write_str("\\par}")?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_optional_object_value(
        &mut self,
        control: &str,
        value: Option<i32>,
    ) -> io::Result<()> {
        if let Some(value) = value {
            self.write_control_word(control, Some(value))?;
        }
        Ok(())
    }

    pub(in super::super) fn write_picture_compatibility(
        &mut self,
        kind: PictureCompatibilityKind,
        picture: &Picture<'_>,
    ) -> io::Result<()> {
        self.write_str(match kind {
            PictureCompatibilityKind::ShapePicture => "{\\*\\shppict",
            PictureCompatibilityKind::NonShapePicture => "{\\nonshppict",
        })?;
        self.write_picture(picture)?;
        self.write_str("}")
    }

    pub(in super::super) fn write_picture_shape_properties(
        &mut self,
        properties: &PictureShapeProperties<'_>,
    ) -> io::Result<()> {
        properties
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\picprop")?;
        if let Some(shape_id) = properties.shape_id {
            self.write_control_word("shplid", Some(shape_id))?;
        }
        for property in &properties.properties {
            self.write_shape_property(property)?;
        }
        self.write_str("}")
    }

    /// Write one inert legacy drawing text box.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_legacy_text_box(&mut self, text_box: &LegacyTextBox<'_>) -> io::Result<()> {
        text_box
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;

        self.write_str("{\\*\\do")?;
        if let Some(anchor) = text_box.horizontal_anchor {
            self.write_control_word(
                match anchor {
                    LegacyHorizontalAnchor::Page => "dobxpage",
                    LegacyHorizontalAnchor::Margin => "dobxmargin",
                    LegacyHorizontalAnchor::Column => "dobxcolumn",
                },
                None,
            )?;
        }
        if let Some(anchor) = text_box.vertical_anchor {
            self.write_control_word(
                match anchor {
                    LegacyVerticalAnchor::Page => "dobypage",
                    LegacyVerticalAnchor::Margin => "dobymargin",
                    LegacyVerticalAnchor::Paragraph => "dobypara",
                },
                None,
            )?;
        }
        if let Some(value) = text_box.z_order {
            self.write_control_word("dodhgt", Some(value))?;
        }
        self.write_control_word("dptxbx", None)?;
        if let Some(value) = text_box.margin {
            self.write_control_word("dptxbxmar", Some(value))?;
        }
        self.write_control_word(
            match text_box.direction {
                LegacyTextDirection::LeftToRightTopToBottom => "dptxlrtb",
                LegacyTextDirection::LeftToRightTopToBottomVertical => "dptxlrtbv",
                LegacyTextDirection::TopToBottomRightToLeft => "dptxtbrl",
                LegacyTextDirection::TopToBottomRightToLeftVertical => "dptxtbrlv",
                LegacyTextDirection::BottomToTopLeftToRight => "dptxbtlr",
            },
            None,
        )?;
        if let Some(value) = text_box.x {
            self.write_control_word("dpx", Some(value))?;
        }
        if let Some(value) = text_box.y {
            self.write_control_word("dpy", Some(value))?;
        }
        if let Some(value) = text_box.width {
            self.write_control_word("dpxsize", Some(value))?;
        }
        if let Some(value) = text_box.height {
            self.write_control_word("dpysize", Some(value))?;
        }
        self.write_str("{\\dptxbxtext ")?;
        self.write_field_story(
            text_box.text.as_ref(),
            &text_box.shapes,
            &text_box.shape_groups,
            &text_box.drawing_order,
            &text_box.story_events,
            &[],
            FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}}")
    }

    /// Write one inert Word 6/95 drawing destination canonically.
    ///
    /// # Errors
    /// Returns an error when writing to the underlying output fails.
    pub fn write_legacy_drawing(&mut self, drawing: &LegacyDrawing<'_>) -> io::Result<()> {
        drawing
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\do")?;
        self.write_control_word(
            match drawing.horizontal_anchor {
                LegacyHorizontalAnchor::Page => "dobxpage",
                LegacyHorizontalAnchor::Margin => "dobxmargin",
                LegacyHorizontalAnchor::Column => "dobxcolumn",
            },
            None,
        )?;
        self.write_control_word(
            match drawing.vertical_anchor {
                LegacyVerticalAnchor::Page => "dobypage",
                LegacyVerticalAnchor::Margin => "dobymargin",
                LegacyVerticalAnchor::Paragraph => "dobypara",
            },
            None,
        )?;
        self.write_control_word("dodhgt", Some(drawing.z_order))?;
        if drawing.locked {
            self.write_control_word("dolock", None)?;
        }
        self.write_legacy_drawing_primitive(&drawing.primitive)?;
        self.write_str("}")
    }

    pub(in super::super) fn write_legacy_geometry(
        &mut self,
        geometry: LegacyDrawingGeometry,
    ) -> io::Result<()> {
        self.write_control_word("dpx", Some(geometry.x))?;
        self.write_control_word("dpy", Some(geometry.y))?;
        self.write_control_word("dpxsize", Some(geometry.width))?;
        self.write_control_word("dpysize", Some(geometry.height))
    }

    pub(in super::super) fn write_legacy_point(
        &mut self,
        point: LegacyDrawingPoint,
    ) -> io::Result<()> {
        self.write_control_word("dpptx", Some(point.x))?;
        self.write_control_word("dppty", Some(point.y))
    }

    pub(in super::super) fn write_legacy_drawing_primitive(
        &mut self,
        primitive: &LegacyDrawingPrimitive<'_>,
    ) -> io::Result<()> {
        match primitive {
            LegacyDrawingPrimitive::Group {
                geometry,
                children,
                end_geometry,
            } => {
                self.write_control_word("dpgroup", None)?;
                let count = i32::try_from(children.len()).map_err(|_err| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "legacy drawing child count overflow",
                    )
                })?;
                self.write_control_word("dpcount", Some(count))?;
                self.write_legacy_geometry(*geometry)?;
                for child in children {
                    self.write_legacy_drawing_primitive(child)?;
                }
                self.write_control_word("dpendgroup", None)?;
                self.write_legacy_geometry(*end_geometry)
            },
            LegacyDrawingPrimitive::Callout(callout) => {
                self.write_control_word("dpcallout", None)?;
                self.write_control_word(
                    match callout.callout_type {
                        LegacyCalloutType::RightAngle => "dpcotright",
                        LegacyCalloutType::Single => "dpcotsingle",
                        LegacyCalloutType::Double => "dpcotdouble",
                        LegacyCalloutType::Triple => "dpcottriple",
                    },
                    None,
                )?;
                if let Some(angle) = callout.angle {
                    self.write_control_word("dpcoa", Some(i32::from(angle)))?;
                }
                if callout.accent {
                    self.write_control_word("dpcoaccent", None)?;
                }
                if callout.smart_attach {
                    self.write_control_word("dpcosmarta", None)?;
                }
                if callout.best_fit {
                    self.write_control_word("dpcobestfit", None)?;
                }
                if callout.minus_x {
                    self.write_control_word("dpcominusx", None)?;
                }
                if callout.minus_y {
                    self.write_control_word("dpcominusy", None)?;
                }
                if callout.border {
                    self.write_control_word("dpcoborder", None)?;
                }
                if let Some(attachment) = callout.attachment {
                    self.write_control_word(
                        match attachment {
                            LegacyCalloutAttachment::Top => "dpcodtop",
                            LegacyCalloutAttachment::Center => "dpcodcenter",
                            LegacyCalloutAttachment::Bottom => "dpcodbottom",
                            LegacyCalloutAttachment::Absolute => "dpcodabs",
                        },
                        None,
                    )?;
                }
                if let Some(value) = callout.descent {
                    self.write_control_word("dpcodescent", Some(value))?;
                }
                self.write_control_word("dpcooffset", Some(callout.offset))?;
                self.write_control_word("dpcolength", Some(callout.length))?;
                self.write_legacy_geometry(callout.geometry)?;
                self.write_legacy_drawing_primitive(&callout.polyline)?;
                self.write_legacy_drawing_primitive(&callout.text_box)?;
                self.write_legacy_properties(callout.properties)
            },
            LegacyDrawingPrimitive::Line {
                start,
                end,
                geometry,
                properties,
            } => {
                self.write_control_word("dpline", None)?;
                self.write_legacy_point(*start)?;
                self.write_legacy_point(*end)?;
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            LegacyDrawingPrimitive::Rectangle {
                rounded,
                geometry,
                properties,
            } => {
                self.write_control_word("dprect", None)?;
                if *rounded {
                    self.write_control_word("dproundr", None)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            LegacyDrawingPrimitive::TextBox {
                text_box,
                properties,
            } => {
                self.write_control_word("dptxbx", None)?;
                if let Some(value) = text_box.margin {
                    self.write_control_word("dptxbxmar", Some(value))?;
                }
                if text_box.direction != LegacyTextDirection::LeftToRightTopToBottom {
                    self.write_control_word(
                        match text_box.direction {
                            LegacyTextDirection::LeftToRightTopToBottom => "dptxlrtb",
                            LegacyTextDirection::LeftToRightTopToBottomVertical => "dptxlrtbv",
                            LegacyTextDirection::TopToBottomRightToLeft => "dptxtbrl",
                            LegacyTextDirection::TopToBottomRightToLeftVertical => "dptxtbrlv",
                            LegacyTextDirection::BottomToTopLeftToRight => "dptxbtlr",
                        },
                        None,
                    )?;
                }
                self.write_legacy_text_box_text(text_box)?;
                self.write_legacy_geometry(LegacyDrawingGeometry {
                    x: text_box.x.unwrap_or(0),
                    y: text_box.y.unwrap_or(0),
                    width: text_box.width.unwrap_or(0),
                    height: text_box.height.unwrap_or(0),
                })?;
                self.write_legacy_properties(*properties)
            },
            LegacyDrawingPrimitive::Ellipse {
                geometry,
                properties,
            } => {
                self.write_control_word("dpellipse", None)?;
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            LegacyDrawingPrimitive::Polyline {
                closed,
                points,
                geometry,
                properties,
            } => {
                self.write_control_word("dppolyline", None)?;
                if *closed {
                    self.write_control_word("dppolygon", None)?;
                }
                let count = i32::try_from(points.len()).map_err(|_err| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "legacy drawing point count overflow",
                    )
                })?;
                self.write_control_word("dppolycount", Some(count))?;
                for point in points {
                    self.write_legacy_point(*point)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
            LegacyDrawingPrimitive::Arc {
                flip_x,
                flip_y,
                geometry,
                properties,
            } => {
                self.write_control_word("dparc", None)?;
                if *flip_x {
                    self.write_control_word("dparcflipx", None)?;
                }
                if *flip_y {
                    self.write_control_word("dparcflipy", None)?;
                }
                self.write_legacy_geometry(*geometry)?;
                self.write_legacy_properties(*properties)
            },
        }
    }

    pub(in super::super) fn write_legacy_text_box_text(
        &mut self,
        text_box: &LegacyTextBox<'_>,
    ) -> io::Result<()> {
        self.write_str("{\\dptxbxtext ")?;
        self.write_field_story(
            text_box.text.as_ref(),
            &text_box.shapes,
            &text_box.shape_groups,
            &text_box.drawing_order,
            &text_box.story_events,
            &[],
            FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}")
    }

    pub(in super::super) fn write_legacy_color(
        &mut self,
        gray: &str,
        red_control: &str,
        green_control: &str,
        blue_control: &str,
        palette_control: &str,
        color: LegacyDrawingColor,
    ) -> io::Result<()> {
        match color {
            LegacyDrawingColor::Gray(value) => {
                self.write_control_word(gray, Some(i32::from(value)))
            },
            LegacyDrawingColor::Rgb {
                red,
                green,
                blue,
                palette,
            } => {
                self.write_control_word(red_control, Some(i32::from(red)))?;
                self.write_control_word(green_control, Some(i32::from(green)))?;
                self.write_control_word(blue_control, Some(i32::from(blue)))?;
                if palette {
                    self.write_control_word(palette_control, None)?;
                }
                Ok(())
            },
        }
    }

    pub(in super::super) fn write_legacy_arrow(
        &mut self,
        prefix: &str,
        arrow: LegacyDrawingArrow,
    ) -> io::Result<()> {
        self.write_control_word(
            &format!(
                "{prefix}{}",
                match arrow.fill {
                    LegacyDrawingArrowFill::Solid => "sol",
                    LegacyDrawingArrowFill::Hollow => "hol",
                }
            ),
            None,
        )?;
        self.write_control_word(&format!("{prefix}l"), Some(arrow.length as i32))?;
        self.write_control_word(&format!("{prefix}w"), Some(arrow.width as i32))
    }

    pub(in super::super) fn write_legacy_properties(
        &mut self,
        properties: LegacyDrawingProperties,
    ) -> io::Result<()> {
        if let Some(line) = properties.line {
            self.write_control_word(
                match line.style {
                    LegacyDrawingLineStyle::Solid => "dplinesolid",
                    LegacyDrawingLineStyle::Hollow => "dplinehollow",
                    LegacyDrawingLineStyle::Dashed => "dplinedash",
                    LegacyDrawingLineStyle::Dotted => "dplinedot",
                    LegacyDrawingLineStyle::DashDot => "dplinedado",
                    LegacyDrawingLineStyle::DashDotDot => "dplinedadodo",
                },
                None,
            )?;
            self.write_legacy_color(
                "dplinegray",
                "dplinecor",
                "dplinecog",
                "dplinecob",
                "dplinepal",
                line.color,
            )?;
            self.write_control_word("dplinew", Some(line.width))?;
        }
        if let Some(fill) = properties.fill {
            self.write_legacy_color(
                "dpfillfggray",
                "dpfillfgcr",
                "dpfillfgcg",
                "dpfillfgcb",
                "dpfillfgpal",
                fill.foreground,
            )?;
            self.write_legacy_color(
                "dpfillbggray",
                "dpfillbgcr",
                "dpfillbgcg",
                "dpfillbgcb",
                "dpfillbgpal",
                fill.background,
            )?;
            self.write_control_word("dpfillpat", Some(fill.pattern as i32))?;
        }
        if let Some(arrow) = properties.start_arrow {
            self.write_legacy_arrow("dpastart", arrow)?;
        }
        if let Some(arrow) = properties.end_arrow {
            self.write_legacy_arrow("dpaend", arrow)?;
        }
        if let Some(shadow) = properties.shadow {
            self.write_control_word("dpshadow", None)?;
            self.write_control_word("dpshadx", Some(shadow.x_offset))?;
            self.write_control_word("dpshady", Some(shadow.y_offset))?;
        }
        Ok(())
    }
}

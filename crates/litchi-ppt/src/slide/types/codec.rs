//! Binary record and `OfficeArt` decoding used by the slide facade.

use super::Slide;
use crate::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::odraw::{FrameKind, ShapeExt as _};
use crate::package::{Error, Result};
use crate::shapes::ShapeEnum;

impl Slide<'_> {
    pub(super) fn parse_animations(&self) -> Result<Vec<ShapeAnimation>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let shapes = crate::odraw::parse(
            &ppdrawing.data,
            &crate::odraw::limits_for_source(self.source_limits),
        )?;
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
    pub(super) fn parse_animation_extension(&self) -> Result<Option<SlideAnimationExtension>> {
        for prog_tags in self.record.find_children(RecordType::ProgTags) {
            for prog_binary_tag in prog_tags.find_children(RecordType::ProgBinaryTag) {
                let Some(tag_name) = prog_binary_tag.find_child(RecordType::CString) else {
                    continue;
                };
                if !super::validation::is_ppt10_tag_name(tag_name) {
                    continue;
                }
                let data = prog_binary_tag
                    .find_child(RecordType::BinaryTagData)
                    .ok_or_else(|| {
                        Error::Corrupted(
                            "___PPT10 programmable tag is missing BinaryTagData".to_string(),
                        )
                    })?;
                return crate::animation::parse_slide_animation_extension(&data.data).map(Some);
            }
        }
        Ok(None)
    }

    /// Parse this slide's shapes, returning them with the non-structural
    /// defects repaired along the way, each located by this slide's persist
    /// ID.
    pub(super) fn parse_shapes(
        &self,
    ) -> Result<(Vec<ShapeEnum<'static>>, crate::leniency::ToleranceReport)> {
        let mut report = crate::leniency::ToleranceReport::default();
        // Find PPDrawing record
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok((Vec::new(), report));
        };

        let shapes = Self::parse_shape_enums(
            &ppdrawing.data,
            self.outline_text_bodies,
            self.source_limits,
            &mut report,
        )?;
        report.locate_from(
            0,
            crate::leniency::DefectLocation {
                persist_id: Some(self.persist_id),
                shape_id: None,
            },
        );
        Ok((shapes, report))
    }

    /// Parse the source-order shape values from one complete PPDrawing
    /// payload. Source-backed selector paths use this helper after reading
    /// exactly one selected Slide record, avoiding a full Document/NotesIndex
    /// materialization while retaining the ordinary shape ordering rules.
    ///
    /// `outline` holds the slide's slide-list text bodies, which supply the
    /// text of shapes that reference them (L-PPT-R1). Resolution only fills
    /// in text: which shapes are returned, and in which order, never depends
    /// on `outline`, so a caller that passes no bodies addresses the same
    /// shapes.
    pub(crate) fn parse_shape_enums(
        ppdrawing_data: &[u8],
        outline: &[crate::OutlineTextBody],
        source_limits: crate::SourceLimits,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Vec<ShapeEnum<'static>>> {
        // Extract Escher shapes from PPDrawing data
        let escher_shapes = crate::odraw::parse(
            ppdrawing_data,
            &crate::odraw::limits_for_source(source_limits),
        )?;

        // Convert Escher shapes to ShapeEnum with full property extraction
        let mut shapes = Vec::with_capacity(escher_shapes.len());
        for (index, escher_shape) in escher_shapes.iter().enumerate() {
            if let Some(mut shape) = Self::convert_odraw_to_shape_enum(
                escher_shape,
                index,
                outline,
                source_limits,
                report,
            )? {
                shape.mark_source_bound_recursive();
                shapes.push(shape);
            }
        }

        Ok(shapes)
    }

    /// Convert an `EscherShape` to `ShapeEnum` with full property extraction.
    ///
    /// # Performance
    ///
    /// - Direct property access (no allocations)
    /// - Pattern matching for type dispatch
    ///
    /// `index` is the shape's zero-based position among the drawing's
    /// top-level shapes. A shape whose content cannot be read is dropped
    /// (`Ok(None)`) and recorded as [`crate::ShapeDefect::UnreadableShape`];
    /// see [`Self::convert_odraw_shape`]. `outline` holds the slide-list
    /// text bodies outline references select; notes and master drawings
    /// pass none.
    pub(crate) fn convert_odraw_to_shape_enum(
        odraw_shape: &litchi_odraw::shape::Shape<'_>,
        index: usize,
        outline: &[crate::OutlineTextBody],
        source_limits: crate::SourceLimits,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Option<ShapeEnum<'static>>> {
        Self::convert_odraw_shape(odraw_shape, 0, index, outline, source_limits, report)
    }

    /// Convert one shape and, for a group, each of its children, isolating
    /// failures per shape (finding F5.11).
    ///
    /// `PowerPoint` and LibreOffice render the remaining shapes of a slide
    /// when one is unreadable, so a content-level failure converting one
    /// shape drops that shape alone and records
    /// [`crate::ShapeDefect::UnreadableShape`] (context: `index`, the
    /// shape's zero-based position among its siblings), discarding repairs
    /// recorded for it. The shape-tree nesting bound, resource limits, and
    /// broken framing of the shape's host records stay fatal.
    fn convert_odraw_shape(
        escher_shape: &litchi_odraw::shape::Shape<'_>,
        depth: usize,
        index: usize,
        outline: &[crate::OutlineTextBody],
        source_limits: crate::SourceLimits,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Option<ShapeEnum<'static>>> {
        if depth >= super::validation::MAX_SHAPE_DEPTH {
            return Err(Error::Corrupted(
                "OfficeArt shape tree exceeds the PPT nesting limit".to_string(),
            ));
        }

        let checkpoint = report.checkpoint();
        let mut converted =
            match Self::convert_odraw_shape_node(escher_shape, outline, source_limits, report) {
                Ok(converted) => converted,
                Err(error) => {
                    report.rollback(checkpoint);
                    if !super::validation::is_isolatable_shape_error(&error)
                        || crate::odraw::validate_host_framing(escher_shape).is_err()
                    {
                        return Err(error);
                    }
                    report.record_at(
                        crate::leniency::ShapeDefect::UnreadableShape,
                        u32::try_from(index).unwrap_or(u32::MAX),
                        crate::leniency::DefectLocation {
                            persist_id: None,
                            shape_id: Some(escher_shape.id()),
                        },
                    );
                    return Ok(None);
                },
            };

        // The shape keeps its stable position among its siblings — `index`
        // counts every sibling, read or not — and its name and alternative
        // text (L-PPT-R4, L-PPT-R3).
        if let Some(shape) = &mut converted {
            let (name, description) = crate::odraw::shape_labels(escher_shape);
            shape.apply_source_identity(crate::shapes::shape::SourceIdentity {
                index,
                name,
                description,
            });
        }

        // Children are converted (and isolated) one by one, outside the
        // group's own isolation, so a fatal error below stays fatal.
        if let Some(ShapeEnum::Group(group)) = &mut converted
            && escher_shape.kind() == litchi_odraw::shape::Kind::Group
        {
            // This follows Apache POI's approach: iterate child shapes and
            // convert them.
            for (child_index, child_escher) in escher_shape.children().iter().enumerate() {
                if let Some(child_shape) = Self::convert_odraw_shape(
                    child_escher,
                    depth + 1,
                    child_index,
                    outline,
                    source_limits,
                    report,
                )? {
                    group.add_decoded_child(child_shape);
                }
            }
        }
        Ok(converted)
    }

    /// The text a placeholder, auto shape, or table cell shows: the inline
    /// text of its `ClientTextbox`, or the slide-list body its outline
    /// reference selects (L-PPT-R1).
    ///
    /// A reference that selects no readable body leaves the text empty, as
    /// POI and LibreOffice leave it, and is recorded as
    /// [`crate::ShapeDefect::UnresolvedOutlineTextRef`] at the shape.
    /// Resolution never drops the shape, so it cannot change which shapes a
    /// slide returns.
    fn shape_text(
        escher_shape: &litchi_odraw::shape::Shape<'_>,
        outline: &[crate::OutlineTextBody],
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Option<String>> {
        use crate::odraw::TextRecords;

        let Some(textbox) = crate::odraw::textbox(escher_shape)? else {
            return Ok(None);
        };
        let max_records = escher_shape.container().record().limits().max_records;
        let max_depth = usize::from(escher_shape.container().record().limits().max_depth);
        let unresolved =
            match crate::odraw::textbox_text_records(textbox.data(), outline, max_records)? {
                TextRecords::Inline(payload) => {
                    return crate::odraw::text_from_textbox_payload_bounded(
                        payload,
                        max_records,
                        max_depth,
                    );
                },
                TextRecords::Outline { index, body } => {
                    match crate::odraw::text_from_textbox_payload_bounded(
                        body.records(),
                        max_records,
                        max_depth,
                    ) {
                        Ok(text) => return Ok(text),
                        Err(error) if super::validation::is_isolatable_shape_error(&error) => index,
                        Err(error) => return Err(error),
                    }
                },
                TextRecords::Unresolved(context) => context,
            };
        report.record_at(
            crate::leniency::ShapeDefect::UnresolvedOutlineTextRef,
            unresolved,
            crate::leniency::DefectLocation {
                persist_id: None,
                shape_id: Some(escher_shape.id()),
            },
        );
        Ok(None)
    }

    /// Convert one shape without its group children.
    fn convert_odraw_shape_node(
        escher_shape: &litchi_odraw::shape::Shape<'_>,
        outline: &[crate::OutlineTextBody],
        source_limits: crate::SourceLimits,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Option<ShapeEnum<'static>>> {
        use crate::shapes::{
            AutoShape, PictureFrameKind, PictureShape, Placeholder, PlaceholderSize,
            PlaceholderType, ShapeEnum, ShapeType, TextBox, shape, shape_enum,
        };
        use crate::slide_extension::HeaderFooterPlaceholder;
        use litchi_odraw::prop::Id;
        use litchi_odraw::shape::Kind;

        let shape_id = escher_shape.id();
        let anchor = crate::odraw::anchor(escher_shape)?;
        let metadata_start = report.defects().len();
        let powerpoint12_shape_metadata =
            crate::odraw::powerpoint12_shape_metadata(escher_shape, Some(report))?;
        report.locate_from(
            metadata_start,
            crate::leniency::DefectLocation {
                persist_id: None,
                shape_id: Some(shape_id),
            },
        );

        if let Some(placeholder_info) = escher_shape.placeholder()? {
            let mut properties = shape::ShapeProperties {
                id: shape_id,
                shape_type: ShapeType::Placeholder,
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
                Self::shape_text(escher_shape, outline, report)?,
            ))));
        }

        if let Some(header_footer) =
            powerpoint12_shape_metadata.and_then(|metadata| metadata.header_footer)
        {
            let placeholder_type = match header_footer {
                HeaderFooterPlaceholder::Date => PlaceholderType::DateAndTime,
                HeaderFooterPlaceholder::SlideNumber => PlaceholderType::SlideNumber,
                HeaderFooterPlaceholder::Footer => PlaceholderType::Footer,
                HeaderFooterPlaceholder::Header => PlaceholderType::Header,
            };
            let mut properties = shape::ShapeProperties {
                id: shape_id,
                shape_type: ShapeType::Placeholder,
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
                Self::shape_text(escher_shape, outline, report)?,
            ))));
        }

        match escher_shape.kind() {
            Kind::TextBox => {
                // Create TextBox with proper properties
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: ShapeType::TextBox,
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
                    outline,
                    source_limits,
                    report,
                )?)))
            },

            Kind::Picture => {
                // Create PictureShape
                let mut picture = PictureShape::new(shape_id);

                picture.set_decoded_frame_kind(match escher_shape.frame_kind()? {
                    FrameKind::Object => PictureFrameKind::OleObject,
                    FrameKind::Media => PictureFrameKind::Media,
                    FrameKind::Picture => PictureFrameKind::Picture,
                });

                if let Some(external_object_id) = escher_shape.external_object_id()? {
                    picture.set_decoded_external_object_id(external_object_id);
                }

                if let Some(a) = anchor {
                    picture.set_decoded_bounds(a.left(), a.top(), a.width(), a.height());
                }
                picture.set_powerpoint12_shape_metadata(powerpoint12_shape_metadata);

                // Extract the one-based BLIP store index from the pib property.
                if let Some(blip_id) = escher_shape.props().get_int(Id::BlipToDisplay) {
                    let blip_index = u32::try_from(blip_id).map_err(|_err| {
                        litchi_odraw::Error::MalformedProperties {
                            reason: "BlipToDisplay must be a positive one-based image identifier",
                        }
                    })?;
                    picture.set_decoded_blip_index(blip_index)?;
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
                    if let Some(width) = escher_shape.props().get_int(Id::LineWidth) {
                        line.set_decoded_width(width);
                    }
                    if let Some(color) = escher_shape.props().get_color(Id::LineColor) {
                        line.set_decoded_color(color.raw());
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
                    group.set_decoded_bounds(a.left(), a.top(), a.width(), a.height());
                }
                group.set_powerpoint12_shape_metadata(powerpoint12_shape_metadata);

                // Child shapes are converted by `convert_odraw_shape`.
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
                    if let Some(cell_anchor) = crate::odraw::anchor(child)?
                        && cell_anchor.width() > 0
                        && cell_anchor.height() > 0
                    {
                        cells.push((child, cell_anchor));
                    }
                }

                let columns: BTreeSet<i32> = cells
                    .iter()
                    .map(|(_, cell_anchor)| cell_anchor.left())
                    .collect();
                let rows: BTreeSet<i32> = cells
                    .iter()
                    .map(|(_, cell_anchor)| cell_anchor.top())
                    .collect();
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
                    table.set_cell_id(row, column, cell.id());
                    table.set_cell_text(
                        row,
                        column,
                        Self::shape_text(cell, outline, report)?.unwrap_or_default(),
                    );
                }

                table.set_grid(row_positions, column_positions);
                Ok(Some(ShapeEnum::Table(table)))
            },

            Kind::Rectangle | Kind::Ellipse | Kind::Callout | Kind::Polygon | Kind::AutoShape => {
                // Create AutoShape
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: ShapeType::AutoShape,
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
                if let Some(text) =
                    Self::shape_text(escher_shape, outline, report)?.filter(|text| !text.is_empty())
                {
                    autoshape.set_decoded_text(text);
                }
                Ok(Some(ShapeEnum::AutoShape(autoshape)))
            },

            // Unknown or unsupported shape types
            Kind::Unknown => Ok(None),
        }
    }

    /// Extract all text from slide and its shapes.
    pub(super) fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // 1. Extract text from direct slide records (TextCharsAtom, etc.)
        // Note: record.extract_text() already recursively processes all children
        let record_text = self.record.extract_text()?;
        let record_trimmed = record_text.trim();
        if !record_trimmed.is_empty() {
            text_parts.push(record_trimmed.to_string());
        }

        // 2. Extract text from Escher/PPDrawing (shapes, text boxes), with
        // outline references resolved into the slide-list text they select
        // (L-PPT-R1). This is separate from regular record text extraction.
        if let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) {
            let escher_text = crate::odraw::text_from_slide_drawing(
                &ppdrawing.data,
                self.outline_text_bodies,
                self.source_limits,
            )?;
            let escher_trimmed = escher_text.trim();
            if !escher_trimmed.is_empty() {
                text_parts.push(escher_trimmed.to_string());
            }
        }

        Ok(if text_parts.is_empty() {
            String::new()
        } else {
            text_parts.join("\n")
        })
    }
}

//! Binary record and OfficeArt decoding used by the slide facade.

use super::Slide;
use crate::animation::{ShapeAnimation, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::odraw::{FrameKind, ShapeExt as _};
use crate::package::{Error, Result};
use crate::shapes::ShapeEnum;

impl<'doc> Slide<'doc> {
    pub(super) fn parse_animations(&self) -> Result<Vec<ShapeAnimation>> {
        let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) else {
            return Ok(Vec::new());
        };
        let shapes = crate::odraw::parse(&ppdrawing.data)?;
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

    pub(super) fn parse_shapes(&self) -> Result<Vec<ShapeEnum<'static>>> {
        // Find PPDrawing record
        let ppdrawing = match self.record.find_child(RecordType::PPDrawing) {
            Some(record) => record,
            None => return Ok(Vec::new()),
        };

        // Extract Escher shapes from PPDrawing data
        let escher_shapes = crate::odraw::parse(&ppdrawing.data)?;

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
        if depth >= super::validation::MAX_SHAPE_DEPTH {
            return Err(Error::Corrupted(
                "OfficeArt shape tree exceeds the PPT nesting limit".to_string(),
            ));
        }

        use crate::shapes::*;
        use crate::slide_extension::HeaderFooterPlaceholder;
        use litchi_odraw::shape::Kind;

        let shape_id = escher_shape.id();
        let anchor = crate::odraw::anchor(escher_shape)?;
        let powerpoint12_shape_metadata = escher_shape.powerpoint12_shape_metadata()?;

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
                escher_shape.text()?,
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
                escher_shape.text()?,
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
                )?)))
            },

            Kind::Picture => {
                // Create PictureShape
                let mut picture = PictureShape::new(shape_id);

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
                if let Some(blip_id) = escher_shape.props().get_int(Id::BlipToDisplay) {
                    let blip_id = u32::try_from(blip_id).map_err(|_| {
                        litchi_odraw::Error::MalformedProperties {
                            reason: "BlipToDisplay must be a positive one-based image identifier",
                        }
                    })?;
                    picture.set_blip_index(blip_id)?;
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
                    if let Some(anchor) = crate::odraw::anchor(child)?
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
    pub(super) fn extract_all_text(&self) -> Result<String> {
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
        if let Some(ppdrawing) = self.record.find_child(RecordType::PPDrawing) {
            let escher_text = crate::odraw::text_from_drawing(&ppdrawing.data)?;
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
}

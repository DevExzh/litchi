//! RTF drawing and shape output.

use super::super::{
    DrawingStoryTextMode, Field, FieldOwner, RtfWriter, Shape, ShapeGroup, ShapeGroupChild,
    ShapeGroupInfo, ShapeProperty, ShapeResult, ShapeThemeColor, ShapeType, StoryDrawing,
    StoryEvent, Write, field, invalid_story_reference, io,
};

impl<W: Write> RtfWriter<W> {
    /// Write the unique inert document-background shape destination.
    pub(in super::super) fn write_document_background(
        &mut self,
        background: Option<&Shape<'_>>,
    ) -> io::Result<()> {
        let Some(shape) = background else {
            return Ok(());
        };
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF background shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF background shape bottom edge overflows",
                )
            })?;
        if shape.properties.len() > 65_536 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF background shape property count exceeds the safety limit",
            ));
        }
        self.write_str("{\\*")?;
        self.write_control_word("background", None)?;
        self.write_str("{")?;
        self.write_control_word("shp", None)?;
        self.write_str("{\\*")?;
        self.write_control_word("shpinst", None)?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if shape.behind_doc
            && !shape
                .info
                .iter()
                .any(|info| matches!(info, ShapeGroupInfo::BelowText(_)))
        {
            self.write_control_word("shpfblwtxt", Some(1))?;
        }
        if shape.locked
            && !shape
                .info
                .iter()
                .any(|info| matches!(info, ShapeGroupInfo::LockAnchor))
        {
            self.write_control_word("shplockanchor", None)?;
        }
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
        for property in &shape.properties {
            if property.name == "shapeType" || property.name == "fBackground" {
                continue;
            }
            self.write_shape_property(property)?;
        }
        self.write_shape_scalar_property("fBackground", "1")?;
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
        self.write_str("}}")
    }

    pub(in super::super) fn write_shape_text(&mut self, shape: &Shape<'_>) -> io::Result<()> {
        self.write_str("{\\shptxt ")?;
        if let Some(background_color) = shape
            .text_formatting
            .and_then(|formatting| formatting.background_color)
        {
            self.write_control_word("cb", Some(i32::from(background_color)))?;
        }
        self.write_field_story(
            shape.text.as_ref(),
            &shape.text_shapes,
            &shape.text_shape_groups,
            &shape.text_drawing_order,
            &shape.text_story_events,
            &[],
            FieldOwner::Other,
            DrawingStoryTextMode::ShapeText,
            0,
        )?;
        self.write_str("}")
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "the writer threads the full drawing context through the pipeline"
    )]
    pub(in super::super) fn write_field_story(
        &mut self,
        text: &str,
        shapes: &[Shape<'_>],
        shape_groups: &[ShapeGroup<'_>],
        drawing_order: &[StoryDrawing],
        story_events: &[StoryEvent],
        fields: &[Field<'_>],
        owner: FieldOwner,
        mode: DrawingStoryTextMode,
        depth: usize,
    ) -> io::Result<()> {
        field::validate_story_events(
            text,
            shapes,
            shape_groups,
            drawing_order,
            story_events,
            "generic-field story",
        )
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        let mut start = 0usize;
        for event in story_events {
            let offset = match *event {
                StoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    shapes
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                StoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    shape_groups
                        .get(index)
                        .ok_or_else(invalid_story_reference)?
                        .position
                },
                StoryEvent::Field(field) => field.position,
                StoryEvent::PageBreak(page_break) => page_break.position,
            };
            let fragment = text.get(start..offset).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF story event splits or leaves its story text",
                )
            })?;
            self.write_drawing_story_fragment(fragment, mode)?;
            match *event {
                StoryEvent::Drawing(StoryDrawing::Shape(index)) => {
                    let shape = shapes.get(index).ok_or_else(invalid_story_reference)?;
                    self.write_root_shape(shape)?;
                },
                StoryEvent::Drawing(StoryDrawing::ShapeGroup(index)) => {
                    let group = shape_groups
                        .get(index)
                        .ok_or_else(invalid_story_reference)?;
                    self.write_shape_group(group, true)?;
                },
                StoryEvent::Field(reference) => {
                    let field = fields
                        .get(reference.field_index)
                        .filter(|field| {
                            field.owner == owner
                                && field.position == reference.position
                                && field.range_end == reference.position
                        })
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidInput,
                                "RTF story has an invalid generic-field owner or reference",
                            )
                        })?;
                    let nested_depth = depth.checked_add(1).ok_or_else(|| {
                        io::Error::new(io::ErrorKind::InvalidInput, "RTF field depth overflow")
                    })?;
                    self.write_field_with_fields(field, fields, nested_depth)?;
                },
                StoryEvent::PageBreak(_) => self.write_str("\\page ")?,
            }
            start = offset;
        }
        let remainder = text.get(start..).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF story event leaves its story text",
            )
        })?;
        self.write_drawing_story_fragment(remainder, mode)
    }

    pub(in super::super) fn write_drawing_story_fragment(
        &mut self,
        value: &str,
        mode: DrawingStoryTextMode,
    ) -> io::Result<()> {
        if matches!(mode, DrawingStoryTextMode::Destination) {
            return self.write_destination_text(value);
        }
        if matches!(mode, DrawingStoryTextMode::Note) {
            return self.write_text(value);
        }
        let mut start = 0usize;
        for (index, character) in value.char_indices() {
            if character != '\n' && character != '\t' {
                continue;
            }
            let fragment = value.get(start..index).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF drawing-story text boundary is invalid",
                )
            })?;
            self.write_destination_text(fragment)?;
            self.write_str(if character == '\n' {
                "\\par "
            } else {
                "\\tab "
            })?;
            start = index + character.len_utf8();
        }
        let remainder = value.get(start..).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF drawing-story text boundary is invalid",
            )
        })?;
        self.write_destination_text(remainder)?;
        Ok(())
    }

    pub(in super::super) fn write_shape_result(
        &mut self,
        result: &ShapeResult<'_>,
    ) -> io::Result<()> {
        result
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        self.write_str("{\\*\\shprslt")?;
        self.write_legacy_drawing(&result.drawing)?;
        self.write_str("}")
    }

    pub(in super::super) fn write_shape_group(
        &mut self,
        group: &ShapeGroup<'_>,
        root: bool,
    ) -> io::Result<()> {
        if root {
            group
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
        } else if group.result.is_some() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF nested shape group cannot contain shprslt",
            ));
        }
        let right = group
            .geometry
            .x
            .checked_add(group.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape group right edge overflows",
                )
            })?;
        let bottom = group
            .geometry
            .y
            .checked_add(group.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF shape group bottom edge overflows",
                )
            })?;
        self.write_str("{\\shpgrp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(group.geometry.x))?;
        self.write_control_word("shptop", Some(group.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(group.geometry.z_order))?;
        for info in &group.info {
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
        for property in &group.properties {
            self.write_shape_property(property)?;
        }
        for child in &group.child_order {
            match *child {
                ShapeGroupChild::Shape(index) => {
                    let shape = group.shapes.get(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF shape group references a missing child shape",
                        )
                    })?;
                    self.write_grouped_shape(shape)?;
                },
                ShapeGroupChild::Group(index) => {
                    let child_group = group.groups.get(index).ok_or_else(|| {
                        io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "RTF shape group references a missing child group",
                        )
                    })?;
                    self.write_shape_group(child_group, false)?;
                },
            }
        }
        self.write_str("}")?;
        if let Some(result) = &group.result {
            self.write_shape_result(result)?;
        }
        self.write_str("}")
    }

    pub(in super::super) fn write_grouped_shape(&mut self, shape: &Shape<'_>) -> io::Result<()> {
        if shape.result.is_some() || !shape.instruction_present {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "RTF grouped shape cannot contain shprslt",
            ));
        }
        let right = shape
            .geometry
            .x
            .checked_add(shape.geometry.width)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF grouped shape right edge overflows",
                )
            })?;
        let bottom = shape
            .geometry
            .y
            .checked_add(shape.geometry.height)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "RTF grouped shape bottom edge overflows",
                )
            })?;
        self.write_str("{\\shp{\\*\\shpinst")?;
        self.write_control_word("shpleft", Some(shape.geometry.x))?;
        self.write_control_word("shptop", Some(shape.geometry.y))?;
        self.write_control_word("shpright", Some(right))?;
        self.write_control_word("shpbottom", Some(bottom))?;
        self.write_control_word("shpz", Some(shape.geometry.z_order))?;
        self.write_shape_info(&shape.info)?;
        if shape.behind_doc
            && !shape
                .info
                .iter()
                .any(|info| matches!(info, ShapeGroupInfo::BelowText(_)))
        {
            self.write_control_word("shpfblwtxt", Some(1))?;
        }
        if shape.locked
            && !shape
                .info
                .iter()
                .any(|info| matches!(info, ShapeGroupInfo::LockAnchor))
        {
            self.write_control_word("shplockanchor", None)?;
        }
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
        self.write_str("}}")
    }

    pub(in super::super) fn write_shape_scalar_property(
        &mut self,
        name: &str,
        value: &str,
    ) -> io::Result<()> {
        self.write_str("{")?;
        self.write_control_word("sp", None)?;
        self.write_str("{")?;
        self.write_control_word("sn", None)?;
        self.write_str(" ")?;
        self.write_text(name)?;
        self.write_str("}{")?;
        self.write_control_word("sv", None)?;
        self.write_str(" ")?;
        self.write_text(value)?;
        self.write_str("}}")
    }

    pub(in super::super) fn write_shape_property(
        &mut self,
        property: &ShapeProperty<'_>,
    ) -> io::Result<()> {
        property
            .validate()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        self.write_str("{\\sp{\\sn ")?;
        self.write_destination_text(property.name.as_ref())?;
        self.write_str("}{\\sv")?;
        if let Some(value) = &property.binary_value {
            self.write_str("{\\*\\svb ")?;
            for byte in value.iter() {
                write!(self.writer, "{byte:02x}")?;
            }
            self.write_str("}")?;
        } else {
            self.write_str(" ")?;
            self.write_destination_text(property.value.as_ref())?;
        }
        self.write_str("}")?;
        if let Some(theme) = property.theme_value {
            self.write_str("{\\*\\hsv")?;
            self.write_control_word(
                match theme.color {
                    ShapeThemeColor::Accent1 => "caccentone",
                    ShapeThemeColor::Accent2 => "caccenttwo",
                    ShapeThemeColor::Accent3 => "caccentthree",
                    ShapeThemeColor::Accent4 => "caccentfour",
                    ShapeThemeColor::Accent5 => "caccentfive",
                    ShapeThemeColor::Accent6 => "caccentsix",
                    ShapeThemeColor::Background1 => "cbackgroundone",
                    ShapeThemeColor::Background2 => "cbackgroundtwo",
                    ShapeThemeColor::Text1 => "ctextone",
                    ShapeThemeColor::Text2 => "ctexttwo",
                },
                None,
            )?;
            self.write_control_word("ctint", Some(i32::from(theme.tint)))?;
            self.write_control_word("cshade", Some(i32::from(theme.shade)))?;
            self.write_str("}")?;
        }
        if let Some(hyperlink) = &property.hyperlink {
            hyperlink
                .validate()
                .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error.to_string()))?;
            self.write_str("{\\hl")?;
            if let Some(location) = &hyperlink.location {
                self.write_str("{\\hlloc ")?;
                self.write_destination_text(location.as_ref())?;
                self.write_str("}")?;
            }
            if let Some(source) = &hyperlink.source {
                self.write_str("{\\hlsrc ")?;
                self.write_destination_text(source.as_ref())?;
                self.write_str("}")?;
            }
            if let Some(friendly_name) = &hyperlink.friendly_name {
                self.write_str("{\\hlfr ")?;
                self.write_destination_text(friendly_name.as_ref())?;
                self.write_str("}")?;
            }
            self.write_str("}")?;
        }
        self.write_str("}")
    }
}

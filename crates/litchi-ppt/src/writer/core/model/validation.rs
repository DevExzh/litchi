//! Validation and atomic mutation gates for the PPT writer model.

use super::{WriteError, Writer};
use crate::encryption::validate_writer_password;
use crate::modify_password::validate_value as validate_modify_password;
use crate::view_info::{SlideViewInfo, ViewKind};
use crate::writer::core::codec::{interaction_for_hyperlink, shape_text_unit_count};

impl Writer {
    pub(in crate::writer::core) fn validate_encryption(&self) -> Result<(), WriteError> {
        if self.modify_password.is_some() && self.encryption.is_none() {
            return Err(WriteError::InvalidData(
                "PowerPoint modify-password output requires password-to-open encryption"
                    .to_string(),
            ));
        }
        if let Some(value) = &self.encryption {
            validate_writer_password(value.profile, value.password.as_str())
                .map_err(WriteError::InvalidData)?;
        }
        if let Some(value) = &self.modify_password {
            validate_modify_password(value.password.as_str())
                .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        }
        Ok(())
    }

    pub(in crate::writer::core) fn validate_smart_tag_references(&self) -> Result<(), WriteError> {
        let smart_tag_count = u32::try_from(self.smart_tags.len()).map_err(|_| {
            WriteError::InvalidData("PowerPoint smart-tag count exceeds u32".to_string())
        })?;
        for run in self
            .slides
            .iter()
            .flat_map(|slide| &slide.shapes)
            .filter_map(|shape| shape.properties.paragraphs.as_ref())
            .flatten()
            .flat_map(|paragraph| &paragraph.runs)
            .filter(|run| !run.smart_tag_indices.is_empty())
        {
            if run.text.is_empty() {
                return Err(WriteError::InvalidData(
                    "PowerPoint smart tags cannot be attached to an empty text run".to_string(),
                ));
            }
            for index in &run.smart_tag_indices {
                if index.as_u32() >= smart_tag_count {
                    return Err(WriteError::InvalidData(format!(
                        "PowerPoint text run references missing smart tag {}",
                        index.as_u32()
                    )));
                }
            }
        }
        Ok(())
    }
}

impl Writer {
    /// Attach a hyperlink to the last shape added on a slide
    pub fn set_last_shape_hyperlink(
        &mut self,
        slide: usize,
        hyperlink_id: u32,
    ) -> Result<(), WriteError> {
        let interaction =
            interaction_for_hyperlink(hyperlink_id, &self.hyperlinks).ok_or_else(|| {
                WriteError::InvalidData(format!("Hyperlink {hyperlink_id} does not exist"))
            })?;
        self.set_last_shape_interaction(slide, interaction)
    }

    /// Add or replace one typed click or mouse-over action on the last shape.
    ///
    /// Validation is atomic. Hyperlink references must identify an existing
    /// writer hyperlink, and a shape can carry at most one action per trigger.
    pub fn set_last_shape_interaction(
        &mut self,
        slide: usize,
        interaction: crate::Interaction,
    ) -> Result<(), WriteError> {
        self.set_last_shape_interaction_with_limits(
            slide,
            interaction,
            crate::InteractionLimits::default(),
        )
    }

    /// Add or replace one shape action with explicit record and name limits.
    pub fn set_last_shape_interaction_with_limits(
        &mut self,
        slide: usize,
        interaction: crate::Interaction,
        limits: crate::InteractionLimits,
    ) -> Result<(), WriteError> {
        interaction
            .validate_with_limits(limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        if interaction.hyperlink_id != 0 && self.hyperlinks.get(interaction.hyperlink_id).is_none()
        {
            return Err(WriteError::InvalidData(format!(
                "Hyperlink {} does not exist",
                interaction.hyperlink_id
            )));
        }
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        if let Some(shape) = slide_data.shapes.last_mut() {
            shape.properties.hyperlink_id = None;
            if let Some(existing) = shape
                .properties
                .interactions
                .iter_mut()
                .find(|existing| existing.trigger == interaction.trigger)
            {
                *existing = interaction;
            } else {
                shape.properties.interactions.push(interaction);
                shape.properties.interactions.sort_by_key(|interaction| {
                    match interaction.trigger {
                        crate::InteractionTrigger::Click => 0,
                        crate::InteractionTrigger::MouseOver => 1,
                    }
                });
            }
            Ok(())
        } else {
            Err(WriteError::InvalidData("No shapes on slide".to_string()))
        }
    }

    /// Remove one trigger from the last shape, returning whether it was present.
    pub fn clear_last_shape_interaction(
        &mut self,
        slide: usize,
        trigger: crate::InteractionTrigger,
    ) -> Result<bool, WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if trigger == crate::InteractionTrigger::Click {
            shape.properties.hyperlink_id = None;
        }
        let old_len = shape.properties.interactions.len();
        shape
            .properties
            .interactions
            .retain(|interaction| interaction.trigger != trigger);
        Ok(shape.properties.interactions.len() != old_len)
    }

    /// Attach a registered hyperlink to one UTF-16 range in the last shape's text.
    pub fn set_last_shape_text_hyperlink(
        &mut self,
        slide: usize,
        range: crate::TextRange,
        hyperlink_id: u32,
    ) -> Result<(), WriteError> {
        let interaction =
            interaction_for_hyperlink(hyperlink_id, &self.hyperlinks).ok_or_else(|| {
                WriteError::InvalidData(format!("Hyperlink {hyperlink_id} does not exist"))
            })?;
        self.set_last_shape_text_interaction(
            slide,
            crate::TextInteraction::new(range, interaction)
                .map_err(|error| WriteError::InvalidData(error.to_string()))?,
        )
    }

    /// Add or replace one trigger/range pair on the last shape's text.
    ///
    /// Text positions are UTF-16 code units. Validation occurs before mutation.
    pub fn set_last_shape_text_interaction(
        &mut self,
        slide: usize,
        interaction: crate::TextInteraction,
    ) -> Result<(), WriteError> {
        self.set_last_shape_text_interaction_with_limits(
            slide,
            interaction,
            crate::TextInteractionLimits::default(),
        )
    }

    /// Add or replace a text action with explicit resource limits.
    pub fn set_last_shape_text_interaction_with_limits(
        &mut self,
        slide: usize,
        interaction: crate::TextInteraction,
        limits: crate::TextInteractionLimits,
    ) -> Result<(), WriteError> {
        if interaction.interaction.hyperlink_id != 0
            && self
                .hyperlinks
                .get(interaction.interaction.hyperlink_id)
                .is_none()
        {
            return Err(WriteError::InvalidData(format!(
                "Hyperlink {} does not exist",
                interaction.interaction.hyperlink_id
            )));
        }
        let slide_data = self
            .slides
            .get(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        let text_units = shape_text_unit_count(&shape.properties)?;
        interaction
            .validate_for_text(text_units, limits)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        let replace_index = shape
            .properties
            .text_interactions
            .iter()
            .position(|existing| {
                existing.range == interaction.range
                    && existing.interaction.trigger == interaction.interaction.trigger
            });
        let prospective_len = shape
            .properties
            .text_interactions
            .len()
            .checked_add(usize::from(replace_index.is_none()))
            .ok_or_else(|| {
                WriteError::InvalidData("Shape text interaction count overflows".to_string())
            })?;
        if prospective_len > limits.max_interactions {
            return Err(WriteError::InvalidData(
                "Shape exceeds the configured text interaction count".to_string(),
            ));
        }

        let shape = self
            .slides
            .get_mut(slide)
            .and_then(|slide| slide.shapes.last_mut())
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        if let Some(index) = replace_index {
            shape.properties.text_interactions[index] = interaction;
        } else {
            shape.properties.text_interactions.push(interaction);
            shape.properties.text_interactions.sort_by_key(|value| {
                (
                    value.range.begin(),
                    value.range.end(),
                    match value.interaction.trigger {
                        crate::InteractionTrigger::Click => 0,
                        crate::InteractionTrigger::MouseOver => 1,
                    },
                )
            });
        }
        Ok(())
    }

    /// Remove one trigger/range pair from the last shape.
    pub fn clear_last_shape_text_interaction(
        &mut self,
        slide: usize,
        range: crate::TextRange,
        trigger: crate::InteractionTrigger,
    ) -> Result<bool, WriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| WriteError::InvalidData(format!("Slide {} does not exist", slide)))?;
        let shape = slide_data
            .shapes
            .last_mut()
            .ok_or_else(|| WriteError::InvalidData("No shapes on slide".to_string()))?;
        let old_len = shape.properties.text_interactions.len();
        shape
            .properties
            .text_interactions
            .retain(|value| value.range != range || value.interaction.trigger != trigger);
        Ok(shape.properties.text_interactions.len() != old_len)
    }
}

impl Writer {
    fn validate_view_info_kind(view: &SlideViewInfo, expected: ViewKind) -> Result<(), WriteError> {
        view.to_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        if view.kind() != expected {
            return Err(WriteError::InvalidData(format!(
                "editing-view kind {:?} does not match {:?}",
                view.kind(),
                expected
            )));
        }
        Ok(())
    }

    /// Set the presentation's slide editing-view preferences, zoom, and guides.
    pub fn set_slide_view_info(&mut self, view: SlideViewInfo) -> Result<(), WriteError> {
        Self::validate_view_info_kind(&view, ViewKind::Slide)?;
        self.slide_view_info = Some(view);
        Ok(())
    }

    /// Restore the writer's canonical default slide editing view.
    pub fn clear_slide_view_info(&mut self) {
        self.slide_view_info = None;
    }

    /// Return the explicit slide editing-view override, if present.
    pub fn slide_view_info(&self) -> Option<&SlideViewInfo> {
        self.slide_view_info.as_ref()
    }

    /// Set the presentation's notes editing-view preferences, zoom, and guides.
    pub fn set_notes_view_info(&mut self, view: SlideViewInfo) -> Result<(), WriteError> {
        Self::validate_view_info_kind(&view, ViewKind::Notes)?;
        self.notes_view_info = Some(view);
        Ok(())
    }

    /// Remove the optional notes editing-view record.
    pub fn clear_notes_view_info(&mut self) {
        self.notes_view_info = None;
    }

    /// Return the explicit notes editing-view, if present.
    pub fn notes_view_info(&self) -> Option<&SlideViewInfo> {
        self.notes_view_info.as_ref()
    }
}

//! Inert legacy RTF drawing text boxes.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_LEGACY_TEXT_BOXES: usize = 16_384;
pub(crate) const MAX_LEGACY_TEXT_BOX_BYTES: usize = 1_048_576;
pub(crate) const MAX_LEGACY_TEXT_BOX_TOTAL_BYTES: usize = 16 * 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyHorizontalAnchor {
    Page,
    Margin,
    Column,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyVerticalAnchor {
    Page,
    Margin,
    Paragraph,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum LegacyTextDirection {
    #[default]
    LeftToRightTopToBottom,
    LeftToRightTopToBottomVertical,
    TopToBottomRightToLeft,
    TopToBottomRightToLeftVertical,
    BottomToTopLeftToRight,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyTextBox<'a> {
    pub text: Cow<'a, str>,
    /// Positional root shapes owned by the legacy text-box story.
    pub shapes: Vec<crate::Shape<'a>>,
    /// Positional root shape groups owned by the legacy text-box story.
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
    /// Exact source order of drawings in the legacy text-box story.
    pub drawing_order: Vec<crate::StoryDrawing>,
    /// Exact source order of drawings, fields, and page breaks in this story.
    pub story_events: Vec<crate::StoryEvent>,
    pub position: usize,
    pub horizontal_anchor: Option<LegacyHorizontalAnchor>,
    pub vertical_anchor: Option<LegacyVerticalAnchor>,
    pub x: Option<i32>,
    pub y: Option<i32>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub margin: Option<i32>,
    pub z_order: Option<i32>,
    pub direction: LegacyTextDirection,
}

impl LegacyTextBox<'_> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.text.is_empty() || self.text.len() > MAX_LEGACY_TEXT_BOX_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box text is empty or exceeds the safety limit".to_string(),
            ));
        }
        if self.text.contains('\0') {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box text contains a NUL character".to_string(),
            ));
        }
        crate::field::validate_story_events(
            self.text.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.story_events,
            "legacy text box",
        )?;
        if self.width.is_some_and(|value| value <= 0)
            || self.height.is_some_and(|value| value <= 0)
            || self.margin.is_some_and(|value| value < 0)
        {
            return Err(RtfError::MalformedDocument(
                "RTF legacy text-box size or margin is outside its valid range".to_string(),
            ));
        }
        Ok(())
    }

    /// Append a validated positional root shape to this legacy text-box story.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape(&mut self, shape: crate::Shape<'_>) -> RtfResult<()> {
        let mut shapes = self.shapes.clone();
        shapes.push(shape.into_owned());
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::Shape(self.shapes.len()));
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &shapes,
            &self.shape_groups,
            &order,
            "legacy text box",
        )?;
        self.shapes = shapes;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len() - 1,
            )));
        Ok(())
    }

    /// Append a validated positional root shape group to this legacy text-box story.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'_>) -> RtfResult<()> {
        let mut groups = self.shape_groups.clone();
        groups.push(group.into_owned());
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::ShapeGroup(self.shape_groups.len()));
        crate::shape::validate_story_drawings(
            self.text.as_ref(),
            &self.shapes,
            &groups,
            &order,
            "legacy text box",
        )?;
        self.shape_groups = groups;
        self.drawing_order = order;
        self.story_events
            .push(crate::StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                self.shape_groups.len() - 1,
            )));
        Ok(())
    }

    /// Clear all drawings owned by this legacy text-box story.
    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = &crate::PageBreak> {
        self.story_events.iter().filter_map(|event| match event {
            crate::StoryEvent::PageBreak(page_break) => Some(page_break),
            crate::StoryEvent::Drawing(_) | crate::StoryEvent::Field(_) => None,
        })
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_page_break(&mut self, position: usize) -> RtfResult<()> {
        crate::field::push_story_page_break(
            &mut self.story_events,
            self.text.as_ref(),
            position,
            "legacy text box",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.story_events
            .retain(|event| !matches!(event, crate::StoryEvent::PageBreak(_)));
    }

    pub(crate) fn into_owned(self) -> LegacyTextBox<'static> {
        LegacyTextBox {
            text: Cow::Owned(self.text.into_owned()),
            shapes: self
                .shapes
                .into_iter()
                .map(crate::Shape::into_owned)
                .collect(),
            shape_groups: self
                .shape_groups
                .into_iter()
                .map(crate::ShapeGroup::into_owned)
                .collect(),
            drawing_order: self.drawing_order,
            story_events: self.story_events,
            position: self.position,
            horizontal_anchor: self.horizontal_anchor,
            vertical_anchor: self.vertical_anchor,
            x: self.x,
            y: self.y,
            width: self.width,
            height: self.height,
            margin: self.margin,
            z_order: self.z_order,
            direction: self.direction,
        }
    }
}

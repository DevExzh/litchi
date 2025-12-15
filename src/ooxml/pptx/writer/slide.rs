/// Slide types and implementation for PPTX presentations.
use crate::common::xml::escape_xml;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::pptx::animations::{
    Animation, AnimationEffect, AnimationSequence, AnimationTrigger,
};
use crate::ooxml::pptx::media::{Media, MediaFormat};
use crate::ooxml::pptx::parts::Comment;
use std::fmt::Write as FmtWrite;

// Import shared format types
use super::super::format::ImageFormat;
use super::shape::MutableShape;

/// A mutable slide in a presentation.
#[derive(Debug, Clone)]
pub struct MutableSlide {
    /// Slide ID (unique identifier)
    pub(crate) slide_id: u32,
    /// Slide title (stored in title placeholder)
    pub(crate) title: Option<String>,
    /// Shapes on the slide
    pub(crate) shapes: Vec<MutableShape>,
    /// Speaker notes for the slide
    pub(crate) notes: Option<String>,
    /// Slide transition effect
    pub(crate) transition: Option<crate::ooxml::pptx::transitions::SlideTransition>,
    /// Slide background
    pub(crate) background: Option<crate::ooxml::pptx::backgrounds::SlideBackground>,
    /// Comments on the slide
    pub(crate) comments: Vec<Comment>,
    /// Media elements (audio/video) on the slide
    pub(crate) media: Vec<Media>,
    /// Animations on the slide
    pub(crate) animations: AnimationSequence,
    /// Whether the slide has been modified
    pub(crate) modified: bool,
}

impl MutableSlide {
    /// Create a new empty slide.
    pub(crate) fn new(slide_id: u32) -> Self {
        Self {
            slide_id,
            title: None,
            shapes: Vec::new(),
            notes: None,
            transition: None,
            background: None,
            comments: Vec::new(),
            media: Vec::new(),
            animations: AnimationSequence::new(),
            modified: false,
        }
    }

    /// Get the slide ID.
    pub fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Set the slide ID.
    ///
    /// This is used internally when duplicating slides to assign new IDs.
    pub(crate) fn set_slide_id(&mut self, slide_id: u32) {
        self.slide_id = slide_id;
        self.modified = true;
    }

    /// Set the slide title.
    pub fn set_title(&mut self, title: &str) {
        self.title = Some(title.to_string());
        self.modified = true;
    }

    /// Get the slide title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set speaker notes for the slide.
    pub fn set_notes(&mut self, notes: &str) {
        self.notes = Some(notes.to_string());
        self.modified = true;
    }

    /// Get the speaker notes for the slide.
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Check if the slide has speaker notes.
    pub fn has_notes(&self) -> bool {
        self.notes.is_some()
    }

    /// Set a transition effect for the slide.
    ///
    /// # Arguments
    /// * `transition` - The transition configuration
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, TransitionType, TransitionSpeed, SlideTransition};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Add a fade transition
    /// let transition = SlideTransition::new(TransitionType::Fade)
    ///     .with_speed(TransitionSpeed::Fast)
    ///     .with_advance_after_ms(3000);
    /// slide.set_transition(transition);
    /// ```
    pub fn set_transition(&mut self, transition: crate::ooxml::pptx::transitions::SlideTransition) {
        self.transition = Some(transition);
        self.modified = true;
    }

    /// Get the transition effect for the slide.
    ///
    /// Returns `None` if no transition is set.
    pub fn transition(&self) -> Option<&crate::ooxml::pptx::transitions::SlideTransition> {
        self.transition.as_ref()
    }

    /// Remove the transition effect from the slide.
    pub fn remove_transition(&mut self) {
        self.transition = None;
        self.modified = true;
    }

    /// Set a background for the slide.
    ///
    /// # Arguments
    /// * `background` - The background configuration
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, SlideBackground};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Set a solid blue background
    /// slide.set_background(SlideBackground::solid("4472C4"));
    /// ```
    pub fn set_background(&mut self, background: crate::ooxml::pptx::backgrounds::SlideBackground) {
        self.background = Some(background);
        self.modified = true;
    }

    /// Get the background for the slide.
    ///
    /// Returns `None` if no background is set.
    pub fn background(&self) -> Option<&crate::ooxml::pptx::backgrounds::SlideBackground> {
        self.background.as_ref()
    }

    /// Remove the background from the slide (use master background).
    pub fn remove_background(&mut self) {
        self.background = None;
        self.modified = true;
    }

    /// Add a text box to the slide.
    pub fn add_text_box(&mut self, text: &str, x: i64, y: i64, width: i64, height: i64) {
        // IDs: 1=group, 2=title, 3+=user shapes
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_text_box(shape_id, text.to_string(), x, y, width, height);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a rectangle to the slide.
    pub fn add_rectangle(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_rectangle(shape_id, x, y, width, height, fill_color);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add an ellipse (circle/oval) to the slide.
    pub fn add_ellipse(
        &mut self,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) {
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_ellipse(shape_id, x, y, width, height, fill_color);
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a picture to the slide from a file.
    pub fn add_picture(
        &mut self,
        image_path: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Result<()> {
        use std::fs;

        let data = fs::read(image_path).map_err(OoxmlError::IoError)?;

        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown image format".to_string()))?;

        let shape_id = (self.shapes.len() + 3) as u32;
        let description = format!("Picture from {}", image_path);
        let shape =
            MutableShape::new_picture(shape_id, data, format, x, y, width, height, description)?;
        self.shapes.push(shape);
        self.modified = true;

        Ok(())
    }

    /// Add a picture to the slide from bytes.
    pub fn add_picture_from_bytes(
        &mut self,
        data: Vec<u8>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        description: Option<String>,
    ) -> Result<()> {
        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown image format".to_string()))?;

        let shape_id = (self.shapes.len() + 3) as u32;
        let desc = description.unwrap_or_else(|| "Picture".to_string());
        let shape = MutableShape::new_picture(shape_id, data, format, x, y, width, height, desc)?;
        self.shapes.push(shape);
        self.modified = true;

        Ok(())
    }

    /// Add a table to the slide.
    ///
    /// # Arguments
    /// * `data` - 2D vector of cell text content (rows x columns)
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `width` - Table width in EMUs
    /// * `height` - Table height in EMUs
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Create a 2x3 table
    /// let data = vec![
    ///     vec!["Header 1".to_string(), "Header 2".to_string(), "Header 3".to_string()],
    ///     vec!["Cell A".to_string(), "Cell B".to_string(), "Cell C".to_string()],
    /// ];
    /// slide.add_table(data, 914400, 914400, 5486400, 1828800);
    /// ```
    pub fn add_table(&mut self, data: Vec<Vec<String>>, x: i64, y: i64, width: i64, height: i64) {
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_table(
            shape_id, x, y, width, height, data, None, None,
            true, // first row is header by default
            true, // band rows by default
        );
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a table to the slide with custom options.
    ///
    /// # Arguments
    /// * `data` - 2D vector of cell text content (rows x columns)
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `width` - Table width in EMUs
    /// * `height` - Table height in EMUs
    /// * `col_widths` - Optional column widths in EMUs
    /// * `row_heights` - Optional row heights in EMUs
    /// * `first_row_header` - Whether the first row should be styled as a header
    /// * `band_rows` - Whether to use alternating row colors
    #[allow(clippy::too_many_arguments)]
    pub fn add_table_with_options(
        &mut self,
        data: Vec<Vec<String>>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        col_widths: Option<Vec<i64>>,
        row_heights: Option<Vec<i64>>,
        first_row_header: bool,
        band_rows: bool,
    ) {
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_table(
            shape_id,
            x,
            y,
            width,
            height,
            data,
            col_widths,
            row_heights,
            first_row_header,
            band_rows,
        );
        self.shapes.push(shape);
        self.modified = true;
    }

    /// Add a group shape to the slide.
    ///
    /// Group shapes contain multiple child shapes that can be manipulated together.
    ///
    /// # Arguments
    /// * `x` - X position of the group in EMUs
    /// * `y` - Y position of the group in EMUs
    /// * `width` - Width of the group in EMUs
    /// * `height` - Height of the group in EMUs
    ///
    /// # Returns
    /// A mutable reference to the newly created group shape for adding child shapes.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// // Create a group and get its index
    /// let group_idx = slide.add_group(914400, 914400, 2743200, 2743200);
    ///
    /// // Add shapes to the group using add_shape_to_group
    /// slide.add_rectangle_to_group(group_idx, 0, 0, 914400, 914400, Some("FF0000".to_string()));
    /// slide.add_ellipse_to_group(group_idx, 914400, 0, 914400, 914400, Some("00FF00".to_string()));
    /// ```
    pub fn add_group(&mut self, x: i64, y: i64, width: i64, height: i64) -> usize {
        let shape_id = (self.shapes.len() + 3) as u32;
        let shape = MutableShape::new_group(shape_id, x, y, width, height, Vec::new());
        self.shapes.push(shape);
        self.modified = true;
        self.shapes.len() - 1
    }

    /// Add a text box to a group shape.
    ///
    /// # Arguments
    /// * `group_idx` - Index of the group shape
    /// * `text` - Text content
    /// * `x` - X position relative to the group
    /// * `y` - Y position relative to the group
    /// * `width` - Width in EMUs
    /// * `height` - Height in EMUs
    ///
    /// # Returns
    /// `true` if the shape was added successfully, `false` if the group index is invalid
    pub fn add_text_box_to_group(
        &mut self,
        group_idx: usize,
        text: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> bool {
        if let Some(shape) = self.shapes.get_mut(group_idx)
            && let Some(children) = shape.get_children_mut()
        {
            let child_id = (children.len() + 100) as u32; // Use different ID range for children
            let child = MutableShape::new_text_box(child_id, text.to_string(), x, y, width, height);
            children.push(child);
            self.modified = true;
            return true;
        }
        false
    }

    /// Add a rectangle to a group shape.
    ///
    /// # Arguments
    /// * `group_idx` - Index of the group shape
    /// * `x` - X position relative to the group
    /// * `y` - Y position relative to the group
    /// * `width` - Width in EMUs
    /// * `height` - Height in EMUs
    /// * `fill_color` - Optional fill color (hex RGB)
    ///
    /// # Returns
    /// `true` if the shape was added successfully, `false` if the group index is invalid
    pub fn add_rectangle_to_group(
        &mut self,
        group_idx: usize,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> bool {
        if let Some(shape) = self.shapes.get_mut(group_idx)
            && let Some(children) = shape.get_children_mut()
        {
            let child_id = (children.len() + 100) as u32;
            let child = MutableShape::new_rectangle(child_id, x, y, width, height, fill_color);
            children.push(child);
            self.modified = true;
            return true;
        }
        false
    }

    /// Add an ellipse to a group shape.
    ///
    /// # Arguments
    /// * `group_idx` - Index of the group shape
    /// * `x` - X position relative to the group
    /// * `y` - Y position relative to the group
    /// * `width` - Width in EMUs
    /// * `height` - Height in EMUs
    /// * `fill_color` - Optional fill color (hex RGB)
    ///
    /// # Returns
    /// `true` if the shape was added successfully, `false` if the group index is invalid
    pub fn add_ellipse_to_group(
        &mut self,
        group_idx: usize,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> bool {
        if let Some(shape) = self.shapes.get_mut(group_idx)
            && let Some(children) = shape.get_children_mut()
        {
            let child_id = (children.len() + 100) as u32;
            let child = MutableShape::new_ellipse(child_id, x, y, width, height, fill_color);
            children.push(child);
            self.modified = true;
            return true;
        }
        false
    }

    /// Get the number of shapes on the slide.
    pub fn shape_count(&self) -> usize {
        self.shapes.len()
    }

    /// Check if the slide has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    // ========================================================================
    // Comments
    // ========================================================================

    /// Add a comment to the slide.
    ///
    /// # Arguments
    /// * `author_id` - ID of the comment author (must match an author in the presentation)
    /// * `text` - Comment text
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::MutablePresentation;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    /// slide.add_comment(0, "This is a comment", 914400, 914400);
    /// ```
    pub fn add_comment(&mut self, author_id: u32, text: &str, x: i64, y: i64) {
        // idx is required for comments - auto-assign based on count
        let idx = self.comments.len() as u32;
        let comment = Comment::new(author_id, text, x, y).with_index(idx);
        self.comments.push(comment);
        self.modified = true;
    }

    /// Add a comment with additional options.
    pub fn add_comment_with_options(
        &mut self,
        author_id: u32,
        text: &str,
        x: i64,
        y: i64,
        datetime: Option<&str>,
        index: Option<u32>,
    ) {
        let mut comment = Comment::new(author_id, text, x, y);
        if let Some(dt) = datetime {
            comment = comment.with_datetime(dt);
        }
        if let Some(idx) = index {
            comment = comment.with_index(idx);
        }
        self.comments.push(comment);
        self.modified = true;
    }

    /// Get all comments on this slide.
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Get the number of comments on this slide.
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    // ========================================================================
    // Media (Audio/Video)
    // ========================================================================

    /// Add an audio file to the slide.
    ///
    /// # Arguments
    /// * `data` - Audio file data
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `width` - Width in EMUs (for the audio icon)
    /// * `height` - Height in EMUs (for the audio icon)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::MutablePresentation;
    /// use std::fs;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    /// let audio_data = fs::read("background.mp3").unwrap();
    /// slide.add_audio(audio_data, 914400, 914400, 914400, 914400);
    /// ```
    pub fn add_audio(&mut self, data: Vec<u8>, x: i64, y: i64, width: i64, height: i64) {
        let media = Media::new(data, x, y, width, height);
        self.media.push(media);
        self.modified = true;
    }

    /// Add a video file to the slide.
    ///
    /// # Arguments
    /// * `data` - Video file data
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `width` - Width in EMUs
    /// * `height` - Height in EMUs
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::MutablePresentation;
    /// use std::fs;
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    /// let video_data = fs::read("intro.mp4").unwrap();
    /// slide.add_video(video_data, 914400, 914400, 4572000, 2571750);
    /// ```
    pub fn add_video(&mut self, data: Vec<u8>, x: i64, y: i64, width: i64, height: i64) {
        let media = Media::new(data, x, y, width, height);
        self.media.push(media);
        self.modified = true;
    }

    /// Add media with explicit format.
    pub fn add_media_with_format(
        &mut self,
        data: Vec<u8>,
        format: MediaFormat,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) {
        let media = Media::with_format(data, format, x, y, width, height);
        self.media.push(media);
        self.modified = true;
    }

    /// Get all media elements on this slide.
    pub fn media(&self) -> &[Media] {
        &self.media
    }

    /// Get the number of media elements on this slide.
    pub fn media_count(&self) -> usize {
        self.media.len()
    }

    // ========================================================================
    // Animations
    // ========================================================================

    /// Add an animation to a shape on the slide.
    ///
    /// # Arguments
    /// * `shape_id` - ID of the shape to animate (must be an existing shape on the slide)
    /// * `effect` - The animation effect type
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::pptx::{MutablePresentation, AnimationEffect};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    /// slide.add_text_box("Animated Text", 914400, 914400, 2743200, 914400);
    /// slide.add_animation(3, AnimationEffect::Fade); // Shape ID 3 is the text box
    /// ```
    pub fn add_animation(&mut self, shape_id: u32, effect: AnimationEffect) {
        let animation = Animation::new(shape_id, effect);
        self.animations.add(animation);
        self.modified = true;
    }

    /// Add an animation with custom settings.
    ///
    /// # Arguments
    /// * `shape_id` - ID of the shape to animate
    /// * `effect` - The animation effect type
    /// * `trigger` - When to trigger the animation
    /// * `duration_ms` - Duration in milliseconds
    /// * `delay_ms` - Delay before starting in milliseconds
    pub fn add_animation_with_options(
        &mut self,
        shape_id: u32,
        effect: AnimationEffect,
        trigger: AnimationTrigger,
        duration_ms: u32,
        delay_ms: u32,
    ) {
        let animation = Animation::new(shape_id, effect)
            .with_trigger(trigger)
            .with_duration(duration_ms)
            .with_delay(delay_ms);
        self.animations.add(animation);
        self.modified = true;
    }

    /// Get the animations on this slide.
    pub fn animations(&self) -> &AnimationSequence {
        &self.animations
    }

    /// Get the number of animations on this slide.
    pub fn animation_count(&self) -> usize {
        self.animations.len()
    }

    /// Clear all animations from the slide.
    pub fn clear_animations(&mut self) {
        self.animations = AnimationSequence::new();
        self.modified = true;
    }

    // ========================================================================
    // Charts and SmartArt
    // ========================================================================

    /// Add a chart to the slide.
    ///
    /// This method registers the chart with the presentation and adds a chart
    /// shape to the slide. The chart data will be embedded as an Excel workbook
    /// when the presentation is saved.
    ///
    /// # Arguments
    /// * `chart_data` - The chart data including type, series, and position
    /// * `pres` - Mutable reference to the presentation (needed to register chart parts)
    ///
    /// # Returns
    /// * `Ok(u32)` - The shape ID of the chart (can be used for animations)
    /// * `Err` if chart registration fails
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::{MutablePresentation, ChartData, ChartSeries, ChartType};
    ///
    /// let mut pres = MutablePresentation::new();
    /// let slide = pres.add_slide().unwrap();
    ///
    /// let chart = ChartData::new(ChartType::Bar, 914400, 914400, 4572000, 2743200)
    ///     .with_title("Sales by Quarter")
    ///     .add_series(
    ///         ChartSeries::new("2024")
    ///             .with_categories(vec!["Q1".into(), "Q2".into(), "Q3".into(), "Q4".into()])
    ///             .with_values(vec![100.0, 150.0, 200.0, 175.0])
    ///     );
    ///
    /// // Note: In practice, you'd need to pass &mut pres which requires different API design
    /// // slide.add_chart(&chart, &mut pres)?;
    /// ```
    pub fn add_chart_shape(
        &mut self,
        chart_idx: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> u32 {
        use super::shape::{MutableShape, ShapeType};

        // IDs: 1=group, 2=title, 3+=user shapes
        let shape_id = (self.shapes.len() + 3) as u32;

        // Create chart shape with placeholder relationship ID
        // The actual relationship ID will be assigned during save
        let shape = MutableShape {
            shape_id,
            shape_type: ShapeType::Chart {
                x,
                y,
                width,
                height,
                chart_rel_id: String::new(), // Will be set during save
                chart_idx,
            },
        };

        self.shapes.push(shape);
        self.modified = true;
        shape_id
    }

    /// Add a SmartArt diagram to the slide.
    ///
    /// This method registers the SmartArt with the presentation and adds a
    /// diagram shape to the slide. The diagram parts will be written when
    /// the presentation is saved.
    ///
    /// # Arguments
    /// * `diagram_idx` - The diagram index from `pres.add_smartart_parts()`
    /// * `x` - X position in EMUs
    /// * `y` - Y position in EMUs
    /// * `width` - Width in EMUs
    /// * `height` - Height in EMUs
    ///
    /// # Returns
    /// The shape ID of the SmartArt (can be used for animations)
    pub fn add_smartart_shape(
        &mut self,
        diagram_idx: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> u32 {
        use super::shape::{MutableShape, ShapeType};

        let shape_id = (self.shapes.len() + 3) as u32;

        // Create SmartArt shape with placeholder relationship IDs
        // The actual relationship IDs will be assigned during save
        let shape = MutableShape {
            shape_id,
            shape_type: ShapeType::SmartArt {
                x,
                y,
                width,
                height,
                data_rel_id: String::new(),
                layout_rel_id: String::new(),
                style_rel_id: String::new(),
                colors_rel_id: String::new(),
                diagram_idx,
            },
        };

        self.shapes.push(shape);
        self.modified = true;
        shape_id
    }

    /// Collect all media from this slide.
    #[allow(dead_code)] // Will be used by package writer
    pub(crate) fn collect_media(&self) -> Vec<(&[u8], MediaFormat)> {
        self.media
            .iter()
            .map(|m| (m.data.as_slice(), m.format))
            .collect()
    }

    /// Collect all images from this slide (from shapes only, not background).
    pub(crate) fn collect_images(&self) -> Vec<(&[u8], ImageFormat)> {
        let mut images = Vec::new();

        for shape in &self.shapes {
            if let Some((data, format)) = shape.get_image_data() {
                images.push((data, format));
            }
        }

        images
    }

    /// Get the background image if this slide has a picture background.
    ///
    /// Returns `Some((image_data, format))` if the background is a picture,
    /// otherwise returns `None`.
    pub(crate) fn get_background_image(&self) -> Option<(&[u8], ImageFormat)> {
        self.background
            .as_ref()
            .and_then(|bg| bg.get_image_data())
            .map(|(data, &format)| (data, format))
    }

    /// Generate slide XML content.
    #[allow(dead_code)] // Public API but not used in the current implementation
    pub(crate) fn to_xml(&self) -> Result<String> {
        self.to_xml_with_rels(None, None)
    }

    /// Generate slide XML content with relationship IDs from the mapper.
    ///
    /// # Arguments
    /// * `slide_index` - The index of this slide (used to look up relationships)
    /// * `rel_mapper` - The relationship mapper containing actual relationship IDs
    pub(crate) fn to_xml_with_rels(
        &self,
        slide_index: Option<usize>,
        rel_mapper: Option<&crate::ooxml::pptx::writer::relmap::RelationshipMapper>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        );
        xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#);
        xml.push_str(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        xml.push_str("<p:cSld>");

        // Add background if present (must come BEFORE spTree per OOXML spec)
        if let Some(ref background) = self.background {
            // For picture backgrounds, we need to get the relationship ID
            let bg_rel_id = if background.get_image_data().is_some() {
                // Get actual relationship ID from mapper
                slide_index.and_then(|si| rel_mapper.and_then(|rm| rm.get_background_id(si)))
            } else {
                None
            };
            xml.push_str(&background.to_xml(bg_rel_id)?);
        }

        xml.push_str("<p:spTree>");

        // Write group shape properties (required)
        xml.push_str("<p:nvGrpSpPr>");
        xml.push_str(r#"<p:cNvPr id="1" name=""/>"#);
        xml.push_str("<p:cNvGrpSpPr/>");
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGrpSpPr>");
        xml.push_str("<p:grpSpPr>");
        xml.push_str("<a:xfrm>");
        xml.push_str(r#"<a:off x="0" y="0"/>"#);
        xml.push_str(r#"<a:ext cx="0" cy="0"/>"#);
        xml.push_str(r#"<a:chOff x="0" y="0"/>"#);
        xml.push_str(r#"<a:chExt cx="0" cy="0"/>"#);
        xml.push_str("</a:xfrm>");
        xml.push_str("</p:grpSpPr>");

        // Write title placeholder if title is set
        if let Some(ref title) = self.title {
            self.write_title_shape(&mut xml, title)?;
        }

        // Write shapes with relationship IDs
        let mut image_counter = 0;
        for shape in &self.shapes {
            use super::shape::{ShapeRelIds, ShapeType};

            // Build relationship IDs based on shape type
            let rel_ids = match &shape.shape_type {
                ShapeType::Picture { .. } => {
                    let rid = slide_index.and_then(|si| {
                        rel_mapper.and_then(|rm| rm.get_image_id(si, image_counter))
                    });
                    image_counter += 1;
                    ShapeRelIds {
                        image_rel_id: rid,
                        ..Default::default()
                    }
                },
                ShapeType::Chart { chart_idx, .. } => {
                    let rid = slide_index
                        .and_then(|si| rel_mapper.and_then(|rm| rm.get_chart_id(si, *chart_idx)));
                    ShapeRelIds {
                        chart_rel_id: rid,
                        ..Default::default()
                    }
                },
                ShapeType::SmartArt { diagram_idx, .. } => {
                    let rids = slide_index.and_then(|si| {
                        rel_mapper.and_then(|rm| rm.get_smartart_ids(si, *diagram_idx))
                    });
                    ShapeRelIds {
                        smartart_rel_ids: rids,
                        ..Default::default()
                    }
                },
                _ => ShapeRelIds::default(),
            };

            shape.to_xml(&mut xml, rel_ids)?;
        }

        // Write media shapes (audio/video)
        let base_shape_id = self.shapes.len() as u32 + 10; // Start after regular shapes
        for (media_idx, media) in self.media.iter().enumerate() {
            let media_rel_ids = slide_index
                .and_then(|si| rel_mapper.and_then(|rm| rm.get_media_ids(si, media_idx)));

            if let Some((video_rid, media_rid, poster_rid)) = media_rel_ids {
                self.write_media_shape(
                    &mut xml,
                    media,
                    base_shape_id + media_idx as u32,
                    video_rid,
                    media_rid,
                    poster_rid,
                )?;
            }
        }

        xml.push_str("</p:spTree>");
        xml.push_str("</p:cSld>");

        xml.push_str(r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#);

        // Add transition if present
        if let Some(ref transition) = self.transition {
            xml.push_str(&transition.to_xml()?);
        }

        // Add timing/animations if present
        if !self.animations.is_empty() {
            xml.push_str(&self.animations.to_xml());
        }

        xml.push_str("</p:sld>");

        Ok(xml)
    }

    /// Generate notes slide XML content.
    pub(crate) fn generate_notes_xml(&self) -> Option<Result<String>> {
        let notes_text = self.notes.as_ref()?;

        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(
            r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        );
        xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#);
        xml.push_str(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        xml.push_str("<p:cSld>");
        xml.push_str("<p:spTree>");

        // Group shape properties
        xml.push_str("<p:nvGrpSpPr>");
        xml.push_str(r#"<p:cNvPr id="1" name=""/>"#);
        xml.push_str("<p:cNvGrpSpPr/>");
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGrpSpPr>");
        xml.push_str("<p:grpSpPr>");
        xml.push_str("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>");
        xml.push_str("<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm>");
        xml.push_str("</p:grpSpPr>");

        // Notes text shape
        xml.push_str("<p:sp>");
        xml.push_str("<p:nvSpPr>");
        xml.push_str(r#"<p:cNvPr id="2" name="Notes Placeholder"/>"#);
        xml.push_str("<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>");
        xml.push_str("<p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr>");
        xml.push_str("</p:nvSpPr>");

        xml.push_str("<p:spPr/>");

        xml.push_str("<p:txBody>");
        xml.push_str("<a:bodyPr/>");
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:r>");
        xml.push_str("<a:rPr lang=\"en-US\" dirty=\"0\"/>");
        if let Err(e) = write!(xml, "<a:t>{}</a:t>", escape_xml(notes_text)) {
            return Some(Err(OoxmlError::Xml(e.to_string())));
        }
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</p:txBody>");
        xml.push_str("</p:sp>");

        xml.push_str("</p:spTree>");
        xml.push_str("</p:cSld>");
        xml.push_str(r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#);
        xml.push_str("</p:notes>");

        Some(Ok(xml))
    }

    /// Write the title placeholder shape.
    fn write_title_shape(&self, xml: &mut String, title: &str) -> Result<()> {
        xml.push_str("<p:sp>");
        xml.push_str("<p:nvSpPr>");
        // Note: ID must be unique within slide. Group shape uses id=1, so title uses id=2.
        xml.push_str(r#"<p:cNvPr id="2" name="Title 1"/>"#);
        xml.push_str("<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>");
        xml.push_str(r#"<p:nvPr><p:ph type="ctrTitle"/></p:nvPr>"#);
        xml.push_str("</p:nvSpPr>");

        xml.push_str("<p:spPr/>");

        xml.push_str("<p:txBody>");
        xml.push_str("<a:bodyPr/>");
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:r>");
        xml.push_str("<a:rPr lang=\"en-US\" dirty=\"0\" smtClean=\"0\"/>");
        write!(xml, "<a:t>{}</a:t>", escape_xml(title))
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</p:txBody>");

        xml.push_str("</p:sp>");

        Ok(())
    }

    /// Write a media (audio/video) shape to XML.
    ///
    /// Media in PPTX is represented as a `<p:pic>` element with audio/video file reference
    /// in the non-visual properties. According to the OOXML spec and python-pptx analysis:
    /// - `<a:videoFile r:link="..."/>` references the OOXML video/audio relationship
    /// - `<p14:media r:embed="..."/>` references the Microsoft media relationship
    /// - `<a:blip r:embed="..."/>` references the poster frame image
    fn write_media_shape(
        &self,
        xml: &mut String,
        media: &Media,
        shape_id: u32,
        video_rel_id: &str,
        media_rel_id: &str,
        poster_rel_id: &str,
    ) -> Result<()> {
        use crate::ooxml::pptx::media::MediaType;

        let is_video = media.media_type() == MediaType::Video;
        let name = media
            .name
            .as_deref()
            .unwrap_or(if is_video { "Video" } else { "Audio" });

        xml.push_str("<p:pic>");

        // Non-visual picture properties
        xml.push_str("<p:nvPicPr>");

        // cNvPr - common non-visual properties with media click action
        write!(
            xml,
            r#"<p:cNvPr id="{}" name="{}">"#,
            shape_id,
            escape_xml(name)
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        // Add click action for media playback
        xml.push_str(r#"<a:hlinkClick r:id="" action="ppaction://media"/>"#);
        xml.push_str("</p:cNvPr>");

        // cNvPicPr - picture-specific non-visual properties
        xml.push_str("<p:cNvPicPr>");
        xml.push_str(r#"<a:picLocks noChangeAspect="1"/>"#);
        xml.push_str("</p:cNvPicPr>");

        // nvPr - non-visual properties with media reference
        xml.push_str("<p:nvPr>");

        // Audio or video file reference (uses r:link referencing OOXML video/audio relationship)
        if is_video {
            write!(xml, r#"<a:videoFile r:link="{}"/>"#, video_rel_id)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        } else {
            write!(xml, r#"<a:audioFile r:link="{}"/>"#, video_rel_id)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }

        // Add p14:media extension (required for PowerPoint 2010+ compatibility)
        // This uses r:embed referencing the Microsoft media relationship
        xml.push_str("<p:extLst>");
        xml.push_str(r#"<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">"#);
        write!(
            xml,
            r#"<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="{}"/>"#,
            media_rel_id
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</p:ext>");
        xml.push_str("</p:extLst>");

        xml.push_str("</p:nvPr>");
        xml.push_str("</p:nvPicPr>");

        // Blip fill - poster frame image (REQUIRED for media shapes)
        xml.push_str("<p:blipFill>");
        write!(xml, r#"<a:blip r:embed="{}"/>"#, poster_rel_id)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
        xml.push_str("</p:blipFill>");

        // Shape properties - position and size
        xml.push_str("<p:spPr>");
        xml.push_str("<a:xfrm>");
        write!(xml, r#"<a:off x="{}" y="{}"/>"#, media.x, media.y)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        write!(
            xml,
            r#"<a:ext cx="{}" cy="{}"/>"#,
            media.width, media.height
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</a:xfrm>");
        xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
        xml.push_str("</p:spPr>");

        xml.push_str("</p:pic>");

        Ok(())
    }
}

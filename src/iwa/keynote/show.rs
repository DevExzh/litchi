//! Keynote Show Structure
//!
//! A show represents the presentation container with metadata and settings.

use super::slide::KeynoteSlide;

/// Represents the overall presentation show
#[derive(Debug, Clone)]
pub struct KeynoteShow {
    /// Show title
    pub title: Option<String>,
    /// Slides in the show
    pub slides: Vec<KeynoteSlide>,
    /// Slide size (width, height) in points
    pub slide_size: Option<(f32, f32)>,
    /// Auto-play setting
    pub auto_play: bool,
    /// Loop presentation
    pub loop_presentation: bool,
}

impl KeynoteShow {
    /// Create a new show
    pub fn new() -> Self {
        Self {
            title: None,
            slides: Vec::new(),
            slide_size: Some((1024.0, 768.0)), // Default size
            auto_play: false,
            loop_presentation: false,
        }
    }

    /// Add a slide to the show
    pub fn add_slide(&mut self, slide: KeynoteSlide) {
        self.slides.push(slide);
    }

    /// Get total number of slides
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Get a specific slide by index
    pub fn get_slide(&self, index: usize) -> Option<&KeynoteSlide> {
        self.slides.get(index)
    }

    /// Get all text content from all slides
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();

        if let Some(ref title) = self.title {
            all_text.push(title.clone());
        }

        for slide in &self.slides {
            all_text.extend(slide.all_text());
        }

        all_text
    }

    /// Check if show is empty
    pub fn is_empty(&self) -> bool {
        self.slides.is_empty()
    }
}

impl Default for KeynoteShow {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_show_creation() {
        let show = KeynoteShow::new();
        assert!(show.is_empty());
        assert_eq!(show.slide_count(), 0);
        assert_eq!(show.slide_size, Some((1024.0, 768.0)));
    }

    #[test]
    fn test_show_add_slide() {
        let mut show = KeynoteShow::new();

        let mut slide = KeynoteSlide::new(0);
        slide.title = Some("Slide 1".to_string());
        show.add_slide(slide);

        assert_eq!(show.slide_count(), 1);
        assert!(!show.is_empty());

        let retrieved = show.get_slide(0);
        assert!(retrieved.is_some());
        assert_eq!(retrieved.unwrap().title, Some("Slide 1".to_string()));
    }

    #[test]
    fn test_show_all_text() {
        let mut show = KeynoteShow::new();
        show.title = Some("My Presentation".to_string());

        let mut slide1 = KeynoteSlide::new(0);
        slide1.title = Some("Slide 1".to_string());
        slide1.text_content.push("Content 1".to_string());
        show.add_slide(slide1);

        let mut slide2 = KeynoteSlide::new(1);
        slide2.title = Some("Slide 2".to_string());
        slide2.text_content.push("Content 2".to_string());
        show.add_slide(slide2);

        let all_text = show.all_text();
        assert!(all_text.contains(&"My Presentation".to_string()));
        assert!(all_text.contains(&"Slide 1".to_string()));
        assert!(all_text.contains(&"Slide 2".to_string()));
        assert!(all_text.contains(&"Content 1".to_string()));
        assert!(all_text.contains(&"Content 2".to_string()));
    }
}

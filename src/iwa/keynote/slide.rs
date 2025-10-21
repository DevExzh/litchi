//! Keynote Slide Structure
//!
//! Slides are the core content units in Keynote presentations.

use crate::iwa::text::TextStorage;

/// Represents a slide in a Keynote presentation
#[derive(Debug, Clone)]
pub struct KeynoteSlide {
    /// Slide index (0-based)
    pub index: usize,
    /// Slide title
    pub title: Option<String>,
    /// Text content on the slide (bullet points, text boxes)
    pub text_content: Vec<String>,
    /// Speaker notes associated with the slide
    pub notes: Option<String>,
    /// Text storages in this slide
    pub text_storages: Vec<TextStorage>,
    /// Build animations on this slide
    pub builds: Vec<BuildAnimation>,
    /// Slide transition
    pub transition: Option<SlideTransition>,
    /// Master slide reference
    pub master_slide_id: Option<u64>,
}

impl KeynoteSlide {
    /// Create a new slide
    pub fn new(index: usize) -> Self {
        Self {
            index,
            title: None,
            text_content: Vec::new(),
            notes: None,
            text_storages: Vec::new(),
            builds: Vec::new(),
            transition: None,
            master_slide_id: None,
        }
    }

    /// Get all text from the slide (title + content + notes)
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref title) = self.title {
            all.push(title.clone());
        }
        all.extend(self.text_content.clone());
        if let Some(ref notes) = self.notes {
            all.push(notes.clone());
        }

        // Include text from storages
        for storage in &self.text_storages {
            let text = storage.plain_text();
            if !text.is_empty() {
                all.push(text.to_string());
            }
        }

        all
    }

    /// Get plain text content as a single string
    pub fn plain_text(&self) -> String {
        self.all_text().join("\n")
    }

    /// Check if slide is empty
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.text_content.is_empty()
            && self.notes.is_none()
            && self.text_storages.is_empty()
    }

    /// Get number of build animations
    pub fn build_count(&self) -> usize {
        self.builds.len()
    }
}

/// Represents a build animation on a slide
#[derive(Debug, Clone)]
pub struct BuildAnimation {
    /// Animation type
    pub animation_type: BuildAnimationType,
    /// Target object reference
    pub target_id: Option<u64>,
    /// Animation duration (in seconds)
    pub duration: f32,
}

/// Types of build animations
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuildAnimationType {
    /// Appear
    Appear,
    /// Dissolve
    Dissolve,
    /// Move in
    MoveIn,
    /// Scale
    Scale,
    /// Fade and scale
    FadeAndScale,
    /// Other/unknown
    Other,
}

impl BuildAnimationType {
    /// Get a human-readable name
    pub fn name(&self) -> &'static str {
        match self {
            Self::Appear => "Appear",
            Self::Dissolve => "Dissolve",
            Self::MoveIn => "Move In",
            Self::Scale => "Scale",
            Self::FadeAndScale => "Fade and Scale",
            Self::Other => "Other",
        }
    }
}

/// Represents a slide transition effect
#[derive(Debug, Clone)]
pub struct SlideTransition {
    /// Transition type
    pub transition_type: TransitionType,
    /// Transition duration (in seconds)
    pub duration: f32,
}

/// Types of slide transitions
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransitionType {
    /// No transition
    None,
    /// Dissolve
    Dissolve,
    /// Push
    Push,
    /// Wipe
    Wipe,
    /// Flip
    Flip,
    /// Cube
    Cube,
    /// Other/unknown
    Other,
}

impl TransitionType {
    /// Get a human-readable name
    pub fn name(&self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Dissolve => "Dissolve",
            Self::Push => "Push",
            Self::Wipe => "Wipe",
            Self::Flip => "Flip",
            Self::Cube => "Cube",
            Self::Other => "Other",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_creation() {
        let mut slide = KeynoteSlide::new(0);
        assert_eq!(slide.index, 0);
        assert!(slide.is_empty());

        slide.title = Some("Introduction".to_string());
        slide.text_content.push("Point 1".to_string());
        slide.text_content.push("Point 2".to_string());

        assert!(!slide.is_empty());
        let text = slide.plain_text();
        assert!(text.contains("Introduction"));
        assert!(text.contains("Point 1"));
    }

    #[test]
    fn test_slide_all_text() {
        let mut slide = KeynoteSlide::new(0);
        slide.title = Some("Title".to_string());
        slide.text_content.push("Content".to_string());
        slide.notes = Some("Notes".to_string());
        slide
            .text_storages
            .push(TextStorage::from_text("Storage".to_string()));

        let all_text = slide.all_text();
        assert_eq!(all_text.len(), 4);
        assert_eq!(all_text[0], "Title");
        assert_eq!(all_text[1], "Content");
        assert_eq!(all_text[2], "Notes");
        assert_eq!(all_text[3], "Storage");
    }

    #[test]
    fn test_build_animation_type_names() {
        assert_eq!(BuildAnimationType::Appear.name(), "Appear");
        assert_eq!(BuildAnimationType::Dissolve.name(), "Dissolve");
        assert_eq!(BuildAnimationType::MoveIn.name(), "Move In");
    }

    #[test]
    fn test_transition_type_names() {
        assert_eq!(TransitionType::None.name(), "None");
        assert_eq!(TransitionType::Dissolve.name(), "Dissolve");
        assert_eq!(TransitionType::Push.name(), "Push");
    }
}

//! Semantic edits for the in-memory presentation snapshot.

use super::Builder;
use crate::Slide;
use litchi_core::Result;

pub(super) fn titled_slide(index: usize, title: &str, text: &str) -> Slide {
    Slide {
        title: Some(title.to_string()),
        text: text.to_string(),
        index,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: Vec::new(),
    }
}

pub(super) fn text_slide(index: usize, text: &str) -> Slide {
    Slide {
        title: None,
        text: text.to_string(),
        index,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: Vec::new(),
    }
}

impl Builder {
    pub(super) fn append_slide(&mut self, mut slide: Slide) -> Result<&mut Self> {
        slide.index = self.slides.len();
        self.slides.push(slide);
        Ok(self)
    }
}

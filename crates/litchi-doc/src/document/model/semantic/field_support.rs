use super::prelude::*;

impl Document {
    pub(super) fn field_story_text(
        &self,
        story: FieldStory,
        start: u32,
        end: u32,
    ) -> Result<String> {
        if start > end {
            return Err(PackageError::Corrupted(
                "field text range has its start after its end".to_string(),
            ));
        }

        let (story_start, story_end) = self.field_story_range(story)?;
        let start = story_start.checked_add(start).ok_or_else(|| {
            PackageError::Corrupted("field text range start overflows".to_string())
        })?;
        let end = story_start
            .checked_add(end)
            .ok_or_else(|| PackageError::Corrupted("field text range end overflows".to_string()))?;
        if end > story_end {
            return Err(PackageError::Corrupted(
                "field text range exceeds its document story".to_string(),
            ));
        }

        Ok(self.text_extractor.text_at_range(start, end).to_string())
    }

    pub(super) fn field_story_range_if_present(&self, story: FieldStory) -> Option<(u32, u32)> {
        story.range(&self.fib)
    }

    pub(super) fn field_story_range(&self, story: FieldStory) -> Result<(u32, u32)> {
        let range = self.field_story_range_if_present(story);
        range.ok_or_else(|| {
            PackageError::Corrupted(format!(
                "field table refers to absent {} story",
                match story {
                    FieldStory::Main => "main document",
                    FieldStory::Header => "header/footer",
                    FieldStory::Footnote => "footnote",
                    FieldStory::Comment => "comment",
                    FieldStory::Endnote => "endnote",
                    FieldStory::Textbox => "textbox",
                    FieldStory::HeaderTextbox => "header textbox",
                }
            ))
        })
    }
}

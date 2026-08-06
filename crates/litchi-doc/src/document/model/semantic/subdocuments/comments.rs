use super::super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Comments
    // ──────────────────────────────────────────────────────────────────

    /// Get all comments in main-document reference order.
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let mut result = Vec::with_capacity(self.comments_table.count());
        for reference in self.comments_table.references() {
            let reference_end = reference.reference_cp.checked_add(1).ok_or_else(|| {
                PackageError::Corrupted("comment reference CP overflows".to_string())
            })?;
            let marker_end = reference.marker_cp.checked_add(1).ok_or_else(|| {
                PackageError::Corrupted("comment marker CP overflows".to_string())
            })?;
            if self
                .text_extractor
                .text_at_range(reference.reference_cp, reference_end)
                != "\u{5}"
                || self
                    .text_extractor
                    .text_at_range(reference.marker_cp, marker_end)
                    != "\u{5}"
            {
                return Err(PackageError::Corrupted(
                    "comment reference or story does not begin with U+0005".to_string(),
                ));
            }
            if let Some(chp_table) = &self.chp_bin_table
                && (!chp_table
                    .runs_in_range(reference.reference_cp, reference_end)
                    .any(|run| run.properties.is_spec)
                    || !chp_table
                        .runs_in_range(reference.marker_cp, marker_end)
                        .any(|run| run.properties.is_spec))
            {
                return Err(PackageError::Corrupted(
                    "comment reference or story marker is missing sprmCFSpec".to_string(),
                ));
            }

            let body_start = reference.marker_cp.checked_add(1).ok_or_else(|| {
                PackageError::Corrupted("comment body start CP overflows".to_string())
            })?;
            let paragraph_mark_cp = reference.text_end_cp.checked_sub(1).ok_or_else(|| {
                PackageError::Corrupted("comment story range is empty".to_string())
            })?;
            if self
                .text_extractor
                .text_at_range(paragraph_mark_cp, reference.text_end_cp)
                != "\r"
            {
                return Err(PackageError::Corrupted(
                    "comment story does not end with a paragraph mark".to_string(),
                ));
            }
            let text = self
                .text_extractor
                .text_at_range(body_start, reference.text_end_cp)
                .to_string();
            let paragraphs =
                self.extract_paragraphs_for_range(body_start, reference.text_end_cp)?;
            let mut comment = Comment::new(
                reference.reference_cp,
                reference.author.clone(),
                reference.descriptor.initials.clone(),
                reference.descriptor.bookmark_tag,
                text,
            );
            comment.range_start = reference.range_start_cp;
            comment.range_end = reference.range_end_cp;
            comment.extended_metadata = reference.extended_metadata;
            comment.paragraphs = paragraphs;
            result.push(comment);
        }
        Ok(result)
    }
}

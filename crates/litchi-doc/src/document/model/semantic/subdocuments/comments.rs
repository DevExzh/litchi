use super::super::prelude::*;
use crate::parts::comments::CommentReference;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Comments
    // ──────────────────────────────────────────────────────────────────

    /// Get all comments in main-document reference order.
    ///
    /// The first malformed comment fails the whole call;
    /// [`Self::comment`] reads one comment at a time, so a malformed comment
    /// fails only its own read.
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let references = self.0.comments_table.references();
        let mut result = Vec::with_capacity(references.len());
        for reference in references {
            result.push(self.build_comment(reference)?);
        }
        Ok(result)
    }

    /// The number of comments, in main-document reference order.
    #[must_use]
    pub fn comment_count(&self) -> usize {
        self.0.comments_table.count()
    }

    /// The comment at zero-based `index` in main-document reference order,
    /// built and checked alone: the value [`Self::comments`] returns at that
    /// index. A malformed comment fails only its own read, so every other
    /// comment stays readable.
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// for index in 0..document.comment_count() {
    ///     match document.comment(Position::new(index)) {
    ///         Ok(Some(comment)) => println!("{}: {}", comment.author, comment.text()),
    ///         Ok(None) => {},
    ///         Err(error) => eprintln!("comment {index} is unreadable: {error}"),
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Corrupted`] when this comment's reference or
    /// story does not begin with the U+0005 annotation mark carrying
    /// `sprmCFSpec`, or its story does not end with a paragraph mark, and the
    /// errors [`Self::paragraphs`] returns for its paragraphs.
    pub fn comment(&self, index: Position) -> Result<Option<Comment>> {
        self.0
            .comments_table
            .references()
            .get(index.get())
            .map(|reference| self.build_comment(reference))
            .transpose()
    }

    fn build_comment(&self, reference: &CommentReference) -> Result<Comment> {
        let reference_end = reference
            .reference_cp
            .checked_add(1)
            .ok_or_else(|| PackageError::Corrupted("comment reference CP overflows".to_string()))?;
        let marker_end = reference
            .marker_cp
            .checked_add(1)
            .ok_or_else(|| PackageError::Corrupted("comment marker CP overflows".to_string()))?;
        if self
            .0
            .text_extractor
            .text_at_range(reference.reference_cp, reference_end)
            != "\u{5}"
            || self
                .0
                .text_extractor
                .text_at_range(reference.marker_cp, marker_end)
                != "\u{5}"
        {
            return Err(PackageError::Corrupted(
                "comment reference or story does not begin with U+0005".to_string(),
            ));
        }
        if let Some(chp_table) = &self.0.chp_bin_table
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
        let paragraph_mark_cp = reference
            .text_end_cp
            .checked_sub(1)
            .ok_or_else(|| PackageError::Corrupted("comment story range is empty".to_string()))?;
        if self
            .0
            .text_extractor
            .text_at_range(paragraph_mark_cp, reference.text_end_cp)
            != "\r"
        {
            return Err(PackageError::Corrupted(
                "comment story does not end with a paragraph mark".to_string(),
            ));
        }
        let text = self
            .0
            .text_extractor
            .text_at_range(body_start, reference.text_end_cp)
            .to_string();
        let paragraphs = self.extract_paragraphs_for_range(body_start, reference.text_end_cp)?;
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
        Ok(comment)
    }
}

//! Borrowed queries over the ordinary main-story body.

use super::prelude::*;

/// A borrowed, allocation-free traversal of ordinary main-story paragraphs.
///
/// The iterator follows the same source-body rules as
/// [`crate::Package::body_paragraphs`]: only carriage returns outside
/// table PAPX runs emit paragraphs, cell markers reset the source boundary,
/// and an unterminated trailing fragment is not emitted. Each item borrows the
/// decoded text owned by the source [`Document`].
pub struct BodyParagraphs<'a> {
    text: &'a str,
    pap_bin_table: Option<&'a crate::parts::pap_bin_table::PapBinTable>,
    story_end: u32,
    current_cp: u32,
    cursor_byte: usize,
    paragraph_start_byte: usize,
    finished: bool,
}

impl<'a> BodyParagraphs<'a> {
    fn new(document: &'a Document) -> Self {
        let (_, story_end) = document.0.fib.get_main_doc_range();
        Self {
            text: document.0.text_extractor.text(),
            pap_bin_table: document.0.pap_bin_table.as_ref(),
            story_end,
            current_cp: 0,
            cursor_byte: 0,
            paragraph_start_byte: 0,
            finished: false,
        }
    }

    fn error(&mut self, message: &'static str) -> Option<Result<&'a str>> {
        self.finished = true;
        Some(Err(PackageError::Corrupted(message.to_string())))
    }
}

impl<'a> Iterator for BodyParagraphs<'a> {
    type Item = Result<&'a str>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }

        loop {
            if self.current_cp >= self.story_end {
                self.finished = true;
                return None;
            }

            let Some(character) = self.text[self.cursor_byte..].chars().next() else {
                return self
                    .error("decoded document text has an inconsistent main-story CP length");
            };
            let byte = self.cursor_byte;
            self.cursor_byte += character.len_utf8();

            let width = match character.len_utf16() {
                1 => 1,
                2 => 2,
                _ => return self.error("DOC character width does not fit in a CP"),
            };
            let cp = self.current_cp;
            let Some(next_cp) = cp.checked_add(width) else {
                return self.error("DOC main-story CP length overflowed");
            };
            if next_cp > self.story_end {
                return self.error("DOC main-story range splits a decoded character");
            }

            self.current_cp = next_cp;
            match character {
                '\r' => {
                    let in_table = self
                        .pap_bin_table
                        .and_then(|table| table.properties_at(cp))
                        .is_some_and(|properties| properties.in_table);
                    let start = self.paragraph_start_byte;
                    self.paragraph_start_byte = self.cursor_byte;
                    if in_table {
                        continue;
                    }
                    return Some(Ok(&self.text[start..byte]));
                },
                '\u{7}' => {
                    // A table-cell marker closes the current cell and never
                    // contributes an ordinary body paragraph.
                    self.paragraph_start_byte = self.cursor_byte;
                },
                _ => {},
            }
        }
    }
}

impl Document {
    /// Return an allocation-free traversal of ordinary main-story paragraphs.
    #[must_use]
    pub fn body_paragraphs(&self) -> BodyParagraphs<'_> {
        BodyParagraphs::new(self)
    }

    /// Return the number of ordinary paragraphs in the main story.
    ///
    /// A paragraph is emitted at a main-story carriage return whose PAPX is
    /// outside a table. Table-cell paragraph marks and cell marks are skipped,
    /// exactly as they are by [`crate::Package::body_paragraphs`]. A
    /// trailing unterminated text fragment is not a paragraph in the DOC
    /// source model.
    pub fn body_paragraph_count(&self) -> Result<usize> {
        self.body_paragraphs().try_fold(0usize, |count, paragraph| {
            paragraph?;
            count.checked_add(1).ok_or_else(|| {
                PackageError::Corrupted("DOC paragraph count overflowed".to_string())
            })
        })
    }

    /// Return one ordinary main-story paragraph by its zero-based position.
    ///
    /// The returned string is a slice of the decoded document text owned by
    /// this [`Document`]. No paragraph collection or text copy is created;
    /// `Ok(None)` means the requested ordinary paragraph does not exist.
    pub fn body_paragraph_text(&self, position: Position) -> Result<Option<&str>> {
        self.body_paragraphs().nth(position.get()).transpose()
    }
}

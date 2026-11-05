use super::super::prelude::*;
use crate::cp_range::CpRange;
use crate::parts::headers::{HeaderFooterStory, HeaderFooterType};

/// Stories per section in `PlcfHdd` ([MS-DOC] §2.3.3).
const STORIES_PER_SECTION: usize = 6;

/// One header or footer slot of a section ([MS-DOC] §2.3.3).
///
/// Every section has six slots: an even-page, odd-page and first-page header
/// and footer. An empty slot uses the story of the same kind of the previous
/// section; in the first section an empty slot has no header or footer.
/// [`Self::source_section`] resolves that inheritance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderFooterSlot {
    section: usize,
    kind: HeaderFooterType,
    cp_range: Option<CpRange>,
    source_section: Option<usize>,
}

impl HeaderFooterSlot {
    /// The zero-based section the slot belongs to.
    #[must_use]
    pub const fn section(&self) -> Position {
        Position::new(self.section)
    }

    /// Which header or footer the slot holds.
    #[must_use]
    pub const fn kind(&self) -> HeaderFooterType {
        self.kind
    }

    /// Whether the slot's own story is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.cp_range.is_none()
    }

    /// The character positions of the slot's own story, `None` when it is
    /// empty. A non-empty story ends with its content's paragraph mark and a
    /// guard paragraph mark.
    #[must_use]
    pub const fn cp_range(&self) -> Option<CpRange> {
        self.cp_range
    }

    /// The section whose story of this kind applies to this section: this
    /// one when its own story is non-empty, otherwise the nearest earlier
    /// section with a non-empty story of this kind, and `None` when no
    /// section up to this one has one.
    #[must_use]
    pub const fn source_section(&self) -> Option<Position> {
        match self.source_section {
            Some(section) => Some(Position::new(section)),
            None => None,
        }
    }

    /// Whether the slot is empty and uses an earlier section's story.
    #[must_use]
    pub const fn is_inherited(&self) -> bool {
        match self.source_section {
            Some(source) => source != self.section,
            None => false,
        }
    }
}

/// The position of `kind` within one section's `PlcfHdd` group.
const fn storage_slot(kind: HeaderFooterType) -> usize {
    match kind {
        HeaderFooterType::EvenPageHeader => 0,
        HeaderFooterType::OddPageHeader => 1,
        HeaderFooterType::EvenPageFooter => 2,
        HeaderFooterType::OddPageFooter => 3,
        HeaderFooterType::FirstPageHeader => 4,
        HeaderFooterType::FirstPageFooter => 5,
    }
}

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Headers / Footers
    // ──────────────────────────────────────────────────────────────────

    /// Get all headers and footers in the document.
    ///
    /// Each section can have up to six stories: first-page header/footer,
    /// even-page header/footer, and odd-page (default) header/footer.
    /// Empty stories (where `start_cp` == `end_cp`) are omitted, so the
    /// section each one belongs to is lost; [`Self::header_footer_slots`]
    /// keeps it.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for hf in doc.headers_footers()? {
    ///     println!("{:?}: {}", hf.header_footer_type, hf.text());
    /// }
    /// ```
    pub fn headers_footers(&self) -> Result<Vec<HeaderFooter>> {
        let table = match &self.0.headers_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::new();
        for story in table.stories() {
            if story.is_empty() {
                continue;
            }
            result.push(self.build_header_footer(story)?);
        }

        Ok(result)
    }

    /// Get only headers (filtering out footers).
    pub fn headers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(HeaderFooter::is_header)
            .collect())
    }

    /// Get only footers (filtering out headers).
    pub fn footers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(HeaderFooter::is_footer)
            .collect())
    }

    /// The number of sections the header and footer table (`PlcfHdd`)
    /// describes, six slots each; 0 when the document has no header
    /// document.
    #[must_use]
    pub fn header_footer_section_count(&self) -> usize {
        self.0
            .headers_table
            .as_ref()
            .map_or(0, |table| table.stories().len() / STORIES_PER_SECTION)
    }

    /// Every header and footer slot, empty ones included: by section, and
    /// within a section in [`HeaderFooterType::all_types`] order.
    ///
    /// Reading slots builds no story.
    ///
    /// ```no_run
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// for slot in document.header_footer_slots() {
    ///     if let Some(header) = document.section_header_footer(slot.section(), slot.kind())? {
    ///         println!("section {} {:?}: {:?}", slot.section().get(), slot.kind(), header.text());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn header_footer_slots(&self) -> Vec<HeaderFooterSlot> {
        let sections = self.header_footer_section_count();
        let mut slots = Vec::with_capacity(sections.saturating_mul(STORIES_PER_SECTION));
        // The nearest section so far with a non-empty story of each kind, so
        // the listing stays linear in the section count.
        let mut sources = [None; STORIES_PER_SECTION];
        for section in 0..sections {
            for &kind in HeaderFooterType::all_types() {
                let Some(story) = self.slot_story(section, kind) else {
                    continue;
                };
                let source = &mut sources[storage_slot(kind)];
                if !story.is_empty() {
                    *source = Some(section);
                }
                slots.push(Self::slot_from_story(section, kind, story, *source));
            }
        }
        slots
    }

    /// The `kind` slot of section `section`, or `None` past the last section.
    ///
    /// Resolving an empty slot's source scans back through the earlier
    /// sections; [`Self::header_footer_slots`] resolves every slot in one
    /// pass.
    #[must_use]
    pub fn header_footer_slot(
        &self,
        section: Position,
        kind: HeaderFooterType,
    ) -> Option<HeaderFooterSlot> {
        let section = section.get();
        let story = self.slot_story(section, kind)?;
        let source_section = (0..=section).rev().find(|&candidate| {
            self.slot_story(candidate, kind)
                .is_some_and(|story| !story.is_empty())
        });
        Some(Self::slot_from_story(section, kind, story, source_section))
    }

    fn slot_from_story(
        section: usize,
        kind: HeaderFooterType,
        story: &HeaderFooterStory,
        source_section: Option<usize>,
    ) -> HeaderFooterSlot {
        HeaderFooterSlot {
            section,
            kind,
            cp_range: (!story.is_empty())
                .then(|| CpRange::new(story.start_cp, story.end_cp))
                .flatten(),
            source_section,
        }
    }

    /// The header or footer of `kind` that applies to section `section`:
    /// the section's own story, or the earlier section's story an empty slot
    /// inherits ([`HeaderFooterSlot::source_section`]). `None` past the last
    /// section and when no header or footer of this kind applies.
    ///
    /// The value is the one [`Self::headers_footers`] returns for that story.
    ///
    /// # Errors
    ///
    /// Returns the errors [`Self::headers_footers`] returns for the story.
    pub fn section_header_footer(
        &self,
        section: Position,
        kind: HeaderFooterType,
    ) -> Result<Option<HeaderFooter>> {
        let Some(source) = self
            .header_footer_slot(section, kind)
            .and_then(|slot| slot.source_section)
        else {
            return Ok(None);
        };
        match self.slot_story(source, kind) {
            Some(story) => self.build_header_footer(story).map(Some),
            None => Ok(None),
        }
    }

    /// The `PlcfHdd` story of slot `kind` of section `section`.
    fn slot_story(&self, section: usize, kind: HeaderFooterType) -> Option<&HeaderFooterStory> {
        let index = section
            .checked_mul(STORIES_PER_SECTION)?
            .checked_add(storage_slot(kind))?;
        let story = self.0.headers_table.as_ref()?.stories().get(index)?;
        debug_assert_eq!(story.story_type, kind, "PlcfHdd slot order");
        Some(story)
    }

    fn build_header_footer(&self, story: &HeaderFooterStory) -> Result<HeaderFooter> {
        let text = self
            .0
            .text_extractor
            .text_at_range(story.start_cp, story.end_cp)
            .to_string();
        let mut hf = HeaderFooter::new(story.story_type, text);
        hf.paragraphs = self.extract_paragraphs_for_range(story.start_cp, story.end_cp)?;
        Ok(hf)
    }
}

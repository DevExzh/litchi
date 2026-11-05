//! Source provenance capture for destination text stories.

use std::ops::Range;

use super::Parser;
use crate::lexer::{ControlWord, Token};
use crate::model::provenance::StorySource;

/// Extend an optional token range so it covers `origin..end`.
pub(super) fn extend(range: &mut Option<Range<usize>>, origin: usize, end: usize) {
    let start = origin.min(end);
    match range {
        Some(current) => {
            current.start = current.start.min(start);
            current.end = current.end.max(end);
        },
        None => *range = Some(start..end),
    }
}

impl Parser<'_> {
    /// Translate a token range into the source byte range it covers.
    fn token_bytes(&self, tokens: &Range<usize>) -> Option<Range<usize>> {
        if tokens.start >= tokens.end {
            return None;
        }
        let spans = self.token_spans?;
        let first = spans.get(tokens.start)?;
        let last = spans.get(tokens.end.checked_sub(1)?)?;
        (first.start <= last.end).then_some(first.start..last.end)
    }

    /// Source byte offset where the token at `index` begins.
    fn token_start(&self, index: usize) -> Option<usize> {
        self.token_spans?.get(index).map(|span| span.start)
    }

    /// Record that the text appended to the active top-level table cell came
    /// from the tokens dispatched since [`Parser::text_origin`].
    pub(super) fn observe_cell_text(&mut self) {
        extend(
            &mut self.current_cell_text_tokens,
            self.text_origin,
            self.pos,
        );
    }

    /// Record the source of the top-level cell about to be added to the row.
    pub(super) fn record_cell_source(&mut self, table: usize, row: usize, cell: usize) {
        let text = self
            .current_cell_text_tokens
            .take()
            .and_then(|tokens| self.token_bytes(&tokens));
        // `\cell` and `\row` are dispatched after the parser advanced past
        // them; the implicit cell finalized by other boundaries has no
        // trustworthy terminator.
        let terminator = self.pos.checked_sub(1).and_then(|index| {
            matches!(
                self.tokens.get(index),
                Some(Token::Control(
                    ControlWord::TableCell | ControlWord::TableRow
                ))
            )
            .then(|| self.token_start(index))
            .flatten()
        });
        self.text_provenance
            .push_cell((table, row, cell), StorySource { text, terminator });
    }

    /// Record that header/footer paragraph text came from `origin..pos`.
    pub(super) fn observe_header_footer_text(&mut self, origin: usize) {
        extend(&mut self.current_hf_text_tokens, origin, self.pos);
    }

    /// Close the active header/footer paragraph's provenance; `terminator`
    /// is the token index of the paragraph mark or closing brace.
    pub(super) fn close_header_footer_paragraph_source(&mut self, terminator: Option<usize>) {
        let text = self
            .current_hf_text_tokens
            .take()
            .and_then(|tokens| self.token_bytes(&tokens));
        let terminator = terminator.and_then(|index| self.token_start(index));
        self.current_hf_paragraph_sources
            .push(StorySource { text, terminator });
    }

    /// Publish the active header/footer story's paragraph provenance under
    /// its final section and story position.
    pub(super) fn record_header_footer_sources(
        &mut self,
        section: usize,
        story: usize,
        paragraphs: usize,
    ) {
        let mut sources = std::mem::take(&mut self.current_hf_paragraph_sources);
        sources.resize(paragraphs, StorySource::default());
        for (paragraph, source) in sources.into_iter().enumerate() {
            self.text_provenance
                .push_header_footer_paragraph((section, story, paragraph), source);
        }
        self.current_hf_text_tokens = None;
    }

    /// Record that note body text came from `origin..pos`.
    pub(super) fn observe_note_text(&mut self, origin: usize) {
        extend(&mut self.current_note_text_tokens, origin, self.pos);
    }

    /// Record the source of the note about to be appended; `terminator` is
    /// the token index of the note's closing brace.
    pub(super) fn record_note_source(&mut self, index: usize, terminator: Option<usize>) {
        let text = self
            .current_note_text_tokens
            .take()
            .and_then(|tokens| self.token_bytes(&tokens));
        let terminator = terminator.and_then(|index| self.token_start(index));
        self.text_provenance
            .push_note(index, StorySource { text, terminator });
    }

    /// Record the source of the annotation about to be appended.
    pub(super) fn record_annotation_source(
        &mut self,
        index: usize,
        text_tokens: Option<Range<usize>>,
        terminator: Option<usize>,
    ) {
        let text = text_tokens.and_then(|tokens| self.token_bytes(&tokens));
        let terminator = terminator.and_then(|index| self.token_start(index));
        self.text_provenance
            .push_annotation(index, StorySource { text, terminator });
    }
}

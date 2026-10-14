//! Literal matching with the find options.

use std::ops::Range;

/// A compiled literal pattern: built once per find, then run over every
/// text group of the scope.
pub(crate) struct Matcher<'p> {
    pattern: &'p str,
    /// The pattern's folded characters, for a case-insensitive find.
    folded: Option<Box<[char]>>,
    whole_word: bool,
}

impl<'p> Matcher<'p> {
    pub(crate) fn new(pattern: &'p str, match_case: bool, whole_word: bool) -> Self {
        Self {
            pattern,
            folded: (!match_case).then(|| pattern.chars().map(fold).collect()),
            whole_word,
        }
    }

    /// Whether `text` may hold a match: the same search without the
    /// whole-word check, so a `false` is exact and a `true` may not be.
    /// Adapters use it to skip containers before decoding them.
    #[cfg(any(feature = "docx", feature = "pptx"))]
    pub(crate) fn may_match(&self, text: &str) -> bool {
        self.candidate(text, 0).is_some()
    }

    /// The leftmost non-overlapping matches in `text`, as byte ranges.
    pub(crate) fn find_iter<'m, 't>(&'m self, text: &'t str) -> Matches<'m, 'p, 't> {
        Matches {
            matcher: self,
            text,
            from: 0,
        }
    }

    /// The first candidate at or after `from`, before the whole-word check.
    fn candidate(&self, text: &str, from: usize) -> Option<Range<usize>> {
        let rest = text.get(from..)?;
        match &self.folded {
            None => rest
                .find(self.pattern)
                .map(|start| from + start..from + start + self.pattern.len()),
            Some(folded) => {
                let first = *folded.first()?;
                for (offset, character) in rest.char_indices() {
                    if fold(character) != first {
                        continue;
                    }
                    let start = from + offset;
                    if let Some(end) = folded_match(text, start, folded) {
                        return Some(start..end);
                    }
                }
                None
            },
        }
    }
}

/// The end of a case-insensitive match of `folded` at `start`, if any.
fn folded_match(text: &str, start: usize, folded: &[char]) -> Option<usize> {
    let mut characters = text.get(start..)?.char_indices();
    for expected in folded {
        let (_, character) = characters.next()?;
        if fold(character) != *expected {
            return None;
        }
    }
    Some(
        characters
            .next()
            .map_or(text.len(), |(offset, _)| start + offset),
    )
}

/// Iterator of [`Matcher::find_iter`].
pub(crate) struct Matches<'m, 'p, 't> {
    matcher: &'m Matcher<'p>,
    text: &'t str,
    from: usize,
}

impl Iterator for Matches<'_, '_, '_> {
    type Item = Range<usize>;

    fn next(&mut self) -> Option<Range<usize>> {
        loop {
            let found = self.matcher.candidate(self.text, self.from)?;
            if !self.matcher.whole_word || is_whole_word(self.text, &found) {
                self.from = found.end;
                return Some(found);
            }
            // Retry one character later: a rejected candidate may overlap
            // an accepted one (`aa` in `baaa ` with whole words).
            let step = self.text[found.start..]
                .chars()
                .next()
                .map_or(1, char::len_utf8);
            self.from = found.start + step;
        }
    }
}

/// The simple, one-to-one lowercase mapping used by a case-insensitive
/// find. A character whose lowercase form is several characters (`İ`)
/// compares as itself; final sigma compares as sigma.
pub(crate) fn fold(character: char) -> char {
    if character.is_ascii() {
        return character.to_ascii_lowercase();
    }
    let mut lower = character.to_lowercase();
    match (lower.next(), lower.next()) {
        (Some('ς'), None) => 'σ',
        (Some(single), None) => single,
        _ => character,
    }
}

/// A word character for `wholeWord`: a letter, a digit or `_`.
fn is_word(character: char) -> bool {
    character.is_alphanumeric() || character == '_'
}

fn is_whole_word(text: &str, found: &Range<usize>) -> bool {
    let before = text[..found.start].chars().next_back();
    let after = text[found.end..].chars().next();
    !before.is_some_and(is_word) && !after.is_some_and(is_word)
}

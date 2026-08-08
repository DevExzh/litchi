//! Context-aware escaping for literal text embedded in Markdown output.

use std::borrow::Cow;

/// Escape literal inline or block text for the CommonMark/GFM-compatible
/// dialect emitted by Litchi.
///
/// Renderer-owned delimiters, HTML tags, and formula delimiters must be added
/// after this function is called. The input is borrowed when no escaping is
/// required.
#[must_use]
pub fn text(input: &str) -> Cow<'_, str> {
    let mut line_start = 0usize;
    let first_special = input.char_indices().find_map(|(index, character)| {
        let special = needs_backslash(input, line_start, index, character);
        if character == '\n' {
            line_start = index + character.len_utf8();
        }
        special.then_some(index)
    });
    let Some(first_special) = first_special else {
        return Cow::Borrowed(input);
    };

    let mut output = String::with_capacity(input.len().saturating_add(8));
    output.push_str(&input[..first_special]);
    line_start = input[..first_special]
        .rfind('\n')
        .map_or(0, |index| index.saturating_add(1));

    for (relative, character) in input[first_special..].char_indices() {
        let index = first_special + relative;
        if needs_backslash(input, line_start, index, character) {
            output.push('\\');
        }
        output.push(character);
        if character == '\n' {
            line_start = index + character.len_utf8();
        }
    }
    Cow::Owned(output)
}

fn needs_backslash(input: &str, line_start: usize, index: usize, character: char) -> bool {
    if matches!(
        character,
        '\\' | '`'
            | '*'
            | '_'
            | '{'
            | '}'
            | '['
            | ']'
            | '<'
            | '>'
            | '&'
            | '#'
            | '+'
            | '-'
            | '!'
            | '|'
            | '~'
            | '='
            | '$'
    ) {
        return true;
    }
    matches!(character, '.' | ')') && is_ordered_list_delimiter(input, line_start, index)
}

fn is_ordered_list_delimiter(input: &str, line_start: usize, index: usize) -> bool {
    let prefix = &input[line_start..index];
    let indent = prefix.bytes().take_while(|byte| *byte == b' ').count();
    if indent > 3 {
        return false;
    }
    let digits = &prefix[indent..];
    (1..=9).contains(&digits.len())
        && digits.bytes().all(|byte| byte.is_ascii_digit())
        && input[index + 1..]
            .chars()
            .next()
            .is_none_or(char::is_whitespace)
}

#[cfg(test)]
mod tests {
    use super::text;
    use std::borrow::Cow;

    #[test]
    fn borrows_plain_text() {
        assert!(matches!(text("plain text."), Cow::Borrowed(_)));
    }

    #[test]
    fn escapes_inline_and_block_syntax_without_changing_line_breaks() {
        assert_eq!(
            text("# title\n1. item\n> quote\n[a] *b* <tag> \\ | $x$"),
            "\\# title\n1\\. item\n\\> quote\n\\[a\\] \\*b\\* \\<tag\\> \\\\ \\| \\$x\\$"
        );
    }

    #[test]
    fn does_not_escape_periods_outside_ordered_list_starts() {
        assert_eq!(text("Version 1.2\nA 1. item"), "Version 1.2\nA 1. item");
    }

    #[test]
    fn escapes_commonmark_ordered_markers_with_zero_to_three_spaces() {
        assert_eq!(
            text("1. zero\n 2) one\n  003. two\n   42) three"),
            "1\\. zero\n 2\\) one\n  003\\. two\n   42\\) three"
        );
    }

    #[test]
    fn handles_commonmark_digit_limits_and_end_of_line_markers() {
        assert_eq!(
            text("999999999. valid\n999999999)\n1000000000. ordinary\n1000000000) ordinary"),
            "999999999\\. valid\n999999999\\)\n1000000000. ordinary\n1000000000) ordinary"
        );
    }

    #[test]
    fn leaves_nonmarkers_unchanged_without_corrupting_digits() {
        assert_eq!(
            text("    1. code\nA1. embedded\n1.nospace\n1)nospace\n１２. unicode digits"),
            "    1. code\nA1. embedded\n1.nospace\n1)nospace\n１２. unicode digits"
        );
    }

    #[test]
    fn escapes_marker_and_unicode_literal_content_independently() {
        assert_eq!(
            text("  7) 漢字 *literal* [連結]"),
            "  7\\) 漢字 \\*literal\\* \\[連結\\]"
        );
    }

    #[test]
    fn escapes_ampersands_that_could_become_character_references() {
        assert_eq!(text("AT&T &copy; &#169;"), "AT\\&T \\&copy; \\&\\#169;");
    }
}

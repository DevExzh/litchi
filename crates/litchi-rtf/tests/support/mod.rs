//! Helpers shared by integration tests.

/// Return the first balanced RTF root group.
///
/// A few vendored fixtures contain additional top-level tokens after their
/// root group. Tests can use this helper to exercise the valid root without
/// changing the fixture itself. The scanner is deliberately limited to the
/// first group and treats escaped syntax and explicit nonnegative `\\binN`
/// payloads as opaque. A missing, negative, overflowing, or truncated binary
/// length is rejected because it is outside this helper's supported subset.
pub(crate) fn first_balanced_rtf_root_prefix(source: &[u8]) -> Option<&[u8]> {
    let root_start = source.iter().position(|byte| !byte.is_ascii_whitespace())?;
    if source[root_start] != b'{' {
        return None;
    }

    let mut depth = 0usize;
    let mut index = root_start;
    while index < source.len() {
        match source[index] {
            b'{' => {
                depth = depth.checked_add(1)?;
                index += 1;
            },
            b'}' => {
                if depth == 0 {
                    return None;
                }
                depth -= 1;
                index += 1;
                if depth == 0 {
                    return Some(&source[root_start..index]);
                }
            },
            b'\\' => {
                index += 1;
                let escaped = *source.get(index)?;
                match escaped {
                    // Escaped group delimiters and backslashes are literal
                    // bytes, not structural RTF syntax.
                    b'{' | b'}' | b'\\' => index += 1,
                    // A hex escape consumes the quote and two validated hex
                    // digits.
                    b'\'' => {
                        let hex_start = index.checked_add(1)?;
                        let hex_end = hex_start.checked_add(2)?;
                        let hex = source.get(hex_start..hex_end)?;
                        if !hex.iter().all(u8::is_ascii_hexdigit) {
                            return None;
                        }
                        index = hex_end;
                    },
                    control if control.is_ascii_alphabetic() => {
                        let word_start = index;
                        while source
                            .get(index)
                            .is_some_and(|byte| byte.is_ascii_alphabetic())
                        {
                            index += 1;
                        }

                        let number_start = index;
                        if source.get(index) == Some(&b'-') {
                            index += 1;
                        }
                        let digit_start = index;
                        while source.get(index).is_some_and(|byte| byte.is_ascii_digit()) {
                            index += 1;
                        }

                        let word_end = number_start;
                        if &source[word_start..word_end] == b"bin" {
                            // Keep the helper's binary subset aligned with
                            // the lexer's i32 control-word parameter.
                            if digit_start == index || source.get(number_start) == Some(&b'-') {
                                return None;
                            }
                            let mut payload_length = 0i32;
                            for &digit in &source[digit_start..index] {
                                payload_length = payload_length
                                    .checked_mul(10)?
                                    .checked_add(i32::from(digit - b'0'))?;
                            }
                            let payload_length = usize::try_from(payload_length).ok()?;
                            if source.get(index) == Some(&b' ') {
                                index += 1;
                            }
                            index = index.checked_add(payload_length)?;
                            if index > source.len() {
                                return None;
                            }
                        }
                    },
                    // These are the control symbols accepted by the
                    // production lexer and consume their symbol byte.
                    b'\n' | b'\r' | b'*' | b'~' | b'-' | b'_' => index += 1,
                    // Unknown control symbols are not ordinary text. Reject
                    // them instead of allowing malformed structure through.
                    _ => return None,
                }
            },
            _ => index += 1,
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use super::first_balanced_rtf_root_prefix;

    #[test]
    fn returns_first_root_without_leading_whitespace() {
        let source = b" \r\n{\\rtf1 Body} trailing";
        assert_eq!(
            first_balanced_rtf_root_prefix(source),
            Some(b"{\\rtf1 Body}".as_slice())
        );
    }

    #[test]
    fn skips_escaped_braces_and_backslashes() {
        let source = br"{\rtf1 \{literal\} \\literal}";
        assert_eq!(
            first_balanced_rtf_root_prefix(source),
            Some(source.as_slice())
        );
    }

    #[test]
    fn accepts_valid_hex_and_rejects_invalid_hex() {
        let valid = br"{\rtf1\'7b}";
        assert_eq!(
            first_balanced_rtf_root_prefix(valid),
            Some(valid.as_slice())
        );

        let non_hex = br"{\rtf1\'7g}";
        assert_eq!(first_balanced_rtf_root_prefix(non_hex), None);

        let truncated = br"{\rtf1\'7}";
        assert_eq!(first_balanced_rtf_root_prefix(truncated), None);
    }

    #[test]
    fn skips_binary_payload_braces() {
        let source = br"{\rtf1\bin4 {\}X}";
        assert_eq!(
            first_balanced_rtf_root_prefix(source),
            Some(source.as_slice())
        );
    }

    #[test]
    fn rejects_missing_negative_overflow_and_truncated_binary_lengths() {
        let missing = br"{\rtf1\bin}";
        assert_eq!(first_balanced_rtf_root_prefix(missing), None);

        let negative = br"{\rtf1\bin-1}";
        assert_eq!(first_balanced_rtf_root_prefix(negative), None);

        let overflow = br"{\rtf1\bin2147483648}";
        assert_eq!(first_balanced_rtf_root_prefix(overflow), None);

        let truncated = br"{\rtf1\bin2147483647 x}";
        assert_eq!(first_balanced_rtf_root_prefix(truncated), None);
    }

    #[test]
    fn rejects_unknown_control_symbols() {
        let source = br"{\rtf1\@}";
        assert_eq!(first_balanced_rtf_root_prefix(source), None);
    }
}

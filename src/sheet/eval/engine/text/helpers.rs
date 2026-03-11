use crate::sheet::CellValue;
use crate::sheet::eval::engine::to_number;

pub(crate) fn to_non_negative_int(value: &CellValue) -> Option<usize> {
    to_number(value).and_then(|n| {
        if n >= 0.0 {
            Some(n.trunc() as usize)
        } else {
            None
        }
    })
}

pub(crate) fn to_positive_int(value: &CellValue) -> Option<usize> {
    to_number(value).and_then(|n| {
        if n > 0.0 {
            Some(n.trunc() as usize)
        } else {
            None
        }
    })
}

pub(crate) fn take_left(s: &str, count: usize) -> String {
    s.chars().take(count).collect()
}

pub(crate) fn take_right(s: &str, count: usize) -> String {
    let chars: Vec<char> = s.chars().collect();
    let len = chars.len();
    let start = len.saturating_sub(count);
    chars[start..].iter().collect()
}

pub(crate) fn take_mid(s: &str, start_num: usize, count: usize) -> String {
    if count == 0 {
        return String::new();
    }
    let chars: Vec<char> = s.chars().collect();
    if start_num == 0 || start_num > chars.len() {
        return String::new();
    }
    let start_idx = start_num - 1;
    let end_idx = (start_idx + count).min(chars.len());
    chars[start_idx..end_idx].iter().collect()
}

fn char_byte_width(ch: char) -> usize {
    if ch as u32 <= 0xFF { 1 } else { 2 }
}

pub(crate) fn dbcs_byte_len(s: &str) -> usize {
    s.chars().map(char_byte_width).sum()
}

fn slice_by_bytes(s: &str, start_byte: usize, byte_count: usize) -> String {
    if byte_count == 0 || start_byte == 0 {
        return String::new();
    }
    let mut byte_pos = 0;
    let mut start_idx: Option<usize> = None;
    let mut taken = 0;
    let mut end_idx = s.len();

    for (idx, ch) in s.char_indices() {
        let width = char_byte_width(ch);
        let next_pos = byte_pos + width;

        if start_idx.is_none() && next_pos >= start_byte {
            start_idx = Some(idx);
        }

        if start_idx.is_some() {
            if taken + width > byte_count {
                end_idx = idx;
                break;
            }
            taken += width;
        }

        byte_pos = next_pos;
    }

    if start_idx.is_none() || taken == 0 {
        return String::new();
    }
    let start_idx = start_idx.unwrap();
    s[start_idx..end_idx].to_string()
}

pub(crate) fn take_left_bytes(s: &str, byte_count: usize) -> String {
    if byte_count == 0 {
        return String::new();
    }
    slice_by_bytes(s, 1, byte_count)
}

pub(crate) fn take_right_bytes(s: &str, byte_count: usize) -> String {
    if byte_count == 0 {
        return String::new();
    }
    let total = dbcs_byte_len(s);
    if byte_count >= total {
        return s.to_string();
    }
    let start_byte = total - byte_count + 1;
    slice_by_bytes(s, start_byte, byte_count)
}

pub(crate) fn take_mid_bytes(s: &str, start_byte: usize, byte_count: usize) -> String {
    if byte_count == 0 {
        return String::new();
    }
    slice_by_bytes(s, start_byte, byte_count)
}

pub(crate) fn dbcs_byte_prefixes(s: &str) -> Vec<usize> {
    let mut prefixes = Vec::with_capacity(s.chars().count() + 1);
    let mut total = 0;
    prefixes.push(0);
    for ch in s.chars() {
        total += char_byte_width(ch);
        prefixes.push(total);
    }
    prefixes
}

pub(crate) fn char_index_from_dbcs_byte(prefixes: &[usize], start_byte: usize) -> Option<usize> {
    if start_byte == 0 {
        return None;
    }
    let total = *prefixes.last().unwrap_or(&0);
    if start_byte > total {
        return None;
    }
    let target = start_byte - 1;
    (0..prefixes.len().saturating_sub(1)).find(|&i| prefixes[i + 1] > target)
}

pub(crate) fn replace_chars_segment(
    s: &str,
    start_num: usize,
    num_chars: usize,
    replacement: &str,
) -> Option<String> {
    if start_num == 0 {
        return None;
    }
    let chars: Vec<char> = s.chars().collect();
    if start_num > chars.len() {
        return None;
    }
    let start_idx = start_num - 1;
    let end_idx = (start_idx + num_chars).min(chars.len());
    let mut out = String::new();
    for ch in &chars[..start_idx] {
        out.push(*ch);
    }
    out.push_str(replacement);
    for ch in &chars[end_idx..] {
        out.push(*ch);
    }
    Some(out)
}

pub(crate) fn replace_bytes_segment(
    s: &str,
    start_byte: usize,
    num_bytes: usize,
    replacement: &str,
) -> Option<String> {
    if start_byte == 0 {
        return None;
    }
    let chars: Vec<char> = s.chars().collect();
    let widths: Vec<usize> = chars.iter().map(|&ch| char_byte_width(ch)).collect();
    let total_bytes: usize = widths.iter().sum();
    if start_byte > total_bytes {
        return None;
    }
    if chars.is_empty() {
        return Some(replacement.to_string());
    }

    let mut bytes_seen = 0;
    let mut start_idx = 0;
    for (idx, width) in widths.iter().enumerate() {
        let next = bytes_seen + width;
        if start_byte <= next {
            start_idx = idx;
            break;
        }
        bytes_seen = next;
    }

    let mut end_idx = start_idx;
    let mut removed = 0;
    if num_bytes == 0 {
        end_idx = start_idx;
    } else {
        while end_idx < chars.len() && removed < num_bytes {
            removed += widths[end_idx];
            end_idx += 1;
        }
    }

    let mut out = String::new();
    for ch in &chars[..start_idx] {
        out.push(*ch);
    }
    out.push_str(replacement);
    for ch in &chars[end_idx..] {
        out.push(*ch);
    }
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    // ===== to_non_negative_int tests =====

    #[test]
    fn test_to_non_negative_int_with_positive() {
        assert_eq!(to_non_negative_int(&CellValue::Int(5)), Some(5));
        assert_eq!(to_non_negative_int(&CellValue::Float(5.7)), Some(5));
    }

    #[test]
    fn test_to_non_negative_int_with_zero() {
        assert_eq!(to_non_negative_int(&CellValue::Int(0)), Some(0));
        assert_eq!(to_non_negative_int(&CellValue::Float(0.0)), Some(0));
    }

    #[test]
    fn test_to_non_negative_int_with_negative() {
        assert_eq!(to_non_negative_int(&CellValue::Int(-5)), None);
        assert_eq!(to_non_negative_int(&CellValue::Float(-1.5)), None);
    }

    #[test]
    fn test_to_non_negative_int_with_non_numeric() {
        assert_eq!(
            to_non_negative_int(&CellValue::String("abc".to_string())),
            None
        );
        assert_eq!(to_non_negative_int(&CellValue::Bool(true)), None);
    }

    // ===== to_positive_int tests =====

    #[test]
    fn test_to_positive_int_with_positive() {
        assert_eq!(to_positive_int(&CellValue::Int(5)), Some(5));
        assert_eq!(to_positive_int(&CellValue::Float(5.7)), Some(5));
    }

    #[test]
    fn test_to_positive_int_with_zero() {
        assert_eq!(to_positive_int(&CellValue::Int(0)), None);
        assert_eq!(to_positive_int(&CellValue::Float(0.0)), None);
    }

    #[test]
    fn test_to_positive_int_with_negative() {
        assert_eq!(to_positive_int(&CellValue::Int(-5)), None);
        assert_eq!(to_positive_int(&CellValue::Float(-1.5)), None);
    }

    // ===== take_left tests =====

    #[test]
    fn test_take_left_basic() {
        assert_eq!(take_left("Hello", 2), "He");
        assert_eq!(take_left("Hello", 5), "Hello");
    }

    #[test]
    fn test_take_left_zero() {
        assert_eq!(take_left("Hello", 0), "");
    }

    #[test]
    fn test_take_left_more_than_length() {
        assert_eq!(take_left("Hello", 10), "Hello");
    }

    #[test]
    fn test_take_left_unicode() {
        // take_left takes characters, not bytes
        assert_eq!(take_left("Hello 世界", 8), "Hello 世界");
        assert_eq!(take_left("Hello 世界", 7), "Hello 世");
    }

    // ===== take_right tests =====

    #[test]
    fn test_take_right_basic() {
        assert_eq!(take_right("Hello", 2), "lo");
        assert_eq!(take_right("Hello", 5), "Hello");
    }

    #[test]
    fn test_take_right_zero() {
        assert_eq!(take_right("Hello", 0), "");
    }

    #[test]
    fn test_take_right_more_than_length() {
        assert_eq!(take_right("Hello", 10), "Hello");
    }

    #[test]
    fn test_take_right_unicode() {
        assert_eq!(take_right("Hello 世界", 3), " 世界");
    }

    // ===== take_mid tests =====

    #[test]
    fn test_take_mid_basic() {
        assert_eq!(take_mid("Hello World", 7, 5), "World");
        assert_eq!(take_mid("Hello", 2, 2), "el");
    }

    #[test]
    fn test_take_mid_zero_count() {
        assert_eq!(take_mid("Hello", 1, 0), "");
    }

    #[test]
    fn test_take_mid_start_zero() {
        assert_eq!(take_mid("Hello", 0, 3), "");
    }

    #[test]
    fn test_take_mid_start_beyond_length() {
        assert_eq!(take_mid("Hello", 10, 3), "");
    }

    #[test]
    fn test_take_mid_partial() {
        assert_eq!(take_mid("Hello", 4, 5), "lo");
    }

    #[test]
    fn test_take_mid_unicode() {
        assert_eq!(take_mid("Hello 世界", 7, 2), "世界");
    }

    // ===== dbcs_byte_len tests =====

    #[test]
    fn test_dbcs_byte_len_ascii() {
        assert_eq!(dbcs_byte_len("Hello"), 5);
    }

    #[test]
    fn test_dbcs_byte_len_unicode() {
        // Characters > 0xFF count as 2 bytes
        assert_eq!(dbcs_byte_len("世界"), 4);
        // "Hello 世界" has: H(1)+e(1)+l(1)+l(1)+o(1)+ (1)+世(2)+界(2) = 10 bytes
        assert_eq!(dbcs_byte_len("Hello 世界"), 10);
    }

    #[test]
    fn test_dbcs_byte_len_empty() {
        assert_eq!(dbcs_byte_len(""), 0);
    }

    // ===== take_left_bytes tests =====

    #[test]
    fn test_take_left_bytes_basic() {
        assert_eq!(take_left_bytes("Hello", 3), "Hel");
    }

    #[test]
    fn test_take_left_bytes_zero() {
        assert_eq!(take_left_bytes("Hello", 0), "");
    }

    #[test]
    fn test_take_left_bytes_more_than_length() {
        assert_eq!(take_left_bytes("Hello", 10), "Hello");
    }

    #[test]
    fn test_take_left_bytes_unicode() {
        // "世" is 2 bytes, "界" is 2 bytes
        assert_eq!(take_left_bytes("世界", 2), "世");
        assert_eq!(take_left_bytes("世界", 4), "世界");
        assert_eq!(take_left_bytes("Hello 世界", 8), "Hello 世");
    }

    // ===== take_right_bytes tests =====

    #[test]
    fn test_take_right_bytes_basic() {
        assert_eq!(take_right_bytes("Hello", 3), "llo");
    }

    #[test]
    fn test_take_right_bytes_zero() {
        assert_eq!(take_right_bytes("Hello", 0), "");
    }

    #[test]
    fn test_take_right_bytes_more_than_length() {
        assert_eq!(take_right_bytes("Hello", 10), "Hello");
    }

    #[test]
    fn test_take_right_bytes_unicode() {
        assert_eq!(take_right_bytes("世界", 2), "界");
        assert_eq!(take_right_bytes("Hello 世界", 5), " 世界");
    }

    // ===== take_mid_bytes tests =====

    #[test]
    fn test_take_mid_bytes_basic() {
        assert_eq!(take_mid_bytes("Hello", 2, 3), "ell");
    }

    #[test]
    fn test_take_mid_bytes_zero_count() {
        assert_eq!(take_mid_bytes("Hello", 1, 0), "");
    }

    #[test]
    fn test_take_mid_bytes_unicode() {
        // "Hello 世界" has bytes: H(1)e(1)l(1)l(1)o(1) (1)世(2)界(2)
        // Byte positions: 1  2  3  4  5  6  7   9   11
        assert_eq!(take_mid_bytes("Hello 世界", 7, 2), "世");
        assert_eq!(take_mid_bytes("Hello 世界", 7, 4), "世界");
    }

    // ===== dbcs_byte_prefixes tests =====

    #[test]
    fn test_dbcs_byte_prefixes_ascii() {
        assert_eq!(dbcs_byte_prefixes("ABC"), vec![0, 1, 2, 3]);
    }

    #[test]
    fn test_dbcs_byte_prefixes_unicode() {
        // "A世B" - A(1) + 世(2) + B(1) = 4 bytes
        assert_eq!(dbcs_byte_prefixes("A世B"), vec![0, 1, 3, 4]);
    }

    #[test]
    fn test_dbcs_byte_prefixes_empty() {
        assert_eq!(dbcs_byte_prefixes(""), vec![0]);
    }

    // ===== char_index_from_dbcs_byte tests =====

    #[test]
    fn test_char_index_from_dbcs_byte_basic() {
        let prefixes = vec![0, 1, 2, 3];
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 1), Some(0));
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 2), Some(1));
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 3), Some(2));
    }

    #[test]
    fn test_char_index_from_dbcs_byte_unicode() {
        // "A世B" - prefixes [0, 1, 3, 4]
        let prefixes = vec![0, 1, 3, 4];
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 1), Some(0)); // A
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 2), Some(1)); // within 世
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 3), Some(1)); // 世 end
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 4), Some(2)); // B
    }

    #[test]
    fn test_char_index_from_dbcs_byte_zero() {
        let prefixes = vec![0, 1, 2, 3];
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 0), None);
    }

    #[test]
    fn test_char_index_from_dbcs_byte_beyond() {
        let prefixes = vec![0, 1, 2, 3];
        assert_eq!(char_index_from_dbcs_byte(&prefixes, 5), None);
    }

    // ===== replace_chars_segment tests =====

    #[test]
    fn test_replace_chars_segment_basic() {
        assert_eq!(
            replace_chars_segment("Hello World", 7, 5, "Universe"),
            Some("Hello Universe".to_string())
        );
    }

    #[test]
    fn test_replace_chars_segment_start_zero() {
        assert_eq!(replace_chars_segment("Hello", 0, 2, "XX"), None);
    }

    #[test]
    fn test_replace_chars_segment_start_beyond() {
        assert_eq!(replace_chars_segment("Hello", 10, 2, "XX"), None);
    }

    #[test]
    fn test_replace_chars_segment_zero_chars() {
        assert_eq!(
            replace_chars_segment("Hello", 3, 0, "XX"),
            Some("HeXXllo".to_string())
        );
    }

    #[test]
    fn test_replace_chars_segment_more_chars() {
        // start_num=4 removes chars[3]=l, and 10 chars from position 4 removes "lo"
        assert_eq!(
            replace_chars_segment("Hello", 4, 10, "!"),
            Some("Hel!".to_string())
        );
    }

    // ===== replace_bytes_segment tests =====

    #[test]
    fn test_replace_bytes_segment_basic() {
        // Each char is 1 byte for ASCII, start_byte=7 (position of 'W'), remove 5 bytes ("World")
        assert_eq!(
            replace_bytes_segment("Hello World", 7, 5, "XYZ"),
            Some("Hello XYZ".to_string())
        );
    }

    #[test]
    fn test_replace_bytes_segment_start_zero() {
        assert_eq!(replace_bytes_segment("Hello", 0, 2, "XX"), None);
    }

    #[test]
    fn test_replace_bytes_segment_start_beyond() {
        assert_eq!(replace_bytes_segment("Hello", 10, 2, "XX"), None);
    }

    #[test]
    fn test_replace_bytes_segment_zero_bytes() {
        assert_eq!(
            replace_bytes_segment("Hello", 3, 0, "XX"),
            Some("HeXXllo".to_string())
        );
    }

    #[test]
    fn test_replace_bytes_segment_empty_string() {
        // Empty string with start_byte=1 returns None (start_byte > total_bytes)
        assert_eq!(replace_bytes_segment("", 1, 0, "XX"), None);
    }

    #[test]
    fn test_replace_bytes_segment_unicode() {
        // "A世C" - A(1 byte), 世(2 bytes), C(1 byte)
        // Start at byte 2 (within 世), remove 1 byte
        let result = replace_bytes_segment("A世C", 2, 1, "X");
        assert!(result.is_some());
    }
}

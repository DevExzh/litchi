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

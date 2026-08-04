//! Safe URI and archive-path classification shared by ODF package owners.

use litchi_core::{Error, Result};

/// Return whether an href must remain an inert linked reference.
pub fn is_linked_href(href: &str) -> bool {
    if href.starts_with('/')
        || href.starts_with('\\')
        || href.starts_with('#')
        || href.contains('\\')
        || href.contains('?')
        || href.contains('#')
    {
        return true;
    }
    let Some(colon) = href.find(':') else {
        return false;
    };
    let scheme = &href[..colon];
    !scheme.is_empty()
        && scheme.as_bytes()[0].is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
}

/// Resolve a relative ODF package href into a safe archive path.
pub fn resolve_package_path(href: &str) -> Result<String> {
    let decoded = percent_decode(href)?;
    if decoded.starts_with('/') || decoded.contains('\\') {
        return Err(Error::InvalidFormat(format!(
            "unsafe package href '{href}'"
        )));
    }
    let mut segments = Vec::new();
    for segment in decoded.split('/') {
        if segment.is_empty() || segment == "." {
            continue;
        }
        if segment == ".." {
            if segments.pop().is_none() {
                return Err(Error::InvalidFormat(format!(
                    "package href escapes the package root: '{href}'"
                )));
            }
            continue;
        }
        if segment
            .chars()
            .any(|character| character == '\0' || character.is_control())
        {
            return Err(Error::InvalidFormat(format!(
                "invalid character in package href '{href}'"
            )));
        }
        segments.push(segment);
    }
    if segments.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "package href has no file path: '{href}'"
        )));
    }
    let path = segments.join("/");
    if path == "mimetype" || path == "META-INF" || path.starts_with("META-INF/") {
        return Err(Error::InvalidFormat(format!(
            "package href targets an administrative entry: '{href}'"
        )));
    }
    Ok(path)
}

fn percent_decode(value: &str) -> Result<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Err(Error::InvalidFormat(format!(
                "invalid percent escape in package href '{value}'"
            )));
        }
        let high = hex_value(bytes[index + 1]);
        let low = hex_value(bytes[index + 2]);
        let (Some(high), Some(low)) = (high, low) else {
            return Err(Error::InvalidFormat(format!(
                "invalid percent escape in package href '{value}'"
            )));
        };
        decoded.push((high << 4) | low);
        index += 3;
    }
    String::from_utf8(decoded)
        .map_err(|_| Error::InvalidFormat("package href is not valid UTF-8".to_string()))
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::{is_linked_href, resolve_package_path};

    #[test]
    fn classifies_non_package_links_without_fetching() {
        assert!(is_linked_href("https://example.invalid/image.png"));
        assert!(is_linked_href("#fragment"));
        assert!(is_linked_href("Pictures/image.png?cache=1"));
        assert!(!is_linked_href("Pictures/image.png"));
    }

    #[test]
    fn normalizes_safe_relative_paths() {
        assert_eq!(
            resolve_package_path("./Pictures/../Pictures/image%20one.png").unwrap(),
            "Pictures/image one.png"
        );
        assert!(resolve_package_path("../../outside.png").is_err());
        assert!(resolve_package_path("META-INF/manifest.xml").is_err());
    }
}

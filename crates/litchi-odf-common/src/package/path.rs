//! Safe URI and archive-path classification shared by ODF package owners.

use litchi_core::{Error, Result};

/// Return whether an href must remain an inert linked reference.
#[must_use]
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
///
/// # Errors
///
/// Returns an error when the href is not a safe, non-administrative relative
/// package path or contains an invalid percent escape.
pub fn resolve_package_path(href: &str) -> Result<String> {
    let path = resolve_relative_package_path(href)?;
    if path == "mimetype" || path == "META-INF" || path.starts_with("META-INF/") {
        return Err(Error::InvalidFormat(format!(
            "package href targets an administrative entry: '{href}'"
        )));
    }
    Ok(path)
}

/// Validate a manifest `full-path` without applying aliases.
///
/// Manifest paths are the package's security metadata keys.  They must use
/// the same spelling as the archive member that callers later request; in
/// particular, accepting `./content.xml`, `foo/../content.xml`, or a percent
/// encoded alias would allow a requested canonical member to miss its
/// encryption descriptor.  The ODF root entry (`/`) is the one intentional
/// absolute path.  Directory entries retain their required trailing slash.
pub(crate) fn validate_manifest_path(path: &str) -> Result<()> {
    if path == "/" {
        return Ok(());
    }
    if path.contains(['?', '#', ':']) {
        return Err(Error::InvalidFormat(format!(
            "unsafe ODF manifest full-path '{path}'"
        )));
    }
    let body = path.strip_suffix('/').unwrap_or(path);
    if body.is_empty() {
        return Err(Error::InvalidFormat(
            "ODF manifest directory entry has an empty path".to_string(),
        ));
    }
    let canonical = resolve_relative_package_path(body)?;
    let expected = if path.ends_with('/') {
        let mut expected = canonical;
        expected.push('/');
        expected
    } else {
        canonical
    };
    if expected != path {
        return Err(Error::InvalidFormat(format!(
            "non-canonical ODF manifest full-path '{path}'"
        )));
    }
    Ok(())
}

fn resolve_relative_package_path(href: &str) -> Result<String> {
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
    Ok(segments.join("/"))
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
        let high_nibble_value = hex_value(bytes[index + 1]);
        let low_nibble_value = hex_value(bytes[index + 2]);
        let (Some(high_nibble), Some(low_nibble)) = (high_nibble_value, low_nibble_value) else {
            return Err(Error::InvalidFormat(format!(
                "invalid percent escape in package href '{value}'"
            )));
        };
        decoded.push((high_nibble << 4) | low_nibble);
        index += 3;
    }
    String::from_utf8(decoded)
        .map_err(|_utf8_error| Error::InvalidFormat("package href is not valid UTF-8".to_string()))
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
    use super::{is_linked_href, resolve_package_path, validate_manifest_path};

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
            resolve_package_path("./Pictures/../Pictures/image%20one.png")
                .unwrap_or_else(|error| panic!("safe package path must resolve: {error}")),
            "Pictures/image one.png"
        );
        assert!(resolve_package_path("../../outside.png").is_err());
        assert!(resolve_package_path("META-INF/manifest.xml").is_err());
    }

    #[test]
    fn manifest_paths_keep_only_the_root_as_absolute_and_allow_directories() {
        assert!(validate_manifest_path("/").is_ok());
        assert!(validate_manifest_path("Pictures/").is_ok());
        assert!(validate_manifest_path("META-INF/documentsignatures.xml").is_ok());

        for alias in [
            "/content.xml",
            "./content.xml",
            "foo/../content.xml",
            "Pictures//image.png",
            "content%2Exml",
            "content\\.xml",
            "content.xml?cache=1",
            "content.xml#fragment",
            "C:content.xml",
            "foo:content.xml",
        ] {
            assert!(
                validate_manifest_path(alias).is_err(),
                "manifest alias must be rejected: {alias}"
            );
        }
    }
}

/// Provides the PackURI value type and utilities for working with package URIs.
///
/// A PackURI represents a part name within an OPC package, following the URI format
/// defined by the Open Packaging Conventions specification.
/// Represents a package URI, which is a partname within an OPC package.
///
/// PackURIs always begin with a forward slash and use forward slashes as path separators,
/// following the OPC specification. They provide access to various components like
/// the base URI (directory), filename, extension, and index.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PackURI {
    /// The full pack URI string (e.g., "/word/document.xml")
    uri: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PartNameConflict {
    Duplicate,
    Equivalent,
    Derived,
}

impl PackURI {
    /// Create a new PackURI from a string.
    ///
    /// # Arguments
    /// * `uri` - The URI string, which must begin with a forward slash
    ///
    /// # Returns
    /// * `Ok(PackURI)` if the URI is valid
    /// * `Err` if the URI is not a canonical, fragment-free package path
    pub fn new<S: Into<String>>(uri: S) -> Result<Self, String> {
        let uri = uri.into();
        if !uri.starts_with('/') {
            return Err(format!("PackURI must begin with slash, got '{}'", uri));
        }
        if uri == "/" {
            return Ok(PackURI { uri });
        }
        if uri == CONTENT_TYPES_URI {
            return Ok(PackURI { uri });
        }
        if uri.ends_with('/') {
            return Err(format!("PackURI must not end with slash, got '{uri}'"));
        }
        if uri.contains("//") {
            return Err(format!(
                "PackURI must not contain empty segments, got '{uri}'"
            ));
        }
        if uri.contains('\\') {
            return Err(format!("PackURI must use forward slashes, got '{uri}'"));
        }
        if uri.contains(['?', '#']) {
            return Err(format!(
                "PackURI must not contain a query or fragment, got '{uri}'"
            ));
        }
        if uri
            .chars()
            .any(|character| character.is_control() || character == ' ')
        {
            return Err(format!(
                "PackURI spaces and control characters must be percent-encoded, got '{uri}'"
            ));
        }
        if uri
            .split('/')
            .skip(1)
            .any(|part| matches!(part, "." | ".."))
        {
            return Err(format!(
                "PackURI must not contain dot segments, got '{uri}'"
            ));
        }
        Self::validate_percent_encoding(&uri)?;
        for segment in uri.split('/').skip(1) {
            if segment.ends_with('.') {
                return Err(format!(
                    "PackURI segments must not end with a dot, got '{uri}'"
                ));
            }
            if !segment.as_bytes().iter().any(|byte| *byte != b'.') {
                return Err(format!(
                    "PackURI segments must contain a non-dot character, got '{uri}'"
                ));
            }
        }
        Ok(PackURI { uri })
    }

    /// Create a PackURI from a relative reference and a base URI.
    ///
    /// This translates a relative reference (like "../styles.xml") onto a base URI
    /// (like "/word") to produce an absolute PackURI (like "/styles.xml").
    ///
    /// # Arguments
    /// * `base_uri` - The base URI to resolve from
    /// * `relative_ref` - The relative reference to resolve
    pub fn from_rel_ref(base_uri: &str, relative_ref: &str) -> Result<Self, String> {
        if relative_ref.is_empty() {
            return Err("Relationship target must not be empty".to_string());
        }
        if relative_ref.contains('\\') {
            return Err(format!(
                "Relationship target must use forward slashes, got '{relative_ref}'"
            ));
        }
        if relative_ref.contains(['?', '#']) {
            return Err(format!(
                "Internal relationship target must not contain a query or fragment, got '{relative_ref}'"
            ));
        }
        if relative_ref
            .split('/')
            .next()
            .is_some_and(|first| first.contains(':'))
        {
            return Err(format!(
                "Internal relationship target must not be an absolute URI, got '{relative_ref}'"
            ));
        }
        // Join the paths using POSIX-style path manipulation
        let joined = if relative_ref.starts_with('/') {
            relative_ref.to_string()
        } else {
            Self::join_paths(base_uri, relative_ref)
        };
        let normalized = Self::normalize_path(&joined)?;
        Self::new(normalized)
    }

    /// Get the base URI (directory portion) of this PackURI.
    ///
    /// For example, "/ppt/slides" for "/ppt/slides/slide1.xml".
    /// For the package pseudo-partname "/", returns "/".
    pub fn base_uri(&self) -> &str {
        if self.uri == "/" {
            return "/";
        }

        if let Some(pos) = self.uri.rfind('/') {
            if pos == 0 { "/" } else { &self.uri[..pos] }
        } else {
            "/"
        }
    }

    /// Get the filename portion of this PackURI.
    ///
    /// For example, "slide1.xml" for "/ppt/slides/slide1.xml".
    /// For the package pseudo-partname "/", returns an empty string.
    pub fn filename(&self) -> &str {
        if let Some(pos) = self.uri.rfind('/') {
            &self.uri[pos + 1..]
        } else {
            ""
        }
    }

    /// Get the extension portion of this PackURI.
    ///
    /// For example, "xml" for "/word/document.xml" (note: no leading period).
    pub fn ext(&self) -> &str {
        let filename = self.filename();
        if let Some(pos) = filename.rfind('.') {
            &filename[pos + 1..]
        } else {
            ""
        }
    }

    /// Get the partname index for tuple partnames, or None for singleton partnames.
    ///
    /// For example, returns 21 for "/ppt/slides/slide21.xml" and None for "/ppt/presentation.xml".
    pub fn idx(&self) -> Option<u32> {
        let filename = self.filename();
        if filename.is_empty() {
            return None;
        }

        // Remove extension to get the name part
        let name_part = if let Some(pos) = filename.rfind('.') {
            &filename[..pos]
        } else {
            filename
        };

        // Try to extract numeric suffix (e.g., "slide21" -> 21)
        let mut digit_start = None;
        for (i, c) in name_part.chars().enumerate() {
            if c.is_ascii_digit() {
                if digit_start.is_none() {
                    digit_start = Some(i);
                }
            } else if digit_start.is_some() {
                // Reset if we encounter a non-digit after digits
                digit_start = None;
            }
        }

        // Parse the numeric suffix if found
        if let Some(start) = digit_start
            && start > 0
            && start < name_part.len()
        {
            return name_part[start..].parse::<u32>().ok();
        }

        None
    }

    /// Get the membername (URI with leading slash stripped).
    ///
    /// This is the form used as the Zip file membername for the package item.
    /// Returns an empty string for the package pseudo-partname "/".
    pub fn membername(&self) -> &str {
        if self.uri == "/" { "" } else { &self.uri[1..] }
    }

    /// Get the relative reference from a base URI to this PackURI.
    ///
    /// For example, PackURI("/ppt/slideLayouts/slideLayout1.xml") would return
    /// "../slideLayouts/slideLayout1.xml" for base_uri "/ppt/slides".
    pub fn relative_ref(&self, base_uri: &str) -> String {
        // Special case for root base URI
        if base_uri == "/" {
            return self.membername().to_string();
        }

        // Calculate relative path
        let from_parts: Vec<&str> = base_uri.split('/').filter(|s| !s.is_empty()).collect();
        let to_parts: Vec<&str> = self.uri.split('/').filter(|s| !s.is_empty()).collect();

        // Find common prefix length
        let common = from_parts
            .iter()
            .zip(to_parts.iter())
            .take_while(|(a, b)| a == b)
            .count();

        // Build relative path
        let mut result = String::new();

        // Add "../" for each part in from_parts after common prefix
        for _ in common..from_parts.len() {
            result.push_str("../");
        }

        // Add parts from to_parts after common prefix
        for (i, part) in to_parts.iter().enumerate().skip(common) {
            if i > common {
                result.push('/');
            }
            result.push_str(part);
        }

        result
    }

    /// Get the PackURI of the .rels part corresponding to this PackURI.
    ///
    /// For example, "/word/_rels/document.xml.rels" for "/word/document.xml".
    pub fn rels_uri(&self) -> Result<PackURI, String> {
        let filename = self.filename();
        let base_uri = self.base_uri();

        let rels_filename = format!("{}.rels", filename);
        let rels_uri_str = if base_uri == "/" {
            format!("/_rels/{}", rels_filename)
        } else {
            format!("{}/_rels/{}", base_uri, rels_filename)
        };

        Self::new(rels_uri_str)
    }

    /// Get the full URI string.
    pub fn as_str(&self) -> &str {
        &self.uri
    }

    /// Whether two part names are equivalent under OPC's ASCII case-insensitive rule.
    pub fn is_equivalent_to(&self, other: &Self) -> bool {
        self.uri.eq_ignore_ascii_case(&other.uri)
    }

    pub(crate) fn conflict_with(&self, other: &Self) -> Option<PartNameConflict> {
        if self.uri == other.uri {
            return Some(PartNameConflict::Duplicate);
        }
        if self.is_equivalent_to(other) {
            return Some(PartNameConflict::Equivalent);
        }
        if is_derived_name(&self.uri, &other.uri) || is_derived_name(&other.uri, &self.uri) {
            return Some(PartNameConflict::Derived);
        }
        None
    }

    /// Helper function to join two paths using forward slashes
    fn join_paths(base: &str, rel: &str) -> String {
        if base.ends_with('/') {
            format!("{}{}", base, rel)
        } else {
            format!("{}/{}", base, rel)
        }
    }

    /// Helper function to normalize a path (resolve ".." and ".")
    fn normalize_path(path: &str) -> Result<String, String> {
        let mut parts = Vec::new();

        for part in path.split('/') {
            match part {
                "" | "." => {
                    // Skip empty and current directory markers
                    if parts.is_empty() {
                        // Keep leading slash
                        parts.push("");
                    }
                },
                ".." => {
                    // Go up one directory
                    if parts.len() > 1 {
                        parts.pop();
                    } else {
                        return Err(format!(
                            "Relationship target escapes the package root: '{path}'"
                        ));
                    }
                },
                _ => {
                    parts.push(part);
                },
            }
        }

        // Handle root case
        if parts.is_empty() || (parts.len() == 1 && parts[0].is_empty()) {
            return Ok("/".to_string());
        }

        Ok(parts.join("/"))
    }

    fn validate_percent_encoding(uri: &str) -> Result<(), String> {
        let bytes = uri.as_bytes();
        let mut index = 0usize;
        while index < bytes.len() {
            if bytes[index] == b'/' {
                index += 1;
                continue;
            }
            if bytes[index] != b'%' {
                if !is_pchar(bytes[index]) {
                    return Err(format!(
                        "PackURI contains a character outside RFC 3986 pchar, got '{uri}'"
                    ));
                }
                index += 1;
                continue;
            }
            let encoded = bytes.get(index + 1..index + 3).ok_or_else(|| {
                format!("PackURI contains incomplete percent encoding, got '{uri}'")
            })?;
            if !encoded.iter().all(u8::is_ascii_hexdigit) {
                return Err(format!(
                    "PackURI contains invalid percent encoding, got '{uri}'"
                ));
            }
            let decoded = (hex_value(encoded[0]) << 4) | hex_value(encoded[1]);
            if matches!(decoded, b'/' | b'\\') {
                return Err(format!(
                    "PackURI must not percent-encode path separators, got '{uri}'"
                ));
            }
            if is_unreserved(decoded) {
                return Err(format!(
                    "PackURI must not percent-encode unreserved characters, got '{uri}'"
                ));
            }
            index += 3;
        }
        Ok(())
    }
}

fn is_unreserved(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~')
}

fn is_pchar(byte: u8) -> bool {
    is_unreserved(byte)
        || matches!(
            byte,
            b'!' | b'$'
                | b'&'
                | b'\''
                | b'('
                | b')'
                | b'*'
                | b'+'
                | b','
                | b';'
                | b'='
                | b':'
                | b'@'
        )
}

fn is_derived_name(candidate: &str, base: &str) -> bool {
    candidate.len() > base.len()
        && candidate
            .get(..base.len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(base))
        && candidate.as_bytes().get(base.len()) == Some(&b'/')
}

fn hex_value(byte: u8) -> u8 {
    match byte {
        b'0'..=b'9' => byte - b'0',
        b'a'..=b'f' => byte - b'a' + 10,
        b'A'..=b'F' => byte - b'A' + 10,
        _ => 0,
    }
}

impl std::fmt::Display for PackURI {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.uri)
    }
}

impl AsRef<str> for PackURI {
    fn as_ref(&self) -> &str {
        &self.uri
    }
}

/// The package pseudo-partname, representing the package itself
pub const PACKAGE_URI: &str = "/";

/// The URI for the [Content_Types].xml part
pub const CONTENT_TYPES_URI: &str = "/[Content_Types].xml";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_packuri_new() {
        assert!(PackURI::new("/word/document.xml").is_ok());
        assert!(PackURI::new("word/document.xml").is_err());
    }

    #[test]
    fn rejects_ambiguous_or_noncanonical_part_names() {
        for invalid in [
            "/word/",
            "/word//document.xml",
            "/word/./document.xml",
            "/word/../document.xml",
            "/word\\document.xml",
            "/word/document.xml?query",
            "/word/document.xml#fragment",
            "/word/my document.xml",
            "/word/%document.xml",
            "/word/%2Fdocument.xml",
            "/word/%5cdocument.xml",
            "/word/%64ocument.xml",
            "/word/document.",
            "/word/.../document.xml",
            "/word/doc[1].xml",
            "/word/doc^1.xml",
            "/word/café.xml",
        ] {
            assert!(PackURI::new(invalid).is_err(), "accepted {invalid}");
        }
        assert!(PackURI::new("/word/my%20document.xml").is_ok());
        assert!(PackURI::new("/word/a!$&'()*+,;=:@.xml").is_ok());
        assert!(PackURI::new("/").is_ok());
    }

    #[test]
    fn detects_equivalent_and_derived_part_names() {
        let document = PackURI::new("/word/document.xml").unwrap();
        let equivalent = PackURI::new("/WORD/DOCUMENT.XML").unwrap();
        let derived = PackURI::new("/word/document.xml/image.gif").unwrap();
        assert!(document.is_equivalent_to(&equivalent));
        assert_eq!(
            document.conflict_with(&equivalent),
            Some(PartNameConflict::Equivalent)
        );
        assert_eq!(
            document.conflict_with(&derived),
            Some(PartNameConflict::Derived)
        );
    }

    #[test]
    fn resolves_internal_targets_without_allowing_root_escape() {
        assert_eq!(
            PackURI::from_rel_ref("/word", "styles/styles.xml")
                .unwrap()
                .as_str(),
            "/word/styles/styles.xml"
        );
        assert_eq!(
            PackURI::from_rel_ref("/word/charts", "../embeddings/data.bin")
                .unwrap()
                .as_str(),
            "/word/embeddings/data.bin"
        );
        assert!(PackURI::from_rel_ref("/", "../escape.xml").is_err());
        assert!(PackURI::from_rel_ref("/word", "https://example.com/a.xml").is_err());
        assert!(PackURI::from_rel_ref("/word", "document.xml#bookmark").is_err());
    }

    #[test]
    fn test_base_uri() {
        let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        assert_eq!(uri.base_uri(), "/ppt/slides");

        let root = PackURI::new("/").unwrap();
        assert_eq!(root.base_uri(), "/");
    }

    #[test]
    fn test_filename() {
        let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        assert_eq!(uri.filename(), "slide1.xml");

        let root = PackURI::new("/").unwrap();
        assert_eq!(root.filename(), "");
    }

    #[test]
    fn test_ext() {
        let uri = PackURI::new("/word/document.xml").unwrap();
        assert_eq!(uri.ext(), "xml");
    }

    #[test]
    fn test_idx() {
        let uri = PackURI::new("/ppt/slides/slide21.xml").unwrap();
        assert_eq!(uri.idx(), Some(21));

        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        assert_eq!(uri.idx(), None);
    }

    #[test]
    fn test_membername() {
        let uri = PackURI::new("/word/document.xml").unwrap();
        assert_eq!(uri.membername(), "word/document.xml");

        let root = PackURI::new("/").unwrap();
        assert_eq!(root.membername(), "");
    }
}

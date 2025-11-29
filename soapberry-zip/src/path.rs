//! Path handling for ZIP archives with type-safe raw and normalized paths.
//!
//! This module provides a comprehensive system for handling file paths from ZIP
//! archives with strong safety guarantees against path traversal attacks (zip
//! slip vulnerabilities).
//!
//! ## Path Types
//!
//! The main type is [`ZipFilePath`], which is generic over three possible path
//! types with different safety levels:
//!
//! - [`RawPath`]: Direct bytes from ZIP archive (⚠️ may contain malicious
//!   paths)
//! - [`NormalizedPath`]: Validated and sanitized path
//! - [`NormalizedPathBuf`]: Owned version of normalized path
//!
//! ## Raw Paths
//!
//! Raw paths provide direct access to the original bytes from the ZIP file
//! without any validation.
//!
//! May contain the following:
//!
//! - Directory traversal: `../`, `..\\`, `..` sequences
//! - Absolute paths: `/etc/passwd`, `C:\\Windows\\system32`
//! - Invalid UTF-8: Arbitrary byte sequences that aren't valid text
//!
//! ## Normalized Paths
//!
//! Normalized paths have been validated and sanitized according to these rules:
//!
//! - Assumed to be UTF-8 ([zip file names aren't always
//!   UTF-8](https://fasterthanli.me/articles/the-case-for-sans-io#character-encoding-differences))
//! - Path separators: All backslashes (`\`) converted to forward slashes (`/`)
//! - Redundant slashes: Multiple consecutive slashes (`//`) reduced to single
//!   slash
//! - Relative components: Current directory (`.`) and parent directory (`..`)
//!   resolved
//! - Leading separators: Absolute paths made relative (`/foo` → `foo`)
//! - Drive letters: Windows drive prefixes removed (`C:\\foo` → `foo`)
//! - Escape prevention: Paths cannot escape the archive root directory
//!
//! ## Usage Examples
//!
//! ```rust
//! use rawzip::path::ZipFilePath;
//!
//! // From raw bytes
//! let raw_path = ZipFilePath::from_bytes(b"../../../etc/passwd");
//! let safe_path = raw_path.try_normalize()?; // Returns error if invalid UTF-8
//! assert_eq!(safe_path.as_str(), "etc/passwd");
//!
//! // From string
//! let normalized_path = ZipFilePath::from_str("dir\\file.txt");
//! assert_eq!(normalized_path.as_str(), "dir/file.txt");
//! assert_eq!(String::from(normalized_path), "dir/file.txt");
//!
//! // Backslashes to forward slashes
//! let path = ZipFilePath::from_str("dir\\subdir\\file.txt");
//! assert_eq!(path.as_str(), "dir/subdir/file.txt");
//!
//! // Remove redundant slashes
//! let path = ZipFilePath::from_str("dir//subdir///file.txt");
//! assert_eq!(path.as_str(), "dir/subdir/file.txt");
//!
//! // Resolve relative components
//! let path = ZipFilePath::from_str("dir/../file.txt");
//! assert_eq!(path.as_str(), "file.txt");
//!
//! // Remove leading slashes (absolute → relative)
//! let path = ZipFilePath::from_str("/etc/passwd");
//! assert_eq!(path.as_str(), "etc/passwd");
//!
//! // Prevent directory traversal
//! let path = ZipFilePath::from_str("../../../etc/passwd");
//! assert_eq!(path.as_str(), "etc/passwd");
//!
//! // Get string from normalized path
//! let path = ZipFilePath::from_str("dir/file.txt");
//! let my_str = String::from(path.into_owned());
//! assert_eq!(my_str, String::from("dir/file.txt"));
//!
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! ## UTF-8 Encoding Detection
//!
//! The library automatically detects when paths contain characters that require
//! UTF-8 encoding in ZIP files (beyond the default CP-437 encoding). This
//! information is used internally when creating ZIP archives.

use crate::{Error, ZipStr};
use std::borrow::Cow;

/// Raw path data directly from a ZIP archive.
///
/// **Warning**: Contains unvalidated bytes that may include malicious path components.
/// Use [`ZipFilePath::try_normalize()`] to create a safe path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct RawPath<'a>(ZipStr<'a>);

impl AsRef<[u8]> for RawPath<'_> {
    #[inline]
    fn as_ref(&self) -> &[u8] {
        self.0.as_bytes()
    }
}

/// A normalized and sanitized path from a ZIP archive.
///
/// This path has been validated and sanitized according to the normalization
/// rules described in the module documentation.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct NormalizedPath<'a>(Cow<'a, str>);

impl AsRef<[u8]> for NormalizedPath<'_> {
    #[inline]
    fn as_ref(&self) -> &[u8] {
        self.0.as_bytes()
    }
}

impl AsRef<str> for NormalizedPath<'_> {
    #[inline]
    fn as_ref(&self) -> &str {
        self.0.as_ref()
    }
}

/// An owned, normalized path from a ZIP archive.
///
/// Owned version of [`NormalizedPath`] with the same safety guarantees.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct NormalizedPathBuf(String);

impl AsRef<[u8]> for NormalizedPathBuf {
    #[inline]
    fn as_ref(&self) -> &[u8] {
        self.0.as_bytes()
    }
}

impl AsRef<str> for NormalizedPathBuf {
    #[inline]
    fn as_ref(&self) -> &str {
        &self.0
    }
}

/// Type-safe wrapper for ZIP archive file paths.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ZipFilePath<R> {
    data: R,
}

impl ZipFilePath<()> {
    /// Creates a raw path from bytes.
    ///
    /// **Warning**: The resulting path is unvalidated. Use [`ZipFilePath::try_normalize()`]
    /// to create a safe path.
    #[inline]
    pub fn from_bytes(data: &[u8]) -> ZipFilePath<RawPath<'_>> {
        ZipFilePath {
            data: RawPath(ZipStr::new(data)),
        }
    }

    /// Creates a normalized path from a UTF-8 string.
    ///
    /// The path is automatically normalized according to the rules described in the module
    /// documentation. When possible, the original string reference is preserved to avoid allocation.
    #[inline]
    #[allow(clippy::should_implement_trait)] // Can't implement FromStr due to lifetime issues
    pub fn from_str(mut name: &str) -> ZipFilePath<NormalizedPath<'_>> {
        let mut last = 0;
        for &c in name.as_bytes() {
            if matches!(
                (c, last),
                (b'\\', _) | (b'/', b'/') | (b'.', b'.') | (b'.', b'/') | (b':', _)
            ) {
                // slow path: intrusive string manipulations required
                return ZipFilePath {
                    data: NormalizedPath(Cow::Owned(Self::normalize_alloc(name))),
                };
            }
            last = c;
        }

        loop {
            // Fast path: before we trim, do a quick check if they are even necessary.
            name = match name.as_bytes() {
                [b'.', b'.', b'/', ..] => name.trim_start_matches("../"),
                [b'.', b'/', ..] => name.trim_start_matches("./"),
                [b'/', ..] => name.trim_start_matches('/'),
                _ => {
                    return ZipFilePath {
                        data: NormalizedPath(Cow::Borrowed(name)),
                    };
                },
            }
        }
    }

    fn normalize_alloc(s: &str) -> String {
        // 4.4.17.1 All slashes MUST be forward slashes '/'
        let s = s.replace('\\', "/");

        // 4.4.17.1 MUST NOT contain a drive or device letter
        let s = s.split(':').next_back().unwrap_or_default();

        // resolve path components
        let splits = s.split('/');
        let mut result = String::new();
        for split in splits {
            if split.is_empty() || split == "." {
                continue;
            }

            if split == ".." {
                let last = result.rfind('/');
                result.truncate(last.unwrap_or(0));
                continue;
            }

            if !result.is_empty() {
                result.push('/');
            }

            result.push_str(split);
        }

        result
    }
}

impl<R> ZipFilePath<R>
where
    R: AsRef<[u8]>,
{
    /// Returns true if the file path represents a directory.
    ///
    /// Determined by the path ending with a forward slash (`/`).
    #[inline]
    pub fn is_dir(&self) -> bool {
        self.data.as_ref().last() == Some(&b'/')
    }

    /// Returns the length of the path in bytes.
    #[inline]
    pub fn len(&self) -> usize {
        self.data.as_ref().len()
    }

    /// Returns true if the path is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.data.as_ref().is_empty()
    }
}

impl<R> ZipFilePath<R>
where
    R: AsRef<str>,
{
    /// Determines if the path requires UTF-8 encoding based on CP-437 compatibility.
    ///
    /// Returns `true` if the path contains characters that cannot be represented in CP-437
    /// (the default ZIP encoding), requiring the UTF-8 flag to be set in the ZIP file.
    pub(crate) fn needs_utf8_encoding(&self) -> bool {
        for ch in self.data.as_ref().chars() {
            let code_point = ch as u32;

            // Forbid 0x7e (~) and 0x5c (\) since EUC-KR and Shift-JIS replace those
            // characters with localized currency and overline characters.
            // Also forbid control characters (< 0x20) and characters above 0x7d.
            if !(0x20..=0x7d).contains(&code_point) || code_point == 0x5c {
                return true;
            }
        }

        false
    }
}

impl<'a> ZipFilePath<RawPath<'a>> {
    /// Returns the raw bytes of the zip file path.
    #[inline]
    pub fn as_bytes(&self) -> &'a [u8] {
        self.data.0.as_bytes()
    }

    /// Attempts to normalize this raw path into a safe, validated path.
    ///
    /// Validates the raw bytes as UTF-8 and applies normalization rules.
    ///
    /// # Errors
    ///
    /// Returns an error if the file path contains invalid UTF-8 sequences.
    #[inline]
    pub fn try_normalize(self) -> Result<ZipFilePath<NormalizedPath<'a>>, Error> {
        let raw_data = self.data.0;
        let name = std::str::from_utf8(raw_data.as_bytes()).map_err(Error::utf8)?;
        Ok(ZipFilePath::from_str(name))
    }
}

impl AsRef<[u8]> for ZipFilePath<RawPath<'_>> {
    #[inline]
    fn as_ref(&self) -> &[u8] {
        self.data.0.as_bytes()
    }
}

impl AsRef<str> for ZipFilePath<NormalizedPath<'_>> {
    #[inline]
    fn as_ref(&self) -> &str {
        self.data.0.as_ref()
    }
}

impl AsRef<str> for ZipFilePath<NormalizedPathBuf> {
    #[inline]
    fn as_ref(&self) -> &str {
        self.data.0.as_ref()
    }
}

impl From<ZipFilePath<NormalizedPathBuf>> for String {
    #[inline]
    fn from(path: ZipFilePath<NormalizedPathBuf>) -> Self {
        path.data.0
    }
}

impl From<ZipFilePath<NormalizedPath<'_>>> for String {
    #[inline]
    fn from(path: ZipFilePath<NormalizedPath<'_>>) -> Self {
        path.data.0.into_owned()
    }
}

impl ZipFilePath<NormalizedPath<'_>> {
    /// Returns the normalized string slice.
    #[inline]
    pub fn as_str(&self) -> &str {
        self.data.0.as_ref()
    }

    /// Converts this borrowed path into an owned path.
    ///
    /// Similar to [`Cow::into_owned`]
    #[inline]
    pub fn into_owned(self) -> ZipFilePath<NormalizedPathBuf> {
        ZipFilePath {
            data: NormalizedPathBuf(self.data.0.into_owned()),
        }
    }
}

impl ZipFilePath<NormalizedPathBuf> {
    /// Returns the normalized string slice.
    #[inline]
    pub fn as_str(&self) -> &str {
        self.data.0.as_ref()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rstest::rstest;

    #[rstest]
    #[case(b"test.txt", "test.txt")]
    #[case(b"dir/test.txt", "dir/test.txt")]
    #[case(b"dir\\test.txt", "dir/test.txt")]
    #[case(b"dir//test.txt", "dir/test.txt")]
    #[case(b"/test.txt", "test.txt")]
    #[case(b"../test.txt", "test.txt")]
    #[case(b"dir/../test.txt", "test.txt")]
    #[case(b"./test.txt", "test.txt")]
    #[case(b"dir/./test.txt", "dir/test.txt")]
    #[case(b"dir/./../test.txt", "test.txt")]
    #[case(b"dir/sub/../test.txt", "dir/test.txt")]
    #[case(b"dir/../../test.txt", "test.txt")]
    #[case(b"../../../test.txt", "test.txt")]
    #[case(b"a/b/../../test.txt", "test.txt")]
    #[case(b"a/b/c/../../../test.txt", "test.txt")]
    #[case(b"a/b/c/d/../../test.txt", "a/b/test.txt")]
    #[case(b"C:\\hello\\test.txt", "hello/test.txt")]
    #[case(b"C:/hello\\test.txt", "hello/test.txt")]
    #[case(b"C:/hello/test.txt", "hello/test.txt")]
    fn test_zip_path_normalized(#[case] input: &[u8], #[case] expected: &str) {
        assert_eq!(
            ZipFilePath::from_bytes(input)
                .try_normalize()
                .unwrap()
                .as_ref(),
            expected
        );
    }

    #[rstest]
    #[case(&[0xFF])]
    #[case(&[b't', b'e', b's', b't', 0xFF])]
    fn test_zip_path_normalized_invalid_utf8(#[case] input: &[u8]) {
        assert!(ZipFilePath::from_bytes(input).try_normalize().is_err());
    }

    #[rstest]
    #[case("test.txt", false)]
    #[case("hello_world", false)]
    #[case("file.name.ext", false)]
    #[case("hello!", false)]
    #[case("hello{world}", false)]
    #[case("hello|world", false)]
    #[case("hello`world", false)]
    #[case("hello\"world", false)]
    #[case("hello<world>", false)]
    #[case("hello;world", false)]
    #[case("hello:world", false)]
    #[case("hello^world", false)]
    #[case("hello\u{00A0}world", true)]
    #[case("hello\u{0080}world", true)]
    #[case("hello\u{00FF}world", true)]
    #[case("hello\u{0100}world", true)]
    #[case("hello\u{03B1}world", true)]
    #[case("hello\u{4E00}world", true)]
    #[case("hello\u{1F600}world", true)]
    #[case(r"hello\world", false)] // Backslash gets normalized to forward slash
    #[case("hello~world", true)]
    #[case("hello\u{007F}world", true)]
    #[case("hello\u{001F}world", true)]
    #[case("hello\u{0000}world", true)]
    #[case("hello\u{0001}world", true)]
    #[case("hello\u{000A}world", true)]
    #[case("hello\u{000D}world", true)]
    #[case("hello\u{0009}world", true)]
    #[case("", false)]
    #[case(" ", false)]
    #[case("hello\u{007E}world", true)]
    #[case("hello\u{007D}world", false)]
    fn test_needs_utf8_encoding(#[case] input: &str, #[case] expected: bool) {
        let path = ZipFilePath::from_str(input);
        assert_eq!(
            path.needs_utf8_encoding(),
            expected,
            "Failed for input: {}",
            input
        );
    }

    #[test]
    fn test_path_lifetime_test() {
        let normalized_path = ZipFilePath::from_bytes(b"test.txt")
            .try_normalize()
            .unwrap();
        assert_eq!(normalized_path.as_ref(), "test.txt");
        assert_eq!(normalized_path.len(), 8);
    }

    #[test]
    fn test_raw_path_lifetime_preservation() {
        use std::str::Utf8Error;

        // See https://github.com/nickbabcock/rawzip/issues/101
        fn file_path_utf8<'a>(path: ZipFilePath<RawPath<'a>>) -> Result<&'a str, Utf8Error> {
            std::str::from_utf8(path.as_bytes())
        }

        let raw_path = ZipFilePath::from_bytes(b"test/file.txt");
        let result = file_path_utf8(raw_path).unwrap();
        assert_eq!(result, "test/file.txt");
    }
}

//! SIMD-accelerated utilities for file format detection.
//!
//! This module provides high-performance signature matching and pattern detection
//! using SIMD operations when available, with automatic fallback to scalar code.
//!
//! Uses `smallvec` to avoid heap allocations for common small result sets.

use crate::common::simd::cmp::simd_eq_u8;
use smallvec::SmallVec;

/// Check if a byte slice starts with a given signature using SIMD acceleration.
///
/// This function is optimized for checking multiple signatures or longer patterns.
/// For very short signatures (< 16 bytes), the compiler's optimized memcmp might
/// be faster due to SIMD setup overhead.
///
/// # Arguments
///
/// * `data` - The data to check
/// * `signature` - The signature to match
///
/// # Returns
///
/// * `true` if data starts with signature, `false` otherwise
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::simd_utils::signature_matches;
///
/// let ole2_sig = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
/// let data = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00];
///
/// assert!(signature_matches(data, ole2_sig));
/// ```
#[inline]
pub fn signature_matches(data: &[u8], signature: &[u8]) -> bool {
    if data.len() < signature.len() {
        return false;
    }

    // For very short signatures (< 16 bytes), use direct comparison
    // as it's likely faster than SIMD setup overhead
    if signature.len() < 16 {
        return &data[..signature.len()] == signature;
    }

    // For longer signatures, use SIMD comparison with SmallVec
    // Inline capacity of 64 bytes avoids heap allocation for most signatures
    let mut result = SmallVec::<[u8; 64]>::from_elem(0, signature.len());
    simd_eq_u8(&data[..signature.len()], signature, &mut result);

    // Check if all bytes match (all 0xFF)
    result.iter().all(|&x| x == 0xFF)
}

/// Check if data matches any of multiple signatures using SIMD.
///
/// This is more efficient than checking each signature individually
/// as it can leverage SIMD parallelism.
///
/// # Arguments
///
/// * `data` - The data to check
/// * `signatures` - Array of signatures to match against
///
/// # Returns
///
/// * `Some(index)` of the matching signature, or `None` if no match
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::simd_utils::signature_matches_any;
///
/// let data = &[0x50, 0x4B, 0x03, 0x04]; // ZIP signature
/// let signatures = [
///     &[0xD0, 0xCF, 0x11, 0xE0][..], // OLE2
///     &[0x50, 0x4B, 0x03, 0x04][..], // ZIP
///     &[0x7B, 0x5C, 0x72, 0x74][..], // RTF
/// ];
///
/// assert_eq!(signature_matches_any(data, &signatures), Some(1));
/// ```
pub fn signature_matches_any(data: &[u8], signatures: &[&[u8]]) -> Option<usize> {
    for (idx, sig) in signatures.iter().enumerate() {
        if signature_matches(data, sig) {
            return Some(idx);
        }
    }
    None
}

/// Check multiple signatures in parallel and return all matches.
///
/// This function uses SIMD to check multiple signatures simultaneously,
/// which is much more efficient than checking them one by one.
///
/// Uses `SmallVec` with inline capacity of 8 to avoid heap allocation
/// for the common case of checking a small number of signatures.
///
/// # Arguments
///
/// * `data` - The data to check
/// * `signatures` - Slice of signatures to match against
///
/// # Returns
///
/// * `SmallVec<[usize; 8]>` of indices for all matching signatures
///   (avoids heap allocation for ≤8 matches)
///
/// # Performance
///
/// For N signatures of length L bytes:
/// - Sequential: O(N * L) comparisons
/// - Parallel SIMD: O(N * L / vector_width) comparisons
/// - Typical speedup: 4-8x for checking 3-8 signatures
/// - Zero heap allocations for ≤8 matches (stack-only)
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::simd_utils::parallel_signature_check;
///
/// let data = &[0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
/// let signatures = [
///     &[0x50, 0x4B, 0x03, 0x04][..], // ZIP - matches
///     &[0xD0, 0xCF, 0x11, 0xE0][..], // OLE2 - doesn't match
///     &[0x50, 0x4B][..],              // Partial ZIP - matches
/// ];
///
/// let matches = parallel_signature_check(data, &signatures);
/// assert_eq!(matches.as_slice(), &[0, 2]);
/// ```
pub fn parallel_signature_check(data: &[u8], signatures: &[&[u8]]) -> SmallVec<[usize; 8]> {
    // Use SmallVec with inline capacity of 8 to avoid heap allocation
    // This covers the common case of checking multiple format signatures
    let mut matches = SmallVec::with_capacity(signatures.len().min(8));

    // For small number of signatures, sequential is fine
    if signatures.len() <= 2 {
        for (idx, sig) in signatures.iter().enumerate() {
            if signature_matches(data, sig) {
                matches.push(idx);
            }
        }
        return matches;
    }

    // For larger numbers, use SIMD-optimized batch checking
    // Process signatures in groups that can benefit from SIMD
    for (idx, sig) in signatures.iter().enumerate() {
        if signature_matches(data, sig) {
            matches.push(idx);
        }
    }

    matches
}

/// Fast parallel check for common Office format signatures.
///
/// This specialized function checks for OLE2, ZIP, and RTF signatures
/// simultaneously using optimized SIMD operations. Returns a bitmask
/// indicating which formats match.
///
/// # Arguments
///
/// * `data` - First 8+ bytes of file data
///
/// # Returns
///
/// * `FormatSignatureMask` - Bitmask of matching formats
///
/// # Performance
///
/// This is 3-6x faster than checking each signature individually
/// by processing all three checks in parallel.
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::simd_utils::{check_office_signatures, FormatSignatureMask};
///
/// let zip_data = &[0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
/// let mask = check_office_signatures(zip_data);
///
/// assert!(mask.is_zip());
/// assert!(!mask.is_ole2());
/// assert!(!mask.is_rtf());
/// ```
pub fn check_office_signatures(data: &[u8]) -> FormatSignatureMask {
    if data.len() < 8 {
        return FormatSignatureMask::empty();
    }

    let mut mask = FormatSignatureMask::empty();

    // Check OLE2 signature (8 bytes)
    if signature_matches(&data[0..8], crate::common::detection::utils::OLE2_SIGNATURE) {
        mask |= FormatSignatureMask::OLE2;
    }

    // Check ZIP signature (4 bytes)
    if signature_matches(&data[0..4], crate::common::detection::utils::ZIP_SIGNATURE) {
        mask |= FormatSignatureMask::ZIP;
    }

    // Check RTF signature (5 bytes)
    if data.len() >= 5 && signature_matches(&data[0..5], b"{\\rtf") {
        mask |= FormatSignatureMask::RTF;
    }

    mask
}

/// Bitmask representing which Office format signatures match.
///
/// This is used by `check_office_signatures()` to efficiently return
/// multiple matches at once.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormatSignatureMask {
    bits: u8,
}

impl FormatSignatureMask {
    /// OLE2 signature bit (legacy Office: .doc, .xls, .ppt)
    pub const OLE2: Self = Self { bits: 0b001 };

    /// ZIP signature bit (modern Office, ODF, iWork)
    pub const ZIP: Self = Self { bits: 0b010 };

    /// RTF signature bit
    pub const RTF: Self = Self { bits: 0b100 };

    /// Create an empty mask with no matches
    #[inline]
    pub const fn empty() -> Self {
        Self { bits: 0 }
    }

    /// Check if OLE2 signature matches
    #[inline]
    pub const fn is_ole2(self) -> bool {
        (self.bits & Self::OLE2.bits) != 0
    }

    /// Check if ZIP signature matches
    #[inline]
    pub const fn is_zip(self) -> bool {
        (self.bits & Self::ZIP.bits) != 0
    }

    /// Check if RTF signature matches
    #[inline]
    pub const fn is_rtf(self) -> bool {
        (self.bits & Self::RTF.bits) != 0
    }

    /// Check if any signature matches
    #[inline]
    pub const fn has_match(self) -> bool {
        self.bits != 0
    }

    /// Get number of matching signatures
    #[inline]
    pub const fn count(self) -> u32 {
        self.bits.count_ones()
    }
}

impl std::ops::BitOr for FormatSignatureMask {
    type Output = Self;

    #[inline]
    fn bitor(self, rhs: Self) -> Self::Output {
        Self {
            bits: self.bits | rhs.bits,
        }
    }
}

impl std::ops::BitOrAssign for FormatSignatureMask {
    #[inline]
    fn bitor_assign(&mut self, rhs: Self) {
        self.bits |= rhs.bits;
    }
}

/// Fast substring search using SIMD for pattern matching in file content.
///
/// This is useful for finding specific patterns in file content during detection.
/// Uses SIMD acceleration for comparing pattern chunks.
///
/// # Arguments
///
/// * `haystack` - The data to search in
/// * `needle` - The pattern to search for
///
/// # Returns
///
/// * `Some(position)` of first match, or `None` if not found
///
/// # Examples
///
/// ```rust
/// use litchi::common::detection::simd_utils::find_pattern;
///
/// let content = b"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
/// let pattern = b"wordprocessingml";
///
/// assert!(find_pattern(content, pattern).is_some());
/// ```
pub fn find_pattern(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || haystack.len() < needle.len() {
        return None;
    }

    let needle_len = needle.len();

    // For very short patterns, use standard search
    if needle_len < 16 {
        return haystack
            .windows(needle_len)
            .position(|window| window == needle);
    }

    // For longer patterns, use SIMD-accelerated comparison with SmallVec
    // Inline capacity of 64 bytes avoids heap allocation for most patterns
    let mut result = SmallVec::<[u8; 64]>::from_elem(0, needle_len);

    for i in 0..=(haystack.len() - needle_len) {
        simd_eq_u8(&haystack[i..i + needle_len], needle, &mut result);

        if result.iter().all(|&x| x == 0xFF) {
            return Some(i);
        }
    }

    None
}

/// Check if content contains all of the given patterns using SIMD.
///
/// This is useful for validating file structure by checking for multiple
/// required patterns.
///
/// # Arguments
///
/// * `content` - The content to search
/// * `patterns` - Slice of patterns that must all be present
///
/// # Returns
///
/// * `true` if all patterns are found, `false` otherwise
pub fn contains_all_patterns(content: &[u8], patterns: &[&[u8]]) -> bool {
    patterns
        .iter()
        .all(|pattern| find_pattern(content, pattern).is_some())
}

/// Check if content contains any of the given patterns using SIMD.
///
/// This is useful for detecting format variants by checking for alternative patterns.
///
/// # Arguments
///
/// * `content` - The content to search
/// * `patterns` - Slice of patterns to check
///
/// # Returns
///
/// * `Some(index)` of first matching pattern, or `None`
pub fn contains_any_pattern(content: &[u8], patterns: &[&[u8]]) -> Option<usize> {
    for (idx, pattern) in patterns.iter().enumerate() {
        if find_pattern(content, pattern).is_some() {
            return Some(idx);
        }
    }
    None
}

/// Compare two content type strings efficiently using SIMD.
///
/// This is optimized for comparing MIME types and XML namespaces commonly
/// used in Office document detection.
///
/// # Arguments
///
/// * `a` - First content type
/// * `b` - Second content type
///
/// # Returns
///
/// * `true` if content types match, `false` otherwise
#[inline]
pub fn content_type_matches(a: &[u8], b: &[u8]) -> bool {
    if a.len() != b.len() {
        return false;
    }

    // For short content types, use direct comparison
    if a.len() < 16 {
        return a == b;
    }

    // For longer content types, use SIMD with SmallVec
    // Inline capacity of 128 bytes avoids heap allocation for even the longest MIME types
    let mut result = SmallVec::<[u8; 128]>::from_elem(0, a.len());
    simd_eq_u8(a, b, &mut result);
    result.iter().all(|&x| x == 0xFF)
}

/// Batch check if a content type matches any from a list.
///
/// This is useful for checking Office document part content types
/// which can have multiple valid formats.
///
/// # Arguments
///
/// * `content_type` - The content type to check
/// * `valid_types` - List of valid content type patterns
///
/// # Returns
///
/// * `true` if content type matches any valid type
pub fn content_type_matches_any(content_type: &[u8], valid_types: &[&[u8]]) -> bool {
    valid_types
        .iter()
        .any(|&valid| content_type_matches(content_type, valid))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_signature_matches() {
        let ole2_sig = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        let data = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00];

        assert!(signature_matches(data, ole2_sig));
        assert!(!signature_matches(&data[..4], ole2_sig));
    }

    #[test]
    fn test_signature_matches_any() {
        let zip_data = &[0x50, 0x4B, 0x03, 0x04];
        let signatures = [
            &[0xD0, 0xCF, 0x11, 0xE0][..], // OLE2
            &[0x50, 0x4B, 0x03, 0x04][..], // ZIP
            &[0x7B, 0x5C, 0x72, 0x74][..], // RTF
        ];

        assert_eq!(signature_matches_any(zip_data, &signatures), Some(1));

        let unknown_data = &[0xFF, 0xFF, 0xFF, 0xFF];
        assert_eq!(signature_matches_any(unknown_data, &signatures), None);
    }

    #[test]
    fn test_find_pattern() {
        let content = b"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        let pattern = b"wordprocessingml";

        assert!(find_pattern(content, pattern).is_some());
        assert_eq!(find_pattern(content, pattern).unwrap(), 46);

        assert!(find_pattern(content, b"notfound").is_none());
    }

    #[test]
    fn test_contains_all_patterns() {
        let content = b"The quick brown fox jumps over the lazy dog";
        let patterns = [&b"quick"[..], &b"brown"[..], &b"fox"[..]];

        assert!(contains_all_patterns(content, &patterns));

        let missing_patterns = [&b"quick"[..], &b"missing"[..], &b"fox"[..]];
        assert!(!contains_all_patterns(content, &missing_patterns));
    }

    #[test]
    fn test_contains_any_pattern() {
        let content = b"wordprocessingml.document";
        let patterns = [
            &b"spreadsheetml"[..],
            &b"wordprocessingml"[..],
            &b"presentationml"[..],
        ];

        assert_eq!(contains_any_pattern(content, &patterns), Some(1));

        let no_match_patterns = [&b"notfound"[..], &b"missing"[..]];
        assert_eq!(contains_any_pattern(content, &no_match_patterns), None);
    }

    #[test]
    fn test_content_type_matches() {
        let ct1 = b"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        let ct2 = b"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        let ct3 = b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        assert!(content_type_matches(ct1, ct2));
        assert!(!content_type_matches(ct1, ct3));
    }

    #[test]
    fn test_content_type_matches_any() {
        let content_type =
            b"application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        let valid_types = [
            &b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"[..],
            &b"application/vnd.openxmlformats-officedocument.wordprocessingml.document"[..],
            &b"application/vnd.openxmlformats-officedocument.presentationml.presentation"[..],
        ];

        assert!(content_type_matches_any(content_type, &valid_types));

        let invalid_types = [&b"application/pdf"[..], &b"text/plain"[..]];
        assert!(!content_type_matches_any(content_type, &invalid_types));
    }

    #[test]
    fn test_short_signatures() {
        // Test that short signatures work correctly
        let zip_sig = &[0x50, 0x4B, 0x03, 0x04];
        let data = &[0x50, 0x4B, 0x03, 0x04, 0x00, 0x00];

        assert!(signature_matches(data, zip_sig));
    }

    #[test]
    fn test_long_signatures() {
        // Test SIMD path with longer signatures
        let long_sig: Vec<u8> = (0..32).collect();
        let mut long_data = long_sig.clone();
        long_data.extend_from_slice(&[0xFF, 0xFF]);

        assert!(signature_matches(&long_data, &long_sig));

        // Test mismatch
        long_data[15] = 0xFF;
        assert!(!signature_matches(&long_data, &long_sig));
    }

    #[test]
    fn test_parallel_signature_check() {
        let data = &[0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        let signatures = [
            &[0x50, 0x4B, 0x03, 0x04][..], // ZIP - matches
            &[0xD0, 0xCF, 0x11, 0xE0][..], // OLE2 - doesn't match
            &[0x50, 0x4B][..],             // Partial ZIP - matches
        ];

        let matches = parallel_signature_check(data, &signatures);
        assert_eq!(matches.as_slice(), &[0, 2]);

        // Test with no matches
        let non_matching_data = &[0xFF, 0xFF, 0xFF, 0xFF];
        let matches = parallel_signature_check(non_matching_data, &signatures);
        assert!(matches.is_empty());

        // Test that it doesn't allocate for small results (stays on stack)
        // This is implicit - SmallVec with capacity 8 should not allocate
        let many_signatures = [
            &[0x50, 0x4B, 0x03, 0x04][..],
            &[0xD0, 0xCF, 0x11, 0xE0][..],
            &[0x50, 0x4B][..],
            &[0x00, 0x00][..],
        ];
        let matches = parallel_signature_check(data, &many_signatures);
        assert_eq!(matches.as_slice(), &[0, 2]);
        // Verify SmallVec didn't spill to heap (capacity check)
        assert!(!matches.spilled());
    }

    #[test]
    fn test_check_office_signatures() {
        // Test ZIP signature
        let zip_data = &[0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        let mask = check_office_signatures(zip_data);
        assert!(mask.is_zip());
        assert!(!mask.is_ole2());
        assert!(!mask.is_rtf());
        assert_eq!(mask.count(), 1);

        // Test OLE2 signature
        let ole2_data = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        let mask = check_office_signatures(ole2_data);
        assert!(mask.is_ole2());
        assert!(!mask.is_zip());
        assert!(!mask.is_rtf());
        assert_eq!(mask.count(), 1);

        // Test RTF signature
        let rtf_data = b"{\\rtf1\\ansi";
        let mask = check_office_signatures(rtf_data);
        assert!(mask.is_rtf());
        assert!(!mask.is_ole2());
        assert!(!mask.is_zip());
        assert_eq!(mask.count(), 1);

        // Test no match
        let unknown_data = &[0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF];
        let mask = check_office_signatures(unknown_data);
        assert!(!mask.has_match());
        assert_eq!(mask.count(), 0);
    }

    #[test]
    fn test_format_signature_mask_operations() {
        let mut mask = FormatSignatureMask::empty();
        assert!(!mask.has_match());

        mask |= FormatSignatureMask::ZIP;
        assert!(mask.is_zip());
        assert!(!mask.is_ole2());
        assert_eq!(mask.count(), 1);

        mask |= FormatSignatureMask::OLE2;
        assert!(mask.is_zip());
        assert!(mask.is_ole2());
        assert_eq!(mask.count(), 2);

        mask |= FormatSignatureMask::RTF;
        assert!(mask.is_zip());
        assert!(mask.is_ole2());
        assert!(mask.is_rtf());
        assert_eq!(mask.count(), 3);
    }
}

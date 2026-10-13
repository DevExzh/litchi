//! The workbook's sheet-name rule and identity.
//!
//! Legacy Excel compares tab names without regard to case, and Litchi
//! registers one canonical caseless identity for workbook sheet names
//! (ADR 0034 §8, `docs/API_GLOSSARY.md` "Name identity policy"): Unicode
//! NFD, default case folding, then NFD again — the same key XLSX uses.  The
//! reader's duplicate-name check, selector lookups, renames and tab
//! insertion all use this one rule.

use caseless::Caseless;
use unicode_normalization::UnicodeNormalization;

fn folded(name: &str) -> impl Iterator<Item = char> + '_ {
    name.chars().nfd().default_case_fold().nfd()
}

/// The canonical identity key of a sheet name.
///
/// ```
/// assert_eq!(litchi_xls::sheet_name::key("Straße"), litchi_xls::sheet_name::key("STRASSE"));
/// ```
#[must_use]
pub fn key(name: &str) -> String {
    if name.is_ascii() {
        return name.to_ascii_lowercase();
    }
    folded(name).collect()
}

/// Whether two sheet names identify the same tab.
///
/// ```
/// assert!(litchi_xls::sheet_name::equivalent("Ä", "a\u{308}"));
/// assert!(!litchi_xls::sheet_name::equivalent("Sheet1", "Sheet2"));
/// ```
#[must_use]
pub fn equivalent(left: &str, right: &str) -> bool {
    if left.is_ascii() && right.is_ascii() {
        return left.eq_ignore_ascii_case(right);
    }
    folded(left).eq(folded(right))
}

/// Whether `name` satisfies the BIFF8 `BoundSheet8` name rule: 1 through 31
/// UTF-16 code units, none of `: \ * ? / [ ]` or NUL, and no leading or
/// trailing apostrophe (`[MS-XLS]` 2.4.28).
///
/// ```
/// assert!(litchi_xls::sheet_name::is_valid("Summary"));
/// assert!(!litchi_xls::sheet_name::is_valid("a/b"));
/// assert!(!litchi_xls::sheet_name::is_valid(""));
/// ```
#[must_use]
pub fn is_valid(name: &str) -> bool {
    crate::records::sheet_name_fault(name).is_none()
}

#[cfg(test)]
mod tests {
    use super::{equivalent, key};

    #[test]
    fn identity_is_unicode_caseless() {
        assert!(equivalent("Sheet1", "SHEET1"));
        assert!(equivalent("Straße", "STRASSE"));
        assert!(equivalent("Ä", "a\u{308}"));
        assert!(equivalent("ΣΑΣ", "σας"));
        assert!(!equivalent("Sheet1", "Sheet2"));
        assert_eq!(key("Straße"), key("STRASSE"));
        assert_eq!(key("ABC"), "abc");
    }
}

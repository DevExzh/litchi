//! Checked CFB storage paths for the inert object owner.
//!
//! `[MS-CFB]` identifies directory entries by their UTF-16 name length and a
//! Unicode simple-uppercase comparison.  Keeping that rule beside target
//! selection prevents invalid names and case-equivalent paths from reaching
//! the package editor, while leaving the stored directory spelling untouched.

use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

const MAX_DIRECTORY_NAME_CODE_UNITS: usize = 31;
const FORBIDDEN_DIRECTORY_NAME_CHARS: [char; 4] = ['/', '\\', ':', '!'];
const STORAGE_OBJECT: u8 = 1;

/// A validated, non-root CFB storage path.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub(crate) struct CfbPath {
    parts: Vec<String>,
}

impl CfbPath {
    pub(crate) fn new(parts: Vec<String>) -> Result<Self, OleError> {
        if parts.is_empty() {
            return Err(OleError::InvalidFormat(
                "object target path is empty".into(),
            ));
        }
        for part in &parts {
            validate_component(part)?;
        }
        Ok(Self { parts })
    }

    pub(crate) fn as_slice(&self) -> &[String] {
        &self.parts
    }

    pub(crate) fn overlaps(&self, other: &Self) -> bool {
        starts_with(&self.parts, &other.parts) || starts_with(&other.parts, &self.parts)
    }

    /// Resolves a host-supplied path to the directory spelling stored in the
    /// CFB.  The operation only reads directory metadata; it never opens a
    /// stream or activates an OLE payload.
    pub(crate) fn resolve<R: Read + Seek>(
        &self,
        ole: &OleFile<R>,
    ) -> Result<Vec<String>, OleError> {
        let mut resolved = Vec::with_capacity(self.parts.len());
        for requested in &self.parts {
            let refs = resolved.iter().map(String::as_str).collect::<Vec<_>>();
            let entry = ole
                .list_directory_entries(&refs)?
                .into_iter()
                .find(|entry| {
                    entry.entry_type == STORAGE_OBJECT && same_component(&entry.name, requested)
                })
                .ok_or_else(|| {
                    OleError::InvalidFormat(format!(
                        "object storage path component {requested:?} not found"
                    ))
                })?;
            resolved.push(entry.name.clone());
        }
        Ok(resolved)
    }
}

/// Iterates the Unicode simple-uppercase UTF-16 comparison units required by
/// `[MS-CFB]` 2.6.4 without retaining a second copy of the component.
struct UppercaseUnits<'a> {
    input: std::str::EncodeUtf16<'a>,
    pending: [u16; 2],
    pending_len: usize,
    pending_index: usize,
}

impl Iterator for UppercaseUnits<'_> {
    type Item = u16;

    fn next(&mut self) -> Option<Self::Item> {
        if self.pending_index < self.pending_len {
            let value = self.pending[self.pending_index];
            self.pending_index += 1;
            return Some(value);
        }

        let unit = self.input.next()?;
        if (0xD800..=0xDFFF).contains(&unit) {
            return Some(unit);
        }

        let character = char::from_u32(u32::from(unit))?;
        let mut uppercase = character.to_uppercase();
        let first = uppercase.next()?;
        if uppercase.next().is_some() {
            // This is a multi-code-point mapping, not a simple mapping.
            return Some(unit);
        }

        self.pending = [0; 2];
        let encoded = first.encode_utf16(&mut self.pending);
        self.pending_len = encoded.len();
        self.pending_index = 1;
        Some(self.pending[0])
    }
}

fn validate_component(component: &str) -> Result<(), OleError> {
    if component.is_empty() {
        return Err(OleError::InvalidFormat(
            "object target path must contain non-empty storage names".into(),
        ));
    }
    if component.contains('\0') {
        return Err(OleError::InvalidFormat(
            "CFB storage names must not contain NUL".into(),
        ));
    }
    if let Some(character) = component
        .chars()
        .find(|character| FORBIDDEN_DIRECTORY_NAME_CHARS.contains(character))
    {
        return Err(OleError::InvalidFormat(format!(
            "CFB storage name contains forbidden character {character:?}"
        )));
    }
    let length = component.encode_utf16().count();
    if length > MAX_DIRECTORY_NAME_CODE_UNITS {
        return Err(OleError::InvalidFormat(format!(
            "CFB storage name uses {length} UTF-16 code units; maximum is {MAX_DIRECTORY_NAME_CODE_UNITS}"
        )));
    }
    Ok(())
}

fn starts_with(path: &[String], prefix: &[String]) -> bool {
    path.len() >= prefix.len()
        && path
            .iter()
            .zip(prefix)
            .all(|(part, expected)| same_component(part, expected))
}

fn same_component(left: &str, right: &str) -> bool {
    left.encode_utf16().count() == right.encode_utf16().count()
        && uppercase_units(left).eq(uppercase_units(right))
}

fn uppercase_units(value: &str) -> UppercaseUnits<'_> {
    UppercaseUnits {
        input: value.encode_utf16(),
        pending: [0; 2],
        pending_len: 0,
        pending_index: 0,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "tests use concise assertions while exercising fallible validation paths"
)]
mod tests {
    use super::{CfbPath, same_component};

    #[test]
    fn cfb_name_comparison_uses_simple_uppercase_without_expansion() {
        assert!(same_component("Pool", "pool"));
        assert!(same_component("Å", "å"));
        assert!(same_component("ſ", "S"));
        assert!(!same_component("ß", "SS"));
        assert!(!same_component("ß", "ẞ"));
    }

    #[test]
    fn path_rejects_names_that_cannot_be_directory_entries() {
        for value in [
            "",
            "bad/name",
            "bad\\name",
            "bad:name",
            "bad!name",
            "bad\0name",
        ] {
            assert!(CfbPath::new(vec![value.to_string()]).is_err(), "{value:?}");
        }
        assert!(CfbPath::new(vec!["😀".repeat(15)]).is_ok());
        assert!(CfbPath::new(vec!["😀".repeat(16)]).is_err());
    }
}

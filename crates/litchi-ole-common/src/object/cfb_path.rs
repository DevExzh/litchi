//! Checked CFB storage paths for the inert object owner.
//!
//! `[MS-CFB]` identifies directory entries by their UTF-16 name length and a
//! simple-uppercase comparison of UTF-16 code points (2.6.4). Every
//! comparison here uses `litchi-cfb`'s [`DirectoryNameKey`], so target
//! selection agrees with the directory tree about which names are one entry.
//! That keeps invalid names and case-equivalent paths from reaching the
//! package editor, and leaves the stored directory spelling untouched.

use litchi_cfb::{DirectoryNameKey, OleError, OleFile};
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
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

    pub(crate) fn try_from_slice(
        parts: &[String],
        resource: &'static str,
    ) -> Result<Self, OleError> {
        if parts.is_empty() {
            return Err(OleError::InvalidFormat("CFB path is empty".into()));
        }
        for part in parts {
            validate_component(part)?;
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(parts.len())
            .map_err(|source| OleError::Allocation { resource, source })?;
        for part in parts {
            let mut value = String::new();
            value
                .try_reserve_exact(part.len())
                .map_err(|source| OleError::Allocation { resource, source })?;
            value.push_str(part);
            owned.push(value);
        }
        Ok(Self { parts: owned })
    }

    pub(crate) fn as_slice(&self) -> &[String] {
        &self.parts
    }

    pub(crate) fn overlaps(&self, other: &Self) -> bool {
        starts_with(&self.parts, &other.parts) || starts_with(&other.parts, &self.parts)
    }

    pub(crate) fn same_as(&self, other: &Self) -> bool {
        same_path(&self.parts, &other.parts)
    }

    pub(crate) fn identity_hash(&self) -> u64 {
        path_identity_hash(&self.parts)
    }

    /// Resolves a host-supplied path to the directory spelling stored in the
    /// CFB.  The operation only reads directory metadata; it never opens a
    /// stream or activates an OLE payload.
    pub(crate) fn resolve<R: Read + Seek>(
        &self,
        ole: &OleFile<R>,
        max_entries: usize,
    ) -> Result<Vec<String>, OleError> {
        let mut resolved = Vec::with_capacity(self.parts.len());
        for requested in &self.parts {
            let refs = resolved.iter().map(String::as_str).collect::<Vec<_>>();
            let mut found = None;
            ole.visit_directory_entry_refs(&refs, max_entries, |entry| {
                if found.is_none()
                    && entry.entry_type == STORAGE_OBJECT
                    && same_component(&entry.name, requested)
                {
                    found = Some(entry.name.clone());
                }
                Ok::<(), OleError>(())
            })?;
            resolved.push(found.ok_or_else(|| {
                OleError::InvalidFormat(format!(
                    "object storage path component {requested:?} not found"
                ))
            })?);
        }
        Ok(resolved)
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

pub(crate) fn same_path(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| same_component(left, right))
}

/// A hash consistent with [`same_path`]: equal paths hash equally.
pub(crate) fn path_identity_hash(path: &[String]) -> u64 {
    let mut hash = DefaultHasher::new();
    path.len().hash(&mut hash);
    for component in path {
        match DirectoryNameKey::new(component) {
            Ok(key) => {
                true.hash(&mut hash);
                key.hash(&mut hash);
            },
            // A name that cannot be a directory entry equals no name (see
            // `same_component`), so any value keeps an index built on this
            // hash sound; lookups confirm a match with `same_path`.
            Err(_) => {
                false.hash(&mut hash);
                component.hash(&mut hash);
            },
        }
    }
    hash.finish()
}

/// Whether two names are the same CFB directory entry, by the comparison
/// `litchi-cfb` uses for its directory tree. A stored name that cannot be a
/// directory entry equals no name; every requested name is validated first.
fn same_component(left: &str, right: &str) -> bool {
    litchi_cfb::directory_names_equal(left, right)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "tests use concise assertions while exercising fallible validation paths"
)]
mod tests {
    use super::{CfbPath, path_identity_hash, same_component};

    #[test]
    fn cfb_name_comparison_is_litchi_cfbs_directory_comparison() {
        assert!(same_component("Pool", "pool"));
        assert!(same_component("Å", "å"));
        assert!(same_component("ſ", "S"));
        assert!(!same_component("ß", "SS"));
        assert!(!same_component("ß", "ẞ"));
        // [MS-CFB] 2.6.4 compares UTF-16 code points and uppercases neither
        // surrogate, so Deseret case pairs are distinct names. (The previous
        // scalar-level fold treated them as one, disagreeing with the
        // directory tree.)
        assert!(!same_component("𐐨", "𐐀"));
        // The iota-subscript letters have a one-code-point simple uppercase
        // mapping although their full mapping expands.
        assert!(same_component("\u{1F80}", "\u{1F88}"));
        for (left, right) in [
            ("Pool", "pool"),
            ("𐐨", "𐐀"),
            ("\u{1F80}", "\u{1F88}"),
            ("ß", "ẞ"),
        ] {
            assert_eq!(
                same_component(left, right),
                litchi_cfb::directory_names_equal(left, right),
                "{left:?} {right:?}"
            );
        }
        assert_eq!(
            path_identity_hash(&["Pool".to_string(), "\u{1F80}".to_string()]),
            path_identity_hash(&["pool".to_string(), "\u{1F88}".to_string()])
        );
        // A stored name that cannot be a directory entry equals nothing, not
        // even itself, and still hashes.
        assert!(!same_component("bad/name", "bad/name"));
        let _ = path_identity_hash(&["bad/name".to_string()]);
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

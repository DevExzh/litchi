//! Relationship-related objects for OPC packages.
//!
//! This module provides types for managing relationships between parts in an OPC package,
//! including internal and external relationships.

use crate::error::{OpcError, Result};
use crate::packuri::PackURI;
use litchi_core::xml::escape_xml;
use std::collections::HashMap;

/// Whether a relationship target is inside or outside the OPC package.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TargetMode {
    /// A target resolved relative to the relationship source part.
    Internal,
    /// An external URI reference preserved without package resolution.
    External,
}

impl TargetMode {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "Internal" => Ok(Self::Internal),
            "External" => Ok(Self::External),
            _ => Err(OpcError::InvalidRelationshipTargetMode(value.to_string())),
        }
    }

    #[inline]
    fn as_xml_value(self) -> &'static str {
        match self {
            Self::Internal => "Internal",
            Self::External => "External",
        }
    }
}

/// A single relationship from a source part to a target.
///
/// Represents a connection between parts in an OPC package, identified by an rId
/// (relationship ID). Can be either internal (pointing to another part) or external
/// (pointing to an external URL).
#[derive(Debug, Clone)]
pub struct Relationship {
    /// Relationship ID (e.g., "rId1", "rId2")
    r_id: String,

    /// Relationship type URI
    reltype: String,

    /// Target reference - either a part URI or external URL
    target_ref: String,

    /// Base URI for resolving relative references
    base_uri: String,

    /// Full source part URI, when known.
    source_uri: Option<String>,

    /// Typed target mode.
    target_mode: TargetMode,
}

impl Relationship {
    /// Create a new relationship.
    ///
    /// # Arguments
    /// * `r_id` - Relationship ID (e.g., "rId1")
    /// * `reltype` - Relationship type URI
    /// * `target_ref` - Target reference (part URI or external URL)
    /// * `base_uri` - Base URI for resolving relative references
    /// * `is_external` - Whether this is an external relationship
    pub fn new(
        r_id: String,
        reltype: String,
        target_ref: String,
        base_uri: String,
        is_external: bool,
    ) -> Self {
        Self::new_with_source(
            r_id,
            reltype,
            target_ref,
            base_uri,
            None,
            if is_external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )
    }

    /// Create a relationship with an explicit target mode.
    pub fn new_with_mode(
        r_id: String,
        reltype: String,
        target_ref: String,
        base_uri: String,
        target_mode: TargetMode,
    ) -> Self {
        Self::new_with_source(r_id, reltype, target_ref, base_uri, None, target_mode)
    }

    fn new_with_source(
        r_id: String,
        reltype: String,
        target_ref: String,
        base_uri: String,
        source_uri: Option<String>,
        target_mode: TargetMode,
    ) -> Self {
        Self {
            r_id,
            reltype,
            target_ref,
            base_uri,
            source_uri,
            target_mode,
        }
    }

    /// Get the relationship ID.
    #[inline]
    pub fn r_id(&self) -> &str {
        &self.r_id
    }

    /// Get the relationship type.
    #[inline]
    pub fn reltype(&self) -> &str {
        &self.reltype
    }

    /// Get the target reference.
    ///
    /// For internal relationships, this is a relative part reference.
    /// For external relationships, this is an absolute URL.
    #[inline]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// Return the path component of the original target URI reference.
    pub fn target_path(&self) -> &str {
        relationship_target_components(&self.target_ref).0
    }

    /// Return the query component without the leading `?`.
    pub fn target_query(&self) -> Option<&str> {
        relationship_target_components(&self.target_ref).1
    }

    /// Return the fragment component without the leading `#`.
    pub fn target_fragment(&self) -> Option<&str> {
        relationship_target_components(&self.target_ref).2
    }

    /// Return the typed target mode.
    #[inline]
    pub fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    /// Check if this is an external relationship.
    #[inline]
    pub fn is_external(&self) -> bool {
        self.target_mode == TargetMode::External
    }

    /// Get the absolute target partname for internal relationships.
    ///
    /// Returns an error if this is an external relationship.
    pub fn target_partname(&self) -> Result<PackURI> {
        if self.is_external() {
            return Err(OpcError::InvalidRelationship(
                "Cannot get target_partname for external relationship".to_string(),
            ));
        }
        let path = self.target_path();
        if path.is_empty() {
            return self
                .source_uri
                .as_deref()
                .filter(|source| *source != "/")
                .ok_or_else(|| {
                    OpcError::InvalidRelationship(
                        "Internal relationship target has no part path".to_string(),
                    )
                })
                .and_then(|source| PackURI::new(source).map_err(OpcError::InvalidPackUri));
        }
        PackURI::from_rel_ref(&self.base_uri, path).map_err(OpcError::InvalidPackUri)
    }
}

pub(crate) fn relationship_target_components(
    reference: &str,
) -> (&str, Option<&str>, Option<&str>) {
    let (before_fragment, fragment) = reference
        .split_once('#')
        .map_or((reference, None), |(before, fragment)| {
            (before, Some(fragment))
        });
    let (path, query) = before_fragment
        .split_once('?')
        .map_or((before_fragment, None), |(path, query)| (path, Some(query)));
    (path, query, fragment)
}

/// Collection of relationships from a single source.
///
/// Uses a HashMap for O(1) lookup by relationship ID while maintaining
/// efficient memory usage by storing references rather than cloning data.
#[derive(Debug)]
pub struct Relationships {
    /// Base URI for resolving relative references
    base_uri: String,

    /// Full source part URI, when this collection belongs to a concrete part.
    source_uri: Option<String>,

    /// Map of relationship ID to Relationship
    rels: HashMap<String, Relationship>,
}

impl Relationships {
    /// Create a new empty relationships collection.
    ///
    /// # Arguments
    /// * `base_uri` - Base URI for resolving relative references
    pub fn new(base_uri: String) -> Self {
        Self {
            base_uri,
            source_uri: None,
            rels: HashMap::new(),
        }
    }

    /// Create a relationship collection owned by a concrete source part.
    pub(crate) fn for_source(source: &PackURI) -> Self {
        Self {
            base_uri: source.base_uri().to_string(),
            source_uri: Some(source.as_str().to_string()),
            rels: HashMap::new(),
        }
    }

    /// Add a relationship to the collection.
    ///
    /// # Arguments
    /// * `reltype` - Relationship type URI
    /// * `target_ref` - Target reference (part URI or external URL)
    /// * `r_id` - Relationship ID
    /// * `is_external` - Whether this is an external relationship
    ///
    /// # Returns
    /// Reference to the newly added relationship
    pub fn add_relationship(
        &mut self,
        reltype: String,
        target_ref: String,
        r_id: String,
        is_external: bool,
    ) -> &Relationship {
        // Preserve compatibility for existing writer call sites without ever
        // replacing an established relationship. New code that needs duplicate
        // diagnostics should use `try_add_relationship`.
        if self.rels.contains_key(&r_id) {
            return self
                .rels
                .get(&r_id)
                .expect("relationship existence was checked");
        }
        let rel = Relationship::new_with_source(
            r_id.clone(),
            reltype,
            target_ref,
            self.base_uri.clone(),
            self.source_uri.clone(),
            if is_external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        );
        self.rels.insert(r_id.clone(), rel);
        // Safe to unwrap since we just inserted it
        self.rels.get(r_id.as_str()).unwrap()
    }

    /// Add a relationship without replacing an existing ID.
    pub fn try_add_relationship(
        &mut self,
        reltype: String,
        target_ref: String,
        r_id: String,
        target_mode: TargetMode,
    ) -> Result<&Relationship> {
        if self.rels.contains_key(&r_id) {
            return Err(OpcError::DuplicateRelationshipId(r_id));
        }
        let relationship = Relationship::new_with_source(
            r_id.clone(),
            reltype,
            target_ref,
            self.base_uri.clone(),
            self.source_uri.clone(),
            target_mode,
        );
        self.rels.insert(r_id.clone(), relationship);
        self.rels.get(&r_id).ok_or_else(|| {
            OpcError::InvalidRelationship("relationship insertion failed".to_string())
        })
    }

    /// Get a relationship by its ID.
    #[inline]
    pub fn get(&self, r_id: &str) -> Option<&Relationship> {
        self.rels.get(r_id)
    }

    /// Get or add a relationship to a target part.
    ///
    /// If a relationship of the given type to the target already exists,
    /// returns that relationship. Otherwise, creates a new one with the
    /// next available rId.
    ///
    /// # Arguments
    /// * `reltype` - Relationship type URI
    /// * `target_ref` - Target reference
    ///
    /// # Returns
    /// Reference to the relationship (existing or newly created)
    pub fn get_or_add(&mut self, reltype: &str, target_ref: &str) -> &Relationship {
        // Check if matching relationship already exists
        for rel in self.rels.values() {
            if rel.reltype() == reltype && rel.target_ref() == target_ref && !rel.is_external() {
                // Return the rId to look it up again (to avoid borrow checker issues)
                let r_id = rel.r_id().to_string();
                return self.rels.get(&r_id).unwrap();
            }
        }

        // Create new relationship with next available rId
        let r_id = self.next_r_id();
        self.add_relationship(reltype.to_string(), target_ref.to_string(), r_id, false)
    }

    /// Get or add an external relationship.
    ///
    /// Similar to `get_or_add` but for external relationships.
    pub fn get_or_add_ext_rel(&mut self, reltype: &str, target_ref: &str) -> String {
        // Check if matching relationship already exists
        for rel in self.rels.values() {
            if rel.reltype() == reltype && rel.target_ref() == target_ref && rel.is_external() {
                return rel.r_id().to_string();
            }
        }

        // Create new relationship with next available rId
        let r_id = self.next_r_id();
        self.add_relationship(
            reltype.to_string(),
            target_ref.to_string(),
            r_id.clone(),
            true,
        );
        r_id
    }

    /// Get the next available relationship ID.
    ///
    /// Generates IDs in the format "rId1", "rId2", etc., filling in gaps
    /// if any exist. Uses efficient integer parsing with atoi_simd.
    fn next_r_id(&self) -> String {
        // Find the highest existing rId number and any gaps
        let mut used_numbers: Vec<u32> = self
            .rels
            .keys()
            .filter_map(|r_id| {
                // Extract number from "rId123" format using fast byte searching
                if r_id.len() > 3 && &r_id[..3] == "rId" {
                    atoi_simd::parse::<u32, false, false>(&r_id.as_bytes()[3..]).ok()
                } else {
                    None
                }
            })
            .collect();

        // Sort to find gaps efficiently
        used_numbers.sort_unstable();

        // Find first gap or use next number
        let mut next_num = 1u32;
        for &num in &used_numbers {
            match num.cmp(&next_num) {
                std::cmp::Ordering::Equal => next_num += 1,
                std::cmp::Ordering::Greater => break,
                std::cmp::Ordering::Less => {},
            }
        }

        format!("rId{}", next_num)
    }

    /// Get the relationship of a specific type.
    ///
    /// Returns an error if no relationship of the type is found,
    /// or if multiple relationships of the type exist.
    pub fn part_with_reltype(&self, reltype: &str) -> Result<&Relationship> {
        let matching: Vec<&Relationship> = self
            .rels
            .values()
            .filter(|rel| rel.reltype() == reltype)
            .collect();

        match matching.len() {
            0 => Err(OpcError::RelationshipNotFound(format!(
                "No relationship of type '{}'",
                reltype
            ))),
            1 => Ok(matching[0]),
            _ => Err(OpcError::InvalidRelationship(format!(
                "Multiple relationships of type '{}'",
                reltype
            ))),
        }
    }

    /// Get an iterator over all relationships.
    #[inline]
    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.rels.values()
    }

    /// Get the number of relationships in the collection.
    #[inline]
    pub fn len(&self) -> usize {
        self.rels.len()
    }

    /// Check if the collection is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.rels.is_empty()
    }

    /// Remove a relationship by its ID.
    pub fn remove(&mut self, r_id: &str) -> Option<Relationship> {
        self.rels.remove(r_id)
    }

    /// Serialize relationships to XML format.
    ///
    /// Generates the XML for a .rels file, with relationships sorted by rId
    /// for consistent output.
    pub fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(1024);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        // Sort relationships by rId for consistent output
        let mut rels: Vec<&Relationship> = self.rels.values().collect();
        rels.sort_by_key(|rel| rel.r_id());

        for rel in rels {
            let target_mode = match rel.target_mode() {
                TargetMode::Internal => "",
                mode => match mode.as_xml_value() {
                    "External" => r#" TargetMode="External""#,
                    _ => "",
                },
            };

            xml.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}"{}/>"#,
                escape_xml(rel.r_id()),
                escape_xml(rel.reltype()),
                escape_xml(rel.target_ref()),
                target_mode
            ));
        }

        xml.push_str("</Relationships>");

        xml
    }
}

impl Default for Relationships {
    fn default() -> Self {
        Self::new("/".to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_relationship_creation() {
        let rel = Relationship::new(
            "rId1".to_string(),
            "http://example.com/rel".to_string(),
            "target.xml".to_string(),
            "/word".to_string(),
            false,
        );

        assert_eq!(rel.r_id(), "rId1");
        assert_eq!(rel.reltype(), "http://example.com/rel");
        assert!(!rel.is_external());
    }

    #[test]
    fn test_next_r_id() {
        let mut rels = Relationships::new("/word".to_string());

        let r_id1 = rels.next_r_id();
        assert_eq!(r_id1, "rId1");

        rels.add_relationship(
            "type1".to_string(),
            "target1".to_string(),
            "rId1".to_string(),
            false,
        );

        let r_id2 = rels.next_r_id();
        assert_eq!(r_id2, "rId2");
    }

    #[test]
    fn test_get_or_add() {
        let mut rels = Relationships::new("/word".to_string());

        let rel1 = rels.get_or_add("type1", "target1");
        assert_eq!(rel1.r_id(), "rId1");

        // Getting the same relationship should return the same rId
        let rel2 = rels.get_or_add("type1", "target1");
        assert_eq!(rel2.r_id(), "rId1");

        // Different target should create new relationship
        let rel3 = rels.get_or_add("type1", "target2");
        assert_eq!(rel3.r_id(), "rId2");
    }

    #[test]
    fn target_components_are_preserved_while_the_path_resolves() {
        let relationship = Relationship::new_with_mode(
            "rId1".to_string(),
            "urn:test".to_string(),
            "../media/image.png?variant=2#preview".to_string(),
            "/word/document".to_string(),
            TargetMode::Internal,
        );
        assert_eq!(relationship.target_path(), "../media/image.png");
        assert_eq!(relationship.target_query(), Some("variant=2"));
        assert_eq!(relationship.target_fragment(), Some("preview"));
        assert_eq!(
            relationship.target_partname().unwrap().as_str(),
            "/word/media/image.png"
        );
        assert_eq!(
            relationship.target_ref(),
            "../media/image.png?variant=2#preview"
        );
    }

    #[test]
    fn fallible_insertion_never_replaces_a_duplicate_id() {
        let mut relationships = Relationships::new("/word".to_string());
        relationships
            .try_add_relationship(
                "urn:first".to_string(),
                "first.xml".to_string(),
                "rId1".to_string(),
                TargetMode::Internal,
            )
            .unwrap();
        assert!(matches!(
            relationships.try_add_relationship(
                "urn:second".to_string(),
                "second.xml".to_string(),
                "rId1".to_string(),
                TargetMode::Internal,
            ),
            Err(OpcError::DuplicateRelationshipId(id)) if id == "rId1"
        ));
        assert_eq!(relationships.get("rId1").unwrap().target_ref(), "first.xml");
    }
}

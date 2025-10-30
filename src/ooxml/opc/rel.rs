use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::PackURI;
/// Relationship-related objects for OPC packages.
///
/// This module provides types for managing relationships between parts in an OPC package,
/// including internal and external relationships.
use std::collections::HashMap;

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

    /// Whether this is an external relationship
    is_external: bool,
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
        Self {
            r_id,
            reltype,
            target_ref,
            base_uri,
            is_external,
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

    /// Check if this is an external relationship.
    #[inline]
    pub fn is_external(&self) -> bool {
        self.is_external
    }

    /// Get the absolute target partname for internal relationships.
    ///
    /// Returns an error if this is an external relationship.
    pub fn target_partname(&self) -> Result<PackURI> {
        if self.is_external {
            return Err(OpcError::InvalidRelationship(
                "Cannot get target_partname for external relationship".to_string(),
            ));
        }
        PackURI::from_rel_ref(&self.base_uri, &self.target_ref).map_err(OpcError::InvalidPackUri)
    }
}

/// Collection of relationships from a single source.
///
/// Uses a HashMap for O(1) lookup by relationship ID while maintaining
/// efficient memory usage by storing references rather than cloning data.
#[derive(Debug)]
pub struct Relationships {
    /// Base URI for resolving relative references
    base_uri: String,

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
        let rel = Relationship::new(
            r_id.clone(),
            reltype,
            target_ref,
            self.base_uri.clone(),
            is_external,
        );
        self.rels.insert(r_id.clone(), rel);
        // Safe to unwrap since we just inserted it
        self.rels.get(r_id.as_str()).unwrap()
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
                    atoi_simd::parse::<u32>(&r_id.as_bytes()[3..]).ok()
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
        xml.push('\n');
        xml.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        xml.push('\n');

        // Sort relationships by rId for consistent output
        let mut rels: Vec<&Relationship> = self.rels.values().collect();
        rels.sort_by_key(|rel| rel.r_id());

        for rel in rels {
            let target_mode = if rel.is_external() {
                r#" TargetMode="External""#
            } else {
                ""
            };

            xml.push_str(&format!(
                r#"  <Relationship Id="{}" Type="{}" Target="{}"{}/>"#,
                Self::escape_xml(rel.r_id()),
                Self::escape_xml(rel.reltype()),
                Self::escape_xml(rel.target_ref()),
                target_mode
            ));
            xml.push('\n');
        }

        xml.push_str("</Relationships>");

        xml
    }

    /// Escape XML special characters.
    #[inline]
    fn escape_xml(s: &str) -> String {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
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
}

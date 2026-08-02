//! YAML front-matter serialization for [`litchi_core::Metadata`]. Lives in the
//! umbrella because `serde_saphyr` is not a dep of `litchi-core`.

use litchi_core::Error;
use litchi_core::Metadata;
use litchi_core::Result;

/// Extension trait providing YAML front-matter serialization for
/// [`litchi_core::Metadata`] (requires the `serde_saphyr` workspace dep,
/// which lives in the umbrella crate, not in `litchi-core`).
///
/// Bring this trait into scope to call `metadata.to_yaml_front_matter()`:
///
/// ```rust,no_run
/// use litchi::MetadataYaml;
/// use litchi_core::Metadata;
///
/// let md = Metadata::default();
/// let _yaml = md.to_yaml_front_matter().unwrap();
/// ```
pub trait MetadataYaml {
    /// Convert metadata to YAML front matter format.
    ///
    /// Returns a string containing the YAML front matter block,
    /// or an empty string if no metadata is available.
    fn to_yaml_front_matter(&self) -> Result<String>;
}

impl MetadataYaml for Metadata {
    fn to_yaml_front_matter(&self) -> Result<String> {
        if !self.has_data() {
            return Ok(String::new());
        }

        let yaml_string = serde_saphyr::to_string(self)
            .map_err(|e| Error::Other(format!("Failed to serialize metadata to YAML: {}", e)))?;

        // Add YAML front matter delimiters
        Ok(format!("---\n{}---\n\n", yaml_string))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_metadata_to_yaml_front_matter() {
        let empty_metadata = Metadata::default();
        let yaml = empty_metadata.to_yaml_front_matter().unwrap();
        assert_eq!(yaml, "");

        let metadata = Metadata {
            title: Some("Test Document".to_string()),
            author: Some("Test Author".to_string()),
            subject: Some("Test Subject".to_string()),
            ..Default::default()
        };

        let yaml = metadata.to_yaml_front_matter().unwrap();
        assert!(yaml.starts_with("---\n"));
        assert!(yaml.ends_with("---\n\n"));
        assert!(yaml.contains("title: Test Document"));
        assert!(yaml.contains("author: Test Author"));
        assert!(yaml.contains("subject: Test Subject"));
    }

    #[test]
    #[cfg(feature = "doc")]
    fn test_ole_metadata_conversion() {
        let ole_metadata = litchi_cfb::OleMetadata {
            title: Some("OLE Document".to_string()),
            author: Some("OLE Author".to_string()),
            codepage: Some(65001),
            ..Default::default()
        };

        let metadata: Metadata = ole_metadata.into();
        assert_eq!(metadata.title, Some("OLE Document".to_string()));
        assert_eq!(metadata.author, Some("OLE Author".to_string()));
        assert_eq!(metadata.codepage, Some(65001));
    }
}

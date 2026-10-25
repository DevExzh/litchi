//! Caller-selected bounds for the raw MS-XLDM storage and metadata codecs.

use crate::error::{Error, LimitKind, Result};
use crate::model::{
    MAX_DIRECTORY_BYTES, MAX_FILES, MAX_PARTITIONS, MAX_PATH_BYTES, MAX_STORAGE_BYTES,
    MAX_XML_DEPTH, MAX_XML_NODES, MAX_XML_TEXT_BYTES,
};

/// Resource policy for borrowed XLDM storage and section-2.5 metadata.
///
/// The format ceilings remain hard limits. A caller may select a narrower
/// policy, but an operation never widens the policy retained by an inspected
/// owner. `DEFAULT` preserves the historical bounded behavior while every raw
/// entry point still requires a policy argument explicitly.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[must_use]
pub struct Limits {
    /// Maximum bytes admitted from one outer storage source.
    pub max_source_bytes: usize,
    /// Maximum bytes emitted by an outer storage or metadata writer.
    pub max_output_bytes: usize,
    /// Maximum virtual-directory XML bytes.
    pub max_directory_bytes: usize,
    /// Maximum virtual-directory and logged-file entries.
    pub max_files: usize,
    /// Maximum partition records.
    pub max_partitions: usize,
    /// Maximum UTF-8 bytes in one storage path or scalar leaf.
    pub max_path_bytes: usize,
    /// Maximum XML nodes in outer storage XML.
    pub max_xml_nodes: usize,
    /// Maximum XML nesting depth in outer storage XML.
    pub max_xml_depth: usize,
    /// Maximum decoded XML text bytes in outer storage XML.
    pub max_xml_text_bytes: usize,
    /// Maximum section-2.5 metadata XML bytes.
    pub max_metadata_bytes: usize,
    /// Maximum section-2.5 XML nodes.
    pub max_metadata_nodes: usize,
    /// Maximum section-2.5 XML nesting depth.
    pub max_metadata_depth: usize,
    /// Maximum decoded section-2.5 text bytes.
    pub max_metadata_text_bytes: usize,
    /// Maximum section-2.5 files discovered from the backup log.
    pub max_metadata_files: usize,
}

impl Limits {
    /// The historical bounded XLDM policy.
    pub const DEFAULT: Self = Self {
        max_source_bytes: MAX_STORAGE_BYTES,
        max_output_bytes: MAX_STORAGE_BYTES,
        max_directory_bytes: MAX_DIRECTORY_BYTES,
        max_files: MAX_FILES,
        max_partitions: MAX_PARTITIONS,
        max_path_bytes: MAX_PATH_BYTES,
        max_xml_nodes: MAX_XML_NODES,
        max_xml_depth: MAX_XML_DEPTH,
        max_xml_text_bytes: MAX_XML_TEXT_BYTES,
        max_metadata_bytes: super::metadata::MAX_METADATA_BYTES,
        max_metadata_nodes: super::metadata::MAX_XML_NODES,
        max_metadata_depth: super::metadata::MAX_XML_DEPTH,
        max_metadata_text_bytes: super::metadata::MAX_TEXT_BYTES,
        max_metadata_files: super::metadata::MAX_METADATA_FILES,
    };

    /// Construct a policy with selected source and output byte ceilings.
    #[must_use = "use the returned resource policy"]
    pub const fn new(max_source_bytes: usize, max_output_bytes: usize) -> Self {
        Self {
            max_source_bytes,
            max_output_bytes,
            ..Self::DEFAULT
        }
    }

    /// Return this policy intersected with another policy.
    #[must_use = "use the intersected resource policy"]
    pub const fn min(self, other: Self) -> Self {
        Self {
            max_source_bytes: min_usize(self.max_source_bytes, other.max_source_bytes),
            max_output_bytes: min_usize(self.max_output_bytes, other.max_output_bytes),
            max_directory_bytes: min_usize(self.max_directory_bytes, other.max_directory_bytes),
            max_files: min_usize(self.max_files, other.max_files),
            max_partitions: min_usize(self.max_partitions, other.max_partitions),
            max_path_bytes: min_usize(self.max_path_bytes, other.max_path_bytes),
            max_xml_nodes: min_usize(self.max_xml_nodes, other.max_xml_nodes),
            max_xml_depth: min_usize(self.max_xml_depth, other.max_xml_depth),
            max_xml_text_bytes: min_usize(self.max_xml_text_bytes, other.max_xml_text_bytes),
            max_metadata_bytes: min_usize(self.max_metadata_bytes, other.max_metadata_bytes),
            max_metadata_nodes: min_usize(self.max_metadata_nodes, other.max_metadata_nodes),
            max_metadata_depth: min_usize(self.max_metadata_depth, other.max_metadata_depth),
            max_metadata_text_bytes: min_usize(
                self.max_metadata_text_bytes,
                other.max_metadata_text_bytes,
            ),
            max_metadata_files: min_usize(self.max_metadata_files, other.max_metadata_files),
        }
    }

    /// Validate the policy against the hard format ceilings.
    pub(crate) fn validate(self) -> Result<Self> {
        check(
            LimitKind::XldmSourceBytes,
            self.max_source_bytes,
            MAX_STORAGE_BYTES,
        )?;
        check(
            LimitKind::XldmOutputBytes,
            self.max_output_bytes,
            MAX_STORAGE_BYTES,
        )?;
        check(
            LimitKind::DirectoryBytes,
            self.max_directory_bytes,
            MAX_DIRECTORY_BYTES,
        )?;
        check(LimitKind::Files, self.max_files, MAX_FILES)?;
        check(LimitKind::Partitions, self.max_partitions, MAX_PARTITIONS)?;
        check(LimitKind::PathBytes, self.max_path_bytes, MAX_PATH_BYTES)?;
        check(LimitKind::XmlNodes, self.max_xml_nodes, MAX_XML_NODES)?;
        check(LimitKind::XmlDepth, self.max_xml_depth, MAX_XML_DEPTH)?;
        check(
            LimitKind::XmlTextBytes,
            self.max_xml_text_bytes,
            MAX_XML_TEXT_BYTES,
        )?;
        check(
            LimitKind::MetadataBytes,
            self.max_metadata_bytes,
            super::metadata::MAX_METADATA_BYTES,
        )?;
        check(
            LimitKind::MetadataNodes,
            self.max_metadata_nodes,
            super::metadata::MAX_XML_NODES,
        )?;
        check(
            LimitKind::MetadataDepth,
            self.max_metadata_depth,
            super::metadata::MAX_XML_DEPTH,
        )?;
        check(
            LimitKind::MetadataTextBytes,
            self.max_metadata_text_bytes,
            super::metadata::MAX_TEXT_BYTES,
        )?;
        check(
            LimitKind::MetadataFiles,
            self.max_metadata_files,
            super::metadata::MAX_METADATA_FILES,
        )?;
        Ok(self)
    }

    /// Maximum source bytes admitted by this policy.
    #[must_use]
    pub const fn max_source_bytes(self) -> usize {
        self.max_source_bytes
    }

    /// Maximum output bytes emitted by this policy.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn check(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

//! Objects that implement reading and writing OPC packages.
//!
//! This module provides the main `OpcPackage` type, which represents an Open Packaging
//! Convention package in memory. It manages parts, relationships, and provides
//! high-level operations for working with office documents.

use crate::constants::relationship_type;
use crate::error::{OpcError, Result};
use crate::limits::ReadLimits;
use crate::members::NonPartMember;
use crate::packuri::{PACKAGE_URI, PackURI, PartNameConflict};
use crate::part::{Part, PartFactory};
use crate::phys_pkg::{OwnedPhysPkgReader, PhysPkgReader};
use crate::pkgreader::PackageReader;
use crate::rel::Relationships;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};
use std::path::Path;

/// Options for saving an OPC package.
#[derive(Debug, Clone, Default)]
pub struct SaveOptions {
    /// Typed font-embedding policy; invalid boolean combinations are impossible.
    pub fonts: FontEmbedding,
}

/// Font publication policy used when an Office package is saved.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FontEmbedding {
    /// Do not discover or publish fonts.
    #[default]
    None,
    /// Publish complete selected font faces.
    Full,
    /// Publish only the glyphs known to be used by the document.
    Subset,
}

/// Main API class for working with OPC packages.
///
/// `OpcPackage` represents an Open Packaging Convention package in memory,
/// providing access to parts, relationships, and package-level operations.
/// Uses efficient data structures and minimal cloning for best performance.
#[allow(
    clippy::module_name_repetitions,
    reason = "OpcPackage is the established public name for the package module's main type."
)]
#[derive(Clone)]
pub struct OpcPackage {
    /// Package-level relationships
    rels: Relationships,

    /// All parts in the package, indexed by partname
    /// Using Box<dyn Part + Send + Sync> for trait objects to allow different part types
    /// `PackURI` keys avoid string allocations compared to String keys
    parts: HashMap<PackURI, Box<dyn Part + Send + Sync>>,

    /// ZIP items the reader found but did not model as parts
    non_part_members: Vec<NonPartMember>,

    /// Save preferences
    save_options: SaveOptions,
}

impl std::fmt::Debug for OpcPackage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("OpcPackage")
            .field("rels", &self.rels)
            .field("parts_count", &self.parts.len())
            .field("non_part_members", &self.non_part_members)
            .field("save_options", &self.save_options)
            .finish()
    }
}

impl OpcPackage {
    /// Create a new empty OPC package.
    #[must_use]
    pub fn new() -> Self {
        Self {
            rels: Relationships::new(PACKAGE_URI.to_string()),
            parts: HashMap::new(),
            non_part_members: Vec::new(),
            save_options: SaveOptions::default(),
        }
    }

    /// ZIP items that were present in the opened archive but are not OPC parts.
    ///
    /// A reader must not reject a package because a ZIP tool left junk in the
    /// archive, but it must not hide the junk either. Each entry names the ZIP
    /// item and why it was not modelled as a part; the bytes stay in the source
    /// archive and are never decompressed.
    #[must_use]
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Set save options for the package.
    pub fn set_save_options(&mut self, options: SaveOptions) {
        self.save_options = options;
    }

    /// Get current save options.
    #[must_use]
    pub fn save_options(&self) -> &SaveOptions {
        &self.save_options
    }

    /// Configure font embedding with one self-documenting policy.
    pub fn with_fonts(&mut self, policy: FontEmbedding) -> &mut Self {
        self.save_options.fonts = policy;
        self
    }

    /// Open an OPC package from a file.
    ///
    /// # Arguments
    /// * `path` - Path to the package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new `OpcPackage` instance loaded with the package contents
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::package::OpcPackage;
    ///
    /// let pkg = OpcPackage::open("document.docx").unwrap();
    /// ```
    ///
    /// # Errors
    /// Returns an error if the file cannot be read or is not a valid OPC package.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open an OPC package from a file with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the file cannot be read, violates `limits`, or is not
    /// a valid OPC package.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: ReadLimits) -> Result<Self> {
        let owned_reader = OwnedPhysPkgReader::open_with_limits(path, limits)?;
        let phys_reader = owned_reader.reader()?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Load an OPC package from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    ///
    /// # Errors
    /// Returns an error if the archive cannot be read or is not a valid OPC package.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Load an OPC package from a reader with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the archive cannot be read, violates `limits`, or is
    /// not a valid OPC package.
    pub fn from_reader_with_limits<R: Read>(reader: R, limits: ReadLimits) -> Result<Self> {
        let owned_reader = OwnedPhysPkgReader::from_reader_with_limits(reader, limits)?;
        let phys_reader = owned_reader.reader()?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Move an owned ZIP archive into the package reader.
    ///
    /// This avoids copying the archive buffer before parts are decompressed.
    ///
    /// # Errors
    /// Returns an error if the archive is not a valid OPC package.
    pub fn from_vec(data: Vec<u8>) -> Result<Self> {
        Self::from_vec_with_limits(data, ReadLimits::default())
    }

    /// Move an owned ZIP archive into the package reader with explicit limits.
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC package.
    pub fn from_vec_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        let owned_reader = OwnedPhysPkgReader::from_bytes_with_limits(data, limits)?;
        let phys_reader = owned_reader.reader()?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Load an OPC package from a byte slice.
    ///
    /// # Arguments
    /// * `data` - The ZIP archive data as a byte slice
    ///
    /// # Errors
    /// Returns an error if the archive is not a valid OPC package.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(data, ReadLimits::default())
    }

    /// Load an OPC package from a byte slice with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC package.
    pub fn from_bytes_with_limits(data: &[u8], limits: ReadLimits) -> Result<Self> {
        let phys_reader = PhysPkgReader::new_with_limits(data, limits)?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Unmarshal a package from a package reader.
    ///
    /// This is the main deserialization logic that converts serialized parts
    /// and relationships into the in-memory object graph.
    ///
    /// Optimized to minimize clones by consuming the package reader and moving data.
    fn unmarshal(mut pkg_reader: PackageReader) -> Result<Self> {
        let mut package = Self::new();

        // Get ownership of package relationships, parts, and non-part members
        let pkg_srels = pkg_reader.take_pkg_srels();
        let sparts = pkg_reader.take_sparts();
        package.non_part_members = pkg_reader.take_non_part_members();

        // Pre-allocate with known capacity to avoid reallocations
        let mut parts_map: HashMap<PackURI, Box<dyn Part + Send + Sync>> =
            HashMap::with_capacity(sparts.len());

        // Create all parts - move data instead of cloning
        for spart in sparts {
            let partname = spart.partname.clone(); // Need to clone partname for the HashMap key
            let mut part = PartFactory::load(
                spart.partname,     // Move
                spart.content_type, // Move
                spart.blob,         // Move (blob is Arc internally if large)
            )?;

            // Load part relationships
            for srel in spart.srels {
                part.rels_mut().try_add_relationship(
                    srel.reltype,    // Move
                    srel.target_ref, // Move
                    srel.r_id,       // Move
                    srel.target_mode,
                )?;
            }

            parts_map.insert(partname, part);
        }

        // Load package relationships - move instead of clone
        for srel in pkg_srels {
            package.rels.try_add_relationship(
                srel.reltype,    // Move
                srel.target_ref, // Move
                srel.r_id,       // Move
                srel.target_mode,
            )?;
        }

        package.parts = parts_map;
        Ok(package)
    }

    /// Get a reference to the main document part.
    ///
    /// For Word documents, this is the document.xml part.
    /// For Excel, the workbook.xml part.
    /// For `PowerPoint`, the presentation.xml part.
    ///
    /// # Errors
    /// Returns an error if the package has no main-document relationship, has
    /// more than one, the relationship is external, or the target part is missing.
    pub fn main_document_part(&self) -> Result<&dyn Part> {
        let mut matching = self.rels.iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::OFFICE_DOCUMENT | relationship_type::STRICT_OFFICE_DOCUMENT
            )
        });
        let rel = matching.next().ok_or_else(|| {
            OpcError::InvalidRelationship("main-document relationship is missing".to_string())
        })?;
        if matching.next().is_some() {
            return Err(OpcError::InvalidRelationship(
                "package has multiple main-document relationships".to_string(),
            ));
        }
        if rel.is_external() {
            return Err(OpcError::InvalidRelationship(
                "main-document relationship cannot be external".to_string(),
            ));
        }
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Get a part by its partname.
    ///
    /// # Arguments
    /// * `partname` - The `PackURI` of the part to retrieve
    ///
    /// # Errors
    /// Returns `OpcError::PartNotFound` if no part with `partname` exists.
    pub fn get_part(&self, partname: &PackURI) -> Result<&dyn Part> {
        if let Some(part) = self.parts.get(partname) {
            let part_ref: &dyn Part = &**part;
            return Ok(part_ref);
        }
        self.find_case_insensitive(partname)
            .map(|(_, part)| part)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Locate a part whose name matches `partname` ignoring ASCII case.
    ///
    /// OPC compares part names case-insensitively, which is why a package
    /// containing two names differing only by case is rejected as ambiguous
    /// when it is read. Because that ambiguity cannot survive loading, this
    /// fallback can match at most one part, and it is what lets a package whose
    /// writer stored `/xl/sharedstrings.xml` still resolve a relationship
    /// targeting `sharedStrings.xml` — without it those parts are simply
    /// unreachable and their content silently reads as absent.
    ///
    /// The exact lookup above is the fast path; this linear scan runs only on a
    /// miss, and the part count is already bounded when the package is read.
    fn find_case_insensitive(&self, partname: &PackURI) -> Option<(&PackURI, &dyn Part)> {
        let wanted = partname.as_str();
        self.parts
            .iter()
            .find(|(name, _)| name.as_str().eq_ignore_ascii_case(wanted))
            .map(|(name, part)| {
                let part_ref: &dyn Part = &**part;
                (name, part_ref)
            })
    }

    /// Get a mutable reference to a part by its partname.
    ///
    /// # Errors
    /// Returns `OpcError::PartNotFound` if no part with `partname` exists.
    pub fn get_part_mut(&mut self, partname: &PackURI) -> Result<&mut dyn Part> {
        // Resolve the stored key first so the borrow of `self.parts` ends before
        // the mutable lookup; see `find_case_insensitive` for why this matches.
        let key = if self.parts.contains_key(partname) {
            partname.clone()
        } else {
            match self.find_case_insensitive(partname) {
                Some((name, _)) => name.clone(),
                None => return Err(OpcError::PartNotFound(partname.to_string())),
            }
        };
        self.parts
            .get_mut(&key)
            .map(|b| {
                let part: &mut dyn Part = &mut **b;
                part
            })
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Get a part by relationship type from the package level.
    ///
    /// # Arguments
    /// * `reltype` - The relationship type URI
    ///
    /// # Errors
    /// Returns an error if no relationship of `reltype` exists or the target
    /// part is missing.
    pub fn part_by_reltype(&self, reltype: &str) -> Result<&dyn Part> {
        let rel = self.rels.part_with_reltype(reltype)?;
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Add a new part to the package.
    ///
    /// # Arguments
    /// * `part` - The part to add
    pub fn add_part(&mut self, part: Box<dyn Part + Send + Sync>) {
        let partname = part.partname().clone();
        self.parts.insert(partname, part);
    }

    /// Try to add a part without replacing an existing or ambiguous part name.
    ///
    /// # Errors
    /// Returns an error if the part's partname duplicates or conflicts with an
    /// existing part name; the existing part is left untouched.
    pub fn try_add_part(&mut self, part: Box<dyn Part + Send + Sync>) -> Result<()> {
        let partname = part.partname().clone();
        self.validate_new_part_name(&partname)?;
        self.parts.insert(partname, part);
        Ok(())
    }

    /// Validate that a new part name would not replace or conflict with an existing part.
    ///
    /// # Errors
    /// Returns an error if `partname` duplicates or conflicts with an existing
    /// part name.
    pub fn validate_new_part_name(&self, partname: &PackURI) -> Result<()> {
        for existing in self.parts.keys() {
            if let Some(conflict) = existing.conflict_with(partname) {
                return Err(part_name_conflict_error(existing, partname, conflict));
            }
        }
        Ok(())
    }

    /// Remove a part by name, returning whether it existed.
    pub fn remove_part(&mut self, partname: &PackURI) -> bool {
        self.parts.remove(partname).is_some()
    }

    /// Get an iterator over all parts in the package.
    pub fn iter_parts(&self) -> impl Iterator<Item = &dyn Part> {
        self.parts.values().map(|b| {
            let part: &dyn Part = &**b;
            part
        })
    }

    /// Get the number of parts in the package.
    #[must_use]
    pub fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// Get a reference to the package-level relationships.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.rels
    }

    /// Get a mutable reference to the package-level relationships.
    pub fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    /// Whether a package-level signature-origin relationship is present.
    ///
    /// This is a cheap capability check. Use [`Self::signatures`] when graph
    /// validation and cryptographic verification are required.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.rels.iter().any(|relationship| {
            relationship.reltype() == relationship_type::DIGITAL_SIGNATURE_ORIGIN
        })
    }

    /// Verifies every OPC signature with the safe strict policy.
    ///
    /// # Errors
    /// Returns an error if the signature graph is ambiguous or spoofed, or if
    /// signature verification fails.
    #[cfg(feature = "sign")]
    pub fn signatures(&self) -> crate::sign::Result<Vec<crate::sign::Report>> {
        crate::sign::signatures(self, &litchi_sign::Policy::strict())
    }

    /// Verifies every OPC signature with an explicit trust-neutral policy.
    ///
    /// # Errors
    /// Returns an error if the signature graph is ambiguous or spoofed, or if
    /// signature verification fails.
    #[cfg(feature = "sign")]
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> crate::sign::Result<Vec<crate::sign::Report>> {
        crate::sign::signatures(self, policy)
    }

    /// Adds a signature while retaining every existing valid signature.
    ///
    /// # Errors
    /// Returns an error if an existing signature is invalid or the new
    /// signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn sign(&mut self, signer: &litchi_sign::Signer) -> crate::sign::Result<PackURI> {
        crate::sign::sign(self, signer, &litchi_sign::Limits::standard())
    }

    /// Adds a signature with explicit authoring resource bounds.
    ///
    /// # Errors
    /// Returns an error if an existing signature is invalid, `limits` are
    /// exceeded, or the new signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> crate::sign::Result<PackURI> {
        crate::sign::sign(self, signer, limits)
    }

    /// Atomically replaces the validated signature graph with one signature.
    ///
    /// # Errors
    /// Returns an error if the signature graph is invalid or the replacement
    /// signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn resign(&mut self, signer: &litchi_sign::Signer) -> crate::sign::Result<PackURI> {
        crate::sign::resign(self, signer, &litchi_sign::Limits::standard())
    }

    /// Atomically replaces signatures with explicit authoring resource bounds.
    ///
    /// # Errors
    /// Returns an error if the signature graph is invalid, `limits` are exceeded,
    /// or the replacement signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> crate::sign::Result<PackURI> {
        crate::sign::resign(self, signer, limits)
    }

    /// Removes all signature relationships and infrastructure parts.
    ///
    /// Deletion is infallible and idempotent, including for a malformed graph.
    pub fn unsign(&mut self) {
        self.strip_signature_graph();
    }

    /// Relate the package to a part.
    ///
    /// Creates or reuses a relationship from the package to the specified part.
    ///
    /// # Arguments
    /// * `partname` - The target part's partname
    /// * `reltype` - The relationship type URI
    ///
    /// # Returns
    /// The relationship ID (rId)
    pub fn relate_to(&mut self, partname: &str, reltype: &str) -> String {
        let rel = self.rels.get_or_add(reltype, partname);
        rel.r_id().to_string()
    }

    /// Add an external relationship (e.g., for hyperlinks).
    ///
    /// # Arguments
    /// * `target_url` - External URL target
    /// * `reltype` - Relationship type
    ///
    /// # Returns
    /// The relationship ID (e.g., "rId1")
    pub fn relate_to_external(&mut self, target_url: &str, reltype: &str) -> String {
        self.rels.get_or_add_ext_rel(reltype, target_url)
    }

    /// Get mutable access to package-level relationships.
    ///
    /// Useful for advanced relationship management.
    pub fn relationships_mut(&mut self) -> &mut Relationships {
        &mut self.rels
    }

    /// Find the next available partname for a part template.
    ///
    /// Useful for creating new parts with sequential numbering (e.g., image1.png, image2.png).
    /// Uses efficient string operations to minimize allocations.
    ///
    /// # Arguments
    /// * `template` - A format string with a %d placeholder for the number
    ///
    /// # Example
    /// ```no_run
    /// # use litchi_opc::package::OpcPackage;
    /// # let mut pkg = OpcPackage::new();
    /// let next_image = pkg.next_partname("/word/media/image%d.png");
    /// ```
    ///
    /// # Errors
    /// Returns an error if the template has no `%d` placeholder, a candidate
    /// partname is invalid, or no free name exists within the bounded search.
    pub fn next_partname(&self, template: &str) -> Result<PackURI> {
        // Find the position of %d in the template for efficient replacement
        let percent_d_pos = template.find("%d").ok_or_else(|| {
            OpcError::InvalidPackUri("Template must contain %d placeholder".to_string())
        })?;

        let mut n = 1u32;
        let mut candidate_bytes = Vec::with_capacity(template.len() + 10); // Pre-allocate

        loop {
            // Clear and reuse the vector for each candidate
            candidate_bytes.clear();

            // Build candidate string more efficiently
            candidate_bytes.extend_from_slice(&template.as_bytes()[..percent_d_pos]);
            candidate_bytes.extend_from_slice(itoa::Buffer::new().format(n).as_bytes());
            candidate_bytes.extend_from_slice(&template.as_bytes()[percent_d_pos + 2..]);

            // Create PackURI from bytes to avoid intermediate string allocation
            let candidate_str = std::str::from_utf8(&candidate_bytes).map_err(|_err| {
                OpcError::InvalidPackUri("Invalid UTF-8 in partname".to_string())
            })?;

            let candidate_uri = PackURI::new(candidate_str).map_err(OpcError::InvalidPackUri)?;
            if !self.parts.contains_key(&candidate_uri) {
                return Ok(candidate_uri);
            }

            n += 1;
            if n > 10000 {
                // Safety limit to prevent infinite loops
                return Err(OpcError::InvalidPackUri(
                    "Too many parts, cannot find next partname".to_string(),
                ));
            }
        }
    }

    /// Check if a part exists in the package.
    #[must_use]
    pub fn contains_part(&self, partname: &PackURI) -> bool {
        self.parts.contains_key(partname)
    }

    /// Atomically save the package to a file.
    ///
    /// Writes and synchronizes a finalized sibling artifact before replacing
    /// the destination. A failure before replacement leaves it untouched.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::package::OpcPackage;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized or the file cannot
    /// be written; the destination is left untouched on failure before replacement.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        crate::pkgwriter::PackageWriter::write(path, self)
    }

    /// Save the package to a stream.
    ///
    /// Writes the complete OPC package including all parts, relationships,
    /// and content types directly to a writer stream. A failure can leave the
    /// caller-owned stream incomplete; the error reports accepted bytes.
    ///
    /// # Arguments
    /// * `writer` - Any sequential writer; seeking is not required
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::package::OpcPackage;
    /// use std::fs::File;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// let file = File::create("output.docx")?;
    /// pkg.to_stream(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized or written to the
    /// stream; the stream may be left incomplete.
    pub fn to_stream<W: Write>(&self, writer: W) -> Result<()> {
        crate::pkgwriter::PackageWriter::write_to_stream(writer, self)
    }

    pub(crate) fn strip_signature_graph(&mut self) {
        let infrastructure: HashSet<PackURI> = self
            .parts
            .values()
            .filter(|part| is_signature_infrastructure(&***part))
            .map(|part| part.partname().clone())
            .collect();

        self.rels.retain(|relationship| {
            !is_signature_relationship(relationship.reltype())
                && !targets_any(relationship, &infrastructure)
        });
        for part in self.parts.values_mut() {
            part.rels_mut().retain(|relationship| {
                !is_signature_relationship(relationship.reltype())
                    && !targets_any(relationship, &infrastructure)
            });
        }
        self.parts
            .retain(|_, part| !is_signature_infrastructure(&**part));
    }
}

impl Default for OpcPackage {
    fn default() -> Self {
        Self::new()
    }
}

fn targets_any(relationship: &crate::Relationship, infrastructure: &HashSet<PackURI>) -> bool {
    if relationship.is_external() {
        return false;
    }
    match relationship.target_partname() {
        Ok(target) => infrastructure
            .iter()
            .any(|part| part.as_str().eq_ignore_ascii_case(target.as_str())),
        Err(_) => is_signature_path(relationship.target_path()),
    }
}

fn is_signature_relationship(kind: &str) -> bool {
    matches!(
        kind,
        relationship_type::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_signature_path(path: &str) -> bool {
    const DIRECTORY: &[u8] = b"/_xmlsignatures/";
    path.as_bytes()
        .get(..DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(DIRECTORY))
}

fn is_signature_infrastructure(part: &dyn Part) -> bool {
    use crate::constants::content_type;

    is_signature_path(part.partname().as_str())
        || matches!(
            part.content_type(),
            content_type::OPC_DIGITAL_SIGNATURE_ORIGIN
                | content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                | content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE
        )
}

fn part_name_conflict_error(
    existing: &PackURI,
    candidate: &PackURI,
    conflict: PartNameConflict,
) -> OpcError {
    match conflict {
        PartNameConflict::Duplicate => OpcError::DuplicatePartName(candidate.to_string()),
        PartNameConflict::Equivalent => OpcError::EquivalentPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
        PartNameConflict::Derived => OpcError::DerivedPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;
    use crate::part::BlobPart;
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::io::Cursor;
    use std::sync::Arc;

    fn create_minimal_docx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();

        // Add [Content_Types].xml
        writer
            .write_deflated(
                "[Content_Types].xml",
                br#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )
            .unwrap();

        // Add _rels/.rels
        writer
            .write_deflated(
                "_rels/.rels",
                br#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )
            .unwrap();

        // Add word/document.xml
        writer
            .write_deflated(
                "word/document.xml",
                br#"<?xml version="1.0"?>
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <body><p><t>Test</t></p></body>
</document>"#,
            )
            .unwrap();

        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn test_open_package() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor).unwrap();

        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn font_embedding_policy_has_only_three_valid_states() {
        let mut package = OpcPackage::new();
        assert_eq!(package.save_options().fonts, FontEmbedding::None);
        package.with_fonts(FontEmbedding::Subset);
        assert_eq!(package.save_options().fonts, FontEmbedding::Subset);
        package.with_fonts(FontEmbedding::Full);
        assert_eq!(package.save_options().fonts, FontEmbedding::Full);
    }

    #[test]
    fn moves_owned_archive_into_package_reader() {
        let pkg = OpcPackage::from_vec(create_minimal_docx()).unwrap();
        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn bounded_package_constructors_reject_oversized_input() {
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OpcPackage::from_bytes_with_limits(b"four", limits),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::InputBytes,
                actual: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn test_main_document_part() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor).unwrap();

        let main_part = pkg.main_document_part().unwrap();
        assert_eq!(
            main_part.content_type(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        );
    }

    #[test]
    fn resolves_one_internal_strict_main_document_part() {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/custom/main.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            "application/xml".to_string(),
            Vec::new(),
        )));
        package.relate_to("custom/main.xml", relationship_type::STRICT_OFFICE_DOCUMENT);

        assert_eq!(package.main_document_part().unwrap().partname(), &uri);

        package.relate_to("other.xml", relationship_type::OFFICE_DOCUMENT);
        assert!(package.main_document_part().is_err());
    }

    #[test]
    fn try_add_part_rejects_conflicts_without_replacing_the_original() {
        let mut package = OpcPackage::new();
        let original_uri = PackURI::new("/word/document.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                original_uri.clone(),
                "application/xml".to_string(),
                b"original".to_vec(),
            )))
            .unwrap();

        for (candidate, expected) in [
            ("/word/document.xml", PartNameConflict::Duplicate),
            ("/WORD/DOCUMENT.XML", PartNameConflict::Equivalent),
            ("/word/document.xml/image.gif", PartNameConflict::Derived),
        ] {
            let error = package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(candidate).unwrap(),
                    "application/octet-stream".to_string(),
                    b"replacement".to_vec(),
                )))
                .unwrap_err();
            assert!(matches!(
                (expected, error),
                (PartNameConflict::Duplicate, OpcError::DuplicatePartName(_))
                    | (
                        PartNameConflict::Equivalent,
                        OpcError::EquivalentPartNames { .. }
                    )
                    | (PartNameConflict::Derived, OpcError::DerivedPartNames { .. })
            ));
            assert_eq!(package.part_count(), 1);
            assert_eq!(package.get_part(&original_uri).unwrap().blob(), b"original");
        }
    }

    #[test]
    fn package_clone_shares_clean_payloads_and_detaches_mutation() {
        let uri = PackURI::new("/custom/data.bin").unwrap();
        let mut source = OpcPackage::new();
        source
            .try_add_part(Box::new(BlobPart::new(
                uri.clone(),
                "application/octet-stream".to_string(),
                b"source".to_vec(),
            )))
            .unwrap();

        let source_blob = source.get_part(&uri).unwrap().blob_arc();
        let mut edited = source.clone();
        let edited_blob = edited.get_part(&uri).unwrap().blob_arc();
        assert!(Arc::ptr_eq(&source_blob, &edited_blob));

        let replacement = Arc::new(b"edited".to_vec());
        edited
            .get_part_mut(&uri)
            .unwrap()
            .set_blob_shared(Arc::clone(&replacement));
        assert_eq!(source.get_part(&uri).unwrap().blob(), b"source");
        assert_eq!(edited.get_part(&uri).unwrap().blob(), b"edited");
        assert!(Arc::ptr_eq(
            &edited.get_part(&uri).unwrap().blob_arc(),
            &replacement
        ));
    }
}

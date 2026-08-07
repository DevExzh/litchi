//! OPC package adapter for the format-neutral `litchi-sign` engine.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashSet};

use litchi_sign::xml::{self, Canon, Profile, Transform};
use litchi_sign::{Coverage, Limits, Policy, Signer, Status, Trust};
use thiserror::Error as ThisError;

use crate::OpcPackage;
use crate::constants::{content_type, relationship_type};
use crate::error::OpcError;
use crate::packuri::PackURI;
use crate::part::{BlobPart, Part};
use crate::rel::{Relationship, Relationships, TargetMode};

pub(crate) const ORIGIN_REL: &str = relationship_type::DIGITAL_SIGNATURE_ORIGIN;
pub(crate) const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
pub(crate) const CERTIFICATE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

const ORIGIN_TYPE: &str = content_type::OPC_DIGITAL_SIGNATURE_ORIGIN;
const SIGNATURE_TYPE: &str = content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE;
const CERTIFICATE_TYPE: &str = content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE;
const RELATIONSHIPS_TYPE: &str = content_type::OPC_RELATIONSHIPS;
const SIGNATURE_DIR: &str = "/_xmlsignatures/";
const ORIGIN_NAME: &str = "/_xmlsignatures/origin.sigs";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Result type for OPC signature operations.
pub type Result<T> = std::result::Result<T, Error>;

/// OPC topology, resolution, or neutral-signature failures.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The package's signature relationships or parts are ambiguous or spoofed.
    #[error("invalid OPC signature graph: {0}")]
    Graph(String),

    /// Cryptographic or `XMLDSig` processing failed in the neutral engine.
    #[error(transparent)]
    Signature(#[from] litchi_sign::Error),

    /// Atomic staging could not update the OPC graph.
    #[error("could not update OPC signature graph: {0}")]
    Package(#[from] OpcError),

    /// Adding a signature is unsafe while an existing signature is invalid.
    #[error("existing signature {part} is invalid")]
    ExistingInvalid {
        /// Signature part that failed integrity or signature verification.
        part: PackURI,
    },
}

/// Verification report paired with its OPC signature part.
#[derive(Debug, Clone)]
pub struct Report {
    part: PackURI,
    details: litchi_sign::Report,
}

impl Report {
    /// Signature XML part URI.
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    /// Trust-neutral `XMLDSig` details.
    #[must_use]
    pub const fn details(&self) -> &litchi_sign::Report {
        &self.details
    }

    /// Whether every signed reference has the expected digest.
    #[must_use]
    pub fn integrity(&self) -> Status {
        self.details.integrity()
    }

    /// Whether the cryptographic signature matches its authenticated key.
    #[must_use]
    pub fn signature(&self) -> Status {
        self.details.signature()
    }

    /// Whether the signature covers every eligible OPC resource.
    #[must_use]
    pub fn coverage(&self) -> Coverage {
        self.details.coverage()
    }

    /// Certificate trust status, which remains neutral until evaluated by a caller.
    #[must_use]
    pub fn trust(&self) -> Trust {
        self.details.trust()
    }

    /// Whether any accepted signature or digest algorithm uses SHA-1.
    #[must_use]
    pub fn uses_sha1(&self) -> bool {
        self.details.uses_sha1()
    }

    /// Claimed signing time, when the signature contains one.
    #[must_use]
    pub fn time(&self) -> Option<&str> {
        self.details.time()
    }

    /// Separates the package URI from the neutral report.
    #[must_use]
    pub fn into_parts(self) -> (PackURI, litchi_sign::Report) {
        (self.part, self.details)
    }
}

/// Discovers and verifies every signature with an explicit policy.
pub(crate) fn signatures(package: &OpcPackage, policy: &Policy) -> Result<Vec<Report>> {
    let graph = Graph::read(package, policy.limits())?;
    verify_graph(package, &graph, policy)
}

fn verify_graph(package: &OpcPackage, graph: &Graph, policy: &Policy) -> Result<Vec<Report>> {
    if graph.signatures.is_empty() {
        return Ok(Vec::new());
    }
    let resolver = PackageResolver::new(package, policy.limits())?;
    let mut reports = Vec::new();
    reports
        .try_reserve(graph.signatures.len())
        .map_err(|_err| litchi_sign::Error::Limit("signature report allocation failed".into()))?;
    for signature in &graph.signatures {
        let part = package.get_part(&signature.part).map_err(graph_error)?;
        let mut certificates = Vec::new();
        certificates
            .try_reserve(signature.certificates.len())
            .map_err(|_err| limit("certificate reference allocation failed"))?;
        for uri in &signature.certificates {
            certificates.push(package.get_part(uri).map_err(graph_error)?.blob());
        }
        let details = xml::verify_with_certs(
            Profile::Package,
            part.blob(),
            &resolver,
            &certificates,
            policy,
        )?;
        reports.push(Report {
            part: signature.part.clone(),
            details,
        });
    }
    Ok(reports)
}

/// Adds one signature after verifying every existing signature.
pub(crate) fn sign(package: &mut OpcPackage, signer: &Signer, limits: &Limits) -> Result<PackURI> {
    let graph = Graph::read(package, limits)?;
    if !graph.signatures.is_empty() {
        let policy = Policy::strict().with_limits(limits.clone());
        for report in verify_graph(package, &graph, &policy)? {
            if report.details.integrity() != Status::Valid
                || report.details.signature() != Status::Valid
            {
                return Err(Error::ExistingInvalid { part: report.part });
            }
        }
    }

    let signature_uri = next_signature_uri(package)?;
    let signature_xml = author(package, signer, limits)?;
    let mut staged = package.clone();
    install(
        &mut staged,
        graph.origin.as_ref(),
        signature_uri.clone(),
        signature_xml,
    )?;
    Graph::read(&staged, limits)?;
    *package = staged;
    Ok(signature_uri)
}

/// Replaces a structurally valid signature graph with one new signature.
pub(crate) fn resign(
    package: &mut OpcPackage,
    signer: &Signer,
    limits: &Limits,
) -> Result<PackURI> {
    Graph::read(package, limits)?;
    let signature_uri = pack_uri("/_xmlsignatures/sig1.xml")?;
    let signature_xml = author(package, signer, limits)?;

    let mut staged = package.clone();
    staged.strip_signature_graph();
    install(&mut staged, None, signature_uri.clone(), signature_xml)?;
    Graph::read(&staged, limits)?;
    *package = staged;
    Ok(signature_uri)
}

fn author(package: &OpcPackage, signer: &Signer, limits: &Limits) -> Result<Vec<u8>> {
    let resolver = PackageResolver::new(package, limits)?;
    let references = resolver.into_author_references()?;
    Ok(xml::author(Profile::Package, signer, &references, limits)?)
}

#[derive(Debug)]
struct Graph {
    origin: Option<PackURI>,
    signatures: Vec<SignatureNode>,
}

#[derive(Debug)]
struct SignatureNode {
    part: PackURI,
    certificates: Vec<PackURI>,
}

impl Graph {
    fn read(package: &OpcPackage, limits: &Limits) -> Result<Self> {
        let mut origins = package
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == ORIGIN_REL);
        let origin_relationship = match origins.next() {
            None => {
                reject_orphan_graph(package)?;
                return Ok(Self {
                    origin: None,
                    signatures: Vec::new(),
                });
            },
            Some(origin) if origins.next().is_none() => origin,
            Some(_) => {
                let count = origins.count().saturating_add(2);
                return Err(Error::Graph(format!(
                    "expected one signature-origin relationship, found {count}"
                )));
            },
        };

        let origin_id = origin_relationship.r_id().to_string();
        let requested_origin = internal_target(origin_relationship, "signature origin")?;
        let origin_part = package.get_part(&requested_origin).map_err(graph_error)?;
        let origin = origin_part.partname().clone();
        require_signature_path(&origin, "signature origin")?;
        require_type(origin_part, ORIGIN_TYPE, "signature origin")?;

        for relationship in package.rels().iter() {
            if relationship.r_id() != origin_id && is_signature_relationship(relationship.reltype())
            {
                return Err(Error::Graph(format!(
                    "unexpected package-level signature relationship {}",
                    relationship.r_id()
                )));
            }
        }
        if origin_part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() != SIGNATURE_REL)
        {
            return Err(Error::Graph(
                "signature origin has an unexpected relationship".into(),
            ));
        }

        let signature_count = origin_part.rels().len();
        if signature_count == 0 {
            return Err(Error::Graph(
                "signature origin has no signature relationships".into(),
            ));
        }
        if signature_count > limits.max_signatures() {
            return Err(limit(format!(
                "OPC signature count {signature_count} exceeds limit {}",
                limits.max_signatures()
            )));
        }

        let mut signatures = Vec::new();
        signatures
            .try_reserve(signature_count)
            .map_err(|_err| limit("signature graph allocation failed"))?;
        let reachable_capacity = signature_count
            .checked_add(1)
            .ok_or_else(|| limit("signature graph capacity overflow"))?;
        let mut reachable = HashSet::new();
        reachable
            .try_reserve(reachable_capacity)
            .map_err(|_err| limit("signature graph allocation failed"))?;
        reachable.insert(origin.clone());
        let mut certificate_count = 0usize;
        let mut certificate_bytes = 0usize;
        for relationship in origin_part.rels().iter() {
            let requested = internal_target(relationship, "signature")?;
            let signature_part = package.get_part(&requested).map_err(graph_error)?;
            let signature_uri = signature_part.partname().clone();
            require_signature_path(&signature_uri, "signature")?;
            if signatures.iter().any(|signature: &SignatureNode| {
                signature
                    .part
                    .as_str()
                    .eq_ignore_ascii_case(signature_uri.as_str())
            }) {
                return Err(Error::Graph(format!(
                    "duplicate signature target {}",
                    signature_uri.as_str()
                )));
            }
            require_type(signature_part, SIGNATURE_TYPE, "signature")?;
            if signature_part.blob().len() > limits.max_signature_bytes() {
                return Err(limit(format!(
                    "signature part {} exceeds the signature byte limit",
                    signature_uri.as_str()
                )));
            }
            if signature_part
                .rels()
                .iter()
                .any(|relationship| relationship.reltype() != CERTIFICATE_REL)
            {
                return Err(Error::Graph(format!(
                    "signature part {} has an unexpected relationship",
                    signature_uri.as_str()
                )));
            }

            let mut certificates = Vec::new();
            let related_certificates = signature_part.rels().len();
            certificate_count = certificate_count
                .checked_add(related_certificates)
                .ok_or_else(|| limit("certificate count overflow"))?;
            if certificate_count > limits.max_certificates() {
                return Err(limit(format!(
                    "related certificate count {certificate_count} exceeds limit {}",
                    limits.max_certificates()
                )));
            }
            certificates
                .try_reserve(related_certificates)
                .map_err(|_err| limit("certificate graph allocation failed"))?;
            for certificate_relationship in signature_part.rels().iter() {
                let requested = internal_target(certificate_relationship, "certificate")?;
                let certificate_part = package.get_part(&requested).map_err(graph_error)?;
                let certificate_uri = certificate_part.partname().clone();
                require_signature_path(&certificate_uri, "certificate")?;
                if certificates.iter().any(|existing: &PackURI| {
                    existing
                        .as_str()
                        .eq_ignore_ascii_case(certificate_uri.as_str())
                }) {
                    return Err(Error::Graph(format!(
                        "duplicate certificate target {}",
                        certificate_uri.as_str()
                    )));
                }
                require_type(certificate_part, CERTIFICATE_TYPE, "certificate")?;
                let bytes = certificate_part.blob().len();
                if bytes > limits.max_certificate_bytes() {
                    return Err(limit(format!(
                        "certificate part {} exceeds the per-certificate byte limit",
                        certificate_uri.as_str()
                    )));
                }
                certificate_bytes = certificate_bytes
                    .checked_add(bytes)
                    .ok_or_else(|| limit("certificate byte count overflow"))?;
                if certificate_bytes > limits.max_total_certificate_bytes() {
                    return Err(limit(format!(
                        "related certificate bytes exceed limit {}",
                        limits.max_total_certificate_bytes()
                    )));
                }
                if !certificate_part.rels().is_empty() {
                    return Err(Error::Graph(format!(
                        "certificate part {} owns relationships",
                        certificate_uri.as_str()
                    )));
                }
                reachable
                    .try_reserve(1)
                    .map_err(|_err| limit("signature graph allocation failed"))?;
                reachable.insert(certificate_uri.clone());
                certificates.push(certificate_uri);
            }
            certificates.sort_by(|left, right| left.as_str().cmp(right.as_str()));
            reachable.insert(signature_uri.clone());
            signatures.push(SignatureNode {
                part: signature_uri,
                certificates,
            });
        }
        signatures.sort_by(|left, right| left.part.as_str().cmp(right.part.as_str()));

        for part in package.iter_parts() {
            if is_infrastructure(part) && !reachable.contains(part.partname()) {
                return Err(Error::Graph(format!(
                    "orphan or spoofed signature infrastructure part {}",
                    part.partname().as_str()
                )));
            }
        }
        reject_inbound_spoofs(package, &reachable, &origin_id)?;

        Ok(Self {
            origin: Some(origin),
            signatures,
        })
    }
}

fn reject_orphan_graph(package: &OpcPackage) -> Result<()> {
    if let Some(part) = package.iter_parts().find(|part| is_infrastructure(*part)) {
        return Err(Error::Graph(format!(
            "orphan or spoofed signature infrastructure part {}",
            part.partname().as_str()
        )));
    }
    if let Some(relationship) = package
        .rels()
        .iter()
        .find(|relationship| is_signature_relationship(relationship.reltype()))
    {
        return Err(Error::Graph(format!(
            "unexpected package-level signature relationship {}",
            relationship.r_id()
        )));
    }
    for part in package.iter_parts() {
        if let Some(relationship) = part
            .rels()
            .iter()
            .find(|relationship| is_signature_relationship(relationship.reltype()))
        {
            return Err(Error::Graph(format!(
                "part {} owns stray signature relationship {}",
                part.partname().as_str(),
                relationship.r_id()
            )));
        }
    }
    Ok(())
}

fn reject_inbound_spoofs(
    package: &OpcPackage,
    reachable: &HashSet<PackURI>,
    origin_id: &str,
) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.r_id() != origin_id && targets_graph(relationship, reachable)? {
            return Err(Error::Graph(format!(
                "package relationship {} points into signature infrastructure",
                relationship.r_id()
            )));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| !reachable.contains(part.partname()))
    {
        for relationship in part.rels().iter() {
            if is_signature_relationship(relationship.reltype()) {
                return Err(Error::Graph(format!(
                    "part {} owns stray signature relationship {}",
                    part.partname().as_str(),
                    relationship.r_id()
                )));
            }
            if targets_graph(relationship, reachable)? {
                return Err(Error::Graph(format!(
                    "part {} points into signature infrastructure",
                    part.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn targets_graph(relationship: &Relationship, reachable: &HashSet<PackURI>) -> Result<bool> {
    if relationship.is_external() {
        return Ok(false);
    }
    match relationship.target_partname() {
        Ok(target) => Ok(reachable
            .iter()
            .any(|part| part.as_str().eq_ignore_ascii_case(target.as_str()))),
        Err(error) if signature_hint(relationship.target_path()) => Err(graph_error(error)),
        Err(_) => Ok(false),
    }
}

fn internal_target(relationship: &Relationship, description: &str) -> Result<PackURI> {
    if relationship.target_mode() != TargetMode::Internal {
        return Err(Error::Graph(format!(
            "{description} relationship must be internal"
        )));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Graph(format!(
            "{description} relationship target cannot contain a query or fragment"
        )));
    }
    relationship.target_partname().map_err(graph_error)
}

fn require_type(part: &dyn Part, expected: &str, description: &str) -> Result<()> {
    if part.content_type() != expected {
        return Err(Error::Graph(format!(
            "{description} part {} has content type {}, expected {expected}",
            part.partname().as_str(),
            part.content_type()
        )));
    }
    Ok(())
}

fn require_signature_path(uri: &PackURI, description: &str) -> Result<()> {
    if !is_signature_path(uri.as_str()) {
        return Err(Error::Graph(format!(
            "{description} part must be under {SIGNATURE_DIR}"
        )));
    }
    Ok(())
}

pub(crate) fn is_signature_relationship(kind: &str) -> bool {
    matches!(kind, ORIGIN_REL | SIGNATURE_REL | CERTIFICATE_REL)
}

pub(crate) fn is_signature_path(path: &str) -> bool {
    path.as_bytes()
        .get(..SIGNATURE_DIR.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(SIGNATURE_DIR.as_bytes()))
}

fn signature_hint(path: &str) -> bool {
    const HINT: &[u8] = b"_xmlsignatures";
    path.as_bytes()
        .windows(HINT.len())
        .any(|window| window.eq_ignore_ascii_case(HINT))
}

pub(crate) fn is_infrastructure(part: &dyn Part) -> bool {
    is_signature_path(part.partname().as_str())
        || matches!(
            part.content_type(),
            ORIGIN_TYPE | SIGNATURE_TYPE | CERTIFICATE_TYPE
        )
}

fn graph_error(error: impl std::fmt::Display) -> Error {
    Error::Graph(error.to_string())
}

fn limit(message: impl Into<String>) -> Error {
    litchi_sign::Error::Limit(message.into()).into()
}

fn pack_uri(value: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(graph_error)
}

enum Resource<'a> {
    Part(&'a dyn Part),
    Relationships {
        relationships: &'a Relationships,
        ids: Vec<String>,
    },
}

struct PackageResolver<'a> {
    entries: BTreeMap<String, Resource<'a>>,
    limits: &'a Limits,
}

impl<'a> PackageResolver<'a> {
    fn new(package: &'a OpcPackage, limits: &'a Limits) -> Result<Self> {
        let shape = resolver_shape(package, limits)?;

        let mut resolver = Self {
            entries: BTreeMap::new(),
            limits,
        };
        let mut parts = Vec::new();
        parts
            .try_reserve(shape.parts)
            .map_err(|_err| limit("OPC part reference allocation failed"))?;
        parts.extend(
            package
                .iter_parts()
                .filter(|part| !is_infrastructure(*part)),
        );
        parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));

        for part in &parts {
            let uri = part_reference_uri(
                part.partname().as_str(),
                part.content_type(),
                limits.max_signature_bytes(),
            )?;
            resolver.insert(uri, Resource::Part(*part))?;
        }
        resolver.insert_relationships(None, package.rels())?;
        for part in parts {
            resolver.insert_relationships(Some(part.partname().as_str()), part.rels())?;
        }
        if resolver.entries.len() != shape.references {
            return Err(Error::Graph(
                "OPC signature reference shape changed during resolution".into(),
            ));
        }
        Ok(resolver)
    }

    fn insert_relationships(
        &mut self,
        source: Option<&str>,
        relationships: &'a Relationships,
    ) -> Result<()> {
        let (count, _) = eligible_relationship_summary(relationships)?;
        if count == 0 {
            return Ok(());
        }
        if self.entries.len() >= self.limits.max_references() {
            return Err(limit("OPC reference count exceeds policy"));
        }
        let uri = relationship_reference_uri(source, self.limits.max_signature_bytes())?;
        let ids = eligible_relationship_ids(relationships, count)?;
        self.insert(uri, Resource::Relationships { relationships, ids })
    }

    fn insert(&mut self, uri: String, resource: Resource<'a>) -> Result<()> {
        if self.entries.len() >= self.limits.max_references() {
            return Err(
                litchi_sign::Error::Limit("OPC reference count exceeds policy".into()).into(),
            );
        }
        if self.entries.contains_key(&uri) {
            return Err(Error::Graph(format!(
                "duplicate eligible signature reference {uri}"
            )));
        }
        self.entries.insert(uri, resource);
        Ok(())
    }

    fn into_author_references(self) -> Result<Vec<xml::Ref<'a>>> {
        let mut references = Vec::new();
        references
            .try_reserve(self.entries.len())
            .map_err(|_err| litchi_sign::Error::Limit("reference allocation failed".into()))?;
        for (uri, resource) in self.entries {
            let reference = match resource {
                Resource::Part(part) => xml::Ref::borrowed_uri(uri, part.blob())?,
                Resource::Relationships { relationships, ids } => {
                    let data = relationship_xml(relationships, &ids, self.limits)?;
                    xml::Ref::owned(uri, data)?
                        .transform(Transform::Relationships(ids))
                        .transform(Transform::Canon(Canon::Inclusive))
                },
            };
            references.push(reference);
        }
        Ok(references)
    }
}

impl xml::Resolver for PackageResolver<'_> {
    fn expected(&self) -> usize {
        self.entries.len()
    }

    fn has(&self, uri: &str) -> bool {
        self.entries.contains_key(uri)
    }

    fn get<'a>(
        &'a self,
        uri: &str,
        transforms: &[Transform],
    ) -> litchi_sign::Result<(Cow<'a, [u8]>, Coverage)> {
        let resource = self.entries.get(uri).ok_or_else(|| {
            litchi_sign::Error::Container(format!("Manifest references unexpected URI {uri}"))
        })?;
        match resource {
            Resource::Part(part) => match transforms {
                [] => Ok((Cow::Borrowed(part.blob()), Coverage::Complete)),
                [Transform::Canon(canon)] => Ok((
                    Cow::Owned(xml::canonicalize(part.blob(), *canon, self.limits)?),
                    Coverage::Complete,
                )),
                _ => Err(litchi_sign::Error::Container(format!(
                    "invalid transform chain for package part {uri}"
                ))),
            },
            Resource::Relationships { relationships, ids } => {
                let (selected, canon) = match transforms {
                    [Transform::Relationships(selected)] => (selected, None),
                    [Transform::Relationships(selected), Transform::Canon(canon)] => {
                        (selected, Some(*canon))
                    },
                    _ => {
                        return Err(litchi_sign::Error::Container(format!(
                            "invalid RelationshipTransform chain for {uri}"
                        )));
                    },
                };
                if selected.len() > self.limits.max_references() {
                    return Err(litchi_sign::Error::Limit(
                        "RelationshipTransform selection exceeds the reference limit".into(),
                    ));
                }
                if selected.iter().any(|id| ids.binary_search(id).is_err()) {
                    return Err(litchi_sign::Error::Container(format!(
                        "RelationshipTransform for {uri} selects an unknown relationship"
                    )));
                }
                let coverage = if selected == ids {
                    Coverage::Complete
                } else {
                    Coverage::Partial
                };
                let mut data = relationship_xml(relationships, selected, self.limits)?;
                if let Some(canon) = canon {
                    data = xml::canonicalize(&data, canon, self.limits)?;
                }
                Ok((Cow::Owned(data), coverage))
            },
        }
    }
}

#[derive(Default)]
struct ResolverShape {
    parts: usize,
    references: usize,
    relationship_ids: usize,
    metadata_bytes: usize,
}

impl ResolverShape {
    fn add_reference(
        &mut self,
        uri_bytes: usize,
        relationship_ids: usize,
        relationship_id_bytes: usize,
        limits: &Limits,
    ) -> Result<()> {
        self.references = self
            .references
            .checked_add(1)
            .ok_or_else(|| limit("OPC reference count overflow"))?;
        if self.references > limits.max_references() {
            return Err(limit(format!(
                "OPC reference count {} exceeds limit {}",
                self.references,
                limits.max_references()
            )));
        }
        self.relationship_ids = self
            .relationship_ids
            .checked_add(relationship_ids)
            .ok_or_else(|| limit("relationship selection count overflow"))?;
        if self.relationship_ids > limits.max_references() {
            return Err(limit(format!(
                "relationship selection count {} exceeds limit {}",
                self.relationship_ids,
                limits.max_references()
            )));
        }
        let added_bytes = uri_bytes
            .checked_add(relationship_id_bytes)
            .ok_or_else(|| limit("OPC reference metadata size overflow"))?;
        self.metadata_bytes = self
            .metadata_bytes
            .checked_add(added_bytes)
            .ok_or_else(|| limit("OPC reference metadata size overflow"))?;
        if self.metadata_bytes > limits.max_signature_bytes() {
            return Err(limit(format!(
                "OPC reference metadata exceeds the signature byte limit of {}",
                limits.max_signature_bytes()
            )));
        }
        Ok(())
    }
}

fn resolver_shape(package: &OpcPackage, limits: &Limits) -> Result<ResolverShape> {
    let part_count = package
        .iter_parts()
        .filter(|part| !is_infrastructure(*part))
        .count();
    if part_count > limits.max_references() {
        return Err(limit(format!(
            "OPC part reference count {part_count} exceeds limit {}",
            limits.max_references()
        )));
    }

    let mut shape = ResolverShape {
        parts: part_count,
        ..ResolverShape::default()
    };
    for part in package
        .iter_parts()
        .filter(|part| !is_infrastructure(*part))
    {
        shape.add_reference(
            part_reference_len(part.partname().as_str(), part.content_type())?,
            0,
            0,
            limits,
        )?;
        let (count, bytes) = eligible_relationship_summary(part.rels())?;
        if count != 0 {
            shape.add_reference(
                relationship_reference_len(Some(part.partname().as_str()))?,
                count,
                bytes,
                limits,
            )?;
        }
    }
    let (count, bytes) = eligible_relationship_summary(package.rels())?;
    if count != 0 {
        shape.add_reference(relationship_reference_len(None)?, count, bytes, limits)?;
    }
    Ok(shape)
}

fn eligible_relationship_summary(relationships: &Relationships) -> Result<(usize, usize)> {
    let mut count = 0usize;
    let mut bytes = 0usize;
    for relationship in relationships
        .iter()
        .filter(|relationship| !is_signature_relationship(relationship.reltype()))
    {
        count = count
            .checked_add(1)
            .ok_or_else(|| limit("relationship selection count overflow"))?;
        bytes = bytes
            .checked_add(relationship.r_id().len())
            .ok_or_else(|| limit("relationship identifier size overflow"))?;
    }
    Ok((count, bytes))
}

fn eligible_relationship_ids(relationships: &Relationships, count: usize) -> Result<Vec<String>> {
    let mut ids = Vec::new();
    ids.try_reserve(count)
        .map_err(|_err| limit("relationship selection allocation failed"))?;
    for relationship in relationships
        .iter()
        .filter(|relationship| !is_signature_relationship(relationship.reltype()))
    {
        ids.push(copy_string(relationship.r_id(), "relationship identifier")?);
    }
    ids.sort();
    Ok(ids)
}

fn copy_string(value: &str, description: &str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve(value.len())
        .map_err(|_err| limit(format!("{description} allocation failed")))?;
    copy.push_str(value);
    Ok(copy)
}

fn part_reference_len(part: &str, content_type: &str) -> Result<usize> {
    joined_len(&[part, "?ContentType=", content_type], "part reference URI")
}

fn part_reference_uri(part: &str, content_type: &str, maximum: usize) -> Result<String> {
    bounded_join(
        &[part, "?ContentType=", content_type],
        maximum,
        "part reference URI",
    )
}

fn relationship_source(source: &str) -> Result<(&str, &str)> {
    let (directory, file) = source
        .rsplit_once('/')
        .ok_or_else(|| Error::Graph(format!("invalid relationship source part {source}")))?;
    if file.is_empty() {
        return Err(Error::Graph(format!(
            "invalid relationship source part {source}"
        )));
    }
    Ok((directory, file))
}

fn relationship_reference_len(source: Option<&str>) -> Result<usize> {
    match source {
        None => joined_len(
            &["/_rels/.rels?ContentType=", RELATIONSHIPS_TYPE],
            "relationship reference URI",
        ),
        Some(source) => {
            let (directory, file) = relationship_source(source)?;
            joined_len(
                &[
                    directory,
                    "/_rels/",
                    file,
                    ".rels?ContentType=",
                    RELATIONSHIPS_TYPE,
                ],
                "relationship reference URI",
            )
        },
    }
}

fn relationship_reference_uri(source: Option<&str>, maximum: usize) -> Result<String> {
    match source {
        None => bounded_join(
            &["/_rels/.rels?ContentType=", RELATIONSHIPS_TYPE],
            maximum,
            "relationship reference URI",
        ),
        Some(source) => {
            let (directory, file) = relationship_source(source)?;
            bounded_join(
                &[
                    directory,
                    "/_rels/",
                    file,
                    ".rels?ContentType=",
                    RELATIONSHIPS_TYPE,
                ],
                maximum,
                "relationship reference URI",
            )
        },
    }
}

fn joined_len(components: &[&str], description: &str) -> Result<usize> {
    components.iter().try_fold(0usize, |length, component| {
        length
            .checked_add(component.len())
            .ok_or_else(|| limit(format!("{description} size overflow")))
    })
}

fn bounded_join(components: &[&str], maximum: usize, description: &str) -> Result<String> {
    let length = joined_len(components, description)?;
    if length > maximum {
        return Err(limit(format!(
            "{description} exceeds the signature byte limit"
        )));
    }
    let mut output = String::new();
    output
        .try_reserve(length)
        .map_err(|_err| limit(format!("{description} allocation failed")))?;
    for component in components {
        output.push_str(component);
    }
    Ok(output)
}

fn relationship_xml(
    relationships: &Relationships,
    ids: &[String],
    limits: &Limits,
) -> litchi_sign::Result<Vec<u8>> {
    let maximum = limits.max_signature_bytes();
    let mut output = Vec::new();
    push(&mut output, b"<Relationships xmlns=\"", maximum)?;
    push_attr(&mut output, RELATIONSHIPS_NS, maximum)?;
    push(&mut output, b"\">", maximum)?;
    for id in ids {
        let relationship = relationships.get(id).ok_or_else(|| {
            litchi_sign::Error::Container(format!("relationship {id} does not exist"))
        })?;
        push(&mut output, b"<Relationship Id=\"", maximum)?;
        push_attr(&mut output, relationship.r_id(), maximum)?;
        push(&mut output, b"\" Target=\"", maximum)?;
        push_attr(&mut output, relationship.target_ref(), maximum)?;
        push(&mut output, b"\" TargetMode=\"", maximum)?;
        push(
            &mut output,
            if relationship.target_mode() == TargetMode::Internal {
                b"Internal"
            } else {
                b"External"
            },
            maximum,
        )?;
        push(&mut output, b"\" Type=\"", maximum)?;
        push_attr(&mut output, relationship.reltype(), maximum)?;
        push(&mut output, b"\"></Relationship>", maximum)?;
    }
    push(&mut output, b"</Relationships>", maximum)?;
    Ok(output)
}

fn push(output: &mut Vec<u8>, bytes: &[u8], maximum: usize) -> litchi_sign::Result<()> {
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| litchi_sign::Error::Limit("relationship XML size overflow".into()))?;
    if new_len > maximum {
        return Err(litchi_sign::Error::Limit(
            "relationship transform exceeds the signature byte limit".into(),
        ));
    }
    output
        .try_reserve(bytes.len())
        .map_err(|_err| litchi_sign::Error::Limit("relationship XML allocation failed".into()))?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn push_attr(output: &mut Vec<u8>, value: &str, maximum: usize) -> litchi_sign::Result<()> {
    for character in value.chars() {
        match character {
            '&' => push(output, b"&amp;", maximum)?,
            '<' => push(output, b"&lt;", maximum)?,
            '"' => push(output, b"&quot;", maximum)?,
            '\t' => push(output, b"&#x9;", maximum)?,
            '\n' => push(output, b"&#xA;", maximum)?,
            '\r' => push(output, b"&#xD;", maximum)?,
            character => {
                let mut encoded = [0; 4];
                push(
                    output,
                    character.encode_utf8(&mut encoded).as_bytes(),
                    maximum,
                )?;
            },
        }
    }
    Ok(())
}

fn next_signature_uri(package: &OpcPackage) -> Result<PackURI> {
    let candidates = package
        .part_count()
        .checked_add(1)
        .ok_or_else(|| Error::Graph("part count overflow".into()))?;
    for index in 1..=candidates {
        let candidate =
            PackURI::new(format!("/_xmlsignatures/sig{index}.xml")).map_err(graph_error)?;
        if !package.contains_part(&candidate) {
            return Ok(candidate);
        }
    }
    Err(Error::Graph("no signature part name is available".into()))
}

fn install(
    package: &mut OpcPackage,
    existing_origin: Option<&PackURI>,
    signature_uri: PackURI,
    signature_xml: Vec<u8>,
) -> Result<()> {
    package.try_add_part(Box::new(BlobPart::new(
        signature_uri.clone(),
        SIGNATURE_TYPE.to_string(),
        signature_xml,
    )))?;

    let origin = match existing_origin {
        Some(origin) => origin.clone(),
        None => {
            let origin = pack_uri(ORIGIN_NAME)?;
            package.try_add_part(Box::new(BlobPart::new(
                origin.clone(),
                ORIGIN_TYPE.to_string(),
                Vec::new(),
            )))?;
            let id = next_relationship_id(package.rels())?;
            package.rels_mut().try_add_relationship(
                ORIGIN_REL.to_string(),
                origin.as_str().to_string(),
                id,
                TargetMode::Internal,
            )?;
            origin
        },
    };

    let id = {
        let origin_part = package.get_part(&origin).map_err(graph_error)?;
        next_relationship_id(origin_part.rels())?
    };
    package
        .get_part_mut(&origin)
        .map_err(graph_error)?
        .rels_mut()
        .try_add_relationship(
            SIGNATURE_REL.to_string(),
            signature_uri.as_str().to_string(),
            id,
            TargetMode::Internal,
        )?;
    Ok(())
}

fn next_relationship_id(relationships: &Relationships) -> Result<String> {
    let candidates = relationships
        .len()
        .checked_add(1)
        .ok_or_else(|| Error::Graph("relationship count overflow".into()))?;
    for index in 1..=candidates {
        let id = format!("rId{index}");
        if relationships.get(&id).is_none() {
            return Ok(id);
        }
    }
    Err(Error::Graph(
        "no relationship identifier is available".into(),
    ))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use p256::ecdsa::SigningKey;

    use super::*;

    const POI_DOCX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.docx");
    const POI_XLSX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.xlsx");
    const POI_PPTX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx");
    const POI_TWICE: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed-twice.docx");
    const MICROSOFT_DOCX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx");
    const MICROSOFT_XLSX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.xlsx");
    const MICROSOFT_PPTX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/xmldsign/ms-office-2010-signed.pptx");

    fn signer(seed: u8) -> Signer {
        let key = SigningKey::from_bytes((&[seed; 32]).into()).expect("valid fixture key");
        Signer::p256(key)
            .time("2026-07-19T12:34:56Z")
            .expect("valid fixture time")
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/document.xml").expect("URI"),
                "application/xml".into(),
                b"<document><value>signed</value></document>".to_vec(),
            )))
            .expect("part");
        package
            .rels_mut()
            .try_add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    .into(),
                "/document.xml".into(),
                "rId1".into(),
                TargetMode::Internal,
            )
            .expect("relationship");
        package
    }

    fn fixture(bytes: &[u8], count: usize) {
        let package = OpcPackage::from_bytes(bytes).expect("fixture opens");
        let reports = package
            .signatures_with(&Policy::compatible())
            .expect("fixture verifies");
        assert_eq!(reports.len(), count);
        for report in reports {
            assert_eq!(report.integrity(), Status::Valid);
            assert_eq!(report.signature(), Status::Valid);
            assert_eq!(report.coverage(), Coverage::Partial);
            assert_eq!(report.trust(), Trust::NotChecked);
            assert!(report.uses_sha1());
            assert!(!report.details().certificates().is_empty());
        }
    }

    #[test]
    fn verifies_real_poi_and_microsoft_office_fixtures() {
        fixture(POI_DOCX, 1);
        fixture(POI_XLSX, 1);
        fixture(POI_PPTX, 1);
        fixture(POI_TWICE, 2);
        fixture(MICROSOFT_DOCX, 1);
        fixture(MICROSOFT_XLSX, 1);
        fixture(MICROSOFT_PPTX, 1);
    }

    #[test]
    fn safe_default_rejects_fixture_sha1() {
        let package = OpcPackage::from_bytes(POI_DOCX).expect("fixture opens");
        assert!(matches!(
            package.signatures(),
            Err(Error::Signature(litchi_sign::Error::Sha1))
        ));
    }

    #[test]
    fn signs_resigns_and_unsigns_with_deterministic_names() {
        let mut package = package();
        assert!(!package.is_signed());
        assert_eq!(
            package.sign(&signer(7)).expect("first").as_str(),
            "/_xmlsignatures/sig1.xml"
        );
        assert_eq!(
            package.sign(&signer(8)).expect("second").as_str(),
            "/_xmlsignatures/sig2.xml"
        );
        let reports = package.signatures().expect("strict verification");
        assert_eq!(reports.len(), 2);
        assert!(reports.iter().all(|report| {
            report.integrity() == Status::Valid
                && report.signature() == Status::Valid
                && report.coverage() == Coverage::Complete
        }));

        assert_eq!(
            package.resign(&signer(9)).expect("resign").as_str(),
            "/_xmlsignatures/sig1.xml"
        );
        assert_eq!(package.signatures().expect("resigned").len(), 1);
        package.unsign();
        package.unsign();
        assert!(!package.is_signed());
        assert!(package.signatures().expect("unsigned").is_empty());
    }

    #[test]
    fn compatibility_reports_partial_coverage_and_strict_rejects_it() {
        let mut missing_part = package();
        missing_part.sign(&signer(7)).expect("sign");
        missing_part
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/unsigned.bin").expect("URI"),
                "application/octet-stream".into(),
                b"unsigned".to_vec(),
            )))
            .expect("part");
        let report = missing_part
            .signatures_with(&Policy::compatible())
            .expect("compatible")
            .remove(0);
        assert_eq!(report.coverage(), Coverage::Partial);
        assert!(matches!(
            missing_part.signatures(),
            Err(Error::Signature(litchi_sign::Error::Container(message)))
                if message.contains("partial")
        ));
        let before = missing_part.part_count();
        assert!(missing_part.sign(&signer(8)).is_err());
        assert_eq!(missing_part.part_count(), before);

        let mut relationships = package();
        relationships.sign(&signer(7)).expect("sign");
        relationships
            .rels_mut()
            .try_add_relationship(
                "urn:test:unsigned".into(),
                "https://example.invalid".into(),
                "rIdUnsigned".into(),
                TargetMode::External,
            )
            .expect("relationship");
        let report = relationships
            .signatures_with(&Policy::compatible())
            .expect("compatible")
            .remove(0);
        assert_eq!(report.coverage(), Coverage::Partial);
        assert!(relationships.signatures().is_err());

        let mut limited = package();
        limited.sign(&signer(7)).expect("sign");
        let limits = Limits::standard().references(1).expect("limits");
        let before = limited.part_count();
        assert!(limited.sign_with(&signer(8), &limits).is_err());
        assert_eq!(limited.part_count(), before);
    }

    #[test]
    fn resource_limits_reject_hostile_signature_graphs_atomically() {
        let mut signatures = package();
        let first = signatures.sign(&signer(7)).expect("first signature");
        let duplicate_xml = signatures
            .get_part(&first)
            .expect("first signature part")
            .blob()
            .to_vec();
        let second = PackURI::new("/_xmlsignatures/sig2.xml").expect("URI");
        signatures
            .try_add_part(Box::new(BlobPart::new(
                second.clone(),
                SIGNATURE_TYPE.into(),
                duplicate_xml,
            )))
            .expect("second signature part");
        let origin = PackURI::new(ORIGIN_NAME).expect("URI");
        signatures
            .get_part_mut(&origin)
            .expect("origin")
            .rels_mut()
            .try_add_relationship(
                SIGNATURE_REL.into(),
                second.as_str().into(),
                "rId2".into(),
                TargetMode::Internal,
            )
            .expect("second signature relationship");

        let one_signature = Limits::standard().signatures(1).expect("limits");
        let policy = Policy::compatible().with_limits(one_signature.clone());
        assert!(matches!(
            signatures.signatures_with(&policy),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("signature count")
        ));
        let before_parts = signatures.part_count();
        let before_origin_relationships =
            signatures.get_part(&origin).expect("origin").rels().len();
        assert!(signatures.sign_with(&signer(8), &one_signature).is_err());
        assert!(signatures.resign_with(&signer(8), &one_signature).is_err());
        assert_eq!(signatures.part_count(), before_parts);
        assert_eq!(
            signatures.get_part(&origin).expect("origin").rels().len(),
            before_origin_relationships
        );

        let mut certificates = package();
        let signature = certificates.sign(&signer(9)).expect("signature");
        for index in 1..=2 {
            let uri = PackURI::new(format!("/_xmlsignatures/cert{index}.cer")).expect("URI");
            certificates
                .try_add_part(Box::new(BlobPart::new(
                    uri,
                    CERTIFICATE_TYPE.into(),
                    vec![index as u8; 2],
                )))
                .expect("certificate part");
            certificates
                .get_part_mut(&signature)
                .expect("signature part")
                .rels_mut()
                .try_add_relationship(
                    CERTIFICATE_REL.into(),
                    format!("cert{index}.cer"),
                    format!("rIdCertificate{index}"),
                    TargetMode::Internal,
                )
                .expect("certificate relationship");
        }

        let one_certificate = Limits::standard().certificates(1).expect("limits");
        let policy = Policy::compatible().with_limits(one_certificate);
        assert!(matches!(
            certificates.signatures_with(&policy),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("certificate count")
        ));

        let three_certificate_bytes = Limits::standard()
            .total_certificate_bytes(3)
            .expect("limits");
        let policy = Policy::compatible().with_limits(three_certificate_bytes);
        assert!(matches!(
            certificates.signatures_with(&policy),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("certificate bytes")
        ));

        let mut parts = package();
        for index in 1..=2 {
            parts
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(format!("/extra{index}.bin")).expect("URI"),
                    "application/octet-stream".into(),
                    vec![index as u8],
                )))
                .expect("extra part");
        }
        let two_references = Limits::standard().references(2).expect("limits");
        let before_parts = parts.part_count();
        assert!(matches!(
            parts.sign_with(&signer(10), &two_references),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("part reference count")
        ));
        assert_eq!(parts.part_count(), before_parts);
        assert!(!parts.is_signed());

        let mut relationships = package();
        for index in 2..=3 {
            relationships
                .rels_mut()
                .try_add_relationship(
                    "urn:test:hostile".into(),
                    format!("https://example.invalid/{index}"),
                    format!("rId{index}"),
                    TargetMode::External,
                )
                .expect("hostile relationship");
        }
        let before_relationships = relationships.rels().len();
        assert!(matches!(
            relationships.sign_with(&signer(11), &two_references),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("relationship selection count")
        ));
        assert_eq!(relationships.rels().len(), before_relationships);
        assert!(!relationships.is_signed());

        let mut metadata = package();
        metadata
            .rels_mut()
            .try_add_relationship(
                "urn:test:hostile".into(),
                "https://example.invalid/metadata".into(),
                "r".repeat(256),
                TargetMode::External,
            )
            .expect("large relationship identifier");
        let compact_metadata = Limits::standard().signature_bytes(128).expect("limits");
        let before_relationships = metadata.rels().len();
        assert!(matches!(
            metadata.sign_with(&signer(12), &compact_metadata),
            Err(Error::Signature(litchi_sign::Error::Limit(message)))
                if message.contains("reference metadata")
        ));
        assert_eq!(metadata.rels().len(), before_relationships);
        assert!(!metadata.is_signed());
    }

    #[test]
    fn verifies_certificate_from_a_related_der_part() {
        let mut package = OpcPackage::from_bytes(POI_DOCX).expect("fixture opens");
        let initial = package
            .signatures_with(&Policy::compatible())
            .expect("initial verification");
        let signature_uri = initial[0].part().clone();
        let certificate = initial[0].details().certificates()[0].der().to_vec();
        let xml = std::str::from_utf8(package.get_part(&signature_uri).expect("signature").blob())
            .expect("UTF-8 signature");
        let start = xml.find("<X509Data>").expect("X509Data");
        let end =
            xml[start..].find("</X509Data>").expect("X509Data end") + start + "</X509Data>".len();
        let mut external_only = Vec::with_capacity(xml.len() - (end - start));
        external_only.extend_from_slice(&xml.as_bytes()[..start]);
        external_only.extend_from_slice(&xml.as_bytes()[end..]);
        let certificate_uri = PackURI::new("/_xmlsignatures/cert1.cer").expect("URI");
        package
            .try_add_part(Box::new(BlobPart::new(
                certificate_uri,
                CERTIFICATE_TYPE.into(),
                certificate,
            )))
            .expect("certificate part");
        let signature = package
            .get_part_mut(&signature_uri)
            .expect("signature part");
        signature.set_blob(external_only);
        signature
            .rels_mut()
            .try_add_relationship(
                CERTIFICATE_REL.into(),
                "cert1.cer".into(),
                "rIdCertificate".into(),
                TargetMode::Internal,
            )
            .expect("certificate relationship");

        let report = package
            .signatures_with(&Policy::compatible())
            .expect("external certificate verifies")
            .remove(0);
        assert_eq!(report.signature(), Status::Valid);
        assert_eq!(report.details().certificates().len(), 1);
    }

    #[test]
    fn rejects_spoofs_duplicates_and_keeps_failed_edits_atomic() {
        let mut spoofed = package();
        spoofed
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(ORIGIN_NAME).expect("URI"),
                "application/octet-stream".into(),
                b"preserve".to_vec(),
            )))
            .expect("spoof part");
        let before = spoofed.clone();
        assert!(spoofed.sign(&signer(7)).is_err());
        assert_eq!(spoofed.part_count(), before.part_count());
        assert_eq!(
            spoofed
                .get_part(&PackURI::new(ORIGIN_NAME).expect("URI"))
                .expect("preserved")
                .blob(),
            b"preserve"
        );

        let mut duplicate = package();
        duplicate.sign(&signer(7)).expect("sign");
        let origin = PackURI::new(ORIGIN_NAME).expect("URI");
        duplicate
            .get_part_mut(&origin)
            .expect("origin")
            .rels_mut()
            .try_add_relationship(
                SIGNATURE_REL.into(),
                "/_xmlsignatures/sig1.xml".into(),
                "rIdDuplicate".into(),
                TargetMode::Internal,
            )
            .expect("duplicate target relationship");
        assert!(matches!(
            duplicate.signatures(),
            Err(Error::Graph(message)) if message.contains("duplicate signature target")
        ));
        duplicate.unsign();
        assert!(!duplicate.is_signed());
    }

    #[test]
    fn signed_package_survives_zip_round_trip() {
        let mut package = package();
        package.sign(&signer(7)).expect("sign");
        let mut output = std::io::Cursor::new(Vec::new());
        package.to_stream(&mut output).expect("serialize");
        let reopened = OpcPackage::from_bytes(output.get_ref()).expect("reopen");
        let reports = reopened.signatures().expect("verify");
        assert_eq!(reports.len(), 1);
        assert_eq!(reports[0].integrity(), Status::Valid);
        assert_eq!(reports[0].signature(), Status::Valid);
    }
}

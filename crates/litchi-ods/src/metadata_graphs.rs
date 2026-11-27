//! Source-checked transactions for inert package RDF metadata graphs.

use std::{fmt, sync::Arc};

use litchi_core::{Error, Position, Result};
use litchi_odf_common::XmlStreamLimits;
use litchi_odf_common::rdf::{Graph, Triple};

use crate::package::Archive;

/// An immutable RDF graph owner bound to exact ODS package bytes.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    graphs: Arc<[Graph]>,
    xml_limits: XmlStreamLimits,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("graphs", &self.graphs.len())
            .finish()
    }
}

impl Snapshot {
    /// Parse an owned ODS package and capture its inert RDF graph catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the package, manifest, or a declared graph is invalid.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_xml_limits(source, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_bytes_with_xml_limits(
        source: Vec<u8>,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        Self::from_arc_with_xml_limits(Arc::from(source), xml_limits)
    }

    fn from_arc_with_xml_limits(source: Arc<[u8]>, xml_limits: XmlStreamLimits) -> Result<Self> {
        let package = Archive::from_bytes_with_xml_limits(source.as_ref().to_vec(), xml_limits)?;
        let graphs = Arc::from(litchi_odf_common::rdf::graphs(package.package())?);
        Ok(Self {
            source,
            graphs,
            xml_limits,
        })
    }

    /// Borrow the exact package bytes captured by this snapshot.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Borrow the RDF graphs in deterministic package-path order.
    #[must_use]
    pub fn graphs(&self) -> &[Graph] {
        &self.graphs
    }

    /// Select one graph by its exact package-relative IRI.
    #[must_use]
    pub fn graph(&self, path: &str) -> Option<&Graph> {
        self.graphs.iter().find(|graph| graph.path == path)
    }

    /// Start a clone-staged, failure-atomic package edit.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            candidate: self.clone(),
        }
    }
}

/// A clone-staged RDF graph edit derived from one immutable [`Snapshot`].
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    candidate: Snapshot,
}

impl Edit {
    /// Borrow the current candidate graph catalog.
    #[must_use]
    pub fn graphs(&self) -> &[Graph] {
        self.candidate.graphs()
    }

    /// Add one validated graph and return its selected package-relative path.
    ///
    /// # Errors
    ///
    /// Returns an error when the path, triples, package, or compact output is invalid.
    pub fn add_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[Triple],
    ) -> Result<String> {
        let package = self.package()?;
        let (bytes, path) =
            litchi_odf_common::rdf::add_graph(package.package(), preferred_path, triples)?;
        validate_authored_xml(&bytes, Some(&path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(path)
    }

    /// Replace one complete graph while preserving its package path.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triples, package, or compact output is invalid.
    pub fn replace_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        if self
            .candidate
            .graph(path)
            .is_some_and(|graph| graph.triples == triples)
        {
            return Ok(());
        }
        let package = self.package()?;
        let bytes = litchi_odf_common::rdf::replace_graph(package.package(), path, triples)?;
        validate_authored_xml(&bytes, Some(path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(())
    }

    /// Remove one graph after validating that no retained graph references it.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is missing, referenced, or the package rebuild fails.
    pub fn remove_graph(&mut self, path: &str) -> Result<()> {
        let package = self.package()?;
        let bytes = litchi_odf_common::rdf::remove_graph(package.package(), path)?;
        validate_authored_xml(&bytes, None, self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(())
    }

    /// Append one triple and return its checked source position.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triple, package, or compact output is invalid.
    pub fn add_triple(&mut self, path: &str, triple: &Triple) -> Result<Position> {
        self.require_compact_graph(path)?;
        let package = self.package()?;
        let (bytes, index) = litchi_odf_common::rdf::add_triple(package.package(), path, triple)?;
        validate_authored_xml(&bytes, Some(path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(Position::new(index))
    }

    /// Replace one triple while retaining its description subject.
    ///
    /// # Errors
    ///
    /// Returns an error when the position, graph, triple, package, or compact output is invalid.
    pub fn replace_triple(
        &mut self,
        path: &str,
        position: Position,
        triple: &Triple,
    ) -> Result<()> {
        let graph = self.graph_or_error(path)?;
        if graph.triples.get(position.get()) == Some(triple) {
            return Ok(());
        }
        self.require_compact_graph(path)?;
        let package = self.package()?;
        let bytes = litchi_odf_common::rdf::replace_triple(
            package.package(),
            path,
            position.get(),
            triple,
        )?;
        validate_authored_xml(&bytes, Some(path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(())
    }

    /// Remove one triple at a checked source position.
    ///
    /// # Errors
    ///
    /// Returns an error when the position, graph, package, or compact output is invalid.
    pub fn remove_triple(&mut self, path: &str, position: Position) -> Result<()> {
        self.require_compact_graph(path)?;
        let package = self.package()?;
        let bytes = litchi_odf_common::rdf::remove_triple(package.package(), path, position.get())?;
        validate_authored_xml(&bytes, Some(path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(())
    }

    /// Move one triple to a final position within the same subject description.
    ///
    /// # Errors
    ///
    /// Returns an error when either position, graph, package, or compact output is invalid.
    pub fn move_triple(&mut self, path: &str, from: Position, to: Position) -> Result<()> {
        let graph = self.graph_or_error(path)?;
        check_position(from, graph.triples.len())?;
        check_position(to, graph.triples.len())?;
        if from == to {
            return Ok(());
        }
        self.require_compact_graph(path)?;
        let package = self.package()?;
        let bytes =
            litchi_odf_common::rdf::move_triple(package.package(), path, from.get(), to.get())?;
        validate_authored_xml(&bytes, Some(path), self.candidate.xml_limits)?;
        self.candidate = Snapshot::from_bytes_with_xml_limits(bytes, self.candidate.xml_limits)?;
        Ok(())
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        self.candidate = self.before.clone();
    }

    /// Publish the fully reparsed package candidate and reversible patch.
    #[must_use]
    pub(crate) fn commit(self) -> Commit {
        let patch = Patch {
            source: self.before.source,
            target: self.candidate.source.clone(),
        };
        Commit {
            snapshot: self.candidate,
            patch,
        }
    }

    fn package(&self) -> Result<Archive> {
        Archive::from_bytes_with_xml_limits(
            self.candidate.source.as_ref().to_vec(),
            self.candidate.xml_limits,
        )
    }

    fn graph_or_error(&self, path: &str) -> Result<&Graph> {
        self.candidate.graph(path).ok_or_else(|| {
            Error::InvalidFormat(format!("RDF metadata graph '{path}' was not found"))
        })
    }

    fn require_compact_graph(&self, path: &str) -> Result<()> {
        let package = self.package()?;
        let bytes = package.package().get_file(path)?;
        litchi_odf_common::compact_xml::validate(
            &bytes,
            litchi_odf_common::compact_xml::Limits::default(),
        )
        .map_err(Error::from)
    }
}

/// An exact-source, reversible RDF graph patch.
#[derive(Clone)]
pub(crate) struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_bytes", &self.source.len())
            .field("target_bytes", &self.target.len())
            .finish()
    }
}

impl Patch {
    /// Return whether this patch changes package bytes.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.source != self.target
    }

    /// Check exact applicability without mutating a snapshot.
    #[must_use]
    pub(crate) fn is_applicable_to(&self, snapshot: &Snapshot) -> bool {
        self.source.as_ref() == snapshot.as_bytes()
    }

    /// Apply this patch only to the exact package snapshot that produced it.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale source or invalid target package.
    pub(crate) fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if !self.is_applicable_to(snapshot) {
            return Err(Error::InvalidFormat(
                "ODS RDF patch source snapshot does not match".to_string(),
            ));
        }
        let target = Snapshot::from_bytes_with_xml_limits(
            self.target.as_ref().to_vec(),
            snapshot.xml_limits,
        )?;
        Ok(Commit {
            snapshot: target,
            patch: self.clone(),
        })
    }
}

/// A validated RDF graph publication.
#[derive(Debug)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Return whether package bytes changed.
    #[must_use]
    pub(crate) fn changed(&self) -> bool {
        self.patch.changed()
    }

    /// Borrow the resulting immutable snapshot.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume this publication into its immutable snapshot.
    #[must_use]
    pub(crate) fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

fn check_position(position: Position, length: usize) -> Result<()> {
    if position.get() < length {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "RDF triple position {} is outside graph length {length}",
            position.get()
        )))
    }
}

fn validate_authored_xml(
    bytes: &[u8],
    graph_path: Option<&str>,
    xml_limits: XmlStreamLimits,
) -> Result<()> {
    let package = Archive::from_bytes_with_xml_limits(bytes.to_vec(), xml_limits)?;
    let manifest = package.package().get_file("META-INF/manifest.xml")?;
    litchi_odf_common::compact_xml::validate(
        &manifest,
        litchi_odf_common::compact_xml::Limits::default(),
    )
    .map_err(Error::from)?;
    if let Some(path) = graph_path {
        let graph = package.package().get_file(path)?;
        litchi_odf_common::compact_xml::validate(
            &graph,
            litchi_odf_common::compact_xml::Limits::default(),
        )
        .map_err(Error::from)?;
    }
    Ok(())
}

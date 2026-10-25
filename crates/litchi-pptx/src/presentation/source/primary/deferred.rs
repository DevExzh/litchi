//! Committed source-backed packages whose output is a deferred stream
//! (record 0356a; the shape of DOCX record 0355a).
//!
//! A family whose change is not a slide overlay (the SVG picture families)
//! commits a new [`Package`] without writing anything. Its output is the
//! family's byte-preserving source publication, run when the package is
//! written ([`Package::write_to`], [`Package::save`]): untouched members stay
//! raw-copied from the retained source, so a commit followed by a save
//! streams exactly as the former `publish_*_to_stream` routes did. Any read
//! or further edit of the committed package materializes that output once
//! into memory, as a source-backed package over the published bytes, and
//! serves every later read from it.

use std::io::Write;
use std::sync::{Arc, OnceLock};

use litchi_core::OwnedSource;

use super::Package;
use crate::Result;

/// Writes a committed artifact to a sink.
pub(super) type Publish = dyn Fn(&mut dyn Write) -> Result<()> + Send + Sync;

/// The pending output of a committed package.
pub(super) struct Deferred {
    publish: Box<Publish>,
    options: crate::OpenOptions,
    materialized: OnceLock<Package>,
}

impl Deferred {
    /// The committed artifact as a source-backed package, published into
    /// memory on first use.
    pub(super) fn package(&self) -> Result<&Package> {
        if let Some(package) = self.materialized.get() {
            return Ok(package);
        }
        let mut bytes = Vec::new();
        (self.publish)(&mut bytes)?;
        let package = Package::from_read_at_with_options(
            Arc::new(OwnedSource::new(bytes)),
            self.options.clone(),
        )?;
        Ok(self.materialized.get_or_init(|| package))
    }

    /// Write the committed artifact.
    pub(super) fn write(&self, writer: &mut dyn Write) -> Result<()> {
        match self.materialized.get() {
            Some(package) => package.write_to(writer),
            None => (self.publish)(writer),
        }
    }
}

impl Package {
    /// A committed package over `self` whose artifact `publish` writes.
    /// `self` must be an ordinary (not deferred) package.
    pub(crate) fn deferred(
        &self,
        publish: impl Fn(&mut dyn Write) -> Result<()> + Send + Sync + 'static,
    ) -> Self {
        let mut options = crate::OpenOptions::default()
            .source_limits(self.owner.limits)
            .cache_limits(self.owner.package.cache_limits());
        options.semantic_limits = self.owner.semantic_limits.clone();
        let options = match self.owner.package.execution_context() {
            Some(context) => options.execution_context(context),
            None => options,
        };
        Self {
            deferred: Some(Arc::new(Deferred {
                publish: Box::new(publish),
                options,
                materialized: OnceLock::new(),
            })),
            ..self.clone()
        }
    }

    /// The package that serves reads: the materialized artifact of a
    /// committed deferred package, or `self`.
    pub(crate) fn effective(&self) -> Result<&Self> {
        match &self.deferred {
            Some(deferred) => deferred.package(),
            None => Ok(self),
        }
    }

    /// Whether `self` holds no staged change of its source: no overlay,
    /// root or order revision, topology or deferred output.
    pub(crate) fn is_pristine(&self) -> bool {
        self.deferred.is_none()
            && self.topology.is_none()
            && self.overlays.is_empty()
            && self.root.edited_bytes().is_none()
    }

    /// Whether two packages are shares of one deferred commit.
    pub(super) fn same_deferred(&self, other: &Self) -> bool {
        match (&self.deferred, &other.deferred) {
            (Some(left), Some(right)) => Arc::ptr_eq(left, right),
            (None, None) => true,
            _ => false,
        }
    }

    /// Write a deferred package's artifact; `None` for an ordinary package.
    pub(super) fn write_deferred(&self, writer: &mut dyn Write) -> Option<Result<()>> {
        self.deferred
            .as_ref()
            .map(|deferred| deferred.write(writer))
    }
}

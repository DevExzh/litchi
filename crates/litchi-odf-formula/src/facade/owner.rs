//! State verbs of the immutable Formula [`Package`] owner.

use std::path::Path;

use litchi_core::detection::FileFormat;
use litchi_odf_common::{Error, LimitKind, OpenOptions, SourceLimitKind, SourceLimits};

use super::{Edit, Formula, Package, Patch};
use crate::codec;
use crate::package::Archive;

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.as_bytes().len())
            .field("mimetype", &self.formula.mimetype())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// The retained physical package (metadata transactions, record 0288).
    pub(crate) fn owned_package(&self) -> &litchi_odf_common::OwnedPackage {
        self.formula.archive.owned()
    }

    /// Adopt bytes a metadata commit published, under this formula's limits.
    pub(crate) fn adopt_formula_bytes(&self, bytes: Vec<u8>) -> litchi_core::Result<Self> {
        let archive = Archive::from_bytes(bytes)?;
        Ok(Self {
            formula: Formula::from_archive(archive, self.formula.limits())?,
        })
    }

    /// Create the minimal valid `.odf` package: one empty block `mrow`.
    ///
    /// The result is deterministic: it reads no clock or randomness, so every
    /// call returns byte-identical packages.
    ///
    /// # Errors
    ///
    /// Returns an error only when the fixed minimal package cannot be built.
    pub fn new() -> Result<Self, Error> {
        crate::Builder::new(crate::authoring::document_root(
            crate::authoring::row(Vec::new()),
            crate::authoring::Display::Block,
        ))
        .build()
    }

    /// Open a Formula package from a path with the default [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when a source or semantic dimension is
    /// exceeded and [`Error::Core`] for I/O, archive, family or `MathML`
    /// validation failures.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Open a Formula package from a path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self, Error> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_options(file, options)
    }

    /// Open a Formula package from bytes (the bytes are copied once).
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Open a Formula package from bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self, Error> {
        let limits = options.source_limits_value();
        check_input_bytes(bytes.len(), limits)?;
        let owned =
            litchi_odf_common::OwnedPackage::from_bytes_with_limits(copy_input(bytes)?, limits)?;
        Self::from_owned_with_options(owned, options)
    }

    /// Open a Formula package from a reader with the default [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader(reader: impl std::io::Read) -> Result<Self, Error> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Open a Formula package from a reader under caller-selected limits. The
    /// reader is consumed through an input-limit guard before indexing.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: OpenOptions,
    ) -> Result<Self, Error> {
        let owned = litchi_odf_common::OwnedPackage::from_reader_with_limits(
            reader,
            options.source_limits_value(),
        )?;
        Self::from_owned_with_options(owned, options)
    }

    /// Adopt a packaged ODF result prepared by smart detection, reusing its
    /// validated archive index instead of indexing the bytes again.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_core::Error::UnexpectedFormat`] (as [`Error::Core`])
    /// when the prepared package is not a Formula package, and the
    /// [`Self::open`] failures otherwise.
    pub fn from_prepared_package(
        prepared: litchi_odf_common::PreparedPackage,
    ) -> Result<Self, Error> {
        match prepared.format() {
            FileFormat::Odf => {
                Self::from_owned_with_options(prepared.into_package(), OpenOptions::new())
            },
            detected => Err(Error::Core(litchi_core::Error::UnexpectedFormat {
                detected,
            })),
        }
    }

    fn from_owned_with_options(
        owned: litchi_odf_common::OwnedPackage,
        options: OpenOptions,
    ) -> Result<Self, Error> {
        let archive = Archive::from_owned_checked(owned, options.semantic_limits_value())?;
        Ok(Self::from_formula(Formula::from_archive(
            archive,
            codec::Limits::default(),
        )?))
    }

    /// Reopen package bytes carried by a durable patch envelope.
    pub(crate) fn from_wire(bytes: &[u8]) -> litchi_core::Result<Self> {
        Ok(Self::from_bytes(bytes)?)
    }

    /// Adopt a validated projection; the package shares its exact bytes.
    pub(crate) const fn from_formula(formula: Formula) -> Self {
        Self { formula }
    }

    /// The read-only `MathML` projection of this package.
    #[must_use]
    pub const fn formula(&self) -> &Formula {
        &self.formula
    }

    /// Start a source-bound whole-root edit of this package.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            replacement: None,
            changes: Vec::new(),
        }
    }

    /// Apply a patch to this exact package.
    ///
    /// # Errors
    ///
    /// Returns an error when this package is not byte-identical to the patch
    /// source.
    pub fn apply(&self, patch: &Patch) -> Result<Self, Error> {
        Ok(patch.apply(self)?)
    }

    /// Atomically replace `path` with this package (full durability).
    ///
    /// The bytes are staged in a sibling temporary file and published with
    /// one rename, so a failure before publication leaves an existing
    /// destination unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] carrying the publication failure.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<(), Error> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replace `path` with this package under explicit durability.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<(), Error> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }

    /// Copy the exact package bytes.
    ///
    /// # Errors
    ///
    /// Returns an allocation failure.
    pub fn to_bytes(&self) -> Result<Vec<u8>, Error> {
        Ok(copy_input(self.as_bytes())?)
    }

    /// Write the exact package bytes to a stream.
    ///
    /// # Errors
    ///
    /// Returns the stream's I/O failure.
    pub fn write_to(&self, mut writer: impl std::io::Write) -> Result<(), Error> {
        writer.write_all(self.as_bytes())?;
        Ok(())
    }

    /// The exact package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.formula.as_bytes()
    }

    /// The identity of the validated archive index retained by this package.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.formula.archive.prepared_index_identity()
    }
}

fn check_input_bytes(length: usize, limits: SourceLimits) -> Result<(), Error> {
    let observed = u64::try_from(length).unwrap_or(u64::MAX);
    if observed > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::Source(SourceLimitKind::InputBytes),
            observed,
            maximum: limits.max_input_bytes(),
        });
    }
    Ok(())
}

fn copy_input(source: &[u8]) -> litchi_core::Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| litchi_core::Error::Allocation {
            resource: "Formula package bytes",
            source: allocation_error,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

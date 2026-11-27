//! Host-supplied storage targets.
//!
//! A target is an opaque semantic key paired with an exact CFB storage path.
//! The common layer never infers that key from a directory name: concrete
//! format crates derive targets from their own records and pass them here.

use super::cfb_path::CfbPath;
use crate::Error;
use litchi_cfb::OleError;
use litchi_cfb::OleFile;
use std::io::{Read, Seek};

/// One host-resolved CFB storage target.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Target {
    key: String,
    path: CfbPath,
}

impl Target {
    /// Creates a target from a host-owned key and exact CFB storage path.
    ///
    /// # Errors
    ///
    /// Returns an error when the key is empty or a path component is not a
    /// valid `[MS-CFB]` directory name.
    pub fn new<S, I>(key_source: impl Into<String>, path_source: I) -> Result<Self, Error>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let key = key_source.into();
        let path = CfbPath::new(path_source.into_iter().map(Into::into).collect())
            .map_err(|_| Error::invalid_selector())?;
        if key.is_empty() {
            return Err(Error::invalid_selector());
        }
        Ok(Self { key, path })
    }

    /// Creates a target whose key is the final storage name.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is empty or contains an invalid part.
    pub fn from_path<I, S>(path_source: I) -> Result<Self, Error>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let path = path_source.into_iter().map(Into::into).collect::<Vec<_>>();
        let key = path.last().cloned().ok_or_else(Error::invalid_selector)?;
        Self::new(key, path)
    }

    /// The host-owned semantic key used by [`super::Objects::get`].
    #[must_use]
    pub fn key(&self) -> &str {
        &self.key
    }

    /// The CFB storage path requested by this target.
    ///
    /// An editor resolves case variants to the stored directory spelling
    /// before capturing an object, so its target catalog exposes the exact
    /// physical path used for subsequent edits.
    #[must_use]
    pub fn path(&self) -> &[String] {
        self.path.as_slice()
    }

    pub(crate) fn resolve<R: Read + Seek>(
        &self,
        ole: &OleFile<R>,
        max_entries: usize,
    ) -> Result<Self, Error> {
        let path = self
            .path
            .resolve(ole, max_entries)
            .map_err(classify_resolution)?;
        Ok(Self {
            key: self.key.clone(),
            path: CfbPath::new(path).map_err(Error::from)?,
        })
    }
}

static EMPTY_TARGETS: Targets = Targets {
    targets: Vec::new(),
};

impl super::snapshot::EmptyValue for Targets {
    fn empty() -> &'static Self {
        &EMPTY_TARGETS
    }

    fn is_empty_value(&self) -> bool {
        self.targets.is_empty()
    }
}

/// A validated collection of host-resolved storage targets.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Targets {
    targets: Vec<Target>,
}

impl Targets {
    /// Moves this collection into shared owner state; the empty (default)
    /// collection allocates nothing.
    pub(crate) fn into_shared(self) -> super::snapshot::Shared<Self> {
        super::snapshot::Shared::new(self)
    }

    /// Creates a target collection and rejects duplicate or overlapping paths.
    ///
    /// # Errors
    ///
    /// Returns an error when a target key is repeated or target paths overlap.
    pub fn new<I>(targets: I) -> Result<Self, Error>
    where
        I: IntoIterator<Item = Target>,
    {
        let mut output = Self::default();
        for target in targets {
            output.push(target)?;
        }
        Ok(output)
    }

    /// Creates a collection containing one target.
    #[must_use]
    pub fn one(target: Target) -> Self {
        Self {
            targets: vec![target],
        }
    }

    /// Returns all targets in deterministic caller-provided order.
    #[must_use]
    pub fn as_slice(&self) -> &[Target] {
        &self.targets
    }

    /// Borrows the targets in caller-provided order.
    pub fn iter(&self) -> std::slice::Iter<'_, Target> {
        self.targets.iter()
    }

    /// Returns the number of selected storages.
    #[must_use]
    pub fn len(&self) -> usize {
        self.targets.len()
    }

    /// Returns whether no storages are selected.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.targets.is_empty()
    }

    #[must_use]
    pub(crate) fn into_vec(self) -> Vec<Target> {
        self.targets
    }

    /// Finds a target by its host-owned semantic key.
    #[must_use]
    pub fn get(&self, key: &str) -> Option<&Target> {
        self.targets.iter().find(|target| target.key == key)
    }

    /// Finds a target by discovery order.
    #[must_use]
    pub fn at(&self, index: usize) -> Option<&Target> {
        self.targets.get(index)
    }

    /// Adds one target after checking key, path, and prefix uniqueness.
    ///
    /// # Errors
    ///
    /// Returns an error when the target key is repeated or its path overlaps
    /// an existing target path.
    pub fn push(&mut self, target: Target) -> Result<(), Error> {
        if self.targets.iter().any(|value| value.key == target.key) {
            return Err(Error::invalid_selector());
        }
        if self
            .targets
            .iter()
            .any(|value| value.path.overlaps(&target.path))
        {
            return Err(Error::invalid_selector());
        }
        self.targets.push(target);
        Ok(())
    }

    pub(crate) fn without(&self, key: &str) -> Result<Self, Error> {
        let index = self
            .targets
            .iter()
            .position(|target| target.key == key)
            .ok_or_else(Error::target_not_found)?;
        let mut output = self.clone();
        output.targets.remove(index);
        Ok(output)
    }
}

fn classify_resolution(error: OleError) -> Error {
    match error {
        OleError::InvalidFormat(message)
            if message.starts_with("object storage path component ")
                && message.ends_with(" not found") =>
        {
            Error::target_not_found()
        },
        error => Error::from(error),
    }
}

impl<'a> IntoIterator for &'a Targets {
    type Item = &'a Target;
    type IntoIter = std::slice::Iter<'a, Target>;

    fn into_iter(self) -> Self::IntoIter {
        self.targets.iter()
    }
}

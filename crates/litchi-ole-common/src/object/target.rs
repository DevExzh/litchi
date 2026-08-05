//! Host-supplied storage targets.
//!
//! A target is an opaque semantic key paired with an exact CFB storage path.
//! The common layer never infers that key from a directory name: concrete
//! format crates derive targets from their own records and pass them here.

use litchi_cfb::OleError;

/// One host-resolved CFB storage target.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Target {
    key: String,
    path: Vec<String>,
}

impl Target {
    /// Creates a target from a host-owned key and exact CFB storage path.
    ///
    /// # Errors
    ///
    /// Returns an error when the key or path contains an empty/repeated part.
    pub fn new<S, I>(key_source: impl Into<String>, path_source: I) -> Result<Self, OleError>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let key = key_source.into();
        let path = path_source.into_iter().map(Into::into).collect::<Vec<_>>();
        validate_parts(&key, &path)?;
        Ok(Self { key, path })
    }

    /// Creates a target whose key is the final storage name.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is empty or contains an invalid part.
    pub fn from_path<I, S>(path_source: I) -> Result<Self, OleError>
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        let path = path_source.into_iter().map(Into::into).collect::<Vec<_>>();
        let key = path
            .last()
            .cloned()
            .ok_or_else(|| OleError::InvalidFormat("object target path is empty".into()))?;
        Self::new(key, path)
    }

    /// The host-owned semantic key used by [`super::Objects::get`].
    #[must_use]
    pub fn key(&self) -> &str {
        &self.key
    }

    /// The exact CFB storage path selected by this target.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }
}

/// A validated collection of host-resolved storage targets.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Targets {
    targets: Vec<Target>,
}

impl Targets {
    /// Creates a target collection and rejects duplicate or overlapping paths.
    ///
    /// # Errors
    ///
    /// Returns an error when a target key is repeated or target paths overlap.
    pub fn new<I>(targets: I) -> Result<Self, OleError>
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
    #[must_use]
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
    pub fn push(&mut self, target: Target) -> Result<(), OleError> {
        if self.targets.iter().any(|value| value.key == target.key) {
            return Err(OleError::InvalidFormat(format!(
                "duplicate object target key {:?}",
                target.key
            )));
        }
        if self.targets.iter().any(|value| {
            value.path == target.path
                || value.path.starts_with(&target.path)
                || target.path.starts_with(&value.path)
        }) {
            return Err(OleError::InvalidFormat(format!(
                "object target path overlaps {:?}",
                target.path
            )));
        }
        self.targets.push(target);
        Ok(())
    }

    pub(crate) fn without(&self, key: &str) -> Result<Self, OleError> {
        let index = self
            .targets
            .iter()
            .position(|target| target.key == key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?;
        let mut output = self.clone();
        output.targets.remove(index);
        Ok(output)
    }
}

impl<'a> IntoIterator for &'a Targets {
    type Item = &'a Target;
    type IntoIter = std::slice::Iter<'a, Target>;

    fn into_iter(self) -> Self::IntoIter {
        self.targets.iter()
    }
}

fn validate_parts(key: &str, path: &[String]) -> Result<(), OleError> {
    if key.is_empty() {
        return Err(OleError::InvalidFormat(
            "object target key must not be empty".into(),
        ));
    }
    if path.is_empty() || path.iter().any(String::is_empty) {
        return Err(OleError::InvalidFormat(
            "object target path must contain non-empty storage names".into(),
        ));
    }
    Ok(())
}

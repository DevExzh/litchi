//! Contextual protection model.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Type {
    None,
    ReadOnlyRecommended,
    ModifyPassword,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Algorithm {
    Sha1,
    Sha256,
    Sha384,
    #[default]
    Sha512,
}

impl Algorithm {
    #[must_use]
    pub const fn uri(self) -> &'static str {
        match self {
            Self::Sha1 => "http://www.w3.org/2000/09/xmldsig#sha1",
            Self::Sha256 => "http://www.w3.org/2001/04/xmlenc#sha256",
            Self::Sha384 => "http://www.w3.org/2001/04/xmldsig-more#sha384",
            Self::Sha512 => "http://www.w3.org/2001/04/xmlenc#sha512",
        }
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_uri(value: &str) -> crate::Result<Self> {
        match value {
            "http://www.w3.org/2000/09/xmldsig#sha1" | "SHA-1" => Ok(Self::Sha1),
            "http://www.w3.org/2001/04/xmlenc#sha256" | "SHA-256" => Ok(Self::Sha256),
            "http://www.w3.org/2001/04/xmldsig-more#sha384" | "SHA-384" => Ok(Self::Sha384),
            "http://www.w3.org/2001/04/xmlenc#sha512" | "SHA-512" => Ok(Self::Sha512),
            _ => Err(crate::Error::Invalid(format!(
                "unsupported presentation protection hash algorithm '{value}'"
            ))),
        }
    }

    pub(crate) fn from_sid(sid: u32) -> crate::Result<Self> {
        match sid {
            4 => Ok(Self::Sha1),
            12 => Ok(Self::Sha256),
            13 => Ok(Self::Sha384),
            14 => Ok(Self::Sha512),
            _ => Err(crate::Error::Invalid(format!(
                "unsupported presentation protection hash SID {sid}"
            ))),
        }
    }

    pub(crate) const fn output_bytes(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }

    pub(crate) const fn sid(self) -> u32 {
        match self {
            Self::Sha1 => 4,
            Self::Sha256 => 12,
            Self::Sha384 => 13,
            Self::Sha512 => 14,
        }
    }
}

#[derive(Clone)]
pub struct Verifier {
    pub(crate) algorithm: Algorithm,
    pub(crate) spin_count: u32,
    pub(crate) hash: String,
    pub(crate) salt: String,
}

impl Verifier {
    #[must_use]
    pub const fn algorithm(&self) -> Algorithm {
        self.algorithm
    }

    #[must_use]
    pub const fn spins(&self) -> u32 {
        self.spin_count
    }

    #[must_use]
    pub fn hash(&self) -> &str {
        &self.hash
    }

    #[must_use]
    pub fn salt(&self) -> &str {
        &self.salt
    }
}

impl std::fmt::Debug for Verifier {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Verifier")
            .field("algorithm", &self.algorithm)
            .field("spin_count", &self.spin_count)
            .field("hash_bytes", &self.hash.len())
            .field("salt_bytes", &self.salt.len())
            .finish_non_exhaustive()
    }
}

#[derive(Debug, Clone, Default)]
pub struct Settings {
    pub read_only_recommended: bool,
    pub(crate) modify: Option<Verifier>,
    pub protect_structure: bool,
    pub protect_windows: bool,
}

impl Settings {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn with_read_only_recommended(mut self, value: bool) -> Self {
        self.read_only_recommended = value;
        self
    }

    #[must_use]
    pub fn with_structure_protection(mut self, value: bool) -> Self {
        self.protect_structure = value;
        self
    }

    #[must_use]
    pub fn with_window_protection(mut self, value: bool) -> Self {
        self.protect_windows = value;
        self
    }

    #[must_use]
    pub fn is_protected(&self) -> bool {
        self.read_only_recommended
            || self.modify.is_some()
            || self.protect_structure
            || self.protect_windows
    }

    #[must_use]
    pub fn protection_type(&self) -> Type {
        if self.modify.is_some() {
            Type::ModifyPassword
        } else if self.read_only_recommended {
            Type::ReadOnlyRecommended
        } else {
            Type::None
        }
    }

    #[must_use]
    pub fn modify(&self) -> Option<&Verifier> {
        self.modify.as_ref()
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_modify_password(&mut self, password: &str) -> crate::Result<()> {
        self.modify = Some(super::codec::generate_verifier(password)?);
        Ok(())
    }

    pub fn clear_modify_password(&mut self) {
        self.modify = None;
    }
}

#[derive(Debug, Clone, Default)]
pub struct Slide {
    pub no_select: bool,
    pub no_move: bool,
    pub no_resize: bool,
    pub no_edit_text: bool,
    pub no_ungroup: bool,
    pub no_change_z_order: bool,
}

impl Slide {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn protect_all(mut self) -> Self {
        self.no_select = true;
        self.no_move = true;
        self.no_resize = true;
        self.no_edit_text = true;
        self.no_ungroup = true;
        self.no_change_z_order = true;
        self
    }

    #[must_use]
    pub fn is_protected(&self) -> bool {
        self.no_select
            || self.no_move
            || self.no_resize
            || self.no_edit_text
            || self.no_ungroup
            || self.no_change_z_order
    }
}

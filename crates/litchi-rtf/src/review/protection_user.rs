//! Inert usernames from the RTF `protusertbl` destination.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

/// Maximum number of protection-user entries retained from one document.
pub const MAX_PROTECTION_USERS: usize = 65_536;
/// Maximum decoded UTF-8 byte length of one protection username.
pub const MAX_PROTECTION_USER_BYTES: usize = 65_536;
/// Maximum aggregate decoded UTF-8 byte length of a protection-user table.
pub const MAX_PROTECTION_USER_TOTAL_BYTES: usize = 16 * 1_048_576;

/// One inert username used by range-level document protection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionUser<'a> {
    /// User or domain-qualified user name exactly as retained by the parser.
    pub name: Cow<'a, str>,
}

impl<'a> ProtectionUser<'a> {
    /// Create and validate one inert protection-user entry.
    pub fn new(name: Cow<'a, str>) -> RtfResult<Self> {
        let user = Self { name };
        user.validate()?;
        Ok(user)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.name.trim().is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF protection username cannot be empty".to_string(),
            ));
        }
        if self.name.len() > MAX_PROTECTION_USER_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF protection username exceeds the safety limit".to_string(),
            ));
        }
        if self.name.contains(['\0', '\r', '\n']) {
            return Err(RtfError::MalformedDocument(
                "RTF protection username contains a forbidden control character".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> ProtectionUser<'static> {
        ProtectionUser {
            name: Cow::Owned(self.name.into_owned()),
        }
    }
}

/// Ordered inert usernames from a present RTF protection-user table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionUserTable<'a> {
    users: Vec<ProtectionUser<'a>>,
}

impl<'a> ProtectionUserTable<'a> {
    /// Create a protection-user table. RTF 1.9.1 requires at least one entry.
    pub fn new(users: Vec<ProtectionUser<'a>>) -> RtfResult<Self> {
        let table = Self { users };
        table.validate()?;
        Ok(table)
    }

    /// Return protection users in source order.
    pub fn users(&self) -> &[ProtectionUser<'a>] {
        &self.users
    }

    /// Append a validated username while enforcing table-wide limits.
    pub fn push(&mut self, user: ProtectionUser<'a>) -> RtfResult<()> {
        user.validate()?;
        if self.users.len() >= MAX_PROTECTION_USERS {
            return Err(RtfError::MalformedDocument(
                "RTF protection-user count exceeds the safety limit".to_string(),
            ));
        }
        let total = self
            .users
            .iter()
            .try_fold(user.name.len(), |total, existing| {
                total.checked_add(existing.name.len()).ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF protection-user aggregate size overflow".to_string(),
                    )
                })
            })?;
        if total > MAX_PROTECTION_USER_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF protection-user aggregate text exceeds the safety limit".to_string(),
            ));
        }
        self.users.push(user);
        Ok(())
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.users.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF protection-user table cannot be empty".to_string(),
            ));
        }
        if self.users.len() > MAX_PROTECTION_USERS {
            return Err(RtfError::MalformedDocument(
                "RTF protection-user count exceeds the safety limit".to_string(),
            ));
        }
        let mut total = 0usize;
        for user in &self.users {
            user.validate()?;
            total = total.checked_add(user.name.len()).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF protection-user aggregate size overflow".to_string(),
                )
            })?;
            if total > MAX_PROTECTION_USER_TOTAL_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF protection-user aggregate text exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> ProtectionUserTable<'static> {
        ProtectionUserTable {
            users: self
                .users
                .into_iter()
                .map(ProtectionUser::into_owned)
                .collect(),
        }
    }
}

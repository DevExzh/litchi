//! Explicit native document identity for fresh Keynote presentations.
//!
//! Apple stores three independent UUIDs in every iWork package: the public
//! document UUID, the UUID of the current saved version, and a private UUID
//! used for local bookkeeping. Two packages that share a document UUID are
//! treated by Keynote and iCloud as revisions of one document, so fresh
//! authoring never invents an identity on its own: [`crate::Builder`]
//! requires an [`Identity`], and [`Identity::generate`] is the only place in
//! this crate that consults randomness (ADR 0034 §9 "Determinism").
//!
//! ```
//! use litchi_keynote::Identity;
//!
//! let fixed = Identity::from_bytes([1; 16], [2; 16], [3; 16])?;
//! assert_eq!(fixed.document(), [1; 16]);
//! assert_eq!(
//!     fixed.to_string(),
//!     "01010101-0101-0101-0101-010101010101 02020202-0202-0202-0202-020202020202 \
//!      03030303-0303-0303-0303-030303030303"
//! );
//! assert!(Identity::from_bytes([1; 16], [1; 16], [3; 16]).is_err());
//! # Ok::<(), litchi_keynote::identity::Error>(())
//! ```

use std::fmt;

use thiserror::Error as ThisError;

/// One of the three UUID roles carried by an [`Identity`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Role {
    /// The stable public document UUID (`Metadata/DocumentIdentifier`).
    DocumentUuid,
    /// The UUID of the current saved version and package revision.
    VersionUuid,
    /// The private UUID used by Keynote's local document bookkeeping.
    PrivateUuid,
}

impl fmt::Display for Role {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::DocumentUuid => "document",
            Self::VersionUuid => "version",
            Self::PrivateUuid => "private",
        })
    }
}

/// An invalid explicit [`Identity`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A UUID is all zero bytes.
    #[error("the Keynote {role} UUID must not be nil")]
    Nil {
        /// The role whose UUID is nil.
        role: Role,
    },
    /// Two roles carry the same UUID.
    #[error("the Keynote {first} and {second} UUIDs must differ")]
    Repeated {
        /// The earlier role.
        first: Role,
        /// The later role carrying the same UUID.
        second: Role,
    },
}

/// The three native UUIDs that identify one Keynote document.
///
/// Values are validated at construction: no UUID is nil and the three are
/// mutually distinct. The textual form used inside the package is the
/// uppercase hyphenated RFC 4122 spelling that Keynote writes.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Identity {
    document: [u8; 16],
    version: [u8; 16],
    private: [u8; 16],
}

impl fmt::Debug for Identity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Identity")
            .field("document", &Hyphenated(&self.document))
            .field("version", &Hyphenated(&self.version))
            .field("private", &Hyphenated(&self.private))
            .finish()
    }
}

/// Formats as the three hyphenated UUIDs (document, version, private)
/// separated by single spaces.
impl fmt::Display for Identity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} {} {}",
            Hyphenated(&self.document),
            Hyphenated(&self.version),
            Hyphenated(&self.private)
        )
    }
}

impl Identity {
    /// The fixed identity carried by [`crate::Package::new`].
    ///
    /// Document `4C495443-4849-4B4E-8000-000000000001`, version
    /// `4C495443-4849-4B4E-8000-000000000002`, private
    /// `4C495443-4849-4B4E-8000-000000000003` (`LITCHIKN` in ASCII followed
    /// by the role number). Every package built from it shares one document
    /// identity; callers that need an independent document use
    /// [`crate::Builder`] with [`Identity::generate`].
    pub const BLANK: Self = Self {
        document: blank_uuid(1),
        version: blank_uuid(2),
        private: blank_uuid(3),
    };

    /// Generate three fresh, mutually distinct RFC 4122 version 4 UUIDs.
    ///
    /// This is the only operation in the crate that consults randomness.
    #[must_use]
    pub fn generate() -> Self {
        let document = fresh_uuid(&[]);
        let version = fresh_uuid(&[document]);
        let private = fresh_uuid(&[document, version]);
        Self {
            document,
            version,
            private,
        }
    }

    /// Build an identity from fixed document, version and private UUID bytes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Nil`] when a UUID is all zero bytes and
    /// [`Error::Repeated`] when two roles carry the same UUID.
    pub fn from_bytes(
        document: [u8; 16],
        version: [u8; 16],
        private: [u8; 16],
    ) -> Result<Self, Error> {
        for (role, value) in [
            (Role::DocumentUuid, document),
            (Role::VersionUuid, version),
            (Role::PrivateUuid, private),
        ] {
            if value == [0; 16] {
                return Err(Error::Nil { role });
            }
        }
        if document == version {
            return Err(Error::Repeated {
                first: Role::DocumentUuid,
                second: Role::VersionUuid,
            });
        }
        if document == private {
            return Err(Error::Repeated {
                first: Role::DocumentUuid,
                second: Role::PrivateUuid,
            });
        }
        if version == private {
            return Err(Error::Repeated {
                first: Role::VersionUuid,
                second: Role::PrivateUuid,
            });
        }
        Ok(Self {
            document,
            version,
            private,
        })
    }

    /// Return the stable public document UUID bytes.
    #[must_use]
    pub const fn document(&self) -> [u8; 16] {
        self.document
    }

    /// Return the current saved-version UUID bytes.
    #[must_use]
    pub const fn version(&self) -> [u8; 16] {
        self.version
    }

    /// Return the private bookkeeping UUID bytes.
    #[must_use]
    pub const fn private(&self) -> [u8; 16] {
        self.private
    }

    /// Return the UUID of one role in its native textual spelling.
    pub(crate) fn text(&self, role: Role) -> String {
        Hyphenated(match role {
            Role::DocumentUuid => &self.document,
            Role::VersionUuid => &self.version,
            Role::PrivateUuid => &self.private,
        })
        .to_string()
    }

    /// Parse a hyphenated UUID in either letter case.
    pub(crate) fn parse_uuid(text: &str) -> Option<[u8; 16]> {
        let bytes = text.as_bytes();
        if bytes.len() != 36 {
            return None;
        }
        let mut output = [0_u8; 16];
        let mut nibbles = 0_usize;
        for (index, byte) in bytes.iter().copied().enumerate() {
            if matches!(index, 8 | 13 | 18 | 23) {
                if byte != b'-' {
                    return None;
                }
                continue;
            }
            let value = char::from(byte).to_digit(16)?;
            let slot = output.get_mut(nibbles / 2)?;
            *slot = (*slot << 4) | u8::try_from(value).ok()?;
            nibbles += 1;
        }
        (nibbles == 32).then_some(output)
    }
}

const fn blank_uuid(role: u8) -> [u8; 16] {
    [
        0x4C, 0x49, 0x54, 0x43, 0x48, 0x49, 0x4B, 0x4E, 0x80, 0, 0, 0, 0, 0, 0, role,
    ]
}

fn fresh_uuid(existing: &[[u8; 16]]) -> [u8; 16] {
    loop {
        let candidate = litchi_core::id::generate_guid_bytes();
        if candidate != [0; 16] && !existing.contains(&candidate) {
            return candidate;
        }
    }
}

struct Hyphenated<'a>(&'a [u8; 16]);

impl fmt::Debug for Hyphenated<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, formatter)
    }
}

impl fmt::Display for Hyphenated<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        for (index, byte) in self.0.iter().enumerate() {
            if matches!(index, 4 | 6 | 8 | 10) {
                formatter.write_str("-")?;
            }
            write!(formatter, "{byte:02X}")?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{Identity, Role};

    #[test]
    fn blank_identity_is_valid_and_spelled_as_documented() {
        let blank = Identity::BLANK;
        assert_eq!(
            Identity::from_bytes(blank.document(), blank.version(), blank.private()),
            Ok(blank)
        );
        assert_eq!(
            blank.text(Role::DocumentUuid),
            "4C495443-4849-4B4E-8000-000000000001"
        );
        assert_eq!(
            blank.text(Role::VersionUuid),
            "4C495443-4849-4B4E-8000-000000000002"
        );
        assert_eq!(
            blank.text(Role::PrivateUuid),
            "4C495443-4849-4B4E-8000-000000000003"
        );
    }

    #[test]
    fn uuid_text_round_trips_in_both_cases() {
        let identity = Identity::generate();
        for role in [Role::DocumentUuid, Role::VersionUuid, Role::PrivateUuid] {
            let text = identity.text(role);
            let parsed = Identity::parse_uuid(&text);
            assert!(parsed.is_some());
            assert_eq!(Identity::parse_uuid(&text.to_ascii_lowercase()), parsed);
        }
        assert_eq!(Identity::parse_uuid("not-a-uuid"), None);
        assert_eq!(
            Identity::parse_uuid("4C495443x4849-4B4E-8000-000000000001"),
            None
        );
    }
}

//! Publication policy for legacy Word document protection.

use crate::package::{Error as PackageError, Result};
use crate::parts::document_properties::DocumentProperties;
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::Patch as ObjectPatch;

use super::Ranges;

const DOP_POINTER: usize = 31;
const FIB_CSW: usize = 32;
const FIB_CSLW: usize = 62;
const WORD97_NFIB: u16 = 0x00C1;
const WORD2000_NFIB: u16 = 0x00D9;
const WORD2002_NFIB: u16 = 0x0101;
const WORD2003_NFIB: u16 = 0x010C;
const WORD2007_NFIB: u16 = 0x0112;

/// Protection observed on a DOC source.
///
/// The state describes the two independent MS-DOC mechanisms that can restrict
/// edits.  A non-empty range table is protected even when the DOP itself does
/// not carry a document-wide lock.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EditProtection {
    /// No document or range-level editing restriction was observed.
    None,
    /// A document-wide DOP restriction is active.
    Document,
    /// At least one range-level protection bookmark is present.
    Ranges,
    /// Both document-wide and range-level restrictions are active.
    DocumentAndRanges,
    /// The host's protection state could not be established safely.
    ///
    /// This state is intentionally fail-closed. Package adapters may expose
    /// an inert source for inspection while refusing changed publication
    /// because a malformed or incomplete host cannot be proven unprotected.
    Unknown,
    /// A structurally valid host carries a DOP or range record outside the
    /// typed grammar, so its protection state is not known precisely.
    ///
    /// This remains distinct from [`Self::None`]. It is retained for
    /// inspection and exact no-op publication only; an authorization cannot
    /// turn an unrecognized protection record into a known safe state.
    Unrecognized,
}

impl EditProtection {
    /// Whether this state requires an explicit edit authorization.
    #[must_use]
    pub const fn is_protected(self) -> bool {
        !matches!(self, Self::None)
    }
}

/// The caller identity and reason retained for an explicit protected-edit
/// authorization.
///
/// This is caller metadata, not a password verifier or cryptographic audit.
/// Word's legacy protection hash is not a security boundary, and this crate
/// does not claim to authenticate a Word account. Callers that intentionally
/// edit a protected source must supply an actor and reason at the API boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionAuthorization {
    actor: String,
    reason: String,
}

impl ProtectionAuthorization {
    /// Creates a caller-supplied authorization record.
    pub fn audited(
        actor: impl Into<String>,
        reason: impl Into<String>,
    ) -> std::result::Result<Self, AuthorizationError> {
        let actor = actor.into();
        if actor.trim().is_empty() {
            return Err(AuthorizationError::EmptyActor);
        }
        let reason = reason.into();
        if reason.trim().is_empty() {
            return Err(AuthorizationError::EmptyReason);
        }
        Ok(Self { actor, reason })
    }

    /// The caller identity supplied with the authorization.
    #[must_use]
    pub fn actor(&self) -> &str {
        &self.actor
    }

    /// The caller-supplied reason for the edit.
    #[must_use]
    pub fn reason(&self) -> &str {
        &self.reason
    }
}

/// Errors produced while constructing an explicit protection authorization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AuthorizationError {
    /// An authorization must identify its caller.
    EmptyActor,
    /// An authorization must state why the protected edit is allowed.
    EmptyReason,
}

impl std::fmt::Display for AuthorizationError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::EmptyActor => "protected-edit authorization requires a non-empty actor",
            Self::EmptyReason => "protected-edit authorization requires a non-empty reason",
        })
    }
}

impl std::error::Error for AuthorizationError {}

/// Publication policy for a DOC editor.
///
/// [`Self::default`] enforces document protection. `AllowProtected` is an
/// explicit, caller-created capability carrying the identity and reason
/// required by ADR-0006. The record does not perform password verification,
/// authentication, or cryptographic auditing.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum ProtectionPolicy {
    /// Reject changed publication while document or range protection is active.
    #[default]
    Enforce,
    /// Permit changed publication with an explicit caller-granted capability.
    AllowProtected(ProtectionAuthorization),
}

impl ProtectionPolicy {
    /// Creates a policy that explicitly authorizes protected publication.
    #[must_use]
    pub fn allow_protected(authorization: ProtectionAuthorization) -> Self {
        Self::AllowProtected(authorization)
    }

    /// Whether this policy contains an explicit protected-edit capability.
    #[must_use]
    pub const fn allows_protected_edits(&self) -> bool {
        matches!(self, Self::AllowProtected(_))
    }

    /// Enforces the policy for one decoded DOC source.
    pub(crate) fn authorize(&self, state: EditProtection) -> Result<()> {
        if matches!(
            state,
            EditProtection::Unknown | EditProtection::Unrecognized
        ) || state.is_protected() && !self.allows_protected_edits()
        {
            return Err(PackageError::ProtectionDenied(state));
        }
        Ok(())
    }
}

/// A whole-CFB patch published by a DOC semantic owner.
///
/// The common OLE patch remains source-bound, but it does not know DOC
/// protection. This wrapper carries the source and replacement protection
/// states so applying a package patch observes the destination policy just
/// like a semantic patch. Callers that intentionally publish protected bytes
/// must provide the same explicit capability used to create the commit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PackagePatch {
    inner: ObjectPatch,
    before_protection: EditProtection,
    after_protection: EditProtection,
}

impl PackagePatch {
    pub(crate) fn new(inner: ObjectPatch, protection: EditProtection) -> Self {
        Self {
            inner,
            before_protection: protection,
            after_protection: protection,
        }
    }

    /// Bytes required as the source of this package patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.inner.before()
    }

    /// Bytes produced by this package patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.inner.after()
    }

    /// Whether the complete CFB artifact is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }

    /// Returns the inverse patch with its destination protection state.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
            before_protection: self.after_protection,
            after_protection: self.before_protection,
        }
    }

    /// Applies the patch under the default DOC publication policy.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        self.apply_with_policy(source, ProtectionPolicy::default())
    }

    /// Applies the patch with an explicit caller-granted protected-edit
    /// capability. The source bytes must still match exactly.
    pub fn apply_with_policy(&self, source: &[u8], policy: ProtectionPolicy) -> Result<Vec<u8>> {
        if source != self.inner.before() {
            return Err(PackageError::InvalidFormat(
                "DOC package patch source snapshot does not match".into(),
            ));
        }
        if !self.inner.is_noop() {
            policy.authorize(self.before_protection)?;
        }
        self.inner.apply(source).map_err(PackageError::from)
    }
}

/// Classifies document and range-level protection from a selected Word FIB
/// and table stream.
pub(crate) fn classify(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<EditProtection> {
    // A missing or truncated counted FIB must never be interpreted as an
    // unprotected document.  In particular, get_table_pointer(DOP) returns
    // None when cbRgFcLcb stops before that pair, which otherwise creates a
    // fail-open path for command-bar and other package owners.
    let effective_nfib = match validate_fib_shape(fib) {
        Ok(value) => value,
        Err(_) => return Ok(EditProtection::Unknown),
    };
    let Some((_, dop_length)) = fib.get_table_pointer(DOP_POINTER) else {
        return Ok(EditProtection::Unknown);
    };
    if dop_length == 0 {
        return Ok(EditProtection::Unknown);
    }
    let document = match document_protected(fib, table_stream, effective_nfib) {
        Ok(document) => document,
        Err(PackageError::Corrupted(_)) => return Ok(EditProtection::Unrecognized),
        Err(error) => return Err(error),
    };
    let ranges = match Ranges::parse(fib, table_stream) {
        Ok(value) => value.is_some_and(|value| !value.ranges().is_empty()),
        Err(PackageError::Corrupted(_)) => return Ok(EditProtection::Unrecognized),
        Err(error) => return Err(error),
    };
    Ok(match (document, ranges) {
        (false, false) => EditProtection::None,
        (true, false) => EditProtection::Document,
        (false, true) => EditProtection::Ranges,
        (true, true) => EditProtection::DocumentAndRanges,
    })
}

fn document_protected(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    effective_nfib: u16,
) -> Result<bool> {
    let (offset, length) = fib
        .get_table_pointer(DOP_POINTER)
        .ok_or_else(|| PackageError::Corrupted("DOP pointer is absent from the FIB".into()))?;
    if length == 0 {
        return Err(PackageError::Corrupted("lcbDop is zero".into()));
    }
    let start = usize::try_from(offset)
        .map_err(|_| PackageError::Corrupted("DOP offset exceeds usize".into()))?;
    let length = usize::try_from(length)
        .map_err(|_| PackageError::Corrupted("DOP length exceeds usize".into()))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| PackageError::Corrupted("DOP range overflows".into()))?;
    let dop = table_stream
        .get(start..end)
        .ok_or_else(|| PackageError::Corrupted("DOP extends beyond the table stream".into()))?;
    if !dop_length_matches_generation(effective_nfib, dop.len()) {
        return Err(PackageError::Corrupted(format!(
            "DOP length {} does not match effective nFib 0x{effective_nfib:04X}",
            dop.len()
        )));
    }
    if dop.len() < 84 {
        return Err(PackageError::Corrupted(
            "DOP is truncated before its protection fields".into(),
        ));
    }

    // The DOP length selects one exact MS-DOC generation. Parsing only a
    // protection prefix would let a truncated Dop2003 (595..615 bytes) or a
    // producer-specific length fall through as unprotected. The typed parser
    // also checks reserved fields and enumerated values, including the
    // Dop2003 modes 4..6.
    let properties = DocumentProperties::parse_bytes(dop)
        .map_err(|error| PackageError::Corrupted(format!("invalid DOP: {error}")))?;
    properties
        .versioned()
        .map_err(|error| PackageError::Corrupted(format!("invalid DOP extension: {error}")))?;
    let base_protection = properties.base().protection();
    let base = base_protection.comments_or_read_only
        || base_protection.form_fields
        || base_protection.tracked_revisions
        || properties.base().protection_key() != 0;
    // Dop2003.fEnforceDocProt and iDocProtCur are at absolute DOP offsets
    // 598..600. All later specified DOP generations contain this complete
    // 616-byte prefix. Mode 7 explicitly means that no editing restrictions
    // are active; modes 0..3 are restricted and modes 4..6 were rejected by
    // the typed parser above.
    let extension = dop.len() >= 616
        && u16::from_le_bytes([dop[598], dop[599]]) & 0x0008 != 0
        && (u16::from_le_bytes([dop[598], dop[599]]) >> 4) & 0x0007 != 7;
    Ok(base || extension)
}

/// Validates the counted FIB shape needed to interpret protection pointers.
///
/// MS-DOC defines `cbRgFcLcb` from the effective `nFib`, which is
/// `FibRgCswNew.nFibNew` whenever `cswNew` is nonzero.  The parser deliberately
/// keeps unknown trailing bytes, so this check is performed at the policy
/// boundary before any protection pointer is treated as absent.
fn validate_fib_shape(fib: &FileInformationBlock) -> Result<u16> {
    let count = fib
        .table_pointer_count()
        .ok_or_else(|| PackageError::Corrupted("FIB table-pointer array is truncated".into()))?;
    let pointer_end = fib
        .minimum_serialized_size()
        .ok_or_else(|| PackageError::Corrupted("FIB table-pointer size overflows".into()))?;
    let raw = fib.raw_data();
    if read_u16(raw, FIB_CSW, "FIB csw")? != 0x000E {
        return Err(PackageError::Corrupted("FIB csw is not 0x000E".into()));
    }
    if read_u16(raw, FIB_CSLW, "FIB cslw")? != 0x0016 {
        return Err(PackageError::Corrupted("FIB cslw is not 0x0016".into()));
    }
    let csw_new = read_u16(raw, pointer_end, "FIB cswNew")?;
    let effective_nfib = if csw_new == 0 {
        fib.version()
    } else {
        // nFibNew is the first short in FibRgCswNew.  A nonzero cswNew with no
        // complete nFibNew is malformed even if the counted pointer array is
        // otherwise present.
        read_u16(
            raw,
            pointer_end
                .checked_add(2)
                .ok_or_else(|| PackageError::Corrupted("FIB nFibNew offset overflows".into()))?,
            "FIB nFibNew",
        )?
    };
    let expected = match effective_nfib {
        WORD97_NFIB => 0x005D,
        WORD2000_NFIB => 0x006C,
        WORD2002_NFIB => 0x0088,
        WORD2003_NFIB => 0x00A4,
        WORD2007_NFIB => 0x00B7,
        _ => {
            return Err(PackageError::Corrupted(format!(
                "unsupported effective FIB version 0x{effective_nfib:04X}"
            )));
        },
    };
    let expected_csw_new = match effective_nfib {
        WORD97_NFIB => 0,
        WORD2000_NFIB | WORD2002_NFIB | WORD2003_NFIB => 2,
        WORD2007_NFIB => 5,
        _ => unreachable!("effective nFib was checked above"),
    };
    let csw_new_accepted = csw_new == expected_csw_new;
    if !csw_new_accepted {
        return Err(PackageError::Corrupted(format!(
            "FIB cswNew {csw_new} does not match effective nFib 0x{effective_nfib:04X} ({expected_csw_new})"
        )));
    }
    let csw_bytes = usize::from(csw_new)
        .checked_mul(2)
        .ok_or_else(|| PackageError::Corrupted("FIB cswNew byte count overflows".into()))?;
    raw.get(
        pointer_end
            ..pointer_end
                .checked_add(csw_bytes)
                .ok_or_else(|| PackageError::Corrupted("FIB FibRgCswNew range overflows".into()))?,
    )
    .ok_or_else(|| PackageError::Corrupted("FIB FibRgCswNew is truncated".into()))?;
    if count != expected {
        return Err(PackageError::Corrupted(format!(
            "FIB cbRgFcLcb {count} does not match effective nFib 0x{effective_nfib:04X} ({expected})"
        )));
    }
    Ok(effective_nfib)
}

/// Returns whether an exact DOP payload length is valid for the selected FIB
/// generation.  The FIB generation and DOP wire size are a coupled version
/// check; a Word 2002 FIB therefore requires its 594-byte Dop2002 payload.
fn dop_length_matches_generation(nfib: u16, length: usize) -> bool {
    match nfib {
        WORD97_NFIB => length == 500,
        WORD2000_NFIB => length == 544,
        WORD2002_NFIB => length == 594,
        WORD2003_NFIB => length == 616,
        WORD2007_NFIB => matches!(length, 674 | 690 | 694),
        _ => false,
    }
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let bytes = data
        .get(
            offset
                ..offset
                    .checked_add(2)
                    .ok_or_else(|| PackageError::Corrupted(format!("{field} offset overflows")))?,
        )
        .ok_or_else(|| PackageError::Corrupted(format!("{field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

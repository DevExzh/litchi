//! Publication policy for legacy Word document protection.

use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use litchi_ole_common::object::Patch as ObjectPatch;

use super::Ranges;

const DOP_POINTER: usize = 31;
const FIB_CSW: usize = 32;
const FIB_CSLW: usize = 62;
const WORD97_NFIB: u16 = 0x00C1;
/// `FibBase.nFib` of the empty document that Word 97 through Word 2003 install
/// for the shell's "New Word Document" command (MS-DOC Appendix A note <11>).
const WORD97_SHELL_TEMPLATE_NFIB: u16 = 0x00C0;
/// `FibBase.nFib` written by the BiDi build of Word 97 (note <11>).
const WORD97_BIDI_NFIB: u16 = 0x00C2;
const WORD2000_NFIB: u16 = 0x00D9;
const WORD2002_NFIB: u16 = 0x0101;
const WORD2003_NFIB: u16 = 0x010C;
const WORD2007_NFIB: u16 = 0x0112;

/// Size of `DopBase`, the prefix shared by every DOP generation (2.7.2).
const DOP_BASE_SIZE: usize = 84;
/// `DopBase` byte holding `fFormNoFields` (0x20) and `fRevMarking` (0x80).
const DOP_FORM_AND_REVISION_BYTE: usize = 5;
const DOP_FORM_NO_FIELDS: u8 = 0x20;
const DOP_REVISION_MARKING: u8 = 0x80;
/// `DopBase` byte holding `fLockAtn` (0x10).
const DOP_ANNOTATION_LOCK_BYTE: usize = 6;
const DOP_LOCK_ANNOTATIONS: u8 = 0x10;
/// `DopBase` byte holding `fProtEnabled` (0x02) and `fLockRev` (0x40).
const DOP_FORM_AND_REVISION_LOCK_BYTE: usize = 7;
const DOP_PROTECT_FORMS: u8 = 0x02;
const DOP_LOCK_REVISIONS: u8 = 0x40;
/// `DopBase.lKeyProtDoc`, the document-protection password hash.
const DOP_PROTECTION_KEY: usize = 78;
/// The `Dop2003` byte holding `fEnforceDocProt` (bit 3) and `iDocProtCur`
/// (bits 4..6); `empty2` occupies the following byte (2.7.7).
const DOP2003_PROTECTION_BYTE: usize = 598;
const DOP2003_ENFORCE_PROTECTION: u8 = 0x08;
const DOP2003_PROTECTION_MODE_SHIFT: u8 = 4;
const DOP2003_PROTECTION_MODE_MASK: u8 = 0x07;
/// `iDocProtCur` 7: "There are no editing restrictions."
const DOP2003_UNRESTRICTED_MODE: u8 = 7;
/// End of the complete 16-bit `Dop2003` unit that carries the enforcement
/// fields. Generations from Word 2003 on define it, so their DOP must reach it.
const DOP2003_PROTECTION_END: usize = 600;

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
    /// A structurally valid host carries a protection record outside the
    /// MS-DOC protection grammar, so its protection state is not known
    /// precisely.
    ///
    /// This covers a DOP outside the table stream or too short to hold the
    /// protection fields of its generation, DOP protection fields that break
    /// an MS-DOC requirement (for example `fLockAtn` together with `fLockRev`,
    /// or a reserved `iDocProtCur`), and a malformed range-protection table. It remains
    /// distinct from [`Self::None`]. It is retained for inspection and exact
    /// no-op publication only; an authorization cannot turn an unrecognized
    /// protection record into a known safe state.
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
    if !dop_holds_generation_protection_fields(effective_nfib, dop.len()) {
        return Err(PackageError::Corrupted(format!(
            "DOP length {} cannot hold the protection fields of effective nFib 0x{effective_nfib:04X}",
            dop.len()
        )));
    }
    dop_restricts_editing(dop)
}

/// Reads the protection-bearing DOP fields, and only those.
///
/// MS-DOC places every document-wide editing restriction in `DopBase`
/// (`fLockAtn`, `fProtEnabled`, `fLockRev` and the password hash
/// `lKeyProtDoc`, 2.7.2) and in the `Dop2003` byte at offset 598
/// (`fEnforceDocProt` and `iDocProtCur`, 2.7.7). `Dop2007`, `Dop2010` and
/// `Dop2013` add no protection field. Every other DOP field is display,
/// statistics, compatibility or typography state that cannot restrict an
/// edit, and several are specified as "MUST be ignored"; this reader does not
/// validate them, so a producer's value there never becomes a protection
/// verdict.
///
/// The requirements MS-DOC places on the protection fields themselves stay
/// enforced; a DOP that breaks one is returned as `Corrupted`, which the
/// classifier reports as [`EditProtection::Unrecognized`]:
/// `fLockAtn` together with `fLockRev`, `fLockRev` without `fRevMarking`,
/// `fFormNoFields` without `fProtEnabled` (all 2.7.2), and a reserved
/// `iDocProtCur` value 4..6 (2.7.7). `fProtEnabled` together with `fLockAtn`
/// or `fLockRev` is only a SHOULD NOT, which Word 97-2003 is documented to
/// break (Appendix A notes <164>, <165> and <167>); such a DOP is protected.
///
/// `fEnforceDocProt` and `iDocProtCur` lie wholly in byte 598, so they are
/// read whenever the DOP reaches that byte, whatever the generation. Producers
/// such as LibreOffice write the `Dop2003` layout into a longer DOP under an
/// older `nFib`, and reading bytes that could carry an enforced restriction
/// can only add a refusal.
fn dop_restricts_editing(dop: &[u8]) -> Result<bool> {
    let base = dop
        .get(..DOP_BASE_SIZE)
        .ok_or_else(|| PackageError::Corrupted("DOP is truncated before DopBase ends".into()))?;
    let revision_marking = base[DOP_FORM_AND_REVISION_BYTE] & DOP_REVISION_MARKING != 0;
    let form_no_fields = base[DOP_FORM_AND_REVISION_BYTE] & DOP_FORM_NO_FIELDS != 0;
    let lock_annotations = base[DOP_ANNOTATION_LOCK_BYTE] & DOP_LOCK_ANNOTATIONS != 0;
    let protect_forms = base[DOP_FORM_AND_REVISION_LOCK_BYTE] & DOP_PROTECT_FORMS != 0;
    let lock_revisions = base[DOP_FORM_AND_REVISION_LOCK_BYTE] & DOP_LOCK_REVISIONS != 0;
    let key = u32::from_le_bytes([
        base[DOP_PROTECTION_KEY],
        base[DOP_PROTECTION_KEY + 1],
        base[DOP_PROTECTION_KEY + 2],
        base[DOP_PROTECTION_KEY + 3],
    ]);
    if lock_annotations && lock_revisions {
        return Err(PackageError::Corrupted(
            "DopBase.fLockAtn and fLockRev are both set".into(),
        ));
    }
    if lock_revisions && !revision_marking {
        return Err(PackageError::Corrupted(
            "DopBase.fLockRev requires fRevMarking".into(),
        ));
    }
    if form_no_fields && !protect_forms {
        return Err(PackageError::Corrupted(
            "DopBase.fFormNoFields requires fProtEnabled".into(),
        ));
    }
    let base_restricted = lock_annotations || protect_forms || lock_revisions || key != 0;
    let enforced_restriction = match dop.get(DOP2003_PROTECTION_BYTE) {
        None => false,
        Some(&flags) => {
            let mode = (flags >> DOP2003_PROTECTION_MODE_SHIFT) & DOP2003_PROTECTION_MODE_MASK;
            if (4..DOP2003_UNRESTRICTED_MODE).contains(&mode) {
                return Err(PackageError::Corrupted(format!(
                    "Dop2003.iDocProtCur has reserved value {mode}"
                )));
            }
            flags & DOP2003_ENFORCE_PROTECTION != 0 && mode != DOP2003_UNRESTRICTED_MODE
        },
    };
    Ok(base_restricted || enforced_restriction)
}

/// Validates the counted FIB shape needed to interpret protection pointers
/// and returns the effective `nFib`.
///
/// MS-DOC 2.5.14 takes `nFib` from `FibRgCswNew.nFibNew` when `cswNew` is
/// nonzero and from `FibBase.nFib` otherwise, and Appendix A note <11> tells
/// readers to treat a `FibBase.nFib` of 0x00C0 or 0x00C2 as 0x00C1.
/// `cbRgFcLcb` must equal the count 2.5.1 assigns to that generation, because
/// it bounds the pointer array that locates the DOP and the range-protection
/// tables. The parser deliberately keeps unknown trailing bytes, so this check
/// is performed at the policy boundary before any protection pointer is
/// treated as absent.
///
/// A nonzero `cswNew` must be the count 2.5.1 assigns to its `nFibNew`, which
/// must be present. A zero `cswNew` is accepted for every
/// generation although 2.5.1 asks for 2 or 5 from Word 2000 on: LibreOffice
/// writes `nFib` 0x0101 with a zero `cswNew`, `FibRgCswNew` carries no
/// protection data, and `cbRgFcLcb` must still match the generation that
/// 2.5.14 selects.
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
        match fib.version() {
            WORD97_SHELL_TEMPLATE_NFIB | WORD97_BIDI_NFIB => WORD97_NFIB,
            nfib => nfib,
        }
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
    let (expected, expected_csw_new) = match effective_nfib {
        WORD97_NFIB => (0x005D, 0),
        WORD2000_NFIB => (0x006C, 2),
        WORD2002_NFIB => (0x0088, 2),
        WORD2003_NFIB => (0x00A4, 2),
        WORD2007_NFIB => (0x00B7, 5),
        _ => {
            return Err(PackageError::Corrupted(format!(
                "unsupported effective FIB version 0x{effective_nfib:04X}"
            )));
        },
    };
    if csw_new != 0 && csw_new != expected_csw_new {
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

/// Returns whether a DOP of `length` bytes holds every protection field the
/// selected FIB generation defines.
///
/// Every generation begins with the 84-byte `DopBase`, which carries the
/// document locks and the password hash. From Word 2003 (0x010C) on the
/// generation also defines the `Dop2003` enforcement unit at bytes 598..600,
/// so a shorter DOP (such as a truncated `Dop2003` of 595..599 bytes) cannot
/// be proven unprotected. For 0x0112, MS-DOC 2.7.1 requires `lcbDop` to be
/// exactly 674, 690 or 694. MS-DOC states no other exact length: a longer
/// DOP only carries fields that restrict nothing or, at byte 598, fields
/// [`dop_restricts_editing`] reads, and a field beyond the recorded length
/// takes its specified default, which for every protection field is
/// unrestricted.
fn dop_holds_generation_protection_fields(nfib: u16, length: usize) -> bool {
    match nfib {
        WORD97_NFIB | WORD2000_NFIB | WORD2002_NFIB => length >= DOP_BASE_SIZE,
        WORD2003_NFIB => length >= DOP2003_PROTECTION_END,
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

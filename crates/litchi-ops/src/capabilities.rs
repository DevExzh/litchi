//! What a file can do through the unified API (record 0010 §7, M1 step R5):
//! [`File::capabilities`] and [`Capabilities::of`].
//!
//! A report names, for one format:
//!
//! - the **state** of the file: [`State::Open`], [`State::Signed`] (reads
//!   allowed) or [`State::Locked`] (every read `document_protected` until
//!   [`Locked::unlock`](crate::Locked)), and whether the handle retained
//!   what an editor is built from;
//! - the **verbs** ([`Verb`]) and **view modes** ([`Mode`]) it answers, each
//!   [`Support::Supported`] or [`Support::Unsupported`] with the code the
//!   refusal returns and its follow-up. Whole-format and protection refusals
//!   take precedence; DOCX exposes its typed property and direct structural
//!   subset, creation and publication while the remaining write verbs remain
//!   `unsupported_capability`, as do unimplemented view modes;
//! - the **elements** ([`Element`]) its resolver knows, in OfficeCLI's help
//!   schema vocabulary: type name and aliases, parents, path forms, children
//!   with their path segments, operations, and properties with their value
//!   type and `get`/`set`/`add` flags, and the keys a selector compares.
//!   Native identity fields and templates stay out of ordinary reports;
//! - the **gaps** ([`Gap`]): registered elements, axes, properties, operation
//!   contexts and view modes the format has that Litchi cannot answer yet,
//!   with the code and the follow-up that closes it.
//!
//! # Generation
//!
//! The report is generated from the same static `ElementSpec`/`PropSpec`
//! table each format adapter's resolver, query compiler and node view run on
//! (record 0010 §7.2). Sharing metadata reduces drift; behavior is also
//! checked by:
//!
//! 1. the spec tables' `const` checks (`check_format_spec!`), and the
//!    capability checks this module runs on every adapter's table
//!    (`check_capabilities!`): path templates, identity keys and gaps are
//!    well formed, and a property is never both readable and a gap;
//! 2. exhaustive `match`es on the adapter list, [`Verb`] and [`Mode`], so an
//!    adapter, verb or mode without its arm does not compile;
//! 3. DOCX, XLSX, XLSB, XLS and PPTX fixture tests checking reported path
//!    templates, queryable kinds, emitted property membership and view modes,
//!    plus tests of protected and pending-format refusals. Complete property liveness
//!    and OfficeCLI parity remain follow-up work;
//! 4. a golden JSON snapshot per format (`tests/golden/capabilities/`,
//!    regenerated only with `LITCHI_BLESS=1`), beside the registry export's
//!    snapshots (`tests/golden/registry/`, [`crate::schema()`]), from which
//!    the OfficeCLI delta table is generated
//!    (`tools/officecli_capability_delta.py`).
//!
//! # JSON
//!
//! `Capabilities` serializes as OfficeCLI's help schema
//! (`schemas/help/_schema.json`) for each element, with the report's state,
//! verbs, views and gaps around them:
//!
//! ```json
//! {
//!   "format": "docx",
//!   "state": "open",
//!   "retainsSource": true,
//!   "verbs": { "get": { "supported": true } },
//!   "elements": [
//!     {
//!       "element": "paragraph",
//!       "paths": { "positional": ["/body/p[N]"] },
//!       "properties": { "style": { "type": "string", "add": false, "set": false, "get": true } }
//!     }
//!   ]
//! }
//! ```
//!
//! This abbreviated example omits other verbs, elements, views and gaps.
//!
//! A path template is *positional* when a step selects by ordinal (`[N]`),
//! *stable* otherwise (semantic names and coordinates). A child's
//! `pathSegment` is the last step of its first rooted template, without
//! predicates.
//!
//! # The M2 seam
//!
//! M1's `PropSpec` is the read-only subset of an M2 property-registry entry
//! (record 0010 §7.3). The write side attaches in two places, without
//! changing this report's shape: each property's access flags (M2's registry
//! entries turn `set` and `add` on where a typed setter exists), and
//! [`Element::operations`], whose `add`, `set` and `remove` M2's handler
//! table answers. DOCX's property and direct structural handlers populate
//! those flags; unsupported selection and leaf contexts remain explicit
//! operation gaps, as do the unimplemented write handlers of other formats.

// In a build without any adapter, `Adapter` (and so `Element`) is
// uninhabited: everything that holds one is unreachable, and the checks and
// helpers that read spec tables are unused. Those lints are allowed in such
// builds only; each adapter adds its feature to this list.
#![cfg_attr(
    not(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb",
        feature = "xls",
        feature = "rtf",
        feature = "odp"
    )),
    allow(
        dead_code,
        unreachable_code,
        unused_imports,
        unused_macros,
        unused_mut,
        unused_variables
    )
)]

use std::fmt;

use litchi_core::FileFormat;
use serde::ser::{Serialize, SerializeMap, SerializeSeq, SerializeStruct, Serializer};

use crate::error::{Code, Error};
pub use crate::node::spec::Subject;
use crate::node::spec::{GapSpec, PropSpec};
use crate::read::{Adapter, with_spec};
pub use crate::registry::Constraint;
use crate::value::ValueType;
use crate::view::Mode;
use crate::{Family, File};

/// What one file, or one format, can do through the unified API.
///
/// Built by [`File::capabilities`] or [`Capabilities::of`]; cheap to copy,
/// and allocation-free to read. `serde::Serialize` writes OfficeCLI's help
/// schema vocabulary (see the [module documentation](self)).
///
/// ```
/// use litchi_core::FileFormat;
/// use litchi_ops::Code;
/// use litchi_ops::capabilities::{Capabilities, Support, Verb};
///
/// // A build without the `pages` feature refuses Apple Pages as a whole.
/// let pages = Capabilities::of(FileFormat::Pages);
/// if cfg!(not(feature = "pages")) {
///     match pages.verb(Verb::Get) {
///         Support::Unsupported { code, follow_up, .. } => {
///             assert_eq!(code, Code::UnsupportedFormat);
///             assert_eq!(follow_up, "enable the `pages` cargo feature");
///         },
///         other => panic!("{other:?}"),
///     }
///     assert_eq!(pages.elements().count(), 0);
/// }
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Capabilities {
    format: Option<FileFormat>,
    state: State,
    retains_source: Option<bool>,
    /// For a file report, whether the PPTX owner can add slides.  Source
    /// owners have this capability through the source-native transaction.
    pptx_slide_add: Option<bool>,
    /// For a file report, whether the PPTX owner can add top-level text-box
    /// shapes through the native source or eager transaction.
    pptx_shape_add: Option<bool>,
    /// For a file report, whether the PPTX owner can remove dependency-free
    /// slides. Removal and slide or top-level text-box insertion have
    /// source-native paths.
    pptx_slide_remove: Option<bool>,
    /// For a file report, whether the PPTX owner can remove top-level slide
    /// shapes. The native path still refuses grouped, master/layout, and
    /// dependency-bearing shapes during operation admission.
    pptx_shape_remove: Option<bool>,
    coverage: Coverage,
    ordinary_output_protected: bool,
}

/// How far the unified read surface covers a format in this build.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Coverage {
    /// The format's adapter has landed: its spec table drives the report.
    Adapter(Adapter),
    /// Reads of the format are refused as a whole.
    Refused(Refusal),
}

/// A whole-format refusal: the code, its follow-up and why.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Refusal {
    code: Code,
    follow_up: &'static str,
    reason: &'static str,
}

impl Refusal {
    const fn support(self) -> Support {
        Support::Unsupported {
            code: self.code,
            follow_up: self.follow_up,
            reason: self.reason,
        }
    }
}

/// The state of a file, as the unified API sees it.
///
/// ```
/// use litchi_ops::capabilities::State;
///
/// assert_eq!(State::Signed.as_str(), "signed");
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum State {
    /// Readable; DOCX's retained-source typed setters are also available.
    Open,
    /// Signed: readable; editing must consume the signature first
    /// ([`Signed::unsign`](crate::Signed::unsign)).
    Signed,
    /// Encrypted: nothing is readable until
    /// [`Locked::unlock`](crate::Locked) consumes the password.
    Locked,
}

impl State {
    /// The JSON spelling: `open`, `signed` or `locked`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Open => "open",
            Self::Signed => "signed",
            Self::Locked => "locked",
        }
    }
}

/// A verb of the unified API: reads, edits, publication and creation.
///
/// ```
/// use litchi_ops::capabilities::Verb;
///
/// assert_eq!(Verb::Query.as_str(), "query");
/// assert!(Verb::Set.is_edit() && !Verb::Get.is_edit());
/// assert!(!Verb::Save.is_edit() && !Verb::Create.is_edit());
/// assert!(Verb::Import.is_edit() && !Verb::Evaluate.is_edit());
/// assert!(Verb::Merge.is_edit());
/// assert!(Verb::Replace.is_edit());
/// assert!(Verb::Refresh.is_edit());
/// assert_eq!(Verb::ALL.len(), 23);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Verb {
    /// [`File::get`] and [`File::root`].
    Get,
    /// [`File::find`].
    Find,
    /// [`File::query`].
    Query,
    /// [`File::view`], [`File::view_to`] and [`File::view_json`]; each mode
    /// is reported on its own ([`Capabilities::view`]).
    View,
    /// Formula evaluation: `File::evaluate` and `File::recalculate`
    /// (record 0266, the `eval` feature).
    Evaluate,
    /// Set properties (M2).
    Set,
    /// Add an element (M2).
    Add,
    /// Remove an element (M2).
    Remove,
    /// Move an element (M2).
    Move,
    /// Swap two elements (M2).
    Swap,
    /// Group mutations in an atomic [`File::edit`] closure.
    Edit,
    /// [`File::batch`] and its single-command [`File::apply`] channel.
    Batch,
    /// CSV/TSV import into a worksheet, [`File::import`] (record 0266).
    Import,
    /// Template merge, [`File::merge`] (record 0274).
    Merge,
    /// Text find/replace: [`File::set`] with the `find=` and `replace=`
    /// properties, OfficeCLI's spelling (record 0273). The match count is
    /// [`crate::Applied::matched`].
    Replace,
    /// Field and table-of-contents refresh, [`File::refresh`] (record 0279).
    Refresh,
    /// Save in the same format (W4).
    Save,
    /// Create a new package of this format (W4).
    Create,
    /// [`File::dump`]: the document as an OfficeCLI v2 batch (record 0336).
    Dump,
    /// [`File::raw`] and [`File::parts`]: package parts as bytes (design
    /// §12, record 0336).
    Raw,
    /// [`File::validate`]: read-only structural validation (record 0336).
    Validate,
    /// [`File::add_part`]: a package part with its content type and
    /// relationship (record 0336).
    AddPart,
    /// [`File::raw_set`]: an audited XPath-addressed XML edit of one part
    /// (record 0336).
    RawSet,
}

impl Verb {
    /// Every verb, in declaration order: the reads, then the edits,
    /// publication and creation.
    pub const ALL: [Self; 23] = [
        Self::Get,
        Self::Find,
        Self::Query,
        Self::View,
        Self::Evaluate,
        Self::Set,
        Self::Add,
        Self::Remove,
        Self::Move,
        Self::Swap,
        Self::Edit,
        Self::Batch,
        Self::Import,
        Self::Merge,
        Self::Replace,
        Self::Refresh,
        Self::Save,
        Self::Create,
        Self::Dump,
        Self::Raw,
        Self::Validate,
        Self::AddPart,
        Self::RawSet,
    ];

    /// The verb's canonical name, using OfficeCLI command names where applicable.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Get => "get",
            Self::Find => "find",
            Self::Query => "query",
            Self::View => "view",
            Self::Evaluate => "evaluate",
            Self::Set => "set",
            Self::Add => "add",
            Self::Remove => "remove",
            Self::Move => "move",
            Self::Swap => "swap",
            Self::Edit => "edit",
            Self::Batch => "batch",
            Self::Import => "import",
            Self::Merge => "merge",
            Self::Replace => "replace",
            Self::Refresh => "refresh",
            Self::Save => "save",
            Self::Create => "create",
            Self::Dump => "dump",
            Self::Raw => "raw",
            Self::Validate => "validate",
            Self::AddPart => "add-part",
            Self::RawSet => "raw-set",
        }
    }

    /// Whether the verb changes an existing document (M2).
    ///
    /// Saving publishes the current snapshot and creating produces a new
    /// package, so neither operation is an edit of an existing document.
    #[must_use]
    pub const fn is_edit(self) -> bool {
        match self {
            Self::Set
            | Self::Add
            | Self::Remove
            | Self::Move
            | Self::Swap
            | Self::Edit
            | Self::Batch
            | Self::Import
            | Self::Merge
            | Self::Replace
            | Self::Refresh
            | Self::AddPart
            | Self::RawSet => true,
            Self::Get
            | Self::Find
            | Self::Query
            | Self::View
            | Self::Evaluate
            | Self::Save
            | Self::Create
            | Self::Dump
            | Self::Raw
            | Self::Validate => false,
        }
    }
}

const _: () = {
    let mut index = 0;
    while index < Verb::ALL.len() {
        assert!(
            Verb::ALL[index] as usize == index,
            "Verb::ALL is out of order"
        );
        index += 1;
    }
};

/// Whether a verb, view mode or element operation is available, and the
/// code its refusal returns when it is not.
///
/// ```
/// use litchi_ops::Code;
/// use litchi_ops::capabilities::Support;
///
/// let support = Support::Unsupported {
///     code: Code::UnsupportedCapability,
///     follow_up: "M2",
///     reason: "edit verbs arrive with the write kernel",
/// };
/// assert!(!support.is_supported());
/// assert_eq!(support.code(), Some(Code::UnsupportedCapability));
/// assert!(Support::Supported.is_supported());
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Support {
    /// Available.
    Supported,
    /// Refused with `code`.
    Unsupported {
        /// The code the refusal returns.
        code: Code,
        /// What makes it available: an action (`Locked::unlock`), a cargo
        /// feature, a leaf request (`L-XLSX-R5`) or a milestone (`M2`).
        follow_up: &'static str,
        /// One sentence on why.
        reason: &'static str,
    },
}

impl Support {
    /// Whether the operation is available.
    #[must_use]
    pub const fn is_supported(self) -> bool {
        matches!(self, Self::Supported)
    }

    /// The code a refusal returns; `None` when supported.
    #[must_use]
    pub const fn code(self) -> Option<Code> {
        match self {
            Self::Supported => None,
            Self::Unsupported { code, .. } => Some(code),
        }
    }

    fn from_gap(gap: &GapSpec) -> Self {
        Self::Unsupported {
            code: gap.code,
            follow_up: gap.follow_up,
            reason: gap.reason,
        }
    }
}

/// The refusal of every read of a locked file.
const LOCKED: Refusal = Refusal {
    code: Code::DocumentProtected,
    follow_up: "Locked::unlock",
    reason: "the document is encrypted; reads need its password",
};

/// Signed owners require explicit signature consumption before edits.
const SIGNED: Refusal = Refusal {
    code: Code::DocumentProtected,
    follow_up: "Signed::unsign",
    reason: "editing requires explicit signature consumption",
};

/// The structural verbs the iWork leaves do not offer yet (record 0296):
/// element creation, and removal, moving and swapping where the leaf has
/// no such transaction (`IWORK_API_REQUESTS.md`).
#[cfg_attr(
    not(feature = "iwork-adapter"),
    allow(dead_code, reason = "read only when an iWork format is compiled in")
)]
const IWORK_STRUCTURE: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "IWORK_API_REQUESTS.md",
    reason: "the iWork leaf has no such structural transaction yet",
};

/// The refusal of every edit verb in M1.
const EDITS: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "M2",
    reason: "edit verbs arrive with the write kernel in milestone M2",
};

/// The refusal of publication for a format whose save owner has not landed
/// in the current W4 wave.
const SAVE_PENDING: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "W4",
    reason: "saving this format is not integrated in the current wave",
};

/// The refusal of creation for a format whose creator has not landed in the
/// current W4 wave.
const CREATE_PENDING: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "W4",
    reason: "creation is not integrated for this format in the current wave",
};

/// A retained edit source is required for file-specific writes.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
const RETAINED_SOURCE: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "reopen with OpenOptions::retain_source(true)",
    reason: "writing requires the retained source used to publish the edit",
};

/// The fallback refusal for a PPTX owner that cannot add either supported
/// element. Retained source owners normally expose both slide and top-level
/// text-box insertion; the value remains for defensive owner states.
#[cfg(feature = "pptx")]
const PPTX_SHAPE_ADD: Refusal = Refusal {
    code: Code::UnsupportedCapability,
    follow_up: "OpenOptions::retain_source(true)",
    reason: "this PPTX owner cannot add a slide or top-level text-box shape",
};

/// The refusal of a locked OOXML package whose format its name does not
/// tell.
const UNKNOWN: Refusal = Refusal {
    code: Code::DocumentProtected,
    follow_up: "Locked::unlock",
    reason: "the encrypted package hides its format until it is unlocked",
};

impl Capabilities {
    /// What the unified API can do with a file of `format` in this build,
    /// opened with the default options: the counterpart of OfficeCLI's
    /// `help <format>`, without a file.
    ///
    #[must_use]
    pub fn of(format: FileFormat) -> Self {
        Self {
            format: Some(format),
            state: State::Open,
            retains_source: None,
            pptx_slide_add: (format == FileFormat::Pptx).then_some(true),
            pptx_shape_add: (format == FileFormat::Pptx).then_some(true),
            pptx_slide_remove: (format == FileFormat::Pptx).then_some(true),
            pptx_shape_remove: (format == FileFormat::Pptx).then_some(true),
            coverage: coverage(Some(format)),
            ordinary_output_protected: false,
        }
    }

    /// The report of `file`.
    pub(crate) fn for_file(file: &File) -> Self {
        let format = file.format();
        let state = match file {
            File::Locked(_) => State::Locked,
            file if file.is_signed() => State::Signed,
            _ => State::Open,
        };
        Self {
            format,
            state,
            retains_source: retains_source(file),
            pptx_slide_add: pptx_slide_add(file),
            pptx_shape_add: pptx_shape_add(file),
            pptx_slide_remove: pptx_slide_remove(file),
            pptx_shape_remove: pptx_shape_remove(file),
            coverage: coverage(format),
            ordinary_output_protected: {
                let mut protected = false;
                #[cfg(feature = "xls")]
                {
                    let file = match file {
                        File::Signed(signed) => signed.file(),
                        file => file,
                    };
                    if let File::Xls(handle) = file {
                        protected = handle.is_encrypted();
                    }
                }
                #[cfg(feature = "ppt")]
                {
                    let file = match file {
                        File::Signed(signed) => signed.file(),
                        file => file,
                    };
                    if let File::Ppt(handle) = file {
                        protected = handle.is_encrypted();
                    }
                }
                #[cfg(feature = "encryption")]
                {
                    let file = match file {
                        File::Signed(signed) => signed.file(),
                        file => file,
                    };
                    #[cfg(feature = "xlsx")]
                    if let File::Xlsx(handle) = file {
                        protected = handle.requires_reencryption();
                    }
                    #[cfg(feature = "pptx")]
                    if let File::Pptx(handle) = file {
                        protected = handle.requires_reencryption();
                    }
                    let _ = file;
                }
                let _ = &mut protected;
                protected
            },
        }
    }

    /// The error a refused read returns: a locked file's, or the whole
    /// format's. An `internal_error` when reads are supported, which the
    /// read entry points never ask for.
    pub(crate) fn refusal(&self) -> Error {
        let refusal = match self.read_refusal() {
            Some(refusal) => refusal,
            None => {
                return Error::new(
                    Code::InternalError,
                    "a supported read was reported as refused",
                );
            },
        };
        let what = match self.format {
            Some(format) => format!("{format:?} documents"),
            None => "this document".to_owned(),
        };
        let error = Error::new(
            refusal.code,
            format!(
                "unified reads (get, find, query, view) of {what} are not available: {}",
                refusal.reason
            ),
        )
        .with_suggestion(format!(
            "follow-up: {}; `File::capabilities()` lists what is available",
            refusal.follow_up
        ));
        match self.format {
            Some(format) => error.with_format(format),
            None => error,
        }
    }

    /// The format, when it is known: `None` only for an encrypted OOXML
    /// package whose name does not tell its format.
    #[must_use]
    pub const fn format(&self) -> Option<FileFormat> {
        self.format
    }

    /// The file's state.
    #[must_use]
    pub const fn state(&self) -> State {
        self.state
    }

    /// Whether the handle retained what an editor is built from
    /// ([`OpenOptions::retain_source`](crate::OpenOptions::retain_source)):
    /// `None` for formats without retention reporting (RTF and OpenDocument), a locked
    /// file, and a report made by [`Capabilities::of`]. A flat OpenDocument
    /// presentation reports `Some(false)`: it has no editable package.
    #[must_use]
    pub const fn retains_source(&self) -> Option<bool> {
        self.retains_source
    }

    /// The whole-format refusal of every read, if any.
    const fn read_refusal(&self) -> Option<Refusal> {
        if let State::Locked = self.state {
            return Some(LOCKED);
        }
        match self.coverage {
            Coverage::Adapter(_) => None,
            Coverage::Refused(refusal) => Some(refusal),
        }
    }

    /// The adapter whose spec table the report lists, if any.
    const fn adapter(&self) -> Option<Adapter> {
        match self.coverage {
            Coverage::Adapter(adapter) => Some(adapter),
            Coverage::Refused(_) => None,
        }
    }

    /// Whether `verb` is available, or the code it refuses with.
    #[must_use]
    pub fn verb(&self, verb: Verb) -> Support {
        if self.state == State::Locked {
            return LOCKED.support();
        }
        // Creation is a format-level operation for `Capabilities::of`, but
        // it is not an operation on an already signed document. Keep a
        // contextual report from advertising it as usable on that handle.
        if (verb.is_edit() || verb == Verb::Create) && self.state == State::Signed {
            return SIGNED.support();
        }
        if self.ordinary_output_protected && (verb.is_edit() || verb == Verb::Save) {
            if self.format == Some(FileFormat::Xls) {
                return Refusal {
                    code: Code::DocumentProtected,
                    follow_up: "M2 XLS encryption-preserving publication",
                    reason: "the unlocked XLS owner refuses ordinary publication to prevent plaintext output",
                }.support();
            }
            #[cfg(feature = "ppt")]
            if self.format == Some(FileFormat::Ppt) {
                return Refusal {
                    code: Code::DocumentProtected,
                    follow_up: "password-preserving PPT publication",
                    reason: "the unlocked encrypted presentation refuses ordinary plaintext publication",
                }
                .support();
            }
            #[cfg(feature = "pptx")]
            if self.adapter() == Some(Adapter::Pptx) {
                return Refusal { code: Code::DocumentProtected, follow_up: "use SaveOptions::reencrypt for encrypted publication; native encrypted edits remain pending", reason: "the unlocked encrypted presentation refuses ordinary plaintext publication and native editing" }.support();
            }
            return Refusal {
                code: Code::DocumentProtected,
                follow_up: "SaveOptions::reencrypt; M2 encryption-profile-preserving native edits",
                reason: "the unlocked encrypted workbook requires re-encryption for publication and refuses native editing",
            }.support();
        }
        if matches!(verb, Verb::Import | Verb::Evaluate) {
            return self.m4_support(verb);
        }
        if verb == Verb::Merge {
            return self.merge_support();
        }
        if verb == Verb::Replace {
            return self.replace_support();
        }
        if verb == Verb::Refresh {
            return self.refresh_support();
        }
        if verb == Verb::Create {
            return self.creation_support();
        }
        if verb == Verb::Save {
            return self.publication_support();
        }
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        if let Some(support) = self.raw_support(verb) {
            return support;
        }
        if verb.is_edit() {
            #[cfg(feature = "doc")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Doc)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "docx")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Docx)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "xlsx")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Xlsx)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "pptx")]
            if verb == Verb::Add && self.adapter() == Some(Adapter::Pptx) {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                if self.pptx_slide_add != Some(true) && self.pptx_shape_add != Some(true) {
                    return PPTX_SHAPE_ADD.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "pptx")]
            if matches!(verb, Verb::Move | Verb::Swap) && self.adapter() == Some(Adapter::Pptx) {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "pptx")]
            if verb == Verb::Remove && self.adapter() == Some(Adapter::Pptx) {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                if self.pptx_slide_remove == Some(true) || self.pptx_shape_remove == Some(true) {
                    return Support::Supported;
                }
                return RETAINED_SOURCE.support();
            }
            #[cfg(feature = "pptx")]
            if matches!(verb, Verb::Set | Verb::Edit | Verb::Batch)
                && self.adapter() == Some(Adapter::Pptx)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "xlsb")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Xlsb)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "xls")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Xls)
            {
                // The canonical XLS Package always retains its editable
                // owner.  Keep the check for contextual reports defensive so
                // a future non-retained owner cannot overclaim writes.
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "ppt")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Ppt)
            {
                if self.retains_source == Some(false) {
                    return RETAINED_SOURCE.support();
                }
                return Support::Supported;
            }
            #[cfg(feature = "rtf")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Rtf)
            {
                return Support::Supported;
            }
            #[cfg(feature = "ods")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Ods)
            {
                return Support::Supported;
            }
            #[cfg(feature = "odt")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Odt)
            {
                return Support::Supported;
            }
            #[cfg(feature = "odp")]
            if matches!(
                verb,
                Verb::Set
                    | Verb::Add
                    | Verb::Remove
                    | Verb::Move
                    | Verb::Swap
                    | Verb::Edit
                    | Verb::Batch
            ) && self.adapter() == Some(Adapter::Odp)
            {
                return Support::Supported;
            }
            #[cfg(odf_objects)]
            // Every edit verb dispatches to the object format's leaf; what
            // a format cannot do is an element-level gap (record 0292).
            if self
                .adapter()
                .and_then(crate::format::odf_tree::writes)
                .is_some()
            {
                return Support::Supported;
            }
            #[cfg(feature = "iwork-adapter")]
            if let Some(writes) = self.adapter().and_then(crate::format::iwork::writes) {
                return match verb {
                    Verb::Set | Verb::Edit | Verb::Batch => Support::Supported,
                    Verb::Add if !writes.addable.is_empty() => Support::Supported,
                    Verb::Remove if !writes.removable.is_empty() => Support::Supported,
                    // Pages exchanges paragraphs (record 0299); Keynote slides,
                    // paragraphs and drawables (record 0298); Numbers sheets
                    // and tables (record 0347).
                    Verb::Move | Verb::Swap if !writes.movable.is_empty() => Support::Supported,
                    _ => IWORK_STRUCTURE.support(),
                };
            }
            return EDITS.support();
        }
        match verb {
            Verb::View => {
                if Mode::ALL.iter().any(|mode| self.view(*mode).is_supported()) {
                    Support::Supported
                } else {
                    // No adapter renders no mode; report the first refusal.
                    self.view(Mode::Text)
                }
            },
            Verb::Get
            | Verb::Find
            | Verb::Query
            | Verb::Set
            | Verb::Add
            | Verb::Remove
            | Verb::Move
            | Verb::Swap
            | Verb::Edit
            | Verb::Batch => Support::Supported,
            Verb::Save => self.publication_support(),
            Verb::Create => self.creation_support(),
            Verb::Import | Verb::Evaluate => self.m4_support(verb),
            Verb::Merge => self.merge_support(),
            Verb::Replace => self.replace_support(),
            Verb::Refresh => self.refresh_support(),
            Verb::Dump | Verb::Raw | Verb::Validate | Verb::AddPart | Verb::RawSet => {
                self.raw_support(verb).unwrap_or_else(|| EDITS.support())
            },
        }
    }

    /// Whether this report's format has an integrated publisher.
    fn publication_support(&self) -> Support {
        if self.format == Some(FileFormat::Rtf) && cfg!(feature = "rtf") {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Ods) && cfg!(feature = "ods") {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Odt) && cfg!(feature = "odt") {
            return Support::Supported;
        }
        if matches!(
            self.format,
            Some(
                FileFormat::Oth
                    | FileFormat::Odm
                    | FileFormat::Odg
                    | FileFormat::Odc
                    | FileFormat::Odi
                    | FileFormat::Odb
                    | FileFormat::Odf
            )
        ) && self.adapter().is_some()
        {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Odp) && cfg!(feature = "odp") {
            return Support::Supported;
        }
        if matches!(
            self.format,
            Some(FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers)
        ) && self.adapter().is_some()
        {
            return Support::Supported;
        }
        #[cfg(feature = "ppt")]
        if self.format == Some(FileFormat::Ppt) {
            if self.retains_source == Some(false) {
                return RETAINED_SOURCE.support();
            }
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Xls) && cfg!(feature = "xls") {
            return Support::Supported;
        }

        if matches!(
            self.format,
            Some(
                FileFormat::Doc
                    | FileFormat::Docx
                    | FileFormat::Xlsx
                    | FileFormat::Pptx
                    | FileFormat::Xlsb,
            )
        ) && self.adapter().is_some()
        {
            Support::Supported
        } else if let Some(refusal) = self.read_refusal() {
            refusal.support()
        } else {
            SAVE_PENDING.support()
        }
    }

    /// Whether this report's format has an integrated creator.
    fn creation_support(&self) -> Support {
        if self.format == Some(FileFormat::Rtf) && cfg!(feature = "rtf") {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Ods) && cfg!(feature = "ods") {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Odt) && cfg!(feature = "odt") {
            return Support::Supported;
        }
        if matches!(
            self.format,
            Some(
                FileFormat::Oth
                    | FileFormat::Odm
                    | FileFormat::Odg
                    | FileFormat::Odc
                    | FileFormat::Odi
                    | FileFormat::Odb
                    | FileFormat::Odf
            )
        ) && self.adapter().is_some()
        {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Odp) && cfg!(feature = "odp") {
            return Support::Supported;
        }
        // Pages (record 0299), Numbers (record 0297) and Keynote (record
        // 0298) create their minimal documents.
        if matches!(
            self.format,
            Some(FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers)
        ) && self.adapter().is_some()
        {
            return Support::Supported;
        }

        #[cfg(feature = "ppt")]
        if self.format == Some(FileFormat::Ppt) {
            return Support::Supported;
        }
        if self.format == Some(FileFormat::Xls) && cfg!(feature = "xls") {
            return Support::Supported;
        }

        if matches!(
            self.format,
            Some(
                FileFormat::Doc
                    | FileFormat::Docx
                    | FileFormat::Xlsx
                    | FileFormat::Pptx
                    | FileFormat::Xlsb,
            )
        ) && self.adapter().is_some()
        {
            Support::Supported
        } else if let Some(refusal) = self.read_refusal() {
            refusal.support()
        } else {
            CREATE_PENDING.support()
        }
    }

    /// The raw-layer and dump verbs (record 0336), `None` for the others.
    fn raw_support(&self, verb: Verb) -> Option<Support> {
        const NO_PARTS: Refusal = Refusal {
            code: Code::UnsupportedCapability,
            follow_up: "record 0336: package parts exist in OOXML and packaged ODF only",
            reason: "this format has no package parts",
        };
        const ODF_WRITES: Refusal = Refusal {
            code: Code::UnsupportedCapability,
            follow_up: "record 0336 follow-up: raw ODF part writes through the ODF package publisher",
            reason: "raw part writes cover OOXML (DOCX, XLSX, PPTX, XLSB)",
        };
        const IWORK_VALIDATION: Refusal = Refusal {
            code: Code::UnsupportedCapability,
            follow_up: "record 0336 follow-up: an iWork package validator from the iWork owners",
            reason: "iWork packages are validated by their owners on open; no report-producing validator exists",
        };
        let format = self.format?;
        let ooxml = matches!(
            format,
            FileFormat::Docx | FileFormat::Xlsx | FileFormat::Pptx | FileFormat::Xlsb
        );
        let odf = matches!(
            format,
            FileFormat::Odt
                | FileFormat::Ods
                | FileFormat::Odp
                | FileFormat::Odg
                | FileFormat::Odc
                | FileFormat::Odi
                | FileFormat::Odb
                | FileFormat::Odf
                | FileFormat::Oth
                | FileFormat::Odm
        );
        let iwork = matches!(
            format,
            FileFormat::Pages | FileFormat::Keynote | FileFormat::Numbers
        );
        Some(match verb {
            Verb::Dump => crate::batch::dump_support(format),
            Verb::Raw if ooxml || odf => Support::Supported,
            Verb::Raw => NO_PARTS.support(),
            Verb::AddPart | Verb::RawSet if ooxml => Support::Supported,
            Verb::AddPart | Verb::RawSet if odf => ODF_WRITES.support(),
            Verb::AddPart | Verb::RawSet => NO_PARTS.support(),
            Verb::Validate if iwork => IWORK_VALIDATION.support(),
            Verb::Validate => Support::Supported,
            _ => return None,
        })
    }

    /// Whether the `mode` view is rendered, or the code it refuses with.
    #[must_use]
    pub fn view(&self, mode: Mode) -> Support {
        if let Some(refusal) = self.read_refusal() {
            return refusal.support();
        }
        match self.adapter().and_then(|adapter| view_gap(adapter, mode)) {
            Some(gap) => Support::from_gap(gap),
            None => Support::Supported,
        }
    }

    /// The element kinds the format's resolver knows, in its spec table's
    /// order (the document root first); none when the format has no adapter
    /// in this build. Listed for a locked file too, whose reads refuse.
    pub fn elements(&self) -> Elements {
        Elements {
            adapter: self.adapter(),
            retains_source: self.retains_source,
            pptx_slide_add: self.pptx_slide_add,
            pptx_shape_add: self.pptx_shape_add,
            pptx_slide_remove: self.pptx_slide_remove,
            pptx_shape_remove: self.pptx_shape_remove,
            next: 0,
        }
    }

    /// The element whose OfficeCLI type name or alias is `name`, ASCII
    /// case-insensitively (`cell`, `col`, `ListObject`).
    #[must_use]
    pub fn element(&self, name: &str) -> Option<Element> {
        self.elements().find(|element| {
            element.name().eq_ignore_ascii_case(name)
                || element
                    .names()
                    .iter()
                    .any(|alias| alias.eq_ignore_ascii_case(name))
        })
    }

    /// Every element, axis, property, operation context and view mode of the format that
    /// Litchi cannot answer yet, with its code and follow-up: the elements
    /// with a gap first, then the format's axis and property gaps, then the
    /// view modes the adapter does not render.
    pub fn gaps(&self) -> Gaps {
        let independent: &'static [GapSpec] = match self.format {
            #[cfg(feature = "ppt")]
            Some(FileFormat::Ppt) => &PPT_CREATION_OPTION_GAPS,
            #[cfg(feature = "rtf")]
            Some(FileFormat::Rtf) => &RTF_CREATION_OPTION_GAPS,
            _ => &[],
        };
        Gaps {
            independent: independent.iter(),
            adapter: self.adapter(),
            phase: GapPhase::Elements,
            next: 0,
        }
    }
}

/// The coverage of `format` in this build.
fn coverage(format: Option<FileFormat>) -> Coverage {
    let Some(format) = format else {
        return Coverage::Refused(UNKNOWN);
    };
    if let Some(adapter) = Adapter::of(format) {
        return Coverage::Adapter(adapter);
    }
    Coverage::Refused(refused(format))
}

/// Why the unified reads of `format` are refused as a whole, when its
/// adapter is not in this build: the M1 step that adds it to an open `File`
/// variant, the cargo feature that is off, or the milestone that gives the
/// format a `File` variant.
fn refused(format: FileFormat) -> Refusal {
    let pending = |follow_up, reason| Refusal {
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    };
    let closed = |follow_up, reason| Refusal {
        code: Code::UnsupportedFormat,
        follow_up,
        reason,
    };
    let feature = |feature: &'static str| Refusal {
        code: Code::UnsupportedFormat,
        follow_up: feature,
        reason: "the format's cargo feature is off in this build",
    };
    match format {
        FileFormat::Docx if cfg!(feature = "docx") => pending(
            "M1 R4a",
            "the DOCX adapter of record 0010 §4.5 has not landed",
        ),
        FileFormat::Xlsb if cfg!(feature = "xlsb") => pending(
            "M1 R4d",
            "the XLSB adapter of record 0010 §4.7 has not landed",
        ),
        FileFormat::Doc if cfg!(feature = "doc") => pending(
            "M1 R6a",
            "the DOC adapter of record 0010 §4.8 has not landed",
        ),
        FileFormat::Odt if cfg!(feature = "odt") => pending(
            "M1 R7",
            "the ODT adapter of record 0010 §4.9 has not landed",
        ),
        FileFormat::Odp if cfg!(feature = "odp") => pending(
            "M1 R7",
            "the ODP adapter of record 0010 §4.9 has not landed",
        ),
        // A format whose adapter has landed but whose feature is off, or a
        // tier 1-3 format without its feature.
        FileFormat::Docx => feature("enable the `docx` cargo feature"),
        FileFormat::Xlsx => feature("enable the `xlsx` cargo feature"),
        FileFormat::Pptx => feature("enable the `pptx` cargo feature"),
        FileFormat::Xlsb => feature("enable the `xlsb` cargo feature"),
        FileFormat::Doc => feature("enable the `doc` cargo feature"),
        FileFormat::Xls => feature("enable the `xls` cargo feature"),
        FileFormat::Ppt => feature("enable the `ppt` cargo feature"),
        FileFormat::Rtf => feature("enable the `rtf` cargo feature"),
        FileFormat::Odt => feature("enable the `odt` cargo feature"),
        FileFormat::Ods => feature("enable the `ods` cargo feature"),
        FileFormat::Odp => feature("enable the `odp` cargo feature"),
        FileFormat::Oth => feature("enable the `oth` cargo feature"),
        FileFormat::Odm => feature("enable the `odm` cargo feature"),
        FileFormat::Odg => feature("enable the `odg` cargo feature"),
        FileFormat::Odc => feature("enable the `odc` cargo feature"),
        FileFormat::Odi => feature("enable the `odi` cargo feature"),
        FileFormat::Odb => feature("enable the `odb` cargo feature"),
        FileFormat::Odf => feature("enable the `odf-formula` cargo feature"),
        FileFormat::Pages => feature("enable the `pages` cargo feature"),
        FileFormat::Keynote => feature("enable the `keynote` cargo feature"),
        FileFormat::Numbers => feature("enable the `numbers` cargo feature"),
        _ => closed(
            "M6",
            "the format has no unified `File` variant before milestone M6",
        ),
    }
}

/// Whether a file retained what an editor is built from; `None` for the
/// read-only formats and a locked file.
fn retains_source(file: &File) -> Option<bool> {
    match file {
        #[cfg(feature = "docx")]
        File::Docx(handle) => Some(handle.retains_source()),
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => Some(handle.retains_source()),
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => Some(handle.retains_source()),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => Some(handle.retains_source()),
        #[cfg(feature = "doc")]
        File::Doc(handle) => Some(handle.retains_source()),
        #[cfg(feature = "xls")]
        File::Xls(handle) => Some(handle.retains_source()),
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => Some(handle.retains_source()),
        File::Signed(signed) => retains_source(signed.file()),
        _ => None,
    }
}

/// Whether a contextual PPTX owner can add a slide through the source-native
/// transaction.
fn pptx_slide_add(file: &File) -> Option<bool> {
    #[cfg(feature = "pptx")]
    {
        match file {
            File::Pptx(handle) => Some(handle.supports_slide_add()),
            File::Signed(signed) => pptx_slide_add(signed.file()),
            _ => None,
        }
    }
    #[cfg(not(feature = "pptx"))]
    {
        let _ = file;
        None
    }
}

/// Whether a contextual PPTX owner can add a text-box shape. Shape creation
/// still requires the eager owner, even though source owners can add slides.
fn pptx_shape_add(file: &File) -> Option<bool> {
    #[cfg(feature = "pptx")]
    {
        match file {
            File::Pptx(handle) => Some(handle.supports_shape_add()),
            File::Signed(signed) => pptx_shape_add(signed.file()),
            _ => None,
        }
    }
    #[cfg(not(feature = "pptx"))]
    {
        let _ = file;
        None
    }
}

/// Whether a contextual PPTX owner can remove a dependency-free slide. The
/// operation itself performs relationship and final-slide admission.
fn pptx_slide_remove(file: &File) -> Option<bool> {
    #[cfg(feature = "pptx")]
    {
        match file {
            File::Pptx(handle) => Some(handle.supports_slide_removal()),
            File::Signed(signed) => pptx_slide_remove(signed.file()),
            _ => None,
        }
    }
    #[cfg(not(feature = "pptx"))]
    {
        let _ = file;
        None
    }
}

/// Whether a contextual PPTX owner can remove a top-level slide shape. The
/// operation itself admits only supported shape topology and dependencies.
fn pptx_shape_remove(file: &File) -> Option<bool> {
    #[cfg(feature = "pptx")]
    {
        match file {
            File::Pptx(handle) => Some(handle.supports_shape_removal()),
            File::Signed(signed) => pptx_shape_remove(signed.file()),
            _ => None,
        }
    }
    #[cfg(not(feature = "pptx"))]
    {
        let _ = file;
        None
    }
}

/// The view modes no M1 adapter renders, with their follow-ups (record 0010
/// §6.1, M1-Δ20, design §10).
static ISSUES_GAP: GapSpec = view_gap_spec(
    "issues",
    "M2 (validate)",
    "the issues view needs the validators M2's `validate` adds (M1-Δ20)",
);
static FORMS_GAP: GapSpec = view_gap_spec(
    "forms",
    "M2 (validate)",
    "the forms view needs the validators M2's `validate` adds (M1-Δ20)",
);
static HTML_GAP: GapSpec = view_gap_spec(
    "html",
    "design §10 (rendering is out of scope)",
    "visual rendering is outside the unified layer",
);
static SVG_GAP: GapSpec = view_gap_spec(
    "svg",
    "design §10 (rendering is out of scope)",
    "visual rendering is outside the unified layer",
);
static SCREENSHOT_GAP: GapSpec = view_gap_spec(
    "screenshot",
    "design §10 (rendering is out of scope)",
    "visual rendering is outside the unified layer",
);
static PDF_GAP: GapSpec = view_gap_spec(
    "pdf",
    "design §10 (rendering is out of scope)",
    "visual rendering is outside the unified layer",
);

/// DOC creator options that the canonical ordinary-document builder cannot
/// honor. They remain visible in the report so callers can discover the
/// accepted flavor, conformance and metadata policy before calling create.
#[cfg(feature = "doc")]
static DOC_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "template",
        },
        code: Code::UnsupportedCapability,
        follow_up: "CreateOptions::flavor(Flavor::Document)",
        reason: "the DOC creator emits ordinary documents and refuses template flavor rather than silently changing the requested content type",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "strict-conformance",
        },
        code: Code::UnsupportedCapability,
        follow_up: "CreateOptions::conformance(Conformance::Transitional)",
        reason: "legacy DOC has one canonical binary writer and does not accept an OOXML Strict conformance request",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "unrepresentable-initial-metadata",
        },
        code: Code::UnsupportedCapability,
        follow_up: "CreateOptions::metadata (OLE property sets, record 0254)",
        reason: "initial metadata is written to the OLE property sets; identifier, hyperlink_base and unzoned timestamps have no property-set slot and are refused rather than dropped, as is metadata combined with CryptoAPI encryption",
    },
];

/// RTF creator options refused by the minimal native builder.
#[cfg(feature = "rtf")]
static RTF_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "template",
        },
        code: Code::UnsupportedCapability,
        follow_up: "use Flavor::Document",
        reason: "RTF template creation is not supported",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "strict-conformance",
        },
        code: Code::UnsupportedCapability,
        follow_up: "use Conformance::Transitional",
        reason: "RTF does not support OOXML Strict conformance",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "initial-metadata",
        },
        code: Code::UnsupportedCapability,
        follow_up: "RTF typed metadata creation",
        reason: "initial RTF metadata authoring is not integrated",
    },
];

/// PPT creator options that the native minimal builder intentionally refuses.
/// These remain visible even though the legacy PPT read adapter is still a
/// separate follow-up, because creation and publication are already exposed
/// by the canonical package owner.
#[cfg(feature = "ppt")]
static PPT_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    ppt_creation_gap(
        "template",
        "use Flavor::Document; M2 PPT template creation",
        "the PPT creator currently authors ordinary presentations and refuses template requests",
    ),
    ppt_creation_gap(
        "strict-conformance",
        "CreateOptions::conformance(Conformance::Transitional)",
        "legacy PPT has one canonical binary writer and refuses an OOXML Strict conformance request",
    ),
    ppt_creation_gap(
        "unrepresentable-initial-metadata",
        "CreateOptions::metadata (OLE property sets, record 0254)",
        "initial metadata is written to the OLE property sets; identifier, hyperlink_base and unzoned timestamps have no property-set slot and are refused rather than dropped",
    ),
];

#[cfg(feature = "ppt")]
const fn ppt_creation_gap(
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "presentation",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

#[cfg(feature = "ppt")]
const _: () = {
    let mut index = 0;
    while index < PPT_CREATION_OPTION_GAPS.len() {
        assert!(crate::node::spec::gap_valid(
            &PPT_CREATION_OPTION_GAPS[index]
        ));
        index += 1;
    }
};

/// DOC structural and batch contexts outside the deliberately narrow body
/// paragraph text setter. `batch` remains supported for atomic text sets;
/// structural commands are listed here so the report does not overclaim it.
#[cfg(feature = "doc")]
static DOC_WRITE_GAPS: [GapSpec; 6] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "metadata",
            context: "path-read",
        },
        code: Code::UnsupportedCapability,
        follow_up: "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        reason: "`set(\"/metadata\", ..)` writes the OLE document properties (root-metadata keys and custom.<name>); the path is set-only, and reads go through File::metadata()",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "body",
            context: "structural-edit",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 DOC structural body editing",
        reason: "the DOC write adapter currently replaces existing ordinary paragraph text and does not add elements",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "body",
            context: "structural-edit",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 DOC structural body editing",
        reason: "the DOC write adapter currently replaces existing ordinary paragraph text and does not remove elements",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "body",
            context: "structural-edit",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 DOC structural body editing",
        reason: "the DOC write adapter currently replaces existing ordinary paragraph text and does not move elements",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "body",
            context: "structural-edit",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 DOC structural body editing",
        reason: "the DOC write adapter currently replaces existing ordinary paragraph text and does not reorder elements",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "batch",
            element: "body",
            context: "non-text-command",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 DOC structural batch commands",
        reason: "DOC atomic batches currently admit paragraph text sets only; structural and read commands remain refused",
    },
];

/// DOCX creator options that are intentionally refused by the leaf creator.
/// They are format-level operation gaps rather than element properties: the
/// default creator is available, while these option classes fail closed with
/// a typed refusal.
#[cfg(feature = "docx")]
static DOCX_CREATION_OPTION_GAPS: [GapSpec; 2] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "strict-conformance",
        },
        code: Code::UnsupportedCapability,
        follow_up: "CreateOptions::conformance(Conformance::Transitional)",
        reason: "the DOCX creator emits Transitional packages and refuses a Strict conformance request rather than silently downgrading it",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "document",
            context: "noncore-initial-metadata",
        },
        code: Code::UnsupportedCapability,
        follow_up: "CreateOptions::metadata (DOCX core properties)",
        reason: "initial metadata outside the DOCX core-properties part is not representable by the creator and is refused rather than dropped",
    },
];

/// XLSX creation options not represented by its native minimal creator.
#[cfg(feature = "xlsx")]
static XLSX_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    xlsx_creation_gap(
        "template",
        "use Flavor::Document; M2 XLSX template creation",
        "the XLSX creator currently authors ordinary workbooks and refuses template requests",
    ),
    xlsx_creation_gap(
        "strict-conformance",
        "CreateOptions::conformance(Conformance::Transitional)",
        "the XLSX creator emits Transitional packages and refuses Strict requests",
    ),
    xlsx_creation_gap(
        "unrepresentable-initial-metadata",
        "CreateOptions::metadata (OOXML document properties, record 0253)",
        "initial metadata is committed through the leaf metadata transaction; content_type, last_backup and codepage have no OOXML property part and are refused rather than dropped",
    ),
];

/// XLS creation options not represented by its native minimal creator.
#[cfg(feature = "xls")]
static XLS_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    xls_creation_gap(
        "template",
        "use Flavor::Document; M2 XLS template creation",
        "the XLS creator currently authors ordinary workbooks and refuses template requests",
    ),
    xls_creation_gap(
        "strict-conformance",
        "CreateOptions::conformance(Conformance::Transitional)",
        "the XLS creator emits binary BIFF8 workbooks and refuses Strict requests",
    ),
    xls_creation_gap(
        "unrepresentable-initial-metadata",
        "CreateOptions::metadata (OLE property sets, record 0254)",
        "initial metadata is written to the OLE property sets; identifier, hyperlink_base and unzoned timestamps have no property-set slot and are refused rather than dropped",
    ),
];

/// XLSB creation options not represented by its native minimal creator.
#[cfg(feature = "xlsb")]
static XLSB_CREATION_OPTION_GAPS: [GapSpec; 3] = [
    xlsb_creation_gap(
        "template",
        "use Flavor::Document; M2 XLSB template creation",
        "the XLSB creator currently authors ordinary workbooks and refuses template requests",
    ),
    xlsb_creation_gap(
        "strict-conformance",
        "CreateOptions::conformance(Conformance::Transitional)",
        "the XLSB creator emits Transitional packages and refuses Strict requests",
    ),
    xlsb_creation_gap(
        "unrepresentable-initial-metadata",
        "CreateOptions::metadata (OOXML document properties, record 0253)",
        "initial metadata is committed through the leaf metadata transaction; content_type, last_backup and codepage have no OOXML property part and are refused rather than dropped",
    ),
];

/// XLSB writes keep a narrow element surface. A range is a selector accepted
/// by the cell handler, while workbook, row, column and other nodes are not
/// structural add/remove elements. Complete-catalog sheet insertion, removal,
/// and ordering dependency refusals remain explicit operation gaps below.
#[cfg(feature = "xlsb")]
static XLSB_WRITE_GAPS: [GapSpec; 25] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "metadata",
            context: "path-read",
        },
        code: Code::UnsupportedCapability,
        follow_up: "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        reason: "`set(\"/metadata\", ..)` writes the OOXML document properties (root-metadata keys and custom.<name>, record 0253); the path is set-only, and reads go through File::metadata() and the root node",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "workbook",
            context: "non-sheet-elements",
        },
        code: Code::UnsupportedElement,
        follow_up: "M2 XLSB structural add handlers for non-sheet elements",
        reason: "XLSB structural addition creates complete workbook tabs; cells, ranges, rows, columns, tables and other nodes require their own native lifecycle owners",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "order-opaque-dependencies",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet insertion",
        reason: "sheet insertion refuses shared-workbook revisions, data-model parts, connections, slicers, timelines, unmodeled workbook records, opaque volatile dependencies, unmodeled title layouts and markup-compatibility title metadata until their dependencies have a native remapping contract",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "vba-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet insertion",
        reason: "sheet insertion refuses workbooks with a VBA project dependency until native catalog identity remapping is available",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "calculation-chain-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB calculation-chain-aware sheet insertion",
        reason: "sheet insertion refuses workbooks with a calculation-chain dependency until native sheet positions can be remapped safely",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "noncontiguous-3d-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB 3-D dependency-aware sheet insertion",
        reason: "sheet insertion refuses a position inside an internal 3-D reference range because it would broaden the range",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "workbook",
            context: "non-sheet-elements",
        },
        code: Code::UnsupportedElement,
        follow_up: "M2 XLSB non-cell removal handlers",
        reason: "XLSB removal deletes complete-catalog sheet tabs and stored cell records; workbook, row, column, table and other non-sheet structural elements require their own native lifecycle owners",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "sheet",
            context: "order-opaque-dependencies",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet removal",
        reason: "sheet removal admits only the selected sheet root and exclusively owned binaryIndex parts; any other outgoing internal relationship, including chart, drawing, comments, tables or adjunct parts, is refused as unsupported until it has a typed dependency disposition, along with shared-workbook revisions, data-model parts, connections, slicers, timelines, pivot-cache and table formula dependencies, unmodeled workbook records, surviving worksheet FRT/AC records and formula groups, opaque volatile dependencies and markup-compatibility title metadata",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "sheet",
            context: "removed-scoped-names",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware scoped-name removal",
        reason: "sheet removal refuses workbook or sheet-scoped names that resolve into a removed tab until native name pruning and reference remapping are available",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "sheet",
            context: "textual-references",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware textual-reference removal",
        reason: "sheet removal refuses formula-like and textual dependency carriers that still reference a removed tab instead of leaving stale names or links",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "cell",
            context: "formula-groups",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware formula-group removal",
        reason: "array, shared and data-table formula-group members require native group lifecycle support before removal",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "cell",
            context: "calculation-chain",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB calculation-chain-aware formula removal",
        reason: "stored formulas cannot be removed while an opaque calculation chain is present",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "cell",
            context: "unmodeled-cell-records",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB additional native cell-record families",
        reason: "semantic removal refuses worksheets containing reserved cell-record families whose coordinate layout is not modeled",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "sheet",
            context: "unbound-tab-visibility",
        },
        code: Code::UnsupportedCapability,
        follow_up: "bind a physical sheet before revealing an unbound tab",
        reason: "an unbound XLSB tab has no sheet payload and must remain veryHidden",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "sheet",
            context: "unmodeled-tab-views",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB multi-window and additional tab-kind view mutation",
        reason: "visibility changes refuse active-tab relocation across multiple workbook windows and selected-view changes on unmodeled dialog or macro tabs",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "sheet",
            context: "sheet-name-markup-compatibility",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB markup-compatibility-aware sheet-name dependency rewrites",
        reason: "sheet-name edits refuse a recognized dependency inside an OOXML markup-compatibility alternative instead of selecting or discarding a branch",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "sheet",
            context: "sheet-name-unmodeled-reference",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB additional sheet-name reference carriers",
        reason: "sheet-name edits refuse formula-like or name-like reference carriers outside the proven rewrite vocabulary so dependent names cannot be left stale",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "sheet",
            context: "order-opaque-dependencies",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet ordering",
        reason: "sheet ordering refuses shared-workbook revisions, data-model parts, connections, slicers, timelines, unmodeled workbook records, opaque volatile dependencies, unmodeled title layouts and markup-compatibility title metadata until their historical or ordinal dependencies have a native remapping contract",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "sheet",
            context: "order-opaque-dependencies",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet ordering",
        reason: "sheet ordering refuses shared-workbook revisions, data-model parts, connections, slicers, timelines, unmodeled workbook records, opaque volatile dependencies, unmodeled title layouts and markup-compatibility title metadata until their historical or ordinal dependencies have a native remapping contract",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "sheet",
            context: "vba-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet ordering",
        reason: "sheet ordering refuses workbooks with a VBA project dependency until native catalog identity remapping is available",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "sheet",
            context: "vba-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB dependency-aware complete-catalog sheet ordering",
        reason: "sheet ordering refuses workbooks with a VBA project dependency until native catalog identity remapping is available",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "sheet",
            context: "calculation-chain-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB calculation-chain-aware sheet ordering",
        reason: "sheet ordering refuses workbooks with a calculation-chain dependency until native sheet positions can be remapped safely",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "sheet",
            context: "calculation-chain-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB calculation-chain-aware sheet ordering",
        reason: "sheet ordering refuses workbooks with a calculation-chain dependency until native sheet positions can be remapped safely",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "sheet",
            context: "noncontiguous-3d-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB 3-D dependency-aware sheet ordering",
        reason: "sheet ordering refuses a permutation that would make an internal 3-D reference noncontiguous",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "sheet",
            context: "noncontiguous-3d-dependency",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSB 3-D dependency-aware sheet ordering",
        reason: "sheet ordering refuses a permutation that would make an internal 3-D reference noncontiguous",
    },
];

#[cfg(feature = "xlsx")]
static XLSX_MANAGED_PUBLICATION_GAP: GapSpec = GapSpec {
    subject: Subject::Operation {
        verb: "save",
        element: "workbook",
        context: "execution-context-budgeting",
    },
    code: Code::UnsupportedCapability,
    follow_up: "managed OOXML open and publication (record 0008 A1)",
    reason: "OpenOptions::execution checks open-boundary cancellation only; unified opens do not yet retain a managed owner for publication budgeting",
};

#[cfg(all(feature = "xlsx", feature = "encryption"))]
static XLSX_SECURITY_OPTION_GAPS: [GapSpec; 4] = [
    encrypted_xlsx_edit_gap("set"),
    encrypted_xlsx_edit_gap("edit"),
    encrypted_xlsx_edit_gap("batch"),
    GapSpec {
        subject: Subject::Operation {
            verb: "save",
            element: "workbook",
            context: "encrypted-without-reencrypt",
        },
        code: Code::DocumentProtected,
        follow_up: "SaveOptions::reencrypt",
        reason: "an encrypted XLSX requires explicit re-encryption for publication; ordinary output never silently emits plaintext",
    },
];

#[cfg(all(feature = "xlsx", feature = "encryption"))]
const fn encrypted_xlsx_edit_gap(verb: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element: "workbook",
            context: "unlocked-encrypted-owner",
        },
        code: Code::DocumentProtected,
        follow_up: "M2 (XLSX encryption-profile-preserving native edits)",
        reason: "the XLSX leaf refuses edits of decrypted packages because its current edit publication cannot preserve the encryption profile",
    }
}

#[cfg(feature = "xlsx")]
const fn xlsx_creation_gap(
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "workbook",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

#[cfg(feature = "xls")]
const fn xls_creation_gap(
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "workbook",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

#[cfg(feature = "xlsb")]
const fn xlsb_creation_gap(
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "workbook",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

#[cfg(feature = "xlsx")]
static XLSX_WRITE_GAPS: [GapSpec; 16] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "metadata",
            context: "path-read",
        },
        code: Code::UnsupportedCapability,
        follow_up: "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        reason: "`set(\"/metadata\", ..)` writes the OOXML document properties (root-metadata keys and custom.<name>, record 0253); the path is set-only, and reads go through File::metadata() and the root node",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "range",
            context: "merge-native-preservation-guards",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX dependency-aware and compatibility-aware merge edits",
        reason: "native merge edits refuse follower content, overlapping geometry, table intersections, group formulas and unmodeled or compatibility-owned payloads instead of discarding them",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "column",
            context: "unit-qualified-width-input",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX OfficeCLI column width input conversion",
        reason: "column width accepts native character-unit numbers; OfficeCLI unit-qualified width conversion remains pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "workbook",
            context: "row-column-table-targets",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX row/column payload and table write admission",
        reason: "row and column layout properties support scoped selectors; whole-row/column cell payloads, styles, table-driven keys and table writes remain pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "batch",
            element: "workbook",
            context: "structural-and-coalesced-commands",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX structure and native transaction coalescing",
        reason: "worksheet lifecycle commands and distinct direct cell payload coalescing are integrated; other structural targets, broader coalescing and read commands remain pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "officecli-input-conventions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX OfficeCLI input policy",
        reason: "cell values preserve typed scalar input; formulas use native formula text without a leading equals sign; OfficeCLI string inference and extended clear modes remain pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "number-format-input-repair",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX explicit number-format compatibility input policy",
        reason: "numberformat accepts checked native codes and OfficeCLI keyword aliases; malformed quotes are rejected under strict defaults rather than repaired by dropping trailing characters",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "fill-replacement-opaque-content",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX lossless replacement of extended fill children",
        reason: "typed solid, pattern and gradient fills support whole-fill replacement; a changed replacement refuses opaque content inside the replaced pattern or gradient child, while exact no-ops retain it",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "border-color-without-line",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX explicit invisible-border color authoring policy",
        reason: "border colors require a resulting visible line; strict input admission rejects colors OfficeCLI silently discards on an unstyled side; clearing a line retains existing source color markup",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "border-opaque-content",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX lossless replacement of extended border content",
        reason: "granular border edits preserve unrelated markup; removing a side or color refuses opaque content that would be discarded",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "cell",
            context: "border-color-projection",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M1 XLSX complete scalar border color projection",
        reason: "ordinary border color readback follows OfficeCLI RGB projection; native theme, indexed, automatic colors and tints remain available through Node::cell_format",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "border-native-only-edges",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX logical and interior border property vocabulary",
        reason: "the native border editor supports Strict start/end edges, interior edges and outline; the unified OfficeCLI vocabulary currently exposes physical left/right, top/bottom and diagonal edges only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "style-resource-compaction",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX generated style resource lifecycle",
        reason: "cell formatting edits reuse matching resources and retain existing stylesheet records; reverting formatting across separate commits can retain unused generated records",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "font-authoring-profile",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX stylesheet creation, markup compatibility and inherited formatting",
        reason: "supported cell formatting edits preserve proven Ignorable markup and require existing direct font and cell-format tables; alternate content, unproven directives, escaped namespace bindings and inherited formatting remain unsupported",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "sheet",
            context: "non-worksheet-tabs",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX non-worksheet tab removal",
        reason: "worksheet removal uses native dependency guards; removal of chart, dialog and macro tabs remains pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "officecli-sheet-input-conventions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 XLSX sheet creation property and replay conventions",
        reason: "sheet creation supports name, hidden and explicit placement, defaulting name to Sheet plus catalog count plus one; ifExists, placeholder claiming and additional sheet properties remain pending",
    },
];

/// Remaining command-channel contexts are discoverable even though the core
/// five mutations and current-version metadata are executable.
#[cfg(feature = "docx")]
static DOCX_BATCH_GAPS: [GapSpec; 6] = [
    batch_gap(
        "text-paragraph-and-soft-breaks",
        "full text-output parity remains incomplete; paragraph addition splits normalized LF into paragraphs with VT inline breaks; run setters update the first literal text node and mark the owning same-paragraph cached field dirty; run addition and paragraph setters normalize controls to inline breaks",
    ),
    batch_gap(
        "read-and-extended-commands",
        "get, query, view, raw, raw-set, add-part, validate and import are not integrated into batch item results",
    ),
    batch_gap(
        "non-v2-dump-metadata",
        "dumpVersion other than explicit v2 is refused for the whole batch; absent metadata uses current semantics",
    ),
    batch_gap(
        "add-from",
        "copying an add source requires typed transfer families not yet integrated into the command channel",
    ),
    batch_gap(
        "move-properties",
        "move-time property assignments are not integrated and are never silently dropped",
    ),
    batch_gap(
        "move-parent-and-relative-anchor",
        "move to a parent combined with before or after is not represented by the current placement model",
    ),
];
#[cfg(feature = "docx")]
const fn batch_gap(context: &'static str, reason: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "batch",
            element: "document",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 (batch command and transfer families)",
        reason,
    }
}

/// Decrypted DOCX owners currently refuse native editing. Publication requires
/// explicit re-encryption to preserve protection. These gaps only exist in
/// builds that expose encryption.
#[cfg(all(feature = "docx", feature = "encryption"))]
static DOCX_SECURITY_OPTION_GAPS: [GapSpec; 8] = [
    encrypted_docx_edit_gap("set"),
    encrypted_docx_edit_gap("add"),
    encrypted_docx_edit_gap("remove"),
    encrypted_docx_edit_gap("move"),
    encrypted_docx_edit_gap("swap"),
    encrypted_docx_edit_gap("edit"),
    encrypted_docx_edit_gap("batch"),
    GapSpec {
        subject: Subject::Operation {
            verb: "save",
            element: "document",
            context: "encrypted-without-reencrypt",
        },
        code: Code::DocumentProtected,
        follow_up: "SaveOptions::reencrypt",
        reason: "an encrypted DOCX cannot be published as plain bytes because that would silently downgrade its protection profile",
    },
];

#[cfg(all(feature = "docx", feature = "encryption"))]
const fn encrypted_docx_edit_gap(verb: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element: "document",
            context: "unlocked-encrypted-owner",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 (DOCX encryption-profile-preserving native edits)",
        reason: "the DOCX leaf refuses edits of decrypted packages because its current edit publication cannot preserve the encryption profile",
    }
}

#[cfg(feature = "doc")]
const _: () = {
    let mut index = 0;
    while index < DOC_CREATION_OPTION_GAPS.len() {
        assert!(crate::node::spec::gap_valid(
            &DOC_CREATION_OPTION_GAPS[index]
        ));
        index += 1;
    }
    let mut index = 0;
    while index < DOC_WRITE_GAPS.len() {
        assert!(crate::node::spec::gap_valid(&DOC_WRITE_GAPS[index]));
        index += 1;
    }
};

#[cfg(feature = "docx")]
const _: () = {
    let mut index = 0;
    while index < DOCX_CREATION_OPTION_GAPS.len() {
        assert!(crate::node::spec::gap_valid(
            &DOCX_CREATION_OPTION_GAPS[index]
        ));
        index += 1;
    }
    let mut index = 0;
    while index < DOCX_BATCH_GAPS.len() {
        assert!(crate::node::spec::gap_valid(&DOCX_BATCH_GAPS[index]));
        index += 1;
    }
    #[cfg(feature = "encryption")]
    {
        let mut index = 0;
        while index < DOCX_SECURITY_OPTION_GAPS.len() {
            assert!(crate::node::spec::gap_valid(
                &DOCX_SECURITY_OPTION_GAPS[index]
            ));
            index += 1;
        }
    }
};

const fn view_gap_spec(
    mode: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::View(mode),
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

const _: () = {
    let gaps = [
        &ISSUES_GAP,
        &FORMS_GAP,
        &HTML_GAP,
        &SVG_GAP,
        &SCREENSHOT_GAP,
        &PDF_GAP,
    ];
    let mut index = 0;
    while index < gaps.len() {
        assert!(
            crate::node::spec::gap_valid(gaps[index]),
            "a view gap lacks a gap code, follow-up or reason"
        );
        index += 1;
    }
};

/// The gap of `mode` for `adapter`: every M1 adapter renders the text,
/// annotated, outline and stats views and refuses the others (record 0010
/// §6.1). The liveness tests check each adapter's views against it.
const fn view_gap(adapter: Adapter, mode: Mode) -> Option<&'static GapSpec> {
    match adapter {
        #[cfg(feature = "doc")]
        Adapter::Doc => m1_view_gap(mode),
        #[cfg(feature = "docx")]
        Adapter::Docx => m1_view_gap(mode),
        #[cfg(feature = "xlsx")]
        Adapter::Xlsx => m1_view_gap(mode),
        #[cfg(feature = "xlsb")]
        Adapter::Xlsb => m1_view_gap(mode),
        #[cfg(feature = "xls")]
        Adapter::Xls => m1_view_gap(mode),
        #[cfg(feature = "pptx")]
        Adapter::Pptx => m1_view_gap(mode),
        #[cfg(feature = "ppt")]
        Adapter::Ppt => m1_view_gap(mode),
        #[cfg(feature = "rtf")]
        Adapter::Rtf => m1_view_gap(mode),
        #[cfg(feature = "ods")]
        Adapter::Ods => m1_view_gap(mode),
        #[cfg(feature = "odt")]
        Adapter::Odt => m1_view_gap(mode),
        #[cfg(feature = "odp")]
        Adapter::Odp => m1_view_gap(mode),
        #[cfg(feature = "odg")]
        Adapter::Odg => m1_view_gap(mode),
        #[cfg(feature = "odc")]
        Adapter::Odc => m1_view_gap(mode),
        #[cfg(feature = "odi")]
        Adapter::Odi => m1_view_gap(mode),
        #[cfg(feature = "odb")]
        Adapter::Odb => m1_view_gap(mode),
        #[cfg(feature = "odf-formula")]
        Adapter::OdfFormula => m1_view_gap(mode),
        #[cfg(feature = "pages")]
        Adapter::Pages => m1_view_gap(mode),
        #[cfg(feature = "keynote")]
        Adapter::Keynote => m1_view_gap(mode),
        #[cfg(feature = "numbers")]
        Adapter::Numbers => m1_view_gap(mode),
    }
}

/// The M1 rule shared by the adapters.
const fn m1_view_gap(mode: Mode) -> Option<&'static GapSpec> {
    match mode {
        Mode::Text | Mode::Annotated | Mode::Outline | Mode::Stats => None,
        Mode::Issues => Some(&ISSUES_GAP),
        Mode::Forms => Some(&FORMS_GAP),
        Mode::Html => Some(&HTML_GAP),
        Mode::Svg => Some(&SVG_GAP),
        Mode::Screenshot => Some(&SCREENSHOT_GAP),
        Mode::Pdf => Some(&PDF_GAP),
    }
}

/// Checks the parts of a format's spec table the capability report reads,
/// at compile time; one invocation per adapter, below.
///
/// - every element has a path template, and every template is a rooted path
///   (`/...`) or Excel notation (`Sheet!A1`);
/// - every element with children has a rooted template for each child, from
///   which the child's path segment is taken;
/// - identity keys are non-empty and unique per element;
/// - an element gap names its own element, and a property gap names an
///   element of the format whose readable properties do not include the
///   key.
macro_rules! check_capabilities {
    ($spec:path) => {
        const _: () = {
            let spec = &$spec;
            let elements = spec.elements;
            let mut i = 0;
            while i < elements.len() {
                let element = &elements[i];
                assert!(!element.paths.is_empty(), "an element has no path template");
                let mut p = 0;
                while p < element.paths.len() {
                    assert!(
                        template_valid(element.paths[p]),
                        "a path template is neither rooted nor Excel notation"
                    );
                    p += 1;
                }
                if !element.children.is_empty() {
                    let mut c = 0;
                    while c < element.children.len() {
                        let child = &elements[element.children[c] as usize];
                        assert!(
                            has_rooted(child.paths),
                            "a child element has no rooted path template"
                        );
                        c += 1;
                    }
                }
                let mut k = 0;
                while k < element.identities.len() {
                    assert!(
                        !element.identities[k].is_empty(),
                        "an identity key is empty"
                    );
                    let mut other = 0;
                    while other < k {
                        assert!(
                            !same(element.identities[k], element.identities[other]),
                            "an identity key repeats"
                        );
                        other += 1;
                    }
                    k += 1;
                }
                if let Some(gap) = element.gap {
                    match gap.subject {
                        $crate::node::spec::Subject::Element(name) => {
                            assert!(
                                same(name, element.name),
                                "an element gap names another element"
                            )
                        },
                        _ => panic!("an element gap is not about an element"),
                    }
                }
                i += 1;
            }
            let mut g = 0;
            while g < spec.gaps.len() {
                if let $crate::node::spec::Subject::Property { element, key } = spec.gaps[g].subject
                {
                    let mut found = false;
                    let mut e = 0;
                    while e < elements.len() {
                        if same(elements[e].name, element) {
                            found = true;
                            let mut p = 0;
                            while p < elements[e].props.len() {
                                assert!(
                                    !same(elements[e].props[p].name, key),
                                    "a property is both readable and a gap"
                                );
                                p += 1;
                            }
                        }
                        e += 1;
                    }
                    assert!(found, "a property gap names no element of the format");
                }
                g += 1;
            }
        };
    };
}

#[cfg(feature = "doc")]
check_capabilities!(crate::format::doc::SPEC);

#[cfg(feature = "docx")]
check_capabilities!(crate::format::docx::SPEC);

/// The DOCX part elements `set`, `add` and `remove` write (record 0319).
#[cfg(feature = "docx")]
const DOCX_PART_WRITES: &[&str] = &[
    "style",
    "abstractNum",
    "num",
    "level",
    "header",
    "footer",
    "footnote",
    "endnote",
];

#[cfg(feature = "xlsx")]
check_capabilities!(crate::format::xlsx::SPEC);

#[cfg(feature = "xlsb")]
check_capabilities!(crate::format::xlsb::SPEC);

#[cfg(feature = "xls")]
check_capabilities!(crate::format::xls::SPEC);

#[cfg(feature = "pptx")]
check_capabilities!(crate::format::pptx::SPEC);

#[cfg(feature = "ppt")]
check_capabilities!(crate::format::ppt::SPEC);

#[cfg(feature = "rtf")]
check_capabilities!(crate::format::rtf::SPEC);

#[cfg(feature = "ods")]
check_capabilities!(crate::format::ods::SPEC);
#[cfg(feature = "odt")]
check_capabilities!(crate::format::odt::SPEC);
#[cfg(feature = "odp")]
check_capabilities!(crate::format::odp::SPEC);
#[cfg(feature = "odg")]
check_capabilities!(crate::format::odf_tree::odg::SPEC);
#[cfg(feature = "odc")]
check_capabilities!(crate::format::odf_tree::odc::SPEC);
#[cfg(feature = "odi")]
check_capabilities!(crate::format::odf_tree::odi::SPEC);
#[cfg(feature = "odb")]
check_capabilities!(crate::format::odf_tree::odb::SPEC);
#[cfg(feature = "odf-formula")]
check_capabilities!(crate::format::odf_tree::formula::SPEC);

#[cfg(feature = "pages")]
check_capabilities!(crate::format::iwork::pages::SPEC);

#[cfg(feature = "keynote")]
check_capabilities!(crate::format::iwork::keynote::SPEC);

#[cfg(feature = "numbers")]
check_capabilities!(crate::format::iwork::numbers::SPEC);

/// Byte equality usable in constant evaluation.
const fn same(a: &str, b: &str) -> bool {
    let (a, b) = (a.as_bytes(), b.as_bytes());
    if a.len() != b.len() {
        return false;
    }
    let mut index = 0;
    while index < a.len() {
        if a[index] != b[index] {
            return false;
        }
        index += 1;
    }
    true
}

/// Whether `template` is rooted (`/...`) or Excel notation (`Sheet!A1`).
const fn template_valid(template: &str) -> bool {
    let bytes = template.as_bytes();
    if bytes.is_empty() {
        return false;
    }
    if bytes[0] == b'/' {
        return true;
    }
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'!' {
            return true;
        }
        index += 1;
    }
    false
}

/// Whether some template of `paths` is rooted.
const fn has_rooted(paths: &[&str]) -> bool {
    let mut index = 0;
    while index < paths.len() {
        if !paths[index].is_empty() && paths[index].as_bytes()[0] == b'/' {
            return true;
        }
        index += 1;
    }
    false
}

/// Whether a path template selects by ordinal: some step's predicate is a
/// single upper-case placeholder (`[N]`, `[K]`).
fn positional(template: &str) -> bool {
    template
        .split('[')
        .skip(1)
        .filter_map(|rest| rest.split_once(']'))
        .any(|(inside, _)| {
            inside.len() == 1 && inside.bytes().all(|byte| byte.is_ascii_uppercase())
        })
}

/// The path segment a child's templates name it by: the last step of its
/// first rooted template, without predicates (`row[N]` gives `row`,
/// `<A1>` stays `<A1>`).
fn path_segment(paths: &[&'static str]) -> Option<&'static str> {
    let template = paths.iter().find(|template| template.starts_with('/'))?;
    let last = template.rsplit('/').next()?;
    let segment = last.split('[').next().unwrap_or(last);
    (!segment.is_empty()).then_some(segment)
}

/// The elements of a report, in spec-table order.
#[derive(Debug, Clone)]
pub struct Elements {
    adapter: Option<Adapter>,
    retains_source: Option<bool>,
    pptx_slide_add: Option<bool>,
    pptx_shape_add: Option<bool>,
    pptx_slide_remove: Option<bool>,
    pptx_shape_remove: Option<bool>,
    next: u16,
}

impl Iterator for Elements {
    type Item = Element;

    fn next(&mut self) -> Option<Element> {
        let adapter = self.adapter?;
        if usize::from(self.next) >= element_count(adapter) {
            return None;
        }
        let element = Element {
            adapter,
            retains_source: self.retains_source,
            pptx_slide_add: self.pptx_slide_add,
            pptx_shape_add: self.pptx_shape_add,
            pptx_slide_remove: self.pptx_slide_remove,
            pptx_shape_remove: self.pptx_shape_remove,
            index: self.next,
        };
        self.next += 1;
        Some(element)
    }
}

/// One element kind of a format, as its resolver knows it: the OfficeCLI
/// help schema's element entry.
///
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Element {
    adapter: Adapter,
    retains_source: Option<bool>,
    pptx_slide_add: Option<bool>,
    pptx_shape_add: Option<bool>,
    pptx_slide_remove: Option<bool>,
    pptx_shape_remove: Option<bool>,
    index: u16,
}

impl Element {
    /// The element's row of its format's spec table.
    fn data(&self) -> Data {
        element_data(self.adapter, self.index)
    }

    /// The OfficeCLI type name (`paragraph`, `cell`, `table`).
    #[must_use]
    pub fn name(&self) -> &'static str {
        self.data().name
    }

    /// Every spelling a path step or selector may use for the element,
    /// lowercase (the type name included).
    fn names(&self) -> &'static [&'static str] {
        self.data().aliases
    }

    /// The other names a path step or selector may use for the element
    /// (OfficeCLI's `elementAliases`): `col` for `column`, `listobject` for
    /// `table`.
    pub fn aliases(&self) -> impl Iterator<Item = &'static str> + Clone {
        let name = self.name();
        self.names()
            .iter()
            .copied()
            .filter(move |alias| !alias.eq_ignore_ascii_case(name))
    }

    /// The element's path templates (`/<sheet>/row[N]`), each with whether
    /// it is positional (selects by ordinal) rather than stable.
    pub fn paths(&self) -> impl Iterator<Item = (&'static str, bool)> + Clone {
        self.data()
            .paths
            .iter()
            .filter(|template| ordinary_template(template))
            .map(|template| (*template, positional(template)))
    }

    /// The elements that may contain this one.
    pub fn parents(&self) -> Related {
        Related {
            element: *self,
            parents: true,
            next: 0,
        }
    }

    /// The elements this one may contain, as the resolver and the query
    /// walker read them.
    pub fn children(&self) -> Related {
        Related {
            element: *self,
            parents: false,
            next: 0,
        }
    }

    /// The path segment a parent's path names this element by (`row`,
    /// `<A1>`, `<sheet>`): the last step of its first rooted template,
    /// without predicates.
    #[must_use]
    pub fn path_segment(&self) -> Option<&'static str> {
        self.paths().find_map(|(path, _)| path_segment(&[path]))
    }

    /// The declared properties, in the order the node view emits them;
    /// write-only entries remain visible with `get == false`.
    pub fn properties(&self) -> Properties {
        Properties {
            aliases: self.data().prop_aliases,
            props: self.data().props.iter(),
            element: *self,
            retains_source: self.retains_source,
        }
    }

    /// The property named `key` (exact, as OfficeCLI's keys are on output).
    #[must_use]
    pub fn property(&self, key: &str) -> Option<Property> {
        self.properties().find(|property| property.name() == key)
    }

    /// The gap when the format has the element but Litchi cannot read it
    /// yet: resolving it is the gap's code, `unsupported_capability`.
    #[must_use]
    pub fn gap(&self) -> Option<Gap> {
        self.data().gap.map(|spec| Gap { spec })
    }

    /// Format-level support; owner-state restrictions are reported by `Capabilities::verb`.
    /// In M1, `get` reads every element
    /// without a gap and `query` matches every such element the query walker
    /// reaches from the document root; M2's handler table answers `add`,
    /// `set` and `remove` here (record 0010 §7.3).
    #[must_use]
    pub fn operations(&self) -> Operations {
        let readable = self.gap().is_none();
        let settable = {
            #[cfg(feature = "docx")]
            {
                (self.adapter == Adapter::Docx
                    && self.retains_source != Some(false)
                    && matches!(
                        self.data().canonical,
                        "p" | "r"
                            | "hyperlink"
                            | "tc"
                            // Table borders, row height and flags (record 0342).
                            | "tbl"
                            | "tr"
                            | "chart"
                            | "series"
                            | "axis"
                            | "picture"
                            | "comment"
                            | "smartart"
                            | "node"
                            | "document"
                            // Table grid columns (record 0345).
                            | "col"
                    )
                    && readable)
                    || (self.adapter == Adapter::Docx
                        && self.retains_source != Some(false)
                        && (DOCX_PART_WRITES.contains(&self.data().canonical)
                            || matches!(
                                self.data().canonical,
                                "tab" | "bookmark" | "field" | "toc"
                            ))
                        && readable)
            }
            #[cfg(not(feature = "docx"))]
            {
                false
            }
        };
        #[cfg(feature = "doc")]
        let settable = settable
            || (self.adapter == Adapter::Doc
                && self.retains_source != Some(false)
                && crate::format::doc::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "xlsx")]
        let settable = settable
            || (self.adapter == Adapter::Xlsx
                && self.retains_source != Some(false)
                && (matches!(
                    self.data().canonical,
                    "workbook"
                        | "cell"
                        | "range"
                        | "sheet"
                        | "row"
                        | "col"
                        | "chart"
                        | "series"
                        | "axis"
                        | "run"
                ) || crate::format::xlsx::FEATURE_ELEMENTS.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "xlsb")]
        let settable = settable
            || (self.adapter == Adapter::Xlsb
                && self.retains_source != Some(false)
                && (matches!(
                    self.data().canonical,
                    "cell" | "range" | "sheet" | "namedrange"
                ) || crate::format::xlsb::FEATURE_ELEMENTS.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "xls")]
        let settable = settable
            || (self.adapter == Adapter::Xls
                && self.retains_source != Some(false)
                && crate::format::xls::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "pptx")]
        let settable = settable
            || (self.adapter == Adapter::Pptx
                && self.retains_source != Some(false)
                && crate::format::pptx::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ppt")]
        let settable = settable
            || (self.adapter == Adapter::Ppt
                && self.retains_source != Some(false)
                && crate::format::ppt::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "rtf")]
        let settable = settable
            || (self.adapter == Adapter::Rtf
                && crate::format::rtf::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ods")]
        let settable = settable
            || (self.adapter == Adapter::Ods
                && crate::format::ods::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odt")]
        let settable = settable
            || (self.adapter == Adapter::Odt
                && crate::format::odt::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odp")]
        let settable = settable
            || (self.adapter == Adapter::Odp
                && self.retains_source != Some(false)
                && crate::format::odp::SETTABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(odf_objects)]
        let settable = settable
            || (crate::format::odf_tree::writes(self.adapter)
                .is_some_and(|writes| writes.settable.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "iwork-adapter")]
        let settable = settable
            || (crate::format::iwork::writes(self.adapter)
                .is_some_and(|writes| writes.settable.contains(&self.data().canonical))
                && readable);
        let structural = {
            #[cfg(feature = "docx")]
            {
                self.adapter == Adapter::Docx
                    && self.retains_source != Some(false)
                    && matches!(
                        self.data().canonical,
                        "p" | "r" | "tbl" | "tr" | "hyperlink" | "comment"
                    )
                    && readable
            }
            #[cfg(not(feature = "docx"))]
            {
                false
            }
        };
        #[cfg(feature = "xlsx")]
        let structural = structural
            || (self.adapter == Adapter::Xlsx
                && self.retains_source != Some(false)
                && (matches!(self.data().canonical, "sheet" | "chart" | "run")
                    || crate::format::xlsx::FEATURE_ELEMENTS.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "pptx")]
        let add = structural
            || (self.adapter == Adapter::Pptx
                && self.retains_source != Some(false)
                && match self.data().canonical {
                    "slide" => self.pptx_slide_add == Some(true),
                    name => {
                        crate::format::pptx::SHAPE_ADDABLE.contains(&name)
                            && self.pptx_shape_add == Some(true)
                    },
                }
                && readable);
        #[cfg(not(feature = "pptx"))]
        let add = structural;
        // XLSX cells are added, not removed: a cell is cleared (record 0338).
        #[cfg(feature = "xlsx")]
        let add = add
            || (self.adapter == Adapter::Xlsx
                && self.retains_source != Some(false)
                && self.data().canonical == "cell"
                && readable);
        // DOCX pictures (records 0251d, 0335), charts (record 0283) and
        // SmartArt graphics and nodes (records 0295, 0335) are addable and
        // removable.
        #[cfg(feature = "docx")]
        let add = add
            || (self.adapter == Adapter::Docx
                && self.retains_source != Some(false)
                && (matches!(
                    self.data().canonical,
                    "picture"
                        | "chart"
                        | "smartart"
                        | "node"
                        | "tab"
                        | "break"
                        | "bookmark"
                        | "field"
                        | "toc"
                        // Cells and grid columns (record 0345).
                        | "tc"
                        | "col"
                ) || DOCX_PART_WRITES.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "xlsb")]
        let add = add
            || (self.adapter == Adapter::Xlsb
                && self.retains_source != Some(false)
                && (matches!(self.data().canonical, "sheet" | "namedrange")
                    || crate::format::xlsb::FEATURE_ELEMENTS.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "doc")]
        let add = add
            || (self.adapter == Adapter::Doc
                && self.retains_source != Some(false)
                && crate::format::doc::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ppt")]
        let add = add
            || (self.adapter == Adapter::Ppt
                && self.retains_source != Some(false)
                && crate::format::ppt::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "xls")]
        let add = add
            || (self.adapter == Adapter::Xls
                && self.retains_source != Some(false)
                && crate::format::xls::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "rtf")]
        let add = add
            || (self.adapter == Adapter::Rtf
                && crate::format::rtf::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ods")]
        let add = add
            || (self.adapter == Adapter::Ods
                && crate::format::ods::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odt")]
        let add = add
            || (self.adapter == Adapter::Odt
                && crate::format::odt::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odp")]
        let add = add
            || (self.adapter == Adapter::Odp
                && self.retains_source != Some(false)
                && crate::format::odp::ADDABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(odf_objects)]
        let add = add
            || (crate::format::odf_tree::writes(self.adapter)
                .is_some_and(|writes| writes.addable.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "iwork-adapter")]
        let add = add
            || (crate::format::iwork::writes(self.adapter)
                .is_some_and(|writes| writes.addable.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "pptx")]
        let remove = structural
            || (self.adapter == Adapter::Pptx
                && self.retains_source != Some(false)
                && ((self.data().canonical == "slide" && self.pptx_slide_remove == Some(true))
                    || (crate::format::pptx::REMOVABLE.contains(&self.data().canonical)
                        && self.pptx_shape_remove == Some(true)))
                && readable);
        #[cfg(not(feature = "pptx"))]
        let remove = structural;
        // DOCX content controls are removable but not yet addable (record
        // 0251a); cells and grid columns are both (record 0345).
        #[cfg(feature = "docx")]
        let remove = remove
            || (self.adapter == Adapter::Docx
                && self.retains_source != Some(false)
                && (matches!(
                    self.data().canonical,
                    "tc" | "sdt"
                        | "col"
                        | "chart"
                        | "picture"
                        | "smartart"
                        | "node"
                        | "tab"
                        | "break"
                        | "bookmark"
                        | "toc"
                ) || DOCX_PART_WRITES.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "xlsb")]
        let remove = remove
            || (self.adapter == Adapter::Xlsb
                && self.retains_source != Some(false)
                && (matches!(self.data().canonical, "cell" | "sheet" | "namedrange")
                    || crate::format::xlsb::FEATURE_ELEMENTS.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "xls")]
        let remove = remove
            || (self.adapter == Adapter::Xls
                && self.retains_source != Some(false)
                && crate::format::xls::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ppt")]
        let remove = remove
            || (self.adapter == Adapter::Ppt
                && self.retains_source != Some(false)
                && crate::format::ppt::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "doc")]
        let remove = remove
            || (self.adapter == Adapter::Doc
                && self.retains_source != Some(false)
                && crate::format::doc::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "rtf")]
        let remove = remove
            || (self.adapter == Adapter::Rtf
                && crate::format::rtf::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "ods")]
        let remove = remove
            || (self.adapter == Adapter::Ods
                && crate::format::ods::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odt")]
        let remove = remove
            || (self.adapter == Adapter::Odt
                && crate::format::odt::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(feature = "odp")]
        let remove = remove
            || (self.adapter == Adapter::Odp
                && self.retains_source != Some(false)
                && crate::format::odp::REMOVABLE.contains(&self.data().canonical)
                && readable);
        #[cfg(odf_objects)]
        let remove = remove
            || (crate::format::odf_tree::writes(self.adapter)
                .is_some_and(|writes| writes.removable.contains(&self.data().canonical))
                && readable);
        #[cfg(feature = "iwork-adapter")]
        let remove = remove
            || (crate::format::iwork::writes(self.adapter)
                .is_some_and(|writes| writes.removable.contains(&self.data().canonical))
                && readable);
        Operations {
            add,
            set: settable,
            get: readable,
            query: readable && reachable(self.adapter, self.index),
            remove,
        }
    }

    /// The keys a selector compares on the element besides its properties
    /// (`text`, `type`, `style`, and `value` for cells), as the query
    /// compiler resolves them.
    pub fn predicate_keys(&self) -> impl Iterator<Item = &'static str> + Clone {
        crate::query::compile::pseudo_keys(self.data().canonical)
    }

    /// The selector spellings of the element's property keys, as the query
    /// compiler maps them (`bold` for a cell's `font.bold`), with the key
    /// each one reads.
    pub fn predicate_aliases(&self) -> impl Iterator<Item = (&'static str, &'static str)> + Clone {
        let element = *self;
        crate::query::compile::key_aliases(self.adapter.family(), self.data().canonical)
            .filter(move |(_, key)| element.property(key).is_some())
    }

    /// Whether a selector's keys that are not properties name table header
    /// columns (`row[Sales>5000]`, record 0010 §5.4): workbook rows.
    pub(crate) fn header_columns(&self) -> bool {
        #[cfg(feature = "xls")]
        if self.adapter == Adapter::Xls {
            return false;
        }
        #[cfg(feature = "ods")]
        if self.adapter == Adapter::Ods {
            return false;
        }
        self.adapter.family() == Family::Workbook && self.data().canonical == "row"
    }
}

/// One row of a format's spec table, without its kind type: what the report
/// reads of an element.
#[derive(Debug, Clone, Copy)]
struct Data {
    name: &'static str,
    canonical: &'static str,
    aliases: &'static [&'static str],
    paths: &'static [&'static str],
    props: &'static [PropSpec],
    prop_aliases: &'static [crate::registry::PropAlias],
    gap: Option<&'static GapSpec>,
}

// Each helper below is one `with_spec!` dispatch, so a build without any
// adapter has no code after an uninhabited match.

/// The number of elements of `adapter`'s spec table.
fn element_count(adapter: Adapter) -> usize {
    with_spec!(adapter, spec => spec.elements.len())
}

/// The element at `index` of `adapter`'s spec table.
fn element_data(adapter: Adapter, index: u16) -> Data {
    with_spec!(adapter, spec => {
        let element = &spec.elements[usize::from(index)];
        let props = element.props;
        #[cfg(feature = "xlsx")]
        let props = if adapter == Adapter::Xlsx {
            match crate::path::navigate::ElementKind::name(element.kind) {
                "workbook" => &crate::format::xlsx::WORKBOOK_CAPABILITY_PROPS,
                "cell" => crate::format::xlsx::CELL_CAPABILITY_PROPS,
                "range" => crate::format::xlsx::RANGE_CAPABILITY_PROPS,
                "sheet" => crate::format::xlsx::SHEET_CAPABILITY_PROPS,
                "row" => crate::format::xlsx::ROW_CAPABILITY_PROPS,
                "col" => crate::format::xlsx::COLUMN_CAPABILITY_PROPS,
                "chart" => crate::format::xlsx::CHART_CAPABILITY_PROPS,
                "series" => crate::format::xlsx::SERIES_CAPABILITY_PROPS,
                "axis" => crate::format::xlsx::AXIS_CAPABILITY_PROPS,
                _ => props,
            }
        } else { props };
        #[cfg(feature = "ppt")]
        let props = if adapter == Adapter::Ppt {
            crate::format::ppt::capability_props(crate::path::navigate::ElementKind::name(
                element.kind,
            ))
            .unwrap_or(props)
        } else {
            props
        };
        #[cfg(feature = "pptx")]
        let props = if adapter == Adapter::Pptx {
            crate::format::pptx::capability_props(crate::path::navigate::ElementKind::name(
                element.kind,
            ))
            .unwrap_or(props)
        } else {
            props
        };
        #[cfg(feature = "xls")]
        let props = if adapter == Adapter::Xls {
            crate::format::xls::capability_props(crate::path::navigate::ElementKind::name(
                element.kind,
            ))
            .unwrap_or(props)
        } else {
            props
        };
        #[cfg(feature = "odp")]
        let props = if adapter == Adapter::Odp {
            crate::format::odp::capability_props(crate::path::navigate::ElementKind::name(
                element.kind,
            ))
            .unwrap_or(props)
        } else {
            props
        };
        #[cfg(feature = "ods")]
        let props = if adapter == Adapter::Ods {
            crate::format::ods::capability_props(crate::path::navigate::ElementKind::name(
                element.kind,
            ))
            .unwrap_or(props)
        } else {
            props
        };

        let canonical = crate::path::navigate::ElementKind::name(element.kind);
        Data {
            name: element.name,
            canonical,
            aliases: element.aliases,
            paths: element.paths,
            props,
            prop_aliases: crate::schema::prop_aliases(adapter, canonical),
            gap: element.gap,
        }
    })
}

/// The index of the parent (`parents`) or child at `next` of the element at
/// `index`.
fn related(adapter: Adapter, index: u16, parents: bool, next: u16) -> Option<u16> {
    with_spec!(adapter, spec => {
        let element = &spec.elements[usize::from(index)];
        let list = if parents { element.parents } else { element.children };
        list.get(usize::from(next))
            .map(|kind| u16::from(crate::path::navigate::ElementKind::ordinal(*kind)))
    })
}

/// Whether the query walker reaches the element at `index` from the
/// document root through the spec's child lists (the root itself is never a
/// match).
fn reachable(adapter: Adapter, index: u16) -> bool {
    with_spec!(adapter, spec => {
        let target = usize::from(index);
        let root = crate::path::navigate::ElementKind::ordinal(spec.root);
        // At most 128 kinds per format (`KindSet`), each pushed once.
        let mut seen = [false; 128];
        let mut stack = [0_u8; 128];
        let mut depth = 1;
        stack[0] = root;
        seen[usize::from(root)] = true;
        let mut found = false;
        while depth > 0 {
            depth -= 1;
            let at = usize::from(stack[depth]);
            for child in spec.elements[at].children {
                let ordinal = crate::path::navigate::ElementKind::ordinal(*child);
                found |= usize::from(ordinal) == target;
                if !seen[usize::from(ordinal)] {
                    seen[usize::from(ordinal)] = true;
                    stack[depth] = ordinal;
                    depth += 1;
                }
            }
        }
        found
    })
}

/// The gap of the element at `index`, or `None` past the last element.
fn element_gap(adapter: Adapter, index: usize) -> Option<Option<&'static GapSpec>> {
    with_spec!(adapter, spec => spec.elements.get(index).map(|element| element.gap))
}

/// The axis or property gap at `index` of `adapter`'s format.
fn format_gap(adapter: Adapter, index: usize) -> Option<&'static GapSpec> {
    with_spec!(adapter, spec => spec.gaps.get(index))
}

#[cfg(feature = "pptx")]
const fn pptx_create_gap(context: &'static str, reason: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "presentation",
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX creation options",
        reason,
    }
}

#[cfg(feature = "pptx")]
static PPTX_WRITE_GAPS: [GapSpec; 15] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "metadata",
            context: "path-read",
        },
        code: Code::UnsupportedCapability,
        follow_up: "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        reason: "`set(\"/metadata\", ..)` writes the OOXML document properties (root-metadata keys and custom.<name>, record 0253); the path is set-only, and reads go through File::metadata() and the root node",
    },
    pptx_create_gap(
        "template",
        "native PPTX template creation is not integrated",
    ),
    pptx_create_gap("strict", "native PPTX Strict creation is not integrated"),
    pptx_create_gap(
        "unrepresentable-metadata",
        "initial metadata is committed through the leaf metadata transaction; content_type, last_backup and codepage have no OOXML property part and are refused rather than dropped",
    ),
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "shape",
            context: "non-append-placement",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX non-append shape placement",
        reason: "PPTX shape creation currently supports append placement only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "shape",
            context: "non-slide-parent",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX group/master/layout shape creation",
        reason: "text-box creation requires a top-level slide parent; group, master and layout insertion remain pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "slide",
            context: "positional-owner",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX positional-owner slide ordering",
        reason: "source-backed PPTX slide ordering maintains p14 sections but still refuses outline slide lists, numeric presentation ranges, and unknown or opaque metadata extensions",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "slide",
            context: "positional-owner",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX positional-owner slide ordering",
        reason: "source-backed PPTX slide ordering maintains p14 sections but still refuses outline slide lists, numeric presentation ranges, and unknown or opaque metadata extensions",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "slide",
            context: "dependency-topology",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX dependency-aware slide removal",
        reason: "native PPTX slide removal removes the notes, comments, pictures, media and other Parts the slide owns exclusively and keeps shared ones; a slide another slide jumps to, or whose relationships name a missing Part, is still refused",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "slide",
            context: "final-slide",
        },
        code: Code::InvalidValue,
        follow_up: "M2 PPTX final-slide removal policy",
        reason: "native PPTX slide removal currently preserves the final slide and cannot remove the last slide from a presentation",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "slide",
            context: "canonical-target",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX equivalent relationship-target spellings",
        reason: "source-backed PPTX slide removal requires canonical owner-relative relationship targets so inverse replay can restore the original OPC spelling",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "shape",
            context: "dependency-topology",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX dependency-aware shape removal",
        reason: "shape removal refuses retained connector or animation references to removed IDs; source-backed removal preserves unused relationships and parts, while eager removal prunes unreachable relationship closure",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "shape",
            context: "non-top-level",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX grouped/master/layout shape removal",
        reason: "native PPTX shape removal currently requires a top-level slide shape; grouped children and master or layout shapes remain outside the admitted topology",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "shape",
            context: "native-shape-text-preservation",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX structural and rich-text editing",
        reason: "text replacement requires an existing slide shape text body and native edit admission; master/layout shapes, text-body creation and rich-text formatting remain pending",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "presentation",
            context: "managed-signature-removal",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPTX managed source-preserving signature removal",
        reason: "explicit signature removal requires native eager materialization, which managed sources refuse",
    },
];

#[cfg(feature = "ppt")]
const PPT_WRITE_GAPS: &[GapSpec] = crate::format::ppt::WRITE_GAPS;

/// Format-level option gaps that do not belong to the read spec table.
fn option_gap(adapter: Adapter, index: usize) -> Option<Option<&'static GapSpec>> {
    #[cfg(feature = "ppt")]
    if adapter == Adapter::Ppt {
        return PPT_WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(feature = "doc")]
    if adapter == Adapter::Doc {
        if let Some(gap) = DOC_CREATION_OPTION_GAPS.get(index) {
            return Some(Some(gap));
        }
        return DOC_WRITE_GAPS
            .get(index.saturating_sub(DOC_CREATION_OPTION_GAPS.len()))
            .map(Some);
    }
    #[cfg(feature = "pptx")]
    if adapter == Adapter::Pptx {
        return PPTX_WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(feature = "docx")]
    if adapter == Adapter::Docx {
        if let Some(gap) = DOCX_CREATION_OPTION_GAPS.get(index) {
            return Some(Some(gap));
        }
        let index = index - DOCX_CREATION_OPTION_GAPS.len();
        if let Some(gap) = DOCX_BATCH_GAPS.get(index) {
            return Some(Some(gap));
        }
        #[cfg(feature = "encryption")]
        return DOCX_SECURITY_OPTION_GAPS
            .get(index - DOCX_BATCH_GAPS.len())
            .map(Some);
    }
    #[cfg(feature = "xlsx")]
    if adapter == Adapter::Xlsx {
        if let Some(gap) = XLSX_CREATION_OPTION_GAPS.get(index) {
            return Some(Some(gap));
        }
        if index == XLSX_CREATION_OPTION_GAPS.len() {
            return Some(Some(&XLSX_MANAGED_PUBLICATION_GAP));
        }
        let index = index.saturating_sub(XLSX_CREATION_OPTION_GAPS.len() + 1);
        if let Some(gap) = XLSX_WRITE_GAPS.get(index) {
            return Some(Some(gap));
        }
        #[cfg(feature = "encryption")]
        return XLSX_SECURITY_OPTION_GAPS
            .get(index - XLSX_WRITE_GAPS.len())
            .map(Some);
    }
    #[cfg(feature = "xlsb")]
    if adapter == Adapter::Xlsb {
        if let Some(gap) = XLSB_CREATION_OPTION_GAPS.get(index) {
            return Some(Some(gap));
        }
        return XLSB_WRITE_GAPS
            .get(index.saturating_sub(XLSB_CREATION_OPTION_GAPS.len()))
            .map(Some);
    }
    #[cfg(feature = "xls")]
    if adapter == Adapter::Xls {
        if let Some(gap) = XLS_CREATION_OPTION_GAPS.get(index) {
            return Some(Some(gap));
        }
        return crate::format::xls::WRITE_GAPS
            .get(index.saturating_sub(XLS_CREATION_OPTION_GAPS.len()))
            .map(Some);
    }
    #[cfg(feature = "rtf")]
    if adapter == Adapter::Rtf {
        return crate::format::rtf::WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(feature = "ods")]
    if adapter == Adapter::Ods {
        return crate::format::ods::WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(feature = "odt")]
    if adapter == Adapter::Odt {
        return crate::format::odt::WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(feature = "odp")]
    if adapter == Adapter::Odp {
        return crate::format::odp::WRITE_GAPS.get(index).map(Some);
    }
    #[cfg(odf_objects)]
    if let Some(writes) = crate::format::odf_tree::writes(adapter) {
        return writes.gaps.get(index).map(Some);
    }
    #[cfg(feature = "iwork-adapter")]
    if let Some(writes) = crate::format::iwork::writes(adapter) {
        return writes.gaps.get(index).map(Some);
    }
    let _ = (adapter, index);
    None
}

/// The verbs that accept an element: OfficeCLI's help schema `operations`.
///
/// ```
/// use litchi_ops::capabilities::Operations;
///
/// let read_only = Operations { add: false, set: false, get: true, query: true, remove: false };
/// assert!(read_only.get && !read_only.set);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "OfficeCLI's help schema flags, one per verb"
)]
pub struct Operations {
    /// `add` creates the element (M2).
    pub add: bool,
    /// `set` changes its properties (M2).
    pub set: bool,
    /// `get` reads it.
    pub get: bool,
    /// `query` matches it.
    pub query: bool,
    /// `remove` deletes it (M2).
    pub remove: bool,
}

/// The parents or children of an [`Element`].
#[derive(Debug, Clone)]
pub struct Related {
    element: Element,
    parents: bool,
    next: u16,
}

impl Iterator for Related {
    type Item = Element;

    fn next(&mut self) -> Option<Element> {
        let element = self.element;
        let index = related(element.adapter, element.index, self.parents, self.next)?;
        self.next += 1;
        Some(Element {
            adapter: element.adapter,
            retains_source: element.retains_source,
            pptx_slide_add: element.pptx_slide_add,
            pptx_shape_add: element.pptx_shape_add,
            pptx_slide_remove: element.pptx_slide_remove,
            pptx_shape_remove: element.pptx_shape_remove,
            index,
        })
    }
}

/// Ordinary templates omit native IDs; semantic names retain their path forms.
fn ordinary_template(template: &str) -> bool {
    if template.contains("<id>") {
        return false;
    }
    let mut rest = template;
    while let Some((_, tail)) = rest.split_once("[@") {
        let key = tail.split(['=', ']']).next().unwrap_or("");
        if !["name", "role", "type"]
            .iter()
            .any(|semantic| key.eq_ignore_ascii_case(semantic))
        {
            return false;
        }
        rest = tail;
    }
    true
}

/// The declared properties of an [`Element`], including write-only entries.
#[derive(Debug, Clone)]
pub struct Properties {
    element: Element,
    aliases: &'static [crate::registry::PropAlias],
    retains_source: Option<bool>,
    props: std::slice::Iter<'static, PropSpec>,
}

impl Iterator for Properties {
    type Item = Property;

    fn next(&mut self) -> Option<Property> {
        let element = self.element;
        self.props
            .by_ref()
            .find(|property| {
                !with_spec!(element.adapter, spec => crate::query::compile::native_key(
                &spec.elements[usize::from(element.index)], property.name))
            })
            .map(|spec| Property {
                spec,
                aliases: self.aliases,
                retains_source: self.retains_source,
                structural_add_supported: {
                    #[cfg(feature = "pptx")]
                    {
                        if self.element.adapter == Adapter::Pptx {
                            match self.element.data().canonical {
                                "slide" => Some(self.element.pptx_slide_add == Some(true)),
                                "shape" => Some(self.element.pptx_shape_add == Some(true)),
                                // Placeholder insertion follows shape insertion
                                // (record 0346).
                                "placeholder" => Some(self.element.pptx_shape_add == Some(true)),
                                _ => None,
                            }
                        } else {
                            None
                        }
                    }
                    #[cfg(not(feature = "pptx"))]
                    {
                        None
                    }
                },
            })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (0, self.props.size_hint().1)
    }
}

/// One property of an element: its key, value type and the verbs that read
/// and write it (OfficeCLI's help schema property entry).
///
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Property {
    spec: &'static PropSpec,
    aliases: &'static [crate::registry::PropAlias],
    retains_source: Option<bool>,
    structural_add_supported: Option<bool>,
}

impl Property {
    /// The key (`font.bold`).
    #[must_use]
    pub const fn name(&self) -> &'static str {
        self.spec.name
    }

    /// The value type, which coerces a selector literal and prints the
    /// value.
    #[must_use]
    pub const fn value_type(&self) -> ValueType {
        self.spec.ty
    }

    /// Whether `get` and `query` read the key.
    #[must_use]
    pub const fn get(&self) -> bool {
        self.spec.access.get
    }

    /// Whether `set` writes the key (M2).
    #[must_use]
    pub const fn set(&self) -> bool {
        match self.retains_source {
            Some(false) => false,
            Some(true) | None => self.spec.access.set,
        }
    }

    /// The other spellings `set` and `add` accept for the key
    /// (`font.strike` for `strike`), as the registry declares them.
    pub fn aliases(&self) -> impl Iterator<Item = &'static str> + Clone {
        let name = self.spec.name;
        self.aliases
            .iter()
            .filter(move |alias| alias.name == name)
            .map(|alias| alias.alias)
    }

    /// The domain a written value must lie in beyond its type.
    #[must_use]
    pub const fn constraint(&self) -> Constraint {
        self.spec.admission.constraint
    }

    /// Whether a write of `null` clears the property.
    #[must_use]
    pub const fn clearable(&self) -> bool {
        self.spec.admission.clear
    }

    /// The gap a write of the key reports when the registry declares it but
    /// its setter is pending, with the follow-up that closes it.
    #[must_use]
    pub const fn write_gap(&self) -> Option<Gap> {
        match self.spec.admission.gap {
            Some(spec) => Some(Gap { spec }),
            None => None,
        }
    }

    /// Whether `add` accepts the key when it creates the element (M2).
    #[must_use]
    pub const fn add(&self) -> bool {
        if matches!(self.structural_add_supported, Some(false)) {
            return false;
        }
        match self.retains_source {
            Some(false) => false,
            Some(true) | None => self.spec.access.add,
        }
    }
}

/// An element, axis, property, operation context or view mode the format has
/// that Litchi cannot answer yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Gap {
    spec: &'static GapSpec,
}

impl Gap {
    /// What is missing.
    #[must_use]
    pub const fn subject(&self) -> Subject {
        self.spec.subject
    }

    /// The code a request that reaches the gap returns:
    /// `unsupported_capability` when the leaf lacks a model,
    /// `unsupported_path` when the format lacks the concept, or
    /// `document_protected` when the option would violate file protection.
    #[must_use]
    pub const fn code(&self) -> Code {
        self.spec.code
    }

    /// The leaf request or milestone that closes the gap (`L-XLSX-R5`,
    /// `M2`).
    #[must_use]
    pub const fn follow_up(&self) -> &'static str {
        self.spec.follow_up
    }

    /// One sentence on why the gap exists.
    #[must_use]
    pub const fn reason(&self) -> &'static str {
        self.spec.reason
    }
}

/// The gaps of a report: element gaps, then axis and property gaps, format
/// option gaps, and view gaps.
#[derive(Debug, Clone)]
pub struct Gaps {
    independent: std::slice::Iter<'static, GapSpec>,
    adapter: Option<Adapter>,
    phase: GapPhase,
    next: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum GapPhase {
    Elements,
    Format,
    Options,
    Views,
}

impl Iterator for Gaps {
    type Item = Gap;

    fn next(&mut self) -> Option<Gap> {
        if let Some(spec) = self.independent.next() {
            return Some(Gap { spec });
        }
        let adapter = self.adapter?;
        loop {
            let index = self.next;
            self.next += 1;
            let found: Option<Option<&'static GapSpec>> = match self.phase {
                GapPhase::Elements => element_gap(adapter, index),
                GapPhase::Format => format_gap(adapter, index).map(|gap| {
                    // View-mode dispatch supplies the authoritative view refusal.
                    (!matches!(gap.subject, Subject::View(_))).then_some(gap)
                }),
                GapPhase::Options => option_gap(adapter, index),
                GapPhase::Views => Mode::ALL.get(index).map(|mode| view_gap(adapter, *mode)),
            };
            match found {
                Some(Some(spec)) => return Some(Gap { spec }),
                Some(None) => {},
                None => {
                    self.next = 0;
                    self.phase = match self.phase {
                        GapPhase::Elements => GapPhase::Format,
                        GapPhase::Format => GapPhase::Options,
                        GapPhase::Options => GapPhase::Views,
                        GapPhase::Views => {
                            self.adapter = None;
                            return None;
                        },
                    };
                },
            }
        }
    }
}

/// The JSON name of a format: its usual extension (`docx`, `keynote`).
const fn format_name(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Doc => "doc",
        FileFormat::Docx => "docx",
        FileFormat::Ppt => "ppt",
        FileFormat::Pptx => "pptx",
        FileFormat::Xls => "xls",
        FileFormat::Xlsx => "xlsx",
        FileFormat::Xlsb => "xlsb",
        FileFormat::Rtf => "rtf",
        FileFormat::Pages => "pages",
        FileFormat::Keynote => "keynote",
        FileFormat::Numbers => "numbers",
        FileFormat::Odt => "odt",
        FileFormat::Ods => "ods",
        FileFormat::Odp => "odp",
        FileFormat::Odg => "odg",
        FileFormat::Odc => "odc",
        FileFormat::Odf => "odf",
        FileFormat::Odi => "odi",
        FileFormat::Odm => "odm",
        FileFormat::Oth => "oth",
        FileFormat::Odb => "odb",
        _ => "unknown",
    }
}

impl Serialize for Capabilities {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut report = serializer.serialize_struct("Capabilities", 7)?;
        report.serialize_field("format", &self.format.map(format_name))?;
        report.serialize_field("state", self.state.as_str())?;
        match self.retains_source {
            Some(retains) => report.serialize_field("retainsSource", &retains)?,
            None => report.skip_field("retainsSource")?,
        }
        report.serialize_field("verbs", &VerbsJson(self))?;
        report.serialize_field("views", &ViewsJson(self))?;
        report.serialize_field(
            "elements",
            &ElementsJson {
                elements: self.elements(),
                format: self.format,
            },
        )?;
        report.serialize_field("gaps", &GapsJson(self.gaps()))?;
        report.end()
    }
}

impl Serialize for Support {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match *self {
            Self::Supported => {
                let mut map = serializer.serialize_map(Some(1))?;
                map.serialize_entry("supported", &true)?;
                map.end()
            },
            Self::Unsupported {
                code,
                follow_up,
                reason,
            } => {
                let mut map = serializer.serialize_map(Some(4))?;
                map.serialize_entry("supported", &false)?;
                map.serialize_entry("code", code.as_str())?;
                map.serialize_entry("followUp", follow_up)?;
                map.serialize_entry("reason", reason)?;
                map.end()
            },
        }
    }
}

struct VerbsJson<'c>(&'c Capabilities);

impl Serialize for VerbsJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(Verb::ALL.len()))?;
        for verb in Verb::ALL {
            map.serialize_entry(verb.as_str(), &self.0.verb(verb))?;
        }
        map.end()
    }
}

struct ViewsJson<'c>(&'c Capabilities);

impl Serialize for ViewsJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(Mode::ALL.len()))?;
        for mode in Mode::ALL {
            map.serialize_entry(mode.name(), &self.0.view(*mode))?;
        }
        map.end()
    }
}

struct ElementsJson {
    elements: Elements,
    format: Option<FileFormat>,
}

impl Serialize for ElementsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for element in self.elements.clone() {
            seq.serialize_element(&ElementJson {
                element,
                format: self.format,
            })?;
        }
        seq.end()
    }
}

/// One element as OfficeCLI's help schema writes it, plus Litchi's
/// `predicates` and `gap` members. Native identity metadata is omitted.
struct ElementJson {
    element: Element,
    format: Option<FileFormat>,
}

impl Serialize for ElementJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.element;
        let mut map = serializer.serialize_map(None)?;
        if let Some(format) = self.format {
            map.serialize_entry("format", format_name(format))?;
        }
        map.serialize_entry("element", element.name())?;
        let mut parents = element.parents();
        match (parents.next(), parents.next()) {
            (None, _) => {},
            (Some(parent), None) => map.serialize_entry("parent", parent.name())?,
            (Some(_), Some(_)) => {
                map.serialize_entry("parent", &Names(element.parents()))?;
            },
        }
        if element.aliases().next().is_some() {
            map.serialize_entry("elementAliases", &Seq(element.aliases()))?;
        }
        map.serialize_entry("operations", &OperationsJson(element.operations()))?;
        map.serialize_entry("paths", &PathsJson(element))?;
        if element.children().next().is_some() {
            map.serialize_entry("children", &ChildrenJson(element))?;
        }
        map.serialize_entry("properties", &PropertiesJson(element))?;
        // The keys a selector compares, for the elements a query can match.
        if element.operations().query {
            map.serialize_entry("predicates", &PredicatesJson(element))?;
        }
        if let Some(gap) = element.gap() {
            map.serialize_entry(
                "gap",
                &GapJson {
                    gap,
                    subject: false,
                },
            )?;
        }
        map.end()
    }
}

/// A JSON array of the items of a clonable iterator.
struct Seq<I>(I);

impl<I> Serialize for Seq<I>
where
    I: Iterator + Clone,
    I::Item: Serialize,
{
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for item in self.0.clone() {
            seq.serialize_element(&item)?;
        }
        seq.end()
    }
}

/// A JSON array of element names.
struct Names(Related);

impl Serialize for Names {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        Seq(self.0.clone().map(|element| element.name())).serialize(serializer)
    }
}

struct OperationsJson(Operations);

impl Serialize for OperationsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let operations = self.0;
        let mut map = serializer.serialize_map(Some(5))?;
        map.serialize_entry("add", &operations.add)?;
        map.serialize_entry("set", &operations.set)?;
        map.serialize_entry("get", &operations.get)?;
        map.serialize_entry("query", &operations.query)?;
        map.serialize_entry("remove", &operations.remove)?;
        map.end()
    }
}

struct PathsJson(Element);

impl Serialize for PathsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.0;
        let mut map = serializer.serialize_map(None)?;
        for (key, positional) in [("stable", false), ("positional", true)] {
            let mut forms = element
                .paths()
                .filter(move |(_, is)| *is == positional)
                .map(|(template, _)| template)
                .peekable();
            if forms.peek().is_some() {
                map.serialize_entry(key, &Seq(forms))?;
            }
        }
        map.end()
    }
}

struct ChildrenJson(Element);

impl Serialize for ChildrenJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for child in self.0.children() {
            seq.serialize_element(&ChildJson(child))?;
        }
        seq.end()
    }
}

struct ChildJson(Element);

impl Serialize for ChildJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(2))?;
        map.serialize_entry("element", self.0.name())?;
        if let Some(segment) = self.0.path_segment() {
            map.serialize_entry("pathSegment", segment)?;
        }
        map.end()
    }
}

struct PropertiesJson(Element);

impl Serialize for PropertiesJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        for property in self.0.properties() {
            map.serialize_entry(property.name(), &PropertyJson(property))?;
        }
        map.end()
    }
}

/// A property as OfficeCLI's help schema writes it: `type`, `values` for an
/// enumeration, then the `add`, `set` and `get` flags.
struct PropertyJson(Property);

impl Serialize for PropertyJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let property = self.0;
        let ty = property.value_type();
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("type", ty.schema_type())?;
        if let ValueType::Token(tokens) = ty {
            map.serialize_entry("values", tokens.values())?;
        }
        map.serialize_entry("add", &property.add())?;
        map.serialize_entry("set", &property.set())?;
        map.serialize_entry("get", &property.get())?;
        map.end()
    }
}

struct PredicatesJson(Element);

impl Serialize for PredicatesJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.0;
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("keys", &Seq(element.predicate_keys()))?;
        if element.predicate_aliases().next().is_some() {
            map.serialize_entry("aliases", &AliasesJson(element))?;
        }
        if element.header_columns() {
            map.serialize_entry("headerColumns", &true)?;
        }
        map.end()
    }
}

/// The selector aliases of an element's keys, as a JSON object.
struct AliasesJson(Element);

impl Serialize for AliasesJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        for (alias, key) in self.0.predicate_aliases() {
            map.serialize_entry(alias, key)?;
        }
        map.end()
    }
}

struct GapsJson(Gaps);

impl Serialize for GapsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for gap in self.0.clone() {
            seq.serialize_element(&GapJson { gap, subject: true })?;
        }
        seq.end()
    }
}

/// A gap: its subject (in the report's `gaps` list; an element's own `gap`
/// member omits it), code, follow-up and reason.
struct GapJson {
    gap: Gap,
    subject: bool,
}

impl Serialize for GapJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let gap = self.gap;
        let mut map = serializer.serialize_map(None)?;
        if self.subject {
            match gap.subject() {
                Subject::Element(element) => {
                    map.serialize_entry("subject", "element")?;
                    map.serialize_entry("element", element)?;
                },
                Subject::Axis(axis) => {
                    map.serialize_entry("subject", "axis")?;
                    map.serialize_entry("axis", axis)?;
                },
                Subject::Property { element, key } => {
                    map.serialize_entry("subject", "property")?;
                    map.serialize_entry("element", element)?;
                    map.serialize_entry("key", key)?;
                },
                Subject::Operation {
                    verb,
                    element,
                    context,
                } => {
                    map.serialize_entry("subject", "operation")?;
                    map.serialize_entry("verb", verb)?;
                    map.serialize_entry("element", element)?;
                    map.serialize_entry("context", context)?;
                },
                Subject::View(mode) => {
                    map.serialize_entry("subject", "view")?;
                    map.serialize_entry("mode", mode)?;
                },
            }
        }
        map.serialize_entry("code", gap.code().as_str())?;
        map.serialize_entry("followUp", gap.follow_up())?;
        map.serialize_entry("reason", gap.reason())?;
        map.end()
    }
}

impl fmt::Display for Support {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Supported => f.write_str("supported"),
            Self::Unsupported {
                code,
                follow_up,
                reason,
            } => write!(f, "{} ({reason}; follow-up: {follow_up})", code.as_str()),
        }
    }
}

mod find;
mod m4;
mod merge;
mod refresh;
#[cfg(test)]
mod tests;

//! Handle for PowerPoint binary (`.ppt`) inputs (record 0003 §2.3, M0a step
//! S5d).
//!
//! The facade reaches the PPT reader by two routes, and the handle has one
//! crate-internal constructor for each:
//!
//! - the filesystem route, where detection opened a positional
//!   [`SourceBackedPackage`] over the file (`Ppt::from_source_package`);
//! - the byte route, where detection parsed the input into an in-memory
//!   [`OleFile`] (`Ppt::from_ole_file`).
//!
//! Both retain the canonical immutable PPT owner and its owned presentation
//! projection. The native owner preserves the original CFB index and source,
//! so the first edit starts from it without reopening it.

use std::io::{Cursor, Write};
use std::mem::size_of;
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_core::{FileFormat, Metadata};
use litchi_ppt::{Package, SourceBackedPackage, SourceLimits};

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::Code;
use crate::node::spec::{GapSpec, PropSpec, Subject};
use crate::{OpenOptions, Result};

mod children;
mod elements;
mod metadata;
mod nav;
mod node;
mod spec;
mod views;
pub(crate) mod write;

pub(crate) use children::Children;
pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub(crate) use spec::SPEC;
use spec::{
    CELL_CAPABILITY_PROPS, COMMENT_CAPABILITY_PROPS, HYPERLINK_CAPABILITY_PROPS,
    NOTES_CAPABILITY_PROPS, PARAGRAPH_CAPABILITY_PROPS, PLACEHOLDER_CAPABILITY_PROPS,
    RUN_CAPABILITY_PROPS, SHAPE_CAPABILITY_PROPS, SLIDE_CAPABILITY_PROPS,
};

/// PPT elements the unified `set` verb writes.
pub(crate) const SETTABLE: &[&str] = &[
    "slide",
    "shape",
    "placeholder",
    "notes",
    "paragraph",
    "comment",
    "tc",
    "hyperlink",
    "run",
];

/// PPT elements the unified `add` verb creates.
pub(crate) const ADDABLE: &[&str] = &["slide", "comment", "hyperlink"];

/// PPT elements the unified `remove` verb deletes.
pub(crate) const REMOVABLE: &[&str] = &["slide", "comment", "hyperlink"];

/// The capability property table of one PPT element, where its write
/// registry differs from its read table.
pub(crate) fn capability_props(name: &str) -> Option<&'static [PropSpec]> {
    Some(match name {
        "shape" => SHAPE_CAPABILITY_PROPS,
        "placeholder" => PLACEHOLDER_CAPABILITY_PROPS,
        "slide" => SLIDE_CAPABILITY_PROPS,
        "notes" => NOTES_CAPABILITY_PROPS,
        "paragraph" => PARAGRAPH_CAPABILITY_PROPS,
        "comment" => COMMENT_CAPABILITY_PROPS,
        "tc" => CELL_CAPABILITY_PROPS,
        "hyperlink" => HYPERLINK_CAPABILITY_PROPS,
        "run" => RUN_CAPABILITY_PROPS,
        _ => return None,
    })
}

/// Operation gaps of the PPT write surface, listed in `capabilities()`.
pub(crate) const WRITE_GAPS: &[GapSpec] = &[
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "ole",
            context: "object-storage",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F8 PPT OLE object writes",
        reason: "embedding an OLE object needs a new ExOleObjStg storage, an ExOleEmbedContainer and a picture frame with its preview; OLE frames are read (progId, linked, frame) only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "ole",
            context: "object-storage",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F8 PPT OLE object writes",
        reason: "OLE frame geometry and object replacement are not written; the frame is read only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "media",
            context: "media-objects",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F9 PPT media writes",
        reason: "a movie or sound needs an external-media object, its sound collection entry and a media picture frame; media frames are read only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "media",
            context: "media-objects",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F9 PPT media writes",
        reason: "media frame geometry and playback are not written through the media element; the frame is read only",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "theme",
            context: "theme-package",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F10 PPT theme writes",
        reason: "the theme is an OPC package inside each master's RoundTripThemeAtom; re-encoding it (and the masters' legacy color schemes that PowerPoint 97-2003 reads) is not modeled",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "chart",
            context: "graph-object",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F7 PPT chart series and axes",
        reason: "the embedded Graph object's chart stream is read through MS-OGRAPH; chart, series and axis writes are not integrated",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "run",
            context: "text-positions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F4 PPT run and line-break insertion and removal",
        reason: "inserting or removing a run or line break changes the text length, which every position-dependent record (style, special-information and interaction runs, fields) of the body would have to follow; runs are formatted and their text replaced in place",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "run",
            context: "text-positions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F4 PPT run and line-break insertion and removal",
        reason: "removing a run changes the text length that the body's position-dependent records hold",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "linebreak",
            context: "text-positions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F4 PPT run and line-break insertion and removal",
        reason: "a line break is a vertical-tab character of the text atom; inserting one changes the text length that the body's position-dependent records hold",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "tr",
            context: "table-layout",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F1 PPT table re-layout",
        reason: "a legacy table is a group of positioned cell shapes and border lines; inserting a row would re-lay out every later cell, line and the group anchor, which the leaf does not model",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "col",
            context: "table-layout",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F1 PPT table re-layout",
        reason: "a legacy table is a group of positioned cell shapes and border lines; inserting a column would re-lay out every later cell, line and the group anchor, which the leaf does not model",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "tr",
            context: "table-layout",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F1 PPT table re-layout",
        reason: "removing a row would re-lay out the table group's later cells and border lines, which the leaf does not model",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "col",
            context: "table-layout",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F1 PPT table re-layout",
        reason: "removing a column would re-lay out the table group's later cells and border lines, which the leaf does not model",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "tc",
            context: "native-cell-text",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0323-F2 PPT nested tables and rich cell text",
        reason: "cell text follows the shape-text closure rules (a changed length refuses over several formatting runs or position-dependent records); tables inside groups and merged-away cells are refused",
    },
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
            verb: "remove",
            element: "slide",
            context: "native-slide-dependencies",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPT dependency-preserving slide removal",
        reason: "slide removal requires native closure admission; unsupported dependent records refuse atomically",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "slide",
            context: "native-visibility-record",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPT visibility atom insertion",
        reason: "slide visibility writes require one native-admitted SSSlideInfoAtom; insertion of missing timing records is not integrated",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "shape",
            context: "native-shape-preservation",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPT nested-shape and rich-text edits",
        reason: "top-level shape, outline placeholder, paragraph and notes text and shape anchors require native closure admission: a changed text length refuses where interactions, field metacharacters or several formatting or language runs hold character positions; nested groups, new text bodies and rich-text formatting remain unsupported",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "batch",
            element: "presentation",
            context: "native-write-coalescing",
        },
        code: Code::UnsupportedCapability,
        follow_up: "M2 PPT broader native edit runs and performance qualification",
        reason: "distinct direct shape text/geometry sets share native transactions; selectors, repeated shapes, reads and structural edits close the run; broader coalescing and frozen-harness qualification remain open",
    },
];

pub(crate) use views::{Json as ViewJson, view_json, view_to};
pub(crate) use write::WriteCommit;

/// An opened legacy PowerPoint (`.ppt`) presentation.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the presentation reader built at open. The
/// handle retains an edit source unless [`OpenOptions::retain_source`] is false.
/// Positional reads retain their source for lazy payloads and freshness checks
/// even when editing is disabled.
#[derive(Clone)]
pub struct Ppt(Arc<Opened<PptState>>);

struct PptState {
    /// The reader every M0a read goes through, built at open exactly as the
    /// facade builds it.
    presentation: litchi_ppt::Presentation,
    /// The OLE property-set projection, read at open as the facade reads it
    /// (record 0003 §1.6); `None` when the streams are absent, unreadable or
    /// empty.
    metadata: Option<Metadata>,
    /// The source limits the reader was built under, inherited by editors.
    source_limits: SourceLimits,
    /// Whether the opened package carries password-to-open encryption.  This
    /// survives an explicit source release so capabilities and publication
    /// keep the protection policy visible without retaining the source owner.
    encrypted: bool,
    /// The source an editor is built from.
    owner: PptOwner,
}

/// The two source states of [`Ppt`] (record 0003 §2.3).
enum PptOwner {
    /// Filesystem route: the positional package detection opened.
    Source(Package),
    /// Byte route: the input bytes and the CFB index detection built, sealed
    /// together.
    Eager(Package),
    /// Edit ownership was released after open (`retain_source(false)`).
    Released,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Ppt>();
    assert!(size_of::<Ppt>() == size_of::<usize>());
};

/// Whether the PowerPoint reader refused an encrypted presentation for want
/// of a password, which makes the input a [`Locked`](crate::Locked) file.
///
/// PowerPoint's only password-to-open scheme is RC4 `CryptoAPI` ([MS-PPT]
/// 2.3.7). A reader built without its `encryption` feature cannot decrypt,
/// and reports that scheme as an unsupported encryption rather than asking
/// for a password; in a build without this crate's `encryption` feature that
/// refusal also means "encrypted". With the feature, the reader asks for the
/// password, and an unsupported encryption is a scheme it cannot decrypt.
pub(crate) fn password_required(error: &litchi_ppt::Error) -> bool {
    match error {
        litchi_ppt::Error::PasswordRequired => true,
        #[cfg(not(feature = "encryption"))]
        litchi_ppt::Error::UnsupportedEncryption(litchi_ppt::EncryptionKind::CryptoApi) => true,
        _ => false,
    }
}

impl Ppt {
    /// Create and adopt a blank legacy PowerPoint package directly through
    /// the native builder. The facade does not serialize and reopen the
    /// package, so the returned handle owns the same canonical package that
    /// the builder produced.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::create::{Conformance, Flavor};
        use crate::error::{Code, Error};

        let refusal = if options.flavor_value() != Flavor::Document {
            Some((
                "the PPT creator does not author template presentations",
                "use Flavor::Document; template authoring is not integrated",
            ))
        } else if options.conformance_value() == Conformance::Strict {
            Some((
                "legacy PPT does not support OOXML Strict conformance",
                "use Conformance::Transitional; Strict authoring is not integrated",
            ))
        } else {
            None
        };
        if let Some((message, suggestion)) = refusal {
            return Err(Error::new(Code::UnsupportedCapability, message)
                .with_format(FileFormat::Ppt)
                .with_suggestion(suggestion));
        }

        let package = match options.metadata_value() {
            Some(metadata) => metadata::create(metadata)?,
            None => Package::new()?,
        };
        Self::from_package(package)
    }

    /// Adopt a newly built canonical package as an eager, editable handle.
    fn from_package(package: Package) -> Result<Self> {
        let presentation = package.presentation()?;
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let source_limits = package.source_limits();
        let encrypted = package.is_encrypted();
        Ok(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
            owner: PptOwner::Eager(package),
        })))
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        !matches!(self.0.owner, PptOwner::Released)
    }

    /// Whether two handles retain the same immutable source snapshot.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }

    /// The retained compound-file length, when an edit owner exists. This
    /// O(1) lower bound is checked before private native-editor allocation;
    /// the exact before/after patch remains the hard post-commit charge.
    pub(crate) fn retained_source_len(&self) -> Option<usize> {
        match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => {
                Some(usize::try_from(package.shared_ole_file().file_size()).unwrap_or(usize::MAX))
            },
            PptOwner::Released => None,
        }
    }

    /// Check the retained source and native edit handoff before a PPT edit
    /// copies the CFB source or builds its immutable slide snapshot. The
    /// exact patch charge remains authoritative after commit.
    pub(crate) fn preflight_snapshot_bytes(&self, maximum: usize, additional: usize) -> Result<()> {
        let source = self.retained_source_len().unwrap_or(0);
        let observed = source
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_ppt::Patch>())
            .saturating_add(size_of::<litchi_ppt::package::edit::Edit>())
            .saturating_add(additional);
        if observed > maximum {
            return Err(crate::transaction::Error::Limit {
                kind: crate::transaction::LimitKind::SnapshotBytes,
                observed,
                maximum,
            }
            .into());
        }
        Ok(())
    }

    /// Admit a primary native PPT edit under the signature, encryption and
    /// source-retention policy shared by the other native formats.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()?;
        if self.is_encrypted() {
            return Err(litchi_ppt::Error::EncryptionPolicy.into());
        }
        if self.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Ppt));
        }
        if !self.retains_source() {
            return Err(crate::Error::new(
                Code::UnsupportedCapability,
                "the PPT source was released before editing",
            )
            .with_format(FileFormat::Ppt)
            .with_suggestion(
                "reopen the presentation without `OpenOptions::retain_source(false)`",
            ));
        }
        Ok(())
    }

    /// Admit a native edit that may still prove an exact signed no-op. A
    /// changed package is rejected by the publication fence, while encrypted
    /// and released owners remain protected before staging.
    pub(crate) fn admit_edit_for_write(&self) -> Result<()> {
        self.check_source_for_write()?;
        if self.is_encrypted() {
            return Err(litchi_ppt::Error::EncryptionPolicy.into());
        }
        if !self.retains_source() {
            return Err(crate::Error::new(
                Code::UnsupportedCapability,
                "the PPT source was released before editing",
            )
            .with_format(FileFormat::Ppt)
            .with_suggestion(
                "reopen the presentation without `OpenOptions::retain_source(false)`",
            ));
        }
        Ok(())
    }

    /// Return the native package that may receive one immutable edit.
    /// Authorized signature removal starts from an unsigned copy of the exact
    /// retained CFB artifact.
    pub(crate) fn editable_package(&self) -> Result<Package> {
        self.admit_edit()?;
        let package = match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => package,
            PptOwner::Released => unreachable!("admit_edit checks source retention"),
        };
        if self.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_encrypted_bytes()?,
                litchi_sign::cfb::Format::Ppt,
            )?;
            return Ok(Package::from_bytes_with_options(
                &bytes,
                litchi_ppt::OpenOptions::new()
                    .source_limits(self.0.source_limits)
                    .leniency(package.leniency()),
            )?);
        }
        Ok(package.clone())
    }

    /// Return the retained package for a setter that may prove an exact
    /// signed no-op. Removal-authorized handles use the explicit unsigned
    /// copy; signed handles retain the original package until publication
    /// proves that the native patch is unchanged.
    pub(crate) fn editable_package_for_write(&self) -> Result<Package> {
        self.admit_edit_for_write()?;
        let package = match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => package,
            PptOwner::Released => unreachable!("admit_edit_for_write checks source retention"),
        };
        if matches!(
            self.signature(),
            Signature::Signed | Signature::RemovalAuthorized
        ) {
            let bytes = signature::unsign_compound_file(
                package.to_encrypted_bytes()?,
                litchi_sign::cfb::Format::Ppt,
            )?;
            return Ok(Package::from_bytes_with_options(
                &bytes,
                litchi_ppt::OpenOptions::new()
                    .source_limits(self.0.source_limits)
                    .leniency(package.leniency()),
            )?);
        }
        Ok(package.clone())
    }

    /// Admit the writable package against the remaining snapshot budget
    /// before signature removal or native edit capture can allocate.
    pub(crate) fn editable_package_for_write_with_snapshot_limit(
        &self,
        maximum: usize,
    ) -> Result<Package> {
        self.preflight_snapshot_bytes(maximum, 0)?;
        self.editable_package_for_write()
    }

    /// Admit an ordinary unsigned package against the remaining snapshot
    /// budget before native edit capture can allocate.
    pub(crate) fn editable_package_with_snapshot_limit(&self, maximum: usize) -> Result<Package> {
        self.preflight_snapshot_bytes(maximum, 0)?;
        self.editable_package()
    }

    /// Return the original native package for a canonical signed no-op edit.
    /// Changed writes use [`Self::editable_package_for_write`] so signature
    /// storages are removed before publication; the empty patch used for an
    /// exact cancellation must remain bound to the signed CFB bytes.
    pub(crate) fn original_package_for_noop(&self) -> Result<Package> {
        self.admit_edit_for_write()?;
        match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => Ok(package.clone()),
            PptOwner::Released => Err(crate::Error::new(
                Code::UnsupportedCapability,
                "the PPT source was released before signed no-op replay",
            )
            .with_format(FileFormat::Ppt)
            .with_suggestion("reopen the presentation with source retention enabled")),
        }
    }

    /// Adopt a committed native package as an eager immutable handle while
    /// retaining the origin and the explicit signature transition.
    pub(crate) fn with_package(&self, package: Package) -> Result<Self> {
        let signature = match self.signature() {
            Signature::RemovalAuthorized => Signature::Unsigned,
            signature => signature,
        };
        Ok(Self::from_package(package)?
            .with_origin(self.origin())
            .with_signature(signature))
    }

    /// One-shot unified PPT setter.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
    ) -> Result<WriteCommit> {
        self.set_with_limits(
            target,
            props,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Set PPT shape or slide properties under an explicit semantic budget.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.set_with_snapshot_budget(target, props, limits, limits.max_snapshot_bytes())
    }

    /// Set through the native writer with a caller-owned remaining snapshot
    /// budget. Transactions use this after earlier leaf commits have already
    /// consumed part of their aggregate limit.
    pub(crate) fn set_with_snapshot_budget<'p, 'v>(
        &self,
        target: crate::target::TargetArg<'p>,
        props: impl crate::props::IntoProps<'v>,
        limits: &crate::transaction::SemanticLimits,
        snapshot_budget: usize,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let props = crate::Props::from_props(props)?;
        if let Some(commit) =
            write::coalesce::single_text_set(self, &target, &props, limits, snapshot_budget)
        {
            self.check_source_for_write()?;
            return self.enforce_signed_noop(Ok(commit));
        }
        let result = write::set_with_snapshot_budget(self, target, props, limits, snapshot_budget);
        self.check_source_for_write()?;
        self.enforce_signed_noop(result)
    }

    /// Admit one direct shape path for a transaction coalescer.
    ///
    /// The coalescer deliberately accepts only an absolute ordinal path to a
    /// top-level shape or placeholder. Queries, slide targets, raw paths and
    /// every structural or content selector remain on the ordinary writer so
    /// their resolution and diagnostics retain the one-shot contract.
    pub(crate) fn shape_plan(
        &self,
        target: &crate::target::TargetArg<'_>,
        props: &crate::Props<'_>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<Option<write::coalesce::Plan>> {
        let mut scene = None;
        self.shape_plan_cached(target, props, limits, &mut scene)
            .map(|plan| plan.map(|plan| plan.with_scene(scene)))
    }

    /// Resolve one direct shape path using a caller-owned slide scene cache.
    /// A transaction run uses this form to avoid reparsing one slide for each
    /// ordinal shape in a homogeneous batch.
    pub(crate) fn shape_plan_cached(
        &self,
        target: &crate::target::TargetArg<'_>,
        props: &crate::Props<'_>,
        limits: &crate::transaction::SemanticLimits,
        scene: &mut Option<write::coalesce::SceneCache>,
    ) -> Result<Option<write::coalesce::Plan>> {
        self.check_source_for_write()?;
        let result = write::coalesce::shape_plan(self, target, props, limits, scene);
        self.check_source_for_write().and(result)
    }

    /// Insert one blank slide under the presentation root.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(WriteCommit, crate::path::Path<'static>)> {
        self.add_with_limits(parent, node, &crate::transaction::SemanticLimits::default())
    }

    /// Insert one blank slide under an explicit semantic budget.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<(WriteCommit, crate::path::Path<'static>)> {
        self.check_source_for_write()?;
        let result = write::add_with_limits(self, parent, node, limits);
        self.check_source_for_write()?;
        result.and_then(|(commit, path)| {
            self.enforce_signed_noop(Ok(commit))
                .map(|commit| (commit, path))
        })
    }

    /// Remove slides selected by a path or a scoped query through the native
    /// primary PPT transaction.
    pub(crate) fn remove<'p>(&self, target: crate::target::TargetArg<'p>) -> Result<WriteCommit> {
        self.remove_with_limits(target, &crate::transaction::SemanticLimits::default())
    }

    /// Remove slides under an explicit semantic expansion and snapshot budget.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: crate::target::TargetArg<'p>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let result = write::remove_with_limits(self, target, limits);
        self.check_source_for_write()?;
        self.enforce_signed_noop(result)
    }

    /// Move one slide to a final sibling position or relative placement.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::placement::Placement<'q>,
    ) -> Result<WriteCommit> {
        self.move_to_with_limits(
            path,
            placement,
            &crate::transaction::SemanticLimits::default(),
        )
    }

    /// Move one slide under an explicit semantic expansion and snapshot budget.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::placement::Placement<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let result = write::move_to_with_limits(self, path, placement, limits);
        self.check_source_for_write()?;
        self.enforce_signed_noop(result)
    }

    /// Exchange two slides through one native primary PPT transaction.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<WriteCommit> {
        self.swap_with_limits(left, right, &crate::transaction::SemanticLimits::default())
    }

    /// Exchange two slides under an explicit semantic expansion and snapshot budget.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &crate::transaction::SemanticLimits,
    ) -> Result<WriteCommit> {
        self.check_source_for_write()?;
        let result = write::swap_with_limits(self, left, right, limits);
        self.check_source_for_write()?;
        self.enforce_signed_noop(result)
    }

    fn enforce_signed_noop(&self, result: Result<WriteCommit>) -> Result<WriteCommit> {
        result.and_then(|commit| {
            if self.signature() != Signature::Signed || commit.after.same_snapshot(self) {
                Ok(commit)
            } else {
                Err(signature::refusal(FileFormat::Ppt))
            }
        })
    }

    /// Opens the presentation over the positional package filesystem
    /// detection opened (the facade's `from_native_ppt_package`).
    ///
    /// The metadata is projected through the package's shared CFB view, the
    /// reader is built, and the package is kept as the source state.
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let package = Package::from_source(package, native_options(options, None))?;
        let presentation = package.presentation()?;
        let source_limits = package.source_limits();
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let encrypted = package.is_encrypted();
        let owner = if options.retain_source {
            PptOwner::Source(package)
        } else {
            PptOwner::Released
        };
        Ok(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
            owner,
        })))
    }

    /// Opens the presentation from the OLE file byte detection already parsed
    /// (the facade's `DetectedFormat::Ppt` arm).
    ///
    /// The PPT package adopts the parse, the metadata streams are projected,
    /// and the reader is built (with `password` when [`Locked::unlock`]
    /// supplies one). Then the parse moves into the retained source; no byte
    /// is copied and the container is not indexed again. This keeps the
    /// input alive for as long as the handle (risk R5);
    /// [`OpenOptions::retain_source`]`(false)` releases it instead.
    ///
    /// An encrypted presentation opened without a password comes back as
    /// [`LegacyOpen::PasswordRequired`] with the container unchanged.
    ///
    /// [`Locked::unlock`]: crate::Locked
    pub(crate) fn from_ole_file(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: &OpenOptions,
        password: Option<&str>,
    ) -> Result<LegacyOpen<Self>> {
        let package = match Package::from_ole_file_or_password_required(
            ole,
            native_options(options, password),
        )? {
            Ok(package) => package,
            Err(ole) => return Ok(LegacyOpen::PasswordRequired(Box::new(ole))),
        };
        let presentation = package.presentation()?;
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let source_limits = package.source_limits();
        let encrypted = package.is_encrypted();
        let owner = if options.retain_source {
            PptOwner::Eager(package)
        } else {
            PptOwner::Released
        };
        Ok(LegacyOpen::Opened(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
            owner,
        }))))
    }

    /// The same presentation under another signature disposition (record
    /// 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// Attach the stable path used by `File::save` without changing the
    /// immutable presentation owner.
    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    /// Return the stable path retained by a path-backed open.
    pub(crate) fn origin(&self) -> Option<Arc<std::path::Path>> {
        self.0.origin()
    }

    /// Whether ordinary publication must preserve an encryption profile.
    pub(crate) fn is_encrypted(&self) -> bool {
        self.0.encrypted
    }

    /// Fence a source-backed owner before publication.  Eager owners use an
    /// immutable owned source and therefore have a stable revision too;
    /// released read projections deliberately have no publication owner.
    pub(crate) fn check_source(&self) -> Result<()> {
        // Every constructor obtains this presentation from the retained
        // package. They share one source identity; checking both would probe
        // the same positional source twice. A released handle retains only
        // the presentation, which remains responsible for its lazy reads.
        match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => {
                package.source_version().map(drop).map_err(Into::into)
            },
            PptOwner::Released => self.0.presentation.check_source().map_err(Into::into),
        }
    }

    /// Source freshness for a mutation that may still prove an exact signed
    /// no-op. Changed native results are rejected at publication.
    pub(crate) fn check_source_for_write(&self) -> Result<()> {
        self.check_source()
    }

    /// Stream the exact native package.  A signed handle preserves the
    /// signature; only `Signed::unsign` reaches the explicit signature
    /// removal adapter.  A released source cannot publish because its edit
    /// owner was intentionally discarded at open.
    /// [`Self::write_to`] into an owned buffer: the leaf's single sized
    /// copy of the retained compound file.
    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        let package = match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => package,
            PptOwner::Released => {
                return Err(crate::Error::new(
                    Code::UnsupportedCapability,
                    "the PPT source was released before publication",
                )
                .with_format(FileFormat::Ppt)
                .with_suggestion(
                    "reopen the presentation without `OpenOptions::retain_source(false)`",
                ));
            },
        };
        self.check_source()?;
        if self.is_encrypted() {
            return Err(litchi_ppt::Error::EncryptionPolicy.into());
        }
        let bytes = if self.signature() == Signature::RemovalAuthorized {
            signature::unsign_compound_file(
                package.to_encrypted_bytes()?,
                litchi_sign::cfb::Format::Ppt,
            )?
        } else {
            package.to_bytes()?
        };
        self.check_source()?;
        Ok(bytes)
    }

    pub(crate) fn write_to<W: Write>(&self, sink: &mut W) -> Result<()> {
        let package = match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => package,
            PptOwner::Released => {
                return Err(crate::Error::new(
                    Code::UnsupportedCapability,
                    "the PPT source was released before publication",
                )
                .with_format(FileFormat::Ppt)
                .with_suggestion(
                    "reopen the presentation without `OpenOptions::retain_source(false)`",
                ));
            },
        };
        self.check_source()?;
        if self.is_encrypted() {
            return Err(litchi_ppt::Error::EncryptionPolicy.into());
        }
        if self.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_encrypted_bytes()?,
                litchi_sign::cfb::Format::Ppt,
            )?;
            sink.write_all(&bytes)?;
            self.check_source()?;
        } else {
            package.write_to(sink)?;
        }
        Ok(())
    }

    /// All text of the presentation, as the facade's `Presentation::text`
    /// returns it.
    ///
    /// # Errors
    ///
    /// Returns a typed PPT error when a slide cannot be decoded.
    pub fn text(&self) -> Result<String> {
        Ok(self.0.presentation.text()?)
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn slide_count(&self) -> Result<usize> {
        Ok(self.0.presentation.slide_count())
    }

    /// The presentation's OLE property-set metadata, read at open; `None`
    /// when the presentation carries none.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(self.0.metadata.clone())
    }

    /// Everything the reader repaired: the open, then every main master's,
    /// slide's and notes page's shapes, which this call parses.
    ///
    /// # Errors
    ///
    /// Returns a typed PPT error when a slide, notes page or shape tree
    /// cannot be parsed.
    pub fn tolerance_report(&self) -> Result<litchi_ppt::ToleranceReport> {
        Ok(self.0.presentation.collect_tolerance_report()?)
    }

    /// The reader the family views read slides from.
    pub(crate) fn reader(&self) -> &litchi_ppt::Presentation {
        &self.0.presentation
    }

    /// The length of the retained in-memory input, if the byte route
    /// retained one.
    #[cfg(test)]
    fn retained_input_len(&self) -> Option<u64> {
        match &self.0.owner {
            PptOwner::Eager(package) => Some(package.shared_ole_file().file_size()),
            PptOwner::Source(_) | PptOwner::Released => None,
        }
    }
}

impl std::fmt::Debug for Ppt {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let owner = match &self.0.owner {
            PptOwner::Source(_) => "source",
            PptOwner::Eager(_) => "eager",
            PptOwner::Released => "released",
        };
        formatter
            .debug_struct("Ppt")
            .field("owner", &owner)
            .finish_non_exhaustive()
    }
}

fn native_options<'a>(
    options: &OpenOptions,
    password: Option<&'a str>,
) -> litchi_ppt::OpenOptions<'a> {
    let native = litchi_ppt::OpenOptions::new().password(password);
    match options.ppt_source_limits {
        Some(limits) => native.source_limits(limits),
        None => native,
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::detect::{DetectedFormat, detect_format_smart};
    use litchi_ole_common::property_set::PropertySetReader;
    use litchi_ppt::raw::Reader;
    use std::path::PathBuf;

    fn fixture_path(name: &str) -> PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt")
            .join(name)
    }

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(fixture_path(name)).expect("PPT fixture")
    }

    /// The facade's byte route: detect, then hand the parse to the handle.
    /// A password-protected presentation reports the leaf's refusal, as the
    /// facade did; the unified opener turns it into `File::Locked`.
    fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Ppt> {
        match detect_format_smart(bytes, &crate::detect::SourceLimits::default()) {
            Some(DetectedFormat::Ppt(ole)) => match Ppt::from_ole_file(ole, options, None)? {
                LegacyOpen::Opened(ppt) => Ok(ppt),
                LegacyOpen::PasswordRequired(_) => Err(litchi_ppt::Error::PasswordRequired.into()),
            },
            other => panic!("detected {other:?}, not PPT"),
        }
    }

    /// The filesystem route: [`crate::open_with`] opens the positional
    /// PowerPoint package.
    #[cfg(any(unix, windows))]
    fn open_path(name: &str, options: &OpenOptions) -> Result<Ppt> {
        open_file(&fixture_path(name), options)
    }

    #[cfg(any(unix, windows))]
    fn open_file(path: &std::path::Path, options: &OpenOptions) -> Result<Ppt> {
        match crate::open_with(path, options)? {
            crate::File::Ppt(ppt) => Ok(ppt),
            other => panic!("{} was not opened as PPT: {other:?}", path.display()),
        }
    }

    /// The leaf pipeline the facade wraps, as an independent oracle.
    fn leaf(name: &str) -> (litchi_ppt::Presentation, Option<Metadata>) {
        let mut package = Reader::from_reader(Cursor::new(fixture(name))).unwrap();
        let metadata = package
            .ole_file()
            .get_metadata()
            .ok()
            .map(Metadata::from)
            .filter(Metadata::has_data);
        (package.presentation().unwrap(), metadata)
    }

    const FIXTURES: &[&str] = &[
        "SampleShow.ppt",
        "empty.ppt",
        "empty_textbox.ppt",
        "text_shapes.ppt",
        "text-margins.ppt",
        "ppt_with_png.ppt",
    ];

    fn assert_matches_leaf(ppt: &Ppt, name: &str) {
        let (presentation, metadata) = leaf(name);
        assert_eq!(ppt.text().unwrap(), presentation.text().unwrap(), "{name}");
        assert_eq!(
            ppt.slide_count().unwrap(),
            presentation.slide_count(),
            "{name}"
        );
        assert_eq!(
            format!("{:?}", ppt.metadata().unwrap()),
            format!("{metadata:?}"),
            "{name}"
        );
        assert_eq!(
            ppt.tolerance_report().unwrap(),
            presentation.collect_tolerance_report().unwrap(),
            "{name}"
        );
        let slides = presentation.slides().unwrap();
        for (index, slide) in slides.iter().enumerate() {
            let selected = ppt.reader().slide_at(index).unwrap().expect("slide");
            assert_eq!(selected.text().unwrap(), slide.text().unwrap(), "{name}");
        }
    }

    #[test]
    fn byte_route_matches_the_leaf_reader() {
        for name in FIXTURES {
            let ppt = open_bytes(fixture(name), &OpenOptions::default()).unwrap();
            assert_matches_leaf(&ppt, name);
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_route_matches_the_leaf_reader() {
        for name in FIXTURES {
            let ppt = open_path(name, &OpenOptions::default()).unwrap();
            assert!(matches!(ppt.0.owner, PptOwner::Source(_)));
            assert_matches_leaf(&ppt, name);
        }
    }

    #[test]
    #[cfg(any(unix, windows))]
    fn positional_admission_preserves_larger_selected_limits() {
        let limits = SourceLimits::builder()
            .max_input_bytes(1024 * 1024 * 1024)
            .unwrap()
            .max_aggregate_payload_bytes(2 * 1024 * 1024 * 1024)
            .unwrap()
            .build()
            .unwrap();
        let ppt = open_path(
            "SampleShow.ppt",
            &OpenOptions::new().ppt_source_limits(limits),
        )
        .unwrap();
        assert_eq!(ppt.0.source_limits, limits);
    }

    #[test]
    fn source_limits_reach_the_reader() {
        let limits = SourceLimits::builder()
            .max_entry_bytes(16)
            .expect("a finite entry limit is valid")
            .build()
            .expect("a checked PPT source limit is valid");
        let options = OpenOptions::default().ppt_source_limits(limits);
        let error = match open_bytes(fixture("SampleShow.ppt"), &options) {
            Ok(_) => panic!("a PPT larger than the package limit opened"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::LimitExceeded);

        #[cfg(any(unix, windows))]
        {
            let error = match open_path("SampleShow.ppt", &options) {
                Ok(_) => panic!("a PPT larger than the stream limit opened"),
                Err(error) => error,
            };
            assert_eq!(error.code(), Code::LimitExceeded);
        }
    }

    #[test]
    fn byte_route_retains_the_input_without_a_copy_and_releases_on_request() {
        let bytes = fixture("SampleShow.ppt");
        let length = u64::try_from(bytes.len()).unwrap();
        let retained = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        // Risk R5: the byte route now keeps exactly the input alive.
        assert_eq!(retained.retained_input_len(), Some(length));

        let released = open_bytes(bytes, &OpenOptions::default().retain_source(false)).unwrap();
        assert_eq!(released.retained_input_len(), None);
        assert_eq!(
            released.editable_package().unwrap_err().code(),
            Code::UnsupportedCapability
        );
        assert_eq!(released.text().unwrap(), retained.text().unwrap());
    }

    /// Every shape text of the editable package equals the leaf's fresh
    /// open; returns how many shapes carry text.
    fn assert_package_reads_like(package: &Package, oracle: &[u8]) -> usize {
        let texts = |package: &Package| -> Vec<Vec<String>> {
            package
                .presentation()
                .unwrap()
                .slides()
                .unwrap()
                .iter()
                .map(|slide| {
                    slide
                        .shapes()
                        .unwrap()
                        .iter()
                        .map(|shape| shape.text().unwrap_or_default())
                        .collect()
                })
                .collect()
        };
        let expected = texts(&Package::from_bytes(oracle).unwrap());
        let read = texts(package);
        assert_eq!(read, expected);
        read.iter()
            .flatten()
            .filter(|text| !text.is_empty())
            .count()
    }

    /// A presentation with three text boxes on each of two slides.
    fn editable_presentation() -> Vec<u8> {
        let mut writer = litchi_ppt::writer::Builder::new();
        for slide_index in 0..2 {
            let slide = writer.add_slide().unwrap();
            for shape_index in 0..3 {
                let x = 10 + shape_index * 260;
                let text = format!("slide-{slide_index}-shape-{shape_index}");
                writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
            }
        }
        let mut output = Cursor::new(Vec::new());
        litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn editable_packages_are_the_retained_source() {
        let bytes = editable_presentation();

        let from_bytes = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        let package = from_bytes.editable_package().unwrap();
        assert_eq!(assert_package_reads_like(&package, &bytes), 6);
        // The retained owner itself, not a copy: the sealed source's
        // process-unique identity carries over.
        let PptOwner::Eager(shared) = &from_bytes.0.owner else {
            panic!("byte route did not retain its source");
        };
        assert_eq!(
            package.source_version().unwrap(),
            shared.shared_ole_file().captured_source_version()
        );

        #[cfg(any(unix, windows))]
        {
            let file = tempfile::NamedTempFile::new().unwrap();
            std::fs::write(file.path(), &bytes).unwrap();
            let from_path = open_file(file.path(), &OpenOptions::default()).unwrap();
            let package = from_path.editable_package().unwrap();
            assert_eq!(assert_package_reads_like(&package, &bytes), 6);
            let PptOwner::Source(retained) = &from_path.0.owner else {
                panic!("filesystem route did not retain its package");
            };
            assert_eq!(
                package.source_version().unwrap(),
                retained.source_version().unwrap()
            );
            let released =
                open_file(file.path(), &OpenOptions::default().retain_source(false)).unwrap();
            assert!(released.editable_package().is_err());
        }
    }

    #[test]
    fn one_source_probe_fences_retained_and_released_presentations() {
        use litchi_core::{ReadAt, SourceVersion};
        use std::io;
        use std::sync::atomic::{AtomicUsize, Ordering};

        struct Source {
            bytes: Vec<u8>,
            probes: AtomicUsize,
            revision: AtomicUsize,
        }
        impl ReadAt for Source {
            fn len(&self) -> io::Result<u64> {
                u64::try_from(self.bytes.len()).map_err(io::Error::other)
            }
            fn version(&self) -> io::Result<SourceVersion> {
                self.probes.fetch_add(1, Ordering::Relaxed);
                Ok(SourceVersion::new(
                    221,
                    self.revision.load(Ordering::Relaxed) as u64,
                ))
            }
            fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
                let offset = usize::try_from(offset).map_err(io::Error::other)?;
                let input = self.bytes.get(offset..).unwrap_or_default();
                let count = input.len().min(output.len());
                output[..count].copy_from_slice(&input[..count]);
                Ok(count)
            }
        }
        let source = Arc::new(Source {
            bytes: Package::new().unwrap().to_bytes().unwrap(),
            probes: AtomicUsize::new(0),
            revision: AtomicUsize::new(0),
        });
        let package = Package::from_read_at(source.clone()).unwrap();
        let retained = Ppt::from_package(package.clone()).unwrap();
        let released = Ppt(Opened::new(PptState {
            presentation: package.presentation().unwrap(),
            metadata: None,
            source_limits: package.source_limits(),
            encrypted: false,
            owner: PptOwner::Released,
        }));
        for handle in [&retained, &released] {
            source.probes.store(0, Ordering::Relaxed);
            handle.check_source().unwrap();
            assert_eq!(source.probes.load(Ordering::Relaxed), 1);
        }
        source.revision.store(1, Ordering::Relaxed);
        for handle in [&retained, &released] {
            source.probes.store(0, Ordering::Relaxed);
            assert_eq!(handle.check_source().unwrap_err().code(), Code::IoError);
            assert_eq!(source.probes.load(Ordering::Relaxed), 1);
        }
    }

    #[test]
    fn handles_are_cheap_shared_clones() {
        let ppt = open_bytes(fixture("SampleShow.ppt"), &OpenOptions::default()).unwrap();
        let clone = ppt.clone();
        assert!(Arc::ptr_eq(&ppt.0, &clone.0));
        std::thread::scope(|scope| {
            scope.spawn(|| assert_eq!(clone.text().unwrap(), ppt.text().unwrap()));
        });
    }
}

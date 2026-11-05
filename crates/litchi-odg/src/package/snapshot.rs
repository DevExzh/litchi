//! Immutable ODG package snapshots and lossless semantic shape edits.

use crate::model::{
    FormControl,
    auxiliary::{Contour, ContourKind, GluePoint, ImageMap, ImageMapArea, ImageMapAreaShape},
    enhanced::{
        DrawingAttribute, DrawingAttributeNamespace, EnhancedGeometry, EnhancedGeometryChild,
        EnhancedGeometryChildKind,
    },
    group::Group,
    layer::Layer,
    page::Page,
    resource::Resource,
    shape::{Properties as ShapeProperties, Shape, ShapeKind},
    style::Style,
    style_resource::{StyleResource, StyleResourceKind},
};
use crate::transition::Transition;
use litchi_core::{
    BlobBundle, BlobLimits, CompositionLimits, DiagnosticFingerprint, Error,
    History as CoreHistory, HistoryLimits, JoinedSubEdits, Metadata, Patch as CorePatch,
    PatchLimits, PatchOperation, Result, Reversible, ReversibleOperation,
};
use litchi_odf_common::{
    Error as OdfError, OpenOptions, compact_xml,
    core::{
        AuthoredXmlFragment, PackageWriter, XmlSourcePart, XmlSplicePublication,
        family::Package as FamilyPackage,
    },
    drawing::Frame,
    media,
    package::{is_linked_href, resolve_package_path},
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use std::{
    borrow::Cow,
    collections::{BTreeMap, BTreeSet},
    fmt::Write as _,
    ops::{Deref, Range},
    path::Path,
    sync::Arc,
};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";
pub(crate) const TEMPLATE_MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics-template";
const BODY_MARKER: &str = "<office:drawing";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FORM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XML_EVENTS: &[u8] = b"http://www.w3.org/2001/xml-events";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SMIL: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_DEPTH: usize = 256;
const MAX_PAGES: usize = 16_384;
const MAX_LAYERS: usize = 16_384;
const MAX_FORM_CONTROLS: usize = 65_536;
const MAX_STYLE_RESOURCES: usize = 65_536;
const MAX_TRANSFER_RESOURCES: usize = 4_096;
const MAX_GROUP_EDITS: usize = 4_096;
const MAX_SHAPES: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_TRANSITION_XML_BYTES: usize = 8 * 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odg";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Dr3d,
    Svg,
    Text,
    Table,
    Form,
    Style,
    Presentation,
    Other,
}

type TextSpans = Vec<Vec<Vec<Option<Range<usize>>>>>;
type NameSpans = Vec<Vec<Option<Range<usize>>>>;
type LayerSpans = Vec<Vec<Option<Range<usize>>>>;
type GeometrySpans = Vec<Vec<[Option<Range<usize>>; 4]>>;
type PointsSpans = Vec<Vec<[Option<Range<usize>>; 2]>>;
type PathSpans = Vec<Vec<Option<Range<usize>>>>;
type ControlSpans = Vec<Vec<Option<Range<usize>>>>;
type PageAttributeSpans = Vec<[Option<Range<usize>>; 2]>;
type ShapeAttributeSpans = [Option<Range<usize>>; 16];

#[cfg(test)]
#[path = "attribute_span_tests.rs"]
mod attribute_span_tests;

#[cfg(test)]
#[path = "attribute_value_tests.rs"]
mod attribute_value_tests;

struct State {
    package: FamilyPackage,
    mimetype: &'static str,
    security: SecurityStatus,
    pages: Vec<Page>,
    layers: Vec<Layer>,
    resources: Vec<Resource>,
    form_controls: Vec<FormControl>,
    styles: Vec<Style>,
    style_resources: Vec<StyleResource>,
    active_content: ActiveContentStatus,
}

/// Inert package security state and mutation policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityStatus {
    signed: bool,
    encrypted: bool,
}

/// Stable capability contract for the supported ODG password and signature lifecycle.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityCapabilities {
    source: SecurityStatus,
    supported: u8,
}

/// One password/signature lifecycle operation that callers can plan before mutation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityLifecycleOperation {
    OpenWithPassword,
    EncryptNew,
    RewriteEncryptedExisting,
    ChangeExistingPassword,
    VerifySignatures,
    SignNew,
    RemoveInvalidatedSignatures,
    ResignExisting,
}

/// Final typed reason an existing-package security transition is unavailable.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityLifecycleRefusal {
    ExistingPackageReencryptionUnavailable,
    ExistingPackageResigningUnavailable,
}

impl SecurityLifecycleRefusal {
    /// The refusal class litchi-ops reports for this reason (record 0351
    /// §2.4): the route does not implement the operation.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        match self {
            Self::ExistingPackageReencryptionUnavailable
            | Self::ExistingPackageResigningUnavailable => {
                litchi_core::edit::RefusalKind::UnsupportedOperation
            },
        }
    }
}

/// Supported, exact-preservation-only, or finally unsupported lifecycle disposition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityLifecycleDisposition {
    Supported,
    ExactSourceOnly,
    Unsupported(SecurityLifecycleRefusal),
}

impl SecurityCapabilities {
    const OPEN_WITH_PASSWORD: u8 = 1 << 0;
    const ENCRYPT_NEW: u8 = 1 << 1;
    const REENCRYPT_EXISTING: u8 = 1 << 2;
    const VERIFY_SIGNATURES: u8 = 1 << 3;
    const SIGN_NEW: u8 = 1 << 4;
    const RESIGN_EXISTING: u8 = 1 << 5;
    const REMOVE_INVALIDATED_SIGNATURES: u8 = 1 << 6;

    /// Source protection inventory used when planning a lifecycle transition.
    #[must_use]
    pub const fn source(self) -> SecurityStatus {
        self.source
    }

    /// Returns the exact support/refusal disposition for one lifecycle operation.
    #[must_use]
    pub const fn disposition(
        self,
        operation: SecurityLifecycleOperation,
    ) -> SecurityLifecycleDisposition {
        match operation {
            SecurityLifecycleOperation::OpenWithPassword if self.can_open_with_password() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::EncryptNew if self.can_encrypt_new() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::RewriteEncryptedExisting if self.source.encrypted => {
                SecurityLifecycleDisposition::ExactSourceOnly
            },
            SecurityLifecycleOperation::RewriteEncryptedExisting => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::ChangeExistingPassword if self.can_reencrypt_existing() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::ChangeExistingPassword => {
                SecurityLifecycleDisposition::Unsupported(
                    SecurityLifecycleRefusal::ExistingPackageReencryptionUnavailable,
                )
            },
            SecurityLifecycleOperation::VerifySignatures if self.can_verify_signatures() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::SignNew if self.can_sign_new() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::RemoveInvalidatedSignatures
                if self.can_remove_invalidated_signatures() =>
            {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::ResignExisting if self.can_resign_existing() => {
                SecurityLifecycleDisposition::Supported
            },
            SecurityLifecycleOperation::ResignExisting => {
                SecurityLifecycleDisposition::Unsupported(
                    SecurityLifecycleRefusal::ExistingPackageResigningUnavailable,
                )
            },
            SecurityLifecycleOperation::OpenWithPassword
            | SecurityLifecycleOperation::EncryptNew
            | SecurityLifecycleOperation::VerifySignatures
            | SecurityLifecycleOperation::SignNew
            | SecurityLifecycleOperation::RemoveInvalidatedSignatures => {
                SecurityLifecycleDisposition::ExactSourceOnly
            },
        }
    }

    /// Whether password-encrypted drawings can be opened for inert inspection.
    #[must_use]
    pub const fn can_open_with_password(self) -> bool {
        self.supported & Self::OPEN_WITH_PASSWORD != 0
    }

    /// Whether a fresh drawing can be authored with password encryption.
    #[must_use]
    pub const fn can_encrypt_new(self) -> bool {
        self.supported & Self::ENCRYPT_NEW != 0
    }

    /// Whether an existing encrypted snapshot can be changed and re-encrypted.
    #[must_use]
    pub const fn can_reencrypt_existing(self) -> bool {
        self.supported & Self::REENCRYPT_EXISTING != 0
    }

    /// Whether inert signature metadata and trust-neutral signature math are available.
    #[must_use]
    pub const fn can_verify_signatures(self) -> bool {
        self.supported & Self::VERIFY_SIGNATURES != 0
    }

    /// Whether a fresh drawing can be authored with a document signature.
    #[must_use]
    pub const fn can_sign_new(self) -> bool {
        self.supported & Self::SIGN_NEW != 0
    }

    /// Whether an existing changed snapshot can be re-signed transactionally.
    #[must_use]
    pub const fn can_resign_existing(self) -> bool {
        self.supported & Self::RESIGN_EXISTING != 0
    }

    /// Whether invalidated signature members can be deliberately removed.
    #[must_use]
    pub const fn can_remove_invalidated_signatures(self) -> bool {
        self.supported & Self::REMOVE_INVALIDATED_SIGNATURES != 0
    }
}

/// Explicit mutation policy for protected drawing packages.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityWritePolicy {
    /// Refuse any write that would invalidate signatures or encryption.
    #[default]
    Refuse,
    /// Deliberately remove stale package signatures while preserving all unsigned payloads.
    RemoveSignatures,
}

/// Explicit disposition for inert active-content provenance during writes.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActiveContentWritePolicy {
    /// Preserve unmodified active-content markup inertly without executing it.
    #[default]
    PreserveInert,
    /// Refuse publication when the source inventories any active-content surface.
    Refuse,
}

/// Bounded inert inventory of active or externally resolved drawing surfaces.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ActiveContentStatus {
    scripts: usize,
    events: usize,
    actions: usize,
    dde: usize,
    external_links: usize,
    embedded_objects: usize,
}

impl ActiveContentStatus {
    /// Whether any active-content surface was inventoried.
    #[must_use]
    pub const fn is_present(self) -> bool {
        self.scripts != 0
            || self.events != 0
            || self.actions != 0
            || self.dde != 0
            || self.external_links != 0
            || self.embedded_objects != 0
    }

    /// Script-bearing element count.
    #[must_use]
    pub const fn scripts(self) -> usize {
        self.scripts
    }

    /// XML-event listener count.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Presentation action/listener count.
    #[must_use]
    pub const fn actions(self) -> usize {
        self.actions
    }

    /// DDE source count.
    #[must_use]
    pub const fn dde(self) -> usize {
        self.dde
    }

    /// External hyperlink count.
    #[must_use]
    pub const fn external_links(self) -> usize {
        self.external_links
    }

    /// Embedded object/plugin/applet/floating-frame count.
    #[must_use]
    pub const fn embedded_objects(self) -> usize {
        self.embedded_objects
    }
}

impl SecurityStatus {
    /// Whether document or macro signature metadata is present.
    #[must_use]
    pub const fn is_signed(self) -> bool {
        self.signed
    }

    /// Whether any manifest member has encryption metadata.
    #[must_use]
    pub const fn is_encrypted(self) -> bool {
        self.encrypted
    }

    /// Whether ordinary semantic rewrite is allowed without changing security lifecycle.
    #[must_use]
    pub const fn allows_rewrite(self) -> bool {
        !self.signed && !self.encrypted
    }
}

/// The immutable, cheap-to-clone owner of one ODG or OTG package.
///
/// `Package` owns every state verb: ingress (`open*`, `from_bytes*`,
/// `from_reader*`, [`Self::from_prepared_package`]), the primary
/// [`Edit`] transaction, patch application ([`Self::apply`]) and output
/// ([`Self::save`], [`Self::save_with_durability`], [`Self::to_bytes`],
/// [`Self::write_to`]). Clones share the exact bytes and the parsed
/// projection; [`Self::drawing`] returns the read-only [`crate::Drawing`]
/// projection. A commit publishes a new package and never mutates its
/// source.
///
/// Unknown package members and unmodeled XML remain in the retained source
/// bytes. Semantic inspection never evaluates controls, scripts, actions, DDE,
/// links, or embedded payloads.
#[derive(Clone)]
pub struct Package(Arc<State>);

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.as_bytes().len())
            .field("template", &self.is_template())
            .finish_non_exhaustive()
    }
}

/// One ODG package flavor: the regular drawing or the OTG template.
#[derive(Clone, Copy)]
struct Flavor {
    mimetype: &'static str,
    family_name: &'static str,
}

const DRAWING: Flavor = Flavor {
    mimetype: MIMETYPE,
    family_name: "ODG",
};

const TEMPLATE: Flavor = Flavor {
    mimetype: TEMPLATE_MIMETYPE,
    family_name: "OTG",
};

impl Package {
    /// The retained physical package (metadata transactions, record 0288).
    pub(crate) fn owned_package(&self) -> &litchi_odf_common::OwnedPackage {
        self.0.package.package()
    }

    /// Creates the minimal valid drawing (one empty `office:drawing` body).
    ///
    /// The bytes are deterministic: no clock or randomness is consulted.
    ///
    /// # Errors
    ///
    /// Returns an error only when the fixed minimal package cannot be built.
    pub fn new() -> std::result::Result<Self, OdfError> {
        crate::Builder::new().build()
    }

    /// Opens an ODG package from a filesystem path with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// Returns [`OdfError::Limit`] when a source or semantic limit is
    /// exceeded and [`OdfError::Core`] when the package cannot be read or is
    /// not a structurally valid ODG.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, OdfError> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens an ODG package from a filesystem path under caller-selected
    /// limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::open_flavor(path, options, DRAWING)
    }

    /// Opens an `OpenDocument` drawing template (OTG) from a filesystem path.
    ///
    /// # Errors
    ///
    /// As [`Self::open`], for a structurally valid OTG.
    pub fn open_template(path: impl AsRef<Path>) -> std::result::Result<Self, OdfError> {
        Self::open_template_with_options(path, OpenOptions::new())
    }

    /// Opens an OTG drawing template from a filesystem path under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open_template`].
    pub fn open_template_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::open_flavor(path, options, TEMPLATE)
    }

    /// Opens a password-protected ODG file for inert inspection.
    ///
    /// Encrypted packages remain read-only: semantic commit and durable
    /// application refuse to strip or silently re-encrypt protected entries.
    ///
    /// # Errors
    ///
    /// As [`Self::open`], plus unsupported encryption metadata or a bad
    /// password.
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::open_with_password_and_options(path, password, OpenOptions::new())
    }

    /// Opens a password-protected ODG file under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open_with_password`].
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let package = FamilyPackage::open_with_password_and_options(
            path,
            password,
            options,
            MIMETYPE,
            BODY_MARKER,
            DRAWING.family_name,
        )?;
        Self::from_family(package, MIMETYPE, options)
    }

    /// Opens an ODG package from borrowed bytes (copied once after the input
    /// size check).
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens an ODG package from borrowed bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_flavor(bytes, options, DRAWING)
    }

    /// Opens an OTG drawing template from borrowed bytes.
    ///
    /// # Errors
    ///
    /// As [`Self::open_template`].
    pub fn from_template_bytes(bytes: &[u8]) -> std::result::Result<Self, OdfError> {
        Self::from_template_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens an OTG drawing template from borrowed bytes under
    /// caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open_template`].
    pub fn from_template_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_flavor(bytes, options, TEMPLATE)
    }

    /// Opens password-protected ODG bytes for inert inspection.
    ///
    /// # Errors
    ///
    /// As [`Self::open_with_password`].
    pub fn from_bytes_with_password(
        bytes: &[u8],
        password: impl Into<String>,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_bytes_with_password_and_options(bytes, password, OpenOptions::new())
    }

    /// Opens password-protected ODG bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open_with_password`].
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: impl Into<String>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        check_input_bytes(bytes.len(), options)?;
        let package = FamilyPackage::from_bytes_with_password_and_options(
            copy_input(bytes)?,
            password,
            options,
            MIMETYPE,
            BODY_MARKER,
            DRAWING.family_name,
        )?;
        Self::from_family(package, MIMETYPE, options)
    }

    /// Opens an ODG package from a reader with the default [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader(reader: impl std::io::Read) -> std::result::Result<Self, OdfError> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Opens an ODG package from a reader under caller-selected limits. The
    /// reader is consumed through an input-limit guard before indexing.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let package = FamilyPackage::from_reader_with_options(
            reader,
            options,
            MIMETYPE,
            BODY_MARKER,
            DRAWING.family_name,
        )?;
        Self::from_family(package, MIMETYPE, options)
    }

    /// Adopts a packaged ODF result prepared by smart detection, reusing its
    /// validated archive index. Both ODG and OTG packages are accepted; the
    /// package MIME type selects the flavor.
    ///
    /// # Errors
    ///
    /// Returns [`OdfError::Core`] carrying [`Error::UnexpectedFormat`] when
    /// the prepared package is another family, or the usual validation error.
    pub fn from_prepared_package(
        prepared: litchi_odf_common::PreparedPackage,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_prepared_package_with_options(prepared, OpenOptions::new())
    }

    /// Adopt a prepared package under the semantic limits in `options`.
    /// The source limits were applied when the archive was prepared; its
    /// existing index and source allocation are reused.
    ///
    /// # Errors
    ///
    /// As [`Self::from_prepared_package`], plus a typed semantic limit failure.
    pub fn from_prepared_package_with_options(
        prepared: litchi_odf_common::PreparedPackage,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        if prepared.format() != litchi_core::detection::FileFormat::Odg {
            return Err(Error::UnexpectedFormat {
                detected: prepared.format(),
            }
            .into());
        }
        let archive = prepared.into_package();
        let flavor = if archive.mimetype()? == TEMPLATE_MIMETYPE {
            TEMPLATE
        } else {
            DRAWING
        };
        let package = FamilyPackage::from_owned_package_with_options(
            archive,
            options,
            flavor.mimetype,
            BODY_MARKER,
            flavor.family_name,
        )?;
        Self::from_family(package, flavor.mimetype, options)
    }

    fn open_flavor(
        path: impl AsRef<Path>,
        options: OpenOptions,
        flavor: Flavor,
    ) -> std::result::Result<Self, OdfError> {
        let package = FamilyPackage::open_with_options(
            path,
            options,
            flavor.mimetype,
            BODY_MARKER,
            flavor.family_name,
        )?;
        Self::from_family(package, flavor.mimetype, options)
    }

    fn from_bytes_flavor(
        bytes: &[u8],
        options: OpenOptions,
        flavor: Flavor,
    ) -> std::result::Result<Self, OdfError> {
        check_input_bytes(bytes.len(), options)?;
        let package = FamilyPackage::from_bytes_with_options(
            copy_input(bytes)?,
            options,
            flavor.mimetype,
            BODY_MARKER,
            flavor.family_name,
        )?;
        Self::from_family(package, flavor.mimetype, options)
    }

    /// Projects a validated family package and checks the page count against
    /// the semantic unit limit.
    fn from_family(
        package: FamilyPackage,
        mimetype: &'static str,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        let owner = Self::from_package(package, mimetype)?;
        let semantic = options.semantic_limits_value();
        if !semantic.is_default() {
            semantic.check(
                litchi_odf_common::SemanticLimitKind::Units,
                u64::try_from(owner.pages().len()).unwrap_or(u64::MAX),
            )?;
        }
        Ok(owner)
    }

    /// Reopens complete ODG bytes produced inside the crate.
    pub(crate) fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(
            FamilyPackage::from_bytes(bytes, MIMETYPE, BODY_MARKER, DRAWING.family_name)?,
            MIMETYPE,
        )
    }

    /// Reopens complete OTG bytes produced inside the crate.
    pub(crate) fn from_template_vec(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(
            FamilyPackage::from_bytes(bytes, TEMPLATE_MIMETYPE, BODY_MARKER, TEMPLATE.family_name)?,
            TEMPLATE_MIMETYPE,
        )
    }

    /// Reopens complete password-encrypted ODG bytes produced inside the crate.
    pub(crate) fn from_encrypted_vec(bytes: Vec<u8>, password: String) -> Result<Self> {
        Self::from_package(
            FamilyPackage::from_bytes_with_password(
                bytes,
                password,
                MIMETYPE,
                BODY_MARKER,
                DRAWING.family_name,
            )?,
            MIMETYPE,
        )
    }

    /// Reopens bytes of the same flavor as `self`.
    pub(crate) fn reopen_same_flavor(&self, bytes: Vec<u8>) -> Result<Self> {
        if self.is_template() {
            Self::from_template_vec(bytes)
        } else {
            Self::from_vec(bytes)
        }
    }

    /// The read-only drawing projection of this package (one reference-count
    /// increment).
    #[must_use]
    pub fn drawing(&self) -> crate::Drawing {
        crate::Drawing::from_package(self.clone())
    }

    /// Returns the identity of the validated archive index retained by this
    /// package, for diagnostics and tests.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.0.package.package().prepared_index_identity()
    }

    /// Repairs tolerated while this package was opened.
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::ToleranceReport {
        self.0.package.tolerance_report()
    }

    /// Applies a primary-transaction patch to this exact package.
    ///
    /// # Errors
    ///
    /// Returns an error when this package is not byte-identical to the patch
    /// source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, OdfError> {
        if !patch.is_applicable_to(self) {
            return Err(OdfError::patch_conflict(
                "odg.apply",
                "the package is not the patch source",
            ));
        }
        Ok(Commit {
            snapshot: patch.apply(self)?,
            changed: !patch.is_noop(),
            patch: patch.clone(),
        })
    }

    /// Atomically replaces `path` with this package (full durability).
    ///
    /// The artifact is staged in a sibling temporary file and published with
    /// one rename; a failure before publication leaves the destination
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`OdfError::Core`] carrying the I/O failure.
    pub fn save(&self, path: impl AsRef<Path>) -> std::result::Result<(), OdfError> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replaces `path` with this package under explicit durability.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> std::result::Result<(), OdfError> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }

    /// Copies the exact package bytes.
    ///
    /// # Errors
    ///
    /// Returns an allocation failure.
    pub fn to_bytes(&self) -> std::result::Result<Vec<u8>, OdfError> {
        Ok(copy_input(self.as_bytes())?)
    }

    /// Writes the exact package bytes to a stream.
    ///
    /// # Errors
    ///
    /// Returns the stream's I/O failure.
    pub fn write_to(&self, mut writer: impl std::io::Write) -> std::result::Result<(), OdfError> {
        writer.write_all(self.as_bytes()).map_err(Error::from)?;
        Ok(())
    }

    fn from_package(package: FamilyPackage, mimetype: &'static str) -> Result<Self> {
        let archive = package.package().package()?;
        let security = SecurityStatus {
            encrypted: archive.manifest().has_encrypted_entries(),
            signed: archive.has_file("META-INF/documentsignatures.xml")
                || archive.has_file("META-INF/macrosignatures.xml"),
        };
        let parsed = parse_content(package.content_xml())?;
        let content_styles = parse_style_definitions(package.content_xml())?;
        let named_styles = package
            .styles_xml()
            .map(parse_style_definitions)
            .transpose()?
            .unwrap_or_default();
        let page_transitions = resolve_page_transitions(
            package.content_xml(),
            package.styles_xml(),
            &parsed.pages,
            &content_styles,
            &named_styles,
        )?;
        let layers = package
            .styles_xml()
            .map(parse_declared_layers)
            .transpose()?
            .unwrap_or_default();
        if parsed.layer_count.saturating_add(layers.len()) > MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        let resources = scan_resources(&package)?;
        let mut styles = content_styles
            .into_iter()
            .chain(named_styles)
            .map(|definition| definition.style)
            .collect::<Vec<_>>();
        styles.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        styles
            .dedup_by(|left, right| left.name() == right.name() && left.family() == right.family());
        let active_content = scan_active_content(package.content_xml(), package.styles_xml())?;
        let mut style_resources = parse_style_resources(package.content_xml())?
            .into_iter()
            .map(|definition| definition.resource)
            .collect::<Vec<_>>();
        if let Some(styles_xml) = package.styles_xml() {
            style_resources.extend(
                parse_style_resources(styles_xml)?
                    .into_iter()
                    .map(|definition| definition.resource),
            );
        }
        style_resources.sort_unstable_by(|left, right| {
            left.kind()
                .cmp(&right.kind())
                .then_with(|| left.name().cmp(right.name()))
        });
        style_resources.dedup();
        let mut pages = parsed.pages;
        for (page, transition) in pages.iter_mut().zip(page_transitions) {
            page.set_transition(transition);
        }
        Ok(Self(Arc::new(State {
            package,
            mimetype,
            security,
            pages,
            form_controls: parsed.form_controls,
            styles,
            style_resources,
            active_content,
            layers,
            resources,
        })))
    }

    /// Returns the exact `content.xml` source.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    /// Returns exact `styles.xml`, when present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    /// Returns common document metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    /// Returns bounded pages in source order.
    #[must_use]
    pub fn pages(&self) -> &[Page] {
        &self.0.pages
    }

    /// Selects one page by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact name is ambiguous.
    pub fn page<'selector>(
        &self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Option<&Page>> {
        let resolved_selector = selector.into();
        match resolved_selector {
            crate::page::Selector::Position(position) => Ok(self.pages().get(position.get())),
            crate::page::Selector::Name(name) => {
                let mut matches = self
                    .pages()
                    .iter()
                    .filter(|page| page.name() == Some(name.as_ref()));
                let selected = matches.next();
                if selected.is_some() && matches.next().is_some() {
                    return invalid("ODG page name selector is ambiguous");
                }
                Ok(selected)
            },
        }
    }

    /// Returns global drawing layers declared by `styles.xml` in source order.
    ///
    /// Page-local declarations are available from [`Page::layers`](crate::page::Page::layers).
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.0.layers
    }

    /// Returns original package bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    /// Whether this snapshot is an OTG drawing template.
    #[must_use]
    pub fn is_template(&self) -> bool {
        self.0.mimetype == TEMPLATE_MIMETYPE
    }

    /// Inert signature/encryption state and rewrite policy for this snapshot.
    #[must_use]
    pub fn security(&self) -> SecurityStatus {
        self.0.security
    }

    /// Returns the stable supported password and signature lifecycle.
    #[must_use]
    pub fn security_capabilities(&self) -> SecurityCapabilities {
        SecurityCapabilities {
            source: self.0.security,
            supported: SecurityCapabilities::OPEN_WITH_PASSWORD
                | SecurityCapabilities::ENCRYPT_NEW
                | SecurityCapabilities::VERIFY_SIGNATURES
                | SecurityCapabilities::SIGN_NEW
                | SecurityCapabilities::REMOVE_INVALIDATED_SIGNATURES,
        }
    }

    /// Reads inert document and macro signature metadata without executing content.
    ///
    /// # Errors
    ///
    /// Returns an error when signature XML cannot be decoded safely.
    pub fn digital_signatures(&self) -> Result<litchi_odf_common::signature::DigitalSignatures> {
        self.0.package.package().digital_signatures()
    }

    /// Verifies document-signature math without making a certificate trust decision.
    ///
    /// # Errors
    ///
    /// Returns an error when signature metadata or referenced bytes cannot be verified.
    pub fn verify_document_signatures(
        &self,
    ) -> Result<Vec<litchi_odf_common::signature::SignatureVerification>> {
        self.0.package.package().verify_document_signatures()
    }

    /// Returns an inert, non-executing active-content inventory.
    #[must_use]
    pub fn active_content(&self) -> ActiveContentStatus {
        self.0.active_content
    }

    /// Lists safe package entry names.
    ///
    /// # Errors
    ///
    /// Returns an error when package member validation fails.
    pub fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    /// Returns package-local image resources referenced by drawing XML.
    #[must_use]
    pub fn resources(&self) -> &[Resource] {
        &self.0.resources
    }

    /// Returns inert form elements carrying `form:id` in source order.
    #[must_use]
    pub fn form_controls(&self) -> &[FormControl] {
        &self.0.form_controls
    }

    /// Returns inert drawing style definitions from content and styles parts.
    #[must_use]
    pub fn style_definitions(&self) -> &[Style] {
        &self.0.styles
    }

    /// Returns inert named gradients, hatches, fill images, markers, opacity, and stroke dashes.
    #[must_use]
    pub fn style_resources(&self) -> &[StyleResource] {
        &self.0.style_resources
    }

    /// Resolves one group root to its complete flattened nested descendant closure.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, non-group shape, or missing source span.
    pub fn group(&self, page: usize, shape: usize) -> Result<Group> {
        let parsed = parse_content(self.content_xml())?;
        group_selection(&parsed, page, shape)
    }

    /// Reads one inventoried package-local resource without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable package member.
    pub fn resource_bytes(&self, resource: usize) -> Result<Option<Vec<u8>>> {
        let selected = self.resources().get(resource).ok_or_else(|| {
            Error::InvalidFormat("ODG resource selector is out of bounds".to_string())
        })?;
        if !selected.is_present() {
            return Ok(None);
        }
        self.0.package.package().get_file(selected.path()).map(Some)
    }

    /// Starts a source-bound semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Edit {
        self.edit_with_policies(
            SecurityWritePolicy::Refuse,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Starts a transaction with an explicit protected-package write policy.
    ///
    /// Encryption is never silently stripped. `RemoveSignatures` permits signed-package edits and
    /// omits the now-stale signature members from the rebuilt package.
    #[must_use]
    pub fn edit_with_security_policy(&self, security_policy: SecurityWritePolicy) -> Edit {
        self.edit_with_policies(security_policy, ActiveContentWritePolicy::PreserveInert)
    }

    /// Starts a transaction with an explicit inert active-content disposition.
    #[must_use]
    pub fn edit_with_active_content_policy(
        &self,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Edit {
        self.edit_with_policies(SecurityWritePolicy::Refuse, active_content_policy)
    }

    /// Starts a transaction with explicit security and inert active-content dispositions.
    #[must_use]
    pub fn edit_with_policies(
        &self,
        security_policy: SecurityWritePolicy,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Edit {
        Edit {
            source: self.clone(),
            content: LazyPart::source(self, EditPart::Content),
            styles: self
                .styles_xml()
                .map(|_| LazyPart::source(self, EditPart::Styles)),
            content_splices: Some(Vec::new()),
            changes: Vec::new(),
            resource_edits: Vec::new(),
            requires_package_projection: false,
            pending_joined: None,
            package_replacement: None,
            security_policy,
            active_content_policy,
        }
    }

    /// Starts an empty deterministic composition for this exact snapshot.
    #[must_use]
    pub fn joined_edits(&self, limits: CompositionLimits) -> JoinedEdits {
        JoinedEdits::new(Lineage::new(self), limits)
    }

    /// Starts explicit bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        CoreHistory::new(self.clone(), limits)
    }

    /// Prepares one compact shape or complete group subtree for checked cross-drawing transfer.
    ///
    /// The plan retains exact source provenance, the referenced local layer declaration, and
    /// package-local resource bytes. It never evaluates embedded or linked content.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, noncompact referenced fragments, unreadable
    /// resources, or unresolved source dependencies.
    pub fn prepare_shape_transfer(&self, page: usize, shape: usize) -> Result<ShapeTransfer> {
        let source_namespaces = transfer_source_namespaces(self)?;
        let mut transfer_budget = TransferBudget::default();
        let parsed = parse_content(self.content_xml())?;
        let selected_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        let selected_shape = selected_page.shapes().get(shape).cloned().ok_or_else(|| {
            Error::InvalidFormat("ODG transfer shape selector is out of bounds".into())
        })?;
        let span = parsed.shape_spans[page][shape]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG transfer shape span is missing".into()))?;
        let raw_xml = self
            .content_xml()
            .get(span.clone())
            .ok_or_else(|| Error::InvalidFormat("ODG transfer shape span is invalid".into()))?;
        transfer_budget.charge(raw_xml.len(), "ODG transfer XML")?;
        let dependency_shapes = selected_page
            .shapes()
            .iter()
            .enumerate()
            .filter(|(index, _shape)| {
                parsed.shape_spans[page][*index]
                    .as_ref()
                    .is_some_and(|candidate| {
                        candidate.start >= span.start && candidate.end <= span.end
                    })
            })
            .map(|(_index, shape_value)| shape_value)
            .collect::<Vec<_>>();
        let mut layers = dependency_shapes
            .iter()
            .filter_map(|shape_value| shape_value.layer())
            .map(|name| resolve_transfer_layer(self, selected_page, name))
            .collect::<Result<Vec<_>>>()?;
        layers.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        layers.dedup_by(|left, right| left.name() == right.name());
        let mut required_styles = dependency_shapes
            .iter()
            .flat_map(|shape_value| [shape_value.style_name(), shape_value.text_style_name()])
            .flatten()
            .map(str::to_owned)
            .collect::<Vec<_>>();
        required_styles.sort_unstable();
        required_styles.dedup();
        let mut controls = dependency_shapes
            .iter()
            .filter_map(|shape_value| shape_value.control_reference())
            .map(str::to_owned)
            .collect::<Vec<_>>();
        controls.sort_unstable();
        controls.dedup();
        let mut style_sources = BTreeMap::new();
        let mut required_style_resources = BTreeSet::new();
        while let Some(style_name) = required_styles.pop() {
            if style_sources.contains_key(&style_name) {
                continue;
            }
            let definition = find_style_definition(self, &style_name)?.ok_or_else(|| {
                Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "odg.package.snapshot",
                    "ODG transfer source has unresolved style",
                ))
            })?;
            transfer_budget.charge(definition.xml.len(), "ODG transfer XML")?;
            if let Some(parent) = style_parent_name(&definition.xml)? {
                required_styles.push(parent);
            }
            for (path, value) in definition.style.properties() {
                let Some((_owner, attribute_name)) = path.rsplit_once('/') else {
                    continue;
                };
                if let Some(kind) = style_resource_reference_kind(attribute_name) {
                    required_style_resources.insert((kind, value.clone()));
                }
            }
            style_sources.insert(style_name, definition);
        }
        let mut style_resource_sources = BTreeMap::new();
        for (kind, value) in required_style_resources {
            let definition = find_style_resource(self, kind, &value)?.ok_or_else(|| {
                Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "odg.package.snapshot",
                    "ODG transfer source has unresolved named style resource",
                ))
            })?;
            transfer_budget.charge(definition.xml.len(), "ODG transfer XML")?;
            style_resource_sources.insert((kind, value), definition);
        }
        for control in &controls {
            if !self
                .form_controls()
                .iter()
                .any(|declared| declared.id() == control)
            {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "odg.package.snapshot",
                    "ODG transfer source has unresolved form control",
                )));
            }
        }
        let mut control_sources = Vec::new();
        for identifier in &controls {
            let position = parsed
                .form_controls
                .iter()
                .position(|control| control.id() == identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer form control is missing".into())
                })?;
            let control_span = parsed.form_control_spans[position]
                .as_ref()
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer form-control span is missing".into())
                })?;
            let raw_control_xml =
                self.content_xml()
                    .get(control_span.clone())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG transfer form-control span is invalid".into())
                    })?;
            transfer_budget.charge(raw_control_xml.len(), "ODG transfer XML")?;
            control_sources.push((parsed.form_controls[position].clone(), control_span.clone()));
        }

        // Namespace closure can add producer declarations to each detached
        // fragment.  Charge those bytes for the complete set before any
        // closure string is allocated, so a rejected aggregate cannot leave
        // several large detached strings behind.
        let mut charge_namespace_closure = |fragment: &str| -> Result<()> {
            let added = transfer_namespace_closure_bytes(fragment, &source_namespaces)?;
            transfer_budget.charge(added, "ODG transfer namespace closure")?;
            Ok(())
        };
        charge_namespace_closure(raw_xml)?;
        for definition in style_sources.values() {
            charge_namespace_closure(&definition.xml)?;
        }
        for definition in style_resource_sources.values() {
            charge_namespace_closure(&definition.xml)?;
        }
        for (_control, span) in &control_sources {
            let fragment = self.content_xml().get(span.clone()).ok_or_else(|| {
                Error::InvalidFormat("ODG transfer form-control span is invalid".into())
            })?;
            charge_namespace_closure(fragment)?;
        }

        // Every source XML span is charged before any detached transfer string
        // is cloned or namespace-closed.  This keeps a rejected aggregate from
        // allocating one large fragment at a time before discovering the cap.
        let xml = close_transfer_fragment_namespaces(raw_xml, &source_namespaces)?;
        validate_transfer_fragment_namespaces(&xml)?;
        compact_xml::validate(xml.as_bytes(), compact_xml::Limits::default())
            .map_err(Error::from)?;
        let styles = style_sources.keys().cloned().collect::<Vec<_>>();
        let mut style_definitions = BTreeMap::new();
        for (style_name, definition) in style_sources {
            let definition_xml =
                close_transfer_fragment_namespaces(&definition.xml, &source_namespaces)?;
            validate_transfer_fragment_namespaces(&definition_xml)?;
            compact_xml::validate(definition_xml.as_bytes(), compact_xml::Limits::default())
                .map_err(Error::from)?;
            style_definitions.insert(
                style_name.clone(),
                TransferStyle {
                    name: style_name,
                    family: definition.style.family().to_owned(),
                    parent: definition.style.parent().map(str::to_owned),
                    xml: definition_xml,
                },
            );
        }
        let mut style_resources = BTreeMap::new();
        for (key, definition) in style_resource_sources {
            let resource_xml =
                close_transfer_fragment_namespaces(&definition.xml, &source_namespaces)?;
            validate_transfer_fragment_namespaces(&resource_xml)?;
            compact_xml::validate(resource_xml.as_bytes(), compact_xml::Limits::default())
                .map_err(Error::from)?;
            style_resources.insert(
                key,
                TransferStyleResource {
                    resource: definition.resource,
                    xml: resource_xml,
                },
            );
        }
        let transfer_fragments = std::iter::once(xml.as_str())
            .chain(style_definitions.values().map(|style| style.xml.as_str()))
            .chain(
                style_resources
                    .values()
                    .map(|resource| resource.xml.as_str()),
            )
            .collect::<Vec<_>>();
        let mut resource_indices = Vec::new();
        for (resource_index, resource) in self.resources().iter().enumerate() {
            if !transfer_fragments
                .iter()
                .any(|fragment| transfer_xml_references(fragment, resource.href()))
            {
                continue;
            }
            if resource_indices
                .iter()
                .any(|index: &usize| self.resources()[*index].path() == resource.path())
            {
                continue;
            }
            if resource_indices.len() >= MAX_TRANSFER_RESOURCES {
                return invalid("ODG transfer resource count exceeds the limit");
            }
            resource_indices.push(resource_index);
        }
        let archive = self.0.package.package();
        let archive_view = archive.package()?;
        for resource_index in &resource_indices {
            let resource = &self.resources()[*resource_index];
            let size = archive_view
                .member_materialized_size(resource.path())?
                .ok_or_else(|| {
                    Error::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "odg.package.snapshot",
                        "ODG transfer resource is missing",
                    ))
                })?;
            if size > MAX_OUTPUT_BYTES as u64 {
                return invalid("ODG transfer resources exceed the byte limit");
            }
            let size = usize::try_from(size).map_err(|_| {
                Error::InvalidFormat("ODG transfer resource size overflows usize".into())
            })?;
            transfer_budget.charge(size, "ODG transfer closure")?;
        }
        let mut resources = Vec::new();
        for resource_index in resource_indices {
            let resource = &self.resources()[resource_index];
            let bytes = self.resource_bytes(resource_index)?.ok_or_else(|| {
                Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG transfer resource is missing",
                ))
            })?;
            if bytes.len() > MAX_OUTPUT_BYTES {
                return invalid("ODG transfer resources exceed the byte limit");
            }
            resources.push(TransferResource {
                href: resource.href().to_owned(),
                path: resource.path().to_owned(),
                media_type: resource.media_type().map(str::to_owned),
                bytes: Some(bytes),
            });
        }
        let control_definitions = control_sources
            .into_iter()
            .map(|(control, control_span)| {
                let raw_control_xml = self.content_xml().get(control_span).ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer form-control span is invalid".into())
                })?;
                let control_xml =
                    close_transfer_fragment_namespaces(raw_control_xml, &source_namespaces)?;
                validate_transfer_fragment_namespaces(&control_xml)?;
                compact_xml::validate(control_xml.as_bytes(), compact_xml::Limits::default())
                    .map_err(Error::from)?;
                Ok(TransferControl {
                    control,
                    xml: control_xml,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(ShapeTransfer {
            source: Lineage::new(self),
            shape: selected_shape,
            xml,
            layers,
            styles,
            style_definitions: style_definitions.into_values().collect(),
            style_resources: style_resources.into_values().collect(),
            controls,
            control_definitions,
            resources,
        })
    }
}

/// Refuses input above the selected source input limit before copying it.
fn check_input_bytes(length: usize, options: OpenOptions) -> std::result::Result<(), OdfError> {
    let limits = options.source_limits_value();
    let observed = u64::try_from(length).unwrap_or(u64::MAX);
    if observed > limits.max_input_bytes() {
        return Err(OdfError::Limit {
            kind: litchi_odf_common::LimitKind::Source(
                litchi_odf_common::SourceLimitKind::InputBytes,
            ),
            observed,
            maximum: limits.max_input_bytes(),
        });
    }
    Ok(())
}

/// Copies caller bytes with a fallible reservation.
fn copy_input(source: &[u8]) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "ODG package bytes",
            source,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

/// A staged source-bound package shape edit.
pub struct Edit {
    source: Package,
    content: LazyPart,
    styles: Option<LazyPart>,
    content_splices: Option<Vec<ContentSplice>>,
    changes: Vec<Change>,
    resource_edits: Vec<ResourceEdit>,
    requires_package_projection: bool,
    pending_joined: Option<JoinedEdits>,
    package_replacement: Option<Package>,
    security_policy: SecurityWritePolicy,
    active_content_policy: ActiveContentWritePolicy,
}

#[derive(Clone, Copy)]
enum EditPart {
    Content,
    Styles,
}

/// A source part borrowed through the owning package until the first mutation.
///
/// The package clone retained here is O(1); only `set_owned` materializes the
/// selected XML part.  `Deref` keeps read-only transaction paths borrowed.
struct LazyPart {
    source: Package,
    part: EditPart,
    owned: Option<String>,
}

impl LazyPart {
    fn source(source: &Package, part: EditPart) -> Self {
        Self {
            source: source.clone(),
            part,
            owned: None,
        }
    }

    fn source_text(&self) -> &str {
        match self.part {
            EditPart::Content => self.source.content_xml(),
            EditPart::Styles => self.source.styles_xml().unwrap_or_default(),
        }
    }

    fn as_str(&self) -> &str {
        self.owned.as_deref().unwrap_or_else(|| self.source_text())
    }

    fn set_owned(&mut self, value: String) {
        self.owned = Some(value);
    }
}

impl Deref for LazyPart {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

#[derive(Debug)]
struct ContentSplice {
    source_range: Range<usize>,
    current_range: Range<usize>,
    expected: Vec<u8>,
    replacement: Vec<u8>,
}

impl Edit {
    /// Appends one admitted disjoint composition after the ordinary staged
    /// operations in this edit.
    ///
    /// The composition is retained without publishing or rebuilding. Lineage,
    /// durable-payload validity, and conflicts with already staged work are
    /// checked before the composition is moved into the edit, so a refusal
    /// leaves this edit unchanged. A second composition is refused with the
    /// rejected plan intact: one plan preserves its original composition
    /// limits, stable identifiers, and deterministic order.
    pub fn join(&mut self, joined: JoinedEdits) -> std::result::Result<&mut Self, JoinError> {
        if self.pending_joined.is_some() {
            return Err(JoinError::new(
                joined_refusal("ODG edit accepts one bounded joined plan"),
                joined,
            ));
        }
        if !joined.lineage().matches(&self.source) {
            return Err(JoinError::new(Error::PatchConflict, joined));
        }
        if let Err(error) = validate_joined_effect_declarations(&joined) {
            return Err(JoinError::new(error, joined));
        }
        let source_fingerprint = fingerprint(self.source.as_bytes());
        let validation_error = joined.sub_edits().find_map(|sub_edit| {
            validate_durable_patch(&sub_edit.payload().inner)
                .err()
                .or_else(|| {
                    validate_joined_preconditions(sub_edit.payload(), &source_fingerprint).err()
                })
        });
        if let Some(error) = validation_error {
            return Err(JoinError::new(error, joined));
        }
        match joined_contains_mixed_package_effects(&joined) {
            Ok(true) => {
                return Err(JoinError::new(
                    joined_refusal(
                        "ODG joined package replacement conflicts with independent staged effects",
                    ),
                    joined,
                ));
            },
            Ok(false) => {},
            Err(error) => return Err(JoinError::new(error, joined)),
        }
        match self.joined_conflicts_with_staged(&joined) {
            Ok(true) => {
                return Err(JoinError::new(
                    joined_refusal("ODG joined edit overlaps staged work"),
                    joined,
                ));
            },
            Ok(false) => {},
            Err(error) => return Err(JoinError::new(error, joined)),
        }
        match joined_has_dependent_conflict(&self.source, &joined) {
            Ok(true) => {
                return Err(JoinError::new(
                    joined_refusal("ODG joined edit has dependent presentation effects"),
                    joined,
                ));
            },
            Ok(false) => {},
            Err(error) => return Err(JoinError::new(error, joined)),
        }
        if let Err(error) = self.validate_joined_page_names(&joined) {
            return Err(JoinError::new(error, joined));
        }
        self.pending_joined = Some(joined);
        Ok(self)
    }

    fn joined_conflicts_with_staged(&self, joined: &JoinedEdits) -> Result<bool> {
        let staged_package_effect = self.requires_package_projection
            || !self.resource_edits.is_empty()
            || self
                .changes
                .iter()
                .any(|change| matches!(change, Change::Structure(_)))
            || self.package_replacement.is_some();
        let direct_conflict =
            joined
                .sub_edits()
                .try_fold(false, |conflict, sub_edit| -> Result<bool> {
                    if conflict {
                        return Ok(true);
                    }
                    let effects = actual_durable_effect_keys(sub_edit.payload())?;
                    Ok(effects.into_iter().any(|effect| {
                        if effect == "package" {
                            return staged_package_effect || !self.changes.is_empty();
                        }
                        if staged_package_effect {
                            return true;
                        }
                        self.changes
                            .iter()
                            .flat_map(|change| change_effect_keys(&self.source, change))
                            .any(|staged| staged == effect)
                    }))
                })?;
        if direct_conflict {
            return Ok(true);
        }
        projected_staged_style_conflict(self, joined)
    }

    fn validate_joined_page_names(&self, joined: &JoinedEdits) -> Result<()> {
        // A replacement changes the page universe, so validating any later
        // selector against the source would be both incorrect and an extra
        // package reopen. The execution segments validate against the active
        // replacement exactly once before staging their operations.
        if joined_contains_package_replacement(joined) {
            return Ok(());
        }
        let mut segment = Vec::new();
        let mut active_replacement = None;
        for sub_edit in joined.sub_edits() {
            for operation in sub_edit.payload().operations() {
                if operation.op == "package.replace" {
                    self.validate_page_name_segment(active_replacement.as_ref(), &segment)?;
                    segment.clear();
                    let blob = durable_blob(sub_edit.payload(), &operation.value)?;
                    let base = active_replacement.as_ref().unwrap_or(&self.source);
                    active_replacement = Some(base.reopen_same_flavor(blob.to_vec())?);
                } else {
                    segment.push(operation);
                }
            }
        }
        self.validate_page_name_segment(active_replacement.as_ref(), &segment)
    }

    fn validate_page_name_segment(
        &self,
        active_replacement: Option<&Package>,
        operations: &[&PatchOperation],
    ) -> Result<()> {
        if let Some(replacement) = active_replacement {
            let edit =
                replacement.edit_with_policies(self.security_policy, self.active_content_policy);
            edit.page_name_targets_for_operations(operations.iter().copied())?;
        } else {
            self.page_name_targets_for_operations(operations.iter().copied())?;
        }
        Ok(())
    }

    fn page_name_targets_for_operations<'a>(
        &self,
        operations: impl Iterator<Item = &'a PatchOperation>,
    ) -> Result<BTreeMap<usize, String>> {
        let mut operations = operations
            .filter(|operation| operation.op == "page.name.set")
            .peekable();
        if operations.peek().is_none() {
            return Ok(BTreeMap::new());
        }
        let parsed = parse_content(&self.content)?;
        let mut names = parsed
            .pages
            .iter()
            .map(|page| page.name().map(str::to_owned))
            .collect::<Vec<_>>();
        let mut targets = BTreeMap::new();
        for operation in operations {
            let page = parse_page_target(&operation.target)?;
            let after = string_value(operation)?.to_owned();
            validate_bounded_value(&after, "ODG page name")?;
            let before = names.get(page).and_then(Option::as_deref).ok_or_else(|| {
                Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.edit.join",
                    "ODG page rename requires an existing draw:name",
                ))
            })?;
            if before != after.as_str()
                && xml_has_attribute(&self.content, DRAW, b"page-name", before)?
            {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.edit.join",
                    "ODG page rename is blocked by a draw:page-name dependency",
                )));
            }
            names[page] = Some(after.clone());
            targets.insert(page, after);
        }
        for (page, name) in &targets {
            if names
                .iter()
                .enumerate()
                .any(|(other, value)| other != *page && value.as_deref() == Some(name.as_str()))
            {
                return Err(joined_refusal("ODG joined page names would duplicate"));
            }
        }
        Ok(targets)
    }

    fn replace_or_insert_shape_attribute(
        &mut self,
        page: usize,
        shape: usize,
        span: Option<Range<usize>>,
        qualified_name: &str,
        value: &str,
    ) -> Result<()> {
        if let Some(attribute_span) = span {
            return self.replace_content_value(&attribute_span, value);
        }
        let parsed = parse_content(&self.content)?;
        let shape_span = parsed
            .shape_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape source span is missing".into()))?;
        let tag_end = start_tag_end(&self.content, shape_span.start)?;
        let insertion = if self.content.as_bytes().get(tag_end.saturating_sub(2)) == Some(&b'/') {
            tag_end.saturating_sub(2)
        } else {
            tag_end.saturating_sub(1)
        };
        let mut attribute = String::new();
        push_attribute(&mut attribute, qualified_name, Some(value))?;
        self.invalidate_content_splices();
        self.content
            .set_owned(insert_xml(&self.content, insertion, &attribute)?);
        Ok(())
    }

    fn replace_content_value(&mut self, span: &Range<usize>, replacement: &str) -> Result<()> {
        let escaped = escaped_xml_string(replacement, MAX_OUTPUT_BYTES, "ODG edited content")?;
        let next = replace_xml_value(&self.content, span, replacement)?;
        if let Some(splices) = &mut self.content_splices {
            stage_content_splice(
                self.source.content_xml().as_bytes(),
                self.content.as_bytes(),
                splices,
                span,
                escaped.as_bytes(),
            )?;
        }
        self.content.set_owned(next);
        Ok(())
    }

    fn replace_content_values(
        &mut self,
        spans: &[&Range<usize>],
        replacements: &[String; 4],
    ) -> Result<()> {
        let mut edits = spans
            .iter()
            .zip(replacements)
            .map(|(span, value)| ((*span).clone(), value.as_str()))
            .collect::<Vec<_>>();
        edits.sort_unstable_by_key(|(span, _)| std::cmp::Reverse(span.start));
        for (span, replacement) in edits {
            self.replace_content_value(&span, replacement)?;
        }
        Ok(())
    }

    fn invalidate_content_splices(&mut self) {
        self.content_splices = None;
    }

    /// Replace `range` of the staged content with balanced `replacement`
    /// markup (or delete it), keeping the edit on the exact-source splice
    /// path when the range maps onto source bytes no earlier edit touched.
    /// Otherwise the edit falls back to a whole-part rewrite, which the
    /// commit admits only for compact sources (record 0288).
    fn splice_content(&mut self, range: &Range<usize>, replacement: &str) -> Result<()> {
        let next = replace_xml(&self.content, range, replacement)?;
        if let Some(splices) = &mut self.content_splices
            && stage_content_splice(
                self.source.content_xml().as_bytes(),
                self.content.as_bytes(),
                splices,
                range,
                replacement.as_bytes(),
            )
            .is_err()
        {
            self.content_splices = None;
        }
        self.content.set_owned(next);
        Ok(())
    }

    /// Insert `child` markup at `at` in the staged content (expanding an
    /// empty owner element whose `/>` starts at `at`).
    fn insert_content_child(&mut self, at: usize, child: &str) -> Result<()> {
        let (range, replacement) = child_insertion(&self.content, at, child)?;
        self.splice_content(&range, &replacement)
    }

    fn require_group_descendant(&self, page: usize, group: usize, descendant: usize) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if !group_selection(&parsed, page, group)?.contains(descendant) {
            return invalid("ODG shape is not owned by the selected group subtree");
        }
        Ok(())
    }

    /// Replaces geometry on one checked descendant owned by a group subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid ownership or the same geometry errors as
    /// [`Self::set_shape_geometry`].
    pub fn set_group_descendant_geometry(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_geometry(page, descendant, x, y, width, height)
    }

    /// Replaces text on one checked descendant owned by a group subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid ownership or a non-lossless text owner.
    pub fn set_group_descendant_text(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_text(page, descendant, text)
    }

    /// Assigns one style to every losslessly addressable descendant style owner.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership, undeclared style, no style owners, or limits.
    pub fn set_group_style_name(
        &mut self,
        page: usize,
        group: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let requested_style = style_name.into();
        validate_bounded_value(&requested_style, "ODG group style name")?;
        let declared_in_content = declares_style_xml(&self.content, &requested_style)?;
        let declared_in_styles = self
            .source
            .styles_xml()
            .map(|styles| declares_style_xml(styles, &requested_style))
            .transpose()?
            .unwrap_or_default();
        if !declared_in_content && !declared_in_styles {
            return invalid("ODG group destination style is not declared");
        }
        let parsed = parse_content(&self.content)?;
        let selection = group_selection(&parsed, page, group)?;
        let targets = selection
            .descendants()
            .iter()
            .copied()
            .filter(|position| parsed.style_name_spans[page][*position].is_some())
            .collect::<Vec<_>>();
        if targets.is_empty() || targets.len() > MAX_GROUP_EDITS {
            return invalid("ODG group style owner count is unsupported");
        }
        for target in targets {
            self.set_shape_style_name(page, target, requested_style.clone())?;
        }
        Ok(())
    }

    /// Replaces every single-span descendant text owner atomically after complete preflight.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership, no lossless text owners, or limits.
    pub fn set_group_text(
        &mut self,
        page: usize,
        group: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let replacement_text = text.into();
        if replacement_text.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement group text exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selection = group_selection(&parsed, page, group)?;
        let targets = selection
            .descendants()
            .iter()
            .copied()
            .filter(|position| matches!(parsed.text_spans[page][*position].as_slice(), [Some(_)]))
            .collect::<Vec<_>>();
        if targets.is_empty() || targets.len() > MAX_GROUP_EDITS {
            return invalid("ODG group text owner count is unsupported");
        }
        for target in targets {
            self.set_shape_text(page, target, replacement_text.clone())?;
        }
        Ok(())
    }

    /// Changes an inert form reference on one checked control descendant.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership or control-reference semantics.
    pub fn set_group_descendant_control_reference(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        reference: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_control_reference(page, descendant, reference)
    }

    /// Renames a page through its existing `draw:name` attribute.
    ///
    /// Page-name references elsewhere in the drawing are dependency checked and cause refusal;
    /// callers must update those owners explicitly rather than leaving dangling references.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, duplicate name, referenced old name, absent
    /// source attribute, or size limit.
    pub fn set_page_name(&mut self, page: usize, name: impl Into<String>) -> Result<()> {
        let after = name.into();
        validate_bounded_value(&after, "ODG page name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let before = selected.name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page rename requires an existing draw:name",
            ))
        })?;
        if parsed
            .pages
            .iter()
            .enumerate()
            .any(|(index, value)| index != page && value.name() == Some(after.as_str()))
        {
            return invalid("ODG page rename would create a duplicate name");
        }
        if before == after {
            return Ok(());
        }
        if xml_has_attribute(&self.content, DRAW, b"page-name", before)? {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page rename is blocked by a draw:page-name dependency",
            )));
        }
        let span = parsed.page_attribute_spans[page][0]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page name source span is missing".into()))?;
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::PageName(PageNameChange {
            page,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes a page's existing drawing-page style reference.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, missing source/style declarations, or limits.
    pub fn set_page_style_name(
        &mut self,
        page: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let after = style_name.into();
        validate_bounded_value(&after, "ODG page style name")?;
        if !declares_style(&self.source, &after)? {
            return invalid("ODG destination page style is not declared");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let before = selected.style_name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page style edit requires an existing draw:style-name",
            ))
        })?;
        if before == after {
            return Ok(());
        }
        let span = parsed.page_attribute_spans[page][1]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page style source span is missing".into()))?;
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::PageStyle(PageStyleChange {
            page,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Sets or clears inert transition metadata on the page's drawing-page style.
    ///
    /// Existing style ownership is respected: automatic styles are edited in
    /// content.xml, while a named style owned by styles.xml is edited in that
    /// part. Unknown style attributes, child elements, namespace choices, and
    /// surrounding producer bytes remain intact.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing page/style owner, an ambiguous style,
    /// malformed XML, or a bounded value violation.
    pub fn set_page_transition(
        &mut self,
        page: usize,
        transition: Option<Transition>,
    ) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let style_name = selected.style_name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page transition requires an existing draw:style-name",
            ))
        })?;
        let current = parse_page_transitions(&self.content, self.styles.as_deref(), &parsed.pages)?
            .get(page)
            .cloned()
            .flatten();
        let desired = transition.filter(|value| !value.is_empty());
        if current == desired {
            return Ok(());
        }
        if transition_style_is_shared(
            &self.content,
            self.styles.as_deref(),
            &parsed.pages,
            page,
            style_name,
        )? {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page transition style is shared by multiple pages",
            )));
        }
        let in_content = parse_style_definitions(&self.content)?
            .iter()
            .filter(|definition| {
                definition.style.name() == style_name && definition.style.family() == "drawing-page"
            })
            .count();
        if in_content > 1 {
            return invalid("ODG page transition style is ambiguous");
        }
        validate_transition_budget(desired.as_ref())?;
        validate_transition_sound_reference(self, desired.as_ref())?;
        validate_transition_xml_ids(self, current.as_ref(), desired.as_ref(), in_content == 1)?;
        if desired.is_none() && current.is_some() {
            let parent = if in_content == 1 {
                parse_style_definitions(&self.content)?
                    .into_iter()
                    .find(|definition| {
                        definition.style.name() == style_name
                            && definition.style.family() == "drawing-page"
                    })
                    .and_then(|definition| definition.style.parent().map(str::to_owned))
            } else {
                self.styles
                    .as_deref()
                    .map(parse_style_definitions)
                    .transpose()?
                    .and_then(|definitions| {
                        definitions
                            .into_iter()
                            .find(|definition| {
                                definition.style.name() == style_name
                                    && definition.style.family() == "drawing-page"
                            })
                            .and_then(|definition| definition.style.parent().map(str::to_owned))
                    })
            };
            if parent.is_some() {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG inherited page transition cannot be cleared without changing its parent style",
                )));
            }
        }
        if in_content == 1 {
            // Borrow the current owner instead of cloning it: the edited copy
            // is the only new allocation, and the reversibility probe reads
            // both strings in place.
            let after_xml = edit_transition_style_xml(&self.content, style_name, desired.as_ref())?;
            if !transition_edit_is_lexically_reversible(
                &self.content,
                &after_xml,
                style_name,
                current.as_ref(),
            ) {
                self.requires_package_projection = true;
            }
            self.content.set_owned(after_xml);
            self.invalidate_content_splices();
        } else if let Some(styles) = &mut self.styles {
            // styles.xml is handed to PackageWriter as a replacement when it
            // changes.  Keep the same gate used by the package rebuild path
            // here, before editing, so a noncompact producer file is refused
            // atomically instead of being changed and failing publication
            // later.
            compact_xml::validate(styles.as_bytes(), compact_xml::Limits::default())
                .map_err(Error::from)?;
            let matches = parse_style_definitions(styles)?
                .iter()
                .filter(|definition| {
                    definition.style.name() == style_name
                        && definition.style.family() == "drawing-page"
                })
                .count();
            if matches != 1 {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG page transition style is not uniquely owned",
                )));
            }
            let after_xml = edit_transition_style_xml(styles, style_name, desired.as_ref())?;
            let requires_projection = !transition_edit_is_lexically_reversible(
                styles,
                &after_xml,
                style_name,
                current.as_ref(),
            );
            styles.set_owned(after_xml);
            if requires_projection {
                self.requires_package_projection = true;
            }
        } else {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page transition style is not declared",
            )));
        }
        self.changes
            .push(Change::PageTransition(Box::new(PageTransitionChange {
                page,
                before: current,
                after: desired,
            })));
        Ok(())
    }

    /// Replaces one shape's sole plain paragraph character-data span.
    ///
    /// Split, mixed, CDATA, and entity-reference text is refused rather than
    /// serialized through a lossy XML model. A transaction owns one edit;
    /// restaging the same selector replaces its pending value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a shape whose text is not one
    /// plain span (several paragraphs, spans, links or fields), and
    /// [`Error::InvalidFormat`] for an invalid selector or a limit violation.
    pub fn set_shape_text(
        &mut self,
        page: usize,
        shape: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape text exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let spans = parsed
            .text_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape source span is missing".to_string()))?;
        if !matches!(spans.as_slice(), [Some(_)]) {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG shape text is not one losslessly replaceable XML span",
            )));
        }
        if selected.text() == after {
            return Ok(());
        }
        let span = spans[0].as_ref().ok_or_else(|| {
            Error::InvalidFormat("ODG shape text source span is missing".to_string())
        })?;
        let before = selected.text().to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Text(TextChange {
            page,
            shape,
            before,
            after,
        }));
        Ok(())
    }

    /// Renames one shape through its existing `draw:name` attribute.
    ///
    /// ODF 1.4 Part 3 §19.197 defines `draw:name` as the reference name for
    /// graphical elements. This preserves the original start tag and attribute
    /// spelling, replacing only the validated attribute-value span.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-bounds selector, an unnamed shape, a
    /// name over the bounded size, or a shape whose source attribute cannot be
    /// losslessly addressed.
    pub fn set_shape_name(
        &mut self,
        page: usize,
        shape: usize,
        name: impl Into<String>,
    ) -> Result<()> {
        let after = name.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape name exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let before = selected.name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG shape rename requires an existing losslessly addressable draw:name",
            ))
        })?;
        let span = parsed
            .name_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape name source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Name(NameChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes a shape's existing layer assignment without normalizing its tag.
    ///
    /// ODF 1.4 Part 3 §§10.2.2-10.2.3 and 19.189 define drawing layers and
    /// their shape assignment. The destination must be one of the declarations
    /// visible through [`Package::layers`].
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, undeclared layer, absent source attribute, or
    /// limit violation.
    pub fn set_shape_layer(
        &mut self,
        page: usize,
        shape: usize,
        layer: impl Into<String>,
    ) -> Result<()> {
        let after = layer.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement layer name exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG page selector is out of bounds".to_string())
        })?;
        let selected = selected_page.shapes().get(shape).ok_or_else(|| {
            Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
        })?;
        let visible_layers = if selected_page.has_layer_set() {
            selected_page.layers()
        } else {
            self.source.layers()
        };
        if !visible_layers
            .iter()
            .any(|declared_layer| declared_layer.name() == after)
        {
            return invalid("ODG destination layer is not declared");
        }
        let before = selected.layer().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG layer change requires an existing losslessly addressable draw:layer",
            ))
        })?;
        let span = parsed
            .layer_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape layer source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Layer(LayerChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Replaces all four existing SVG geometry attributes as one operation.
    ///
    /// # Errors
    ///
    /// Returns an error if the checked selectors fail or the shape does not
    /// own four losslessly addressable geometry attributes.
    pub fn set_shape_geometry(
        &mut self,
        page: usize,
        shape: usize,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Result<()> {
        let after = [x.into(), y.into(), width.into(), height.into()];
        validate_geometry(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = [
            selected.x(),
            selected.y(),
            selected.width(),
            selected.height(),
        ]
        .map(|value| value.map(str::to_owned));
        let spans = parsed
            .geometry_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape geometry spans are missing".into()))?;
        let ranges = spans
            .iter()
            .map(|span| {
                span.as_ref().ok_or_else(|| {
                    Error::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "odg.package.snapshot",
                        "ODG geometry edit requires existing x, y, width, and height attributes",
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if before
            .iter()
            .zip(&after)
            .all(|(source_value, target_value)| {
                source_value.as_deref() == Some(target_value.as_str())
            })
        {
            return Ok(());
        }
        self.replace_content_values(&ranges, &after)?;
        self.changes.push(Change::Geometry(GeometryChange {
            page,
            shape,
            before: before.map(Option::unwrap_or_default),
            after,
        }));
        Ok(())
    }

    /// Sets or inserts a lexical `draw:transform` on one checked shape.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, lexical controls, provenance, or output limits.
    pub fn set_shape_transform(
        &mut self,
        page: usize,
        shape: usize,
        transform: impl Into<String>,
    ) -> Result<()> {
        let after = transform.into();
        validate_advanced_geometry_value(&after, "ODG transform")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if selected.kind().is_three_dimensional() {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG dr3d transforms are inert and cannot be edited through draw:transform",
            )));
        }
        if selected.transform() == Some(after.as_str()) {
            return Ok(());
        }
        let span = parsed.transform_spans[page][shape].clone();
        self.replace_or_insert_shape_attribute(page, shape, span, "draw:transform", &after)?;
        self.changes.push(Change::Structure(
            StructureChange::ShapeAdvancedGeometryChanged { page, shape },
        ));
        Ok(())
    }

    /// Sets or inserts a polygon/polyline view box and lexical point list atomically.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selector owns a polygon or polyline and values are bounded.
    pub fn set_shape_points(
        &mut self,
        page: usize,
        shape: usize,
        view_box: impl Into<String>,
        points: impl Into<String>,
    ) -> Result<()> {
        let target_view_box = view_box.into();
        let target_points = points.into();
        validate_advanced_geometry_value(&target_view_box, "ODG polygon view box")?;
        validate_advanced_geometry_value(&target_points, "ODG polygon points")?;
        if !is_integer_list(&target_view_box, 4) {
            return invalid("ODG polygon view box is not four integers");
        }
        if !is_points(&target_points) {
            return invalid("ODG polygon points are not an ODF point list");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if !matches!(selected.kind(), ShapeKind::Polygon | ShapeKind::Polyline) {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG point-list edit requires a polygon or polyline",
            )));
        }
        if selected.view_box() == Some(target_view_box.as_str())
            && selected.points() == Some(target_points.as_str())
        {
            return Ok(());
        }
        let view_box_span = parsed.points_spans[page][shape][0].clone();
        self.replace_or_insert_shape_attribute(
            page,
            shape,
            view_box_span,
            "svg:viewBox",
            &target_view_box,
        )?;
        let reparsed = parse_content(&self.content)?;
        let points_span = reparsed.points_spans[page][shape][1].clone();
        self.replace_or_insert_shape_attribute(
            page,
            shape,
            points_span,
            "draw:points",
            &target_points,
        )?;
        self.changes.push(Change::Structure(
            StructureChange::ShapeAdvancedGeometryChanged { page, shape },
        ));
        Ok(())
    }

    /// Sets or inserts lexical line endpoints on a line, connector, or measure shape.
    ///
    /// # Errors
    ///
    /// Returns an error for an incompatible kind, selector failure, or invalid lexical values.
    pub fn set_shape_line_geometry(
        &mut self,
        page: usize,
        shape: usize,
        x1: impl Into<String>,
        y1: impl Into<String>,
        x2: impl Into<String>,
        y2: impl Into<String>,
    ) -> Result<()> {
        let after = [x1.into(), y1.into(), x2.into(), y2.into()];
        validate_geometry(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if !matches!(
            selected.kind(),
            ShapeKind::Line | ShapeKind::Connector | ShapeKind::Measure
        ) {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG line-endpoint edit requires a line, connector, or measure",
            )));
        }
        let before = selected.line_geometry();
        if before
            .iter()
            .zip(&after)
            .all(|(source_value, target_value)| *source_value == Some(target_value.as_str()))
        {
            return Ok(());
        }
        for (index, (qualified_name, value)) in ["svg:x1", "svg:y1", "svg:x2", "svg:y2"]
            .into_iter()
            .zip(&after)
            .enumerate()
        {
            let current = parse_content(&self.content)?;
            let attribute_span = current.line_geometry_spans[page][shape][index].clone();
            self.replace_or_insert_shape_attribute(
                page,
                shape,
                attribute_span,
                qualified_name,
                value,
            )?;
        }
        self.changes.push(Change::Structure(
            StructureChange::ShapeAdvancedGeometryChanged { page, shape },
        ));
        Ok(())
    }

    /// Changes an existing graphic style reference without normalizing XML.
    ///
    /// # Errors
    ///
    /// Returns an error for a checked-selector failure or missing source attribute.
    pub fn set_shape_style_name(
        &mut self,
        page: usize,
        shape: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let after = style_name.into();
        validate_bounded_value(&after, "ODG shape style name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = selected.style_name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG style edit requires an existing draw:style-name",
            ))
        })?;
        let span = parsed
            .style_name_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape style span is missing".into()))?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Style(StyleChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes an existing SVG path-data attribute without normalizing XML.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selected shape is a path with an existing,
    /// losslessly addressable `svg:d` attribute.
    pub fn set_shape_path_data(
        &mut self,
        page: usize,
        shape: usize,
        path_data: impl Into<String>,
    ) -> Result<()> {
        let after = path_data.into();
        validate_path_data(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if selected.kind() != ShapeKind::Path {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG path-data edit requires a draw:path shape",
            )));
        }
        let before = selected.path_data().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG path-data edit requires an existing svg:d",
            ))
        })?;
        let span = parsed
            .path_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape path-data span is missing".into()))?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Path(PathChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes an existing inert form-control reference without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selected control shape has an existing,
    /// losslessly addressable `draw:control` attribute.
    pub fn set_shape_control_reference(
        &mut self,
        page: usize,
        shape: usize,
        reference: impl Into<String>,
    ) -> Result<()> {
        let after = reference.into();
        validate_bounded_value(&after, "ODG form-control reference")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if selected.kind() != ShapeKind::Control {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG form-control edit requires a draw:control shape",
            )));
        }
        let before = selected.control_reference().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG form-control edit requires an existing draw:control",
            ))
        })?;
        let span = parsed
            .control_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG form-control source span is missing".into())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes
            .push(Change::ControlReference(ControlReferenceChange {
                page,
                shape,
                before: before_owned,
                after,
            }));
        Ok(())
    }

    /// Inserts a detached page at a checked source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid position, duplicate page identity, or limit violation.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached page values transfer ownership into the transaction"
    )]
    pub fn insert_page(&mut self, position: usize, page: Page) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if position > parsed.pages.len() {
            return invalid("ODG page insertion position is out of bounds");
        }
        if parsed.pages.len() >= MAX_PAGES {
            return invalid("ODG page count exceeds the limit");
        }
        if let Some(name) = page.name()
            && parsed.pages.iter().any(|value| value.name() == Some(name))
        {
            return invalid("ODG inserted page name is already present");
        }
        let mut page = page;
        let transition = page.transition().filter(|value| !value.is_empty()).cloned();
        let generated_style = if let Some(transition) = transition.as_ref() {
            validate_transition_budget(Some(transition))?;
            validate_transition_sound_reference(self, Some(transition))?;
            validate_transition_xml_ids(self, None, Some(transition), true)?;
            let name = unique_page_transition_style_name(
                &self.content,
                self.styles.as_deref(),
                &parsed.pages,
            )?;
            page.set_style_name(Some(name.clone()));
            Some((
                name.clone(),
                serialize_detached_page_style(&name, transition)?,
            ))
        } else {
            None
        };
        let styled_content: Cow<'_, str> =
            if let Some((_name, style_xml)) = generated_style.as_ref() {
                Cow::Owned(insert_automatic_style(&self.content, style_xml)?)
            } else {
                Cow::Borrowed(self.content.as_str())
            };
        let styled_parsed = parse_content(&styled_content)?;
        let at = if position == styled_parsed.pages.len() {
            styled_parsed.drawing_insert_position
        } else {
            styled_parsed.page_spans[position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?
                .start
        };
        let xml = serialize_page(&page)?;
        if generated_style.is_some() {
            let content = insert_child_xml(&styled_content, at, &xml)?;
            self.invalidate_content_splices();
            self.content.set_owned(content);
        } else {
            self.insert_content_child(at, &xml)?;
        }
        if let Some((name, _style_xml)) = generated_style {
            self.changes
                .push(Change::Structure(StructureChange::StyleInserted { name }));
        }
        self.changes
            .push(Change::Structure(StructureChange::PageInserted {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(())
    }

    /// Appends a detached page.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate identity or a resource limit.
    pub fn add_page(&mut self, page: Page) -> Result<()> {
        let position = parse_content(&self.content)?.pages.len();
        self.insert_page(position, page)
    }

    /// Removes one page selected by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is absent, ambiguous, or unaddressable.
    pub fn remove_page<'selector>(
        &mut self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Page> {
        let parsed = parse_content(&self.content)?;
        let position = resolve_page_position(&parsed.pages, selector.into())?;
        let page = parsed.pages[position].clone();
        let span = parsed.page_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        if let Some(name) = page.name()
            && xml_has_attribute(&content, DRAW, b"page-name", name)?
        {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page removal is blocked by a draw:page-name dependency",
            )));
        }
        let span = span.clone();
        self.splice_content(&span, "")?;
        self.changes
            .push(Change::Structure(StructureChange::PageRemoved {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(page)
    }

    /// Moves the page at `from` so it ends at position `to` of the page
    /// order. The page's bytes move unchanged; on a producer drawing both
    /// edits stay exact-source splices.
    ///
    /// # Errors
    ///
    /// Returns an error for a position out of bounds or a missing span.
    pub fn move_page(&mut self, from: usize, to: usize) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if from >= parsed.pages.len() || to >= parsed.pages.len() {
            return invalid("ODG page move position is out of bounds");
        }
        if from == to {
            return Ok(());
        }
        let span = |index: usize| {
            parsed.page_spans[index]
                .clone()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))
        };
        let moved = span(from)?;
        let at = if to > from {
            span(to)?.end
        } else {
            span(to)?.start
        };
        self.relocate_content(moved, at)?;
        self.changes
            .push(Change::Structure(StructureChange::PageMoved { from, to }));
        Ok(())
    }

    /// Moves the shape at `from` of `page` so it ends at position `to` of the
    /// page's shape order.
    ///
    /// # Errors
    ///
    /// Returns an error for a selector out of bounds or a missing span.
    pub fn move_shape(&mut self, page: usize, from: usize, to: usize) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let spans = parsed
            .shape_spans
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if from >= spans.len() || to >= spans.len() {
            return invalid("ODG shape move position is out of bounds");
        }
        if from == to {
            return Ok(());
        }
        let span = |index: usize| {
            spans[index]
                .clone()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))
        };
        let moved = span(from)?;
        let at = if to > from {
            span(to)?.end
        } else {
            span(to)?.start
        };
        self.relocate_content(moved, at)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeMoved {
                page,
                from,
                to,
            }));
        Ok(())
    }

    /// Moves the top-level shape at `from` of `page` into page `to_page`, so
    /// it ends at position `to` of that page's shape order. The shape's bytes
    /// move unchanged; on a producer drawing both edits stay exact-source
    /// splices, and an empty destination page element is expanded.
    ///
    /// A shape that is part of a group, an anchor inside a group, and any
    /// identity dependency that would cross pages (an `xml:id`/`draw:id` of
    /// the moved subtree referenced from outside it, or a connector in the
    /// subtree glued to a shape outside it) are refused rather than left
    /// dangling.
    ///
    /// # Errors
    ///
    /// Returns an error for a selector out of bounds, a missing span, a
    /// group member, or a cross-page dependency.
    pub fn move_shape_to_page(
        &mut self,
        page: usize,
        from: usize,
        to_page: usize,
        to: usize,
    ) -> Result<()> {
        if page == to_page {
            let count = parse_content(&self.content)?
                .shape_spans
                .get(page)
                .map_or(0, Vec::len);
            if to >= count {
                return invalid("ODG shape move position is out of bounds");
            }
            return self.move_shape(page, from, to);
        }
        let parsed = parse_content(&self.content)?;
        let spans = parsed
            .shape_spans
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let targets = parsed.shape_spans.get(to_page).ok_or_else(|| {
            Error::InvalidFormat("ODG destination page selector is out of bounds".into())
        })?;
        let moved = spans
            .get(from)
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?
            .clone()
            .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?;
        let nested = |spans: &[Option<Range<usize>>], span: &Range<usize>| {
            spans
                .iter()
                .flatten()
                .any(|outer| outer != span && outer.start <= span.start && span.end <= outer.end)
        };
        if nested(spans, &moved) {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG group members move with their group",
            )));
        }
        let at = match targets.get(to) {
            Some(anchor) => {
                let anchor = anchor
                    .as_ref()
                    .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?;
                if nested(targets, anchor) {
                    return Err(Error::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "odg.package.snapshot",
                        "ODG shape move position lies inside a group",
                    )));
                }
                anchor.start
            },
            None if to == targets.len() => {
                parsed.page_insert_positions[to_page].ok_or_else(|| {
                    Error::InvalidFormat("ODG page insertion point is missing".into())
                })?
            },
            None => return invalid("ODG shape move position is out of bounds"),
        };
        check_shape_identity(&self.content, &moved)?;
        let bytes = self
            .content
            .get(moved.clone())
            .ok_or_else(|| Error::InvalidFormat("ODG relocated span is invalid".into()))?
            .to_owned();
        if at >= moved.end {
            self.insert_content_child(at, &bytes)?;
            self.splice_content(&moved, "")?;
        } else if at <= moved.start {
            self.splice_content(&moved, "")?;
            self.insert_content_child(at, &bytes)?;
        } else {
            return invalid("ODG relocation target lies inside the moved element");
        }
        self.changes
            .push(Change::Structure(StructureChange::ShapeTransferred {
                page,
                from,
                to_page,
                to,
            }));
        Ok(())
    }

    /// Move the staged bytes of `span` to the insertion point `at` (outside
    /// the span), as a removal and an insertion.
    fn relocate_content(&mut self, span: Range<usize>, at: usize) -> Result<()> {
        let bytes = self
            .content
            .get(span.clone())
            .ok_or_else(|| Error::InvalidFormat("ODG relocated span is invalid".into()))?
            .to_owned();
        if at >= span.end {
            self.splice_content(&(at..at), &bytes)?;
            self.splice_content(&span, "")
        } else if at <= span.start {
            self.splice_content(&span, "")?;
            self.splice_content(&(at..at), &bytes)
        } else {
            invalid("ODG relocation target lies inside the moved element")
        }
    }

    /// Inserts a detached shape at a checked page shape position.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, undeclared layers, inert-only
    /// shape kinds, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached shape values transfer ownership into the transaction"
    )]
    pub fn insert_shape(&mut self, page: usize, position: usize, shape: Shape) -> Result<()> {
        if shape.kind().is_three_dimensional() {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG 3D shapes are inert read-only owners",
            )));
        }
        if shape.source_backed() {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "odg.package.snapshot",
                "ODG parsed shapes require prepare_shape_transfer for source-preserving insertion",
            )));
        }
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if position > selected_page.shapes().len() {
            return invalid("ODG shape insertion position is out of bounds");
        }
        validate_shape_layer(selected_page, self.source.layers(), &shape)?;
        if parsed
            .pages
            .iter()
            .map(|value| value.shapes().len())
            .sum::<usize>()
            >= MAX_SHAPES
        {
            return invalid("ODG shape count exceeds the limit");
        }
        let at = if position == selected_page.shapes().len() {
            parsed.page_insert_positions[page]
                .ok_or_else(|| Error::InvalidFormat("ODG page insertion point is missing".into()))?
        } else {
            parsed.shape_spans[page][position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?
                .start
        };
        let xml = serialize_shape(&shape)?;
        self.insert_content_child(at, &xml)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeInserted {
                page,
                position,
                kind: shape.kind(),
            }));
        Ok(())
    }

    /// Appends a detached shape to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, dependencies, or limits.
    pub fn add_shape(&mut self, page: usize, shape: Shape) -> Result<()> {
        let position = parse_content(&self.content)?
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?
            .shapes()
            .len();
        self.insert_shape(page, position, shape)
    }

    /// Adds or replaces one bounded inert automatic style definition.
    ///
    /// Arbitrary qualified property attributes are retained as data and never evaluated.
    /// Definitions owned by `styles.xml` are not silently shadowed.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid properties, ambiguous names, external ownership, or limits.
    pub fn put_style(&mut self, style: &Style) -> Result<()> {
        if let Some(parent) = style.parent() {
            let declared_in_content = declares_style_xml(&self.content, parent)?;
            let declared_in_styles = self
                .source
                .styles_xml()
                .map(|styles| declares_style_xml(styles, parent))
                .transpose()?
                .unwrap_or_default();
            if !declared_in_content && !declared_in_styles {
                return invalid("ODG parent style is not declared");
            }
        }
        let xml = serialize_style(style)?;
        let matches = parse_style_definitions(&self.content)?
            .into_iter()
            .filter(|definition| definition.style.name() == style.name())
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return invalid("ODG automatic style name is ambiguous");
        }
        let externally_owned = self
            .source
            .styles_xml()
            .map(|styles| declares_style_xml(styles, style.name()))
            .transpose()?
            .unwrap_or_default();
        let (content, change) = if let Some(existing) = matches.first() {
            if existing.style == *style {
                return Ok(());
            }
            (
                replace_xml(&self.content, &existing.span, &xml)?,
                StructureChange::StyleReplaced {
                    name: style.name().to_owned(),
                },
            )
        } else {
            if externally_owned {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG style owned by styles.xml cannot be shadowed",
                )));
            }
            (
                insert_automatic_style(&self.content, &xml)?,
                StructureChange::StyleInserted {
                    name: style.name().to_owned(),
                },
            )
        };
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_style_definitions(&self.content)?;
        self.changes.push(Change::Structure(change));
        Ok(())
    }

    /// Removes one content-owned automatic style after checking all known references.
    ///
    /// # Errors
    ///
    /// Returns an error for missing/ambiguous/external ownership or a live style dependency.
    pub fn remove_style(&mut self, name: &str) -> Result<Style> {
        let matches = parse_style_definitions(&self.content)?
            .into_iter()
            .filter(|definition| definition.style.name() == name)
            .collect::<Vec<_>>();
        let [existing] = matches.as_slice() else {
            return invalid("ODG content style selector is missing or ambiguous");
        };
        if xml_has_attribute(&self.content, DRAW, b"style-name", name)?
            || xml_has_attribute(&self.content, DRAW, b"text-style-name", name)?
            || xml_has_attribute(&self.content, STYLE, b"parent-style-name", name)?
        {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG style removal is blocked by a live dependency",
            )));
        }
        let content = remove_xml(&self.content, &existing.span)?;
        let removed = existing.style.clone();
        self.invalidate_content_splices();
        self.content.set_owned(content);
        self.changes
            .push(Change::Structure(StructureChange::StyleRemoved {
                name: name.to_owned(),
            }));
        Ok(removed)
    }

    /// Adds or replaces one content-owned named drawing resource.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid attributes, ambiguous names, external ownership, or limits.
    pub fn put_style_resource(&mut self, resource: &StyleResource) -> Result<()> {
        let xml = serialize_style_resource(resource)?;
        let matches = parse_style_resources(&self.content)?
            .into_iter()
            .filter(|definition| {
                definition.resource.kind() == resource.kind()
                    && definition.resource.name() == resource.name()
            })
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return invalid("ODG named style-resource selector is ambiguous");
        }
        let externally_owned = self
            .source
            .styles_xml()
            .map(|styles| declares_style_resource_xml(styles, resource.kind(), resource.name()))
            .transpose()?
            .unwrap_or_default();
        let (content, change) = if let Some(existing) = matches.first() {
            if existing.resource == *resource {
                return Ok(());
            }
            (
                replace_xml(&self.content, &existing.span, &xml)?,
                StructureChange::StyleResourceReplaced {
                    kind: resource.kind(),
                    name: resource.name().to_owned(),
                },
            )
        } else {
            if externally_owned {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG named resource owned by styles.xml cannot be shadowed",
                )));
            }
            (
                insert_automatic_style(&self.content, &xml)?,
                StructureChange::StyleResourceInserted {
                    kind: resource.kind(),
                    name: resource.name().to_owned(),
                },
            )
        };
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_style_resources(&self.content)?;
        self.changes.push(Change::Structure(change));
        Ok(())
    }

    /// Removes one content-owned named drawing resource after dependency checks.
    ///
    /// # Errors
    ///
    /// Returns an error for missing/ambiguous ownership or a live style dependency.
    pub fn remove_style_resource(
        &mut self,
        kind: StyleResourceKind,
        name: &str,
    ) -> Result<StyleResource> {
        let matches = parse_style_resources(&self.content)?
            .into_iter()
            .filter(|definition| {
                definition.resource.kind() == kind && definition.resource.name() == name
            })
            .collect::<Vec<_>>();
        let [existing] = matches.as_slice() else {
            return invalid("ODG content named style-resource selector is missing or ambiguous");
        };
        for reference in style_resource_reference_locals(kind) {
            if xml_has_attribute(&self.content, DRAW, reference, name)? {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG named style-resource removal is blocked by a live dependency",
                )));
            }
        }
        let content = remove_xml(&self.content, &existing.span)?;
        let removed = existing.resource.clone();
        self.invalidate_content_splices();
        self.content.set_owned(content);
        self.changes
            .push(Change::Structure(StructureChange::StyleResourceRemoved {
                kind,
                name: name.to_owned(),
            }));
        Ok(removed)
    }

    /// Appends an empty structural group to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page selector or limit violation.
    pub fn add_group(&mut self, page: usize, name: impl Into<String>) -> Result<()> {
        self.add_shape(page, Shape::new(ShapeKind::Group).with_name(name))
    }

    /// Adds an inert form-control declaration without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate/invalid identifiers or output limits.
    pub fn add_form_control(&mut self, control: &FormControl) -> Result<()> {
        validate_bounded_value(control.id(), "ODG form-control identifier")?;
        if let Some(name) = control.name() {
            validate_bounded_value(name, "ODG form-control name")?;
        }
        let parsed = parse_content(&self.content)?;
        if parsed.form_controls.len() >= MAX_FORM_CONTROLS {
            return invalid("ODG form-control count exceeds the limit");
        }
        if parsed
            .form_controls
            .iter()
            .any(|value| value.id() == control.id())
        {
            return invalid("ODG form-control identifier is already present");
        }
        let control_xml = serialize_form_control(control)?;
        let content = if let Some(at) = parsed.forms_insert_position {
            let form_xml = format!(
                "<form:form xmlns:form=\"{}\" form:name=\"Litchi\">{control_xml}</form:form>",
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_child_xml(&self.content, at, &form_xml)?
        } else {
            let forms_xml = format!(
                "<office:forms xmlns:office=\"{}\" xmlns:form=\"{}\"><form:form form:name=\"Litchi\">{control_xml}</form:form></office:forms>",
                std::str::from_utf8(OFFICE).unwrap_or_default(),
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_xml(&self.content, parsed.drawing_start_position, &forms_xml)?
        };
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlInserted {
                id: control.id().to_owned(),
            }));
        Ok(())
    }

    /// Replaces one inert form declaration while preserving its referenced `form:id`.
    ///
    /// Arbitrary bounded form attributes remain data only and are never activated.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing identifier, changed identity, invalid attributes, or limits.
    pub fn replace_form_control(&mut self, control: &FormControl) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let mut matches = parsed
            .form_controls
            .iter()
            .enumerate()
            .filter(|(_index, candidate)| candidate.id() == control.id());
        let (position, _before) = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODG form-control selector did not match".into())
        })?;
        if matches.next().is_some() {
            return invalid("ODG form-control identifier is ambiguous");
        }
        let span = parsed.form_control_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG form-control span is missing".into()))?;
        let xml = serialize_form_control(control)?;
        let content = replace_xml(&self.content, span, &xml)?;
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlReplaced {
                id: control.id().to_owned(),
            }));
        Ok(())
    }

    /// Removes an inert form declaration only when no drawing shape references it.
    ///
    /// # Errors
    ///
    /// Returns an error for absent/ambiguous identifiers or a live `draw:control` dependency.
    pub fn remove_form_control(&mut self, identifier: &str) -> Result<FormControl> {
        let parsed = parse_content(&self.content)?;
        if parsed.pages.iter().any(|page| {
            page.shapes()
                .iter()
                .any(|shape| shape.control_reference() == Some(identifier))
        }) {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG form-control removal is blocked by a drawing shape",
            )));
        }
        let mut matches = parsed
            .form_controls
            .iter()
            .enumerate()
            .filter(|(_index, control)| control.id() == identifier);
        let (position, control) = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODG form-control selector did not match".into())
        })?;
        if matches.next().is_some() {
            return invalid("ODG form-control identifier is ambiguous");
        }
        let removed = control.clone();
        let span = parsed.form_control_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG form-control span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        self.invalidate_content_splices();
        self.content.set_owned(content);
        self.changes
            .push(Change::Structure(StructureChange::FormControlRemoved {
                id: identifier.to_owned(),
            }));
        Ok(removed)
    }

    /// Inserts a prepared cross-drawing shape/group and its dependency closure.
    ///
    /// Missing page-local layers and noncolliding package resources are copied. Graphic/text
    /// styles and form controls must already exist in the destination; unresolved dependencies
    /// and differing resource-path collisions are refused before publication.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, unresolved dependencies, collisions, source
    /// policy violations, or output limits.
    pub fn insert_shape_transfer(
        &mut self,
        page: usize,
        position: usize,
        transfer: &ShapeTransfer,
    ) -> Result<()> {
        ShapeTransfer::validate_destination(&self.content, page)?;
        transfer.validate_aggregate()?;
        let mut transfer_xml = transfer.xml.clone();
        let mut resource_remaps = BTreeMap::new();
        for resource in &transfer.resources {
            let destination = self.stage_transferred_resource(resource)?;
            if destination != resource.path {
                resource_remaps.insert(resource.href.clone(), destination);
            }
        }

        let mut destination_style_resources = parse_style_resources(&self.content)?;
        if let Some(styles) = self.source.styles_xml() {
            destination_style_resources.extend(parse_style_resources(styles)?);
        }
        let mut reserved_resource_names = destination_style_resources
            .iter()
            .map(|definition| definition.resource.name().to_owned())
            .collect::<Vec<_>>();
        let mut style_resource_remaps = BTreeMap::new();
        for transferred in &transfer.style_resources {
            let kind = transferred.resource.kind();
            let source_name = transferred.resource.name();
            let resource_dependency_remapped = resource_remaps
                .keys()
                .any(|href| transfer_xml_references(&transferred.xml, href));
            let collision = destination_style_resources.iter().find(|definition| {
                definition.resource.kind() == kind && definition.resource.name() == source_name
            });
            let destination_name = if collision.is_some_and(|definition| {
                definition.xml != transferred.xml || resource_dependency_remapped
            }) {
                unique_collision_name(
                    source_name,
                    transferred.xml.as_bytes(),
                    &reserved_resource_names,
                )
            } else {
                source_name.to_owned()
            };
            reserved_resource_names.push(destination_name.clone());
            style_resource_remaps.insert((kind, source_name.to_owned()), destination_name.clone());
            let already_present = !resource_dependency_remapped
                && destination_style_resources.iter().any(|definition| {
                    definition.resource.kind() == kind
                        && definition.resource.name() == destination_name
                        && definition.xml == transferred.xml
                });
            if already_present {
                continue;
            }
            let mut resource_xml = rewrite_qualified_attribute_values(
                &transferred.xml,
                &[b"draw:name"],
                source_name,
                &destination_name,
            )?;
            for (before, after) in &resource_remaps {
                resource_xml = rewrite_qualified_attribute_values(
                    &resource_xml,
                    &[b"xlink:href"],
                    before,
                    after,
                )?;
            }
            resource_xml = ensure_transfer_namespaces(
                &resource_xml,
                &[
                    ("draw", DRAW),
                    ("svg", SVG),
                    ("xlink", XLINK),
                    ("dr3d", DR3D),
                    ("table", TABLE),
                    ("smil", SMIL),
                ],
            )?;
            let content = insert_automatic_style(&self.content, &resource_xml)?;
            self.invalidate_content_splices();
            self.content.set_owned(content);
            self.changes
                .push(Change::Structure(StructureChange::StyleResourceInserted {
                    kind,
                    name: destination_name,
                }));
        }

        let mut destination_styles = parse_style_definitions(&self.content)?;
        if let Some(styles) = self.source.styles_xml() {
            destination_styles.extend(parse_style_definitions(styles)?);
        }
        let mut reserved_style_names = destination_styles
            .iter()
            .map(|definition| definition.style.name().to_owned())
            .collect::<Vec<_>>();
        let mut style_remaps = BTreeMap::new();
        for transferred in &transfer.style_definitions {
            let mut dependency_probe = transferred.xml.clone();
            for (before, after) in &resource_remaps {
                dependency_probe = rewrite_qualified_attribute_values(
                    &dependency_probe,
                    &[b"xlink:href"],
                    before,
                    after,
                )?;
            }
            for ((kind, before), after) in &style_resource_remaps {
                dependency_probe = rewrite_qualified_attribute_values(
                    &dependency_probe,
                    style_resource_qualified_references(*kind),
                    before,
                    after,
                )?;
            }
            let dependency_remapped = dependency_probe != transferred.xml;
            let collision = destination_styles
                .iter()
                .find(|definition| definition.style.name() == transferred.name);
            let destination_name = if collision.is_some_and(|definition| {
                definition.xml != transferred.xml
                    || definition.style.family() != transferred.family
                    || dependency_remapped
            }) {
                unique_collision_name(
                    &transferred.name,
                    transferred.xml.as_bytes(),
                    &reserved_style_names,
                )
            } else {
                transferred.name.clone()
            };
            reserved_style_names.push(destination_name.clone());
            style_remaps.insert(transferred.name.clone(), destination_name);
        }
        for transferred in &transfer.style_definitions {
            let destination_name = style_remaps
                .get(&transferred.name)
                .ok_or_else(|| Error::InvalidFormat("ODG style remap is missing".into()))?;
            let already_present = destination_styles.iter().any(|definition| {
                definition.style.name() == destination_name && definition.xml == transferred.xml
            });
            if already_present {
                continue;
            }
            let mut style_xml = transferred.xml.clone();
            for (before, after) in &style_remaps {
                style_xml = rewrite_qualified_attribute_values(
                    &style_xml,
                    &[b"style:name", b"style:parent-style-name"],
                    before,
                    after,
                )?;
            }
            for (before, after) in &resource_remaps {
                style_xml = rewrite_qualified_attribute_values(
                    &style_xml,
                    &[b"xlink:href"],
                    before,
                    after,
                )?;
            }
            for ((kind, before), after) in &style_resource_remaps {
                style_xml = rewrite_qualified_attribute_values(
                    &style_xml,
                    style_resource_qualified_references(*kind),
                    before,
                    after,
                )?;
            }
            style_xml = ensure_transfer_namespaces(
                &style_xml,
                &[
                    ("style", STYLE),
                    ("draw", DRAW),
                    ("svg", SVG),
                    ("fo", FO),
                    ("presentation", PRESENTATION),
                    ("smil", SMIL),
                    ("xlink", XLINK),
                ],
            )?;
            let content = insert_automatic_style(&self.content, &style_xml)?;
            self.invalidate_content_splices();
            self.content.set_owned(content);
            self.changes
                .push(Change::Structure(StructureChange::StyleInserted {
                    name: destination_name.clone(),
                }));
        }
        for (before, after) in &style_remaps {
            transfer_xml = rewrite_qualified_attribute_values(
                &transfer_xml,
                &[b"draw:style-name", b"draw:text-style-name"],
                before,
                after,
            )?;
        }
        for ((kind, before), after) in &style_resource_remaps {
            transfer_xml = rewrite_qualified_attribute_values(
                &transfer_xml,
                style_resource_qualified_references(*kind),
                before,
                after,
            )?;
        }
        for (before, after) in &resource_remaps {
            transfer_xml =
                rewrite_qualified_attribute_values(&transfer_xml, &[b"xlink:href"], before, after)?;
        }

        for transferred in &transfer.control_definitions {
            let parsed = parse_content(&self.content)?;
            let collision = parsed
                .form_controls
                .iter()
                .position(|candidate| candidate.id() == transferred.control.id());
            if collision.is_some_and(|control_position| {
                parsed.form_control_spans[control_position]
                    .as_ref()
                    .is_some_and(|span| self.content[span.clone()] == transferred.xml)
            }) {
                continue;
            }
            let mut control_xml = transferred.xml.clone();
            let destination_id = if collision.is_some() {
                let reserved = parsed
                    .form_controls
                    .iter()
                    .map(|candidate| candidate.id().to_owned())
                    .collect::<Vec<_>>();
                let id = unique_collision_name(
                    transferred.control.id(),
                    transferred.xml.as_bytes(),
                    &reserved,
                );
                transfer_xml = rewrite_qualified_attribute_values(
                    &transfer_xml,
                    &[b"draw:control"],
                    transferred.control.id(),
                    &id,
                )?;
                control_xml = rewrite_qualified_attribute_values(
                    &control_xml,
                    &[b"form:id"],
                    transferred.control.id(),
                    &id,
                )?;
                id
            } else {
                transferred.control.id().to_owned()
            };
            control_xml = ensure_transfer_namespaces(
                &control_xml,
                &[("form", FORM), ("xlink", b"http://www.w3.org/1999/xlink")],
            )?;
            self.insert_transferred_form_control(&destination_id, &control_xml)?;
        }
        transfer_xml = ensure_transfer_namespaces(
            &transfer_xml,
            &[
                ("office", OFFICE),
                ("draw", DRAW),
                ("text", TEXT),
                ("svg", SVG),
                ("style", STYLE),
                ("form", FORM),
                ("fo", FO),
                ("xlink", XLINK),
                ("presentation", PRESENTATION),
                ("dr3d", DR3D),
                ("table", TABLE),
                ("smil", SMIL),
            ],
        )?;
        let initial = parse_content(&self.content)?;
        let initial_page = initial.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        let needs_local_layer_set = !initial_page.has_layer_set()
            && transfer.layers.iter().any(|required| {
                !self
                    .source
                    .layers()
                    .iter()
                    .any(|global| global.name() == required.name())
            });
        if needs_local_layer_set {
            let mut closure = initial_page
                .shapes()
                .iter()
                .filter_map(|shape| shape.layer())
                .map(|name| resolve_transfer_layer(&self.source, initial_page, name))
                .collect::<Result<Vec<_>>>()?;
            closure.extend(transfer.layers.iter().cloned());
            closure.sort_unstable_by(|left, right| left.name().cmp(right.name()));
            closure.dedup_by(|left, right| left.name() == right.name());
            for layer in closure {
                self.add_layer(page, layer)?;
            }
        }
        for layer in &transfer.layers {
            let parsed = parse_content(&self.content)?;
            let destination_page = parsed.pages.get(page).ok_or_else(|| {
                Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
            })?;
            let visible = if destination_page.has_layer_set() {
                destination_page.layers()
            } else {
                self.source.layers()
            };
            if !visible.iter().any(|value| value.name() == layer.name()) {
                self.add_layer(page, layer.clone())?;
            }
        }
        let parsed = parse_content(&self.content)?;
        let destination_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        if position > destination_page.shapes().len() {
            return invalid("ODG transfer shape position is out of bounds");
        }
        if parsed
            .pages
            .iter()
            .map(|value| value.shapes().len())
            .sum::<usize>()
            >= MAX_SHAPES
        {
            return invalid("ODG shape count exceeds the limit");
        }
        let at = if position == destination_page.shapes().len() {
            parsed.page_insert_positions[page]
                .ok_or_else(|| Error::InvalidFormat("ODG page insertion point is missing".into()))?
        } else {
            parsed.shape_spans[page][position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?
                .start
        };
        let content = insert_child_xml(&self.content, at, &transfer_xml)?;
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeInserted {
                page,
                position,
                kind: transfer.shape.kind(),
            }));
        Ok(())
    }

    /// Removes one shape; removing a group owns and removes its complete subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid checked selector or missing source span.
    pub fn remove_shape(&mut self, page: usize, shape: usize) -> Result<Shape> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let span = parsed.shape_spans[page][shape]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?;
        let span = span.clone();
        self.splice_content(&span, "")?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeRemoved {
                page,
                position: shape,
                kind: selected.kind(),
            }));
        Ok(selected)
    }

    /// Adds a page-local layer declaration.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, duplicate names, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached layer values transfer ownership into the transaction"
    )]
    pub fn add_layer(&mut self, page: usize, layer: Layer) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .layers()
            .iter()
            .any(|value| value.name() == layer.name())
        {
            return invalid("ODG page-local layer name is already present");
        }
        let layer_xml = serialize_layer(&layer)?;
        let content = if selected.has_layer_set() {
            let at = parsed.layer_set_insert_positions[page].ok_or_else(|| {
                Error::InvalidFormat("ODG page-local layer-set insertion point is missing".into())
            })?;
            insert_child_xml(&self.content, at, &layer_xml)?
        } else {
            let page_span = parsed.page_spans[page]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
            let xml = format!(
                "<draw:layer-set xmlns:draw=\"{}\">{layer_xml}</draw:layer-set>",
                std::str::from_utf8(DRAW).unwrap_or_default()
            );
            let empty_at = page_span.end.saturating_sub(2);
            if self.content.as_bytes().get(empty_at..page_span.end) == Some(b"/>") {
                insert_child_xml(&self.content, empty_at, &xml)?
            } else {
                let at = start_tag_end(&self.content, page_span.start)?;
                insert_xml(&self.content, at, &xml)?
            }
        };
        self.invalidate_content_splices();
        self.content.set_owned(content);
        self.changes
            .push(Change::Structure(StructureChange::LayerInserted {
                page,
                name: layer.name().to_owned(),
            }));
        Ok(())
    }

    /// Removes an unreferenced page-local layer by exact name.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous name or a live shape dependency.
    pub fn remove_layer(&mut self, page: usize, name: &str) -> Result<Layer> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .shapes()
            .iter()
            .any(|shape| shape.layer() == Some(name))
        {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG layer removal is blocked by a shape assignment",
            )));
        }
        let mut matches = selected
            .layers()
            .iter()
            .enumerate()
            .filter(|(_, layer)| layer.name() == name);
        let (position, matched_layer) = matches
            .next()
            .ok_or_else(|| Error::InvalidFormat("ODG layer selector did not match".into()))?;
        if matches.next().is_some() {
            return invalid("ODG layer name selector is ambiguous");
        }
        let removed_layer = matched_layer.clone();
        let span = parsed.layer_element_spans[page][position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG layer span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        self.invalidate_content_splices();
        self.content.set_owned(content);
        self.changes
            .push(Change::Structure(StructureChange::LayerRemoved {
                page,
                name: name.to_owned(),
            }));
        Ok(removed_layer)
    }

    /// Adds or replaces one referenced package-local resource and manifest entry.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, media type, or size limit.
    pub fn set_resource(
        &mut self,
        resource: usize,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let target_media_type = media_type.into();
        validate_media_type(&target_media_type)?;
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG resource exceeds the output limit");
        }
        self.stage_resource(resource, Some(target_media_type), Some(bytes))
    }

    /// Adds a noncolliding package-local media/resource member.
    ///
    /// The member remains inert until drawing XML references it. Existing paths and unsafe package
    /// paths are refused rather than overwritten.
    ///
    /// # Errors
    ///
    /// Returns an error for unsafe/colliding paths, invalid media types, or size limits.
    pub fn add_resource(
        &mut self,
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let owned_path = path.into();
        let owned_media_type = media_type.into();
        validate_resource_path(&owned_path)?;
        validate_media_type(&owned_media_type)?;
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG resource exceeds the output limit");
        }
        if self
            .source
            .files()?
            .iter()
            .any(|value| value == &owned_path)
            || self
                .resource_edits
                .iter()
                .any(|edit| edit.path == owned_path)
        {
            return invalid("ODG resource path is already present");
        }
        self.resource_edits.push(ResourceEdit {
            resource: self
                .source
                .resources()
                .len()
                .saturating_add(self.resource_edits.len()),
            path: owned_path,
            before_media_type: None,
            after_media_type: Some(owned_media_type),
            before_bytes: None,
            after_bytes: Some(bytes),
        });
        Ok(())
    }

    /// Removes a package member only when drawing XML has no live reference to its path.
    ///
    /// # Errors
    ///
    /// Returns an error for unsafe/absent paths or a live drawing resource dependency.
    pub fn remove_unreferenced_resource(&mut self, path: &str) -> Result<()> {
        validate_resource_path(path)?;
        if self
            .source
            .resources()
            .iter()
            .any(|resource| resource.path() == path)
        {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG resource removal is blocked by a drawing reference",
            )));
        }
        let archive = self.source.0.package.package();
        if !archive.has_file(path)? {
            return invalid("ODG resource path is absent");
        }
        let package = archive.package()?;
        self.resource_edits.push(ResourceEdit {
            resource: self.source.resources().len(),
            path: path.to_owned(),
            before_media_type: package.manifest().get_media_type(path).map(str::to_owned),
            after_media_type: None,
            before_bytes: Some(package.get_file(path)?),
            after_bytes: None,
        });
        Ok(())
    }

    /// Removes one package-local resource while retaining its inert reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid resource selector.
    pub fn remove_resource(&mut self, resource: usize) -> Result<()> {
        self.stage_resource(resource, None, None)
    }

    fn stage_resource(
        &mut self,
        resource: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let selected =
            self.source.resources().get(resource).ok_or_else(|| {
                Error::InvalidFormat("ODG resource selector is out of bounds".into())
            })?;
        let before_bytes = self.source.resource_bytes(resource)?;
        let before_media_type = selected.media_type().map(str::to_owned);
        if let Some(edit) = self
            .resource_edits
            .iter_mut()
            .find(|edit| edit.resource == resource)
        {
            edit.after_media_type = after_media_type;
            edit.after_bytes = after_bytes;
        } else {
            self.resource_edits.push(ResourceEdit {
                resource,
                path: selected.path().to_owned(),
                before_media_type: before_media_type.clone(),
                after_media_type,
                before_bytes: before_bytes.clone(),
                after_bytes,
            });
        }
        self.resource_edits.retain(|edit| {
            edit.before_media_type != edit.after_media_type || edit.before_bytes != edit.after_bytes
        });
        Ok(())
    }

    fn stage_transferred_resource(&mut self, resource: &TransferResource) -> Result<String> {
        validate_resource_path(&resource.path)?;
        let Some(bytes) = &resource.bytes else {
            return Ok(resource.path.clone());
        };
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG transferred resource exceeds the output limit");
        }
        if let Some(staged) = self
            .resource_edits
            .iter()
            .find(|edit| edit.path == resource.path)
            && staged.after_bytes.as_ref() == Some(bytes)
            && staged.after_media_type == resource.media_type
        {
            return Ok(resource.path.clone());
        }
        let archive = self.source.0.package.package();
        if archive.has_file(&resource.path)? {
            let existing = archive.get_file(&resource.path)?;
            let existing_media_type = archive
                .package()?
                .manifest()
                .get_media_type(&resource.path)
                .map(str::to_owned);
            if existing == *bytes && existing_media_type == resource.media_type {
                return Ok(resource.path.clone());
            }
        }
        let destination = if self
            .resource_edits
            .iter()
            .any(|edit| edit.path == resource.path)
            || archive.has_file(&resource.path)?
        {
            unique_resource_path(&self.source, &self.resource_edits, &resource.path, bytes)?
        } else {
            resource.path.clone()
        };
        let media_type = resource
            .media_type
            .clone()
            .unwrap_or_else(|| "application/octet-stream".to_string());
        validate_media_type(&media_type)?;
        self.resource_edits.push(ResourceEdit {
            resource: self
                .source
                .resources()
                .len()
                .saturating_add(self.resource_edits.len()),
            path: destination.clone(),
            before_media_type: None,
            after_media_type: Some(media_type),
            before_bytes: None,
            after_bytes: Some(bytes.clone()),
        });
        Ok(destination)
    }

    fn insert_transferred_form_control(&mut self, id: &str, xml: &str) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if parsed.form_controls.len() >= MAX_FORM_CONTROLS {
            return invalid("ODG form-control count exceeds the limit");
        }
        if parsed
            .form_controls
            .iter()
            .any(|control| control.id() == id)
        {
            return invalid("ODG transferred form-control identifier is already present");
        }
        if xml.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG transferred form-control exceeds the output limit");
        }
        let content = if let Some(at) = parsed.forms_insert_position {
            let prefix = format!(
                "<form:form xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" form:name=\"Litchi\">",
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            let suffix = "</form:form>";
            let capacity = prefix
                .len()
                .checked_add(xml.len())
                .and_then(|size| size.checked_add(suffix.len()))
                .ok_or_else(|| Error::InvalidFormat("ODG transferred form size overflow".into()))?;
            let mut form_xml =
                precharge_xml_capacity(prefix, capacity, "ODG transferred form-control")?;
            form_xml.push_str(xml);
            form_xml.push_str(suffix);
            insert_child_xml(&self.content, at, &form_xml)?
        } else {
            let prefix = format!(
                "<office:forms xmlns:office=\"{}\" xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"><form:form form:name=\"Litchi\">",
                std::str::from_utf8(OFFICE).unwrap_or_default(),
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            let suffix = "</form:form></office:forms>";
            let capacity = prefix
                .len()
                .checked_add(xml.len())
                .and_then(|size| size.checked_add(suffix.len()))
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transferred forms size overflow".into())
                })?;
            let mut forms_xml =
                precharge_xml_capacity(prefix, capacity, "ODG transferred form-controls")?;
            forms_xml.push_str(xml);
            forms_xml.push_str(suffix);
            insert_xml(&self.content, parsed.drawing_start_position, &forms_xml)?
        };
        self.invalidate_content_splices();
        self.content.set_owned(content);
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlInserted {
                id: id.to_owned(),
            }));
        Ok(())
    }

    fn stage_durable_operation(
        &mut self,
        operation: &PatchOperation,
        source_fingerprint: &str,
        page_name_targets: Option<&BTreeMap<usize, String>>,
    ) -> Result<()> {
        if operation
            .preconditions
            .get("source")
            .and_then(serde_json::Value::as_str)
            != Some(source_fingerprint)
        {
            return Err(Error::PatchConflict);
        }
        match operation.op.as_str() {
            "page.name.set" => {
                self.stage_page_name_operation(
                    parse_page_target(&operation.target)?,
                    string_value(operation)?.to_owned(),
                    page_name_targets,
                )?;
            },
            "page.style.set" => {
                self.set_page_style_name(
                    parse_page_target(&operation.target)?,
                    string_value(operation)?,
                )?;
            },
            "page.transition.set" => {
                self.set_page_transition(
                    parse_page_target(&operation.target)?,
                    transition_from_value(&operation.value)?,
                )?;
            },
            "shape.geometry.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                let values = geometry_value(&operation.value)?;
                self.set_shape_geometry(
                    page, shape, &values[0], &values[1], &values[2], &values[3],
                )?;
            },
            "shape.control.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_control_reference(page, shape, string_value(operation)?)?;
            },
            "shape.layer.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_layer(page, shape, string_value(operation)?)?;
            },
            "shape.name.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_name(page, shape, string_value(operation)?)?;
            },
            "shape.path.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_path_data(page, shape, string_value(operation)?)?;
            },
            "shape.style.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_style_name(page, shape, string_value(operation)?)?;
            },
            "shape.text.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                self.set_shape_text(page, shape, string_value(operation)?)?;
            },
            _ => return invalid("ODG durable patch operation is unsupported"),
        }
        Ok(())
    }

    fn stage_page_name_operation(
        &mut self,
        page: usize,
        after: String,
        page_name_targets: Option<&BTreeMap<usize, String>>,
    ) -> Result<()> {
        if page_name_targets.is_none() {
            return self.set_page_name(page, after);
        }
        validate_bounded_value(&after, "ODG page name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let before = selected.name().ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page rename requires an existing draw:name",
            ))
        })?;
        if before == after {
            return Ok(());
        }
        if xml_has_attribute(&self.content, DRAW, b"page-name", before)? {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG page rename is blocked by a draw:page-name dependency",
            )));
        }
        let span = parsed.page_attribute_spans[page][0]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page name source span is missing".into()))?;
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::PageName(PageNameChange {
            page,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    fn apply_pending_joined(&mut self) -> Result<()> {
        let Some(joined) = self.pending_joined.as_ref() else {
            return Ok(());
        };
        if self.joined_conflicts_with_staged(joined)? {
            return invalid("ODG joined edit overlaps staged work");
        }
        let source_fingerprint = fingerprint(self.source.as_bytes());
        for sub_edit in joined.sub_edits() {
            validate_durable_patch(&sub_edit.payload().inner)?;
            validate_joined_preconditions(sub_edit.payload(), &source_fingerprint)?;
        }
        self.validate_joined_page_names(joined)?;
        let joined = self
            .pending_joined
            .take()
            .expect("pending joined plan was borrowed above");
        let mut staged_before_replacement = Vec::new();
        let mut replacement_operations: Option<Vec<PatchOperation>> = None;
        for sub_edit in joined.into_sub_edits() {
            let durable = sub_edit.into_payload();
            for operation in durable.operations() {
                if operation.op == "package.replace" {
                    if let Some(operations) = &replacement_operations {
                        self.commit_replacement_operations(operations, &source_fingerprint)?;
                    } else {
                        self.stage_durable_operations(
                            &staged_before_replacement,
                            &source_fingerprint,
                        )?;
                        staged_before_replacement.clear();
                    }
                    let active = self.package_replacement.as_ref().unwrap_or(&self.source);
                    enforce_security_policy(active, self.security_policy)?;
                    enforce_active_content_policy(active, self.active_content_policy)?;
                    let blob = durable_blob(&durable, &operation.value)?;
                    let replacement = active.reopen_same_flavor(blob.to_vec())?;
                    enforce_active_content_policy(&replacement, self.active_content_policy)?;
                    self.package_replacement = Some(replacement);
                    replacement_operations = Some(Vec::new());
                } else if let Some(operations) = &mut replacement_operations {
                    operations.push(operation.clone());
                } else {
                    staged_before_replacement.push(operation.clone());
                }
            }
        }
        if let Some(operations) = &replacement_operations {
            self.commit_replacement_operations(operations, &source_fingerprint)?;
        } else {
            self.stage_durable_operations(&staged_before_replacement, &source_fingerprint)?;
        }
        Ok(())
    }

    fn stage_durable_operations(
        &mut self,
        operations: &[PatchOperation],
        source_fingerprint: &str,
    ) -> Result<()> {
        if operations.is_empty() {
            return Ok(());
        }
        let page_name_targets = self.page_name_targets_for_operations(operations.iter())?;
        for operation in operations {
            self.stage_durable_operation(operation, source_fingerprint, Some(&page_name_targets))?;
        }
        Ok(())
    }

    fn commit_replacement_operations(
        &mut self,
        operations: &[PatchOperation],
        source_fingerprint: &str,
    ) -> Result<()> {
        if operations.is_empty() {
            return Ok(());
        }
        let replacement = self
            .package_replacement
            .take()
            .expect("package replacement was retained for ordered operations");
        let mut replacement_edit =
            replacement.edit_with_policies(self.security_policy, self.active_content_policy);
        let page_name_targets =
            replacement_edit.page_name_targets_for_operations(operations.iter())?;
        for operation in operations {
            replacement_edit.stage_durable_operation(
                operation,
                source_fingerprint,
                Some(&page_name_targets),
            )?;
        }
        self.package_replacement = Some(replacement_edit.commit_core()?.into_package());
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the edited package.
    ///
    /// An edit that leaves every part unchanged publishes the exact source
    /// bytes and reports [`Diagnostics::changed`] as `false`.
    ///
    /// # Errors
    ///
    /// Returns an error when source policy, rebuilding, parsing, or typed readback fails.
    pub fn commit(self) -> std::result::Result<Commit, OdfError> {
        Ok(self.commit_core()?)
    }

    pub(crate) fn commit_core(mut self) -> Result<Commit> {
        self.apply_pending_joined()?;
        if let Some(snapshot) = self.package_replacement.take() {
            let changed = self.source.as_bytes() != snapshot.as_bytes();
            let patch = Patch {
                source: self.source,
                target: snapshot.clone(),
                changes: Vec::new(),
                resource_changes: Vec::new(),
                requires_package_projection: true,
            };
            return Ok(Commit {
                snapshot,
                patch,
                changed,
            });
        }
        if self.content.as_str() == self.source.content_xml()
            && self.styles.as_deref() == self.source.styles_xml()
            && self.resource_edits.is_empty()
        {
            return Ok(Commit::unchanged(self.source));
        }
        enforce_security_policy(&self.source, self.security_policy)?;
        enforce_active_content_policy(&self.source, self.active_content_policy)?;
        let target_active_content = scan_active_content(&self.content, self.styles.as_deref())?;
        if self.active_content_policy == ActiveContentWritePolicy::Refuse
            && target_active_content.is_present()
        {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG active-content write policy refuses the staged inventory",
            )));
        }
        let replacements = self
            .resource_edits
            .iter()
            .map(|edit| ResourceReplacement {
                path: &edit.path,
                media_type: edit.after_media_type.as_deref().unwrap_or_default(),
                bytes: edit.after_bytes.as_deref(),
            })
            .collect::<Vec<_>>();
        // An untouched styles.xml must remain an archive copy.  Passing it as
        // replacement would make PackageWriter revalidate producer
        // formatting that was never edited, and would reject otherwise valid
        // packages with noncompact styles.xml.  A changed styles.xml is still
        // published through the normal compact replacement path.
        let published_styles = if self.styles.as_deref() == self.source.styles_xml() {
            None
        } else {
            self.styles.as_deref()
        };
        let requires_package_projection;
        let rebuilt = if let Some(splices) = &self.content_splices {
            requires_package_projection = self.requires_package_projection
                || self.source.security().is_signed()
                || ensure_compact_rewrite_source(&self.source).is_err();
            let publication = content_splice_publication(&self.source, splices)?;
            rebuild_spliced(
                &self.source,
                publication,
                published_styles,
                &replacements,
                self.security_policy,
            )?
        } else {
            requires_package_projection = self.requires_package_projection;
            ensure_compact_rewrite_source(&self.source)?;
            compact_xml::validate(self.content.as_bytes(), compact_xml::Limits::default())
                .map_err(Error::from)?;
            rebuild(
                &self.source,
                &self.content,
                published_styles,
                &replacements,
                self.security_policy,
            )?
        };
        let snapshot = self.source.reopen_same_flavor(rebuilt)?;
        if snapshot.content_xml() != self.content.as_str() {
            return invalid("ODG package edit failed exact content readback");
        }
        if snapshot.styles_xml() != self.styles.as_deref() {
            return invalid("ODG package edit failed exact styles readback");
        }
        for edit in &self.resource_edits {
            let archive = snapshot.0.package.package().package()?;
            if archive.manifest().get_media_type(&edit.path) != edit.after_media_type.as_deref() {
                return invalid("ODG resource edit failed manifest readback");
            }
            let actual = if snapshot.0.package.package().has_file(&edit.path)? {
                Some(snapshot.0.package.package().get_file(&edit.path)?)
            } else {
                None
            };
            if actual != edit.after_bytes {
                return invalid("ODG resource edit failed byte readback");
            }
        }
        let resource_changes = self
            .resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        Ok(Commit {
            patch: Patch {
                source: self.source,
                target: snapshot.clone(),
                changes: self.changes,
                resource_changes,
                requires_package_projection,
            },
            snapshot,
            changed: true,
        })
    }
}

/// One semantic operation published by a unified ODG package transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    ControlReference(ControlReferenceChange),
    Text(TextChange),
    Name(NameChange),
    Layer(LayerChange),
    Geometry(GeometryChange),
    Style(StyleChange),
    Path(PathChange),
    PageName(PageNameChange),
    PageStyle(PageStyleChange),
    PageTransition(Box<PageTransitionChange>),
    Structure(StructureChange),
}

fn change_effect_keys(_source: &Package, change: &Change) -> Vec<String> {
    let (target, operation) = match change {
        Change::ControlReference(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.control.set",
        ),
        Change::Text(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.text.set",
        ),
        Change::Name(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.name.set",
        ),
        Change::Layer(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.layer.set",
        ),
        Change::Geometry(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.geometry.set",
        ),
        Change::Style(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.style.set",
        ),
        Change::Path(value) => (
            format!("page/{}/shape/{}", value.page, value.shape),
            "shape.path.set",
        ),
        Change::PageName(value) => (format!("page/{}", value.page), "page.name.set"),
        Change::PageStyle(value) => (format!("page/{}", value.page), "page.style.set"),
        Change::PageTransition(value) => (format!("page/{}", value.page), "page.transition.set"),
        Change::Structure(_) => return vec!["package".to_owned()],
    };
    vec![format!("{target}#{operation}")]
}

/// One reversible page-name change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageNameChange {
    page: usize,
    before: String,
    after: String,
}

impl PageNameChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible drawing-page style-reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageStyleChange {
    page: usize,
    before: String,
    after: String,
}

impl PageStyleChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible inert drawing-page transition change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageTransitionChange {
    page: usize,
    before: Option<Transition>,
    after: Option<Transition>,
}

impl PageTransitionChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// Transition metadata expected before application.
    #[must_use]
    pub const fn before(&self) -> Option<&Transition> {
        self.before.as_ref()
    }

    /// Transition metadata produced after application.
    #[must_use]
    pub const fn after(&self) -> Option<&Transition> {
        self.after.as_ref()
    }
}

/// One reversible inert form-control reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ControlReferenceChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl ControlReferenceChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible semantic shape-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl TextChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Text expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Text produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible `draw:name` change for a shape.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl NameChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible drawing-layer assignment change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LayerChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl LayerChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// The layer name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The layer name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible four-attribute geometry change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GeometryChange {
    page: usize,
    shape: usize,
    before: [String; 4],
    after: [String; 4],
}

impl GeometryChange {
    /// Page position at the time of this operation.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// Shape position at the time of this operation.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Source `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn before(&self) -> &[String; 4] {
        &self.before
    }

    /// Target `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn after(&self) -> &[String; 4] {
        &self.after
    }
}

/// One reversible graphic-style reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

/// One reversible SVG path-data change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PathChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl PathChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

impl StyleChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A structural page, layer, shape, or group operation.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum StructureChange {
    PageInserted {
        position: usize,
        name: Option<String>,
    },
    /// A page moved from one source-order position to another.
    PageMoved {
        from: usize,
        to: usize,
    },
    /// A shape moved within one page's shape order.
    ShapeMoved {
        page: usize,
        from: usize,
        to: usize,
    },
    /// A top-level shape moved from position `from` of `page` to position
    /// `to` of another page, `to_page`.
    ShapeTransferred {
        page: usize,
        from: usize,
        to_page: usize,
        to: usize,
    },
    PageRemoved {
        position: usize,
        name: Option<String>,
    },
    LayerInserted {
        page: usize,
        name: String,
    },
    LayerRemoved {
        page: usize,
        name: String,
    },
    ShapeInserted {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
    ShapeRemoved {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
    ShapeAdvancedGeometryChanged {
        page: usize,
        shape: usize,
    },
    FormControlInserted {
        id: String,
    },
    FormControlRemoved {
        id: String,
    },
    FormControlReplaced {
        id: String,
    },
    StyleInserted {
        name: String,
    },
    StyleRemoved {
        name: String,
    },
    StyleReplaced {
        name: String,
    },
    StyleResourceInserted {
        kind: StyleResourceKind,
        name: String,
    },
    StyleResourceRemoved {
        kind: StyleResourceKind,
        name: String,
    },
    StyleResourceReplaced {
        kind: StyleResourceKind,
        name: String,
    },
}

/// One package-local resource replacement or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    #[must_use]
    pub const fn resource(&self) -> usize {
        self.resource
    }

    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn before_media_type(&self) -> Option<&str> {
        self.before_media_type.as_deref()
    }

    #[must_use]
    pub fn after_media_type(&self) -> Option<&str> {
        self.after_media_type.as_deref()
    }

    #[must_use]
    pub const fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    #[must_use]
    pub const fn after_size(&self) -> Option<usize> {
        self.after_size
    }
}

struct ResourceEdit {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_bytes: Option<Vec<u8>>,
    after_bytes: Option<Vec<u8>>,
}

impl ResourceEdit {
    fn change(&self) -> ResourceChange {
        ResourceChange {
            resource: self.resource,
            path: self.path.clone(),
            before_media_type: self.before_media_type.clone(),
            after_media_type: self.after_media_type.clone(),
            before_size: self.before_bytes.as_ref().map(Vec::len),
            after_size: self.after_bytes.as_ref().map(Vec::len),
        }
    }
}

struct ResourceReplacement<'a> {
    path: &'a str,
    media_type: &'a str,
    bytes: Option<&'a [u8]>,
}

/// A committed package publication and its exact-source patch.
pub struct Commit {
    snapshot: Package,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Package) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                changes: Vec::new(),
                resource_changes: Vec::new(),
                requires_package_projection: false,
            },
            snapshot,
            changed: false,
        }
    }

    /// Borrows the published immutable package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.snapshot
    }

    /// Borrows the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free facts about this commit.
    #[must_use]
    pub fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            changed: self.changed,
            change_count: self.patch.changes.len(),
            resource_change_count: self.patch.resource_changes.len(),
        }
    }

    /// Consumes the commit into its published package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.snapshot
    }

    /// Consumes the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Content-free facts about one [`Commit`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    change_count: usize,
    resource_change_count: usize,
}

impl Diagnostics {
    /// Whether publication rebuilt the package bytes.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The number of semantic operations the commit published.
    #[must_use]
    pub const fn change_count(&self) -> usize {
        self.change_count
    }

    /// The number of package-local resource changes the commit published.
    #[must_use]
    pub const fn resource_change_count(&self) -> usize {
        self.resource_change_count
    }
}

/// An exact-source-checked reversible package patch.
#[derive(Clone)]
pub struct Patch {
    source: Package,
    target: Package,
    changes: Vec<Change>,
    resource_changes: Vec<ResourceChange>,
    requires_package_projection: bool,
}

impl Patch {
    /// Whether this patch authorizes the supplied exact source bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Package) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Whether this patch preserves the exact package bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source.as_bytes() == self.target.as_bytes()
    }

    /// Applies this patch only to its exact source package; the public route
    /// is [`Package::apply`].
    pub(crate) fn apply(&self, source: &Package) -> Result<Package> {
        if !self.is_applicable_to(source) {
            return invalid("ODG package patch source does not match");
        }
        Ok(self.target.clone())
    }

    /// The semantic change represented by this patch.
    #[must_use]
    pub fn change(&self) -> Option<&TextChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Text(value) => Some(value),
            Change::ControlReference(_)
            | Change::Name(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::PageTransition(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic `draw:name` change, when this is a name patch.
    #[must_use]
    pub fn name_change(&self) -> Option<&NameChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Name(value) => Some(value),
            Change::ControlReference(_)
            | Change::Text(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::PageTransition(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic drawing-layer change, when present.
    #[must_use]
    pub fn layer_change(&self) -> Option<&LayerChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Layer(value) => Some(value),
            Change::ControlReference(_)
            | Change::Text(_)
            | Change::Name(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::PageTransition(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic page transition change, when present.
    #[must_use]
    pub fn page_transition_change(&self) -> Option<&PageTransitionChange> {
        self.changes.iter().find_map(|change| match change {
            Change::PageTransition(value) => Some(value.as_ref()),
            Change::ControlReference(_)
            | Change::Text(_)
            | Change::Name(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::Structure(_) => None,
        })
    }

    /// All semantic operations in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Package-local resource changes in source selector order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Composes adjacent exact-lineage patches.
    ///
    /// # Errors
    ///
    /// Returns an error unless this target is byte-identical to `next`'s source.
    pub fn then(&self, next: &Self) -> Result<Self> {
        if self.target.as_bytes() != next.source.as_bytes() {
            return Err(Error::PatchConflict);
        }
        let mut changes = self.changes.clone();
        changes.extend_from_slice(&next.changes);
        let mut resource_changes = self.resource_changes.clone();
        resource_changes.extend_from_slice(&next.resource_changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
            resource_changes,
            requires_package_projection: self.requires_package_projection
                || next.requires_package_projection,
        })
    }

    /// An exact-source patch restoring the original package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().rev().map(inverse_change).collect(),
            resource_changes: self
                .resource_changes
                .iter()
                .rev()
                .map(inverse_resource_change)
                .collect(),
            requires_package_projection: self.requires_package_projection,
        }
    }

    /// Projects this exact-source patch into the shared durable semantic wire format.
    ///
    /// # Errors
    ///
    /// Returns an error when a semantic operation exceeds the durable patch limits.
    pub fn durable(&self) -> Result<DurablePatch> {
        let limits = durable_limits();
        let source = fingerprint(self.source.as_bytes());
        let target = fingerprint(self.target.as_bytes());
        if !self.resource_changes.is_empty()
            || self.requires_package_projection
            || self
                .changes
                .iter()
                .any(|change| matches!(change, Change::Structure(_)))
        {
            return package_replacement_patch(self, limits, &source, &target);
        }
        let mut operations = Vec::new();
        for change in &self.changes {
            operations.push(change_operation(change, limits, &source, &target)?);
        }
        CorePatch::<Reversible>::new(
            limits,
            DURABLE_FORMAT,
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
        .map(|inner| DurablePatch { inner })
        .map_err(durable_error)
    }

    /// Prepares this patch as an independently joinable sub-edit.
    ///
    /// # Errors
    ///
    /// Returns an error when durable projection or bounded effect construction fails.
    pub fn prepare(
        &self,
        identifier: impl Into<String>,
        limits: CompositionLimits,
    ) -> Result<SubEdit> {
        let durable = self.durable()?;
        let writes = durable
            .operations()
            .iter()
            .map(durable_effect_keys)
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .flatten()
            .collect::<Vec<_>>();
        SubEdit::new(
            Lineage::new(&self.source),
            limits,
            identifier,
            Vec::<String>::new(),
            writes,
            durable,
        )
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG sub-edit: {error}")))
    }
}

/// Exact ODG source lineage used by deterministic sub-edit composition.
#[derive(Clone, PartialEq, Eq)]
pub struct Lineage(Arc<[u8]>);

impl Lineage {
    fn new(snapshot: &Package) -> Self {
        Self(Arc::from(snapshot.as_bytes()))
    }

    fn matches(&self, snapshot: &Package) -> bool {
        self.0.as_ref() == snapshot.as_bytes()
    }
}

/// One independently prepared ODG semantic patch.
pub type SubEdit = litchi_core::SubEdit<Lineage, DurablePatch>;

/// Deterministically ordered, provably disjoint ODG sub-edits.
pub type JoinedEdits = JoinedSubEdits<Lineage, DurablePatch>;

/// Recoverable refusal while staging a bounded joined plan on an edit.
///
/// The rejected plan is retained so callers can correct or route it to a
/// different edit without reconstructing its payloads. The error itself is
/// content-free; the plan is available only through the explicit recovery
/// accessor.
#[derive(Debug)]
pub struct JoinError {
    error: Error,
    rejected: Box<JoinedEdits>,
}

impl JoinError {
    fn new(error: Error, rejected: JoinedEdits) -> Self {
        Self {
            error,
            rejected: Box::new(rejected),
        }
    }

    /// The fixed classified reason for refusing the plan.
    #[must_use]
    pub const fn error(&self) -> &Error {
        &self.error
    }

    /// Recover the rejected plan for retry or alternate routing.
    #[must_use]
    pub fn into_rejected(self) -> JoinedEdits {
        *self.rejected
    }

    /// Recover both the refusal and the rejected plan.
    #[must_use]
    pub fn into_parts(self) -> (Error, JoinedEdits) {
        (self.error, *self.rejected)
    }
}

/// Non-mutating three-way ODG merge plan.
pub type MergePlan = litchi_core::ThreeWayMergePlan<Lineage, DurablePatch>;

/// Explicit bounded ODG undo/redo history.
pub type History = CoreHistory<Package>;

/// A bounded, provenance-bound shape or complete group-subtree transfer plan.
#[derive(Clone)]
pub struct ShapeTransfer {
    source: Lineage,
    shape: Shape,
    xml: String,
    layers: Vec<Layer>,
    styles: Vec<String>,
    style_definitions: Vec<TransferStyle>,
    style_resources: Vec<TransferStyleResource>,
    controls: Vec<String>,
    control_definitions: Vec<TransferControl>,
    resources: Vec<TransferResource>,
}

#[derive(Default)]
struct TransferBudget {
    bytes: usize,
}

impl TransferBudget {
    fn charge(&mut self, bytes: usize, owner: &str) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat(format!("{owner} size overflow")))?;
        if self.bytes > MAX_OUTPUT_BYTES {
            return invalid(format!("{owner} aggregate exceeds the output limit"));
        }
        Ok(())
    }
}

impl ShapeTransfer {
    fn validate_aggregate(&self) -> Result<()> {
        let mut budget = TransferBudget::default();
        budget.charge(self.xml.len(), "ODG transfer XML")?;
        for style in &self.style_definitions {
            budget.charge(style.xml.len(), "ODG transfer XML")?;
        }
        for resource in &self.style_resources {
            budget.charge(resource.xml.len(), "ODG transfer XML")?;
        }
        for control in &self.control_definitions {
            budget.charge(control.xml.len(), "ODG transfer XML")?;
        }
        for resource in &self.resources {
            if let Some(bytes) = &resource.bytes {
                budget.charge(bytes.len(), "ODG transfer resources")?;
            }
        }
        Ok(())
    }

    /// Root shape semantics retained by the transfer.
    #[must_use]
    pub const fn shape(&self) -> &Shape {
        &self.shape
    }

    /// Required declared layers in stable name order.
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.layers
    }

    /// Required graphic/text style names in stable order.
    #[must_use]
    pub fn styles(&self) -> &[String] {
        &self.styles
    }

    /// Exact compact style-definition closure in stable name order.
    #[must_use]
    pub fn style_definitions(&self) -> &[TransferStyle] {
        &self.style_definitions
    }

    /// Exact named gradient/hatch/fill/marker/opacity/dash closure in stable order.
    #[must_use]
    pub fn style_resources(&self) -> &[TransferStyleResource] {
        &self.style_resources
    }

    /// Required inert form-control identifiers in stable order.
    #[must_use]
    pub fn controls(&self) -> &[String] {
        &self.controls
    }

    /// Inert form-control declaration closure in stable identifier order.
    #[must_use]
    pub fn control_definitions(&self) -> &[TransferControl] {
        &self.control_definitions
    }

    /// Package-local resource closure in source occurrence order.
    #[must_use]
    pub fn resources(&self) -> &[TransferResource] {
        &self.resources
    }

    /// Content-free fingerprint of the exact source artifact.
    #[must_use]
    pub fn source_fingerprint(&self) -> String {
        DiagnosticFingerprint::of(self.source.0.as_ref()).as_hex()
    }

    fn validate_destination(content: &str, page: usize) -> Result<()> {
        if parse_content(content)?.pages.get(page).is_none() {
            return invalid("ODG transfer destination page is out of bounds");
        }
        Ok(())
    }
}

/// One exact compact style dependency retained by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferStyle {
    name: String,
    family: String,
    parent: Option<String>,
    xml: String,
}

/// One exact compact named drawing resource retained by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferStyleResource {
    resource: StyleResource,
    xml: String,
}

impl TransferStyleResource {
    /// Typed inert source resource.
    #[must_use]
    pub const fn resource(&self) -> &StyleResource {
        &self.resource
    }
}

impl TransferStyle {
    /// Source style name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Source style family.
    #[must_use]
    pub fn family(&self) -> &str {
        &self.family
    }

    /// Optional parent-style dependency.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }
}

/// One exact inert form-control dependency retained by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferControl {
    control: FormControl,
    xml: String,
}

impl TransferControl {
    /// Parsed inert form semantics.
    #[must_use]
    pub const fn control(&self) -> &FormControl {
        &self.control
    }
}

/// One inert package-local resource retained by a shape transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferResource {
    href: String,
    path: String,
    media_type: Option<String>,
    bytes: Option<Vec<u8>>,
}

impl TransferResource {
    /// Exact source hyperlink spelling retained for collision rewriting.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Safe package-member path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Declared manifest media type, when present.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Retained inert bytes, or `None` for a missing source reference.
    #[must_use]
    pub fn bytes(&self) -> Option<&[u8]> {
        self.bytes.as_deref()
    }
}

/// A durable, versioned ODG semantic patch.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    /// Parses canonical deterministic JSON under the ODG patch limits.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, malformed, over-limit, or wrong-format input.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_limits())
            .map_err(durable_error)?;
        if inner.format() != DURABLE_FORMAT {
            return invalid("durable patch is not an ODG patch");
        }
        validate_durable_patch(&inner)?;
        validate_durable_patch(&inner.inverse())?;
        Ok(Self { inner })
    }

    /// Applies every operation after exact source precondition checks.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, unsupported operations, security refusal, or failed
    /// whole-package readback.
    pub fn apply(&self, source: &Package) -> Result<Package> {
        self.apply_with_policies(
            source,
            SecurityWritePolicy::Refuse,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies this patch under an explicit signature-write policy.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, invalid policy, unsupported encryption, or readback.
    pub fn apply_with_security_policy(
        &self,
        source: &Package,
        security_policy: SecurityWritePolicy,
    ) -> Result<Package> {
        self.apply_with_policies(
            source,
            security_policy,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies this patch under explicit security and inert active-content dispositions.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, policy refusal, unsupported operations, or readback.
    pub fn apply_with_policies(
        &self,
        source: &Package,
        security_policy: SecurityWritePolicy,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Result<Package> {
        apply_durable_patch(source, self, true, security_policy, active_content_policy)
    }

    /// Returns the inverse durable patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Forward semantic operations in deterministic order.
    #[must_use]
    pub fn operations(&self) -> &[PatchOperation] {
        self.inner.operations()
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch exceeds its retained serialization limits.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(durable_error)
    }

    /// Content-free diagnostic fingerprint of the canonical wire envelope.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded canonical serialization fails.
    pub fn fingerprint(&self) -> Result<DiagnosticFingerprint> {
        self.inner.fingerprint().map_err(durable_error)
    }
}

fn package_replacement_patch(
    patch: &Patch,
    limits: PatchLimits,
    source: &str,
    target: &str,
) -> Result<DurablePatch> {
    let mut forward_blobs = BlobBundle::new(limits.blobs());
    let forward_id = forward_blobs
        .insert(patch.target.as_bytes())
        .map_err(durable_error)?;
    let mut reverse_blobs = BlobBundle::new(limits.blobs());
    let reverse_id = reverse_blobs
        .insert(patch.source.as_bytes())
        .map_err(durable_error)?;
    let operation = reversible_operation(
        limits,
        "package.replace",
        "package",
        source,
        target,
        serde_json::Value::String(reverse_id.as_hex()),
        serde_json::Value::String(forward_id.as_hex()),
    )?;
    CorePatch::<Reversible>::new(
        limits,
        DURABLE_FORMAT,
        [operation],
        forward_blobs,
        reverse_blobs,
    )
    .map(|inner| DurablePatch { inner })
    .map_err(durable_error)
}

fn change_operation(
    change: &Change,
    limits: PatchLimits,
    source: &str,
    target: &str,
) -> Result<ReversibleOperation> {
    let (name, semantic_target, before, after) = match change {
        Change::ControlReference(value) => (
            "shape.control.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Text(value) => (
            "shape.text.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Name(value) => (
            "shape.name.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Layer(value) => (
            "shape.layer.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Geometry(value) => (
            "shape.geometry.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::json!(value.before),
            serde_json::json!(value.after),
        ),
        Change::Style(value) => (
            "shape.style.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Path(value) => (
            "shape.path.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::PageName(value) => (
            "page.name.set",
            format!("page/{}", value.page),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::PageStyle(value) => (
            "page.style.set",
            format!("page/{}", value.page),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::PageTransition(value) => (
            "page.transition.set",
            format!("page/{}", value.page),
            transition_value(value.before.as_ref()),
            transition_value(value.after.as_ref()),
        ),
        Change::Structure(_) => {
            return invalid("structural ODG change requires package replacement projection");
        },
    };
    reversible_operation(
        limits,
        name,
        &semantic_target,
        source,
        target,
        before,
        after,
    )
}

fn transition_value(value: Option<&Transition>) -> serde_json::Value {
    let Some(value) = value else {
        return serde_json::Value::Null;
    };
    let mut object = serde_json::Map::new();
    let fields = [
        ("transition_type", value.transition_type()),
        ("style", value.style()),
        ("speed", value.speed()),
        ("smil_type", value.smil_type()),
        ("smil_subtype", value.smil_subtype()),
        ("direction", value.direction()),
        ("fade_color", value.fade_color()),
        ("duration", value.duration()),
    ];
    for (name, field) in fields {
        if let Some(field) = field {
            object.insert(name.to_owned(), serde_json::Value::String(field.to_owned()));
        }
    }
    if let Some(sound) = value.sound() {
        let mut sound_object = serde_json::Map::new();
        sound_object.insert(
            "href".to_owned(),
            serde_json::Value::String(sound.href().to_owned()),
        );
        if let Some(play_full) = sound.play_full() {
            sound_object.insert("play_full".to_owned(), serde_json::Value::Bool(play_full));
        }
        if sound.actuate_on_request() {
            sound_object.insert(
                "actuate_on_request".to_owned(),
                serde_json::Value::Bool(true),
            );
        }
        if let Some(show) = sound.show() {
            sound_object.insert(
                "show".to_owned(),
                serde_json::Value::String(show.to_owned()),
            );
        }
        if let Some(xml_id) = sound.xml_id() {
            sound_object.insert(
                "xml_id".to_owned(),
                serde_json::Value::String(xml_id.to_owned()),
            );
        }
        object.insert("sound".to_owned(), serde_json::Value::Object(sound_object));
    }
    serde_json::Value::Object(object)
}

fn transition_from_value(value: &serde_json::Value) -> Result<Option<Transition>> {
    let Some(object) = value.as_object() else {
        if value.is_null() {
            return Ok(None);
        }
        return invalid("ODG durable page transition value is not an object");
    };
    let mut transition = Transition::new();
    transition.set_transition_type(transition_string(object, "transition_type")?)?;
    transition.set_style(transition_string(object, "style")?)?;
    transition.set_speed(transition_string(object, "speed")?)?;
    transition.set_smil_type(transition_string(object, "smil_type")?)?;
    transition.set_smil_subtype(transition_string(object, "smil_subtype")?)?;
    transition.set_direction(transition_string(object, "direction")?)?;
    transition.set_fade_color(transition_string(object, "fade_color")?)?;
    transition.set_duration(transition_string(object, "duration")?)?;
    if let Some(value) = object.get("sound") {
        if !value.is_null() {
            let sound = value.as_object().ok_or_else(|| {
                Error::InvalidFormat("ODG durable transition sound is not an object".into())
            })?;
            let href = required_transition_string(sound, "href")?;
            let sound = crate::transition::Sound::new(href)?
                .with_play_full(transition_bool(sound, "play_full")?)
                .with_actuate_on_request(
                    transition_bool(sound, "actuate_on_request")?.unwrap_or(false),
                )
                .with_show(transition_string(sound, "show")?)?
                .with_xml_id(transition_string(sound, "xml_id")?)?;
            transition.set_sound(Some(sound));
        }
    }
    Ok((!transition.is_empty()).then_some(transition))
}

fn transition_string(
    object: &serde_json::Map<String, serde_json::Value>,
    key: &str,
) -> Result<Option<String>> {
    let Some(value) = object.get(key) else {
        return Ok(None);
    };
    if value.is_null() {
        return Ok(None);
    }
    value
        .as_str()
        .map(str::to_owned)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("ODG durable transition {key} is not a string"))
        })
        .map(Some)
}

fn required_transition_string(
    object: &serde_json::Map<String, serde_json::Value>,
    key: &str,
) -> Result<String> {
    transition_string(object, key)?.ok_or_else(|| {
        Error::InvalidFormat(format!("ODG durable transition sound {key} is missing"))
    })
}

fn transition_bool(
    object: &serde_json::Map<String, serde_json::Value>,
    key: &str,
) -> Result<Option<bool>> {
    let Some(value) = object.get(key) else {
        return Ok(None);
    };
    if value.is_null() {
        return Ok(None);
    }
    value
        .as_bool()
        .ok_or_else(|| Error::InvalidFormat(format!("ODG durable transition {key} is not boolean")))
        .map(Some)
}

fn string_value(operation: &PatchOperation) -> Result<&str> {
    operation
        .value
        .as_str()
        .ok_or_else(|| Error::InvalidFormat("ODG durable patch value is not a string".to_string()))
}

fn geometry_value(value: &serde_json::Value) -> Result<[String; 4]> {
    let values = value.as_array().ok_or_else(|| {
        Error::InvalidFormat("ODG durable geometry value is not an array".to_string())
    })?;
    if values.len() != 4 {
        return invalid("ODG durable geometry value must contain four attributes");
    }
    let parsed = values
        .iter()
        .map(|attribute| {
            attribute.as_str().map(str::to_owned).ok_or_else(|| {
                Error::InvalidFormat("ODG durable geometry attribute is not a string".to_string())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    parsed
        .try_into()
        .map_err(|_values| Error::InvalidFormat("ODG durable geometry is invalid".to_string()))
}

fn durable_blob_id(value: &serde_json::Value) -> Result<&str> {
    let identifier = value.as_str().ok_or_else(|| {
        Error::InvalidFormat("ODG durable package blob identifier is not a string".to_string())
    })?;
    if identifier.len() != 64
        || !identifier
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return invalid("ODG durable package blob identifier is invalid");
    }
    Ok(identifier)
}

fn durable_blob<'patch>(
    patch: &'patch DurablePatch,
    value: &serde_json::Value,
) -> Result<&'patch [u8]> {
    let identifier = durable_blob_id(value)?;
    let blob_id = patch
        .inner
        .blobs()
        .ids()
        .find(|candidate| candidate.as_hex() == identifier)
        .ok_or_else(|| Error::InvalidFormat("ODG durable package blob is missing".to_string()))?;
    patch
        .inner
        .blobs()
        .get(blob_id)
        .ok_or_else(|| Error::InvalidFormat("ODG durable package blob is missing".to_string()))
}

fn apply_durable_patch(
    source: &Package,
    patch: &DurablePatch,
    check_source: bool,
    security_policy: SecurityWritePolicy,
    active_content_policy: ActiveContentWritePolicy,
) -> Result<Package> {
    validate_durable_patch(&patch.inner)?;
    let source_fingerprint = fingerprint(source.as_bytes());
    if check_source
        && patch.operations().iter().any(|operation| {
            operation
                .preconditions
                .get("source")
                .and_then(serde_json::Value::as_str)
                != Some(source_fingerprint.as_str())
        })
    {
        return Err(Error::PatchConflict);
    }
    let mut current = source.clone();
    for operation in patch.operations() {
        if operation.op == "package.replace" {
            enforce_security_policy(&current, security_policy)?;
            enforce_active_content_policy(&current, active_content_policy)?;
            let blob = durable_blob(patch, &operation.value)?;
            current = current.reopen_same_flavor(blob.to_vec())?;
            enforce_active_content_policy(&current, active_content_policy)?;
            continue;
        }
        let mut edit = current.edit_with_policies(security_policy, active_content_policy);
        edit.stage_durable_operation(operation, &source_fingerprint, None)?;
        current = edit.commit_core()?.into_package();
    }
    Ok(current)
}

fn durable_error(error: litchi_core::PatchError) -> Error {
    let message = error.to_string();
    drop(error);
    Error::InvalidFormat(format!("invalid ODG durable patch: {message}"))
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(8, MAX_OUTPUT_BYTES, MAX_OUTPUT_BYTES),
        4 * 1024 * 1024,
        1_024,
        32,
        MAX_TEXT_BYTES,
        32 * 1024 * 1024,
    )
}

fn fingerprint(bytes: &[u8]) -> String {
    DiagnosticFingerprint::of(bytes).as_hex()
}

fn reversible_operation(
    limits: PatchLimits,
    name: &str,
    semantic_target: &str,
    source: &str,
    target: &str,
    before: serde_json::Value,
    after: serde_json::Value,
) -> Result<ReversibleOperation> {
    let forward = PatchOperation::new(
        limits,
        name,
        semantic_target,
        source_precondition(source),
        after,
    )
    .map_err(durable_error)?;
    let inverse = PatchOperation::new(
        limits,
        name,
        semantic_target,
        source_precondition(target),
        before,
    )
    .map_err(durable_error)?;
    Ok(ReversibleOperation::new(forward, inverse))
}

fn source_precondition(source: &str) -> BTreeMap<String, serde_json::Value> {
    BTreeMap::from([(
        "source".to_string(),
        serde_json::Value::String(source.to_string()),
    )])
}

fn parse_shape_target(target: &str) -> Result<(usize, usize)> {
    let Some(remainder) = target.strip_prefix("page/") else {
        return invalid("ODG durable patch target is invalid");
    };
    let Some((page_text, shape_text)) = remainder.split_once("/shape/") else {
        return invalid("ODG durable patch target is invalid");
    };
    if shape_text.contains('/') {
        return invalid("ODG durable patch target is invalid");
    }
    let page = page_text
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat("ODG durable page target is invalid".to_string()))?;
    let shape = shape_text.parse::<usize>().map_err(|_error| {
        Error::InvalidFormat("ODG durable shape target is invalid".to_string())
    })?;
    Ok((page, shape))
}

fn parse_page_target(target: &str) -> Result<usize> {
    let page_text = target
        .strip_prefix("page/")
        .filter(|value| !value.contains('/'))
        .ok_or_else(|| Error::InvalidFormat("ODG durable page target is invalid".to_string()))?;
    page_text
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat("ODG durable page target is invalid".to_string()))
}

fn validate_durable_patch(patch: &CorePatch<Reversible>) -> Result<()> {
    for operation in patch.operations() {
        if !matches!(
            operation.op.as_str(),
            "package.replace"
                | "page.name.set"
                | "page.style.set"
                | "page.transition.set"
                | "shape.geometry.set"
                | "shape.control.set"
                | "shape.layer.set"
                | "shape.name.set"
                | "shape.path.set"
                | "shape.style.set"
                | "shape.text.set"
        ) {
            return invalid("ODG durable patch operation is unsupported");
        }
        if operation.preconditions.len() != 1
            || operation
                .preconditions
                .get("source")
                .and_then(serde_json::Value::as_str)
                .is_none_or(|value| value.len() != 64 || !value.is_ascii())
        {
            return invalid("ODG durable patch source precondition is invalid");
        }
        match operation.op.as_str() {
            "package.replace" => {
                if operation.target != "package" {
                    return invalid("ODG durable package target is invalid");
                }
                let identifier = durable_blob_id(&operation.value)?;
                if !patch
                    .blobs()
                    .ids()
                    .any(|candidate| candidate.as_hex() == identifier)
                {
                    return invalid("ODG durable package blob is missing");
                }
            },
            "shape.geometry.set" => {
                parse_shape_target(&operation.target)?;
                geometry_value(&operation.value)?;
            },
            "page.name.set" | "page.style.set" => {
                parse_page_target(&operation.target)?;
                string_value(operation)?;
            },
            "page.transition.set" => {
                parse_page_target(&operation.target)?;
                transition_from_value(&operation.value)?;
            },
            _ => {
                parse_shape_target(&operation.target)?;
                string_value(operation)?;
            },
        }
    }
    Ok(())
}

fn validate_joined_preconditions(patch: &DurablePatch, source_fingerprint: &str) -> Result<()> {
    if patch.operations().iter().any(|operation| {
        operation
            .preconditions
            .get("source")
            .and_then(serde_json::Value::as_str)
            != Some(source_fingerprint)
    }) {
        return Err(Error::PatchConflict);
    }
    Ok(())
}

fn validate_joined_effect_declarations(joined: &JoinedEdits) -> Result<()> {
    let mut accepted = BTreeSet::new();
    for sub_edit in joined.sub_edits() {
        let declared = sub_edit
            .effects()
            .writes()
            .map(canonical_declared_effect_key)
            .collect::<BTreeSet<_>>();
        let actual = actual_durable_effect_keys(sub_edit.payload())?
            .into_iter()
            .collect::<BTreeSet<_>>();
        if !actual.is_subset(&declared) {
            return Err(joined_refusal(
                "ODG joined payload effects do not match declarations",
            ));
        }
        if actual.iter().any(|effect| !accepted.insert(effect.clone())) {
            return Err(joined_refusal("ODG joined payload effects overlap"));
        }
    }
    Ok(())
}

fn durable_effect_keys(operation: &PatchOperation) -> Result<Vec<String>> {
    let key = match operation.op.as_str() {
        "package.replace" => "package".to_owned(),
        op @ ("page.name.set" | "page.style.set" | "page.transition.set") => {
            format!("page/{}#{op}", parse_page_target(&operation.target)?)
        },
        op @ ("shape.geometry.set" | "shape.control.set" | "shape.layer.set" | "shape.name.set"
        | "shape.path.set" | "shape.style.set" | "shape.text.set") => {
            let (page, shape) = parse_shape_target(&operation.target)?;
            format!("page/{page}/shape/{shape}#{op}")
        },
        _ => format!("{}#{}", operation.target, operation.op),
    };
    Ok(vec![key])
}

fn actual_durable_effect_keys(patch: &DurablePatch) -> Result<Vec<String>> {
    patch
        .operations()
        .iter()
        .try_fold(Vec::new(), |mut effects, operation| {
            effects.extend(durable_effect_keys(operation)?);
            Ok(effects)
        })
}

fn canonical_declared_effect_key(effect: &str) -> String {
    if effect == "package" {
        return effect.to_owned();
    }
    let Some((target, operation)) = effect.rsplit_once('#') else {
        return effect.to_owned();
    };
    match operation {
        "page.name.set" | "page.style.set" | "page.transition.set" => parse_page_target(target)
            .map(|page| format!("page/{page}#{operation}"))
            .unwrap_or_else(|_| effect.to_owned()),
        "shape.geometry.set" | "shape.control.set" | "shape.layer.set" | "shape.name.set"
        | "shape.path.set" | "shape.style.set" | "shape.text.set" => parse_shape_target(target)
            .map(|(page, shape)| format!("page/{page}/shape/{shape}#{operation}"))
            .unwrap_or_else(|_| effect.to_owned()),
        _ => effect.to_owned(),
    }
}

const ORDINARY_STYLE_GROUP: usize = usize::MAX;

#[derive(Clone)]
struct StyleProjection {
    pages: Vec<Option<String>>,
    parents: Arc<BTreeMap<String, Option<String>>>,
}

impl StyleProjection {
    fn from_package(package: &Package) -> Result<Self> {
        Ok(Self {
            pages: package
                .pages()
                .iter()
                .map(|page| page.style_name().map(str::to_owned))
                .collect(),
            parents: Arc::new(drawing_page_style_parents(
                package.content_xml(),
                package.styles_xml(),
            )?),
        })
    }

    fn page_style(&self, page: usize) -> Option<String> {
        self.pages.get(page).cloned().flatten()
    }

    fn set_page_style(&mut self, page: usize, style: String) {
        if let Some(current) = self.pages.get_mut(page) {
            *current = Some(style);
        }
    }
}

#[derive(Clone)]
enum StyleDependencyEvent {
    Retarget {
        group: usize,
        page: usize,
        source_style: Option<String>,
        destination: String,
    },
    Transition {
        group: usize,
        page: usize,
        source_style: Option<String>,
        parents: Arc<BTreeMap<String, Option<String>>>,
    },
}

fn joined_has_dependent_conflict(source: &Package, joined: &JoinedEdits) -> Result<bool> {
    if !joined
        .sub_edits()
        .any(|sub_edit| has_style_dependency_operation(sub_edit.payload().operations()))
    {
        return Ok(false);
    }
    // Cross-sub-edit package/semantic mixing has already been rejected by
    // `joined_contains_mixed_package_effects`. Any remaining style/transition
    // operations around a replacement are in one ordered sub-edit, where
    // intra-sub-edit dependencies are intentional and valid. Do not reopen
    // replacement blobs merely to walk a dependency graph that cannot cross a
    // composition boundary.
    if joined_contains_package_replacement(joined) {
        return Ok(false);
    }
    let events = projected_joined_style_events(source, joined)?;
    style_events_conflict(&effective_style_events(events))
}

fn projected_staged_style_conflict(edit: &Edit, joined: &JoinedEdits) -> Result<bool> {
    if !edit
        .changes
        .iter()
        .any(|change| matches!(change, Change::PageStyle(_) | Change::PageTransition(_)))
    {
        return Ok(false);
    }
    let mut projection = StyleProjection::from_package(&edit.source)?;
    let mut events = Vec::new();
    for change in &edit.changes {
        match change {
            Change::PageStyle(value) => {
                events.push(StyleDependencyEvent::Retarget {
                    group: ORDINARY_STYLE_GROUP,
                    page: value.page,
                    source_style: projection.page_style(value.page),
                    destination: value.after.clone(),
                });
                projection.set_page_style(value.page, value.after.clone());
            },
            Change::PageTransition(value) => {
                events.push(StyleDependencyEvent::Transition {
                    group: ORDINARY_STYLE_GROUP,
                    page: value.page,
                    source_style: projection.page_style(value.page),
                    parents: projection.parents.clone(),
                });
            },
            _ => {},
        }
    }
    events.extend(projected_joined_style_events_from_projection(
        edit.source.clone(),
        projection,
        joined,
    )?);
    style_events_conflict(&effective_style_events(events))
}

fn has_style_dependency_operation<'a>(
    operations: impl IntoIterator<Item = &'a PatchOperation>,
) -> bool {
    operations.into_iter().any(|operation| {
        matches!(
            operation.op.as_str(),
            "page.style.set" | "page.transition.set"
        )
    })
}

fn joined_contains_package_replacement(joined: &JoinedEdits) -> bool {
    joined.sub_edits().any(|sub_edit| {
        sub_edit
            .payload()
            .operations()
            .iter()
            .any(|operation| operation.op == "package.replace")
    })
}

fn projected_joined_style_events(
    source: &Package,
    joined: &JoinedEdits,
) -> Result<Vec<StyleDependencyEvent>> {
    projected_joined_style_events_from_projection(
        source.clone(),
        StyleProjection::from_package(source)?,
        joined,
    )
}

fn projected_joined_style_events_from_projection(
    mut active_package: Package,
    mut projection: StyleProjection,
    joined: &JoinedEdits,
) -> Result<Vec<StyleDependencyEvent>> {
    let mut events = Vec::new();
    for (group, sub_edit) in joined.sub_edits().enumerate() {
        for operation in sub_edit.payload().operations() {
            if operation.op == "package.replace" {
                let blob = durable_blob(sub_edit.payload(), &operation.value)?;
                active_package = active_package.reopen_same_flavor(blob.to_vec())?;
                projection = StyleProjection::from_package(&active_package)?;
                continue;
            }
            match operation.op.as_str() {
                "page.style.set" => {
                    let page = parse_page_target(&operation.target)?;
                    let destination = string_value(operation)?.to_owned();
                    events.push(StyleDependencyEvent::Retarget {
                        group,
                        page,
                        source_style: projection.page_style(page),
                        destination: destination.clone(),
                    });
                    projection.set_page_style(page, destination);
                },
                "page.transition.set" => {
                    let page = parse_page_target(&operation.target)?;
                    events.push(StyleDependencyEvent::Transition {
                        group,
                        page,
                        source_style: projection.page_style(page),
                        parents: projection.parents.clone(),
                    });
                },
                _ => {},
            }
        }
    }
    Ok(events)
}

fn effective_style_events(events: Vec<StyleDependencyEvent>) -> Vec<StyleDependencyEvent> {
    // Cross-group dependency checks observe each group's effective final style
    // state.  Retargets overwritten later in the same group are not external
    // dependencies; transitions remain because they observe the intermediate
    // style at their position in the ordered payload.
    let mut initial_styles = BTreeMap::new();
    let mut final_retargets = BTreeMap::new();
    for (index, event) in events.iter().enumerate() {
        if let StyleDependencyEvent::Retarget {
            group,
            page,
            source_style,
            ..
        } = event
        {
            initial_styles
                .entry((*group, *page))
                .or_insert_with(|| source_style.clone());
            final_retargets.insert((*group, *page), index);
        }
    }
    events
        .into_iter()
        .enumerate()
        .filter_map(|(index, event)| match &event {
            StyleDependencyEvent::Retarget {
                group,
                page,
                source_style: _,
                destination,
            } => {
                if final_retargets.get(&(*group, *page)) != Some(&index) {
                    return None;
                }
                if initial_styles
                    .get(&(*group, *page))
                    .and_then(Option::as_deref)
                    == Some(destination.as_str())
                {
                    return None;
                }
                Some(event)
            },
            StyleDependencyEvent::Transition { .. } => Some(event),
        })
        .collect()
}

fn style_events_conflict(events: &[StyleDependencyEvent]) -> Result<bool> {
    for (index, left) in events.iter().enumerate() {
        for right in events[index.saturating_add(1)..].iter() {
            if style_event_group(left) == style_event_group(right) {
                continue;
            }
            if style_event_pair_conflicts(left, right)? {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn style_event_group(event: &StyleDependencyEvent) -> usize {
    match event {
        StyleDependencyEvent::Retarget { group, .. }
        | StyleDependencyEvent::Transition { group, .. } => *group,
    }
}

fn style_event_pair_conflicts(
    left: &StyleDependencyEvent,
    right: &StyleDependencyEvent,
) -> Result<bool> {
    match (left, right) {
        (
            StyleDependencyEvent::Retarget {
                page: left_page,
                destination,
                ..
            },
            StyleDependencyEvent::Transition {
                page: right_page,
                source_style,
                parents,
                ..
            },
        )
        | (
            StyleDependencyEvent::Transition {
                page: right_page,
                source_style,
                parents,
                ..
            },
            StyleDependencyEvent::Retarget {
                page: left_page,
                destination,
                ..
            },
        ) => {
            if left_page == right_page {
                return Ok(true);
            }
            match source_style.as_deref() {
                Some(source_style) => style_chain_contains(parents, destination, source_style),
                None => Ok(false),
            }
        },
        (
            StyleDependencyEvent::Transition {
                source_style: left_style,
                parents: left_parents,
                ..
            },
            StyleDependencyEvent::Transition {
                source_style: right_style,
                parents: right_parents,
                ..
            },
        ) => {
            let Some(left_style) = left_style.as_deref() else {
                return Ok(false);
            };
            let Some(right_style) = right_style.as_deref() else {
                return Ok(false);
            };
            if left_style == right_style {
                return Ok(true);
            }
            if style_chain_contains(left_parents, left_style, right_style)? {
                return Ok(true);
            }
            style_chain_contains(right_parents, right_style, left_style)
        },
        _ => Ok(false),
    }
}

fn joined_contains_mixed_package_effects(joined: &JoinedEdits) -> Result<bool> {
    let sub_edits = joined
        .sub_edits()
        .map(|sub_edit| actual_durable_effect_keys(sub_edit.payload()))
        .collect::<Result<Vec<_>>>()?;
    Ok(sub_edits.iter().enumerate().any(|(index, left)| {
        let left_package = left.iter().any(|effect| effect == "package");
        let left_other = left.iter().any(|effect| effect != "package");
        sub_edits[index.saturating_add(1)..].iter().any(|right| {
            let right_package = right.iter().any(|effect| effect == "package");
            let right_other = right.iter().any(|effect| effect != "package");
            (left_package && right_other) || (left_other && right_package)
        })
    }))
}

struct Parsed {
    control_spans: ControlSpans,
    form_controls: Vec<FormControl>,
    form_control_spans: Vec<Option<Range<usize>>>,
    forms_insert_position: Option<usize>,
    drawing_start_position: usize,
    pages: Vec<Page>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    layer_count: usize,
    geometry_spans: GeometrySpans,
    line_geometry_spans: GeometrySpans,
    points_spans: PointsSpans,
    transform_spans: Vec<Vec<Option<Range<usize>>>>,
    path_spans: PathSpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    page_spans: Vec<Option<Range<usize>>>,
    page_attribute_spans: PageAttributeSpans,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_insert_positions: Vec<Option<usize>>,
    drawing_insert_position: usize,
}

fn group_selection(parsed: &Parsed, page: usize, shape: usize) -> Result<Group> {
    let selected = parsed
        .pages
        .get(page)
        .and_then(|value| value.shapes().get(shape))
        .ok_or_else(|| Error::InvalidFormat("ODG group selector is out of bounds".into()))?;
    if selected.kind() != ShapeKind::Group {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odg.package.snapshot",
            "ODG group operation requires a draw:g root",
        )));
    }
    let root = parsed.shape_spans[page][shape]
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("ODG group source span is missing".into()))?;
    let descendants = parsed.shape_spans[page]
        .iter()
        .enumerate()
        .filter(|(position, span)| {
            *position != shape
                && span.as_ref().is_some_and(|candidate| {
                    candidate.start >= root.start && candidate.end <= root.end
                })
        })
        .map(|(position, _span)| position)
        .collect();
    Ok(Group::parsed(page, shape, descendants))
}

struct ActiveShape {
    depth: usize,
    page: usize,
    shape: usize,
    start: usize,
    kind: ShapeKind,
}

struct ActiveFormControl {
    depth: usize,
    control: usize,
    start: usize,
}

#[derive(Clone, Copy)]
enum AccessibilityKind {
    Description,
    Title,
}

struct ActiveAccessibility {
    depth: usize,
    page: usize,
    shape: usize,
    kind: AccessibilityKind,
}

struct Scanner {
    depth: usize,
    root_seen: bool,
    body_seen: bool,
    drawing_seen: bool,
    body_depth: Option<usize>,
    drawing_depth: Option<usize>,
    drawing_start_position: Option<usize>,
    forms_depth: Option<usize>,
    forms_insert_position: Option<usize>,
    form_controls: Vec<FormControl>,
    form_control_spans: Vec<Option<Range<usize>>>,
    active_form_controls: Vec<ActiveFormControl>,
    pages: Vec<Page>,
    page_depths: Vec<usize>,
    page_starts: Vec<usize>,
    layer_sets: Vec<(usize, Option<usize>)>,
    active_shapes: Vec<ActiveShape>,
    active_accessibility: Option<ActiveAccessibility>,
    control_spans: ControlSpans,
    paragraph_depths: Vec<usize>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    geometry_spans: GeometrySpans,
    line_geometry_spans: GeometrySpans,
    points_spans: PointsSpans,
    transform_spans: Vec<Vec<Option<Range<usize>>>>,
    path_spans: PathSpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_count: usize,
    shape_count: usize,
    text_bytes: usize,
    page_spans: Vec<Option<Range<usize>>>,
    page_attribute_spans: PageAttributeSpans,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_starts: Vec<(usize, usize, usize)>,
    layer_set_insert_positions: Vec<Option<usize>>,
    active_layers: Vec<(usize, usize, usize, usize)>,
    drawing_insert_position: Option<usize>,
}

impl Scanner {
    fn new() -> Self {
        Self {
            depth: 0,
            root_seen: false,
            body_seen: false,
            drawing_seen: false,
            body_depth: None,
            drawing_depth: None,
            drawing_start_position: None,
            forms_depth: None,
            forms_insert_position: None,
            form_controls: Vec::new(),
            form_control_spans: Vec::new(),
            active_form_controls: Vec::new(),
            pages: Vec::new(),
            page_depths: Vec::new(),
            page_starts: Vec::new(),
            layer_sets: Vec::new(),
            active_shapes: Vec::new(),
            active_accessibility: None,
            control_spans: Vec::new(),
            paragraph_depths: Vec::new(),
            text_spans: Vec::new(),
            name_spans: Vec::new(),
            layer_spans: Vec::new(),
            geometry_spans: Vec::new(),
            line_geometry_spans: Vec::new(),
            points_spans: Vec::new(),
            transform_spans: Vec::new(),
            path_spans: Vec::new(),
            style_name_spans: Vec::new(),
            layer_count: 0,
            shape_count: 0,
            text_bytes: 0,
            page_spans: Vec::new(),
            page_attribute_spans: Vec::new(),
            page_insert_positions: Vec::new(),
            shape_spans: Vec::new(),
            layer_element_spans: Vec::new(),
            layer_set_starts: Vec::new(),
            layer_set_insert_positions: Vec::new(),
            active_layers: Vec::new(),
            drawing_insert_position: None,
        }
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        self.depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
        if self.depth > MAX_DEPTH {
            return invalid("ODG XML nesting exceeds the limit");
        }
        self.observe(reader, namespace, element, tag, tag_start, empty)?;
        if empty {
            self.depth = self.depth.saturating_sub(1);
        }
        Ok(())
    }

    fn observe(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        let local_name = element.local_name();
        let local = local_name.as_ref();
        if self.depth == 1 {
            if self.root_seen
                || namespace != NamespaceKind::Office
                || local != b"document-content"
                || empty
            {
                return invalid("ODG content.xml requires one office:document-content root");
            }
            self.root_seen = true;
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"body" {
            if self.body_seen || self.depth != 2 || empty {
                return invalid("ODG content.xml requires one non-empty office:body");
            }
            self.body_seen = true;
            self.body_depth = Some(self.depth);
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"forms" {
            if self.forms_depth.is_some()
                || self.forms_insert_position.is_some()
                || self.drawing_depth != Some(self.depth - 1)
            {
                return invalid("ODG office:forms is misplaced or duplicated");
            }
            if empty {
                self.forms_insert_position = Some(tag_start + tag.len() - 2);
            } else {
                self.forms_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"drawing" {
            if self.drawing_seen || self.body_depth != Some(self.depth - 1) {
                return invalid("ODG office:drawing is misplaced or duplicated");
            }
            self.drawing_seen = true;
            self.drawing_start_position = Some(if empty {
                tag_start + tag.len() - 2
            } else {
                tag_start + tag.len()
            });
            if empty {
                self.drawing_insert_position = Some(tag_start + tag.len() - 2);
            } else {
                self.drawing_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Form
            && self
                .forms_depth
                .is_some_and(|forms_depth| self.depth > forms_depth)
            && let Some(identifier) = attribute(reader, element, FORM, b"id")?
        {
            if self.form_controls.len() >= MAX_FORM_CONTROLS {
                return invalid("ODG form-control count exceeds the limit");
            }
            if self
                .form_controls
                .iter()
                .any(|control| control.id() == identifier)
            {
                return invalid("ODG form-control identifier is duplicated");
            }
            let control = self.form_controls.len();
            let element_name = std::str::from_utf8(local).map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG form-control element name: {error}"))
            })?;
            let element_name = clone_bounded_string(
                element_name,
                MAX_TEXT_BYTES,
                "ODG form-control element name",
            )?;
            self.form_controls.push(FormControl::parsed(
                identifier,
                attribute(reader, element, FORM, b"name")?,
                element_name,
                arbitrary_attributes(
                    reader,
                    element,
                    &[(FORM, b"id".as_slice()), (FORM, b"name".as_slice())],
                )?,
            ));
            self.form_control_spans
                .push(empty.then_some(tag_start..tag_start + tag.len()));
            if !empty {
                self.active_form_controls.push(ActiveFormControl {
                    depth: self.depth,
                    control,
                    start: tag_start,
                });
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"layer-set" {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer-set is outside draw:page".to_string())
            })?;
            self.pages[page].mark_layer_set();
            if empty {
                self.layer_set_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.layer_sets.push((self.depth, Some(page)));
                self.layer_set_starts.push((self.depth, page, tag_start));
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer"
            && self
                .layer_sets
                .last()
                .is_some_and(|(depth, _)| *depth + 1 == self.depth)
        {
            self.add_layer(reader, element)?;
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer is outside draw:page".to_string())
            })?;
            let layer = self.pages[page].layers().len() - 1;
            if empty {
                self.layer_element_spans[page][layer] = Some(tag_start..tag_start + tag.len());
            } else {
                self.active_layers
                    .push((self.depth, page, layer, tag_start));
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"page" {
            if self.drawing_depth != Some(self.depth - 1) {
                return invalid("ODG draw:page is outside office:drawing");
            }
            if self.pages.len() >= MAX_PAGES {
                return invalid("ODG page count exceeds the limit");
            }
            self.pages.push(Page::parsed(
                attribute(reader, element, DRAW, b"name")?,
                attribute(reader, element, XML, b"id")?,
                attribute(reader, element, DRAW, b"style-name")?,
                attribute(reader, element, DRAW, b"master-page-name")?,
                None,
            ));
            self.text_spans.push(Vec::new());
            self.control_spans.push(Vec::new());
            self.name_spans.push(Vec::new());
            self.layer_spans.push(Vec::new());
            self.geometry_spans.push(Vec::new());
            self.line_geometry_spans.push(Vec::new());
            self.points_spans.push(Vec::new());
            self.transform_spans.push(Vec::new());
            self.path_spans.push(Vec::new());
            self.style_name_spans.push(Vec::new());
            self.page_spans.push(None);
            self.page_attribute_spans.push([
                attribute_source_span(reader, element, tag, tag_start, DRAW, b"name")?,
                attribute_source_span(reader, element, tag, tag_start, DRAW, b"style-name")?,
            ]);
            self.page_insert_positions.push(None);
            self.shape_spans.push(Vec::new());
            self.layer_element_spans.push(Vec::new());
            self.layer_set_insert_positions.push(None);
            if empty {
                let page = self.pages.len() - 1;
                self.page_spans[page] = Some(tag_start..tag_start + tag.len());
                self.page_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.page_depths.push(self.depth);
                self.page_starts.push(tag_start);
            }
            return Ok(());
        }
        if let Some(kind) = shape_kind(namespace, local) {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG drawing shape is outside draw:page".to_string())
            })?;
            let direct_parent = self
                .active_shapes
                .last()
                .filter(|parent| parent.depth.checked_add(1) == Some(self.depth));
            if kind == ShapeKind::ThreeDimensionalScene {
                if let Some(parent) = direct_parent
                    && !matches!(
                        parent.kind,
                        ShapeKind::Group | ShapeKind::ThreeDimensionalScene
                    )
                {
                    return invalid("ODG dr3d:scene has an invalid parent");
                }
            } else if kind.is_three_dimensional()
                && !direct_parent
                    .is_some_and(|parent| parent.kind == ShapeKind::ThreeDimensionalScene)
            {
                return invalid("ODG 3D drawing objects require a dr3d:scene parent");
            } else if self
                .active_shapes
                .last()
                .is_some_and(|parent| parent.kind == ShapeKind::ThreeDimensionalScene)
                && !kind.is_three_dimensional()
            {
                return invalid("ODG dr3d:scene contains a non-3D drawing object");
            }
            if kind.is_three_dimensional() {
                validate_three_dimensional_attributes(reader, element, kind)?;
            }
            if kind == ShapeKind::ThreeDimensionalLight {
                let direction =
                    required_attribute(reader, element, DR3D, b"direction", "dr3d:light")?;
                if !is_vector3d(&direction) {
                    return invalid("ODG dr3d:light direction is not a vector3D");
                }
            }
            if matches!(
                kind,
                ShapeKind::ThreeDimensionalExtrude | ShapeKind::ThreeDimensionalRotate
            ) {
                let view_box =
                    required_attribute(reader, element, SVG, b"viewBox", "dr3d path shape")?;
                let _path_data = required_attribute(reader, element, SVG, b"d", "dr3d path shape")?;
                if !is_integer_list(&view_box, 4) {
                    return invalid("ODG 3D path shape viewBox is not four integers");
                }
            }
            if self.shape_count >= MAX_SHAPES {
                return invalid("ODG shape count exceeds the limit");
            }
            self.shape_count += 1;
            let name = attribute(reader, element, DRAW, b"name")?;
            let page_name = self.pages[page].name().map(str::to_string);
            let frame = if kind == ShapeKind::Frame {
                Some(frame(reader, element, name.clone(), page_name)?)
            } else {
                None
            };
            let z_index = optional_u32_attribute(reader, element, DRAW, b"z-index")?;
            // Geometry and line endpoints are one ownership boundary: both
            // groups are needed before lexical validation and retain the
            // legacy request order for the ordered fallback below.
            let [x, y, width, height, x1, y1, x2, y2] = attribute_values(
                reader,
                element,
                [
                    (SVG, b"x"),
                    (SVG, b"y"),
                    (SVG, b"width"),
                    (SVG, b"height"),
                    (SVG, b"x1"),
                    (SVG, b"y1"),
                    (SVG, b"x2"),
                    (SVG, b"y2"),
                ],
            )?;
            let geometry = [x, y, width, height];
            let line_geometry = [x1, y1, x2, y2];
            validate_shape_lexical_attributes(
                kind,
                &geometry,
                &line_geometry,
                attribute(reader, element, SVG, b"viewBox")?.as_deref(),
                attribute(reader, element, DRAW, b"points")?.as_deref(),
            )?;
            let shape = self.pages[page].shapes().len();
            // These are the eight remaining ShapeProperties values. Keep
            // this batch after lexical validation so validation still reads
            // viewBox and points in its established order before ownership
            // moves into the parsed shape.
            let [
                control_reference,
                layer,
                path_data,
                transform,
                points,
                view_box,
                style_name,
                text_style_name,
            ] = attribute_values(
                reader,
                element,
                [
                    (DRAW, b"control"),
                    (DRAW, b"layer"),
                    (SVG, b"d"),
                    (DRAW, b"transform"),
                    (DRAW, b"points"),
                    (SVG, b"viewBox"),
                    (DRAW, b"style-name"),
                    (DRAW, b"text-style-name"),
                ],
            )?;
            self.pages[page].push_shape(Shape::parsed(
                ShapeProperties {
                    control_reference,
                    geometry,
                    layer,
                    name,
                    path_data,
                    transform,
                    points,
                    view_box,
                    line_geometry,
                    style_name,
                    text_style_name,
                    z_index,
                },
                kind,
                frame,
            ));
            self.text_spans[page].push(Vec::new());
            let [
                control_span,
                name_span,
                layer_span,
                x_span,
                y_span,
                width_span,
                height_span,
                x1_span,
                y1_span,
                x2_span,
                y2_span,
                view_box_span,
                points_span,
                transform_span,
                path_span,
                style_name_span,
            ] = shape_attribute_source_spans(reader, element, tag, tag_start)?;
            self.control_spans[page].push(control_span);
            self.name_spans[page].push(name_span);
            self.layer_spans[page].push(layer_span);
            self.geometry_spans[page].push([x_span, y_span, width_span, height_span]);
            self.line_geometry_spans[page].push([x1_span, y1_span, x2_span, y2_span]);
            self.points_spans[page].push([view_box_span, points_span]);
            self.transform_spans[page].push(transform_span);
            self.path_spans[page].push(path_span);
            self.style_name_spans[page].push(style_name_span);
            self.shape_spans[page].push(empty.then_some(tag_start..tag_start + tag.len()));
            if !empty {
                self.active_shapes.push(ActiveShape {
                    depth: self.depth,
                    page,
                    shape,
                    start: tag_start,
                    kind,
                });
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Svg
            && matches!(local, b"title" | b"desc")
            && !empty
            && let Some(active) = self.active_shapes.last()
            && active.depth + 1 == self.depth
        {
            self.active_accessibility = Some(ActiveAccessibility {
                depth: self.depth,
                page: active.page,
                shape: active.shape,
                kind: if local == b"title" {
                    AccessibilityKind::Title
                } else {
                    AccessibilityKind::Description
                },
            });
            return Ok(());
        }
        if !self.active_shapes.is_empty()
            && namespace == NamespaceKind::Text
            && local == b"p"
            && !empty
        {
            self.paragraph_depths.push(self.depth);
        }
        Ok(())
    }

    fn add_layer(&mut self, reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
        if self.layer_count >= MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        self.layer_count += 1;
        let name = required_attribute(reader, element, DRAW, b"name", "draw:layer")?;
        let protected = optional_bool_attribute(reader, element, DRAW, b"protected")?;
        let layer = Layer::parsed(
            name,
            attribute(reader, element, DRAW, b"display")?,
            protected,
        );
        if let Some((_, Some(page))) = self.layer_sets.last() {
            self.pages[*page].push_layer(layer.clone());
            self.layer_element_spans[*page].push(None);
        }
        Ok(())
    }

    fn text(&mut self, span: Option<Range<usize>>, value: &str) -> Result<()> {
        if let Some(accessibility) = &self.active_accessibility {
            self.text_bytes = self.text_bytes.checked_add(value.len()).ok_or_else(|| {
                Error::InvalidFormat("ODG text extraction size overflow".to_string())
            })?;
            if self.text_bytes > MAX_TEXT_BYTES {
                return invalid("ODG text extraction exceeds the limit");
            }
            let shape = self.pages[accessibility.page]
                .shape_mut(accessibility.shape)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG active accessibility shape disappeared".to_string())
                })?;
            match accessibility.kind {
                AccessibilityKind::Description => shape.push_description(value),
                AccessibilityKind::Title => shape.push_title(value),
            }
            return Ok(());
        }
        if self.paragraph_depths.is_empty() {
            return Ok(());
        }
        let Some(active) = self.active_shapes.last() else {
            return Ok(());
        };
        let (page, shape_index) = (active.page, active.shape);
        self.text_bytes = self
            .text_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("ODG text extraction size overflow".to_string()))?;
        if self.text_bytes > MAX_TEXT_BYTES {
            return invalid("ODG text extraction exceeds the limit");
        }
        let shape = self.pages[page]
            .shape_mut(shape_index)
            .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
        shape.push_text(value);
        self.text_spans[page][shape_index].push(span);
        Ok(())
    }

    fn end(
        &mut self,
        namespace: NamespaceKind,
        local: &[u8],
        tag_start: usize,
        tag_end: usize,
    ) -> Result<()> {
        if self
            .active_form_controls
            .last()
            .is_some_and(|control| control.depth == self.depth)
        {
            let active = self.active_form_controls.pop().ok_or_else(|| {
                Error::InvalidFormat("ODG active form control disappeared".into())
            })?;
            self.form_control_spans[active.control] = Some(active.start..tag_end);
        }
        if self
            .active_accessibility
            .as_ref()
            .is_some_and(|active| active.depth == self.depth)
        {
            self.active_accessibility = None;
        }
        if self.paragraph_depths.last() == Some(&self.depth)
            && namespace == NamespaceKind::Text
            && local == b"p"
        {
            self.paragraph_depths.pop();
        }
        if self
            .active_shapes
            .last()
            .is_some_and(|shape| shape.depth == self.depth)
        {
            let active = self
                .active_shapes
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
            self.shape_spans[active.page][active.shape] = Some(active.start..tag_end);
        }
        if self
            .active_layers
            .last()
            .is_some_and(|layer| layer.0 == self.depth)
        {
            let (_, page, layer, start) = self
                .active_layers
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active layer disappeared".to_string()))?;
            self.layer_element_spans[page][layer] = Some(start..tag_end);
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer-set"
            && self
                .layer_sets
                .last()
                .is_some_and(|set| set.0 == self.depth)
        {
            self.layer_sets.pop();
            let (_, page, _) = self.layer_set_starts.pop().ok_or_else(|| {
                Error::InvalidFormat("ODG active layer-set disappeared".to_string())
            })?;
            self.layer_set_insert_positions[page] = Some(tag_start);
        }
        if namespace == NamespaceKind::Draw
            && local == b"page"
            && self.page_depths.last() == Some(&self.depth)
        {
            self.page_depths.pop();
            let start = self
                .page_starts
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active page disappeared".to_string()))?;
            let page = self.pages.len() - 1;
            self.page_spans[page] = Some(start..tag_end);
            self.page_insert_positions[page] = Some(tag_start);
        }
        if self.drawing_depth == Some(self.depth) {
            self.drawing_depth = None;
            self.drawing_insert_position = Some(tag_start);
        }
        if namespace == NamespaceKind::Office
            && local == b"forms"
            && self.forms_depth == Some(self.depth)
        {
            self.forms_depth = None;
            self.forms_insert_position = Some(tag_start);
        }
        if self.body_depth == Some(self.depth) {
            self.body_depth = None;
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth underflow".to_string()))?;
        Ok(())
    }

    fn current_page(&self) -> Option<usize> {
        self.page_depths
            .last()
            .filter(|depth| self.depth > **depth)
            .map(|_| self.pages.len() - 1)
    }

    fn finish(self) -> Result<Parsed> {
        if self.depth != 0
            || !self.root_seen
            || !self.body_seen
            || !self.drawing_seen
            || self.body_depth.is_some()
            || self.drawing_depth.is_some()
            || self.forms_depth.is_some()
            || !self.page_depths.is_empty()
            || !self.layer_sets.is_empty()
            || !self.active_form_controls.is_empty()
            || self.active_accessibility.is_some()
        {
            return invalid("ODG content.xml has an incomplete drawing structure");
        }
        Ok(Parsed {
            control_spans: self.control_spans,
            form_controls: self.form_controls,
            form_control_spans: self.form_control_spans,
            forms_insert_position: self.forms_insert_position,
            drawing_start_position: self.drawing_start_position.ok_or_else(|| {
                Error::InvalidFormat("ODG drawing start position is missing".to_string())
            })?,
            pages: self.pages,
            text_spans: self.text_spans,
            name_spans: self.name_spans,
            layer_spans: self.layer_spans,
            geometry_spans: self.geometry_spans,
            line_geometry_spans: self.line_geometry_spans,
            points_spans: self.points_spans,
            transform_spans: self.transform_spans,
            path_spans: self.path_spans,
            style_name_spans: self.style_name_spans,
            layer_count: self.layer_count,
            page_spans: self.page_spans,
            page_attribute_spans: self.page_attribute_spans,
            page_insert_positions: self.page_insert_positions,
            shape_spans: self.shape_spans,
            layer_element_spans: self.layer_element_spans,
            layer_set_insert_positions: self.layer_set_insert_positions,
            drawing_insert_position: self.drawing_insert_position.ok_or_else(|| {
                Error::InvalidFormat("ODG drawing insertion point is missing".to_string())
            })?,
        })
    }
}

fn parse_content(xml: &str) -> Result<Parsed> {
    validate_raw_xml_events(xml, "ODG content.xml event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner::new();
    let mut has_enhanced_geometry = false;
    let mut has_auxiliary_shapes = false;
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG content.xml: {error}")))?;
        let namespace = classify(&resolved_namespace);
        validate_borrowed_xml_event(&borrowed_event, "ODG content.xml event")?;
        if let Event::Start(element) | Event::Empty(element) = &borrowed_event {
            let local = element.local_name();
            // Discover optional inventories during the mandatory namespace-aware
            // scan. Misplaced owners still select their validating pass, and
            // producer aliases or namespace rebinding cannot hide them. A pass
            // that is skipped here has nothing to validate beyond the raw
            // admission that already covered this whole part above.
            has_enhanced_geometry |= namespace == NamespaceKind::Draw
                && matches!(
                    local.as_ref(),
                    b"enhanced-geometry" | b"equation" | b"handle"
                );
            has_auxiliary_shapes |= namespace == NamespaceKind::Dr3d
                || namespace == NamespaceKind::Draw
                    && matches!(
                        local.as_ref(),
                        b"frame"
                            | b"image"
                            | b"image-map"
                            | b"area-rectangle"
                            | b"area-circle"
                            | b"area-polygon"
                            | b"contour-polygon"
                            | b"contour-path"
                            | b"glue-point"
                    );
        }
        let end = position(&reader)?;
        match borrowed_event {
            Event::Start(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                false,
            )?,
            Event::Empty(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                true,
            )?,
            Event::End(element) => {
                scanner.end(namespace, element.local_name().as_ref(), start, end)?;
            },
            Event::Text(text) => {
                let value = text_value(&text)?;
                scanner.text(Some(start..end), &value)?;
            },
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::InvalidFormat(format!("invalid ODG CDATA: {error}")))?;
                scanner.text(None, &value)?;
            },
            Event::GeneralRef(reference) => {
                let value = reference_value(&reference)?;
                scanner.text(None, &value)?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG content.xml"),
            Event::Eof => {
                let mut parsed = scanner.finish()?;
                if has_enhanced_geometry {
                    parse_enhanced_geometry(xml, &mut parsed)?;
                }
                if has_auxiliary_shapes {
                    parse_shape_auxiliary_children(xml, &mut parsed)?;
                }
                return Ok(parsed);
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
}

/// Admit one borrowed quick-xml event before any parser path can call
/// `into_owned` or retain one of its fields. The reader normally exposes
/// borrowed slices, but `Event::into_owned` clones every variant, including
/// comments, processing instructions, and oversized start-tag attributes that
/// this family ignores. Checking the raw event and each raw attribute here
/// keeps those ignored inputs bounded just like semantic fields.
fn validate_borrowed_xml_event(event: &Event<'_>, owner: &str) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            validate_raw_xml_value(element.as_ref(), owner)?;
            for raw_attribute in element.attributes() {
                let attribute = raw_attribute.map_err(|error| {
                    Error::InvalidFormat(format!("invalid {owner} attribute: {error}"))
                })?;
                validate_raw_xml_value(attribute.key.as_ref(), owner)?;
                validate_raw_xml_value(attribute.value.as_ref(), owner)?;
            }
        },
        Event::End(element) => validate_raw_xml_value(element.as_ref(), owner)?,
        Event::Text(text) | Event::Comment(text) | Event::DocType(text) => {
            validate_raw_xml_value(text.as_ref(), owner)?;
        },
        Event::CData(text) => validate_raw_xml_value(text.as_ref(), owner)?,
        Event::Decl(declaration) => validate_raw_xml_value(declaration.as_ref(), owner)?,
        Event::PI(instruction) => validate_raw_xml_value(instruction.as_ref(), owner)?,
        Event::GeneralRef(reference) => validate_raw_xml_value(reference.as_ref(), owner)?,
        Event::Eof => {},
    }
    Ok(())
}

/// Admit every raw event before handing the same XML to `NsReader`.
///
/// `NsReader::read_event` updates its namespace stack while reading a start
/// tag, and `read_resolved_event` can clone an unknown prefix while resolving
/// the element name. The raw reader has neither behaviour: it only borrows
/// the event from the source. Running this admission pass first therefore
/// makes the per-event byte limit apply before namespace bookkeeping can own a
/// prefix or declaration value.
fn validate_raw_xml_events(xml: &str, owner: &str) -> Result<()> {
    validate_xml_source_size(xml, owner)?;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid {owner}: {error}")))?;
        validate_borrowed_xml_event(&event, owner)?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat(format!("{owner} depth overflow")))?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "{owner} nesting exceeds the limit"
                    )));
                }
            },
            Event::Empty(_) => {
                let virtual_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat(format!("{owner} depth overflow")))?;
                if virtual_depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "{owner} nesting exceeds the limit"
                    )));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
            Event::Eof => break,
        }
    }
    Ok(())
}

struct ActiveEnhancedGeometry {
    page: usize,
    shape: usize,
    depth: usize,
    attributes: Vec<DrawingAttribute>,
    children: Vec<EnhancedGeometryChild>,
}

struct ActiveAuxiliaryShape {
    page: usize,
    shape: usize,
    depth: usize,
    kind: ShapeKind,
    frame_auxiliary_phase: u8,
    frame_event_listeners_seen: bool,
    frame_title_seen: bool,
    frame_description_seen: bool,
    image_map_seen: bool,
    contour_seen: bool,
    three_d_child_phase: u8,
}

fn parse_enhanced_geometry(xml: &str, parsed: &mut Parsed) -> Result<()> {
    validate_raw_xml_events(xml, "ODG enhanced-geometry event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut next_page = 0usize;
    let mut current_page = None;
    let mut current_page_depth = None;
    let mut shape_positions = vec![0usize; parsed.pages.len()];
    let mut active_shapes = Vec::new();
    let mut active_geometry = None;
    let mut active_empty_child_depth = None;
    let mut attribute_count = 0usize;
    loop {
        let (resolved_namespace, borrowed_event) =
            reader.read_resolved_event().map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG enhanced-geometry XML: {error}"))
            })?;
        let namespace = classify(&resolved_namespace);
        validate_borrowed_xml_event(&borrowed_event, "ODG enhanced-geometry event")?;
        match borrowed_event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                if active_empty_child_depth.is_some() {
                    return invalid("ODG enhanced-geometry equation or handle has child elements");
                }
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"page" {
                    if current_page.is_some() || next_page >= parsed.pages.len() {
                        return invalid("ODG enhanced-geometry page scope is invalid");
                    }
                    current_page = Some(next_page);
                    current_page_depth = Some(depth);
                    next_page = next_page.saturating_add(1);
                } else if let Some(kind) = shape_kind(namespace, local.as_ref()) {
                    let page = current_page.ok_or_else(|| {
                        Error::InvalidFormat("ODG enhanced shape is outside draw:page".into())
                    })?;
                    let shape = shape_positions[page];
                    shape_positions[page] = shape.saturating_add(1);
                    if shape >= parsed.pages[page].shapes().len() {
                        return invalid("ODG enhanced shape source order is invalid");
                    }
                    active_shapes.push(ActiveAuxiliaryShape {
                        page,
                        shape,
                        depth,
                        kind,
                        frame_auxiliary_phase: 0,
                        frame_event_listeners_seen: false,
                        frame_title_seen: false,
                        frame_description_seen: false,
                        image_map_seen: false,
                        contour_seen: false,
                        three_d_child_phase: 0,
                    });
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"enhanced-geometry"
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Custom && shape.depth.saturating_add(1) == depth
                    })
                {
                    if active_geometry.is_some() {
                        return invalid("ODG custom shape has duplicate enhanced geometry");
                    }
                    let shape = active_shapes.last().ok_or_else(|| {
                        Error::InvalidFormat("ODG enhanced geometry has no owner".into())
                    })?;
                    if parsed.pages[shape.page].shapes()[shape.shape]
                        .enhanced_geometry()
                        .is_some()
                    {
                        return invalid("ODG custom shape has duplicate enhanced geometry");
                    }
                    let attributes = drawing_attributes(&reader, &element, &mut attribute_count)?;
                    validate_enhanced_geometry_attributes(&attributes)?;
                    active_geometry = Some(ActiveEnhancedGeometry {
                        page: shape.page,
                        shape: shape.shape,
                        depth,
                        attributes,
                        children: Vec::new(),
                    });
                } else if let Some(kind) = enhanced_child_kind(namespace, local.as_ref())
                    && active_geometry
                        .as_ref()
                        .is_some_and(|geometry| geometry.depth.saturating_add(1) == depth)
                {
                    let attributes = drawing_attributes(&reader, &element, &mut attribute_count)?;
                    validate_enhanced_child_attributes(&attributes)?;
                    if let Some(geometry) = active_geometry.as_mut() {
                        geometry
                            .children
                            .push(EnhancedGeometryChild::parsed(kind, attributes));
                    }
                    active_empty_child_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                if active_empty_child_depth.is_some() {
                    return invalid("ODG enhanced-geometry equation or handle has child elements");
                }
                let virtual_depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"page" {
                    if next_page >= parsed.pages.len() {
                        return invalid("ODG enhanced-geometry page scope is invalid");
                    }
                    next_page = next_page.saturating_add(1);
                } else if let Some(kind) = shape_kind(namespace, local.as_ref()) {
                    let page = current_page.ok_or_else(|| {
                        Error::InvalidFormat("ODG enhanced shape is outside draw:page".into())
                    })?;
                    let shape = shape_positions[page];
                    shape_positions[page] = shape.saturating_add(1);
                    if shape >= parsed.pages[page].shapes().len() {
                        return invalid("ODG enhanced shape source order is invalid");
                    }
                    let _ = kind;
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"enhanced-geometry"
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Custom
                            && shape.depth.saturating_add(1) == virtual_depth
                    })
                {
                    let shape = active_shapes.last().ok_or_else(|| {
                        Error::InvalidFormat("ODG enhanced geometry has no owner".into())
                    })?;
                    if parsed.pages[shape.page].shapes()[shape.shape]
                        .enhanced_geometry()
                        .is_some()
                    {
                        return invalid("ODG custom shape has duplicate enhanced geometry");
                    }
                    let geometry = EnhancedGeometry::parsed(
                        {
                            let attributes =
                                drawing_attributes(&reader, &element, &mut attribute_count)?;
                            validate_enhanced_geometry_attributes(&attributes)?;
                            attributes
                        },
                        Vec::new(),
                    )?;
                    parsed.pages[shape.page]
                        .shape_mut(shape.shape)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG enhanced shape is missing".into())
                        })?
                        .set_enhanced_geometry(geometry);
                } else if let Some(kind) = enhanced_child_kind(namespace, local.as_ref())
                    && active_geometry
                        .as_ref()
                        .is_some_and(|geometry| geometry.depth.saturating_add(1) == virtual_depth)
                {
                    let attributes = drawing_attributes(&reader, &element, &mut attribute_count)?;
                    validate_enhanced_child_attributes(&attributes)?;
                    if let Some(geometry) = active_geometry.as_mut() {
                        geometry
                            .children
                            .push(EnhancedGeometryChild::parsed(kind, attributes));
                    }
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if active_empty_child_depth == Some(depth) {
                    active_empty_child_depth = None;
                }
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"enhanced-geometry"
                    && active_geometry
                        .as_ref()
                        .is_some_and(|geometry| geometry.depth == depth)
                {
                    let geometry = active_geometry.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG enhanced geometry source is missing".into())
                    })?;
                    let parsed_geometry =
                        EnhancedGeometry::parsed(geometry.attributes, geometry.children)?;
                    parsed.pages[geometry.page]
                        .shape_mut(geometry.shape)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG enhanced shape is missing".into())
                        })?
                        .set_enhanced_geometry(parsed_geometry);
                }
                if shape_kind(namespace, local.as_ref()).is_some()
                    && active_shapes
                        .last()
                        .is_some_and(|shape| shape.depth == depth)
                {
                    active_shapes.pop();
                }
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"page"
                    && current_page_depth == Some(depth)
                {
                    current_page = None;
                    current_page_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG enhanced XML depth underflow".into())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG enhanced geometry"),
            Event::Eof => break,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
            Event::GeneralRef(_) => {
                if active_empty_child_depth.is_some() {
                    return invalid(
                        "ODG enhanced-geometry equation or handle must have empty content",
                    );
                }
            },
            Event::CData(text) => {
                if active_empty_child_depth.is_some()
                    && !text
                        .as_ref()
                        .iter()
                        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return invalid(
                        "ODG enhanced-geometry equation or handle must have empty content",
                    );
                }
            },
            Event::Text(text) => {
                if active_empty_child_depth.is_some()
                    && !text
                        .as_ref()
                        .iter()
                        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return invalid(
                        "ODG enhanced-geometry equation or handle must have empty content",
                    );
                }
            },
        }
    }
    if depth != 0
        || current_page.is_some()
        || !active_shapes.is_empty()
        || active_geometry.is_some()
        || active_empty_child_depth.is_some()
    {
        return invalid("ODG enhanced-geometry XML is incomplete");
    }
    Ok(())
}

struct ActiveAuxiliaryImage {
    depth: usize,
}

struct ActiveAuxiliaryMap {
    page: usize,
    shape: usize,
    depth: usize,
    start: usize,
    areas: Vec<ImageMapArea>,
    bytes: usize,
}

struct ActiveAuxiliaryArea {
    depth: usize,
    start: usize,
    area: ImageMapArea,
}

fn parse_shape_auxiliary_children(xml: &str, parsed: &mut Parsed) -> Result<()> {
    validate_raw_xml_events(xml, "ODG drawing-child event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut next_page = 0usize;
    let mut current_page = None;
    let mut current_page_depth = None;
    let mut shape_positions = vec![0usize; parsed.pages.len()];
    let mut active_shapes: Vec<ActiveAuxiliaryShape> = Vec::new();
    let mut active_image = None;
    let mut active_map = None;
    let mut active_area = None;
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, borrowed_event) =
            reader.read_resolved_event().map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG drawing-child XML: {error}"))
            })?;
        let namespace = classify(&resolved_namespace);
        validate_borrowed_xml_event(&borrowed_event, "ODG drawing-child event")?;
        let end = position(&reader)?;
        match borrowed_event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"page" {
                    if current_page.is_some() || next_page >= parsed.pages.len() {
                        return invalid("ODG drawing-child page scope is invalid");
                    }
                    current_page = Some(next_page);
                    current_page_depth = Some(depth);
                    next_page = next_page.saturating_add(1);
                } else if let Some(kind) = shape_kind(namespace, local.as_ref()) {
                    let page = current_page.ok_or_else(|| {
                        Error::InvalidFormat("ODG drawing shape is outside draw:page".into())
                    })?;
                    update_auxiliary_shape_parent_order(&mut active_shapes, depth, kind)?;
                    let shape = next_auxiliary_shape(&mut shape_positions, parsed, page)?;
                    active_shapes.push(ActiveAuxiliaryShape {
                        page,
                        shape,
                        depth,
                        kind,
                        frame_auxiliary_phase: 0,
                        frame_event_listeners_seen: false,
                        frame_title_seen: false,
                        frame_description_seen: false,
                        image_map_seen: false,
                        contour_seen: false,
                        three_d_child_phase: 0,
                    });
                } else if is_frame_payload(namespace, local.as_ref())
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Frame && shape.depth.saturating_add(1) == depth
                    })
                {
                    let frame = active_shapes.last_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG frame payload has no frame owner".into())
                    })?;
                    if frame.frame_auxiliary_phase != 0 {
                        return invalid("ODG draw:frame payload is out of order");
                    }
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"image"
                    && active_shapes
                        .last()
                        .is_some_and(|shape| shape.depth.saturating_add(1) == depth)
                {
                    if let Some(shape) = active_shapes.last_mut().filter(|shape| {
                        shape.kind == ShapeKind::Frame && shape.depth.saturating_add(1) == depth
                    }) && shape.frame_auxiliary_phase != 0
                    {
                        return invalid("ODG draw:image is out of draw:frame order");
                    }
                    if active_image.is_some() {
                        return invalid("ODG draw:image owner is duplicated");
                    }
                    let shape = active_shapes.last().ok_or_else(|| {
                        Error::InvalidFormat("ODG draw:image has no drawing owner".into())
                    })?;
                    if shape.kind != ShapeKind::Frame {
                        return invalid("ODG draw:image must be owned by draw:frame");
                    }
                    active_image = Some(ActiveAuxiliaryImage { depth });
                } else if namespace == NamespaceKind::Draw && local.as_ref() == b"image-map" {
                    {
                        let shape = active_shapes
                            .last_mut()
                            .filter(|shape| {
                                shape.kind == ShapeKind::Frame
                                    && shape.depth.saturating_add(1) == depth
                            })
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG draw:image-map must be a direct draw:frame child".into(),
                                )
                            })?;
                        if shape.image_map_seen || shape.frame_auxiliary_phase > 2 {
                            return invalid("ODG draw:image-map is duplicated or out of order");
                        }
                        shape.image_map_seen = true;
                        shape.frame_auxiliary_phase = 3;
                    }
                    let (page, shape) =
                        auxiliary_image_owner(&active_shapes, depth).ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODG draw:image-map must be a direct draw:frame child".into(),
                            )
                        })?;
                    if active_map.is_some() {
                        return invalid("ODG draw:image-map owner is duplicated");
                    }
                    active_map = Some(ActiveAuxiliaryMap {
                        page,
                        shape,
                        depth,
                        start,
                        areas: Vec::new(),
                        bytes: 0,
                    });
                } else if area_kind(namespace, local.as_ref()).is_some() {
                    let map = active_map
                        .as_ref()
                        .filter(|map| map.depth.saturating_add(1) == depth)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG image-map area has no map owner".into())
                        })?;
                    if map.areas.len() >= 65_536 {
                        return invalid("ODG image-map area count exceeds the limit");
                    }
                    let area = parse_image_map_area(&reader, &element, start..end, xml)?;
                    active_area = Some(ActiveAuxiliaryArea { depth, start, area });
                } else if namespace == NamespaceKind::Office
                    && local.as_ref() == b"event-listeners"
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Frame && shape.depth.saturating_add(1) == depth
                    })
                {
                    let frame = active_shapes.last_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG frame event-listeners has no owner".into())
                    })?;
                    if frame.frame_event_listeners_seen || frame.frame_auxiliary_phase != 0 {
                        return invalid("ODG draw:frame event-listeners are out of order");
                    }
                    frame.frame_event_listeners_seen = true;
                    frame.frame_auxiliary_phase = 1;
                } else if namespace == NamespaceKind::Svg
                    && matches!(local.as_ref(), b"title" | b"desc")
                    && active_shapes.last().is_some_and(|shape| {
                        shape.depth.saturating_add(1) == depth
                            && matches!(
                                shape.kind,
                                ShapeKind::Frame | ShapeKind::ThreeDimensionalScene
                            )
                    })
                {
                    let shape = active_shapes.last_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG accessibility element has no owner".into())
                    })?;
                    if shape.kind == ShapeKind::Frame {
                        if local.as_ref() == b"title" {
                            if shape.frame_title_seen || shape.frame_auxiliary_phase > 3 {
                                return invalid("ODG draw:frame title is out of order");
                            }
                            shape.frame_title_seen = true;
                            shape.frame_auxiliary_phase = 4;
                        } else {
                            if shape.frame_description_seen || shape.frame_auxiliary_phase > 4 {
                                return invalid("ODG draw:frame description is out of order");
                            }
                            shape.frame_description_seen = true;
                            shape.frame_auxiliary_phase = 5;
                        }
                    } else if shape.three_d_child_phase != 0
                        || (local.as_ref() == b"title" && shape.frame_title_seen)
                        || (local.as_ref() == b"title" && shape.frame_description_seen)
                        || (local.as_ref() == b"desc" && shape.frame_description_seen)
                    {
                        return invalid("ODG dr3d:scene accessibility is out of order");
                    } else if local.as_ref() == b"title" {
                        shape.frame_title_seen = true;
                    } else {
                        shape.frame_description_seen = true;
                    }
                }
                // RNG empty content also permits paired XML tags. Consume the
                // closing tag without allowing child elements or text content.
                if namespace == NamespaceKind::Draw
                    && matches!(local.as_ref(), b"contour-polygon" | b"contour-path")
                {
                    {
                        let shape = active_shapes
                            .last_mut()
                            .filter(|shape| {
                                shape.kind == ShapeKind::Frame
                                    && shape.depth.saturating_add(1) == depth
                            })
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG contour must be a direct draw:frame child".into(),
                                )
                            })?;
                        if shape.contour_seen {
                            return invalid("ODG draw:frame has duplicate contours");
                        }
                        shape.contour_seen = true;
                        if shape.frame_auxiliary_phase > 5 {
                            return invalid("ODG draw:frame contour is out of order");
                        }
                        shape.frame_auxiliary_phase = 6;
                    }
                    let (page, shape_index) = auxiliary_image_owner(&active_shapes, depth)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG contour has no image owner".into())
                        })?;
                    let end = empty_auxiliary_end(&mut reader)?;
                    let contour = parse_contour(&reader, &element, start..end, xml)?;
                    let owner = parsed.pages[page].shape_mut(shape_index).ok_or_else(|| {
                        Error::InvalidFormat("ODG contour shape is missing".into())
                    })?;
                    if owner.contours().len() >= 65_536 {
                        return invalid("ODG contour count exceeds the limit");
                    }
                    owner.push_contour(contour);
                    depth -= 1;
                }
                if namespace == NamespaceKind::Draw && local.as_ref() == b"glue-point" {
                    if let Some(shape) = active_shapes.last_mut().filter(|shape| {
                        shape.kind == ShapeKind::Frame && shape.depth.saturating_add(1) == depth
                    }) {
                        if shape.frame_auxiliary_phase > 2 {
                            return invalid("ODG draw:glue-point is out of draw:frame order");
                        }
                        shape.frame_auxiliary_phase = 2;
                    }
                    let (page_index, shape_index, shape_kind) = {
                        let shape = active_shapes
                            .last()
                            .filter(|shape| shape.depth.saturating_add(1) == depth)
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG glue point has no direct shape owner".into(),
                                )
                            })?;
                        (shape.page, shape.shape, shape.kind)
                    };
                    if !allows_glue_point(shape_kind) {
                        return invalid("ODG glue-point owner does not permit glue points");
                    }
                    if shape_kind == ShapeKind::ThreeDimensionalScene {
                        let scene = active_shapes.last_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODG scene glue-point owner disappeared".into())
                        })?;
                        if scene.three_d_child_phase > 3 {
                            return invalid("ODG dr3d:scene glue-point is out of order");
                        }
                        scene.three_d_child_phase = 3;
                    }
                    let end = empty_auxiliary_end(&mut reader)?;
                    let glue_point = parse_glue_point(&reader, &element, start..end, xml)?;
                    let owner =
                        parsed.pages[page_index]
                            .shape_mut(shape_index)
                            .ok_or_else(|| {
                                Error::InvalidFormat("ODG glue-point shape is missing".into())
                            })?;
                    if owner.glue_points().len() >= 65_536 {
                        return invalid("ODG glue-point count exceeds the limit");
                    }
                    owner.push_glue_point(glue_point);
                    depth -= 1;
                }
            },
            Event::Empty(element) => {
                let virtual_depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"page" {
                    if next_page >= parsed.pages.len() {
                        return invalid("ODG drawing-child page scope is invalid");
                    }
                    next_page = next_page.saturating_add(1);
                } else if let Some(kind) = shape_kind(namespace, local.as_ref()) {
                    let page = current_page.ok_or_else(|| {
                        Error::InvalidFormat("ODG drawing shape is outside draw:page".into())
                    })?;
                    update_auxiliary_shape_parent_order(&mut active_shapes, virtual_depth, kind)?;
                    let shape = next_auxiliary_shape(&mut shape_positions, parsed, page)?;
                    if active_image.is_some() || active_map.is_some() || active_area.is_some() {
                        return invalid("ODG nested drawing owner is misplaced");
                    }
                    let _ = kind;
                    let _ = shape;
                } else if is_frame_payload(namespace, local.as_ref())
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Frame
                            && shape.depth.saturating_add(1) == virtual_depth
                    })
                {
                    let frame = active_shapes.last().ok_or_else(|| {
                        Error::InvalidFormat("ODG frame payload has no frame owner".into())
                    })?;
                    if frame.frame_auxiliary_phase != 0 {
                        return invalid("ODG draw:frame payload is out of order");
                    }
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"image"
                    && active_shapes
                        .last()
                        .is_some_and(|shape| shape.depth.saturating_add(1) == virtual_depth)
                {
                    if let Some(shape) = active_shapes.last().filter(|shape| {
                        shape.kind == ShapeKind::Frame
                            && shape.depth.saturating_add(1) == virtual_depth
                    }) && shape.frame_auxiliary_phase != 0
                    {
                        return invalid("ODG draw:image is out of draw:frame order");
                    }
                    if active_image.is_some() {
                        return invalid("ODG draw:image owner is duplicated");
                    }
                    let shape = active_shapes.last().ok_or_else(|| {
                        Error::InvalidFormat("ODG draw:image has no drawing owner".into())
                    })?;
                    if shape.kind != ShapeKind::Frame {
                        return invalid("ODG draw:image must be owned by draw:frame");
                    }
                    let _ = shape;
                } else if namespace == NamespaceKind::Draw && local.as_ref() == b"image-map" {
                    {
                        let shape = active_shapes
                            .last_mut()
                            .filter(|shape| {
                                shape.kind == ShapeKind::Frame
                                    && shape.depth.saturating_add(1) == virtual_depth
                            })
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG draw:image-map must be a direct draw:frame child".into(),
                                )
                            })?;
                        if shape.image_map_seen || shape.frame_auxiliary_phase > 2 {
                            return invalid("ODG draw:image-map is duplicated or out of order");
                        }
                        shape.image_map_seen = true;
                        shape.frame_auxiliary_phase = 3;
                    }
                    let (page, shape) = auxiliary_image_owner(&active_shapes, virtual_depth)
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODG draw:image-map must be a direct draw:frame child".into(),
                            )
                        })?;
                    let map = ImageMap::parsed(
                        Vec::new(),
                        source_fragment(xml, start..end, "ODG image-map")?,
                    )?;
                    attach_image_map(parsed, page, shape, map)?;
                } else if let Some(kind) = area_kind(namespace, local.as_ref()) {
                    let map = active_map
                        .as_mut()
                        .filter(|map| map.depth.saturating_add(1) == virtual_depth)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG image-map area has no map owner".into())
                        })?;
                    if map.areas.len() >= 65_536 {
                        return invalid("ODG image-map area count exceeds the limit");
                    }
                    charge_source_span(xml, start..end, &mut map.bytes, "ODG image-map area")?;
                    let area = parse_image_map_area(&reader, &element, start..end, xml)?;
                    map.areas.push(area);
                    let _ = kind;
                } else if namespace == NamespaceKind::Office
                    && local.as_ref() == b"event-listeners"
                    && active_shapes.last().is_some_and(|shape| {
                        shape.kind == ShapeKind::Frame
                            && shape.depth.saturating_add(1) == virtual_depth
                    })
                {
                    let frame = active_shapes.last_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG frame event-listeners has no owner".into())
                    })?;
                    if frame.frame_event_listeners_seen || frame.frame_auxiliary_phase != 0 {
                        return invalid("ODG draw:frame event-listeners are out of order");
                    }
                    frame.frame_event_listeners_seen = true;
                    frame.frame_auxiliary_phase = 1;
                } else if namespace == NamespaceKind::Svg
                    && matches!(local.as_ref(), b"title" | b"desc")
                    && active_shapes.last().is_some_and(|shape| {
                        shape.depth.saturating_add(1) == virtual_depth
                            && matches!(
                                shape.kind,
                                ShapeKind::Frame | ShapeKind::ThreeDimensionalScene
                            )
                    })
                {
                    let shape = active_shapes.last_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG accessibility element has no owner".into())
                    })?;
                    if shape.kind == ShapeKind::Frame {
                        if local.as_ref() == b"title" {
                            if shape.frame_title_seen || shape.frame_auxiliary_phase > 3 {
                                return invalid("ODG draw:frame title is out of order");
                            }
                            shape.frame_title_seen = true;
                            shape.frame_auxiliary_phase = 4;
                        } else {
                            if shape.frame_description_seen || shape.frame_auxiliary_phase > 4 {
                                return invalid("ODG draw:frame description is out of order");
                            }
                            shape.frame_description_seen = true;
                            shape.frame_auxiliary_phase = 5;
                        }
                    } else if shape.three_d_child_phase != 0
                        || (local.as_ref() == b"title" && shape.frame_title_seen)
                        || (local.as_ref() == b"title" && shape.frame_description_seen)
                        || (local.as_ref() == b"desc" && shape.frame_description_seen)
                    {
                        return invalid("ODG dr3d:scene accessibility is out of order");
                    } else if local.as_ref() == b"title" {
                        shape.frame_title_seen = true;
                    } else {
                        shape.frame_description_seen = true;
                    }
                } else if namespace == NamespaceKind::Draw
                    && matches!(local.as_ref(), b"contour-polygon" | b"contour-path")
                {
                    {
                        let shape = active_shapes
                            .last_mut()
                            .filter(|shape| {
                                shape.kind == ShapeKind::Frame
                                    && shape.depth.saturating_add(1) == virtual_depth
                            })
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG contour must be a direct draw:frame child".into(),
                                )
                            })?;
                        if shape.contour_seen {
                            return invalid("ODG draw:frame has duplicate contours");
                        }
                        shape.contour_seen = true;
                        if shape.frame_auxiliary_phase > 5 {
                            return invalid("ODG draw:frame contour is out of order");
                        }
                        shape.frame_auxiliary_phase = 6;
                    }
                    let (page, shape_index) = auxiliary_image_owner(&active_shapes, virtual_depth)
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODG contour must be a direct draw:frame child".into(),
                            )
                        })?;
                    let contour = parse_contour(&reader, &element, start..end, xml)?;
                    let shape = parsed.pages[page].shape_mut(shape_index).ok_or_else(|| {
                        Error::InvalidFormat("ODG contour shape is missing".into())
                    })?;
                    if shape.contours().len() >= 65_536 {
                        return invalid("ODG contour count exceeds the limit");
                    }
                    shape.push_contour(contour);
                } else if namespace == NamespaceKind::Draw && local.as_ref() == b"glue-point" {
                    if let Some(shape) = active_shapes.last_mut().filter(|shape| {
                        shape.kind == ShapeKind::Frame
                            && shape.depth.saturating_add(1) == virtual_depth
                    }) {
                        if shape.frame_auxiliary_phase > 2 {
                            return invalid("ODG draw:glue-point is out of draw:frame order");
                        }
                        shape.frame_auxiliary_phase = 2;
                    }
                    let (page_index, shape_index, shape_kind) = {
                        let shape = active_shapes
                            .last()
                            .filter(|shape| shape.depth.saturating_add(1) == virtual_depth)
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODG glue point has no direct shape owner".into(),
                                )
                            })?;
                        (shape.page, shape.shape, shape.kind)
                    };
                    if !allows_glue_point(shape_kind) {
                        return invalid("ODG glue-point owner does not permit glue points");
                    }
                    if shape_kind == ShapeKind::ThreeDimensionalScene {
                        let scene = active_shapes.last_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODG scene glue-point owner disappeared".into())
                        })?;
                        if scene.three_d_child_phase > 3 {
                            return invalid("ODG dr3d:scene glue-point is out of order");
                        }
                        scene.three_d_child_phase = 3;
                    }
                    let glue_point = parse_glue_point(&reader, &element, start..end, xml)?;
                    let owner =
                        parsed.pages[page_index]
                            .shape_mut(shape_index)
                            .ok_or_else(|| {
                                Error::InvalidFormat("ODG glue-point shape is missing".into())
                            })?;
                    if owner.glue_points().len() >= 65_536 {
                        return invalid("ODG glue-point count exceeds the limit");
                    }
                    owner.push_glue_point(glue_point);
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if active_area.as_ref().is_some_and(|area| {
                    area.depth == depth && area_kind(namespace, local.as_ref()).is_some()
                }) {
                    let area = active_area.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG image-map area source is missing".into())
                    })?;
                    let map = active_map.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("ODG image-map area has no map owner".into())
                    })?;
                    charge_source_span(xml, area.start..end, &mut map.bytes, "ODG image-map area")?;
                    let source_xml = source_fragment(xml, area.start..end, "ODG image-map area")?;
                    let mut value = area.area;
                    value = ImageMapArea::parsed(
                        value.shape().clone(),
                        value.href().map(str::to_owned),
                        value.target_frame_name().map(str::to_owned),
                        value.show().map(str::to_owned),
                        value.no_href(),
                        value.name().map(str::to_owned),
                        source_xml,
                    )?;
                    map.areas.push(value);
                }
                if active_map
                    .as_ref()
                    .is_some_and(|map| map.depth == depth && local.as_ref() == b"image-map")
                {
                    let mut map = active_map.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG image-map source is missing".into())
                    })?;
                    charge_source_span(xml, map.start..end, &mut map.bytes, "ODG image-map")?;
                    let image_map = ImageMap::parsed(
                        map.areas,
                        source_fragment(xml, map.start..end, "ODG image-map")?,
                    )?;
                    attach_image_map(parsed, map.page, map.shape, image_map)?;
                }
                if active_image
                    .as_ref()
                    .is_some_and(|image| image.depth == depth && local.as_ref() == b"image")
                {
                    active_image = None;
                }
                if shape_kind(namespace, local.as_ref()).is_some()
                    && active_shapes
                        .last()
                        .is_some_and(|shape| shape.depth == depth)
                {
                    active_shapes.pop();
                }
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"page"
                    && current_page_depth == Some(depth)
                {
                    current_page = None;
                    current_page_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG drawing-child XML depth underflow".into())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG drawing-child XML"),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0
        || current_page.is_some()
        || !active_shapes.is_empty()
        || active_image.is_some()
        || active_map.is_some()
        || active_area.is_some()
    {
        return invalid("ODG drawing-child XML is incomplete");
    }
    Ok(())
}

fn next_auxiliary_shape(
    shape_positions: &mut [usize],
    parsed: &Parsed,
    page: usize,
) -> Result<usize> {
    let shape = *shape_positions
        .get(page)
        .ok_or_else(|| Error::InvalidFormat("ODG drawing-child page is missing".into()))?;
    let next = shape
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG drawing-child shape count overflow".into()))?;
    shape_positions[page] = next;
    if shape >= parsed.pages[page].shapes().len() {
        return invalid("ODG drawing-child shape source order is invalid");
    }
    Ok(shape)
}

fn empty_auxiliary_end(reader: &mut NsReader<&[u8]>) -> Result<usize> {
    loop {
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG empty drawing child: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG empty drawing child event")?;
        match event {
            Event::End(_) => return position(reader),
            Event::Comment(_) | Event::PI(_) => {},
            Event::Text(text)
                if text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n')) => {},
            _ => return invalid("ODG contour or glue-point must have empty content"),
        }
    }
}

fn auxiliary_image_owner(
    active_shapes: &[ActiveAuxiliaryShape],
    depth: usize,
) -> Option<(usize, usize)> {
    active_shapes
        .last()
        .filter(|shape| shape.depth.saturating_add(1) == depth && shape.kind == ShapeKind::Frame)
        .map(|shape| (shape.page, shape.shape))
}

fn is_frame_payload(namespace: NamespaceKind, local: &[u8]) -> bool {
    matches!(
        (namespace, local),
        (NamespaceKind::Draw, b"text-box")
            | (NamespaceKind::Draw, b"object")
            | (NamespaceKind::Draw, b"object-ole")
            | (NamespaceKind::Draw, b"applet")
            | (NamespaceKind::Draw, b"floating-frame")
            | (NamespaceKind::Draw, b"plugin")
            | (NamespaceKind::Table, b"table")
    )
}

fn update_auxiliary_shape_parent_order(
    active_shapes: &mut [ActiveAuxiliaryShape],
    depth: usize,
    kind: ShapeKind,
) -> Result<()> {
    if active_shapes
        .iter()
        .rev()
        .any(|shape| shape.kind == ShapeKind::Frame && depth > shape.depth.saturating_add(1))
    {
        return invalid("ODG draw:frame payload cannot contain drawing shapes");
    }
    let Some(parent) = active_shapes
        .iter_mut()
        .rev()
        .find(|shape| shape.depth.saturating_add(1) == depth)
    else {
        return Ok(());
    };
    if parent.kind == ShapeKind::Frame {
        return invalid("ODG draw:frame contains a shape outside its payload grammar");
    }
    if parent.kind != ShapeKind::ThreeDimensionalScene {
        return Ok(());
    }
    // ODF 1.4 `shapes3d` includes `dr3d-scene`, so a nested scene is an
    // ordinary 3D child of its parent scene.
    if !kind.is_three_dimensional() {
        return invalid("ODG dr3d:scene contains a non-3D drawing object");
    }
    if kind == ShapeKind::ThreeDimensionalLight {
        if parent.three_d_child_phase >= 2 {
            return invalid("ODG dr3d:scene light is out of order");
        }
        parent.three_d_child_phase = 1;
    } else {
        if parent.three_d_child_phase >= 3 {
            return invalid("ODG dr3d:scene shape is out of order");
        }
        parent.three_d_child_phase = 2;
    }
    Ok(())
}

const fn allows_glue_point(kind: ShapeKind) -> bool {
    !kind.is_three_dimensional() || matches!(kind, ShapeKind::ThreeDimensionalScene)
}

fn attach_image_map(parsed: &mut Parsed, page: usize, shape: usize, map: ImageMap) -> Result<()> {
    let owner = parsed.pages[page]
        .shape_mut(shape)
        .ok_or_else(|| Error::InvalidFormat("ODG image-map shape is missing".into()))?;
    if !owner.set_image_map(map) {
        return invalid("ODG drawing shape has duplicate image maps");
    }
    Ok(())
}

fn area_kind(namespace: NamespaceKind, local: &[u8]) -> Option<u8> {
    if namespace != NamespaceKind::Draw {
        return None;
    }
    match local {
        b"area-rectangle" => Some(0),
        b"area-circle" => Some(1),
        b"area-polygon" => Some(2),
        _ => None,
    }
}

fn parse_image_map_area(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    span: Range<usize>,
    xml: &str,
) -> Result<ImageMapArea> {
    let local = element.local_name();
    let shape = match local.as_ref() {
        b"area-rectangle" => ImageMapAreaShape::Rectangle {
            x: required_attribute(reader, element, SVG, b"x", "image-map area")?,
            y: required_attribute(reader, element, SVG, b"y", "image-map area")?,
            width: required_attribute(reader, element, SVG, b"width", "image-map area")?,
            height: required_attribute(reader, element, SVG, b"height", "image-map area")?,
        },
        b"area-circle" => ImageMapAreaShape::Circle {
            cx: required_attribute(reader, element, SVG, b"cx", "image-map area")?,
            cy: required_attribute(reader, element, SVG, b"cy", "image-map area")?,
            r: required_attribute(reader, element, SVG, b"r", "image-map area")?,
        },
        b"area-polygon" => ImageMapAreaShape::Polygon {
            x: required_attribute(reader, element, SVG, b"x", "image-map area")?,
            y: required_attribute(reader, element, SVG, b"y", "image-map area")?,
            width: required_attribute(reader, element, SVG, b"width", "image-map area")?,
            height: required_attribute(reader, element, SVG, b"height", "image-map area")?,
            view_box: required_attribute(reader, element, SVG, b"viewBox", "image-map area")?,
            points: required_attribute(reader, element, DRAW, b"points", "image-map area")?,
        },
        _ => return invalid("ODG image-map area kind is unsupported"),
    };
    if let ImageMapAreaShape::Polygon {
        view_box, points, ..
    } = &shape
        && (!is_integer_list(view_box, 4) || !is_points(points))
    {
        return invalid("ODG image-map polygon geometry is invalid");
    }
    let geometry_valid = match &shape {
        ImageMapAreaShape::Rectangle {
            x,
            y,
            width,
            height,
        }
        | ImageMapAreaShape::Polygon {
            x,
            y,
            width,
            height,
            ..
        } => [x, y, width, height]
            .iter()
            .all(|value| is_odf_length(value)),
        ImageMapAreaShape::Circle { cx, cy, r } => {
            [cx, cy, r].iter().all(|value| is_odf_length(value))
        },
    };
    if !geometry_valid {
        return invalid("ODG image-map geometry is not an ODF length");
    }
    let link_type = attribute(reader, element, XLINK, b"type")?;
    let href = attribute(reader, element, XLINK, b"href")?;
    if href.is_some() {
        if link_type.as_deref() != Some("simple") {
            return invalid("ODG image-map xlink:type must be 'simple' when xlink:href is present");
        }
    } else if let Some(link_type) = link_type {
        if link_type != "simple" {
            return invalid("ODG image-map xlink:type must be 'simple'");
        }
        return invalid("ODG image-map xlink:type requires xlink:href");
    }
    let show = attribute(reader, element, XLINK, b"show")?;
    if show
        .as_deref()
        .is_some_and(|value| !matches!(value, "new" | "replace"))
    {
        return invalid("ODG image-map xlink:show is invalid");
    }
    let no_href = attribute(reader, element, DRAW, b"nohref")?;
    let no_href = match no_href.as_deref() {
        None => false,
        Some("nohref") => true,
        Some(_) => return invalid("ODG image-map draw:nohref is invalid"),
    };
    ImageMapArea::parsed(
        shape,
        href,
        attribute(reader, element, OFFICE, b"target-frame-name")?,
        show,
        no_href,
        attribute(reader, element, OFFICE, b"name")?,
        source_fragment(xml, span, "ODG image-map area")?,
    )
}

fn parse_contour(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    span: Range<usize>,
    xml: &str,
) -> Result<Contour> {
    let local = element.local_name();
    let kind = match local.as_ref() {
        b"contour-polygon" => ContourKind::Polygon,
        b"contour-path" => ContourKind::Path,
        _ => return invalid("ODG contour kind is unsupported"),
    };
    let recreate_on_edit =
        required_attribute(reader, element, DRAW, b"recreate-on-edit", "contour")?;
    if !matches!(recreate_on_edit.as_str(), "true" | "false") {
        return invalid("ODG contour draw:recreate-on-edit is invalid");
    }
    match kind {
        ContourKind::Polygon => {
            let view_box = required_attribute(reader, element, SVG, b"viewBox", "contour polygon")?;
            let points = required_attribute(reader, element, DRAW, b"points", "contour polygon")?;
            if !is_integer_list(&view_box, 4) || !is_points(&points) {
                return invalid("ODG contour polygon geometry is invalid");
            }
        },
        ContourKind::Path => {
            let view_box = required_attribute(reader, element, SVG, b"viewBox", "contour path")?;
            required_attribute(reader, element, SVG, b"d", "contour path")?;
            if !is_integer_list(&view_box, 4) {
                return invalid("ODG contour path viewBox is not four integers");
            }
        },
    }
    for (name, value) in [
        (
            b"width".as_slice(),
            attribute(reader, element, SVG, b"width")?,
        ),
        (
            b"height".as_slice(),
            attribute(reader, element, SVG, b"height")?,
        ),
    ] {
        if let Some(value) = value
            && !is_odf_length(&value)
        {
            return invalid(format!("ODG contour {name:?} is not an ODF length"));
        }
    }
    let mut attribute_count = 0usize;
    Contour::parsed(
        kind,
        drawing_attributes(reader, element, &mut attribute_count)?,
        source_fragment(xml, span, "ODG contour")?,
    )
}

fn parse_glue_point(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    span: Range<usize>,
    xml: &str,
) -> Result<GluePoint> {
    let align = attribute(reader, element, DRAW, b"align")?;
    if align.as_deref().is_some_and(|value| {
        !matches!(
            value,
            "top-left"
                | "top"
                | "top-right"
                | "left"
                | "center"
                | "right"
                | "bottom-left"
                | "bottom-right"
        )
    }) {
        return invalid("ODG glue-point draw:align is invalid");
    }
    let escape_direction =
        required_attribute(reader, element, DRAW, b"escape-direction", "glue-point")?;
    if !matches!(
        escape_direction.as_str(),
        "auto" | "left" | "right" | "up" | "down" | "horizontal" | "vertical"
    ) {
        return invalid("ODG glue-point escape direction is invalid");
    }
    let id = required_attribute(reader, element, DRAW, b"id", "glue-point")?;
    let integer = id.strip_prefix('+').unwrap_or(&id);
    if integer.is_empty() || !integer.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid("ODG glue-point draw:id is not a non-negative integer");
    }
    let x = required_attribute(reader, element, SVG, b"x", "glue-point")?;
    let y = required_attribute(reader, element, SVG, b"y", "glue-point")?;
    if !is_odf_distance_or_percent(&x) || !is_odf_distance_or_percent(&y) {
        return invalid("ODG glue-point coordinates are not distances or percentages");
    }
    GluePoint::parsed(
        id,
        x,
        y,
        align,
        escape_direction,
        source_fragment(xml, span, "ODG glue-point")?,
    )
}

fn charge_source_span(xml: &str, span: Range<usize>, total: &mut usize, owner: &str) -> Result<()> {
    let source = xml.get(span).ok_or_else(|| {
        Error::InvalidFormat(format!("{owner} source span is outside the source XML"))
    })?;
    let charged = total
        .checked_add(source.len())
        .ok_or_else(|| Error::InvalidFormat(format!("{owner} source size overflow")))?;
    if charged > 8 * 1024 * 1024 {
        return Err(Error::InvalidFormat(format!(
            "{owner} aggregate exceeds the byte limit"
        )));
    }
    *total = charged;
    Ok(())
}

fn source_fragment(xml: &str, span: Range<usize>, owner: &str) -> Result<String> {
    let source = xml.get(span).ok_or_else(|| {
        Error::InvalidFormat(format!("{owner} source span is outside the source XML"))
    })?;
    if source.len() > 8 * 1024 * 1024 {
        return Err(Error::InvalidFormat(format!(
            "{owner} source exceeds the byte limit"
        )));
    }
    let mut fragment = String::new();
    fragment
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG source fragment",
            source: allocation_error,
        })?;
    fragment.push_str(source);
    Ok(fragment)
}

fn clone_xml_span(xml: &str, span: Range<usize>, resource: &'static str) -> Result<String> {
    let source = xml
        .get(span)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} source span is invalid")))?;
    clone_bounded_string(source, MAX_OUTPUT_BYTES, resource)
}

fn enhanced_child_kind(
    namespace: NamespaceKind,
    local: &[u8],
) -> Option<EnhancedGeometryChildKind> {
    (namespace == NamespaceKind::Draw).then_some(match local {
        b"equation" => EnhancedGeometryChildKind::Equation,
        b"handle" => EnhancedGeometryChildKind::Handle,
        _ => return None,
    })
}

fn drawing_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    count: &mut usize,
) -> Result<Vec<DrawingAttribute>> {
    let mut attributes = Vec::new();
    for raw in element.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG enhanced-geometry attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match classify(&namespace) {
            NamespaceKind::Draw => DrawingAttributeNamespace::Drawing,
            NamespaceKind::Svg => DrawingAttributeNamespace::Svg,
            NamespaceKind::Dr3d => DrawingAttributeNamespace::Dr3d,
            NamespaceKind::Other
            | NamespaceKind::Office
            | NamespaceKind::Text
            | NamespaceKind::Table
            | NamespaceKind::Form
            | NamespaceKind::Style
            | NamespaceKind::Presentation => continue,
        };
        let local = std::str::from_utf8(local.as_ref()).map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG attribute name: {error}"))
        })?;
        let local = clone_bounded_string(local, MAX_TEXT_BYTES, "ODG attribute name")?;
        validate_raw_xml_value(attribute.value.as_ref(), "ODG enhanced-geometry attribute")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG attribute value: {error}"))
            })?;
        validate_decoded_xml_value(value.as_ref(), "ODG enhanced-geometry attribute")?;
        let value = own_bounded_cow(value, MAX_TEXT_BYTES, "ODG enhanced-geometry attribute")?;
        let attribute = DrawingAttribute::parsed(namespace, local, value)?;
        if attributes.iter().any(|existing: &DrawingAttribute| {
            existing.namespace() == attribute.namespace()
                && existing.local_name() == attribute.local_name()
        }) {
            return invalid("ODG enhanced-geometry attribute is duplicated");
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("ODG enhanced-geometry attribute count overflow".into())
        })?;
        if *count > MAX_SHAPES {
            return invalid("ODG enhanced-geometry attribute count exceeds the limit");
        }
        attributes.push(attribute);
    }
    Ok(attributes)
}

fn validate_enhanced_geometry_attributes(attributes: &[DrawingAttribute]) -> Result<()> {
    for attribute in attributes {
        let value = attribute.value();
        match (attribute.namespace(), attribute.local_name()) {
            (DrawingAttributeNamespace::Svg, "viewBox") if !is_integer_list(value, 4) => {
                return invalid("ODG enhanced-geometry svg:viewBox is not four integers");
            },
            (DrawingAttributeNamespace::Drawing, name)
                if matches!(
                    name,
                    "mirror-vertical"
                        | "mirror-horizontal"
                        | "extrusion-allowed"
                        | "text-path-allowed"
                        | "concentric-gradient-fill-allowed"
                        | "extrusion"
                        | "extrusion-light-face"
                        | "extrusion-first-light-harsh"
                        | "extrusion-second-light-harsh"
                        | "extrusion-metal"
                        | "extrusion-color"
                        | "text-path"
                        | "text-path-same-letter-heights"
                ) && !matches!(value, "true" | "false") =>
            {
                return invalid("ODG enhanced-geometry Boolean attribute is invalid");
            },
            (DrawingAttributeNamespace::Drawing, name)
                if matches!(
                    name,
                    "extrusion-brightness"
                        | "extrusion-diffusion"
                        | "extrusion-first-light-level"
                        | "extrusion-second-light-level"
                        | "extrusion-shininess"
                ) && !is_odf_percent_in_range(value) =>
            {
                return invalid("ODG enhanced-geometry percentage attribute is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-specularity")
                if !is_odf_nonnegative_percent(value) =>
            {
                return invalid("ODG enhanced-geometry specularity is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-depth")
                if !is_length_double_list(value) =>
            {
                return invalid("ODG enhanced-geometry extrusion depth is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-first-light-direction")
            | (DrawingAttributeNamespace::Drawing, "extrusion-second-light-direction")
                if !is_vector3d(value) =>
            {
                return invalid("ODG enhanced-geometry light direction is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-rotation-center")
                if !is_vector3d(value) =>
            {
                return invalid("ODG enhanced-geometry rotation center is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-viewpoint") if !is_point3d(value) => {
                return invalid("ODG enhanced-geometry viewpoint is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-number-of-line-segments")
                if !is_odf_integer(value) =>
            {
                return invalid("ODG enhanced-geometry line-segment count is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-metal-type")
                if !is_namespaced_token(value) =>
            {
                return invalid("ODG enhanced-geometry metal type is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-rotation-angle")
                if !is_token_list(value, 2) =>
            {
                return invalid("ODG enhanced-geometry rotation angle is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-skew")
                if !is_double_token_pair(value) =>
            {
                return invalid("ODG enhanced-geometry skew is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "extrusion-origin")
                if !is_extrusion_origin(value) =>
            {
                return invalid("ODG enhanced-geometry origin is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "glue-point-type")
                if !matches!(value, "none" | "segments" | "rectangle") =>
            {
                return invalid("ODG enhanced-geometry glue-point type is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "text-path-mode")
                if !matches!(value, "normal" | "path" | "shape") =>
            {
                return invalid("ODG enhanced-geometry text-path mode is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "text-path-scale")
                if !matches!(value, "path" | "shape") =>
            {
                return invalid("ODG enhanced-geometry text-path scale is invalid");
            },
            (DrawingAttributeNamespace::Drawing, "path-stretchpoint-x")
            | (DrawingAttributeNamespace::Drawing, "path-stretchpoint-y")
                if !is_odf_double(value) =>
            {
                return invalid("ODG enhanced-geometry stretch point is invalid");
            },
            (DrawingAttributeNamespace::Dr3d, "projection")
                if !matches!(value, "parallel" | "perspective") =>
            {
                return invalid("ODG enhanced-geometry projection is invalid");
            },
            (DrawingAttributeNamespace::Dr3d, "shade-mode")
                if !matches!(value, "flat" | "phong" | "gouraud" | "draft") =>
            {
                return invalid("ODG enhanced-geometry shade mode is invalid");
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_enhanced_child_attributes(attributes: &[DrawingAttribute]) -> Result<()> {
    for attribute in attributes {
        if attribute.namespace() == DrawingAttributeNamespace::Drawing
            && matches!(
                attribute.local_name(),
                "handle-mirror-horizontal" | "handle-mirror-vertical" | "handle-switched"
            )
            && !matches!(attribute.value(), "true" | "false")
        {
            return invalid("ODG enhanced-geometry handle Boolean attribute is invalid");
        }
    }
    Ok(())
}

fn is_odf_double(value: &str) -> bool {
    matches!(value, "INF" | "-INF" | "NaN") || value.parse::<f64>().is_ok()
}

fn is_odf_percent_in_range(value: &str) -> bool {
    is_odf_percent(value)
}

fn is_odf_nonnegative_percent(value: &str) -> bool {
    is_odf_percent(value)
        && value
            .strip_suffix('%')
            .and_then(|number| number.parse::<f64>().ok())
            .is_some_and(|number| number.is_finite() && number >= 0.0)
}

fn is_namespaced_token(value: &str) -> bool {
    let Some((prefix, local)) = value.split_once(':') else {
        return false;
    };
    !prefix.is_empty()
        && !local.is_empty()
        && prefix.chars().all(is_ncname_char)
        && prefix.chars().next().is_some_and(is_ncname_start)
        && local.chars().all(is_ncname_char)
        && local.chars().next().is_some_and(is_ncname_start)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn is_token_list(value: &str, expected: usize) -> bool {
    let values = value.split_ascii_whitespace().collect::<Vec<_>>();
    values.len() == expected && values.iter().all(|value| !value.is_empty())
}

fn is_double_token_pair(value: &str) -> bool {
    let mut values = value.split_ascii_whitespace();
    values.next().is_some_and(is_odf_double)
        && values.next().is_some_and(|_| true)
        && values.next().is_none()
}

fn is_extrusion_origin(value: &str) -> bool {
    let values = value.split_ascii_whitespace().collect::<Vec<_>>();
    values.len() == 2
        && values.iter().all(|value| {
            value
                .parse::<f64>()
                .is_ok_and(|number| number.is_finite() && (-0.5..=0.5).contains(&number))
        })
}

fn is_length_double_list(value: &str) -> bool {
    let mut values = value.split_ascii_whitespace();
    values.next().is_some_and(is_odf_length)
        && values.next().is_some_and(is_odf_double)
        && values.next().is_none()
}

fn is_point3d(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix('(')
        .and_then(|value| value.strip_suffix(')'))
    else {
        return false;
    };
    let mut values = inner.split_ascii_whitespace();
    (0..3).all(|_| {
        values.next().is_some_and(|value| {
            ["cm", "mm", "in", "pt", "pc"]
                .iter()
                .any(|unit| value.strip_suffix(unit).is_some_and(is_odf_decimal))
        })
    }) && values.next().is_none()
}

fn parse_declared_layers(xml: &str) -> Result<Vec<Layer>> {
    validate_raw_xml_events(xml, "ODG styles.xml event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut layer_sets = Vec::<usize>::new();
    let mut layers = Vec::new();
    loop {
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG styles.xml: {error}")))?;
        validate_borrowed_xml_event(&event, "ODG styles.xml event")?;
        let namespace = classify(&resolved_namespace);
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"layer-set" {
                    layer_sets.push(depth);
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets.last().is_some_and(|set| *set + 1 == depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::Empty(element) => {
                let virtual_depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets
                        .last()
                        .is_some_and(|set| *set + 1 == virtual_depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Draw
                    && element.local_name().as_ref() == b"layer-set"
                    && layer_sets.last() == Some(&depth)
                {
                    layer_sets.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG styles XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG styles.xml"),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || !layer_sets.is_empty() {
        return invalid("ODG styles.xml has incomplete layer declarations");
    }
    Ok(layers)
}

fn scan_resources(package: &FamilyPackage) -> Result<Vec<Resource>> {
    let archive = package.package().package()?;
    let images = media::scan_package(package.content_xml(), package.styles_xml(), &archive)?;
    let mut resources = Vec::new();
    for (occurrence, image) in images.into_iter().enumerate() {
        match image.source {
            media::Source::PackagePart {
                href,
                path,
                manifest_media_type,
            } => resources.push(Resource::new(
                occurrence,
                href,
                path,
                manifest_media_type,
                true,
            )),
            media::Source::MissingPackagePart {
                href,
                resolved_path,
            } => resources.push(Resource::new(occurrence, href, resolved_path, None, false)),
            media::Source::Inline { .. }
            | media::Source::Linked { .. }
            | media::Source::Missing
            | _ => {},
        }
    }
    Ok(resources)
}

fn scan_active_content(content: &str, styles: Option<&str>) -> Result<ActiveContentStatus> {
    let mut status = ActiveContentStatus::default();
    scan_active_xml(content, &mut status)?;
    if let Some(style_xml) = styles {
        scan_active_xml(style_xml, &mut status)?;
    }
    Ok(status)
}

fn scan_active_xml(xml: &str, status: &mut ActiveContentStatus) -> Result<()> {
    validate_raw_xml_events(xml, "ODG active-content event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG active-content inventory XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG active-content event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if resolved_bound(&resolved_namespace, SCRIPT)
                    || resolved_bound(&resolved_namespace, OFFICE) && local == b"scripts"
                {
                    status.scripts = checked_active_count(status.scripts)?;
                }
                if resolved_bound(&resolved_namespace, XML_EVENTS) {
                    status.events = checked_active_count(status.events)?;
                }
                if resolved_bound(&resolved_namespace, PRESENTATION)
                    && matches!(local, b"event-listener" | b"show")
                {
                    status.actions = checked_active_count(status.actions)?;
                }
                if resolved_bound(&resolved_namespace, OFFICE) && local == b"dde-source" {
                    status.dde = checked_active_count(status.dde)?;
                }
                if resolved_bound(&resolved_namespace, DRAW)
                    && matches!(
                        local,
                        b"object" | b"object-ole" | b"plugin" | b"applet" | b"floating-frame"
                    )
                {
                    status.embedded_objects = checked_active_count(status.embedded_objects)?;
                }
                if attribute(&reader, &element, XLINK, b"href")?
                    .as_deref()
                    .is_some_and(is_external_href)
                {
                    status.external_links = checked_active_count(status.external_links)?;
                }
            },
            Event::DocType(_) => return invalid("DTD XML is prohibited in ODG active content"),
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    Ok(())
}

fn checked_active_count(count: usize) -> Result<usize> {
    let next_count = count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG active-content count overflow".into()))?;
    if next_count > MAX_SHAPES {
        return invalid("ODG active-content count exceeds the limit");
    }
    Ok(next_count)
}

fn is_external_href(href: &str) -> bool {
    href.contains("://")
        || href.starts_with("file:")
        || href.starts_with("mailto:")
        || href.starts_with("data:")
}

#[derive(Clone)]
struct ParsedStyleDefinition {
    style: Style,
    xml: String,
    span: Range<usize>,
}

#[derive(Clone)]
struct ParsedStyleResource {
    resource: StyleResource,
    xml: String,
    span: Range<usize>,
}

struct ActiveStyleResource {
    depth: usize,
    start: usize,
    resource: StyleResource,
}

fn parse_style_resources(xml: &str) -> Result<Vec<ParsedStyleResource>> {
    validate_raw_xml_events(xml, "ODG style-resource event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut active = Vec::new();
    let mut definitions = Vec::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG named style-resource XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG style-resource event")?;
        let namespace = classify(&resolved_namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                if namespace == NamespaceKind::Draw
                    && let Some(kind) = style_resource_kind(element.local_name().as_ref())
                {
                    if definitions.len().saturating_add(active.len()) >= MAX_STYLE_RESOURCES {
                        return invalid("ODG named style-resource count exceeds the limit");
                    }
                    let name = required_attribute(
                        &reader,
                        &element,
                        DRAW,
                        b"name",
                        "named style resource",
                    )?;
                    let attributes =
                        arbitrary_attributes(&reader, &element, &[(DRAW, b"name".as_slice())])?;
                    active.push(ActiveStyleResource {
                        depth,
                        start,
                        resource: StyleResource::parsed(kind, name, attributes),
                    });
                }
            },
            Event::Empty(element) => {
                if namespace == NamespaceKind::Draw
                    && let Some(kind) = style_resource_kind(element.local_name().as_ref())
                {
                    if definitions.len().saturating_add(active.len()) >= MAX_STYLE_RESOURCES {
                        return invalid("ODG named style-resource count exceeds the limit");
                    }
                    let name = required_attribute(
                        &reader,
                        &element,
                        DRAW,
                        b"name",
                        "named style resource",
                    )?;
                    definitions.push(ParsedStyleResource {
                        resource: StyleResource::parsed(
                            kind,
                            name,
                            arbitrary_attributes(&reader, &element, &[(DRAW, b"name".as_slice())])?,
                        ),
                        xml: clone_xml_span(xml, start..end, "ODG style-resource XML")?,
                        span: start..end,
                    });
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Draw
                    && style_resource_kind(element.local_name().as_ref()).is_some()
                    && active
                        .last()
                        .is_some_and(|resource| resource.depth == depth)
                {
                    let completed_resource = active.pop().ok_or_else(|| {
                        Error::InvalidFormat("ODG named style-resource span is missing".into())
                    })?;
                    definitions.push(ParsedStyleResource {
                        resource: completed_resource.resource,
                        xml: clone_xml_span(
                            xml,
                            completed_resource.start..end,
                            "ODG style-resource XML",
                        )?,
                        span: completed_resource.start..end,
                    });
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG named style-resource depth underflow".into())
                })?;
            },
            Event::DocType(_) => return invalid("DTD XML is prohibited in ODG style resources"),
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || !active.is_empty() {
        return invalid("ODG named style-resource XML is incomplete");
    }
    Ok(definitions)
}

fn style_resource_kind(local: &[u8]) -> Option<StyleResourceKind> {
    match local {
        b"gradient" => Some(StyleResourceKind::Gradient),
        b"hatch" => Some(StyleResourceKind::Hatch),
        b"fill-image" => Some(StyleResourceKind::FillImage),
        b"marker" => Some(StyleResourceKind::Marker),
        b"opacity" => Some(StyleResourceKind::Opacity),
        b"stroke-dash" => Some(StyleResourceKind::StrokeDash),
        _ => None,
    }
}

fn style_resource_reference_locals(kind: StyleResourceKind) -> &'static [&'static [u8]] {
    const GRADIENT: &[&[u8]] = &[b"fill-gradient-name"];
    const HATCH: &[&[u8]] = &[b"fill-hatch-name"];
    const FILL_IMAGE: &[&[u8]] = &[b"fill-image-name"];
    const MARKER: &[&[u8]] = &[b"marker-start", b"marker-end"];
    const OPACITY: &[&[u8]] = &[b"opacity-name"];
    const STROKE_DASH: &[&[u8]] = &[b"stroke-dash"];
    match kind {
        StyleResourceKind::Gradient => GRADIENT,
        StyleResourceKind::Hatch => HATCH,
        StyleResourceKind::FillImage => FILL_IMAGE,
        StyleResourceKind::Marker => MARKER,
        StyleResourceKind::Opacity => OPACITY,
        StyleResourceKind::StrokeDash => STROKE_DASH,
    }
}

fn style_resource_qualified_references(kind: StyleResourceKind) -> &'static [&'static [u8]] {
    const GRADIENT: &[&[u8]] = &[b"draw:fill-gradient-name"];
    const HATCH: &[&[u8]] = &[b"draw:fill-hatch-name"];
    const FILL_IMAGE: &[&[u8]] = &[b"draw:fill-image-name"];
    const MARKER: &[&[u8]] = &[b"draw:marker-start", b"draw:marker-end"];
    const OPACITY: &[&[u8]] = &[b"draw:opacity-name"];
    const STROKE_DASH: &[&[u8]] = &[b"draw:stroke-dash"];
    match kind {
        StyleResourceKind::Gradient => GRADIENT,
        StyleResourceKind::Hatch => HATCH,
        StyleResourceKind::FillImage => FILL_IMAGE,
        StyleResourceKind::Marker => MARKER,
        StyleResourceKind::Opacity => OPACITY,
        StyleResourceKind::StrokeDash => STROKE_DASH,
    }
}

fn style_resource_reference_kind(attribute: &str) -> Option<StyleResourceKind> {
    match attribute {
        "draw:fill-gradient-name" => Some(StyleResourceKind::Gradient),
        "draw:fill-hatch-name" => Some(StyleResourceKind::Hatch),
        "draw:fill-image-name" => Some(StyleResourceKind::FillImage),
        "draw:marker-start" | "draw:marker-end" => Some(StyleResourceKind::Marker),
        "draw:opacity-name" => Some(StyleResourceKind::Opacity),
        "draw:stroke-dash" => Some(StyleResourceKind::StrokeDash),
        _ => None,
    }
}

fn declares_style_resource_xml(xml: &str, kind: StyleResourceKind, name: &str) -> Result<bool> {
    parse_style_resources(xml).map(|definitions| {
        definitions.iter().any(|definition| {
            definition.resource.kind() == kind && definition.resource.name() == name
        })
    })
}

fn find_style_resource(
    snapshot: &Package,
    kind: StyleResourceKind,
    name: &str,
) -> Result<Option<ParsedStyleResource>> {
    let mut matches = parse_style_resources(snapshot.content_xml())?
        .into_iter()
        .filter(|definition| {
            definition.resource.kind() == kind && definition.resource.name() == name
        })
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return invalid("ODG named style-resource definition is ambiguous");
    }
    if let Some(definition) = matches.pop() {
        return Ok(Some(definition));
    }
    if let Some(styles) = snapshot.styles_xml() {
        let mut style_matches = parse_style_resources(styles)?
            .into_iter()
            .filter(|definition| {
                definition.resource.kind() == kind && definition.resource.name() == name
            })
            .collect::<Vec<_>>();
        if style_matches.len() > 1 {
            return invalid("ODG named style-resource definition is ambiguous");
        }
        return Ok(style_matches.pop());
    }
    Ok(None)
}

struct ActiveStyleDefinition {
    depth: usize,
    start: usize,
    name: String,
    family: String,
    parent: Option<String>,
    properties: BTreeMap<String, String>,
}

fn parse_style_definitions(xml: &str) -> Result<Vec<ParsedStyleDefinition>> {
    validate_raw_xml_events(xml, "ODG style-definition event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut active: Option<ActiveStyleDefinition> = None;
    let mut definitions = Vec::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG style catalog XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG style event")?;
        let namespace = classify(&resolved_namespace);
        let active_namespace = resolved_bound(&resolved_namespace, SCRIPT)
            || resolved_bound(&resolved_namespace, XML_EVENTS);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if active.is_some()
                    && (active_namespace
                        || namespace == NamespaceKind::Office && local.as_ref() == b"scripts")
                {
                    return invalid("active XML is prohibited in ODG style definitions");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    if active.is_some() {
                        return invalid("ODG style definitions cannot be nested");
                    }
                    active = Some(ActiveStyleDefinition {
                        depth,
                        start,
                        name: required_attribute(&reader, &element, STYLE, b"name", "style")?,
                        family: required_attribute(&reader, &element, STYLE, b"family", "style")?,
                        parent: attribute(&reader, &element, STYLE, b"parent-style-name")?,
                        properties: BTreeMap::new(),
                    });
                } else if namespace == NamespaceKind::Style
                    && local.as_ref().ends_with(b"-properties")
                    && let Some(value) = &mut active
                    && value.depth.checked_add(1) == Some(depth)
                {
                    let owner = std::str::from_utf8(local.as_ref()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG style property name: {error}"))
                    })?;
                    for (name, property) in arbitrary_attributes(&reader, &element, &[])? {
                        if value
                            .properties
                            .insert(format!("style:{owner}/{name}"), property)
                            .is_some()
                        {
                            return invalid("ODG style property attribute is duplicated");
                        }
                    }
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if active.is_some()
                    && (active_namespace
                        || namespace == NamespaceKind::Office && local.as_ref() == b"scripts")
                {
                    return invalid("active XML is prohibited in ODG style definitions");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    let name = required_attribute(&reader, &element, STYLE, b"name", "style")?;
                    let family = required_attribute(&reader, &element, STYLE, b"family", "style")?;
                    let parent = attribute(&reader, &element, STYLE, b"parent-style-name")?;
                    definitions.push(ParsedStyleDefinition {
                        style: Style::parsed(name, family, parent, BTreeMap::new()),
                        xml: clone_xml_span(xml, start..end, "ODG style XML")?,
                        span: start..end,
                    });
                } else if namespace == NamespaceKind::Style
                    && local.as_ref().ends_with(b"-properties")
                    && let Some(value) = &mut active
                    && value.depth == depth
                {
                    let owner = std::str::from_utf8(local.as_ref()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG style property name: {error}"))
                    })?;
                    for (name, property) in arbitrary_attributes(&reader, &element, &[])? {
                        if value
                            .properties
                            .insert(format!("style:{owner}/{name}"), property)
                            .is_some()
                        {
                            return invalid("ODG style property attribute is duplicated");
                        }
                    }
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Style
                    && local.as_ref() == b"style"
                    && active.as_ref().is_some_and(|value| value.depth == depth)
                {
                    let value = active.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG style source span is missing".into())
                    })?;
                    definitions.push(ParsedStyleDefinition {
                        style: Style::parsed(
                            value.name,
                            value.family,
                            value.parent,
                            value.properties,
                        ),
                        xml: clone_xml_span(xml, value.start..end, "ODG style XML")?,
                        span: value.start..end,
                    });
                }
                depth = depth.saturating_sub(1);
            },
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::DocType(_) => {
                return invalid("DTD XML is prohibited in ODG style catalogs");
            },
            Event::PI(_) if active.is_some() => {
                return invalid("processing instructions are prohibited in ODG style definitions");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || active.is_some() {
        return invalid("ODG style catalog XML is incomplete");
    }
    Ok(definitions)
}

/// Check direct and inherited page references to a transition style owner.
fn transition_style_is_shared(
    content: &str,
    styles: Option<&str>,
    pages: &[Page],
    selected_page: usize,
    selected_style: &str,
) -> Result<bool> {
    if pages.len() <= 1 {
        return Ok(false);
    }
    let parents = drawing_page_style_parents(content, styles)?;
    for (index, page) in pages.iter().enumerate() {
        if index == selected_page {
            continue;
        }
        if let Some(name) = page.style_name()
            && style_chain_contains(&parents, name, selected_style)?
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn drawing_page_style_parents(
    content: &str,
    styles: Option<&str>,
) -> Result<BTreeMap<String, Option<String>>> {
    let content_styles = parse_style_definitions(content)?;
    let named_styles = styles
        .map(parse_style_definitions)
        .transpose()?
        .unwrap_or_default();
    // Use the same family filter and automatic-style shadowing as transition
    // resolution. Content definitions are inserted last so they shadow a
    // same-named named style and its parent closure.
    let mut parents = BTreeMap::new();
    for definition in named_styles.iter().chain(&content_styles) {
        if definition.style.family() == "drawing-page" {
            parents.insert(
                definition.style.name().to_owned(),
                definition.style.parent().map(str::to_owned),
            );
        }
    }
    Ok(parents)
}

fn style_chain_contains(
    parents: &BTreeMap<String, Option<String>>,
    start: &str,
    target: &str,
) -> Result<bool> {
    let mut current = Some(start);
    let mut visited = BTreeSet::new();
    let mut depth = 0usize;
    while let Some(name) = current {
        if depth > MAX_DEPTH {
            return invalid("ODG drawing-page style inheritance exceeds the limit");
        }
        if !visited.insert(name) {
            return invalid("ODG drawing-page style inheritance is cyclic");
        }
        if name == target {
            return Ok(true);
        }
        current = parents.get(name).and_then(|parent| parent.as_deref());
        depth += 1;
    }
    Ok(false)
}

/// Resolve inert transition metadata referenced by each drawing page.
///
/// Automatic styles in content.xml shadow named styles from styles.xml. The
/// resolution is limited to the page's drawing-page style owner; a same-named
/// graphic style is never treated as a transition source.
fn parse_page_transitions(
    content: &str,
    styles: Option<&str>,
    pages: &[Page],
) -> Result<Vec<Option<Transition>>> {
    let content_styles = parse_style_definitions(content)?;
    let styles_styles = styles
        .map(parse_style_definitions)
        .transpose()?
        .unwrap_or_default();
    resolve_page_transitions(content, styles, pages, &content_styles, &styles_styles)
}

fn resolve_page_transitions(
    content: &str,
    styles: Option<&str>,
    pages: &[Page],
    content_styles: &[ParsedStyleDefinition],
    styles_styles: &[ParsedStyleDefinition],
) -> Result<Vec<Option<Transition>>> {
    let mut definitions = BTreeMap::new();
    for definition in styles_styles {
        if definition.style.family() != "drawing-page" {
            continue;
        }
        if definitions
            .insert(definition.style.name(), (definition, styles))
            .is_some()
        {
            return invalid("ODG drawing-page style definition is ambiguous");
        }
    }
    let mut content_names = BTreeSet::new();
    for definition in content_styles {
        if definition.style.family() != "drawing-page" {
            continue;
        }
        if !content_names.insert(definition.style.name()) {
            return invalid("ODG automatic drawing-page style definition is ambiguous");
        }
        // Automatic styles in content.xml shadow same-named definitions from
        // styles.xml, including their parent-style closure.
        definitions.insert(definition.style.name(), (definition, Some(content)));
    }
    // Direct transition attributes are source-scoped but independent of the
    // page whose inheritance chain reaches them. Keep only successful direct
    // parses for this resolution pass; recursive results remain per-page so
    // cycle and depth checks still run for every page.
    let mut direct_transitions = BTreeMap::new();
    pages
        .iter()
        .map(|page| -> Result<Option<Transition>> {
            let Some(name) = page.style_name() else {
                return Ok(None);
            };
            resolve_transition_style(
                name,
                &definitions,
                &mut BTreeSet::new(),
                0,
                &mut direct_transitions,
            )
        })
        .collect()
}

fn resolve_transition_style<'a>(
    name: &str,
    definitions: &BTreeMap<&'a str, (&'a ParsedStyleDefinition, Option<&'a str>)>,
    visiting: &mut BTreeSet<String>,
    depth: usize,
    direct_transitions: &mut BTreeMap<&'a str, Transition>,
) -> Result<Option<Transition>> {
    if depth > MAX_DEPTH {
        return invalid("ODG drawing-page style inheritance exceeds the limit");
    }
    let Some((definition, source)) = definitions.get(name).copied() else {
        return Ok(None);
    };
    if !visiting.insert(name.to_owned()) {
        return invalid("ODG drawing-page style inheritance is cyclic");
    }
    let mut transition = match direct_transitions.get(definition.style.name()) {
        Some(transition) => transition.clone(),
        None => {
            let transition = transition_values_from_style(&definition.style, source)?;
            direct_transitions.insert(definition.style.name(), transition.clone());
            transition
        },
    };
    if let Some(parent) = definition.style.parent()
        && let Some(parent_transition) =
            resolve_transition_style(parent, definitions, visiting, depth + 1, direct_transitions)?
    {
        transition.inherit_from(&parent_transition);
    }
    visiting.remove(name);
    Ok((!transition.is_empty()).then_some(transition))
}

fn transition_values_from_style(style: &Style, xml: Option<&str>) -> Result<Transition> {
    if style.family() != "drawing-page" {
        return Ok(Transition::new());
    }
    // Parse the transition attributes from the source-scoped owner rather
    // than the generic style projection.  The latter intentionally retains
    // raw QName prefixes, so using it here would miss a valid `p:*` alias for
    // the presentation namespace.
    let Some(xml) = xml else {
        return Ok(Transition::new());
    };
    let Some(owner) = find_transition_style_owner(xml, style.name())? else {
        return Ok(Transition::new());
    };
    let property = |name: &str| {
        owner
            .attributes
            .get(name)
            .map(|attribute| attribute.decoded.clone())
    };
    let transition = Transition::from_parts(
        property("presentation:transition-type"),
        property("presentation:transition-style"),
        property("presentation:transition-speed"),
        property("smil:type"),
        property("smil:subtype"),
        property("smil:direction"),
        property("smil:fadeColor"),
        property("presentation:duration"),
        owner.sound,
    )?;
    Ok(transition)
}

fn parse_transition_sound_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<crate::transition::Sound> {
    let href = required_attribute(reader, element, XLINK, b"href", "transition sound")?;
    let link_type = required_attribute(reader, element, XLINK, b"type", "transition sound")?;
    if link_type != "simple" {
        return invalid("ODG transition sound xlink:type must be 'simple'");
    }
    let play_full = optional_bool_attribute(reader, element, PRESENTATION, b"play-full")?;
    let actuate = attribute(reader, element, XLINK, b"actuate")?;
    if let Some(value) = actuate.as_deref()
        && value != "onRequest"
    {
        return invalid("ODG transition sound xlink:actuate must be 'onRequest'");
    }
    let show = attribute(reader, element, XLINK, b"show")?;
    let xml_id = attribute(reader, element, XML, b"id")?;
    crate::transition::Sound::new(href).and_then(|sound| {
        sound
            .with_play_full(play_full)
            .with_actuate_on_request(actuate.is_some())
            .with_show(show)?
            .with_xml_id(xml_id)
    })
}

#[derive(Clone)]
struct TransitionAttributeSpan {
    value: Range<usize>,
    token: Range<usize>,
    decoded: String,
}

struct TransitionStyleOwner {
    style_span: Range<usize>,
    style_open_span: Range<usize>,
    style_close_start: Option<usize>,
    property_open_span: Option<Range<usize>>,
    property_span: Option<Range<usize>>,
    sound_span: Option<Range<usize>>,
    sound: Option<crate::transition::Sound>,
    attributes: BTreeMap<String, TransitionAttributeSpan>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TransitionNamespaceBinding {
    Canonical,
    Foreign,
    Unbound,
}

fn transition_edit_is_lexically_reversible(
    source: &str,
    target: &str,
    style_name: &str,
    before: Option<&Transition>,
) -> bool {
    edit_transition_style_xml(target, style_name, before).is_ok_and(|restored| restored == source)
}

fn edit_transition_style_xml(
    source: &str,
    style_name: &str,
    transition: Option<&Transition>,
) -> Result<String> {
    validate_xml_source_size(source, "ODG transition XML")?;
    let owner = find_transition_style_owner(source, style_name)?;
    let Some(owner) = owner else {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odg.package.snapshot",
            "ODG drawing-page style has no transition owner",
        )));
    };
    let fields = transition_fields(transition);
    if owner.property_open_span.is_none() {
        let Some(transition) = transition else {
            return clone_bounded_string(source, MAX_OUTPUT_BYTES, "ODG transition XML");
        };
        let property = serialize_transition_properties(transition)?;
        let at = if let Some(close) = owner.style_close_start {
            close
        } else {
            owner
                .style_open_span
                .end
                .checked_sub(2)
                .ok_or_else(|| Error::InvalidFormat("ODG style start span is invalid".into()))?
        };
        return if owner.style_close_start.is_some() {
            insert_xml(source, at, &property)
        } else {
            insert_child_xml(source, at, &property)
        };
    }
    let mut current = clone_bounded_string(source, MAX_OUTPUT_BYTES, "ODG transition XML")?;
    for (qualified, value) in fields {
        current = rewrite_transition_attribute(&current, style_name, qualified, value)?;
    }
    let current_owner = find_transition_style_owner(&current, style_name)?
        .ok_or_else(|| Error::InvalidFormat("ODG transition style disappeared".into()))?;
    match transition.and_then(Transition::sound) {
        Some(sound) => {
            let sound_xml = serialize_transition_sound(sound)?;
            if let Some(span) = current_owner.sound_span {
                if current_owner.sound.as_ref() != Some(sound) {
                    return Err(Error::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "odg.package.snapshot",
                        "ODG transition sound edit would discard producer markup",
                    )));
                }
                let _ = span;
            } else if let Some(open) = current_owner.property_open_span {
                if current
                    .get(open.clone())
                    .is_some_and(|tag| tag.ends_with("/>"))
                {
                    let tag = current.get(open.clone()).ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODG transition property source span is invalid".into(),
                        )
                    })?;
                    let name_end = tag
                        .as_bytes()
                        .iter()
                        .enumerate()
                        .skip(1)
                        .find(|(_, byte)| {
                            byte.is_ascii_whitespace() || **byte == b'/' || **byte == b'>'
                        })
                        .map(|(index, _)| index)
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG transition property name is invalid".into())
                        })?;
                    let element_name = &tag[1..name_end];
                    let start_tag = tag.strip_suffix("/>").ok_or_else(|| {
                        Error::InvalidFormat("ODG transition property tag is invalid".into())
                    })?;
                    let replacement = format!(
                        "{start_tag}>{sound_xml}</{element_name}>",
                        element_name = element_name
                    );
                    current = replace_xml(&current, &open, &replacement)?;
                } else {
                    let at = insertion_before_tag_close(&current, &open)?;
                    current = insert_xml(&current, at, &sound_xml)?;
                }
            }
        },
        None => {
            if let Some(span) = current_owner.sound_span {
                let _ = span;
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedContent,
                    "odg.package.snapshot",
                    "ODG transition sound removal would discard producer markup",
                )));
            }
        },
    }
    Ok(current)
}

fn transition_fields(transition: Option<&Transition>) -> [(&'static str, Option<&str>); 8] {
    [
        (
            "presentation:transition-type",
            transition.and_then(Transition::transition_type),
        ),
        (
            "presentation:transition-style",
            transition.and_then(Transition::style),
        ),
        (
            "presentation:transition-speed",
            transition.and_then(Transition::speed),
        ),
        ("smil:type", transition.and_then(Transition::smil_type)),
        (
            "smil:subtype",
            transition.and_then(Transition::smil_subtype),
        ),
        ("smil:direction", transition.and_then(Transition::direction)),
        (
            "smil:fadeColor",
            transition.and_then(Transition::fade_color),
        ),
        (
            "presentation:duration",
            transition.and_then(Transition::duration),
        ),
    ]
}

fn validate_transition_budget(transition: Option<&Transition>) -> Result<()> {
    let Some(transition) = transition else {
        return Ok(());
    };
    let attribute_bytes = transition_fields(Some(transition))
        .iter()
        .map(|(name, value)| name.len().saturating_add(value.map_or(0, str::len)))
        .sum::<usize>();
    let sound_bytes = transition.sound().map_or(0, |sound| {
        sound
            .href()
            .len()
            .saturating_add(sound.show().map_or(0, str::len))
            .saturating_add(sound.xml_id().map_or(0, str::len))
    });
    if attribute_bytes.saturating_add(sound_bytes) > MAX_TRANSITION_XML_BYTES {
        return invalid("ODG transition metadata exceeds the aggregate limit");
    }
    Ok(())
}

fn validate_transition_sound_reference(
    transaction: &Edit,
    transition: Option<&Transition>,
) -> Result<()> {
    let Some(href) = transition
        .and_then(Transition::sound)
        .map(crate::transition::Sound::href)
    else {
        return Ok(());
    };
    if href.is_empty() {
        // The schema permits an empty anyIRI.  It denotes a same-document
        // reference and has no package member to resolve.
        return Ok(());
    }
    if is_linked_href(href) {
        return Ok(());
    }
    let path = resolve_package_path(href)?;
    if let Some(edit) = transaction
        .resource_edits
        .iter()
        .find(|edit| edit.path == path)
    {
        if edit.after_bytes.is_some() {
            return Ok(());
        }
        return invalid("ODG transition sound reference targets a removed resource");
    }
    if !transaction.source.0.package.package().has_file(&path)? {
        return invalid("ODG transition sound reference targets a missing resource");
    }
    Ok(())
}

fn validate_transition_xml_ids(
    transaction: &Edit,
    current: Option<&Transition>,
    desired: Option<&Transition>,
    in_content: bool,
) -> Result<()> {
    let Some(identifier) = desired
        .and_then(Transition::sound)
        .and_then(crate::transition::Sound::xml_id)
    else {
        return Ok(());
    };
    // XML IDs are document-scoped.  `content.xml` and `styles.xml` are
    // separate XML documents, so equal IDs across those parts are legal.
    let occurrences = if in_content {
        collect_xml_ids(&transaction.content)?
    } else {
        transaction
            .styles
            .as_deref()
            .map(collect_xml_ids)
            .transpose()?
            .unwrap_or_default()
    };
    let current_occurrence = current
        .and_then(Transition::sound)
        .and_then(crate::transition::Sound::xml_id)
        .is_some_and(|value| value == identifier);
    let allowed = usize::from(current_occurrence);
    if occurrences.get(identifier).copied().unwrap_or_default() > allowed {
        return invalid("ODG transition sound xml:id is already used in the package");
    }
    Ok(())
}

fn collect_xml_ids(xml: &str) -> Result<BTreeMap<String, usize>> {
    validate_raw_xml_events(xml, "ODG XML-id event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut identifiers = BTreeMap::new();
    loop {
        let (_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid ODG XML while checking xml:id values: {error}"
            ))
        })?;
        validate_borrowed_xml_event(&event, "ODG xml:id event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if let Some(identifier) = attribute(&reader, &element, XML, b"id")? {
                    let entry = identifiers.entry(identifier).or_insert(0usize);
                    *entry = entry
                        .checked_add(1)
                        .ok_or_else(|| Error::InvalidFormat("ODG xml:id count overflow".into()))?;
                    if *entry > MAX_SHAPES {
                        return invalid("ODG xml:id count exceeds the limit");
                    }
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG XML"),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    Ok(identifiers)
}

fn serialize_transition_properties(transition: &Transition) -> Result<String> {
    let prefix = "<style:drawing-page-properties xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:smil=\"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0\"";
    let sound_capacity = transition
        .sound()
        .map(transition_sound_capacity)
        .transpose()?;
    let mut capacity = prefix.len();
    for (name, value) in transition_fields(Some(transition)) {
        capacity = capacity
            .checked_add(serialized_attribute_len_allow_empty(name, value)?)
            .ok_or_else(|| Error::InvalidFormat("ODG transition XML size overflow".into()))?;
    }
    if transition.sound().is_some() {
        capacity = capacity
            .checked_add(1)
            .and_then(|size| size.checked_add(sound_capacity.unwrap_or_default()))
            .and_then(|size| size.checked_add("</style:drawing-page-properties>".len()))
            .ok_or_else(|| Error::InvalidFormat("ODG transition XML size overflow".into()))?;
    } else {
        capacity = capacity
            .checked_add(2)
            .ok_or_else(|| Error::InvalidFormat("ODG transition XML size overflow".into()))?;
    }
    let mut xml =
        precharge_xml_capacity(String::from(prefix), capacity, "ODG transition properties")?;
    for (name, value) in transition_fields(Some(transition)) {
        if matches!(name, "smil:type" | "smil:subtype") {
            push_attribute_allow_empty(&mut xml, name, value)?;
        } else {
            push_attribute(&mut xml, name, value)?;
        }
    }
    if let Some(sound) = transition.sound() {
        xml.push('>');
        xml.push_str(&serialize_transition_sound(sound)?);
        xml.push_str("</style:drawing-page-properties>");
    } else {
        xml.push_str("/>");
    }
    Ok(xml)
}

fn transition_sound_capacity(sound: &crate::transition::Sound) -> Result<usize> {
    let prefix = "<presentation:sound xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xlink:type=\"simple\"";
    let mut capacity = prefix.len();
    capacity = capacity
        .checked_add(serialized_attribute_len_allow_empty(
            "xlink:href",
            Some(sound.href()),
        )?)
        .ok_or_else(|| Error::InvalidFormat("ODG transition sound size overflow".into()))?;
    if sound.actuate_on_request() {
        capacity = capacity
            .checked_add(serialized_attribute_len(
                "xlink:actuate",
                Some("onRequest"),
            )?)
            .ok_or_else(|| Error::InvalidFormat("ODG transition sound size overflow".into()))?;
    }
    for (name, value) in [
        ("xlink:show", sound.show()),
        ("xml:id", sound.xml_id()),
        (
            "presentation:play-full",
            sound
                .play_full()
                .map(|value| if value { "true" } else { "false" }),
        ),
    ] {
        capacity = capacity
            .checked_add(serialized_attribute_len(name, value)?)
            .ok_or_else(|| Error::InvalidFormat("ODG transition sound size overflow".into()))?;
    }
    capacity
        .checked_add(2)
        .ok_or_else(|| Error::InvalidFormat("ODG transition sound size overflow".into()))
}

fn serialize_transition_sound(sound: &crate::transition::Sound) -> Result<String> {
    let prefix = "<presentation:sound xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xlink:type=\"simple\"";
    let mut xml = precharge_xml_capacity(
        String::from(prefix),
        transition_sound_capacity(sound)?,
        "ODG transition sound",
    )?;
    push_attribute_allow_empty(&mut xml, "xlink:href", Some(sound.href()))?;
    if sound.actuate_on_request() {
        push_attribute(&mut xml, "xlink:actuate", Some("onRequest"))?;
    }
    push_attribute(&mut xml, "xlink:show", sound.show())?;
    push_attribute(&mut xml, "xml:id", sound.xml_id())?;
    push_attribute(
        &mut xml,
        "presentation:play-full",
        sound
            .play_full()
            .map(|value| if value { "true" } else { "false" }),
    )?;
    xml.push_str("/>");
    Ok(xml)
}

fn push_attribute_allow_empty(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(attribute_value) = value else {
        return Ok(());
    };
    if attribute_value.len() > MAX_TEXT_BYTES || attribute_value.contains('\0') {
        return invalid("ODG XML attribute value is invalid");
    }
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    push_escaped_xml(output, attribute_value, "ODG transition attribute")?;
    output.push('"');
    Ok(())
}

fn rewrite_transition_attribute(
    source: &str,
    style_name: &str,
    qualified: &str,
    value: Option<&str>,
) -> Result<String> {
    let owner = find_transition_style_owner(source, style_name)?
        .ok_or_else(|| Error::InvalidFormat("ODG transition style disappeared".into()))?;
    let Some(open) = owner.property_open_span else {
        return Err(Error::InvalidFormat(
            "ODG transition property source span is missing".into(),
        ));
    };
    if let Some(attribute) = owner.attributes.get(qualified) {
        return match value {
            Some(value) => replace_xml_value(source, &attribute.value, value),
            None => remove_xml(source, &attribute.token),
        };
    }
    let Some(value) = value else {
        return clone_bounded_string(source, MAX_OUTPUT_BYTES, "ODG transition XML");
    };
    let mut insertion = String::new();
    let prefix = qualified
        .split_once(':')
        .map(|(prefix, _)| prefix)
        .unwrap_or("");
    let binding = source_transition_namespace_binding(source, &open, prefix)?;
    let attribute_name = match binding {
        TransitionNamespaceBinding::Canonical | TransitionNamespaceBinding::Unbound => {
            qualified.to_owned()
        },
        TransitionNamespaceBinding::Foreign => {
            let alias = unused_transition_namespace_prefix(source, &open, prefix)?;
            let local = qualified
                .split_once(':')
                .map(|(_, local)| local)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transition attribute is unqualified".into())
                })?;
            let namespace = transition_namespace_uri(prefix)?;
            if source_transition_namespace_binding_for(source, &open, &alias, namespace.as_bytes())?
                != TransitionNamespaceBinding::Canonical
            {
                write!(insertion, " xmlns:{alias}=\"{namespace}\"").map_err(|_| {
                    Error::InvalidFormat("ODG transition namespace write failed".into())
                })?;
            }
            format!("{alias}:{local}")
        },
    };
    if binding == TransitionNamespaceBinding::Unbound {
        let namespace = transition_namespace_uri(prefix)?;
        write!(insertion, " xmlns:{prefix}=\"{namespace}\"")
            .map_err(|_| Error::InvalidFormat("ODG transition namespace write failed".into()))?;
    }
    write_attribute(&mut insertion, &attribute_name, value)?;
    let at = insertion_before_tag_close(source, &open)?;
    insert_xml(source, at, &insertion)
}

fn transition_namespace_uri(prefix: &str) -> Result<&'static str> {
    match prefix {
        "presentation" => Ok("urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"),
        "smil" => Ok("urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"),
        _ => invalid("ODG transition attribute namespace is unsupported"),
    }
}

fn source_transition_namespace_binding(
    source: &str,
    open: &Range<usize>,
    prefix: &str,
) -> Result<TransitionNamespaceBinding> {
    let expected = transition_namespace_uri(prefix)?;
    source_transition_namespace_binding_for(source, open, prefix, expected.as_bytes())
}

fn source_transition_namespace_binding_for(
    source: &str,
    open: &Range<usize>,
    prefix: &str,
    expected: &[u8],
) -> Result<TransitionNamespaceBinding> {
    // Every caller passes the exact source that `find_transition_style_owner`
    // has just admitted with `validate_raw_xml_events`, so the raw per-event
    // and depth bounds already hold before this namespace-aware reader runs.
    let probe = format!("{prefix}:__litchi_namespace_probe");
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    loop {
        let start = position(&reader)?;
        if start > open.start {
            return Ok(TransitionNamespaceBinding::Unbound);
        }
        let (_resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transition namespace source: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transition namespace event")?;
        let end = position(&reader)?;
        if start != open.start {
            if matches!(event, Event::Eof) {
                return Ok(TransitionNamespaceBinding::Unbound);
            }
            continue;
        }
        if end < open.end {
            return Ok(TransitionNamespaceBinding::Unbound);
        }
        let (namespace, _local) = reader.resolver().resolve_attribute(QName(probe.as_bytes()));
        let binding = match namespace {
            ResolveResult::Bound(Namespace(uri)) if *uri == *expected => {
                TransitionNamespaceBinding::Canonical
            },
            ResolveResult::Bound(_) => TransitionNamespaceBinding::Foreign,
            ResolveResult::Unbound | ResolveResult::Unknown(_) => {
                TransitionNamespaceBinding::Unbound
            },
        };
        return Ok(binding);
    }
}

fn unused_transition_namespace_prefix(
    source: &str,
    open: &Range<usize>,
    occupied_prefix: &str,
) -> Result<String> {
    let expected = transition_namespace_uri(occupied_prefix)?.as_bytes();
    let base = match occupied_prefix {
        "presentation" => "litchi_presentation",
        "smil" => "litchi_smil",
        _ => "litchi_transition",
    };
    for index in 0..MAX_DEPTH {
        let candidate = if index == 0 {
            base.to_owned()
        } else {
            format!("{base}{index}")
        };
        match source_transition_namespace_binding_for(source, open, &candidate, expected)? {
            TransitionNamespaceBinding::Canonical | TransitionNamespaceBinding::Unbound => {
                return Ok(candidate);
            },
            TransitionNamespaceBinding::Foreign => {},
        }
    }
    invalid("ODG transition namespace prefix space is exhausted")
}

fn insertion_before_tag_close(source: &str, open: &Range<usize>) -> Result<usize> {
    let tag = source
        .get(open.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG transition start tag span is invalid".into()))?;
    if tag.ends_with("/>") {
        Ok(open.end - 2)
    } else if tag.ends_with('>') {
        Ok(open.end - 1)
    } else {
        invalid("ODG transition start tag is unterminated")
    }
}

fn write_attribute(target: &mut String, name: &str, value: &str) -> Result<()> {
    target.push(' ');
    target.push_str(name);
    target.push_str("=\"");
    push_escaped_xml(target, value, "ODG transition attribute")?;
    target.push('"');
    Ok(())
}

fn find_transition_style_owner(
    xml: &str,
    style_name: &str,
) -> Result<Option<TransitionStyleOwner>> {
    validate_raw_xml_events(xml, "ODG transition-style event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut style: Option<(usize, usize, Range<usize>, bool)> = None;
    let mut property: Option<(usize, usize, Range<usize>)> = None;
    let mut active_sound: Option<(usize, usize)> = None;
    let mut owner = None;
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transition XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transition event")?;
        let namespace = classify(&resolved_namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if active_sound.is_some() {
                    return invalid("ODG transition sound must have empty content");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    let name = required_attribute(&reader, &element, STYLE, b"name", "style")?;
                    let family = required_attribute(&reader, &element, STYLE, b"family", "style")?;
                    if name == style_name && family == "drawing-page" {
                        if style.is_some() || owner.is_some() {
                            return invalid("ODG transition style is ambiguous");
                        }
                        style = Some((depth, start, start..end, false));
                    }
                } else if namespace == NamespaceKind::Style
                    && local.as_ref() == b"drawing-page-properties"
                    && style
                        .as_ref()
                        .is_some_and(|(style_depth, _, _, _)| *style_depth + 1 == depth)
                {
                    if property.is_some() {
                        return invalid("ODG drawing-page transition property is duplicated");
                    }
                    let tag = xml.as_bytes().get(start..end).ok_or_else(|| {
                        Error::InvalidFormat("ODG transition property span is invalid".into())
                    })?;
                    property = Some((depth, start, start..end));
                    let attributes = transition_attribute_spans(&reader, &element, tag, start)?;
                    if let Some((_, _, open)) = property.as_mut() {
                        let mut transition_owner = TransitionStyleOwner {
                            style_span: 0..0,
                            style_open_span: 0..0,
                            style_close_start: None,
                            property_open_span: Some(open.clone()),
                            property_span: None,
                            sound_span: None,
                            sound: None,
                            attributes,
                        };
                        if let Some((_, style_start, style_open, _)) = style.as_ref() {
                            transition_owner.style_span = *style_start..0;
                            transition_owner.style_open_span = style_open.clone();
                        }
                        owner = Some(transition_owner);
                    }
                } else if namespace == NamespaceKind::Presentation
                    && local.as_ref() == b"sound"
                    && property
                        .as_ref()
                        .is_some_and(|(property_depth, _, _)| *property_depth + 1 == depth)
                {
                    if let Some(current) = owner.as_mut() {
                        if current.sound_span.is_some() {
                            return invalid("ODG transition sound is duplicated");
                        }
                        active_sound = Some((depth, start));
                        current.sound = Some(parse_transition_sound_element(&reader, &element)?);
                    }
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if active_sound.is_some() {
                    return invalid("ODG transition sound must have empty content");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    let name = required_attribute(&reader, &element, STYLE, b"name", "style")?;
                    let family = required_attribute(&reader, &element, STYLE, b"family", "style")?;
                    if name == style_name && family == "drawing-page" {
                        if style.is_some() || owner.is_some() {
                            return invalid("ODG transition style is ambiguous");
                        }
                        owner = Some(TransitionStyleOwner {
                            style_span: start..end,
                            style_open_span: start..end,
                            style_close_start: None,
                            property_open_span: None,
                            property_span: None,
                            sound_span: None,
                            sound: None,
                            attributes: BTreeMap::new(),
                        });
                    }
                } else if namespace == NamespaceKind::Style
                    && local.as_ref() == b"drawing-page-properties"
                    && style
                        .as_ref()
                        .is_some_and(|(style_depth, _, _, _)| *style_depth + 1 == depth + 1)
                {
                    if property.is_some() {
                        return invalid("ODG drawing-page transition property is duplicated");
                    }
                    let tag = xml.as_bytes().get(start..end).ok_or_else(|| {
                        Error::InvalidFormat("ODG transition property span is invalid".into())
                    })?;
                    let attributes = transition_attribute_spans(&reader, &element, tag, start)?;
                    if let Some((_, style_start, style_open, _)) = style.as_ref() {
                        owner = Some(TransitionStyleOwner {
                            style_span: *style_start..end,
                            style_open_span: style_open.clone(),
                            style_close_start: None,
                            property_open_span: Some(start..end),
                            property_span: Some(start..end),
                            sound_span: None,
                            sound: None,
                            attributes,
                        });
                    }
                } else if namespace == NamespaceKind::Presentation
                    && local.as_ref() == b"sound"
                    && property
                        .as_ref()
                        .is_some_and(|(property_depth, _, _)| *property_depth + 1 == depth + 1)
                {
                    if let Some(current) = owner.as_mut() {
                        if current.sound_span.is_some() {
                            return invalid("ODG transition sound is duplicated");
                        }
                        current.sound_span = Some(start..end);
                        current.sound = Some(parse_transition_sound_element(&reader, &element)?);
                    }
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if active_sound
                    .as_ref()
                    .is_some_and(|(sound_depth, _)| *sound_depth == depth)
                {
                    if namespace != NamespaceKind::Presentation || local.as_ref() != b"sound" {
                        return invalid("ODG transition sound is incomplete");
                    }
                    let (_, sound_start) = active_sound.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG transition sound source is missing".into())
                    })?;
                    if let Some(current) = owner.as_mut() {
                        current.sound_span = Some(sound_start..end);
                    }
                }
                if property
                    .as_ref()
                    .is_some_and(|(property_depth, _, _)| *property_depth == depth)
                {
                    if let Some(current) = owner.as_mut() {
                        if let Some((_, property_start, _)) = property.take() {
                            current.property_span = Some(property_start..end);
                        }
                    }
                }
                if style
                    .as_ref()
                    .is_some_and(|(style_depth, _, _, empty)| !*empty && *style_depth == depth)
                    && namespace == NamespaceKind::Style
                    && local.as_ref() == b"style"
                {
                    if let Some((_, style_start, style_open, _)) = style.take() {
                        if let Some(current) = owner.as_mut() {
                            current.style_span = style_start..end;
                            current.style_close_start = Some(start);
                        } else {
                            owner = Some(TransitionStyleOwner {
                                style_span: style_start..end,
                                style_open_span: style_open,
                                style_close_start: Some(start),
                                property_open_span: None,
                                property_span: None,
                                sound_span: None,
                                sound: None,
                                attributes: BTreeMap::new(),
                            });
                        }
                    }
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG transition XML depth underflow".into())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG transition XML"),
            Event::Eof => break,
            Event::CData(text) if active_sound.is_some() => {
                if !text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return invalid("ODG transition sound must have empty content");
                }
            },
            Event::Text(text) if active_sound.is_some() => {
                if !text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
                {
                    return invalid("ODG transition sound must have empty content");
                }
            },
            Event::GeneralRef(_) if active_sound.is_some() => {
                return invalid("ODG transition sound must have empty content");
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || style.is_some() || property.is_some() || active_sound.is_some() {
        return invalid("ODG transition XML is incomplete");
    }
    Ok(owner)
}

fn transition_attribute_spans(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<BTreeMap<String, TransitionAttributeSpan>> {
    let mut result = BTreeMap::new();
    for raw in element.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transition attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let qualified = match (namespace, local.as_ref()) {
            (ResolveResult::Bound(Namespace(uri)), b"transition-type") if *uri == *PRESENTATION => {
                "presentation:transition-type"
            },
            (ResolveResult::Bound(Namespace(uri)), b"transition-style")
                if *uri == *PRESENTATION =>
            {
                "presentation:transition-style"
            },
            (ResolveResult::Bound(Namespace(uri)), b"transition-speed")
                if *uri == *PRESENTATION =>
            {
                "presentation:transition-speed"
            },
            (ResolveResult::Bound(Namespace(uri)), b"duration") if *uri == *PRESENTATION => {
                "presentation:duration"
            },
            (ResolveResult::Bound(Namespace(uri)), b"type") if *uri == *SMIL => "smil:type",
            (ResolveResult::Bound(Namespace(uri)), b"subtype") if *uri == *SMIL => "smil:subtype",
            (ResolveResult::Bound(Namespace(uri)), b"direction") if *uri == *SMIL => {
                "smil:direction"
            },
            (ResolveResult::Bound(Namespace(uri)), b"fadeColor") if *uri == *SMIL => {
                "smil:fadeColor"
            },
            _ => continue,
        };
        let raw_name = attribute.key.as_ref();
        validate_raw_xml_value(raw_name, "ODG transition attribute name")?;
        validate_raw_xml_value(attribute.value.as_ref(), "ODG transition attribute value")?;
        let (value_start, value_end) = attribute_value_span(tag, raw_name)?;
        let token = attribute_token_span(tag, raw_name)?;
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG transition attribute value: {error}"))
            })?;
        validate_decoded_xml_value(decoded.as_ref(), "ODG transition attribute value")?;
        let decoded = own_bounded_cow(decoded, MAX_TEXT_BYTES, "ODG transition attribute value")?;
        if result
            .insert(
                qualified.to_owned(),
                TransitionAttributeSpan {
                    value: tag_start + value_start..tag_start + value_end,
                    token: tag_start + token.start..tag_start + token.end,
                    decoded,
                },
            )
            .is_some()
        {
            return invalid("ODG transition attribute is duplicated");
        }
    }
    Ok(result)
}

fn checked_xml_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
    if next_depth > MAX_DEPTH {
        return invalid("ODG XML nesting exceeds the limit");
    }
    Ok(next_depth)
}

fn push_declared_layer(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    layers: &mut Vec<Layer>,
) -> Result<()> {
    if layers.len() >= MAX_LAYERS {
        return invalid("ODG declared layer count exceeds the limit");
    }
    layers.push(Layer::parsed(
        required_attribute(reader, element, DRAW, b"name", "draw:layer")?,
        attribute(reader, element, DRAW, b"display")?,
        optional_bool_attribute(reader, element, DRAW, b"protected")?,
    ));
    Ok(())
}

fn text_value(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
    // The raw escaped form bounds every decoder/unescape allocation below.
    validate_raw_xml_value(text.as_ref(), "ODG text")?;
    let decoded = text
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text: {error}")))?;
    // Check the borrowed Cow before converting it into an owned String.
    validate_decoded_xml_value(decoded.as_ref(), "ODG text")?;
    let unescaped = quick_xml::escape::unescape(decoded.as_ref())
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text escape: {error}")))?;
    let unescaped = own_bounded_cow(unescaped, MAX_TEXT_BYTES, "ODG text")?;
    validate_decoded_xml_value(&unescaped, "ODG text")?;
    Ok(unescaped)
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    validate_raw_xml_value(reference.as_ref(), "ODG character reference")?;
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODG character reference: {error}"))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "apos" => Ok("'".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("ODG custom entities are not allowed"),
    }
}

fn content_splice_publication(
    source: &Package,
    splices: &[ContentSplice],
) -> Result<XmlSplicePublication> {
    let source_part = XmlSourcePart::load(source.0.package.package(), "content.xml")?;
    if source_part.bytes() != source.content_xml().as_bytes() {
        return invalid("ODG content splice has different package provenance");
    }
    let mut publication = XmlSplicePublication::new(source_part.clone());
    for splice in splices {
        let proof = source_part.checked_range(splice.source_range.clone(), &splice.expected)?;
        let fragment = if splice.replacement.is_empty() {
            AuthoredXmlFragment::deletion()
        } else if splice.replacement.first() == Some(&b'<') {
            AuthoredXmlFragment::markup(splice.replacement.clone())?
        } else {
            AuthoredXmlFragment::text(splice.replacement.clone())?
        };
        publication.replace(proof, fragment)?;
    }
    Ok(publication)
}

fn rebuild_spliced(
    source: &Package,
    content: XmlSplicePublication,
    styles: Option<&str>,
    replacements: &[ResourceReplacement<'_>],
    security_policy: SecurityWritePolicy,
) -> Result<Vec<u8>> {
    enforce_security_policy(source, security_policy)?;
    let archive = source.0.package.package();
    let mut writer = PackageWriter::new_bounded(
        MAX_OUTPUT_BYTES,
        litchi_odf_common::core::StreamingArchiveLimits::default(),
    );
    writer.set_mimetype(source.0.mimetype)?;
    content.publish(&mut writer)?;
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes())?;
    } else if source.0.package.package().has_file("styles.xml")? {
        // Preserve an untouched producer styles.xml through the checked splice
        // path.  Exact-source publication intentionally accepts formatting
        // that authored XML publication would reject.
        XmlSplicePublication::new(XmlSourcePart::load(
            source.0.package.package(),
            "styles.xml",
        )?)
        .publish(&mut writer)?;
    }
    for path in ["meta.xml", "settings.xml"] {
        if archive.has_file(path)? {
            XmlSplicePublication::new(XmlSourcePart::load(archive, path)?).publish(&mut writer)?;
        }
    }
    let mut excluded = replacements
        .iter()
        .map(|replacement| replacement.path.to_owned())
        .collect::<Vec<_>>();
    excluded.push("settings.xml".to_string());
    writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
    for replacement in replacements {
        if let Some(bytes) = replacement.bytes {
            writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
        }
    }
    writer.finish_to_bounded_bytes()
}

fn rebuild(
    source: &Package,
    content: &str,
    styles: Option<&str>,
    replacements: &[ResourceReplacement<'_>],
    security_policy: SecurityWritePolicy,
) -> Result<Vec<u8>> {
    enforce_security_policy(source, security_policy)?;
    let archive = source.0.package.package();
    let mut writer = PackageWriter::new_bounded(
        MAX_OUTPUT_BYTES,
        litchi_odf_common::core::StreamingArchiveLimits::default(),
    );
    writer.set_mimetype(source.0.mimetype)?;
    writer.add_file("content.xml", content.as_bytes())?;
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes())?;
    } else if source.0.package.package().has_file("styles.xml")? {
        writer.add_file(
            "styles.xml",
            &source.0.package.package().get_file("styles.xml")?,
        )?;
    }
    for path in ["meta.xml", "settings.xml"] {
        if archive.has_file(path)? {
            writer.add_file(path, &archive.get_file(path)?)?;
        }
    }
    let mut excluded = replacements
        .iter()
        .map(|replacement| replacement.path.to_owned())
        .collect::<Vec<_>>();
    excluded.push("settings.xml".to_string());
    writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
    for replacement in replacements {
        if let Some(bytes) = replacement.bytes {
            writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
        }
    }
    writer.finish_to_bounded_bytes()
}

fn enforce_security_policy(source: &Package, policy: SecurityWritePolicy) -> Result<()> {
    if source.security().is_signed() && policy == SecurityWritePolicy::Refuse {
        return invalid("ODG package edits refuse signed packages");
    }
    if source.security().is_encrypted() {
        return invalid("ODG package edits refuse encrypted packages");
    }
    Ok(())
}

fn enforce_active_content_policy(source: &Package, policy: ActiveContentWritePolicy) -> Result<()> {
    if policy == ActiveContentWritePolicy::Refuse && source.active_content().is_present() {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odg.package.snapshot",
            "ODG active-content write policy refuses the source inventory",
        )));
    }
    Ok(())
}

fn ensure_compact_rewrite_source(source: &Package) -> Result<()> {
    let archive = source.0.package.package();
    for path in source.files()? {
        let is_xml = Path::new(&path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"));
        if is_xml && path != "META-INF/manifest.xml" {
            compact_xml::validate(&archive.get_file(&path)?, compact_xml::Limits::default())
                .map_err(Error::from)?;
        }
    }
    Ok(())
}

fn replace_xml_value(source: &str, span: &Range<usize>, replacement: &str) -> Result<String> {
    if span.start > span.end || span.end > source.len() {
        return invalid("ODG text source span is invalid");
    }
    let escaped_replacement_len = escaped_xml_len(replacement)?;
    let capacity = source
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|size| size.checked_add(escaped_replacement_len))
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".to_string()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..span.start]);
    push_escaped_xml(&mut output, replacement, "ODG edited content")?;
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn stage_content_splice(
    source: &[u8],
    current: &[u8],
    splices: &mut Vec<ContentSplice>,
    range: &Range<usize>,
    replacement: &[u8],
) -> Result<()> {
    let actual = current
        .get(range.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice range is invalid".into()))?;
    if let Some(index) = splices
        .iter()
        .position(|splice| ranges_overlap_or_conflict(&splice.current_range, range))
    {
        if splices[index].current_range != *range || actual != splices[index].replacement {
            return invalid("ODG content splice overlaps an earlier semantic edit");
        }
        let old_end = splices[index].current_range.end;
        let new_end = range
            .start
            .checked_add(replacement.len())
            .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
        splices[index].current_range.end = new_end;
        splices[index].replacement = replacement.to_vec();
        shift_current_splices(splices, index, old_end, new_end)?;
        return Ok(());
    }

    // A markup edit touching an earlier markup edit (two pages appended at
    // the same point, a removal next to an insertion) joins it: two
    // publication edits may not share one source boundary.
    if is_markup_splice(replacement)
        && let Some(index) = splices.iter().position(|splice| {
            is_markup_splice(&splice.replacement)
                && (splice.current_range.end == range.start
                    || splice.current_range.start == range.end)
        })
    {
        return merge_content_splice(source, actual, splices, index, range, replacement);
    }
    let (removed_before, added_before) = splices
        .iter()
        .filter(|splice| splice.current_range.end <= range.start)
        .try_fold((0usize, 0usize), |(removed, added), splice| {
            Ok::<_, Error>((
                removed
                    .checked_add(splice.source_range.len())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG content splice size overflow".into())
                    })?,
                added
                    .checked_add(splice.current_range.len())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG content splice size overflow".into())
                    })?,
            ))
        })?;
    let source_start = range
        .start
        .checked_sub(added_before)
        .and_then(|value| value.checked_add(removed_before))
        .ok_or_else(|| Error::InvalidFormat("ODG content splice mapping is invalid".into()))?;
    let source_end = source_start
        .checked_add(range.len())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
    let source_range = source_start..source_end;
    let expected = source
        .get(source_range.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice source range is invalid".into()))?;
    if expected != actual {
        return invalid("ODG content splice lost exact source provenance");
    }
    let new_end = range
        .start
        .checked_add(replacement.len())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
    let insertion_index = splices.len();
    splices.push(ContentSplice {
        source_range,
        current_range: range.start..new_end,
        expected: expected.to_vec(),
        replacement: replacement.to_vec(),
    });
    shift_current_splices(splices, insertion_index, range.end, new_end)?;
    splices.sort_unstable_by_key(|splice| splice.current_range.start);
    Ok(())
}

/// Whether a splice replacement is a deletion or balanced markup (not an
/// escaped text or attribute value).
fn is_markup_splice(replacement: &[u8]) -> bool {
    replacement.first().is_none_or(|byte| *byte == b'<')
}

/// Join the markup edit of `range` with the adjacent staged splice `index`.
fn merge_content_splice(
    source: &[u8],
    actual: &[u8],
    splices: &mut [ContentSplice],
    index: usize,
    range: &Range<usize>,
    replacement: &[u8],
) -> Result<()> {
    let overflow = || Error::InvalidFormat("ODG content splice size overflow".into());
    let held = &splices[index];
    let append = held.current_range.end == range.start;
    let length = range.len();
    let source_range = if append {
        held.source_range.start
            ..held
                .source_range
                .end
                .checked_add(length)
                .ok_or_else(overflow)?
    } else {
        held.source_range
            .start
            .checked_sub(length)
            .ok_or_else(overflow)?..held.source_range.end
    };
    let expected = source
        .get(source_range.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice source range is invalid".into()))?
        .to_vec();
    let mut joined = Vec::new();
    joined
        .try_reserve_exact(
            held.replacement
                .len()
                .checked_add(replacement.len())
                .ok_or_else(overflow)?,
        )
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG content splice",
            source: allocation_error,
        })?;
    let current_start = if append {
        joined.extend_from_slice(&held.replacement);
        joined.extend_from_slice(replacement);
        held.current_range.start
    } else {
        joined.extend_from_slice(replacement);
        joined.extend_from_slice(&held.replacement);
        range.start
    };
    // The bytes the new edit removes must be the source bytes next to the
    // held splice (no other edit rewrote them).
    let removed = if append {
        &expected[expected.len() - length..]
    } else {
        &expected[..length]
    };
    if removed != actual {
        return invalid("ODG content splice lost exact source provenance");
    }
    let old_end = range.end;
    let new_end = range
        .start
        .checked_add(replacement.len())
        .ok_or_else(overflow)?;
    let current_end = current_start
        .checked_add(joined.len())
        .ok_or_else(overflow)?;
    let splice = &mut splices[index];
    splice.source_range = source_range;
    splice.expected = expected;
    splice.replacement = joined;
    splice.current_range = current_start..current_end;
    shift_current_splices(splices, index, old_end, new_end)?;
    splices.sort_unstable_by_key(|splice| splice.current_range.start);
    Ok(())
}

fn shift_current_splices(
    splices: &mut [ContentSplice],
    changed: usize,
    old_end: usize,
    new_end: usize,
) -> Result<()> {
    for (index, splice) in splices.iter_mut().enumerate() {
        if index == changed || splice.current_range.start < old_end {
            continue;
        }
        splice.current_range.start = splice
            .current_range
            .start
            .checked_sub(old_end)
            .and_then(|offset| new_end.checked_add(offset))
            .ok_or_else(|| Error::InvalidFormat("ODG content splice shift is invalid".into()))?;
        splice.current_range.end = splice
            .current_range
            .end
            .checked_sub(old_end)
            .and_then(|offset| new_end.checked_add(offset))
            .ok_or_else(|| Error::InvalidFormat("ODG content splice shift is invalid".into()))?;
    }
    Ok(())
}

fn ranges_overlap_or_conflict(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
        || (left.start == left.end && right.start == right.end && left.start == right.start)
}

fn insert_xml(source: &str, at: usize, xml: &str) -> Result<String> {
    if at > source.len() || !source.is_char_boundary(at) {
        return invalid("ODG XML insertion point is invalid");
    }
    let capacity = source
        .len()
        .checked_add(xml.len())
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".into()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..at]);
    output.push_str(xml);
    output.push_str(&source[at..]);
    Ok(output)
}

fn insert_automatic_style(source: &str, style: &str) -> Result<String> {
    validate_raw_xml_events(source, "ODG automatic-style event")?;
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut automatic_depth = None;
    let mut insertion = None;
    let mut body_start = None;
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG automatic-style owner XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG automatic-style event")?;
        let namespace = classify(&resolved_namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Office && local.as_ref() == b"automatic-styles" {
                    if automatic_depth.replace(depth).is_some() {
                        return invalid("ODG automatic-styles owner is duplicated");
                    }
                } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
                    body_start = Some(start);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"automatic-styles"
                    && insertion.replace(end.saturating_sub(2)).is_some()
                {
                    return invalid("ODG automatic-styles owner is duplicated");
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"automatic-styles"
                    && automatic_depth == Some(depth)
                {
                    insertion = Some(start);
                    automatic_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG automatic styles");
            },
            Event::Eof => break,
            Event::CData(_) | Event::Comment(_) | Event::Decl(_) | Event::Text(_) => {},
        }
    }
    if let Some(at) = insertion {
        return insert_child_xml(source, at, style);
    }
    let at = body_start
        .ok_or_else(|| Error::InvalidFormat("ODG office:body source span is missing".into()))?;
    let prefix = format!(
        "<office:automatic-styles xmlns:office=\"{}\">",
        std::str::from_utf8(OFFICE).unwrap_or_default()
    );
    let suffix = "</office:automatic-styles>";
    let owner_len = prefix
        .len()
        .checked_add(style.len())
        .and_then(|size| size.checked_add(suffix.len()))
        .ok_or_else(|| Error::InvalidFormat("ODG automatic-style size overflow".into()))?;
    if source
        .len()
        .checked_add(owner_len)
        .is_none_or(|size| size > MAX_OUTPUT_BYTES)
    {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut owner = String::new();
    owner
        .try_reserve_exact(owner_len)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG automatic styles",
            source: allocation_error,
        })?;
    owner.push_str(&prefix);
    owner.push_str(style);
    owner.push_str(suffix);
    insert_xml(source, at, &owner)
}

/// The replaced range and balanced replacement markup that insert `child`
/// at `at`: an insertion point, or the `/>` of an empty owner element, which
/// is expanded whole (`<owner …/>` becomes `<owner …>child</owner>`) so the
/// replacement is balanced markup the splice publication accepts.
fn child_insertion(source: &str, at: usize, child: &str) -> Result<(Range<usize>, String)> {
    if source.as_bytes().get(at..at.saturating_add(2)) != Some(b"/>") {
        if at > source.len() || !source.is_char_boundary(at) {
            return invalid("ODG XML insertion point is invalid");
        }
        return Ok((at..at, child.to_owned()));
    }
    let element_start = source
        .get(..at)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner start is missing".into()))?;
    let name_start = element_start + 1;
    let name_end = source
        .as_bytes()
        .get(name_start..at)
        .and_then(|bytes| {
            bytes
                .iter()
                .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
                .map(|offset| name_start + offset)
        })
        .unwrap_or(at);
    let name = source
        .get(name_start..name_end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner name is missing".into()))?;
    let open = source[element_start..at].trim_end();
    let mut replacement = String::new();
    replacement
        .try_reserve_exact(open.len() + child.len() + name.len() + 4)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    replacement.push_str(open);
    replacement.push('>');
    replacement.push_str(child);
    replacement.push_str("</");
    replacement.push_str(name);
    replacement.push('>');
    Ok((element_start..at + 2, replacement))
}

fn insert_child_xml(source: &str, at: usize, child: &str) -> Result<String> {
    if source.as_bytes().get(at..at.saturating_add(2)) != Some(b"/>") {
        return insert_xml(source, at, child);
    }
    let element_start = source
        .get(..at)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner start is missing".into()))?;
    let name_start = element_start + 1;
    let name_end = source
        .as_bytes()
        .get(name_start..at)
        .and_then(|bytes| {
            bytes
                .iter()
                .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
                .map(|offset| name_start + offset)
        })
        .unwrap_or(at);
    let name = source
        .get(name_start..name_end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner name is missing".into()))?;
    let replacement_len = 1usize
        .checked_add(child.len())
        .and_then(|size| size.checked_add(name.len()))
        .and_then(|size| size.checked_add(3))
        .ok_or_else(|| Error::InvalidFormat("ODG XML insertion size overflow".into()))?;
    let capacity = source
        .len()
        .checked_sub(2)
        .and_then(|size| size.checked_add(replacement_len))
        .ok_or_else(|| Error::InvalidFormat("ODG XML insertion size overflow".into()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let replacement = format!(">{child}</{name}>");
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..at]);
    output.push_str(&replacement);
    output.push_str(&source[at + 2..]);
    Ok(output)
}

fn remove_xml(source: &str, span: &Range<usize>) -> Result<String> {
    if span.start > span.end
        || span.end > source.len()
        || !source.is_char_boundary(span.start)
        || !source.is_char_boundary(span.end)
    {
        return invalid("ODG XML removal span is invalid");
    }
    let capacity = source.len() - (span.end - span.start);
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..span.start]);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn replace_xml(source: &str, span: &Range<usize>, replacement: &str) -> Result<String> {
    let without = remove_xml(source, span)?;
    insert_xml(&without, span.start, replacement)
}

fn start_tag_end(source: &str, start: usize) -> Result<usize> {
    source
        .get(start..)
        .and_then(|tail| tail.find('>').map(|offset| start + offset + 1))
        .ok_or_else(|| Error::InvalidFormat("ODG element start tag is unterminated".into()))
}

fn unique_page_transition_style_name(
    content: &str,
    styles: Option<&str>,
    pages: &[Page],
) -> Result<String> {
    let mut names = parse_style_definitions(content)?
        .into_iter()
        .map(|definition| definition.style.name().to_owned())
        .collect::<BTreeSet<_>>();
    if let Some(styles) = styles {
        names.extend(
            parse_style_definitions(styles)?
                .into_iter()
                .map(|definition| definition.style.name().to_owned()),
        );
    }
    names.extend(pages.iter().filter_map(Page::style_name).map(str::to_owned));
    for index in 0..MAX_PAGES {
        let candidate = format!("LitchiPageTransition{index}");
        if names.insert(candidate.clone()) {
            return Ok(candidate);
        }
    }
    invalid("ODG detached page transition style names are exhausted")
}

fn serialize_detached_page_style(name: &str, transition: &Transition) -> Result<String> {
    validate_bounded_value(name, "ODG detached page transition style name")?;
    let properties = serialize_transition_properties(transition)?;
    let prefix = format!(
        "<style:style xmlns:style=\"{}\"",
        std::str::from_utf8(STYLE).unwrap_or_default()
    );
    let name_capacity = serialized_attribute_len("style:name", Some(name))?;
    let family_capacity = serialized_attribute_len("style:family", Some("drawing-page"))?;
    let mut capacity = prefix.len();
    capacity = capacity
        .checked_add(name_capacity)
        .and_then(|size| size.checked_add(family_capacity))
        .and_then(|size| size.checked_add(1))
        .and_then(|size| size.checked_add(properties.len()))
        .and_then(|size| size.checked_add("</style:style>".len()))
        .ok_or_else(|| Error::InvalidFormat("ODG detached page style size overflow".into()))?;
    let mut xml = precharge_xml_capacity(prefix, capacity, "ODG detached page style")?;
    push_attribute(&mut xml, "style:name", Some(name))?;
    push_attribute(&mut xml, "style:family", Some("drawing-page"))?;
    xml.push('>');
    xml.push_str(&properties);
    xml.push_str("</style:style>");
    Ok(xml)
}

fn serialize_page(page: &Page) -> Result<String> {
    let prefix = format!(
        "<draw:page xmlns:draw=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default()
    );
    let mut capacity = prefix.len();
    for (name, value) in [
        ("draw:name", page.name()),
        ("xml:id", page.xml_id()),
        ("draw:style-name", page.style_name()),
        ("draw:master-page-name", page.master_page_name()),
    ] {
        capacity = capacity
            .checked_add(serialized_attribute_len(name, value)?)
            .ok_or_else(|| Error::InvalidFormat("ODG page XML size overflow".into()))?;
    }
    capacity = capacity
        .checked_add("></draw:page>".len())
        .ok_or_else(|| Error::InvalidFormat("ODG page XML size overflow".into()))?;
    let mut xml = precharge_xml_capacity(prefix, capacity, "ODG serialized page")?;
    push_attribute(&mut xml, "draw:name", page.name())?;
    push_attribute(&mut xml, "xml:id", page.xml_id())?;
    push_attribute(&mut xml, "draw:style-name", page.style_name())?;
    push_attribute(&mut xml, "draw:master-page-name", page.master_page_name())?;
    xml.push_str("></draw:page>");
    Ok(xml)
}

fn serialize_layer(layer: &Layer) -> Result<String> {
    validate_bounded_value(layer.name(), "ODG layer name")?;
    let prefix = "<draw:layer";
    let mut capacity = prefix.len();
    let display_capacity = serialized_attribute_len("draw:display", layer.display())?;
    capacity = capacity
        .checked_add(serialized_attribute_len("draw:name", Some(layer.name()))?)
        .and_then(|size| size.checked_add(display_capacity))
        .ok_or_else(|| Error::InvalidFormat("ODG layer XML size overflow".into()))?;
    if let Some(protected) = layer.protected() {
        capacity = capacity
            .checked_add(serialized_attribute_len(
                "draw:protected",
                Some(if protected { "true" } else { "false" }),
            )?)
            .ok_or_else(|| Error::InvalidFormat("ODG layer XML size overflow".into()))?;
    }
    capacity = capacity
        .checked_add(2)
        .ok_or_else(|| Error::InvalidFormat("ODG layer XML size overflow".into()))?;
    let mut xml = precharge_xml_capacity(String::from(prefix), capacity, "ODG serialized layer")?;
    push_attribute(&mut xml, "draw:name", Some(layer.name()))?;
    push_attribute(&mut xml, "draw:display", layer.display())?;
    if let Some(protected) = layer.protected() {
        push_attribute(
            &mut xml,
            "draw:protected",
            Some(if protected { "true" } else { "false" }),
        )?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn serialize_form_control(control: &FormControl) -> Result<String> {
    validate_bounded_value(control.id(), "ODG form-control identifier")?;
    validate_xml_local_name(control.element(), "ODG form-control element")?;
    let prefix = format!(
        "<form:{} xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        control.element(),
        std::str::from_utf8(FORM).unwrap_or_default()
    );
    let mut capacity = prefix.len();
    let name_capacity = serialized_attribute_len("form:name", control.name())?;
    capacity = capacity
        .checked_add(serialized_attribute_len("form:id", Some(control.id()))?)
        .and_then(|size| size.checked_add(name_capacity))
        .ok_or_else(|| Error::InvalidFormat("ODG form-control XML size overflow".into()))?;
    for (name, value) in control.attributes() {
        validate_xml_qualified_name(name, "ODG form-control attribute")?;
        let prefix = name.split_once(':').map(|(prefix, _local)| prefix);
        if !matches!(prefix, Some("form" | "xlink")) {
            return invalid("ODG form-control attribute namespace is unsupported");
        }
        if matches!(name.as_str(), "form:id" | "form:name") {
            return invalid("ODG form-control arbitrary attributes duplicate identity");
        }
        capacity = capacity
            .checked_add(serialized_attribute_len(name, Some(value))?)
            .ok_or_else(|| Error::InvalidFormat("ODG form-control XML size overflow".into()))?;
    }
    capacity = capacity
        .checked_add(2)
        .ok_or_else(|| Error::InvalidFormat("ODG form-control XML size overflow".into()))?;
    let mut xml = precharge_xml_capacity(prefix, capacity, "ODG serialized form-control")?;
    push_attribute(&mut xml, "form:id", Some(control.id()))?;
    push_attribute(&mut xml, "form:name", control.name())?;
    for (name, value) in control.attributes() {
        push_attribute(&mut xml, name, Some(value))?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn serialize_style(style: &Style) -> Result<String> {
    validate_bounded_value(style.name(), "ODG style name")?;
    validate_bounded_value(style.family(), "ODG style family")?;
    let prefix = format!(
        "<style:style xmlns:style=\"{}\" xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:fo=\"{}\"",
        std::str::from_utf8(STYLE).unwrap_or_default(),
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(FO).unwrap_or_default()
    );
    let mut capacity = prefix.len();
    let family_capacity = serialized_attribute_len("style:family", Some(style.family()))?;
    let parent_capacity = serialized_attribute_len("style:parent-style-name", style.parent())?;
    capacity = capacity
        .checked_add(serialized_attribute_len("style:name", Some(style.name()))?)
        .and_then(|size| size.checked_add(family_capacity))
        .and_then(|size| size.checked_add(parent_capacity))
        .ok_or_else(|| Error::InvalidFormat("ODG style XML size overflow".into()))?;
    let mut owners: BTreeMap<&str, Vec<(&str, &str)>> = BTreeMap::new();
    for (path, value) in style.properties() {
        let (owner, name) = path.split_once('/').ok_or_else(|| {
            Error::InvalidFormat("ODG style property owner path is invalid".into())
        })?;
        validate_style_property_owner(owner)?;
        validate_style_property_name(name)?;
        owners
            .entry(owner)
            .or_default()
            .push((name, value.as_str()));
    }
    for (owner, properties) in &owners {
        capacity = capacity
            .checked_add(1)
            .and_then(|size| size.checked_add(owner.len()))
            .ok_or_else(|| Error::InvalidFormat("ODG style XML size overflow".into()))?;
        for (name, value) in properties {
            capacity = capacity
                .checked_add(serialized_attribute_len(name, Some(value))?)
                .ok_or_else(|| Error::InvalidFormat("ODG style XML size overflow".into()))?;
        }
        capacity = capacity
            .checked_add(2)
            .ok_or_else(|| Error::InvalidFormat("ODG style XML size overflow".into()))?;
    }
    capacity = capacity
        .checked_add(if owners.is_empty() {
            2
        } else {
            ">".len() + "</style:style>".len()
        })
        .ok_or_else(|| Error::InvalidFormat("ODG style XML size overflow".into()))?;
    let mut xml = precharge_xml_capacity(prefix, capacity, "ODG serialized style")?;
    push_attribute(&mut xml, "style:name", Some(style.name()))?;
    push_attribute(&mut xml, "style:family", Some(style.family()))?;
    push_attribute(&mut xml, "style:parent-style-name", style.parent())?;
    if owners.is_empty() {
        xml.push_str("/>");
        return Ok(xml);
    }
    xml.push('>');
    for (owner, properties) in owners {
        xml.push('<');
        xml.push_str(owner);
        for (name, value) in properties {
            push_attribute(&mut xml, name, Some(value))?;
        }
        xml.push_str("/>");
    }
    xml.push_str("</style:style>");
    Ok(xml)
}

fn serialize_style_resource(resource: &StyleResource) -> Result<String> {
    validate_bounded_value(resource.name(), "ODG named style-resource name")?;
    let prefix = format!(
        "<draw:{} xmlns:draw=\"{}\" xmlns:xlink=\"{}\" xmlns:svg=\"{}\"",
        resource.kind().element(),
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(XLINK).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
    );
    let mut capacity = prefix
        .len()
        .checked_add(serialized_attribute_len(
            "draw:name",
            Some(resource.name()),
        )?)
        .ok_or_else(|| Error::InvalidFormat("ODG style-resource XML size overflow".into()))?;
    for (name, value) in resource.attributes() {
        validate_xml_qualified_name(name, "ODG named style-resource attribute")?;
        if name == "draw:name" {
            return invalid("ODG named style-resource attributes duplicate identity");
        }
        let prefix = name.split_once(':').map(|(prefix, _local)| prefix);
        if !matches!(prefix, Some("draw" | "svg" | "xlink")) {
            return invalid("ODG named style-resource attribute namespace is unsupported");
        }
        capacity = capacity
            .checked_add(serialized_attribute_len(name, Some(value))?)
            .ok_or_else(|| Error::InvalidFormat("ODG style-resource XML size overflow".into()))?;
    }
    capacity = capacity
        .checked_add(2)
        .ok_or_else(|| Error::InvalidFormat("ODG style-resource XML size overflow".into()))?;
    let mut xml = precharge_xml_capacity(prefix, capacity, "ODG serialized style-resource")?;
    push_attribute(&mut xml, "draw:name", Some(resource.name()))?;
    for (name, value) in resource.attributes() {
        push_attribute(&mut xml, name, Some(value))?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn validate_style_property_owner(value: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid("ODG style property owner requires a qualified name");
    };
    if prefix != "style" || !local.ends_with("-properties") {
        return invalid("ODG style property owner is unsupported");
    }
    validate_xml_local_name(local, "ODG style property owner")
}

fn validate_style_property_name(value: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid("ODG style property requires a qualified name");
    };
    if !matches!(prefix, "style" | "draw" | "svg" | "fo") {
        return invalid("ODG style property uses an unsupported namespace prefix");
    }
    validate_xml_local_name(local, "ODG style property")
}

fn validate_xml_local_name(value: &str, context: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 256
        || !value.bytes().enumerate().all(|(index, byte)| {
            byte.is_ascii_alphanumeric() || byte == b'_' || index > 0 && byte == b'-'
        })
    {
        return invalid(format!("{context} is invalid"));
    }
    Ok(())
}

fn validate_xml_qualified_name(value: &str, context: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid(format!("{context} requires a qualified name"));
    };
    if !matches!(prefix, "form" | "xlink" | "draw" | "svg") {
        return invalid(format!("{context} uses an unsupported namespace prefix"));
    }
    validate_xml_local_name(local, context)
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let extra = match character {
            '&' | '<' | '>' | '\'' | '"' => match character {
                '&' => 4,
                '<' | '>' => 3,
                '\'' | '"' => 5,
                _ => 0,
            },
            _ => 0,
        };
        total
            .checked_add(character.len_utf8())
            .and_then(|size| size.checked_add(extra))
            .ok_or_else(|| Error::InvalidFormat("ODG XML escaped size overflow".into()))
    })
}

fn push_escaped_xml(output: &mut String, value: &str, resource: &'static str) -> Result<()> {
    let escaped_len = escaped_xml_len(value)?;
    let required = output
        .len()
        .checked_add(escaped_len)
        .ok_or_else(|| Error::InvalidFormat("ODG escaped XML size overflow".into()))?;
    if required > output.capacity() {
        output
            .try_reserve_exact(required - output.len())
            .map_err(|allocation_error| Error::Allocation {
                resource,
                source: allocation_error,
            })?;
    }
    let mut start = 0usize;
    for (index, character) in value.char_indices() {
        let replacement = match character {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '\'' => "&apos;",
            '"' => "&quot;",
            _ => continue,
        };
        output.push_str(&value[start..index]);
        output.push_str(replacement);
        start = index + character.len_utf8();
    }
    output.push_str(&value[start..]);
    Ok(())
}

fn escaped_xml_string(value: &str, maximum: usize, resource: &'static str) -> Result<String> {
    let escaped_len = escaped_xml_len(value)?;
    if escaped_len > maximum {
        return invalid(format!("{resource} exceeds the output limit"));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(escaped_len)
        .map_err(|allocation_error| Error::Allocation {
            resource,
            source: allocation_error,
        })?;
    push_escaped_xml(&mut output, value, resource)?;
    Ok(output)
}

fn serialized_attribute_len(name: &str, value: Option<&str>) -> Result<usize> {
    let Some(value) = value else {
        return Ok(0);
    };
    validate_bounded_value(value, "ODG XML attribute value")?;
    serialized_attribute_len_allow_empty(name, Some(value))
}

fn serialized_attribute_len_allow_empty(name: &str, value: Option<&str>) -> Result<usize> {
    let Some(value) = value else {
        return Ok(0);
    };
    if value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return invalid("ODG XML attribute value");
    }
    let escaped = escaped_xml_len(value)?;
    1usize
        .checked_add(name.len())
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(escaped))
        .and_then(|size| size.checked_add(1))
        .ok_or_else(|| Error::InvalidFormat("ODG XML attribute size overflow".into()))
}

fn validate_xml_source_size(source: &str, owner: &str) -> Result<()> {
    if source.len() > MAX_OUTPUT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{owner} exceeds the output limit"
        )));
    }
    Ok(())
}

fn clone_bounded_string(source: &str, maximum: usize, resource: &'static str) -> Result<String> {
    if source.len() > maximum {
        return Err(Error::InvalidFormat(format!(
            "{resource} exceeds the byte limit"
        )));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| Error::Allocation {
            resource,
            source: allocation_error,
        })?;
    output.push_str(source);
    Ok(output)
}

fn own_bounded_cow<'a>(
    value: Cow<'a, str>,
    maximum: usize,
    resource: &'static str,
) -> Result<String> {
    match value {
        Cow::Borrowed(value) => clone_bounded_string(value, maximum, resource),
        Cow::Owned(value) => {
            if value.len() > maximum {
                return Err(Error::InvalidFormat(format!(
                    "{resource} exceeds the byte limit"
                )));
            }
            Ok(value)
        },
    }
}

fn precharge_xml_capacity(
    mut xml: String,
    capacity: usize,
    resource: &'static str,
) -> Result<String> {
    if capacity > MAX_OUTPUT_BYTES || capacity < xml.len() {
        return invalid("ODG serialized XML exceeds the output limit");
    }
    xml.try_reserve_exact(capacity - xml.len())
        .map_err(|allocation_error| Error::Allocation {
            resource,
            source: allocation_error,
        })?;
    Ok(xml)
}

fn validate_text_content(value: &str, owner: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES
        || value.contains('\0')
        || value
            .chars()
            .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return invalid(owner);
    }
    Ok(())
}

fn shape_serialized_capacity(shape: &Shape, element: &str) -> Result<usize> {
    let (prefix, local) = element.split_once(':').unwrap_or(("draw", element));
    let mut prefix_xml = format!(
        "<{prefix}:{local} xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(TEXT).unwrap_or_default()
    );
    if prefix == "dr3d" {
        prefix_xml.push_str(" xmlns:dr3d=\"");
        prefix_xml.push_str(std::str::from_utf8(DR3D).unwrap_or_default());
        prefix_xml.push('"');
    }
    if shape.control_reference().is_some() && shape.kind() != ShapeKind::Control {
        return invalid("ODG draw:control is only supported on detached control shapes");
    }
    for value in [
        shape.name(),
        shape.layer(),
        shape.control_reference(),
        shape.style_name(),
        shape.text_style_name(),
        shape.x(),
        shape.y(),
        shape.width(),
        shape.height(),
    ]
    .into_iter()
    .flatten()
    {
        validate_bounded_value(value, "ODG XML attribute value")?;
    }
    for value in [shape.x(), shape.y(), shape.width(), shape.height()]
        .into_iter()
        .flatten()
    {
        if !is_odf_length(value) {
            return invalid("ODG shape coordinate or size is not an ODF length");
        }
    }
    if let Some(transform) = shape.transform() {
        validate_advanced_geometry_value(transform, "ODG transform")?;
    }
    if shape.points().is_some()
        && (!matches!(shape.kind(), ShapeKind::Polygon | ShapeKind::Polyline)
            || shape.view_box().is_none())
    {
        return invalid("ODG points require a polygon/polyline with a paired view box");
    }
    if matches!(shape.kind(), ShapeKind::Polygon | ShapeKind::Polyline)
        && shape.view_box().is_some()
        && shape.points().is_none()
    {
        return invalid("ODG polygon/polyline view box requires paired points");
    }
    if let Some(points) = shape.points()
        && !is_points(points)
    {
        return invalid("ODG polygon points are not an ODF point list");
    }
    if let Some(view_box) = shape.view_box()
        && !is_integer_list(view_box, 4)
    {
        return invalid("ODG polygon view box is not four integers");
    }
    let line_geometry = shape.line_geometry();
    validate_line_geometry(shape.kind(), &line_geometry, true)?;
    for endpoint in line_geometry.iter().flatten().copied() {
        if !is_odf_length(endpoint) {
            return invalid("ODG line endpoint is not an ODF length");
        }
    }
    if shape.path_data().is_some() && shape.kind() != ShapeKind::Path {
        return invalid("ODG svg:d is only supported on detached path shapes");
    }
    if let Some(path_data) = shape.path_data() {
        validate_path_data(path_data)?;
    }
    if let Some(title) = shape.title() {
        validate_text_content(title, "ODG shape title is invalid")?;
    }
    if let Some(description) = shape.description() {
        validate_text_content(description, "ODG shape description is invalid")?;
    }
    validate_text_content(shape.text(), "ODG shape text is invalid")?;

    let mut capacity = prefix_xml.len();
    for (name, value) in [
        ("draw:name", shape.name()),
        ("draw:layer", shape.layer()),
        ("draw:control", shape.control_reference()),
        ("draw:style-name", shape.style_name()),
        ("draw:text-style-name", shape.text_style_name()),
        (
            "draw:z-index",
            shape.z_index().map(|value| value.to_string()).as_deref(),
        ),
        ("svg:x", shape.x()),
        ("svg:y", shape.y()),
        ("svg:width", shape.width()),
        ("svg:height", shape.height()),
        ("draw:transform", shape.transform()),
        ("svg:viewBox", shape.view_box()),
        ("draw:points", shape.points()),
        ("svg:x1", line_geometry[0]),
        ("svg:y1", line_geometry[1]),
        ("svg:x2", line_geometry[2]),
        ("svg:y2", line_geometry[3]),
        ("svg:d", shape.path_data()),
    ] {
        capacity = capacity
            .checked_add(serialized_attribute_len(name, value)?)
            .ok_or_else(|| Error::InvalidFormat("ODG shape size overflow".into()))?;
    }
    if shape.title().is_some() || shape.description().is_some() || !shape.text().is_empty() {
        capacity = capacity
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODG shape size overflow".into()))?;
        for (open, close, value) in [
            ("<svg:title>", "</svg:title>", shape.title()),
            ("<svg:desc>", "</svg:desc>", shape.description()),
            (
                "<text:p>",
                "</text:p>",
                (!shape.text().is_empty()).then_some(shape.text()),
            ),
        ] {
            if let Some(value) = value {
                let escaped = escaped_xml_len(value)?;
                capacity = capacity
                    .checked_add(open.len())
                    .and_then(|size| size.checked_add(close.len()))
                    .and_then(|size| size.checked_add(escaped))
                    .ok_or_else(|| Error::InvalidFormat("ODG shape text size overflow".into()))?;
            }
        }
        capacity = capacity
            .checked_add(4 + prefix.len() + local.len())
            .ok_or_else(|| Error::InvalidFormat("ODG shape size overflow".into()))?;
    } else {
        capacity = capacity
            .checked_add(2)
            .ok_or_else(|| Error::InvalidFormat("ODG shape size overflow".into()))?;
    }
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG serialized shape exceeds the output limit");
    }
    Ok(capacity)
}

fn serialize_shape(shape: &Shape) -> Result<String> {
    let element = shape.kind().element_name();
    let (prefix, local) = element.split_once(':').unwrap_or(("draw", element));
    let capacity = shape_serialized_capacity(shape, element)?;
    let mut xml = format!(
        "<{prefix}:{local} xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(TEXT).unwrap_or_default()
    );
    xml.try_reserve_exact(capacity.saturating_sub(xml.len()))
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG serialized shape",
            source: allocation_error,
        })?;
    if prefix == "dr3d" {
        xml.push_str(" xmlns:dr3d=\"");
        xml.push_str(std::str::from_utf8(DR3D).unwrap_or_default());
        xml.push('"');
    }
    push_attribute(&mut xml, "draw:name", shape.name())?;
    push_attribute(&mut xml, "draw:layer", shape.layer())?;
    if shape.control_reference().is_some() && shape.kind() != ShapeKind::Control {
        return invalid("ODG draw:control is only supported on detached control shapes");
    }
    push_attribute(&mut xml, "draw:control", shape.control_reference())?;
    push_attribute(&mut xml, "draw:style-name", shape.style_name())?;
    push_attribute(&mut xml, "draw:text-style-name", shape.text_style_name())?;
    if let Some(z_index) = shape.z_index() {
        push_attribute(&mut xml, "draw:z-index", Some(&z_index.to_string()))?;
    }
    push_attribute(&mut xml, "svg:x", shape.x())?;
    push_attribute(&mut xml, "svg:y", shape.y())?;
    push_attribute(&mut xml, "svg:width", shape.width())?;
    push_attribute(&mut xml, "svg:height", shape.height())?;
    if let Some(transform) = shape.transform() {
        validate_advanced_geometry_value(transform, "ODG transform")?;
    }
    push_attribute(&mut xml, "draw:transform", shape.transform())?;
    if shape.points().is_some()
        && (!matches!(shape.kind(), ShapeKind::Polygon | ShapeKind::Polyline)
            || shape.view_box().is_none())
    {
        return invalid("ODG points require a polygon/polyline with a paired view box");
    }
    if matches!(shape.kind(), ShapeKind::Polygon | ShapeKind::Polyline)
        && shape.view_box().is_some()
        && shape.points().is_none()
    {
        return invalid("ODG polygon/polyline view box requires paired points");
    }
    if let Some(points) = shape.points() {
        validate_advanced_geometry_value(points, "ODG polygon points")?;
    }
    if let Some(view_box) = shape.view_box() {
        validate_advanced_geometry_value(view_box, "ODG polygon view box")?;
    }
    push_attribute(&mut xml, "svg:viewBox", shape.view_box())?;
    push_attribute(&mut xml, "draw:points", shape.points())?;
    let line_geometry = shape.line_geometry();
    validate_line_geometry(shape.kind(), &line_geometry, true)?;
    for endpoint in line_geometry.iter().flatten().copied() {
        validate_advanced_geometry_value(endpoint, "ODG line endpoint")?;
    }
    push_attribute(&mut xml, "svg:x1", line_geometry[0])?;
    push_attribute(&mut xml, "svg:y1", line_geometry[1])?;
    push_attribute(&mut xml, "svg:x2", line_geometry[2])?;
    push_attribute(&mut xml, "svg:y2", line_geometry[3])?;
    if shape.path_data().is_some() && shape.kind() != ShapeKind::Path {
        return invalid("ODG svg:d is only supported on detached path shapes");
    }
    if let Some(path_data) = shape.path_data() {
        validate_path_data(path_data)?;
    }
    push_attribute(&mut xml, "svg:d", shape.path_data())?;
    if shape.title().is_none() && shape.description().is_none() && shape.text().is_empty() {
        xml.push_str("/>");
        return Ok(xml);
    }
    xml.push('>');
    if let Some(title) = shape.title() {
        xml.push_str("<svg:title>");
        push_escaped_xml(&mut xml, title, "ODG serialized shape")?;
        xml.push_str("</svg:title>");
    }
    if let Some(description) = shape.description() {
        xml.push_str("<svg:desc>");
        push_escaped_xml(&mut xml, description, "ODG serialized shape")?;
        xml.push_str("</svg:desc>");
    }
    if !shape.text().is_empty() {
        xml.push_str("<text:p>");
        push_escaped_xml(&mut xml, shape.text(), "ODG serialized shape")?;
        xml.push_str("</text:p>");
    }
    xml.push_str("</");
    xml.push_str(prefix);
    xml.push(':');
    xml.push_str(local);
    xml.push('>');
    Ok(xml)
}

fn push_attribute(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(attribute_value) = value else {
        return Ok(());
    };
    validate_bounded_value(attribute_value, "ODG XML attribute value")?;
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    push_escaped_xml(output, attribute_value, "ODG XML attribute")?;
    output.push('"');
    Ok(())
}

fn validate_bounded_value(value: &str, owner: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} is invalid")));
    }
    Ok(())
}

fn validate_raw_xml_value(value: &[u8], owner: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!("{owner} exceeds the limit")));
    }
    Ok(())
}

fn validate_decoded_xml_value(value: &str, owner: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} exceeds the limit")));
    }
    Ok(())
}

fn validate_geometry(values: &[String; 4]) -> Result<()> {
    for value in values {
        validate_bounded_value(value, "ODG geometry value")?;
        if !is_odf_length(value) {
            return invalid("ODG geometry value is not an ODF length");
        }
    }
    Ok(())
}

fn validate_path_data(value: &str) -> Result<()> {
    validate_bounded_value(value, "ODG path data")?;
    if value.chars().any(char::is_control) {
        return invalid("ODG path data contains a control character");
    }
    Ok(())
}

fn validate_advanced_geometry_value(value: &str, owner: &str) -> Result<()> {
    validate_bounded_value(value, owner)?;
    if value.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(format!(
            "{owner} contains a control character"
        )));
    }
    Ok(())
}

fn validate_shape_layer(page: &Page, global_layers: &[Layer], shape: &Shape) -> Result<()> {
    let Some(layer) = shape.layer() else {
        return Ok(());
    };
    let visible = if page.has_layer_set() {
        page.layers()
    } else {
        global_layers
    };
    if !visible.iter().any(|value| value.name() == layer) {
        return invalid("ODG inserted shape references an undeclared layer");
    }
    Ok(())
}

fn resolve_transfer_layer(snapshot: &Package, page: &Page, name: &str) -> Result<Layer> {
    let visible = if page.has_layer_set() {
        page.layers()
    } else {
        snapshot.layers()
    };
    visible
        .iter()
        .find(|layer| layer.name() == name)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODG transfer shape references undeclared layer '{name}'"
            ))
        })
}

fn transfer_xml_references(xml: &str, href: &str) -> bool {
    let escaped = quick_xml::escape::escape(href);
    xml.contains(&format!("xlink:href=\"{escaped}\""))
        || xml.contains(&format!("xlink:href='{escaped}'"))
}

/// Collects root and inherited namespace declarations available to fragments
/// copied out of the two XML parts.  A source package may declare producer
/// namespaces such as `loext` only on its document root; those declarations
/// must travel with an extracted shape or the destination document would have
/// an unbound prefix.  Conflicting aliases are withheld and cause an explicit
/// refusal below.
fn transfer_source_namespaces(snapshot: &Package) -> Result<BTreeMap<String, String>> {
    let mut namespaces = BTreeMap::new();
    let mut conflicts = BTreeSet::new();
    for xml in [Some(snapshot.content_xml()), snapshot.styles_xml()] {
        let Some(xml) = xml else {
            continue;
        };
        for (prefix, uri) in transfer_root_namespaces(xml)? {
            if let Some(previous) = namespaces.get(&prefix)
                && previous != &uri
            {
                conflicts.insert(prefix.clone());
            } else {
                namespaces.insert(prefix, uri);
            }
        }
    }
    for prefix in conflicts {
        namespaces.remove(&prefix);
    }
    Ok(namespaces)
}

/// Returns namespace prefixes and declarations found in one self-contained
/// transfer fragment.  Declarations are collected at every depth so a
/// declaration owned by a nested opaque child remains valid in its original
/// scope; `close_transfer_fragment_namespaces` only adds declarations that are
/// absent from the fragment altogether.
fn transfer_fragment_namespaces(
    xml: &str,
) -> Result<(BTreeSet<String>, BTreeMap<String, String>, BTreeSet<String>)> {
    validate_xml_source_size(xml, "ODG transfer fragment")?;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut declarations = BTreeMap::<String, String>::new();
    let mut used = BTreeSet::<String>::new();
    let mut sensitive_aliases = BTreeSet::<String>::new();
    let mut root_seen = false;
    loop {
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer fragment XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transfer fragment event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                record_transfer_qname(
                    element.name().as_ref(),
                    &mut used,
                    &mut sensitive_aliases,
                    !root_seen,
                )?;
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG transfer namespace: {error}"))
                    })?;
                    let name = attribute.key.as_ref();
                    validate_raw_xml_value(name, "ODG transfer namespace name")?;
                    if name == b"xmlns" {
                        continue;
                    }
                    if let Some(prefix) = name.strip_prefix(b"xmlns:") {
                        validate_raw_xml_value(prefix, "ODG transfer namespace prefix")?;
                        let prefix = std::str::from_utf8(prefix).map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODG transfer namespace prefix: {error}"
                            ))
                        })?;
                        validate_raw_xml_value(
                            attribute.value.as_ref(),
                            "ODG transfer namespace URI",
                        )?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid ODG transfer namespace URI: {error}"
                                ))
                            })?;
                        validate_decoded_xml_value(value.as_ref(), "ODG transfer namespace URI")?;
                        let value =
                            own_bounded_cow(value, MAX_TEXT_BYTES, "ODG transfer namespace URI")?;
                        let prefix = clone_bounded_string(
                            prefix,
                            MAX_TEXT_BYTES,
                            "ODG transfer namespace prefix",
                        )?;
                        declarations.insert(prefix, value);
                    } else {
                        record_transfer_qname(name, &mut used, &mut sensitive_aliases, false)?;
                    }
                }
                root_seen = true;
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transfer fragments");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
    if !root_seen {
        return invalid("ODG transfer fragment has no root element");
    }
    Ok((used, declarations, sensitive_aliases))
}

/// Checks that every namespace used by a transfer fragment can be reproduced
/// after the fragment is moved under a different content root.  Rewriting is
/// intentionally limited to canonical ODF prefixes: an alias for a known
/// dependency namespace would otherwise make resource/style closure checks
/// silently miss a reference.  Such fragments are refused before any
/// transaction state is changed.
fn validate_transfer_fragment_namespaces(xml: &str) -> Result<()> {
    let (used, declarations, sensitive_aliases) = transfer_fragment_namespaces(xml)?;
    for prefix in used {
        if prefix == "xml" {
            continue;
        }
        let expected = transfer_namespace_uri(&prefix);
        let declared = declarations.get(&prefix).map(String::as_str);
        match (expected, declared) {
            (Some(expected), Some(declared)) if declared != expected => {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG transfer fragment binds canonical prefix to another namespace",
                )));
            },
            (Some(_), _) => {},
            (None, Some(uri)) if transfer_namespace_uri_bytes(uri).is_some() => {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG transfer fragment aliases known namespace with prefix",
                )));
            },
            (None, Some(_)) => {},
            (None, None) if sensitive_aliases.contains(&prefix) => {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG transfer fragment cannot prove dependency namespace for prefix",
                )));
            },
            (None, None) => {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odg.package.snapshot",
                    "ODG transfer fragment cannot prove namespace for prefix",
                )));
            },
        }
    }
    Ok(())
}

/// Adds source-root declarations needed by an extracted fragment.  Unknown
/// producer prefixes are copied verbatim when their source-root URI is
/// unambiguous; otherwise the transfer is refused instead of publishing XML
/// whose namespace identity depends on the destination root.
fn close_transfer_fragment_namespaces(
    xml: &str,
    available: &BTreeMap<String, String>,
) -> Result<String> {
    let (used, declarations, _sensitive_aliases) = transfer_fragment_namespaces(xml)?;
    let mut missing = BTreeMap::new();
    for prefix in used {
        if prefix == "xml"
            || transfer_namespace_uri(&prefix).is_some()
            || declarations.contains_key(&prefix)
        {
            continue;
        }
        let uri = available.get(&prefix).ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG transfer fragment cannot prove namespace closure for prefix",
            ))
        })?;
        missing.insert(prefix, uri.clone());
    }
    ensure_transfer_namespace_values(xml, &missing)
}

fn transfer_namespace_closure_bytes(
    xml: &str,
    available: &BTreeMap<String, String>,
) -> Result<usize> {
    let (used, declarations, _sensitive_aliases) = transfer_fragment_namespaces(xml)?;
    used.into_iter().try_fold(0usize, |total, prefix| {
        if prefix == "xml"
            || transfer_namespace_uri(&prefix).is_some()
            || declarations.contains_key(&prefix)
        {
            return Ok(total);
        }
        let uri = available.get(&prefix).ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odg.package.snapshot",
                "ODG transfer fragment cannot prove namespace closure for prefix",
            ))
        })?;
        // `ensure_transfer_namespace_values` writes exactly:
        // ` xmlns:{prefix}="{uri}"`.
        let escaped_uri_len = escaped_xml_len(uri)?;
        let declaration_bytes = 1usize
            .checked_add("xmlns:".len())
            .and_then(|size| size.checked_add(prefix.len()))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(escaped_uri_len))
            .and_then(|size| size.checked_add(1))
            .ok_or_else(|| Error::InvalidFormat("ODG namespace closure size overflow".into()))?;
        total
            .checked_add(declaration_bytes)
            .ok_or_else(|| Error::InvalidFormat("ODG namespace closure size overflow".into()))
    })
}

fn transfer_root_namespaces(xml: &str) -> Result<BTreeMap<String, String>> {
    validate_xml_source_size(xml, "ODG transfer root XML")?;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer root XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transfer root event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let mut declarations = BTreeMap::new();
                for raw_attribute in element.attributes() {
                    let attribute = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid ODG transfer root namespace declaration: {error}"
                        ))
                    })?;
                    let key = attribute.key.as_ref();
                    validate_raw_xml_value(key, "ODG transfer root namespace name")?;
                    let prefix = if key == b"xmlns" {
                        ""
                    } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                        validate_raw_xml_value(prefix, "ODG transfer root namespace prefix")?;
                        std::str::from_utf8(prefix).map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODG transfer root namespace prefix: {error}"
                            ))
                        })?
                    } else {
                        continue;
                    };
                    validate_raw_xml_value(
                        attribute.value.as_ref(),
                        "ODG transfer root namespace URI",
                    )?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODG transfer root namespace URI: {error}"
                            ))
                        })?;
                    validate_decoded_xml_value(value.as_ref(), "ODG transfer root namespace URI")?;
                    let value =
                        own_bounded_cow(value, MAX_TEXT_BYTES, "ODG transfer root namespace URI")?;
                    let prefix = clone_bounded_string(
                        prefix,
                        MAX_TEXT_BYTES,
                        "ODG transfer root namespace prefix",
                    )?;
                    if declarations.insert(prefix, value).is_some() {
                        return invalid("ODG transfer root namespace declaration is duplicated");
                    }
                }
                return Ok(declarations);
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transfer roots");
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::CData(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::End(_) | Event::Text(_) | Event::CData(_) | Event::Eof => {
                return invalid("ODG transfer root start tag is missing");
            },
        }
    }
}

fn record_transfer_qname(
    name: &[u8],
    used: &mut BTreeSet<String>,
    sensitive_aliases: &mut BTreeSet<String>,
    root: bool,
) -> Result<()> {
    validate_raw_xml_value(name, "ODG transfer qualified name")?;
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return Ok(());
    };
    validate_raw_xml_value(&name[..separator], "ODG transfer qualified-name prefix")?;
    let prefix = std::str::from_utf8(&name[..separator]).map_err(|error| {
        Error::InvalidFormat(format!("invalid ODG transfer qualified name: {error}"))
    })?;
    if prefix.is_empty() {
        return invalid("ODG transfer qualified name has an empty prefix");
    }
    used.insert(clone_bounded_string(
        prefix,
        MAX_TEXT_BYTES,
        "ODG transfer qualified-name prefix",
    )?);
    if root || transfer_sensitive_local_name(&name[separator + 1..]) {
        sensitive_aliases.insert(clone_bounded_string(
            prefix,
            MAX_TEXT_BYTES,
            "ODG transfer qualified-name prefix",
        )?);
    }
    Ok(())
}

fn transfer_sensitive_local_name(name: &[u8]) -> bool {
    matches!(
        name,
        b"href"
            | b"name"
            | b"style-name"
            | b"text-style-name"
            | b"parent-style-name"
            | b"fill-gradient-name"
            | b"fill-hatch-name"
            | b"fill-image-name"
            | b"marker-start"
            | b"marker-end"
            | b"stroke-dash"
            | b"opacity"
            | b"control"
            | b"id"
    )
}

fn transfer_namespace_uri(prefix: &str) -> Option<&'static str> {
    match prefix {
        "office" => std::str::from_utf8(OFFICE).ok(),
        "draw" => std::str::from_utf8(DRAW).ok(),
        "dr3d" => std::str::from_utf8(DR3D).ok(),
        "svg" => std::str::from_utf8(SVG).ok(),
        "text" => std::str::from_utf8(TEXT).ok(),
        "table" => std::str::from_utf8(TABLE).ok(),
        "style" => std::str::from_utf8(STYLE).ok(),
        "form" => std::str::from_utf8(FORM).ok(),
        "fo" => std::str::from_utf8(FO).ok(),
        "presentation" => std::str::from_utf8(PRESENTATION).ok(),
        "smil" => std::str::from_utf8(SMIL).ok(),
        "xlink" => std::str::from_utf8(XLINK).ok(),
        "xml" => std::str::from_utf8(XML).ok(),
        _ => None,
    }
}

fn transfer_namespace_uri_bytes(uri: &str) -> Option<&'static [u8]> {
    [
        OFFICE,
        DRAW,
        DR3D,
        SVG,
        TEXT,
        TABLE,
        STYLE,
        FORM,
        FO,
        PRESENTATION,
        SMIL,
        XLINK,
        XML,
    ]
    .into_iter()
    .find(|candidate| std::str::from_utf8(candidate).ok() == Some(uri))
}

fn declares_style(snapshot: &Package, name: &str) -> Result<bool> {
    if declares_style_xml(snapshot.content_xml(), name)? {
        return Ok(true);
    }
    snapshot
        .styles_xml()
        .map(|xml| declares_style_xml(xml, name))
        .transpose()
        .map(Option::unwrap_or_default)
}

fn declares_style_xml(xml: &str, name: &str) -> Result<bool> {
    parse_style_definitions(xml).map(|definitions| {
        definitions
            .iter()
            .any(|definition| definition.style.name() == name)
    })
}

fn find_style_definition(snapshot: &Package, name: &str) -> Result<Option<ParsedStyleDefinition>> {
    let mut matches = parse_style_definitions(snapshot.content_xml())?
        .into_iter()
        .filter(|definition| definition.style.name() == name)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return invalid("ODG style definition is ambiguous");
    }
    if let Some(definition) = matches.pop() {
        return Ok(Some(definition));
    }
    if let Some(styles) = snapshot.styles_xml() {
        let mut style_matches = parse_style_definitions(styles)?
            .into_iter()
            .filter(|definition| definition.style.name() == name)
            .collect::<Vec<_>>();
        if style_matches.len() > 1 {
            return invalid("ODG style definition is ambiguous");
        }
        return Ok(style_matches.pop());
    }
    Ok(None)
}

fn style_parent_name(xml: &str) -> Result<Option<String>> {
    validate_raw_xml_events(xml, "ODG transferred style event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transferred style XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transferred style event")?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if classify(&namespace) == NamespaceKind::Style
                    && element.local_name().as_ref() == b"style" =>
            {
                return attribute(&reader, &element, STYLE, b"parent-style-name");
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transferred styles");
            },
            Event::Eof => return Ok(None),
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::Start(_)
            | Event::Empty(_) => {},
        }
    }
}

fn xml_has_attribute(xml: &str, namespace: &[u8], local: &[u8], value: &str) -> Result<bool> {
    validate_raw_xml_events(xml, "ODG dependency event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (_resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG dependency XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG dependency event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if attribute(&reader, &element, namespace, local)?.as_deref() == Some(value) {
                    return Ok(true);
                }
            },
            // Character references and the predefined entities in text
            // (LibreOffice writes `&apos;`) are inert; custom entities are
            // refused by `reference_value`.
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG dependencies");
            },
            Event::Eof => return Ok(false),
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
}

fn resolve_page_position(pages: &[Page], selector: crate::page::Selector<'_>) -> Result<usize> {
    match selector {
        crate::page::Selector::Position(position) => pages
            .get(position.get())
            .map(|_| position.get())
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into())),
        crate::page::Selector::Name(name) => {
            let mut matches = pages
                .iter()
                .enumerate()
                .filter(|(_, page)| page.name() == Some(name.as_ref()));
            let selected = matches
                .next()
                .ok_or_else(|| Error::InvalidFormat("ODG page selector did not match".into()))?;
            if matches.next().is_some() {
                return invalid("ODG page name selector is ambiguous");
            }
            Ok(selected.0)
        },
    }
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || !media_type.contains('/')
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return invalid("ODG resource media type is invalid");
    }
    Ok(())
}

fn validate_resource_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.len() > 4_096
        || path.starts_with('/')
        || path.ends_with('/')
        || path.contains('\\')
        || path.contains(':')
        || path
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
        || path.bytes().any(|byte| byte.is_ascii_control())
    {
        return invalid("ODG resource path is unsafe");
    }
    Ok(())
}

fn unique_collision_name(base: &str, bytes: &[u8], reserved: &[String]) -> String {
    let fingerprint = DiagnosticFingerprint::of(bytes).as_hex();
    let suffix = &fingerprint[..12];
    let candidate = format!("{base}_litchi_{suffix}");
    if !reserved.iter().any(|name| name == &candidate) {
        return candidate;
    }
    for ordinal in 2usize..=MAX_TRANSFER_RESOURCES {
        let numbered_candidate = format!("{base}_litchi_{suffix}_{ordinal}");
        if !reserved.iter().any(|name| name == &numbered_candidate) {
            return numbered_candidate;
        }
    }
    format!("{base}_litchi_{fingerprint}")
}

fn unique_resource_path(
    source: &Package,
    staged: &[ResourceEdit],
    path: &str,
    bytes: &[u8],
) -> Result<String> {
    let (stem, extension) = path
        .rsplit_once('.')
        .map_or((path, ""), |(stem, extension)| (stem, extension));
    let fingerprint = DiagnosticFingerprint::of(bytes).as_hex();
    for width in [12usize, 24, 64] {
        let suffix = &fingerprint[..width];
        let candidate = if extension.is_empty() {
            format!("{stem}_litchi_{suffix}")
        } else {
            format!("{stem}_litchi_{suffix}.{extension}")
        };
        validate_resource_path(&candidate)?;
        if !source.files()?.iter().any(|value| value == &candidate)
            && !staged.iter().any(|edit| edit.path == candidate)
        {
            return Ok(candidate);
        }
    }
    invalid("ODG transferred resource collision could not be remapped")
}

fn rewrite_qualified_attribute_values(
    xml: &str,
    qualified_names: &[&[u8]],
    before: &str,
    after: &str,
) -> Result<String> {
    if before == after {
        return clone_bounded_string(xml, MAX_OUTPUT_BYTES, "ODG transfer fragment");
    }
    validate_xml_source_size(xml, "ODG transfer fragment")?;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut spans = Vec::new();
    loop {
        let start = position_reader(&reader)?;
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer fragment XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transfer fragment rewrite event")?;
        let end = position_reader(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let tag = xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer start tag span is invalid".into())
                })?;
                for raw_attribute in element.attributes() {
                    let parsed = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG transfer attribute: {error}"))
                    })?;
                    if !qualified_names
                        .iter()
                        .any(|name| parsed.key.as_ref() == *name)
                    {
                        continue;
                    }
                    validate_raw_xml_value(parsed.value.as_ref(), "ODG transfer attribute value")?;
                    let value = parsed
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODG transfer attribute value: {error}"
                            ))
                        })?;
                    validate_decoded_xml_value(value.as_ref(), "ODG transfer attribute value")?;
                    if value == before {
                        let (value_start, value_end) =
                            attribute_value_span(tag, parsed.key.as_ref())?;
                        spans.push(start + value_start..start + value_end);
                    }
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transfer fragments");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
    let mut output = clone_bounded_string(xml, MAX_OUTPUT_BYTES, "ODG transfer fragment")?;
    spans.sort_unstable_by_key(|span| std::cmp::Reverse(span.start));
    for span in spans {
        output = replace_xml_value(&output, &span, after)?;
    }
    Ok(output)
}

fn ensure_transfer_namespaces(xml: &str, namespaces: &[(&str, &[u8])]) -> Result<String> {
    let namespaces = namespaces
        .iter()
        .map(|(prefix, uri)| {
            validate_raw_xml_value(prefix.as_bytes(), "ODG transfer namespace prefix")?;
            validate_raw_xml_value(uri, "ODG transfer namespace URI")?;
            let uri = std::str::from_utf8(uri).map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG transfer namespace: {error}"))
            })?;
            let uri = clone_bounded_string(uri, MAX_TEXT_BYTES, "ODG transfer namespace URI")?;
            let prefix =
                clone_bounded_string(prefix, MAX_TEXT_BYTES, "ODG transfer namespace prefix")?;
            Ok((prefix, uri))
        })
        .collect::<Result<BTreeMap<_, _>>>()?;
    ensure_transfer_namespace_values(xml, &namespaces)
}

fn ensure_transfer_namespace_values(
    xml: &str,
    namespaces: &BTreeMap<String, String>,
) -> Result<String> {
    validate_xml_source_size(xml, "ODG transfer fragment")?;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let (root_start, root_end) = loop {
        let start = position_reader(&reader)?;
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer fragment XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transfer namespace event")?;
        let end = position_reader(&reader)?;
        match event {
            Event::Start(_) | Event::Empty(_) => break (start, end),
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("active XML is prohibited in ODG transfer fragments");
            },
            Event::CData(_) | Event::End(_) | Event::Text(_) | Event::Eof => {
                return invalid("ODG transfer fragment start tag is missing");
            },
        }
    };
    let tag = xml
        .get(root_start..root_end)
        .ok_or_else(|| Error::InvalidFormat("ODG transfer fragment start tag is invalid".into()))?;
    let name_end = tag
        .bytes()
        .enumerate()
        .skip(1)
        .find(|(_index, byte)| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .map(|(index, _byte)| root_start + index)
        .ok_or_else(|| Error::InvalidFormat("ODG transfer fragment name is invalid".into()))?;
    let mut declaration_capacity = 0usize;
    for (prefix, namespace_uri) in namespaces {
        validate_decoded_xml_value(prefix, "ODG transfer namespace prefix")?;
        validate_decoded_xml_value(namespace_uri, "ODG transfer namespace URI")?;
        if !transfer_root_declares_namespace(xml, prefix)? {
            let escaped_uri_len = escaped_xml_len(namespace_uri)?;
            let declaration_bytes = 1usize
                .checked_add("xmlns:".len())
                .and_then(|size| size.checked_add(prefix.len()))
                .and_then(|size| size.checked_add(2))
                .and_then(|size| size.checked_add(escaped_uri_len))
                .and_then(|size| size.checked_add(1))
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer namespace declaration size overflow".into())
                })?;
            declaration_capacity = declaration_capacity
                .checked_add(declaration_bytes)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer namespace declaration size overflow".into())
                })?;
        }
    }
    let mut declarations = precharge_xml_capacity(
        String::new(),
        declaration_capacity,
        "ODG transfer namespace declarations",
    )?;
    for (prefix, namespace_uri) in namespaces {
        if !transfer_root_declares_namespace(xml, prefix)? {
            write!(declarations, " xmlns:{prefix}=\"").map_err(|error| {
                Error::InvalidFormat(format!("ODG transfer namespace write failed: {error}"))
            })?;
            push_escaped_xml(
                &mut declarations,
                namespace_uri,
                "ODG transfer namespace declarations",
            )?;
            declarations.push('"');
        }
    }
    insert_xml(xml, name_end, &declarations)
}

fn transfer_root_declares_namespace(xml: &str, prefix: &str) -> Result<bool> {
    let expected = format!("xmlns:{prefix}");
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer fragment XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG transfer namespace declaration event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                return element.attributes().try_fold(false, |found, raw| {
                    let attribute = raw.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid ODG transfer namespace declaration: {error}"
                        ))
                    })?;
                    Ok(found || attribute.key.as_ref() == expected.as_bytes())
                });
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("active XML is prohibited in ODG transfer fragments");
            },
            Event::CData(_) | Event::End(_) | Event::Text(_) | Event::Eof => {
                return invalid("ODG transfer fragment start tag is missing");
            },
        }
    }
}

fn inverse_change(change: &Change) -> Change {
    match change {
        Change::ControlReference(value) => Change::ControlReference(ControlReferenceChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Text(value) => Change::Text(TextChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Name(value) => Change::Name(NameChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Layer(value) => Change::Layer(LayerChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Geometry(value) => Change::Geometry(GeometryChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Style(value) => Change::Style(StyleChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Path(value) => Change::Path(PathChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::PageName(value) => Change::PageName(PageNameChange {
            page: value.page,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::PageStyle(value) => Change::PageStyle(PageStyleChange {
            page: value.page,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::PageTransition(value) => Change::PageTransition(Box::new(PageTransitionChange {
            page: value.page,
            before: value.after.clone(),
            after: value.before.clone(),
        })),
        Change::Structure(value) => Change::Structure(match value {
            StructureChange::PageInserted { position, name } => StructureChange::PageRemoved {
                position: *position,
                name: name.clone(),
            },
            StructureChange::PageRemoved { position, name } => StructureChange::PageInserted {
                position: *position,
                name: name.clone(),
            },
            StructureChange::LayerInserted { page, name } => StructureChange::LayerRemoved {
                page: *page,
                name: name.clone(),
            },
            StructureChange::LayerRemoved { page, name } => StructureChange::LayerInserted {
                page: *page,
                name: name.clone(),
            },
            StructureChange::ShapeInserted {
                page,
                position,
                kind,
            } => StructureChange::ShapeRemoved {
                page: *page,
                position: *position,
                kind: *kind,
            },
            StructureChange::ShapeRemoved {
                page,
                position,
                kind,
            } => StructureChange::ShapeInserted {
                page: *page,
                position: *position,
                kind: *kind,
            },
            StructureChange::PageMoved { from, to } => StructureChange::PageMoved {
                from: *to,
                to: *from,
            },
            StructureChange::ShapeMoved { page, from, to } => StructureChange::ShapeMoved {
                page: *page,
                from: *to,
                to: *from,
            },
            StructureChange::ShapeTransferred {
                page,
                from,
                to_page,
                to,
            } => StructureChange::ShapeTransferred {
                page: *to_page,
                from: *to,
                to_page: *page,
                to: *from,
            },
            StructureChange::ShapeAdvancedGeometryChanged { page, shape } => {
                StructureChange::ShapeAdvancedGeometryChanged {
                    page: *page,
                    shape: *shape,
                }
            },
            StructureChange::FormControlInserted { id } => {
                StructureChange::FormControlRemoved { id: id.clone() }
            },
            StructureChange::FormControlRemoved { id } => {
                StructureChange::FormControlInserted { id: id.clone() }
            },
            StructureChange::FormControlReplaced { id } => {
                StructureChange::FormControlReplaced { id: id.clone() }
            },
            StructureChange::StyleInserted { name } => {
                StructureChange::StyleRemoved { name: name.clone() }
            },
            StructureChange::StyleRemoved { name } => {
                StructureChange::StyleInserted { name: name.clone() }
            },
            StructureChange::StyleReplaced { name } => {
                StructureChange::StyleReplaced { name: name.clone() }
            },
            StructureChange::StyleResourceInserted { kind, name } => {
                StructureChange::StyleResourceRemoved {
                    kind: *kind,
                    name: name.clone(),
                }
            },
            StructureChange::StyleResourceRemoved { kind, name } => {
                StructureChange::StyleResourceInserted {
                    kind: *kind,
                    name: name.clone(),
                }
            },
            StructureChange::StyleResourceReplaced { kind, name } => {
                StructureChange::StyleResourceReplaced {
                    kind: *kind,
                    name: name.clone(),
                }
            },
        }),
    }
}

fn inverse_resource_change(change: &ResourceChange) -> ResourceChange {
    ResourceChange {
        resource: change.resource,
        path: change.path.clone(),
        before_media_type: change.after_media_type.clone(),
        after_media_type: change.before_media_type.clone(),
        before_size: change.after_size,
        after_size: change.before_size,
    }
}

fn frame(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: Option<String>,
    page_name: Option<String>,
) -> Result<Frame> {
    Ok(Frame {
        name,
        xml_id: attribute(reader, element, XML, b"id")?,
        title: None,
        description: None,
        anchor_type: attribute(reader, element, TEXT, b"anchor-type")?,
        x: attribute(reader, element, SVG, b"x")?,
        y: attribute(reader, element, SVG, b"y")?,
        width: attribute(reader, element, SVG, b"width")?,
        height: attribute(reader, element, SVG, b"height")?,
        end_cell_address: None,
        page_name,
        sheet_name: None,
        sheet_shape: false,
    })
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        if parsed_attribute.key.local_name().as_ref() != local {
            continue;
        }
        let (namespace, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected) && name.as_ref() == local {
            // Check the encoded attribute before quick-xml can materialize its Cow.
            validate_raw_xml_value(parsed_attribute.value.as_ref(), "ODG attribute value")?;
            let decoded = parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODG attribute value: {error}"))
                })?;
            // Keep the decoded Cow borrowed/temporary until its size is admitted.
            validate_decoded_xml_value(decoded.as_ref(), "ODG attribute value")?;
            let decoded = own_bounded_cow(decoded, MAX_TEXT_BYTES, "ODG attribute value")?;
            if value.replace(decoded).is_some() {
                return invalid("ODG element has a duplicate namespaced attribute");
            }
        }
    }
    Ok(value)
}

fn attribute_values<const N: usize>(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    requests: [(&[u8], &[u8]); N],
) -> Result<[Option<String>; N]> {
    let mut values = std::array::from_fn(|_| None);
    let mut batch_error = false;

    for raw_attribute in element.attributes() {
        let parsed_attribute = match raw_attribute {
            Ok(parsed_attribute) => parsed_attribute,
            Err(_) => {
                batch_error = true;
                break;
            },
        };
        let parsed_local = parsed_attribute.key.local_name();
        if !requests
            .iter()
            .any(|(_, wanted)| parsed_local.as_ref() == *wanted)
        {
            continue;
        }
        let (namespace, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        for (slot, (expected, wanted)) in requests.iter().enumerate() {
            if parsed_local.as_ref() != *wanted
                || !resolved_bound(&namespace, expected)
                || name.as_ref() != *wanted
            {
                continue;
            }
            if values[slot].is_some() {
                batch_error = true;
                break;
            }
            // Apply the same bounded admission as `attribute` before a value
            // is decoded or retained. Any refusal replays the ordered helper
            // below, which reports the established first error.
            if validate_raw_xml_value(parsed_attribute.value.as_ref(), "ODG attribute value")
                .is_err()
            {
                batch_error = true;
                break;
            }
            let decoded = match parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            {
                Ok(decoded) => decoded,
                Err(_) => {
                    batch_error = true;
                    break;
                },
            };
            if validate_decoded_xml_value(decoded.as_ref(), "ODG attribute value").is_err() {
                batch_error = true;
                break;
            }
            let Ok(decoded) = own_bounded_cow(decoded, MAX_TEXT_BYTES, "ODG attribute value")
            else {
                batch_error = true;
                break;
            };
            values[slot] = Some(decoded);
        }
        if batch_error {
            break;
        }
    }

    if !batch_error {
        return Ok(values);
    }

    // The fast traversal intentionally discards partial values on every
    // error. Replaying the legacy helper in request order preserves the
    // established first-error precedence even when raw attribute order differs
    // from the order in which callers consume the groups.
    drop(values);
    let mut ordered = std::array::from_fn(|_| None);
    for (slot, (expected, wanted)) in requests.iter().enumerate() {
        ordered[slot] = attribute(reader, element, expected, wanted)?;
    }
    Ok(ordered)
}

fn arbitrary_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    excluded: &[(&[u8], &[u8])],
) -> Result<BTreeMap<String, String>> {
    let mut values = BTreeMap::new();
    for raw_attribute in element.attributes() {
        let parsed = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let raw_name = parsed.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(parsed.key);
        if excluded.iter().any(|(expected, wanted)| {
            resolved_bound(&namespace, expected) && local.as_ref() == *wanted
        }) {
            continue;
        }
        validate_raw_xml_value(raw_name, "ODG arbitrary attribute name")?;
        let name = std::str::from_utf8(raw_name).map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG attribute name: {error}"))
        })?;
        let name = clone_bounded_string(name, MAX_TEXT_BYTES, "ODG arbitrary attribute name")?;
        // The raw escaped bytes bound normalization before it can allocate.
        validate_raw_xml_value(parsed.value.as_ref(), "ODG arbitrary attribute value")?;
        let value = parsed
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODG attribute value: {error}"))
            })?;
        // Admit the decoded Cow before retaining an owned attribute value.
        validate_decoded_xml_value(value.as_ref(), "ODG arbitrary attribute value")?;
        let value = own_bounded_cow(value, MAX_TEXT_BYTES, "ODG arbitrary attribute value")?;
        if values.insert(name, value).is_some() {
            return invalid("ODG element has a duplicate arbitrary attribute");
        }
    }
    Ok(values)
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
    owner: &str,
) -> Result<String> {
    attribute(reader, element, expected, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODG {owner} requires a namespaced attribute")))
}

fn optional_bool_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<bool>> {
    attribute(reader, element, expected, local)?
        .map(|value| match value.as_str() {
            "false" => Ok(false),
            "true" => Ok(true),
            _ => invalid("ODG Boolean attribute is not true or false"),
        })
        .transpose()
}

fn optional_u32_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<u32>> {
    attribute(reader, element, expected, local)?
        .map(|value| {
            value.parse::<u32>().map_err(|_error| {
                Error::InvalidFormat("ODG integer attribute is invalid".to_string())
            })
        })
        .transpose()
}

fn is_vector3d(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix('(')
        .and_then(|value| value.strip_suffix(')'))
    else {
        return false;
    };
    let mut values = inner.split(' ').filter(|value| !value.is_empty());
    let valid = (0..3).all(|_| values.next().is_some_and(is_odf_decimal));
    valid && values.next().is_none()
}

fn validate_three_dimensional_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: ShapeKind,
) -> Result<()> {
    if attribute(reader, element, DRAW, b"transform")?.is_some() {
        return invalid("ODG 3D shapes use dr3d:transform, not draw:transform");
    }
    if let Some(transform) = attribute(reader, element, DR3D, b"transform")? {
        validate_optional_lexical_value(&transform, "ODG dr3d:transform")?;
    }
    match kind {
        ShapeKind::ThreeDimensionalScene => {
            for local in [b"vrp".as_slice(), b"vpn", b"vup"] {
                if let Some(value) = attribute(reader, element, DR3D, local)?
                    && !is_vector3d(&value)
                {
                    return invalid("ODG dr3d scene vector is invalid");
                }
            }
            if let Some(value) = attribute(reader, element, DR3D, b"projection")?
                && !matches!(value.as_str(), "parallel" | "perspective")
            {
                return invalid("ODG dr3d projection is invalid");
            }
            for local in [b"distance".as_slice(), b"focal-length"] {
                if let Some(value) = attribute(reader, element, DR3D, local)?
                    && !is_odf_length(&value)
                {
                    return invalid("ODG dr3d scene distance is invalid");
                }
            }
            if let Some(value) = attribute(reader, element, DR3D, b"shade-mode")?
                && !matches!(value.as_str(), "flat" | "phong" | "gouraud" | "draft")
            {
                return invalid("ODG dr3d shade mode is invalid");
            }
            if let Some(value) = attribute(reader, element, DR3D, b"ambient-color")?
                && !is_odf_color(&value)
            {
                return invalid("ODG dr3d ambient color is invalid");
            }
            let _ = optional_bool_attribute(reader, element, DR3D, b"lighting-mode")?;
        },
        ShapeKind::ThreeDimensionalLight => {
            if let Some(value) = attribute(reader, element, DR3D, b"diffuse-color")?
                && !is_odf_color(&value)
            {
                return invalid("ODG dr3d diffuse color is invalid");
            }
            let _ = optional_bool_attribute(reader, element, DR3D, b"enabled")?;
            let _ = optional_bool_attribute(reader, element, DR3D, b"specular")?;
        },
        ShapeKind::ThreeDimensionalCube => {
            for local in [b"min-edge".as_slice(), b"max-edge"] {
                if let Some(value) = attribute(reader, element, DR3D, local)?
                    && !is_vector3d(&value)
                {
                    return invalid("ODG dr3d cube edge is invalid");
                }
            }
        },
        ShapeKind::ThreeDimensionalSphere => {
            for local in [b"center".as_slice(), b"size"] {
                if let Some(value) = attribute(reader, element, DR3D, local)?
                    && !is_vector3d(&value)
                {
                    return invalid("ODG dr3d sphere vector is invalid");
                }
            }
        },
        ShapeKind::ThreeDimensionalExtrude | ShapeKind::ThreeDimensionalRotate => {},
        _ => {},
    }
    Ok(())
}

fn validate_optional_lexical_value(value: &str, owner: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return invalid(owner);
    }
    Ok(())
}

fn is_integer_list(value: &str, expected: usize) -> bool {
    let mut values = value.split_ascii_whitespace();
    let valid = (0..expected).all(|_| values.next().is_some_and(is_odf_integer));
    valid && values.next().is_none()
}

fn is_points(value: &str) -> bool {
    let mut points = value.split(' ');
    let Some(first) = points.next() else {
        return false;
    };
    is_point(first) && points.all(is_point)
}

fn is_point(point: &str) -> bool {
    let Some((x, y)) = point.split_once(',') else {
        return false;
    };
    is_points_integer(x) && is_points_integer(y)
}

fn is_points_integer(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
}

fn is_odf_integer(value: &str) -> bool {
    let value = value
        .strip_prefix('-')
        .or_else(|| value.strip_prefix('+'))
        .unwrap_or(value);
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
}

fn is_odf_decimal(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    let Some((whole, fraction)) = value.split_once('.') else {
        return !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit());
    };
    if whole.is_empty() {
        !fraction.is_empty() && fraction.bytes().all(|byte| byte.is_ascii_digit())
    } else {
        whole.bytes().all(|byte| byte.is_ascii_digit())
            && fraction.bytes().all(|byte| byte.is_ascii_digit())
    }
}

fn is_odf_length(value: &str) -> bool {
    ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .any(|unit| value.strip_suffix(unit).is_some_and(is_odf_decimal))
}

fn is_odf_percent(value: &str) -> bool {
    value.strip_suffix('%').is_some_and(is_odf_decimal)
}

fn is_odf_distance_or_percent(value: &str) -> bool {
    is_odf_length(value) || is_odf_percent(value)
}

fn is_odf_color(value: &str) -> bool {
    value.len() == 7
        && value.starts_with('#')
        && value.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit)
}

fn validate_shape_lexical_attributes(
    kind: ShapeKind,
    geometry: &[Option<String>; 4],
    line_geometry: &[Option<String>; 4],
    view_box: Option<&str>,
    points: Option<&str>,
) -> Result<()> {
    for value in geometry.iter().flatten() {
        if !is_odf_length(value) {
            return invalid("ODG shape coordinate or size is not an ODF length");
        }
    }
    validate_line_geometry(kind, line_geometry, false)?;
    if let Some(value) = view_box
        && !is_integer_list(value, 4)
    {
        return invalid("ODG shape viewBox is not four integers");
    }
    if let Some(value) = points
        && !is_points(value)
    {
        return invalid("ODG shape points are not an ODF point list");
    }
    if points.is_some()
        && (!matches!(kind, ShapeKind::Polygon | ShapeKind::Polyline) || view_box.is_none())
    {
        return invalid("ODG shape points require a polygon or polyline with a paired view box");
    }
    Ok(())
}

fn validate_line_geometry<T: AsRef<str>>(
    kind: ShapeKind,
    values: &[Option<T>; 4],
    require_complete: bool,
) -> Result<()> {
    let present = values.iter().flatten().count();
    match kind {
        ShapeKind::Line | ShapeKind::Measure if require_complete && present != 4 => {
            return invalid("ODG line geometry requires four endpoints");
        },
        ShapeKind::Connector
            if require_complete
                && (values[0].is_some() != values[1].is_some()
                    || values[2].is_some() != values[3].is_some()) =>
        {
            return invalid("ODG connector endpoints must be paired");
        },
        _ if present != 0
            && !matches!(
                kind,
                ShapeKind::Line | ShapeKind::Connector | ShapeKind::Measure
            ) =>
        {
            return invalid("ODG line endpoints require a line-like shape");
        },
        _ => {},
    }
    for value in values.iter().flatten() {
        if !is_odf_length(value.as_ref()) {
            return invalid("ODG line endpoint is not an ODF length");
        }
    }
    Ok(())
}

fn shape_name_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<Option<Range<usize>>> {
    attribute_source_span(reader, element, tag, tag_start, DRAW, b"name")
}

fn shape_attribute_source_spans(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<ShapeAttributeSpans> {
    // The legacy ordered helpers remain the error oracle for malformed spans,
    // raw attributes, and duplicate targets, so their precedence is unchanged.
    match shape_attribute_source_spans_batch(reader, element, tag, tag_start) {
        Ok(spans) => Ok(spans),
        Err(_) => Ok([
            attribute_source_span(reader, element, tag, tag_start, DRAW, b"control")?,
            shape_name_span(reader, element, tag, tag_start)?,
            attribute_source_span(reader, element, tag, tag_start, DRAW, b"layer")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"x")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"y")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"width")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"height")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"x1")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"y1")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"x2")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"y2")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"viewBox")?,
            attribute_source_span(reader, element, tag, tag_start, DRAW, b"points")?,
            attribute_source_span(reader, element, tag, tag_start, DRAW, b"transform")?,
            attribute_source_span(reader, element, tag, tag_start, SVG, b"d")?,
            attribute_source_span(reader, element, tag, tag_start, DRAW, b"style-name")?,
        ]),
    }
}

fn shape_attribute_source_spans_batch(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<ShapeAttributeSpans> {
    const REQUESTS: [(&[u8], &[u8]); 16] = [
        (DRAW, b"control"),
        (DRAW, b"name"),
        (DRAW, b"layer"),
        (SVG, b"x"),
        (SVG, b"y"),
        (SVG, b"width"),
        (SVG, b"height"),
        (SVG, b"x1"),
        (SVG, b"y1"),
        (SVG, b"x2"),
        (SVG, b"y2"),
        (SVG, b"viewBox"),
        (DRAW, b"points"),
        (DRAW, b"transform"),
        (SVG, b"d"),
        (DRAW, b"style-name"),
    ];
    let mut keys: [Option<QName<'_>>; REQUESTS.len()] = [None; REQUESTS.len()];
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let parsed_local = parsed_attribute.key.local_name();
        let Some(slot) = REQUESTS
            .iter()
            .position(|(_, wanted)| parsed_local.as_ref() == *wanted)
        else {
            continue;
        };
        let (expected, _) = REQUESTS[slot];
        let (namespace, _) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected)
            && keys[slot].replace(parsed_attribute.key).is_some()
        {
            return invalid("ODG element has duplicate namespaced attributes");
        }
    }
    let mut spans: ShapeAttributeSpans = std::array::from_fn(|_| None);
    for (slot, key) in keys.into_iter().enumerate() {
        if let Some(key) = key {
            let (start, end) = attribute_value_span(tag, key.as_ref())?;
            spans[slot] = Some(tag_start + start..tag_start + end);
        }
    }
    Ok(spans)
}

fn attribute_source_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
    expected: &[u8],
    wanted_local: &[u8],
) -> Result<Option<Range<usize>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        if parsed_attribute.key.local_name().as_ref() != wanted_local {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected)
            && local.as_ref() == wanted_local
            && key
                .replace(parsed_attribute.key.as_ref().to_vec())
                .is_some()
        {
            return invalid("ODG element has duplicate namespaced attributes");
        }
    }
    let Some(name_key) = key else {
        return Ok(None);
    };
    let (start, end) = attribute_value_span(tag, &name_key)?;
    Ok(Some(tag_start + start..tag_start + end))
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return invalid("ODG shape attribute is missing '='");
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| Error::InvalidFormat("ODG shape attribute is not quoted".to_string()))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= tag.len() {
            return invalid("ODG shape attribute is unterminated");
        }
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    invalid("ODG shape name span was not found")
}

fn attribute_token_span(tag: &[u8], wanted: &[u8]) -> Result<Range<usize>> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        let whitespace_start = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return invalid("ODG transition attribute is missing '='");
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| Error::InvalidFormat("ODG transition attribute is not quoted".into()))?;
        cursor += 1;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return invalid("ODG transition attribute is unterminated");
        }
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok(whitespace_start..cursor);
        }
    }
    invalid("ODG transition attribute span was not found")
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| {
        Error::InvalidFormat("ODG XML position exceeds platform limits".to_string())
    })
}

fn position_reader(reader: &quick_xml::Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| {
        Error::InvalidFormat("ODG XML position exceeds platform limits".to_string())
    })
}

fn shape_kind(namespace: NamespaceKind, local: &[u8]) -> Option<ShapeKind> {
    if namespace == NamespaceKind::Dr3d {
        return Some(match local {
            b"scene" => ShapeKind::ThreeDimensionalScene,
            b"light" => ShapeKind::ThreeDimensionalLight,
            b"cube" => ShapeKind::ThreeDimensionalCube,
            b"sphere" => ShapeKind::ThreeDimensionalSphere,
            b"extrude" => ShapeKind::ThreeDimensionalExtrude,
            b"rotate" => ShapeKind::ThreeDimensionalRotate,
            _ => return None,
        });
    }
    (namespace == NamespaceKind::Draw).then_some(match local {
        b"caption" => ShapeKind::Caption,
        b"circle" => ShapeKind::Circle,
        b"connector" => ShapeKind::Connector,
        b"control" => ShapeKind::Control,
        b"custom-shape" => ShapeKind::Custom,
        b"ellipse" => ShapeKind::Ellipse,
        b"frame" => ShapeKind::Frame,
        b"g" => ShapeKind::Group,
        b"line" => ShapeKind::Line,
        b"measure" => ShapeKind::Measure,
        b"path" => ShapeKind::Path,
        b"page-thumbnail" => ShapeKind::PageThumbnail,
        b"polygon" => ShapeKind::Polygon,
        b"polyline" => ShapeKind::Polyline,
        b"rect" => ShapeKind::Rectangle,
        b"regular-polygon" => ShapeKind::RegularPolygon,
        _ => return None,
    })
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DRAW => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(uri)) if *uri == DR3D => NamespaceKind::Dr3d,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == SVG => NamespaceKind::Svg,
        ResolveResult::Bound(Namespace(uri)) if *uri == FORM => NamespaceKind::Form,
        ResolveResult::Bound(Namespace(uri)) if *uri == STYLE => NamespaceKind::Style,
        ResolveResult::Bound(Namespace(uri)) if *uri == PRESENTATION => NamespaceKind::Presentation,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

/// Refuse a cross-page move of the subtree at `span` when an identity
/// dependency would cross the page boundary: an `xml:id`/`draw:id` the
/// subtree declares that any attribute outside it names, or a connector
/// end (`draw:start-shape`/`draw:end-shape`) inside it that names an
/// identity declared outside it.
fn check_shape_identity(xml: &str, span: &Range<usize>) -> Result<()> {
    validate_raw_xml_events(xml, "ODG dependency event")?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut inside_ids = Vec::new();
    let mut inside_refs = Vec::new();
    let mut outside_values = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let (_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG dependency XML: {error}"))
        })?;
        validate_borrowed_xml_event(&event, "ODG dependency event")?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let inside = span.start <= start && start < span.end;
                for raw in element.attributes() {
                    let parsed = raw.map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG attribute: {error}"))
                    })?;
                    let value = String::from_utf8_lossy(parsed.value.as_ref()).into_owned();
                    let (namespace, local) = reader.resolver().resolve_attribute(parsed.key);
                    let is_id = (parsed.key.as_ref() == b"xml:id")
                        || (resolved_bound(&namespace, DRAW) && local.as_ref() == b"id");
                    let is_glue = resolved_bound(&namespace, DRAW)
                        && matches!(local.as_ref(), b"start-shape" | b"end-shape");
                    if inside {
                        if is_id {
                            inside_ids.push(value);
                        } else if is_glue {
                            inside_refs.push(value);
                        }
                    } else if !is_id {
                        outside_values.push(value);
                    }
                }
            },
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG dependencies");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
    if inside_ids
        .iter()
        .any(|id| outside_values.iter().any(|value| value == id))
    {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odg.package.snapshot",
            "ODG shape move to another page is blocked by a reference to its identity",
        )));
    }
    if inside_refs.iter().any(|value| !inside_ids.contains(value)) {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odg.package.snapshot",
            "ODG connector move to another page is blocked by its glued shapes",
        )));
    }
    Ok(())
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn joined_refusal(reason: &'static str) -> Error {
    Error::Refused(litchi_core::edit::Refusal::new(
        litchi_core::edit::RefusalKind::UnsupportedOperation,
        "odg.edit.join",
        reason,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::events::{BytesCData, BytesPI, BytesText};

    fn lazy_source_package() -> Package {
        let content = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"/></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;
        let styles = br#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:styles/></office:document-styles>"#;
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(MIMETYPE).unwrap();
        writer.add_file("content.xml", content).unwrap();
        writer.add_file("styles.xml", styles).unwrap();
        Package::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn edit_parts_stay_borrowed_for_noop_and_pure_replacement() {
        let source = lazy_source_package();
        let edit = source.edit();
        assert!(edit.content.owned.is_none());
        assert_eq!(edit.content.as_ptr(), source.content_xml().as_ptr());
        let styles = edit.styles.as_ref().expect("fixture has styles.xml");
        assert!(styles.owned.is_none());
        assert_eq!(styles.as_ptr(), source.styles_xml().unwrap().as_ptr());
        drop(edit);

        let mut mutation = source.edit();
        mutation.set_page_name(0, "Materialized page").unwrap();
        assert!(mutation.content.owned.is_some());
        assert!(
            mutation
                .styles
                .as_ref()
                .is_some_and(|part| part.owned.is_none())
        );
        drop(mutation);

        let no_op = source.edit().commit().unwrap();
        assert!(!no_op.diagnostics().changed());
        assert_eq!(no_op.package().as_bytes(), source.as_bytes());

        let mut replacement_edit = source.edit();
        replacement_edit
            .add_page(Page::new("Replacement page"))
            .unwrap();
        let replacement = replacement_edit.commit().unwrap();
        let limits = CompositionLimits::new(2, 2, 4, 2);
        let mut joined = source.joined_edits(limits);
        joined
            .join(replacement.patch().prepare("replacement", limits).unwrap())
            .unwrap();
        let mut owner = source.edit();
        owner.join(joined).unwrap();
        owner.apply_pending_joined().unwrap();
        assert!(owner.package_replacement.is_some());
        assert!(owner.content.owned.is_none());
        assert!(
            owner
                .styles
                .as_ref()
                .is_some_and(|part| part.owned.is_none())
        );
        let committed = owner.commit_core().unwrap();
        assert!(committed.diagnostics().changed());
    }

    #[test]
    fn borrowed_event_admission_rejects_oversized_ignored_inputs() {
        let oversized = "x".repeat(MAX_TEXT_BYTES + 1);
        assert!(
            validate_borrowed_xml_event(&Event::Text(BytesText::new(&oversized)), "text probe")
                .is_err()
        );
        assert!(
            validate_borrowed_xml_event(&Event::CData(BytesCData::new(&oversized)), "cdata probe")
                .is_err()
        );
        assert!(
            validate_borrowed_xml_event(
                &Event::Comment(BytesText::new(&oversized)),
                "comment probe"
            )
            .is_err()
        );
        assert!(
            validate_borrowed_xml_event(
                &Event::PI(BytesPI::new(&oversized)),
                "processing-instruction probe"
            )
            .is_err()
        );

        let start = format!("draw:rect value=\"{oversized}\"");
        assert!(
            validate_borrowed_xml_event(
                &Event::Start(BytesStart::from_content(&start, "draw:rect".len())),
                "start-tag probe"
            )
            .is_err()
        );
    }

    #[test]
    fn bounded_clone_admission_refuses_before_reservation() {
        let source = "0123456789";
        let error = clone_bounded_string(source, 9, "allocation probe").unwrap_err();
        assert!(
            matches!(error, Error::InvalidFormat(reason) if reason.contains("allocation probe"))
        );
    }

    #[test]
    fn attribute_value_batch_applies_the_ordered_value_admission() {
        let oversized = "x".repeat(MAX_TEXT_BYTES + 1);
        let xml = format!(
            r#"<d:rect xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" d:name="kept" d:layer="{oversized}"/>"#
        );
        let mut reader = NsReader::from_str(&xml);
        let element = match reader.read_resolved_event().unwrap().1 {
            Event::Empty(element) => element.into_owned(),
            other => panic!("expected an empty element, got {other:?}"),
        };
        let requests = [(DRAW, b"name".as_slice()), (DRAW, b"layer".as_slice())];
        let batch = attribute_values(&reader, &element, requests).unwrap_err();
        let ordered = attribute(&reader, &element, DRAW, b"layer").unwrap_err();
        assert_eq!(batch.to_string(), ordered.to_string());
        assert!(
            matches!(batch, Error::InvalidFormat(reason) if reason.contains("ODG attribute value"))
        );
    }
}

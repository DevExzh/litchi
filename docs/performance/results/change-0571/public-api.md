# The public API added by change 0571, verbatim

Recorded here because the API surface is the part of this change that a reviewer
most needs to see exactly, and because ADR 0010 constrains what may appear on it.

## The opaque handle, in the crate that owns the OPC-to-ZIP translation

```rust
// crates/litchi-opc/src/prepared.rs  ->  litchi_opc::PreparedOoxmlSource
pub struct PreparedOoxmlSource { /* private: package, main_part_content_type, prepared_version */ }
impl std::fmt::Debug for PreparedOoxmlSource { /* content type and revision only */ }

impl PreparedOoxmlSource {
    pub fn new(package: SourceBackedPackage, main_part_content_type: &str) -> Result<Self>;
    #[must_use] pub fn main_part_content_type(&self) -> &str;
    #[must_use] pub const fn prepared_source_version(&self) -> SourceVersion;
    pub fn source_version(&self) -> Result<SourceVersion>;
    pub fn ensure_current(&self) -> Result<()>;
    pub fn into_package(self) -> Result<SourceBackedPackage>;
}
```

`into_package` **consumes** the handle and the type is not `Clone`, so one
prepared snapshot reaches exactly one adapter. No archive type, physical
identifier, lock or executor appears: the surface is the crate's own OPC package
type, a string slice, a source version and the crate's error.

## The detection entry point, in the facade

```rust
#[cfg(all(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
          any(unix, windows)))]
pub struct PreparedDetection { /* private: format, prepared */ }

impl PreparedDetection {
    #[must_use] pub const fn format(&self) -> litchi_core::detection::FileFormat;
    #[must_use] pub const fn is_prepared(&self) -> bool;
    #[must_use] pub fn prepared(&self) -> Option<&crate::opc::PreparedOoxmlSource>;
    #[must_use] pub fn into_prepared(self) -> Option<crate::opc::PreparedOoxmlSource>;
    #[must_use] pub fn into_parts(self)
        -> (litchi_core::detection::FileFormat, Option<crate::opc::PreparedOoxmlSource>);
}

pub fn detect_and_prepare<P: AsRef<Path>>(path: P) -> Option<PreparedDetection>;
pub fn detect_and_prepare_with_limits<P: AsRef<Path>>(
    path: P, limits: crate::opc::ReadLimits,
) -> Option<PreparedDetection>;
```

`format()` is always the classification. `prepared()` returns `None` for anything
that is not an OOXML package, and `detect_and_prepare` returns `None` exactly
when the existing detector does.

## The adopters

```rust
#[cfg(feature = "docx")]                          // impl Document
pub fn from_prepared(prepared: crate::opc::PreparedOoxmlSource) -> Result<Self>;
#[cfg(feature = "pptx")]                          // impl Presentation
pub fn from_prepared(prepared: crate::opc::PreparedOoxmlSource) -> Result<Self>;
#[cfg(any(feature = "xlsx", feature = "xlsb"))]   // impl Workbook
pub fn from_prepared(prepared: crate::opc::PreparedOoxmlSource) -> Result<Self>;
```

## The typed wrong-format error

```rust
// crates/litchi-core/src/error/types.rs, in the existing #[non_exhaustive] enum Error
    /// The input is a recognized Office format, but not one this opener owns.
    #[error("Detected {detected:?}, which this opener does not handle")]
    UnexpectedFormat {
        /// Neutral classification produced by detection.
        detected: crate::detection::FileFormat,
    },
```

The format type already lives in the same crate as the error and the enum is
non-exhaustive, so this needed no new home and is not a breaking change.

## Supporting classifier functions

```rust
// impl DetectedFormat
#[must_use] pub const fn format(&self) -> litchi_core::detection::FileFormat;

// detection_smart::ooxml
#[must_use] pub fn format_for_main_content_type(content_type: &str) -> Option<FileFormat>;
pub fn try_classify_ooxml_source_backed_package(
    package: &litchi_opc::SourceBackedPackage,
) -> crate::opc::Result<Option<(FileFormat, &str)>>;
```

The existing classifier keeps its signature and delegates, so no existing caller
changed. The content-type table is shared by the forward catalog scan and by
`format_for_main_content_type`, which is what prevents a handle from disagreeing
with the classification that produced it.

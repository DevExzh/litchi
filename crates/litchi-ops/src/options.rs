//! Options for opening an input: [`OpenOptions`] and [`SourceMode`]
//! (record 0003 §2.5, decision D5).
//!
//! Plain `open(path)` needs none of this. [`OpenOptions::default`] is the
//! common case, and every knob left unset keeps the leaf crate's own bounded
//! default. Options never carry credentials: an encrypted input opens as
//! [`Locked`](crate::Locked), and the password is passed to its consuming
//! `unlock` transition instead (ADR 0006).

use litchi_core::ExecutionContext;

/// How an opened input is held while it is read.
///
/// ```
/// use litchi_ops::{OpenOptions, SourceMode};
///
/// assert_eq!(SourceMode::default(), SourceMode::Auto);
/// let options = OpenOptions::new().source_mode(SourceMode::Eager);
/// # let _ = options;
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum SourceMode {
    /// Keep the existing per-route choice: source-backed (lazy, positional
    /// reads over the retained input) where the leaf and target support it,
    /// eager otherwise.
    #[default]
    Auto,
    /// Materialize the leaf's eager owner at open, even where a
    /// source-backed owner exists.
    Eager,
}

/// Options for one open call.
///
/// The defaults are the common case: [`SourceMode::Auto`], source retention
/// on, no caller-supplied execution context, and every per-format limit left
/// at the leaf crate's bounded default. Builder methods consume and return
/// the options, so they chain.
///
/// `OpenOptions` holds limits and policy only. It never holds credentials,
/// so it stays `Clone` and safe to log.
///
/// ```
/// use litchi_ops::{OpenOptions, SourceMode};
///
/// // The common case: nothing to set.
/// let plain = OpenOptions::default();
/// # let _ = plain;
///
/// // A caller that caps the input size and wants eager owners.
/// let tuned = OpenOptions::new()
///     .max_input_bytes(64 << 20)
///     .source_mode(SourceMode::Eager)
///     .retain_source(false);
/// # let _ = tuned;
/// ```
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct OpenOptions {
    /// Ceiling on the physical input size. `None` keeps each route's own
    /// ceiling.
    pub(crate) max_input_bytes: Option<u64>,
    /// Caller-supplied ADR 0031 execution context. `None` runs sequentially
    /// under the leaf defaults; no ambient pool is created.
    pub(crate) execution: Option<ExecutionContext>,
    /// Source-backed versus eager owner selection.
    pub(crate) source_mode: SourceMode,
    /// Whether the handle retains the source its edit entry is built from
    /// (record 0003 §2.3; record 0007 Q3 default `true`).
    pub(crate) retain_source: bool,
    /// OOXML package read limits shared by DOCX, PPTX, XLSX and XLSB.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    pub(crate) ooxml_limits: Option<litchi_opc::SourceLimits>,
    /// Outer OOXML encryption limits retained by password-bearing opens.
    #[cfg(feature = "encryption")]
    pub(crate) ooxml_encryption_limits: litchi_crypto::ooxml::Limits,
    /// OOXML source-backed part cache policy.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    pub(crate) ooxml_cache: Option<litchi_opc::SourceCacheLimits>,
    /// Legacy Word options other than credentials.
    #[cfg(feature = "doc")]
    pub(crate) doc: Option<DocOptions>,
    /// Legacy `PowerPoint` source limits.
    #[cfg(feature = "ppt")]
    pub(crate) ppt_source_limits: Option<litchi_ppt::SourceLimits>,
    /// Native RTF source limits and explicit repair policy.
    #[cfg(feature = "rtf")]
    pub(crate) rtf_options: Option<litchi_rtf::OpenOptions>,
    /// ODF ingress options (source and semantic limits) shared by ODT, ODS
    /// and ODP, including the other ODF text/object families.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    pub(crate) odf_options: Option<litchi_odf_common::OpenOptions>,
    /// Pages ingress options (source and semantic limits).
    #[cfg(feature = "pages")]
    pub(crate) pages_options: Option<litchi_pages::OpenOptions>,
    /// Keynote ingress options (source and semantic limits).
    #[cfg(feature = "keynote")]
    pub(crate) keynote_options: Option<litchi_keynote::OpenOptions>,
    /// Numbers ingress options (source and semantic limits).
    #[cfg(feature = "numbers")]
    pub(crate) numbers_options: Option<litchi_numbers::OpenOptions>,
}

impl OpenOptions {
    /// Options for the common case; the same as [`OpenOptions::default`].
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new();
    /// # let _ = options;
    /// ```
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_input_bytes: None,
            execution: None,
            source_mode: SourceMode::Auto,
            retain_source: true,
            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            ooxml_limits: None,
            #[cfg(feature = "encryption")]
            ooxml_encryption_limits: litchi_crypto::ooxml::Limits::DEFAULT,
            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            ooxml_cache: None,
            #[cfg(feature = "doc")]
            doc: None,
            #[cfg(feature = "ppt")]
            ppt_source_limits: None,
            #[cfg(feature = "rtf")]
            rtf_options: None,
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            odf_options: None,
            #[cfg(feature = "pages")]
            pages_options: None,
            #[cfg(feature = "keynote")]
            keynote_options: None,
            #[cfg(feature = "numbers")]
            numbers_options: None,
        }
    }

    /// Refuse inputs whose physical size exceeds `bytes`.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new().max_input_bytes(16 << 20);
    /// # let _ = options;
    /// ```
    #[must_use]
    pub const fn max_input_bytes(mut self, bytes: u64) -> Self {
        self.max_input_bytes = Some(bytes);
        self
    }

    /// Run open-time work under a caller-owned ADR 0031 execution context.
    ///
    /// In this release an open checks the context's cancellation token
    /// before and after detection and returns
    /// [`Code::Cancelled`](crate::Code::Cancelled) when it is set. The leaf
    /// parsers do not yet charge the context's budget: the source-backed DOCX
    /// owner refuses markup-compatibility content under a context (record
    /// 0008, finding A1), so no opener builds a context-managed owner until
    /// that is fixed.
    ///
    /// ```
    /// use std::num::{NonZeroU64, NonZeroUsize};
    ///
    /// use litchi_core::{Budget, CancellationSource, Limits, Profile};
    /// use litchi_core::{ExecutionContext, ExecutionLimits};
    ///
    /// let budget = Budget::root("open", Limits::for_profile(Profile::Desktop));
    /// let (_source, token) = CancellationSource::pair();
    /// let limits = ExecutionLimits::new(
    ///     NonZeroUsize::MIN,
    ///     NonZeroUsize::MIN,
    ///     NonZeroU64::new(1 << 20).unwrap(),
    ///     0,
    /// )
    /// .unwrap();
    /// let context = ExecutionContext::new(budget, token, limits);
    /// let options = litchi_ops::OpenOptions::new().execution(context);
    /// # let _ = options;
    /// ```
    #[must_use]
    pub fn execution(mut self, context: ExecutionContext) -> Self {
        self.execution = Some(context);
        self
    }

    /// Choose source-backed or eager owners; see [`SourceMode`].
    ///
    /// ```
    /// use litchi_ops::{OpenOptions, SourceMode};
    ///
    /// let options = OpenOptions::new().source_mode(SourceMode::Eager);
    /// # let _ = options;
    /// ```
    #[must_use]
    pub const fn source_mode(mut self, mode: SourceMode) -> Self {
        self.source_mode = mode;
        self
    }

    /// Whether the opened handle keeps the source its edit entry is built
    /// from. The default is `true`, so write verbs need no second open.
    /// `false` releases the input after open on the routes that would
    /// otherwise keep it alive only for editing; write verbs then refuse.
    ///
    /// ```
    /// let read_only = litchi_ops::OpenOptions::new().retain_source(false);
    /// # let _ = read_only;
    /// ```
    #[must_use]
    pub const fn retain_source(mut self, retain: bool) -> Self {
        self.retain_source = retain;
        self
    }

    /// Replace the OOXML package read limits used for DOCX, PPTX, XLSX and
    /// XLSB inputs.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .ooxml_limits(litchi_opc::SourceLimits::default());
    /// # let _ = options;
    /// ```
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[must_use]
    pub const fn ooxml_limits(mut self, limits: litchi_opc::SourceLimits) -> Self {
        self.ooxml_limits = Some(limits);
        self
    }

    /// Replace the outer OOXML encryption limits used by password-bearing
    /// DOCX, PPTX, XLSX and XLSB opens.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub const fn ooxml_encryption_limits(mut self, limits: litchi_crypto::ooxml::Limits) -> Self {
        self.ooxml_encryption_limits = limits;
        self
    }

    /// Replace the source-backed OOXML part cache policy.
    ///
    /// ```
    /// let cache = litchi_opc::SourceCacheLimits::new(8 << 20, 64).unwrap();
    /// let options = litchi_ops::OpenOptions::new().ooxml_cache(cache);
    /// # let _ = options;
    /// ```
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[must_use]
    pub const fn ooxml_cache(mut self, limits: litchi_opc::SourceCacheLimits) -> Self {
        self.ooxml_cache = Some(limits);
        self
    }

    /// Replace the legacy Word options.
    ///
    /// ```
    /// use litchi_ops::{DocOptions, OpenOptions};
    ///
    /// let doc = DocOptions::new().leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
    /// let options = OpenOptions::new().doc(doc);
    /// # let _ = options;
    /// ```
    #[cfg(feature = "doc")]
    #[must_use]
    pub const fn doc(mut self, options: DocOptions) -> Self {
        self.doc = Some(options);
        self
    }

    /// Replace the legacy `PowerPoint` source limits.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .ppt_source_limits(litchi_ppt::SourceLimits::default());
    /// # let _ = options;
    /// ```
    #[cfg(feature = "ppt")]
    #[must_use]
    pub const fn ppt_source_limits(mut self, limits: litchi_ppt::SourceLimits) -> Self {
        self.ppt_source_limits = Some(limits);
        self
    }

    /// Replace the native RTF source limits and repair policy.
    #[cfg(feature = "rtf")]
    #[must_use]
    pub const fn rtf_options(mut self, options: litchi_rtf::OpenOptions) -> Self {
        self.rtf_options = Some(options);
        self
    }

    /// Replace the ODF ingress options (source and semantic limits) used for
    /// packaged and flat ODF inputs opened from bytes or a path.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .odf_options(litchi_odf_common::OpenOptions::new());
    /// # let _ = options;
    /// ```
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    #[must_use]
    pub const fn odf_options(mut self, options: litchi_odf_common::OpenOptions) -> Self {
        self.odf_options = Some(options);
        self
    }

    /// Replace the Pages ingress options (source and semantic limits) used
    /// for Pages packages and package directories (record 0306).
    ///
    /// Semantic limits apply on the single detection pass; source limits
    /// other than the leaf defaults make the leaf parse the package again
    /// under them, because detection prepares the container under the hard
    /// ceilings.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .pages_options(litchi_pages::OpenOptions::new());
    /// # let _ = options;
    /// ```
    #[cfg(feature = "pages")]
    #[must_use]
    pub const fn pages_options(mut self, options: litchi_pages::OpenOptions) -> Self {
        self.pages_options = Some(options);
        self
    }

    /// Replace the Keynote ingress options (source and semantic limits) used
    /// for Keynote packages and package directories (record 0306).
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .keynote_options(litchi_keynote::OpenOptions::new());
    /// # let _ = options;
    /// ```
    #[cfg(feature = "keynote")]
    #[must_use]
    pub const fn keynote_options(mut self, options: litchi_keynote::OpenOptions) -> Self {
        self.keynote_options = Some(options);
        self
    }

    /// Replace the Numbers ingress options (source and semantic limits) used
    /// for Numbers packages and package directories (record 0306); following
    /// the same ingress-option contract as the Pages setter above.
    ///
    /// ```
    /// let options = litchi_ops::OpenOptions::new()
    ///     .numbers_options(litchi_numbers::OpenOptions::new());
    /// # let _ = options;
    /// ```
    #[cfg(feature = "numbers")]
    #[must_use]
    pub const fn numbers_options(mut self, options: litchi_numbers::OpenOptions) -> Self {
        self.numbers_options = Some(options);
        self
    }

    /// The selected ODF ingress options, or the defaults.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    pub(crate) fn odf_open_options(&self) -> litchi_odf_common::OpenOptions {
        self.odf_options.unwrap_or_default()
    }

    /// Convert the owner-specific ingress settings to the limits required by
    /// the raw detection probes. This is kept private so public raw callers
    /// cannot accidentally recover an `OpenOptions` exception.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp",
        feature = "pages",
        feature = "keynote",
        feature = "numbers",
        odf_objects
    ))]
    pub(crate) fn detection_limits(&self) -> crate::detect::SourceLimits {
        let mut limits = crate::detect::SourceLimits::new();
        if let Some(maximum) = self.max_input_bytes {
            limits = limits.max_input_bytes(maximum);
        }

        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        if let Some(source_limits) = self.ooxml_limits {
            limits = limits.ooxml_source_limits(source_limits);
        }

        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        if let Some(options) = self.odf_options {
            limits = limits.odf_source_limits(options.source_limits_value());
        }

        #[cfg(feature = "rtf")]
        if let Some(options) = self.rtf_options {
            limits = limits.rtf_source_limits(options.source_limits_value());
        }

        #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
        {
            // A format-specific DOC or PPT ceiling must not reject another
            // CFB format before classification. Owner admission applies that
            // format's policy after detection; only the common source ceiling
            // applies to every CFB candidate here.
            let cfb_input = self.max_input_bytes;
            if let Some(value) = cfb_input {
                let value = value.min(litchi_cfb::OleFileLimits::MAX_INPUT_BYTES);
                if let Ok(source_limits) = litchi_cfb::OleFileLimits::new(value) {
                    limits = limits.cfb_source_limits(source_limits);
                }
            }
        }

        limits
    }
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self::new()
    }
}

/// Legacy Word (`.doc`) open options, without credentials.
///
/// `litchi_doc::OpenOptions` also carries a password and is therefore not
/// `Clone`; this type keeps only the policy knobs.
///
/// ```
/// use litchi_ops::DocOptions;
///
/// let strict = DocOptions::default();
/// assert_eq!(strict, DocOptions::new());
/// ```
#[cfg(feature = "doc")]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub struct DocOptions {
    pub(crate) leniency: litchi_doc::Leniency,
    pub(crate) source_limits: Option<litchi_doc::SourceLimits>,
}

#[cfg(feature = "doc")]
impl DocOptions {
    /// Leaf defaults: [`litchi_doc::Leniency::Strict`].
    #[must_use]
    pub const fn new() -> Self {
        Self {
            leniency: litchi_doc::Leniency::Strict,
            source_limits: None,
        }
    }

    /// Set finite DOC package and stream ingestion limits.
    ///
    /// ```
    /// let limits = litchi_doc::SourceLimits::builder()
    ///     .max_input_bytes(1_000_000)?.max_entry_bytes(500_000)?
    ///     .max_aggregate_bytes(800_000)?.build()?;
    /// let options = litchi_ops::DocOptions::new().source_limits(limits);
    /// # Ok::<(), litchi_doc::Error>(())
    /// ```
    #[must_use]
    pub const fn source_limits(mut self, limits: litchi_doc::SourceLimits) -> Self {
        self.source_limits = Some(limits);
        self
    }

    /// Set the stylesheet-defect policy.
    ///
    /// ```
    /// let options = litchi_ops::DocOptions::new()
    ///     .leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
    /// # let _ = options;
    /// ```
    #[must_use]
    pub const fn leniency(mut self, leniency: litchi_doc::Leniency) -> Self {
        self.leniency = leniency;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const fn assert_traits<T: Send + Sync + Clone + std::fmt::Debug>() {}
    const _: () = {
        assert_traits::<OpenOptions>();
        assert_traits::<SourceMode>();
    };
    #[cfg(feature = "doc")]
    const _: () = assert_traits::<DocOptions>();

    #[test]
    fn default_is_the_common_case() {
        let options = OpenOptions::default();
        assert_eq!(options.max_input_bytes, None);
        assert!(options.execution.is_none());
        assert_eq!(options.source_mode, SourceMode::Auto);
        assert!(
            options.retain_source,
            "record 0007 Q3: retain_source defaults to true"
        );
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        {
            assert_eq!(options.ooxml_limits, None);
            assert_eq!(options.ooxml_cache, None);
        }
        #[cfg(feature = "doc")]
        assert_eq!(options.doc, None);
        #[cfg(feature = "ppt")]
        assert_eq!(options.ppt_source_limits, None);
        #[cfg(feature = "rtf")]
        assert_eq!(options.rtf_options, None);
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        assert_eq!(options.odf_options, None);
    }

    #[test]
    fn new_matches_default() {
        let a = OpenOptions::new();
        let b = OpenOptions::default();
        assert_eq!(format!("{a:?}"), format!("{b:?}"));
    }

    #[test]
    fn builder_sets_common_knobs() {
        let options = OpenOptions::new()
            .max_input_bytes(1024)
            .source_mode(SourceMode::Eager)
            .retain_source(false);
        assert_eq!(options.max_input_bytes, Some(1024));
        assert_eq!(options.source_mode, SourceMode::Eager);
        assert!(!options.retain_source);
        // Later calls overwrite earlier ones.
        let options = options.max_input_bytes(7).retain_source(true);
        assert_eq!(options.max_input_bytes, Some(7));
        assert!(options.retain_source);
    }

    #[test]
    fn builder_sets_execution_context() {
        use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits, Profile};
        use std::num::{NonZeroU64, NonZeroUsize};

        let budget = Budget::root("test", Limits::for_profile(Profile::Desktop));
        let (_source, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("valid limits");
        let options = OpenOptions::new().execution(ExecutionContext::new(budget, token, limits));
        let context = options.execution.as_ref().expect("context set");
        assert_eq!(context.limits().workers(), NonZeroUsize::MIN);
        // Cloning the options shares the caller's context rather than
        // dropping it.
        assert!(options.clone().execution.is_some());
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[test]
    fn builder_sets_ooxml_sub_options() {
        let limits = litchi_opc::SourceLimits::default();
        let cache = litchi_opc::SourceCacheLimits::new(4096, 2).expect("non-zero cache");
        let options = OpenOptions::new().ooxml_limits(limits).ooxml_cache(cache);
        assert_eq!(options.ooxml_limits, Some(limits));
        assert_eq!(options.ooxml_cache, Some(cache));
    }

    #[cfg(feature = "doc")]
    #[test]
    fn builder_sets_doc_sub_options() {
        assert_eq!(DocOptions::default().leniency, litchi_doc::Leniency::Strict);
        let doc = DocOptions::new().leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
        let options = OpenOptions::new().doc(doc);
        assert_eq!(options.doc, Some(doc));
    }

    #[cfg(all(feature = "doc", feature = "ppt"))]
    #[test]
    fn doc_source_ceiling_does_not_reject_ppt_detection() {
        let mut builder = litchi_ppt::writer::Builder::new();
        builder.add_slide().unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let doc = litchi_doc::SourceLimits::builder()
            .max_input_bytes(bytes.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        let options = OpenOptions::new().doc(DocOptions::new().source_limits(doc));
        assert_eq!(
            crate::detect::detect_format_smart(bytes.clone(), &options.detection_limits())
                .map(|value| value.format()),
            Some(litchi_core::detection::FileFormat::Ppt),
        );
        let common = options.max_input_bytes((bytes.len() - 1) as u64);
        assert!(crate::detect::detect_format_smart(bytes, &common.detection_limits()).is_none());
    }

    #[cfg(feature = "ppt")]
    #[test]
    fn builder_sets_ppt_sub_options() {
        let limits = litchi_ppt::SourceLimits::default();
        let options = OpenOptions::new().ppt_source_limits(limits);
        assert_eq!(options.ppt_source_limits, Some(limits));
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    #[test]
    fn builder_sets_odf_sub_options() {
        let odf = litchi_odf_common::OpenOptions::new()
            .source_limits(litchi_odf_common::SourceLimits::default());
        let options = OpenOptions::new().odf_options(odf);
        assert_eq!(options.odf_options, Some(odf));
    }
}

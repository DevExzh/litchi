//! Static File dispatch to unified read adapters.

use crate::node::Node;
use crate::path::IntoPath;
use crate::query::{IntoQuery, Matches};
use crate::view::Mode;
use crate::{Error, Family, File, Result};
use litchi_core::FileFormat;

#[cfg_attr(
    not(any(
        feature = "doc",
        feature = "docx",
        feature = "xlsx",
        feature = "pptx",
        feature = "ppt",
        feature = "xlsb",
        feature = "xls",
        feature = "rtf",
        feature = "ods",
        feature = "odt",
        feature = "odp"
    )),
    allow(
        clippy::only_used_in_recursion,
        reason = "without a unified format adapter, arguments only pass through signed wrappers"
    )
)]
impl File {
    /// Report available unified operations without reading document payloads.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// let file = litchi_ops::open("report.docx")?;
    /// assert!(file.capabilities().verb(litchi_ops::capabilities::Verb::Get).is_supported());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[must_use]
    pub fn capabilities(&self) -> crate::capabilities::Capabilities {
        crate::capabilities::Capabilities::for_file(self)
    }

    /// The document root, without entering its children.
    ///
    /// # Errors
    /// Returns source revision errors, `document_protected` for locked files,
    /// or `unsupported_capability` until the format's adapter is integrated.
    pub fn root(&self) -> Result<Node<'_>> {
        self.get("/")
    }

    /// Resolve a positional or semantic path to a typed node.
    /// Native identity selection requires [`crate::path::raw::path`].
    ///
    /// # Errors
    /// Returns `invalid_path` for malformed or non-admitted native paths,
    /// `path_not_found` for missing nodes, typed leaf errors for failed reads,
    /// and protection/capability errors as described by [`Self::root`].
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// let file = litchi_ops::open("report.docx")?;
    /// let paragraph = file.get("/body/p[1]")?;
    /// assert_eq!(paragraph.kind(), "paragraph");
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn get<'p>(&self, path: impl IntoPath<'p>) -> Result<Node<'_>> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::docx::Nav::new(handle), &path)
                    .map(Node::Docx)
            },
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::doc::Nav::new(handle), &path)
                    .map(Node::Doc)
            },
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::xlsx::Nav::new(handle), &path)
                    .map(Node::Xlsx)
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::pptx::Nav::new(handle), &path)
                    .map(Node::Pptx)
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::ppt::Nav::new(handle), &path)
                    .map(Node::Ppt)
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::xlsb::Nav::new(handle), &path)
                    .map(Node::Xlsb)
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::xls::Nav::new(handle), &path)
                    .map(Node::Xls)
            },
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::rtf::Nav::new(handle), &path)
                    .map(Node::Rtf)
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::ods::Nav::new(handle), &path)
                    .map(Node::Ods)
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::odt::Nav::new(handle), &path)
                    .map(Node::Odt)
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::get(&crate::format::odp::Nav::new(handle), &path)
                    .map(Node::Odp)
            },
            Self::Signed(signed) => signed.file().get(path),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            other => Err(pending(other)),
        }
    }

    /// Resolve a path, returning `None` only when the target does not exist.
    ///
    /// # Errors
    /// As [`Self::get`], except a missing target is `Ok(None)`.
    pub fn find<'p>(&self, path: impl IntoPath<'p>) -> Result<Option<Node<'_>>> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::docx::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Docx))
            },
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::doc::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Doc))
            },
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::xlsx::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Xlsx))
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::pptx::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Pptx))
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::ppt::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Ppt))
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::xlsb::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Xlsb))
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::xls::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Xls))
            },
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::rtf::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Rtf))
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::ods::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Ods))
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::odt::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Odt))
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => {
                let path = path.into_path()?;
                crate::path::resolve::find(&crate::format::odp::Nav::new(handle), &path)
                    .map(|node| node.map(Node::Odp))
            },
            Self::Signed(signed) => signed.file().find(path),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            other => Err(pending(other)),
        }
    }

    /// Stream matching nodes in document order, optionally reusing a compiled plan.
    /// Payload traversal starts when the iterator is advanced. Native predicates require
    /// [`crate::query::raw::query`]. Unknown keys produce compile diagnostics.
    ///
    /// # Errors
    /// Returns typed selector, format, protection and capability errors before
    /// iteration. Lazy read errors are iterator items; an error ends iteration.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// let file = litchi_ops::open("report.docx")?;
    /// let first = file.query("p[style=Heading1]")?.first()?;
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn query<'q>(&self, selector: impl IntoQuery<'q>) -> Result<Matches<'_, 'q>> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => {
                crate::query::run(crate::format::docx::Nav::new(handle), selector)
                    .map(Matches::docx)
            },
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                crate::query::run(crate::format::doc::Nav::new(handle), selector).map(Matches::doc)
            },
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                crate::query::run(crate::format::xlsx::Nav::new(handle), selector)
                    .map(Matches::xlsx)
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                crate::query::run(crate::format::pptx::Nav::new(handle), selector)
                    .map(Matches::pptx)
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                crate::query::run(crate::format::ppt::Nav::new(handle), selector).map(Matches::ppt)
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                crate::query::run(crate::format::xlsb::Nav::new(handle), selector)
                    .map(Matches::xlsb)
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                crate::query::run(crate::format::xls::Nav::new(handle), selector).map(Matches::xls)
            },
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                crate::query::run(crate::format::rtf::Nav::new(handle), selector).map(Matches::rtf)
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                crate::query::run(crate::format::ods::Nav::new(handle), selector).map(Matches::ods)
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) => {
                crate::query::run(crate::format::odt::Nav::new(handle), selector).map(Matches::odt)
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => {
                crate::query::run(crate::format::odp::Nav::new(handle), selector).map(Matches::odp)
            },
            Self::Signed(signed) => signed.file().query(selector),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            other => Err(pending(other)),
        }
    }

    /// Render text, annotated text, outline or statistics using default options.
    ///
    /// # Errors
    /// Returns typed read/protection errors or `unsupported_capability` for a
    /// format or mode whose renderer is not yet integrated.
    pub fn view(&self, mode: Mode) -> Result<String> {
        let mut output = String::new();
        self.view_to(mode, &mut output)?;
        Ok(output)
    }

    /// Build a typed JSON view without serializing through an intermediate string.
    ///
    /// # Errors
    /// As [`Self::view`]. Serialize the returned value with any serde serializer.
    pub fn view_json(&self, mode: Mode) -> Result<crate::view::Json<'_>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                crate::format::doc::views::view_json(handle, mode).map(crate::view::Json::doc)
            },
            #[cfg(feature = "docx")]
            Self::Docx(handle) => crate::format::docx::views::view_json(
                handle,
                mode,
                &crate::view::mode::Options::default(),
            )
            .map(crate::view::Json::docx),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                crate::format::xlsx::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::xlsx)
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                crate::format::pptx::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::pptx)
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                crate::format::ppt::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::ppt)
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                crate::format::xlsb::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::xlsb)
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                crate::format::xls::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::xls)
            },
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                crate::format::rtf::views::view_json(handle, mode).map(crate::view::Json::rtf)
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                crate::format::ods::view_json(handle, mode, &crate::view::mode::Options::default())
                    .map(crate::view::Json::ods)
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) => {
                crate::format::odt::views::view_json(handle, mode).map(crate::view::Json::odt)
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => crate::format::odp::views::view_json(
                handle,
                mode,
                &crate::view::mode::Options::default(),
            )
            .map(crate::view::Json::odp),
            Self::Signed(signed) => signed.file().view_json(mode),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            other => Err(pending(other)),
        }
    }

    /// Stream a document view into a text sink. An error may leave partial output.
    ///
    /// # Errors
    /// As [`Self::view`], with `io_error` when the sink rejects output.
    pub fn view_to<W: std::fmt::Write>(&self, mode: Mode, output: &mut W) -> Result<()> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => crate::format::docx::views::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => crate::format::doc::views::view_to(handle, mode, output),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => crate::format::xlsx::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => crate::format::pptx::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => crate::format::ppt::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => crate::format::xlsb::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => crate::format::xls::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => crate::format::rtf::views::view_to(handle, mode, output),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => crate::format::odt::views::view_to(handle, mode, output),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => crate::format::ods::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => crate::format::odp::views::view_to(
                handle,
                mode,
                &crate::view::mode::Options::default(),
                output,
            ),
            Self::Signed(signed) => signed.file().view_to(mode, output),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            other => Err(pending(other)),
        }
    }
}

fn pending(file: &File) -> Error {
    file.capabilities().refusal()
}

/// The formats whose unified-read adapter has landed (record 0010 §10.2).
///
/// Matches on it are exhaustive, so an adapter added here without its arms
/// does not compile.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(crate) enum Adapter {
    /// Legacy binary Word (`.doc`).
    #[cfg(feature = "doc")]
    Doc,
    /// WordprocessingML (record 0051).
    #[cfg(feature = "docx")]
    Docx,
    /// SpreadsheetML (record 0055).
    #[cfg(feature = "xlsx")]
    Xlsx,
    #[cfg(feature = "pptx")]
    Pptx,
    #[cfg(feature = "ppt")]
    Ppt,
    /// OpenDocument Text (`.odt`, tier-3 adapter, record 0276).
    #[cfg(feature = "odt")]
    Odt,
    #[cfg(feature = "xlsb")]
    Xlsb,
    #[cfg(feature = "xls")]
    Xls,
    /// Rich Text Format (`.rtf`, tier-3 read adapter).
    #[cfg(feature = "rtf")]
    Rtf,
    /// OpenDocument Spreadsheet (`.ods`, tier-3 adapter, record 0277).
    #[cfg(feature = "ods")]
    Ods,
    /// OpenDocument Presentation (`.odp`, record 0278).
    #[cfg(feature = "odp")]
    Odp,
}

impl Adapter {
    /// The adapter of `format`, when it has landed and its cargo feature is
    /// on.
    pub(crate) const fn of(format: FileFormat) -> Option<Self> {
        match format {
            #[cfg(feature = "doc")]
            FileFormat::Doc => Some(Self::Doc),
            #[cfg(feature = "docx")]
            FileFormat::Docx => Some(Self::Docx),
            #[cfg(feature = "xlsx")]
            FileFormat::Xlsx => Some(Self::Xlsx),
            #[cfg(feature = "odt")]
            FileFormat::Odt => Some(Self::Odt),
            #[cfg(feature = "pptx")]
            FileFormat::Pptx => Some(Self::Pptx),
            #[cfg(feature = "ppt")]
            FileFormat::Ppt => Some(Self::Ppt),
            #[cfg(feature = "xlsb")]
            FileFormat::Xlsb => Some(Self::Xlsb),
            #[cfg(feature = "xls")]
            FileFormat::Xls => Some(Self::Xls),
            #[cfg(feature = "rtf")]
            FileFormat::Rtf => Some(Self::Rtf),
            #[cfg(feature = "ods")]
            FileFormat::Ods => Some(Self::Ods),
            #[cfg(feature = "odp")]
            FileFormat::Odp => Some(Self::Odp),
            _ => None,
        }
    }

    /// The format the adapter reads.
    pub(crate) const fn format(self) -> FileFormat {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc => FileFormat::Doc,
            #[cfg(feature = "docx")]
            Self::Docx => FileFormat::Docx,
            #[cfg(feature = "odt")]
            Self::Odt => FileFormat::Odt,
            #[cfg(feature = "xlsx")]
            Self::Xlsx => FileFormat::Xlsx,
            #[cfg(feature = "pptx")]
            Self::Pptx => FileFormat::Pptx,
            #[cfg(feature = "ppt")]
            Self::Ppt => FileFormat::Ppt,
            #[cfg(feature = "xlsb")]
            Self::Xlsb => FileFormat::Xlsb,
            #[cfg(feature = "xls")]
            Self::Xls => FileFormat::Xls,
            #[cfg(feature = "rtf")]
            Self::Rtf => FileFormat::Rtf,
            #[cfg(feature = "ods")]
            Self::Ods => FileFormat::Ods,
            #[cfg(feature = "odp")]
            Self::Odp => FileFormat::Odp,
        }
    }

    /// The format's family, which names its element vocabulary.
    pub(crate) const fn family(self) -> Family {
        Family::of(self.format())
    }
}

/// Runs `$body` with `$spec` bound to the `&'static FormatSpec<Kind>` of
/// `$adapter`: the one table its resolver, query compiler, node view and
/// capability report read (record 0010 §7.2). Static dispatch: one arm per
/// adapter, monomorphic in each.
macro_rules! with_spec {
    ($adapter:expr, $spec:ident => $body:expr) => {
        match $adapter {
            #[cfg(feature = "doc")]
            $crate::read::Adapter::Doc => {
                let $spec = &$crate::format::doc::SPEC;
                $body
            },
            #[cfg(feature = "docx")]
            $crate::read::Adapter::Docx => {
                let $spec = &$crate::format::docx::SPEC;
                $body
            },
            #[cfg(feature = "xlsx")]
            $crate::read::Adapter::Xlsx => {
                let $spec = &$crate::format::xlsx::SPEC;
                $body
            },
            #[cfg(feature = "pptx")]
            $crate::read::Adapter::Pptx => {
                let $spec = &$crate::format::pptx::SPEC;
                $body
            },
            #[cfg(feature = "ppt")]
            $crate::read::Adapter::Ppt => {
                let $spec = &$crate::format::ppt::SPEC;
                $body
            },
            #[cfg(feature = "xlsb")]
            $crate::read::Adapter::Xlsb => {
                let $spec = &$crate::format::xlsb::SPEC;
                $body
            },
            #[cfg(feature = "xls")]
            $crate::read::Adapter::Xls => {
                let $spec = &$crate::format::xls::SPEC;
                $body
            },
            #[cfg(feature = "rtf")]
            $crate::read::Adapter::Rtf => {
                let $spec = &$crate::format::rtf::SPEC;
                $body
            },
            #[cfg(feature = "ods")]
            $crate::read::Adapter::Ods => {
                let $spec = &$crate::format::ods::SPEC;
                $body
            },
            #[cfg(feature = "odt")]
            $crate::read::Adapter::Odt => {
                let $spec = &$crate::format::odt::SPEC;
                $body
            },
            #[cfg(feature = "odp")]
            $crate::read::Adapter::Odp => {
                let $spec = &$crate::format::odp::SPEC;
                $body
            },
        }
    };
}

pub(crate) use with_spec;

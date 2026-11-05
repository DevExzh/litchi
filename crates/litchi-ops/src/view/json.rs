//! A view as OfficeCLI's JSON (record 0010 §6.1): what
//! [`File::view_json`](crate::File::view_json) returns.

// In a build without any adapter, `Json` is uninhabited; those lints are
// allowed in such builds only. Each adapter adds its feature to this list.
#![cfg_attr(
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
        feature = "odp",
        odf_objects,
        feature = "iwork-adapter"
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

use std::convert::Infallible;
use std::fmt;
use std::marker::PhantomData;

use serde::ser::{Serialize, Serializer};

/// A typed, serializable document view returned by [`crate::File::view_json`].
/// Serialization preserves the format adapter's shape without an intermediate
/// JSON string or the CLI envelope. Debug output omits document content.
///
/// ```
/// # #[cfg(feature = "docx")]
/// # {
/// let file = litchi_ops::open(concat!(env!("CARGO_MANIFEST_DIR"),
///     "/../../test-data/ooxml/docx/drawing.docx"))?;
/// let view = file.view_json(litchi_ops::view::Mode::Stats)?;
/// let json = serde_json::to_value(&view).expect("serializes");
/// assert!(json.is_object());
/// # }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub struct Json<'f> {
    inner: Inner<'f>,
}

/// One variant per format adapter (static dispatch, ADR 0004).
enum Inner<'f> {
    #[cfg(feature = "doc")]
    Doc(crate::format::doc::views::Json),
    #[cfg(feature = "docx")]
    Docx(Box<crate::format::docx::views::Json>),
    #[cfg(feature = "xlsx")]
    Xlsx(crate::format::xlsx::ViewJson),
    #[cfg(feature = "pptx")]
    Pptx(crate::format::pptx::ViewJson),
    #[cfg(feature = "ppt")]
    Ppt(crate::format::ppt::ViewJson),
    #[cfg(feature = "xlsb")]
    Xlsb(crate::format::xlsb::ViewJson),
    #[cfg(feature = "xls")]
    Xls(crate::format::xls::ViewJson),
    #[cfg(feature = "rtf")]
    Rtf(crate::format::rtf::views::Json),
    #[cfg(feature = "ods")]
    Ods(crate::format::ods::ViewJson),
    #[cfg(feature = "odt")]
    Odt(crate::format::odt::views::Json),
    #[cfg(feature = "odp")]
    Odp(crate::format::odp::views::Json),
    #[cfg(odf_objects)]
    Object(crate::format::odf_tree::views::Json),
    #[cfg(feature = "iwork-adapter")]
    Iwork(crate::format::iwork::views::Json),
    /// Keeps `'f` in use; uninhabited.
    #[allow(dead_code, reason = "uninhabited: only keeps the lifetime in use")]
    Never(Infallible, PhantomData<&'f ()>),
}

impl Json<'_> {
    /// A DOC view.
    #[cfg(feature = "doc")]
    pub(crate) fn doc(json: crate::format::doc::views::Json) -> Self {
        Self {
            inner: Inner::Doc(json),
        }
    }

    /// An DOCX view.
    #[cfg(feature = "docx")]
    pub(crate) fn docx(json: crate::format::docx::views::Json) -> Self {
        Self {
            inner: Inner::Docx(Box::new(json)),
        }
    }

    /// An XLSX view.
    #[cfg(feature = "xlsx")]
    pub(crate) fn xlsx(json: crate::format::xlsx::ViewJson) -> Self {
        Self {
            inner: Inner::Xlsx(json),
        }
    }
    #[cfg(feature = "pptx")]
    pub(crate) fn pptx(json: crate::format::pptx::ViewJson) -> Self {
        Self {
            inner: Inner::Pptx(json),
        }
    }
    #[cfg(feature = "ppt")]
    pub(crate) fn ppt(json: crate::format::ppt::ViewJson) -> Self {
        Self {
            inner: Inner::Ppt(json),
        }
    }
    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb(json: crate::format::xlsb::ViewJson) -> Self {
        Self {
            inner: Inner::Xlsb(json),
        }
    }
    #[cfg(feature = "xls")]
    pub(crate) fn xls(json: crate::format::xls::ViewJson) -> Self {
        Self {
            inner: Inner::Xls(json),
        }
    }

    /// An ODT view.
    #[cfg(feature = "odt")]
    pub(crate) fn odt(json: crate::format::odt::views::Json) -> Self {
        Self {
            inner: Inner::Odt(json),
        }
    }

    /// An RTF view.
    #[cfg(feature = "rtf")]
    pub(crate) fn rtf(json: crate::format::rtf::views::Json) -> Self {
        Self {
            inner: Inner::Rtf(json),
        }
    }

    /// An ODS view.
    #[cfg(feature = "ods")]
    pub(crate) fn ods(json: crate::format::ods::ViewJson) -> Self {
        Self {
            inner: Inner::Ods(json),
        }
    }

    /// An ODP view.
    #[cfg(feature = "odp")]
    pub(crate) fn odp(json: crate::format::odp::views::Json) -> Self {
        Self {
            inner: Inner::Odp(json),
        }
    }

    /// An object-tree view (ODG, ODC, ODI, ODB, ODF formula).
    #[cfg(odf_objects)]
    pub(crate) fn object(json: crate::format::odf_tree::views::Json) -> Self {
        Self {
            inner: Inner::Object(json),
        }
    }

    /// An iWork object-tree view (Pages, Keynote, Numbers).
    #[cfg(feature = "iwork-adapter")]
    pub(crate) fn iwork(json: crate::format::iwork::views::Json) -> Self {
        Self {
            inner: Inner::Iwork(json),
        }
    }
}

impl Serialize for Json<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match &self.inner {
            #[cfg(feature = "doc")]
            Inner::Doc(json) => json.serialize(serializer),
            #[cfg(feature = "docx")]
            Inner::Docx(json) => json.serialize(serializer),
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(json) => json.serialize(serializer),
            #[cfg(feature = "pptx")]
            Inner::Pptx(json) => json.serialize(serializer),
            #[cfg(feature = "ppt")]
            Inner::Ppt(json) => json.serialize(serializer),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(json) => json.serialize(serializer),
            #[cfg(feature = "xls")]
            Inner::Xls(json) => json.serialize(serializer),
            #[cfg(feature = "rtf")]
            Inner::Rtf(json) => json.serialize(serializer),
            #[cfg(feature = "ods")]
            Inner::Ods(json) => json.serialize(serializer),
            #[cfg(feature = "odt")]
            Inner::Odt(json) => json.serialize(serializer),
            #[cfg(feature = "odp")]
            Inner::Odp(json) => json.serialize(serializer),
            #[cfg(odf_objects)]
            Inner::Object(json) => json.serialize(serializer),
            #[cfg(feature = "iwork-adapter")]
            Inner::Iwork(json) => json.serialize(serializer),
            Inner::Never(never, _) => match *never {},
        }
    }
}

impl fmt::Debug for Json<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Json { .. }")
    }
}

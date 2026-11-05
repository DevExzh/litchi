//! Typed budgets of the mail-merge family (record 0355a).

use litchi_ooxml_common::mce::Limits as MceLimits;
use litchi_opc::SourceLimits;

/// The bounded resources of the mail-merge family, carried by [`crate::Error::MailMergeLimit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Serialized or parsed mail-merge XML bytes.
    XmlBytes,
    /// Odso field maps.
    FieldMaps,
    /// Mail-merge recipients.
    Recipients,
    /// Mail-merge XML attributes on one node.
    XmlAttributes,
    /// Mail-merge XML nodes.
    XmlNodes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::XmlBytes => "mail-merge XML bytes",
            Self::FieldMaps => "odso field maps",
            Self::Recipients => "mail-merge recipients",
            Self::XmlAttributes => "mail-merge XML attributes on one node",
            Self::XmlNodes => "mail-merge XML nodes",
        }
    }
}

/// Derive the mail-merge MCE policy retained by an OPC package.
///
/// A detached caller supplies an MCE profile directly.  Package owners use
/// this adapter so a raw settings or recipient codec cannot widen the
/// physical part and XML-depth policy under which the package was admitted.
pub(crate) fn mce_limits_for_source(source: &SourceLimits) -> MceLimits {
    let mut limits = MceLimits::default();
    let part_bytes = usize::try_from(source.max_part_bytes()).unwrap_or(usize::MAX);
    limits.max_input_bytes = limits.max_input_bytes.min(part_bytes);
    limits.max_output_bytes = limits.max_output_bytes.min(part_bytes);
    limits.max_depth = limits.max_depth.min(source.max_depth());
    limits
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

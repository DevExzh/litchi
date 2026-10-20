//! The typed budgets of the MS-OVBA project codec (record 0360).

/// A MS-OVBA project codec budget that was exceeded.
///
/// Every limit error of this crate is
/// [`crate::Error::Limit`] with one of these kinds; its display is the
/// resource's human-readable name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The bound on standalone VBA CFB bytes.
    CfbBytes,
    /// The bound on VBA CFB stream bytes.
    CfbStreamBytes,
    /// The bound on compressed VBA stream bytes.
    CompressedStreamBytes,
    /// The bound on decompressed VBA stream bytes.
    DecompressedStreamBytes,
    /// The bound on VBA designer property count.
    DesignerProperties,
    /// The bound on VBA designer stream bytes.
    DesignerStreamBytes,
    /// The bound on VBA designer count.
    Designers,
    /// The bound on VBA input string bytes.
    InputStringBytes,
    /// The bound on VBA license key bytes.
    LicenseKeyBytes,
    /// The bound on PROJECTlk stream bytes.
    LicenseStreamBytes,
    /// The bound on VBA license count.
    Licenses,
    /// The bound on PROJECTwm module-name bytes.
    ModuleNameBytes,
    /// The bound on VBA module count.
    Modules,
    /// The bound on PROJECT path bytes.
    ProjectPathBytes,
    /// The bound on PROJECT stream bytes.
    ProjectStreamBytes,
    /// The bound on PROJECT text bytes.
    ProjectTextBytes,
    /// The bound on PROJECT text line bytes.
    ProjectTextLineBytes,
    /// The bound on VBA reference count.
    References,
    /// The bound on VBA string bytes.
    StringBytes,
    /// The bound on aggregate VBA module source bytes.
    TotalSourceBytes,
    /// The bound on PROJECTwm Unicode module-name bytes.
    UnicodeModuleNameBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::CfbBytes => "standalone VBA CFB bytes",
            Self::CfbStreamBytes => "VBA CFB stream bytes",
            Self::CompressedStreamBytes => "compressed VBA stream bytes",
            Self::DecompressedStreamBytes => "decompressed VBA stream bytes",
            Self::DesignerProperties => "VBA designer property count",
            Self::DesignerStreamBytes => "VBA designer stream bytes",
            Self::Designers => "VBA designer count",
            Self::InputStringBytes => "VBA input string bytes",
            Self::LicenseKeyBytes => "VBA license key bytes",
            Self::LicenseStreamBytes => "PROJECTlk stream bytes",
            Self::Licenses => "VBA license count",
            Self::ModuleNameBytes => "PROJECTwm module-name bytes",
            Self::Modules => "VBA module count",
            Self::ProjectPathBytes => "PROJECT path bytes",
            Self::ProjectStreamBytes => "PROJECT stream bytes",
            Self::ProjectTextBytes => "PROJECT text bytes",
            Self::ProjectTextLineBytes => "PROJECT text line bytes",
            Self::References => "VBA reference count",
            Self::StringBytes => "VBA string bytes",
            Self::TotalSourceBytes => "aggregate VBA module source bytes",
            Self::UnicodeModuleNameBytes => "PROJECTwm Unicode module-name bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

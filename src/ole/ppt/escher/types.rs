/// Escher record types.
///
/// Based on Microsoft Office Drawing specification and Apache POI implementation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum EscherRecordType {
    /// Unknown or unsupported record type
    Unknown = 0x0000,

    // Container records (0xF000 - 0xF00F)
    /// Drawing Group Container
    DggContainer = 0xF000,
    /// Blip Store Container
    BStoreContainer = 0xF001,
    /// Drawing Container
    DgContainer = 0xF002,
    /// Shape Group Container
    SpgrContainer = 0xF003,
    /// Shape Container
    SpContainer = 0xF004,
    /// Solver Container
    SolverContainer = 0xF005,

    // Atom records
    /// File Drawing Group atom
    Dgg = 0xF006,
    /// Blip Store Entry
    BSE = 0xF007,
    /// Drawing atom
    Dg = 0xF008,
    /// Shape Group atom
    Spgr = 0xF009,
    /// Shape atom
    Sp = 0xF00A,
    /// Shape Options
    Opt = 0xF00B,
    /// Client Anchor
    ClientAnchor = 0xF010,
    /// Client Data
    ClientData = 0xF011,
    /// Client Textbox (contains text)
    ClientTextbox = 0xF00D,
    /// Child Anchor
    ChildAnchor = 0xF00F,

    // Blip records
    /// JPEG Blip
    BlipJpeg = 0xF01D,
    /// PNG Blip
    BlipPng = 0xF01E,
    /// DIB Blip
    BlipDib = 0xF01F,
    /// TIFF Blip
    BlipTiff = 0xF029,
    /// EMF Blip
    BlipEmf = 0xF01A,
    /// WMF Blip
    BlipWmf = 0xF01B,
    /// PICT Blip
    BlipPict = 0xF01C,

    // Text records
    /// Secondary Opt (Shape Options)
    SecondaryOpt = 0xF121,
    /// Tertiary Opt
    TertiaryOpt = 0xF122,

    // Split menu colors
    /// Split Menu Colors
    SplitMenuColors = 0xF11E,

    // Color MRU
    /// Color MRU
    ColorMRU = 0xF11A,

    // Connector rule
    /// Connector Rule
    ConnectorRule = 0xF012,
    /// Align Rule
    AlignRule = 0xF013,
    /// Arc Rule
    ArcRule = 0xF014,
    /// Client Rule
    ClientRule = 0xF015,
    /// Callout Rule
    CalloutRule = 0xF017,
}

impl EscherRecordType {
    /// Check if this is a container record type.
    ///
    /// Container records have version field 0xF (15) and can contain child records.
    #[inline]
    pub const fn is_container(self) -> bool {
        matches!(
            self,
            Self::DggContainer
                | Self::BStoreContainer
                | Self::DgContainer
                | Self::SpgrContainer
                | Self::SpContainer
                | Self::SolverContainer
        )
    }

    /// Check if this record type can contain text.
    #[inline]
    pub const fn can_contain_text(self) -> bool {
        matches!(self, Self::ClientTextbox | Self::SpContainer)
    }

    /// Check if this is a BLIP (image) record type.
    #[inline]
    pub const fn is_blip(self) -> bool {
        matches!(
            self,
            Self::BlipEmf
                | Self::BlipWmf
                | Self::BlipPict
                | Self::BlipJpeg
                | Self::BlipPng
                | Self::BlipDib
                | Self::BlipTiff
        )
    }
}

impl From<u16> for EscherRecordType {
    fn from(value: u16) -> Self {
        match value {
            0xF000 => Self::DggContainer,
            0xF001 => Self::BStoreContainer,
            0xF002 => Self::DgContainer,
            0xF003 => Self::SpgrContainer,
            0xF004 => Self::SpContainer,
            0xF005 => Self::SolverContainer,
            0xF006 => Self::Dgg,
            0xF007 => Self::BSE,
            0xF008 => Self::Dg,
            0xF009 => Self::Spgr,
            0xF00A => Self::Sp,
            0xF00B => Self::Opt,
            0xF00D => Self::ClientTextbox,
            0xF00F => Self::ChildAnchor,
            0xF010 => Self::ClientAnchor,
            0xF011 => Self::ClientData,
            0xF012 => Self::ConnectorRule,
            0xF013 => Self::AlignRule,
            0xF014 => Self::ArcRule,
            0xF015 => Self::ClientRule,
            0xF017 => Self::CalloutRule,
            0xF01A => Self::BlipEmf,
            0xF01B => Self::BlipWmf,
            0xF01C => Self::BlipPict,
            0xF01D => Self::BlipJpeg,
            0xF01E => Self::BlipPng,
            0xF01F => Self::BlipDib,
            0xF029 => Self::BlipTiff,
            0xF11A => Self::ColorMRU,
            0xF11E => Self::SplitMenuColors,
            0xF121 => Self::SecondaryOpt,
            0xF122 => Self::TertiaryOpt,
            _ => Self::Unknown,
        }
    }
}

impl From<EscherRecordType> for u16 {
    fn from(record_type: EscherRecordType) -> Self {
        record_type as u16
    }
}

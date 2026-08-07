//! Checked, zero-copy `OfficeArt` record parsing.

use crate::{Error, Result};

/// The semantic kind encoded by an `OfficeArt` record header.
///
/// Unknown values retain their original numeric representation so callers can
/// inspect or round-trip records introduced by newer producers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RecordKind {
    /// Drawing-group container.
    DggContainer,
    /// BLIP-store container.
    BStoreContainer,
    /// Drawing container.
    DgContainer,
    /// Shape-group container.
    SpgrContainer,
    /// Shape container.
    SpContainer,
    /// Solver container.
    SolverContainer,
    /// File drawing-group atom.
    Dgg,
    /// BLIP-store entry.
    Bse,
    /// Drawing atom.
    Dg,
    /// Shape-group atom.
    Spgr,
    /// Shape atom.
    Sp,
    /// Primary shape-options atom.
    Opt,
    /// Client textbox atom.
    ClientTextbox,
    /// Child anchor atom.
    ChildAnchor,
    /// Client anchor atom.
    ClientAnchor,
    /// Client data atom.
    ClientData,
    /// Connector rule.
    ConnectorRule,
    /// Alignment rule.
    AlignRule,
    /// Arc rule.
    ArcRule,
    /// Client rule.
    ClientRule,
    /// Callout rule.
    CalloutRule,
    /// Enhanced Metafile BLIP.
    BlipEmf,
    /// Windows Metafile BLIP.
    BlipWmf,
    /// Macintosh PICT BLIP.
    BlipPict,
    /// JPEG BLIP.
    BlipJpeg,
    /// PNG BLIP.
    BlipPng,
    /// Device-independent bitmap BLIP.
    BlipDib,
    /// TIFF BLIP.
    BlipTiff,
    /// Most-recently-used color atom.
    ColorMru,
    /// Split-menu colors atom.
    SplitMenuColors,
    /// Secondary shape-options atom.
    SecondaryOpt,
    /// Tertiary shape-options atom.
    TertiaryOpt,
    /// An unrecognized record kind, retaining its wire value.
    Unknown(u16),
}

impl RecordKind {
    /// Converts a wire value into a lossless semantic kind.
    #[must_use]
    pub const fn from_raw(raw: u16) -> Self {
        match raw {
            0xF000 => Self::DggContainer,
            0xF001 => Self::BStoreContainer,
            0xF002 => Self::DgContainer,
            0xF003 => Self::SpgrContainer,
            0xF004 => Self::SpContainer,
            0xF005 => Self::SolverContainer,
            0xF006 => Self::Dgg,
            0xF007 => Self::Bse,
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
            0xF01D | 0xF02A => Self::BlipJpeg,
            0xF01E => Self::BlipPng,
            0xF01F => Self::BlipDib,
            0xF029 => Self::BlipTiff,
            0xF11A => Self::ColorMru,
            0xF11E => Self::SplitMenuColors,
            0xF121 => Self::SecondaryOpt,
            0xF122 => Self::TertiaryOpt,
            value => Self::Unknown(value),
        }
    }

    /// Returns the canonical wire value for this kind.
    #[must_use]
    pub const fn raw(self) -> u16 {
        match self {
            Self::DggContainer => 0xF000,
            Self::BStoreContainer => 0xF001,
            Self::DgContainer => 0xF002,
            Self::SpgrContainer => 0xF003,
            Self::SpContainer => 0xF004,
            Self::SolverContainer => 0xF005,
            Self::Dgg => 0xF006,
            Self::Bse => 0xF007,
            Self::Dg => 0xF008,
            Self::Spgr => 0xF009,
            Self::Sp => 0xF00A,
            Self::Opt => 0xF00B,
            Self::ClientTextbox => 0xF00D,
            Self::ChildAnchor => 0xF00F,
            Self::ClientAnchor => 0xF010,
            Self::ClientData => 0xF011,
            Self::ConnectorRule => 0xF012,
            Self::AlignRule => 0xF013,
            Self::ArcRule => 0xF014,
            Self::ClientRule => 0xF015,
            Self::CalloutRule => 0xF017,
            Self::BlipEmf => 0xF01A,
            Self::BlipWmf => 0xF01B,
            Self::BlipPict => 0xF01C,
            Self::BlipJpeg => 0xF01D,
            Self::BlipPng => 0xF01E,
            Self::BlipDib => 0xF01F,
            Self::BlipTiff => 0xF029,
            Self::ColorMru => 0xF11A,
            Self::SplitMenuColors => 0xF11E,
            Self::SecondaryOpt => 0xF121,
            Self::TertiaryOpt => 0xF122,
            Self::Unknown(raw) => raw,
        }
    }

    /// Returns whether records of this kind contain child records.
    #[must_use]
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

    /// Returns whether this kind can carry text-related records.
    #[must_use]
    pub const fn can_contain_text(self) -> bool {
        matches!(self, Self::ClientTextbox | Self::SpContainer)
    }

    /// Returns whether this kind stores an image BLIP.
    #[must_use]
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
        ) || matches!(self, Self::Unknown(raw) if raw >= 0xF018 && raw <= 0xF117)
    }
}

impl From<u16> for RecordKind {
    fn from(raw: u16) -> Self {
        Self::from_raw(raw)
    }
}

impl From<RecordKind> for u16 {
    fn from(kind: RecordKind) -> Self {
        kind.raw()
    }
}

/// A checked, zero-copy view of one `OfficeArt` record.
#[derive(Clone)]
pub struct Record<'data> {
    kind: RecordKind,
    raw_kind: u16,
    version: u8,
    instance: u16,
    len: u32,
    data: &'data [u8],
}

impl core::fmt::Debug for Record<'_> {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        formatter
            .debug_struct("Record")
            .field("kind", &self.kind)
            .field("raw_kind", &self.raw_kind)
            .field("version", &self.version)
            .field("instance", &self.instance)
            .field("len", &self.len)
            .finish_non_exhaustive()
    }
}

impl<'data> Record<'data> {
    #[cfg(test)]
    pub(crate) fn from_parts(
        kind: RecordKind,
        version: u8,
        instance: u16,
        data: &'data [u8],
    ) -> Result<Self> {
        let len = u32::try_from(data.len()).map_err(|_err| Error::ArithmeticOverflow {
            context: "test record body length",
        })?;
        Ok(Self {
            kind,
            raw_kind: kind.raw(),
            version: version & 0x0F,
            instance: instance & 0x0FFF,
            len,
            data,
        })
    }

    /// Parses a record beginning at `offset` and returns it with bytes consumed.
    ///
    /// The declared `recLen` is authoritative. Truncated atoms, containers, and
    /// BLIPs are all rejected rather than silently shortened.
    ///
    /// # Errors
    ///
    /// Returns `Error::ArithmeticOverflow` if the offset or length arithmetic
    /// cannot be represented, `Error::TruncatedHeader` if fewer than eight
    /// header bytes remain at `offset`, or `Error::TruncatedPayload` if the
    /// declared `recLen` extends past the supplied slice.
    pub fn parse(data: &'data [u8], offset: usize) -> Result<(Self, usize)> {
        let header_end = offset.checked_add(8).ok_or(Error::ArithmeticOverflow {
            context: "record header extent",
        })?;
        let header = data.get(offset..header_end).ok_or(Error::TruncatedHeader {
            offset,
            available: data.len().saturating_sub(offset),
        })?;

        let ver_inst = u16::from_le_bytes([header[0], header[1]]);
        let raw_kind = u16::from_le_bytes([header[2], header[3]]);
        let len = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
        let body_len = usize::try_from(len).map_err(|_err| Error::ArithmeticOverflow {
            context: "record body length",
        })?;
        let body_end = header_end
            .checked_add(body_len)
            .ok_or(Error::ArithmeticOverflow {
                context: "record body extent",
            })?;
        let body = data
            .get(header_end..body_end)
            .ok_or(Error::TruncatedPayload {
                offset,
                declared: len,
                available: data.len().saturating_sub(header_end),
            })?;
        let consumed = body_end
            .checked_sub(offset)
            .ok_or(Error::ArithmeticOverflow {
                context: "record consumed length",
            })?;

        Ok((
            Self {
                kind: RecordKind::from_raw(raw_kind),
                raw_kind,
                version: (ver_inst & 0x000F) as u8,
                instance: (ver_inst >> 4) & 0x0FFF,
                len,
                data: body,
            },
            consumed,
        ))
    }

    /// Returns the semantic record kind.
    #[must_use]
    pub const fn kind(&self) -> RecordKind {
        self.kind
    }

    /// Returns the exact record-kind value read from the wire.
    #[must_use]
    pub const fn raw_kind(&self) -> u16 {
        self.raw_kind
    }

    /// Returns the four-bit record version.
    #[must_use]
    pub const fn version(&self) -> u8 {
        self.version
    }

    /// Returns the twelve-bit record instance.
    #[must_use]
    pub const fn instance(&self) -> u16 {
        self.instance
    }

    /// Returns the declared body length.
    #[must_use]
    pub const fn len(&self) -> u32 {
        self.len
    }

    /// Returns whether the record body is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Returns the borrowed body bytes.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.data
    }

    /// Returns whether this record can contain child records.
    #[must_use]
    pub const fn is_container(&self) -> bool {
        self.version == 0x0F
    }

    /// Returns whether this record kind can contain text.
    #[must_use]
    pub const fn can_contain_text(&self) -> bool {
        self.kind.can_contain_text()
    }

    /// Returns the body offset when this record borrows from `parent`.
    #[must_use]
    pub fn data_offset(&self, parent: &[u8]) -> Option<usize> {
        let offset = (self.data.as_ptr() as usize).checked_sub(parent.as_ptr() as usize)?;
        let end = offset.checked_add(self.data.len())?;
        (end <= parent.len()).then_some(offset)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    #[test]
    fn parses_without_copying() {
        let bytes = [
            0x02, 0x00, 0x0A, 0xF0, 0x04, 0x00, 0x00, 0x00, 0xAA, 0xBB, 0xCC, 0xDD,
        ];
        let (record, consumed) = Record::parse(&bytes, 0).expect("valid record");

        assert_eq!(record.kind(), RecordKind::Sp);
        assert_eq!(record.raw_kind(), 0xF00A);
        assert_eq!(record.data(), &bytes[8..]);
        assert_eq!(record.data_offset(&bytes), Some(8));
        assert_eq!(consumed, bytes.len());
    }

    #[test]
    fn preserves_unknown_kinds() {
        let bytes = [0x00, 0x00, 0x34, 0x12, 0, 0, 0, 0];
        let (record, _) = Record::parse(&bytes, 0).expect("valid record");

        assert_eq!(record.kind(), RecordKind::Unknown(0x1234));
        assert_eq!(record.raw_kind(), 0x1234);
    }

    #[test]
    fn rejects_truncated_container_and_blip() {
        for raw_kind in [0xF004_u16, 0xF01E] {
            let mut bytes = vec![0x0F, 0x00];
            bytes.extend_from_slice(&raw_kind.to_le_bytes());
            bytes.extend_from_slice(&8_u32.to_le_bytes());
            bytes.extend_from_slice(&[1, 2]);

            assert!(matches!(
                Record::parse(&bytes, 0),
                Err(Error::TruncatedPayload { .. })
            ));
        }
    }

    #[test]
    fn rejects_offset_overflow() {
        assert!(matches!(
            Record::parse(&[], usize::MAX),
            Err(Error::ArithmeticOverflow { .. })
        ));
    }

    #[test]
    fn container_proof_uses_the_header_version() {
        let wrong_version = [0x00, 0x00, 0x04, 0xF0, 0, 0, 0, 0];
        let unknown_container = [0x0F, 0x00, 0x34, 0x12, 0, 0, 0, 0];

        assert!(
            !Record::parse(&wrong_version, 0)
                .expect("valid record")
                .0
                .is_container()
        );
        assert!(
            Record::parse(&unknown_container, 0)
                .expect("valid record")
                .0
                .is_container()
        );
    }
}

//! Typed, lossless Word 2003 Document Properties extension.

use super::document_properties_97::DopExtensionError;
use super::document_properties_2002::Dop2002;

const DOP2002_SIZE: usize = 594;
const DOP2003_SIZE: usize = 616;
const EXTENSION_SIZE: usize = DOP2003_SIZE - DOP2002_SIZE;
const COMPATIBILITY_FLAGS: usize = 0;
const PROTECTION_FLAGS: usize = 4;
const PAGE_LOCK_WIDTH: usize = 6;
const PAGE_LOCK_HEIGHT: usize = 10;
const PAGE_LOCK_FONT_PERCENT: usize = 14;
const STATE_TOOLBARS: usize = 18;
const RESERVED_BYTE: usize = 19;
const LIST_CLEANUP_LIMIT: usize = 20;
const UNUSED_COMPATIBILITY_BIT: u32 = 1 << 9;
const EMPTY_COMPATIBILITY_MASK: u32 = 0xffff_e000;
const EMPTY_PROTECTION_MASK: u16 = 0xff00;
const KNOWN_TOOLBAR_MASK: u8 = 0x07;

/// Editing restriction selected when document protection is enforced.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentProtectionMode {
    TrackedChanges,
    CommentsAndPermissionRanges,
    FormFields,
    PermissionRanges,
    Unrestricted,
}

impl DocumentProtectionMode {
    fn parse(raw: u16) -> Result<Self, DopExtensionError> {
        match raw {
            0 => Ok(Self::TrackedChanges),
            1 => Ok(Self::CommentsAndPermissionRanges),
            2 => Ok(Self::FormFields),
            3 => Ok(Self::PermissionRanges),
            7 => Ok(Self::Unrestricted),
            _ => Err(DopExtensionError::new(format!(
                "invalid Dop2003 document protection mode {raw}"
            ))),
        }
    }

    const fn raw(self) -> u16 {
        match self {
            Self::TrackedChanges => 0,
            Self::CommentsAndPermissionRanges => 1,
            Self::FormFields => 2,
            Self::PermissionRanges => 3,
            Self::Unrestricted => 7,
        }
    }
}

/// Virtual-page geometry used by reading-mode ink lockdown.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ReadingModePageLock {
    pub width_twips: u32,
    pub height_twips: u32,
    pub font_scale_percent: u32,
}

/// Toolbars that Word displayed because of document state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentStateToolbars(u8);

impl DocumentStateToolbars {
    pub const REVIEWING: u8 = 0x01;
    pub const WEB: u8 = 0x02;
    pub const MAIL_MERGE: u8 = 0x04;

    pub fn from_raw(raw: u8) -> Result<Self, DopExtensionError> {
        if raw & !KNOWN_TOOLBAR_MASK != 0 {
            Err(DopExtensionError::new(format!(
                "invalid Dop2003 document-state toolbar mask {raw:#04x}"
            )))
        } else {
            Ok(Self(raw))
        }
    }

    pub const fn raw(self) -> u8 {
        self.0
    }

    pub const fn contains(self, toolbar: u8) -> bool {
        self.0 & toolbar != 0
    }
}

/// Typed, lossless Word 2003 DOP extension.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dop2003 {
    raw: [u8; EXTENSION_SIZE],
    pub treat_comment_lock_as_read_only: bool,
    pub style_lock: bool,
    pub auto_format_override: bool,
    pub save_custom_xml_data_only: bool,
    pub apply_custom_xml_transform: bool,
    pub style_lock_enforced: bool,
    pub legacy_comment_lock_fallback: bool,
    pub ignore_mixed_xml_content: bool,
    pub show_xml_placeholder_text: bool,
    pub word_97_compatibility_ui: bool,
    pub lock_document_theme: bool,
    pub lock_quick_format_style_set: bool,
    pub reading_mode_ink_lockdown: bool,
    pub show_ink_annotations: bool,
    pub remove_annotation_date_time: bool,
    pub enforce_document_protection: bool,
    pub document_protection_mode: DocumentProtectionMode,
    pub display_background_shapes: bool,
    pub reading_mode_page_lock: ReadingModePageLock,
    pub document_state_toolbars: DocumentStateToolbars,
    pub largest_list_id_for_cleanup: u16,
}

impl Dop2003 {
    pub fn parse(dop: &[u8]) -> Result<Self, DopExtensionError> {
        if dop.len() < DOP2003_SIZE {
            return Err(DopExtensionError::new("Dop2003 is shorter than 616 bytes"));
        }
        Dop2002::parse(dop)?;
        let extension = &dop[DOP2002_SIZE..DOP2003_SIZE];
        let compatibility = le_u32(extension, COMPATIBILITY_FLAGS);
        let protection = le_u16(extension, PROTECTION_FLAGS);
        if compatibility & EMPTY_COMPATIBILITY_MASK != 0 {
            return Err(DopExtensionError::new(
                "Dop2003 empty compatibility bits must be zero",
            ));
        }
        if protection & EMPTY_PROTECTION_MASK != 0 {
            return Err(DopExtensionError::new(
                "Dop2003 empty protection bits must be zero",
            ));
        }
        if extension[RESERVED_BYTE] != 0 {
            return Err(DopExtensionError::new(
                "Dop2003 reserved toolbar byte must be zero",
            ));
        }
        let style_lock = compatibility & (1 << 1) != 0;
        let style_lock_enforced = compatibility & (1 << 5) != 0;
        if style_lock_enforced && !style_lock {
            return Err(DopExtensionError::new(
                "Dop2003 enforced style lock requires style lock",
            ));
        }
        let mut raw = [0u8; EXTENSION_SIZE];
        raw.copy_from_slice(extension);
        Ok(Self {
            raw,
            treat_comment_lock_as_read_only: compatibility & 1 != 0,
            style_lock,
            auto_format_override: compatibility & (1 << 2) != 0,
            save_custom_xml_data_only: compatibility & (1 << 3) != 0,
            apply_custom_xml_transform: compatibility & (1 << 4) != 0,
            style_lock_enforced,
            legacy_comment_lock_fallback: compatibility & (1 << 6) != 0,
            ignore_mixed_xml_content: compatibility & (1 << 7) != 0,
            show_xml_placeholder_text: compatibility & (1 << 8) != 0,
            word_97_compatibility_ui: compatibility & (1 << 10) != 0,
            lock_document_theme: compatibility & (1 << 11) != 0,
            lock_quick_format_style_set: compatibility & (1 << 12) != 0,
            reading_mode_ink_lockdown: protection & 1 != 0,
            show_ink_annotations: protection & (1 << 1) != 0,
            remove_annotation_date_time: protection & (1 << 2) != 0,
            enforce_document_protection: protection & (1 << 3) != 0,
            document_protection_mode: DocumentProtectionMode::parse((protection >> 4) & 7)?,
            display_background_shapes: protection & (1 << 7) != 0,
            reading_mode_page_lock: ReadingModePageLock {
                width_twips: le_u32(extension, PAGE_LOCK_WIDTH),
                height_twips: le_u32(extension, PAGE_LOCK_HEIGHT),
                font_scale_percent: le_u32(extension, PAGE_LOCK_FONT_PERCENT),
            },
            document_state_toolbars: DocumentStateToolbars::from_raw(extension[STATE_TOOLBARS])?,
            largest_list_id_for_cleanup: le_u16(extension, LIST_CLEANUP_LIMIT),
        })
    }

    pub fn write_into(mut self, dop: &mut [u8]) -> Result<(), DopExtensionError> {
        if dop.len() < DOP2003_SIZE {
            return Err(DopExtensionError::new(
                "Dop2003 target is shorter than 616 bytes",
            ));
        }
        if self.style_lock_enforced && !self.style_lock {
            return Err(DopExtensionError::new(
                "Dop2003 enforced style lock requires style lock",
            ));
        }
        DocumentStateToolbars::from_raw(self.document_state_toolbars.raw())?;
        let mut compatibility = le_u32(&self.raw, COMPATIBILITY_FLAGS) & UNUSED_COMPATIBILITY_BIT;
        compatibility |= u32::from(self.treat_comment_lock_as_read_only);
        compatibility |= u32::from(self.style_lock) << 1;
        compatibility |= u32::from(self.auto_format_override) << 2;
        compatibility |= u32::from(self.save_custom_xml_data_only) << 3;
        compatibility |= u32::from(self.apply_custom_xml_transform) << 4;
        compatibility |= u32::from(self.style_lock_enforced) << 5;
        compatibility |= u32::from(self.legacy_comment_lock_fallback) << 6;
        compatibility |= u32::from(self.ignore_mixed_xml_content) << 7;
        compatibility |= u32::from(self.show_xml_placeholder_text) << 8;
        compatibility |= u32::from(self.word_97_compatibility_ui) << 10;
        compatibility |= u32::from(self.lock_document_theme) << 11;
        compatibility |= u32::from(self.lock_quick_format_style_set) << 12;
        let mut protection = u16::from(self.reading_mode_ink_lockdown);
        protection |= u16::from(self.show_ink_annotations) << 1;
        protection |= u16::from(self.remove_annotation_date_time) << 2;
        protection |= u16::from(self.enforce_document_protection) << 3;
        protection |= self.document_protection_mode.raw() << 4;
        protection |= u16::from(self.display_background_shapes) << 7;
        put_u32(&mut self.raw, COMPATIBILITY_FLAGS, compatibility);
        put_u16(&mut self.raw, PROTECTION_FLAGS, protection);
        put_u32(
            &mut self.raw,
            PAGE_LOCK_WIDTH,
            self.reading_mode_page_lock.width_twips,
        );
        put_u32(
            &mut self.raw,
            PAGE_LOCK_HEIGHT,
            self.reading_mode_page_lock.height_twips,
        );
        put_u32(
            &mut self.raw,
            PAGE_LOCK_FONT_PERCENT,
            self.reading_mode_page_lock.font_scale_percent,
        );
        self.raw[STATE_TOOLBARS] = self.document_state_toolbars.raw();
        self.raw[RESERVED_BYTE] = 0;
        put_u16(
            &mut self.raw,
            LIST_CLEANUP_LIMIT,
            self.largest_list_id_for_cleanup,
        );
        dop[DOP2002_SIZE..DOP2003_SIZE].copy_from_slice(&self.raw);
        Ok(())
    }
}

fn le_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn le_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        data[offset..offset + 4]
            .try_into()
            .expect("fixed-width slice"),
    )
}

fn put_u16(data: &mut [u8], offset: usize, value: u16) {
    data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn put_u32(data: &mut [u8], offset: usize, value: u32) {
    data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_dop2003() -> Vec<u8> {
        let mut dop = vec![0u8; DOP2003_SIZE];
        dop[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
        dop
    }

    #[test]
    fn parses_complete_typed_extension() {
        let mut dop = valid_dop2003();
        put_u32(
            &mut dop[DOP2002_SIZE..],
            COMPATIBILITY_FLAGS,
            1 | (1 << 1) | (1 << 5) | (1 << 9) | (1 << 12),
        );
        put_u16(
            &mut dop[DOP2002_SIZE..],
            PROTECTION_FLAGS,
            1 | (1 << 1) | (3 << 4) | (1 << 7),
        );
        put_u32(&mut dop[DOP2002_SIZE..], PAGE_LOCK_WIDTH, 12_240);
        put_u32(&mut dop[DOP2002_SIZE..], PAGE_LOCK_HEIGHT, 15_840);
        put_u32(&mut dop[DOP2002_SIZE..], PAGE_LOCK_FONT_PERCENT, 125);
        dop[DOP2002_SIZE + STATE_TOOLBARS] =
            DocumentStateToolbars::REVIEWING | DocumentStateToolbars::MAIL_MERGE;
        put_u16(&mut dop[DOP2002_SIZE..], LIST_CLEANUP_LIMIT, 42);

        let value = Dop2003::parse(&dop).unwrap();
        assert!(value.style_lock_enforced);
        assert_eq!(
            value.document_protection_mode,
            DocumentProtectionMode::PermissionRanges
        );
        assert_eq!(value.reading_mode_page_lock.font_scale_percent, 125);
        assert!(
            value
                .document_state_toolbars
                .contains(DocumentStateToolbars::MAIL_MERGE)
        );
        assert_eq!(value.largest_list_id_for_cleanup, 42);
    }

    #[test]
    fn round_trip_preserves_the_undefined_bit() {
        let mut dop = valid_dop2003();
        put_u32(
            &mut dop[DOP2002_SIZE..],
            COMPATIBILITY_FLAGS,
            UNUSED_COMPATIBILITY_BIT,
        );
        let parsed = Dop2003::parse(&dop).unwrap();
        let mut output = dop.clone();
        parsed.write_into(&mut output).unwrap();
        assert_eq!(output, dop);
    }

    #[test]
    fn rejects_invalid_invariants_and_reserved_values() {
        let mut lock = valid_dop2003();
        put_u32(&mut lock[DOP2002_SIZE..], COMPATIBILITY_FLAGS, 1 << 5);
        assert!(Dop2003::parse(&lock).is_err());

        let mut mode = valid_dop2003();
        put_u16(&mut mode[DOP2002_SIZE..], PROTECTION_FLAGS, 4 << 4);
        assert!(Dop2003::parse(&mode).is_err());

        let mut empty = valid_dop2003();
        put_u32(&mut empty[DOP2002_SIZE..], COMPATIBILITY_FLAGS, 1 << 13);
        assert!(Dop2003::parse(&empty).is_err());

        let mut toolbar = valid_dop2003();
        toolbar[DOP2002_SIZE + STATE_TOOLBARS] = 0x08;
        assert!(Dop2003::parse(&toolbar).is_err());
    }

    #[test]
    fn writes_typed_mutations_and_checks_target_size() {
        let dop = valid_dop2003();
        let mut value = Dop2003::parse(&dop).unwrap();
        value.reading_mode_ink_lockdown = true;
        value.reading_mode_page_lock.width_twips = 9_000;
        value.document_protection_mode = DocumentProtectionMode::Unrestricted;
        let mut output = dop.clone();
        value.write_into(&mut output).unwrap();
        let reparsed = Dop2003::parse(&output).unwrap();
        assert!(reparsed.reading_mode_ink_lockdown);
        assert_eq!(reparsed.reading_mode_page_lock.width_twips, 9_000);
        assert_eq!(
            reparsed.document_protection_mode,
            DocumentProtectionMode::Unrestricted
        );
        assert!(Dop2003::parse(&output[..DOP2003_SIZE - 1]).is_err());
    }
}

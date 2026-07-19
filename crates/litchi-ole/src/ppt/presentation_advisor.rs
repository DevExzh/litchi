//! Presentation Advisor preferences from MS-PPT 2.4.6.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::records::PptRecord;

const PRESENTATION_ADVISOR_RECORD_TYPE: u16 = 0x177a;
const ADVISOR_RULE_MASK: u32 = 0x07ff;

/// A presentation-style rule that PowerPoint can suppress warnings for.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum PowerPointAdvisorRule {
    CaseStyleTitle = 0,
    CaseStyleBody = 1,
    EndPunctuationTitle = 2,
    EndPunctuationBody = 3,
    TooManyBullets = 4,
    FontSizeTitle = 5,
    FontSizeBody = 6,
    NumberOfLinesTitle = 7,
    NumberOfLinesBody = 8,
    TooManyFonts = 9,
    PrintTip = 10,
}

impl PowerPointAdvisorRule {
    /// Every MS-PPT 2.4.6 rule in bit order.
    pub const ALL: [Self; 11] = [
        Self::CaseStyleTitle,
        Self::CaseStyleBody,
        Self::EndPunctuationTitle,
        Self::EndPunctuationBody,
        Self::TooManyBullets,
        Self::FontSizeTitle,
        Self::FontSizeBody,
        Self::NumberOfLinesTitle,
        Self::NumberOfLinesBody,
        Self::TooManyFonts,
        Self::PrintTip,
    ];

    const fn bit(self) -> u32 {
        1 << self as u8
    }
}

/// Compact Presentation Advisor suppression settings.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointPresentationAdvisorSettings {
    disabled_mask: u16,
}

impl PowerPointPresentationAdvisorSettings {
    /// Create settings with every advisor rule enabled.
    pub const fn new() -> Self {
        Self { disabled_mask: 0 }
    }

    /// Parse one strict `PresAdvisorFlags9Atom`.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type_raw != PRESENTATION_ADVISOR_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "PresAdvisorFlags9Atom has an invalid record header or size".to_string(),
            ));
        }
        let flags = u32::from_le_bytes(record.data[0..4].try_into().unwrap());
        if flags & !ADVISOR_RULE_MASK != 0 {
            return Err(PptError::Corrupted(
                "PresAdvisorFlags9Atom has nonzero reserved bits".to_string(),
            ));
        }
        Ok(Self {
            disabled_mask: flags as u16,
        })
    }

    /// Discover the single advisor atom in the PPT9 document tag.
    pub(crate) fn parse_document(document: &PptRecord) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == PRESENTATION_ADVISOR_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(PptError::Corrupted(
                "PPT9 document tag contains multiple PresAdvisorFlags9Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Return whether warnings for `rule` are disabled.
    pub const fn is_disabled(self, rule: PowerPointAdvisorRule) -> bool {
        self.disabled_mask as u32 & rule.bit() != 0
    }

    /// Disable warnings for `rule`.
    pub fn disable(&mut self, rule: PowerPointAdvisorRule) {
        self.disabled_mask |= rule.bit() as u16;
    }

    /// Enable warnings for `rule`.
    pub fn enable(&mut self, rule: PowerPointAdvisorRule) {
        self.disabled_mask &= !(rule.bit() as u16);
    }

    /// Iterate over disabled rules in specification bit order.
    pub fn disabled_rules(self) -> impl Iterator<Item = PowerPointAdvisorRule> {
        PowerPointAdvisorRule::ALL
            .into_iter()
            .filter(move |rule| self.is_disabled(*rule))
    }

    /// Encode the exact PowerPoint 9 atom.
    pub fn to_record(self) -> PptRecord {
        let flags = u32::from(self.disabled_mask);
        PptRecord {
            record_type: PptRecordType::from(PRESENTATION_ADVISOR_RECORD_TYPE),
            record_type_raw: PRESENTATION_ADVISOR_RECORD_TYPE,
            version: 0,
            instance: 0,
            data_length: 4,
            data: flags.to_le_bytes().to_vec(),
            children: Vec::new(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_every_typed_advisor_rule() {
        let mut settings = PowerPointPresentationAdvisorSettings::new();
        for rule in PowerPointAdvisorRule::ALL {
            settings.disable(rule);
        }
        let parsed = PowerPointPresentationAdvisorSettings::parse(&settings.to_record()).unwrap();
        assert_eq!(
            parsed.disabled_rules().collect::<Vec<_>>(),
            PowerPointAdvisorRule::ALL
        );
        for rule in PowerPointAdvisorRule::ALL {
            assert!(parsed.is_disabled(rule));
        }
    }

    #[test]
    fn supports_independent_enable_and_rejects_reserved_bits() {
        let mut settings = PowerPointPresentationAdvisorSettings::new();
        settings.disable(PowerPointAdvisorRule::TooManyFonts);
        settings.disable(PowerPointAdvisorRule::PrintTip);
        settings.enable(PowerPointAdvisorRule::TooManyFonts);
        assert_eq!(
            settings.disabled_rules().collect::<Vec<_>>(),
            vec![PowerPointAdvisorRule::PrintTip]
        );

        let mut record = settings.to_record();
        record.data.copy_from_slice(&0x0800u32.to_le_bytes());
        assert!(PowerPointPresentationAdvisorSettings::parse(&record).is_err());
    }
}

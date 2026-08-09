//! Presentation Advisor preferences from MS-PPT 2.4.6.

use crate::consts::RecordType;

use super::package::{Error, Result};
use super::records::Record;

const PRESENTATION_ADVISOR_RECORD_TYPE: u16 = 0x177a;
const ADVISOR_RULE_MASK: u32 = 0x07ff;

/// A presentation-style rule that `PowerPoint` can suppress warnings for.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum AdvisorRule {
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

impl AdvisorRule {
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

    const fn bit(self) -> u16 {
        1u16 << (self as u8)
    }
}

/// Compact Presentation Advisor suppression settings.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PresentationAdvisorSettings` is the established public API name; renaming it would break downstream crates"
)]
pub struct PresentationAdvisorSettings {
    disabled_mask: u16,
}

impl PresentationAdvisorSettings {
    /// Create settings with every advisor rule enabled.
    #[must_use]
    pub const fn new() -> Self {
        Self { disabled_mask: 0 }
    }

    /// Parse one strict `PresAdvisorFlags9Atom`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or size is invalid or reserved
    /// flag bits are set.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type_raw != PRESENTATION_ADVISOR_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(Error::Corrupted(
                "PresAdvisorFlags9Atom has an invalid record header or size".to_string(),
            ));
        }
        let flags = u32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]);
        if flags & !ADVISOR_RULE_MASK != 0 {
            return Err(Error::Corrupted(
                "PresAdvisorFlags9Atom has nonzero reserved bits".to_string(),
            ));
        }
        let disabled_mask = u16::try_from(flags).map_err(|_err| {
            Error::Corrupted("PresAdvisorFlags9Atom flags exceed 16 bits".to_string())
        })?;
        Ok(Self { disabled_mask })
    }

    /// Discover the single advisor atom in the PPT9 document tag.
    pub(crate) fn parse_document(document: &Record) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == PRESENTATION_ADVISOR_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return Err(Error::Corrupted(
                "PPT9 document tag contains multiple PresAdvisorFlags9Atom records".to_string(),
            ));
        }
        Self::parse(record).map(Some)
    }

    /// Return whether warnings for `rule` are disabled.
    #[must_use]
    pub const fn is_disabled(self, rule: AdvisorRule) -> bool {
        self.disabled_mask & rule.bit() != 0
    }

    /// Disable warnings for `rule`.
    pub fn disable(&mut self, rule: AdvisorRule) {
        self.disabled_mask |= rule.bit();
    }

    /// Enable warnings for `rule`.
    pub fn enable(&mut self, rule: AdvisorRule) {
        self.disabled_mask &= !rule.bit();
    }

    /// Iterate over disabled rules in specification bit order.
    pub fn disabled_rules(self) -> impl Iterator<Item = AdvisorRule> {
        AdvisorRule::ALL
            .into_iter()
            .filter(move |rule| self.is_disabled(*rule))
    }

    /// Encode the exact `PowerPoint` 9 atom.
    #[must_use]
    pub fn to_record(self) -> Record {
        let flags = u32::from(self.disabled_mask);
        Record {
            record_type: RecordType::from(PRESENTATION_ADVISOR_RECORD_TYPE),
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
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_every_typed_advisor_rule() {
        let mut settings = PresentationAdvisorSettings::new();
        for rule in AdvisorRule::ALL {
            settings.disable(rule);
        }
        let parsed = PresentationAdvisorSettings::parse(&settings.to_record()).unwrap();
        assert_eq!(
            parsed.disabled_rules().collect::<Vec<_>>(),
            AdvisorRule::ALL
        );
        for rule in AdvisorRule::ALL {
            assert!(parsed.is_disabled(rule));
        }
    }

    #[test]
    fn supports_independent_enable_and_rejects_reserved_bits() {
        let mut settings = PresentationAdvisorSettings::new();
        settings.disable(AdvisorRule::TooManyFonts);
        settings.disable(AdvisorRule::PrintTip);
        settings.enable(AdvisorRule::TooManyFonts);
        assert_eq!(
            settings.disabled_rules().collect::<Vec<_>>(),
            vec![AdvisorRule::PrintTip]
        );

        let mut record = settings.to_record();
        record.data.copy_from_slice(&0x0800u32.to_le_bytes());
        assert!(PresentationAdvisorSettings::parse(&record).is_err());
    }
}

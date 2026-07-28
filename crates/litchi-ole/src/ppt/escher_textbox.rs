use super::super::consts::PptRecordType;
use super::TextRuler;
use super::package::{PptError, Result};
/// EscherTextboxWrapper implementation.
///
/// Based on Apache POI's EscherTextboxWrapper, this wraps an Escher textbox record
/// and provides access to its child PPT records (TextCharsAtom, TextBytesAtom, StyleTextPropAtom).
use super::records::PptRecord;
use super::text_interaction::{
    PowerPointTextInteraction, PowerPointTextInteractionLimits, text_units_from_records,
};
use super::text_run::{ParagraphRun, TextRun, TextRunExtractor};

/// Wrapper around Escher textbox data.
///
/// Based on POI's EscherTextboxWrapper. Parses child records from the textbox data.
#[derive(Debug, Clone)]
pub struct EscherTextboxWrapper {
    /// The raw Escher textbox data
    data: Vec<u8>,
    /// Child PPT records found in the textbox
    child_records: Vec<PptRecord>,
    /// Extracted text
    text: String,
    /// Text split into character-formatting runs
    runs: Vec<TextRun>,
    /// Text split into paragraph-formatting runs
    paragraph_runs: Vec<ParagraphRun>,
    /// Text ruler carried by the textbox, when present
    text_ruler: Option<TextRuler>,
    /// Range-anchored click and mouse-over actions.
    text_interactions: Vec<PowerPointTextInteraction>,
    /// Header/footer metacharacter placeholders in the textbox.
    metachars: Vec<super::text_metachar::PowerPointTextMetachar>,
    /// Outline text references tying the textbox to outline text bodies.
    outline_text_refs: Vec<super::text_si_exception::PowerPointOutlineTextRef>,
}

impl EscherTextboxWrapper {
    /// Create a new wrapper from Escher textbox data.
    ///
    /// Based on POI's EscherTextboxWrapper constructor which calls
    /// Record.findChildRecords(data, 0, data.length).
    pub fn new(data: Vec<u8>) -> Result<Self> {
        Self::new_with_interaction_limits(data, PowerPointTextInteractionLimits::default())
    }

    /// Create a wrapper with explicit text-interaction resource limits.
    pub fn new_with_interaction_limits(
        data: Vec<u8>,
        limits: PowerPointTextInteractionLimits,
    ) -> Result<Self> {
        // Parse child records from the escher data
        let child_records = Self::find_child_records(&data)?;

        let mut extractor = TextRunExtractor::new();
        extractor.extract_from_records(&child_records)?;
        let text = extractor.text().to_string();
        let runs = extractor.runs().to_vec();
        let paragraph_runs = extractor.paragraph_runs().to_vec();
        let mut ruler_records = child_records
            .iter()
            .filter(|record| record.record_type == PptRecordType::TextRulerAtom);
        let text_ruler = ruler_records
            .next()
            .map(|record| TextRuler::parse(&record.data))
            .transpose()?;
        if ruler_records.next().is_some() {
            return Err(PptError::Corrupted(
                "ClientTextbox contains multiple TextRulerAtom records".to_string(),
            ));
        }
        let text_interactions = Self::text_interactions_from_records(&child_records, limits)?;
        let metachars = super::text_metachar::metachars_from_records(child_records.iter())?;
        let outline_text_refs = child_records
            .iter()
            .filter(|record| record.record_type == crate::consts::PptRecordType::OutlineTextRefAtom)
            .map(super::text_si_exception::PowerPointOutlineTextRef::parse_record)
            .collect::<Result<Vec<_>>>()?;

        Ok(Self {
            data,
            child_records,
            text,
            runs,
            paragraph_runs,
            text_ruler,
            text_interactions,
            metachars,
            outline_text_refs,
        })
    }

    /// Find child PPT records in the Escher textbox data.
    ///
    /// Based on POI's Record.findChildRecords().
    fn find_child_records(data: &[u8]) -> Result<Vec<PptRecord>> {
        let mut records = Vec::new();
        let mut offset = 0;

        while offset < data.len() {
            let header_end = offset.checked_add(8).ok_or_else(|| {
                PptError::Corrupted("ClientTextbox record offset overflow".to_string())
            })?;
            if header_end > data.len() {
                return Err(PptError::Corrupted(
                    "Truncated record header in ClientTextbox".to_string(),
                ));
            }

            let payload_length = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            let payload_length = usize::try_from(payload_length).map_err(|_| {
                PptError::Corrupted("ClientTextbox record size overflow".to_string())
            })?;
            let record_end = header_end.checked_add(payload_length).ok_or_else(|| {
                PptError::Corrupted("ClientTextbox record size overflow".to_string())
            })?;
            if record_end > data.len() {
                return Err(PptError::Corrupted(
                    "Record extends beyond ClientTextbox data".to_string(),
                ));
            }

            let (record, consumed) = PptRecord::parse(&data[offset..record_end], 0)?;
            if consumed != record_end - offset {
                return Err(PptError::Corrupted(
                    "ClientTextbox child record was only partially parsed".to_string(),
                ));
            }
            records.push(record);
            offset = record_end;
        }

        Ok(records)
    }

    /// Parse only text-range actions without retaining or decoding the textbox.
    pub fn parse_text_interactions_with_limits(
        data: &[u8],
        limits: PowerPointTextInteractionLimits,
    ) -> Result<Vec<PowerPointTextInteraction>> {
        let records = Self::find_child_records(data)?;
        Self::text_interactions_from_records(&records, limits)
    }

    fn text_interactions_from_records(
        records: &[PptRecord],
        limits: PowerPointTextInteractionLimits,
    ) -> Result<Vec<PowerPointTextInteraction>> {
        let has_text_interactions = records.iter().any(|record| {
            matches!(
                record.record_type,
                PptRecordType::InteractiveInfo | PptRecordType::TextInteractiveInfoAtom
            )
        });
        if !has_text_interactions {
            return Ok(Vec::new());
        }
        let mut headers = records
            .iter()
            .filter(|record| record.record_type == PptRecordType::TextHeaderAtom);
        let header = headers.next().ok_or_else(|| {
            PptError::Corrupted("Interactive ClientTextbox has no TextHeaderAtom".to_string())
        })?;
        if headers.next().is_some()
            || header.version != 0
            || header.data_length != 4
            || header.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "Interactive ClientTextbox has an invalid TextHeaderAtom".to_string(),
            ));
        }
        let _ = crate::ppt::PowerPointTextType::parse(u32::from_le_bytes(
            header.data[..4].try_into().unwrap(),
        ))?;
        let text_units = text_units_from_records(records)?;
        PowerPointTextInteraction::parse_records(records, text_units, limits)
    }

    /// Get the extracted text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get child records.
    pub fn child_records(&self) -> &[PptRecord] {
        &self.child_records
    }

    /// Get the character-formatting runs extracted from `StyleTextPropAtom`.
    pub fn runs(&self) -> &[TextRun] {
        &self.runs
    }

    /// Get paragraph-formatting runs extracted from `StyleTextPropAtom`.
    pub fn paragraph_runs(&self) -> &[ParagraphRun] {
        &self.paragraph_runs
    }

    /// Get the textbox-specific ruler, when present.
    pub fn text_ruler(&self) -> Option<&TextRuler> {
        self.text_ruler.as_ref()
    }

    /// Strictly paired text-range interactions in record order.
    pub fn text_interactions(&self) -> &[PowerPointTextInteraction] {
        &self.text_interactions
    }

    /// Header/footer metacharacter placeholders in this textbox, in record
    /// order (MS-PPT 2.9.47-2.9.52). Placeholders are never substituted,
    /// formatted, or laid out.
    pub fn metachars(&self) -> &[super::text_metachar::PowerPointTextMetachar] {
        &self.metachars
    }

    /// Outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78) tying
    /// this textbox to outline text bodies.
    pub fn outline_text_refs(&self) -> &[super::text_si_exception::PowerPointOutlineTextRef] {
        &self.outline_text_refs
    }

    /// Find a StyleTextPropAtom record.
    pub fn find_style_text_prop_atom(&self) -> Option<&PptRecord> {
        self.child_records
            .iter()
            .find(|r| r.record_type == PptRecordType::StyleTextPropAtom)
    }

    /// Get the raw data.
    pub fn data(&self) -> &[u8] {
        &self.data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escher_textbox_wrapper_creation() {
        // Create a simple textbox with a TextCharsAtom
        // PPT record header: ver/inst (2 bytes) | type (2 bytes) | length (4 bytes)
        let mut data = Vec::new();

        // TextCharsAtom: Record type 0x0FA0 (4000)
        data.extend_from_slice(&[0x00, 0x00]); // Version/instance (ver=0, inst=0)
        data.extend_from_slice(&[0xA0, 0x0F]); // Record type 0x0FA0 (little-endian)
        data.extend_from_slice(&[0x0A, 0x00, 0x00, 0x00]); // Length: 10 bytes (little-endian)

        // Text data (UTF-16LE): "Hello"
        data.extend_from_slice(&[
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
        ]);

        let wrapper = EscherTextboxWrapper::new(data).unwrap();
        assert!(wrapper.text().contains("Hello") || !wrapper.text().is_empty());
        assert!(!wrapper.child_records().is_empty());
    }

    #[test]
    fn rejects_truncated_child_record_instead_of_resynchronizing() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&4008u16.to_le_bytes());
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(b"bad");

        let error = EscherTextboxWrapper::new(data).unwrap_err();
        assert!(error.to_string().contains("extends beyond ClientTextbox"));
    }

    #[test]
    fn allows_client_textboxes_without_inline_text() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&3999u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());

        let wrapper = EscherTextboxWrapper::new(data).unwrap();
        assert_eq!(wrapper.text(), "");
        assert!(wrapper.runs().is_empty());
    }

    #[test]
    fn rejects_multiple_text_rulers() {
        let mut data = Vec::new();
        for _ in 0..2 {
            data.extend_from_slice(&0u16.to_le_bytes());
            data.extend_from_slice(&4006u16.to_le_bytes());
            data.extend_from_slice(&4u32.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
        }

        let error = EscherTextboxWrapper::new(data).unwrap_err();
        assert!(error.to_string().contains("multiple TextRulerAtom"));
    }
}

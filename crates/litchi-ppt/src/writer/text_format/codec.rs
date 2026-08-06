//! Binary encoders for PowerPoint text atoms and style properties.
//!
//! The public builder remains available from this module and the
//! writer::text_format facade.

use super::semantic::{
    Paragraph, TextAlign, TextColor, TextDirection, TextFontAlign, char_mask, para_mask,
};
use super::validation;

// =============================================================================
// Text Properties Builder
// =============================================================================

/// Builder for TextCharsAtom/TextBytesAtom and StyleTextPropAtom
pub struct TextPropsBuilder {
    paragraphs: Vec<Paragraph>,
}

impl TextPropsBuilder {
    /// Create a new builder
    pub fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
        }
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.paragraphs.push(para);
    }

    /// Build TextCharsAtom (UTF-16LE text), adding CR between paragraphs.
    ///
    /// The final paragraph break is implicit and is not stored in the text atom.
    pub fn build_text_chars(&self) -> Vec<u8> {
        let mut data = Vec::new();
        for (i, para) in self.paragraphs.iter().enumerate() {
            for run in &para.runs {
                for ch in run.text.encode_utf16() {
                    data.extend_from_slice(&ch.to_le_bytes());
                }
            }
            // Add paragraph separator (CR) for all paragraphs including the last
            // This makes the text length match the StyleTextPropAtom char counts
            if i < self.paragraphs.len() - 1 {
                data.extend_from_slice(&0x000Du16.to_le_bytes()); // CR between paragraphs
            }
        }
        data
    }

    /// Build StyleTextPropAtom containing paragraph and character formatting
    ///
    /// According to MS-PPT spec:
    /// - Sum of paragraph character counts = total text length + 1
    /// - Sum of character run counts = total text length + 1
    /// - The +1 accounts for an implicit terminating character
    pub fn build_style_text_prop(&self) -> std::io::Result<Vec<u8>> {
        let mut data = Vec::new();

        if self.paragraphs.is_empty() {
            data.extend_from_slice(&1u32.to_le_bytes());
            data.extend_from_slice(&0i16.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&1u32.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            return Ok(data);
        }

        // Paragraph properties (TextPFRun entries)
        // Each paragraph covers its runs + CR separator (except last paragraph gets +1 for terminator)
        for para in &self.paragraphs {
            let para_text_len = para.runs.iter().try_fold(0u32, |total, run| {
                let count = u32::try_from(run.text.encode_utf16().count()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint text run exceeds the PPT size limit",
                    )
                })?;
                total.checked_add(count).ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint paragraph exceeds the PPT size limit",
                    )
                })
            })?;

            // Character count: text + CR (or +1 for last paragraph terminator)
            // +1 for either CR separator or implicit terminating character
            let char_count = para_text_len.checked_add(1).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "PowerPoint paragraph exceeds the PPT size limit",
                )
            })?;
            data.extend_from_slice(&char_count.to_le_bytes());

            validation::validate_indent_level(para.indent_level)?;
            data.extend_from_slice(&para.indent_level.to_le_bytes());

            // Build mask based on what properties are set
            let mut mask = para.explicit_mask;
            if para.alignment != TextAlign::Left {
                mask |= para_mask::ALIGNMENT;
            }
            if para.line_spacing != 100 {
                mask |= para_mask::LINE_SPACING;
            }
            if para.space_before != 0 {
                mask |= para_mask::SPACE_BEFORE;
            }
            if para.space_after != 0 {
                mask |= para_mask::SPACE_AFTER;
            }
            if para.left_margin != 0 {
                mask |= para_mask::LEFT_MARGIN;
            }
            if para.indent != 0 {
                mask |= para_mask::INDENT;
            }
            if para.bullet_enabled.is_some() || para.bullet_char.is_some() {
                mask |= para_mask::HAS_BULLET;
            }
            if para.bullet_char.is_some() {
                mask |= para_mask::BULLET_CHAR;
            }
            if para.bullet_font_enabled.is_some() || para.bullet_font_index.is_some() {
                mask |= para_mask::BULLET_HAS_FONT;
            }
            if para.bullet_font_index.is_some() {
                mask |= para_mask::BULLET_FONT;
            }
            if para.bullet_size_enabled.is_some() || para.bullet_size.is_some() {
                mask |= para_mask::BULLET_HAS_SIZE;
            }
            if para.bullet_size.is_some() {
                mask |= para_mask::BULLET_SIZE;
            }
            if para.bullet_color_enabled.is_some() || para.bullet_color.is_some() {
                mask |= para_mask::BULLET_HAS_COLOR;
            }
            if para.bullet_color.is_some() {
                mask |= para_mask::BULLET_COLOR;
            }
            if para.default_tab_size.is_some() {
                mask |= para_mask::DEFAULT_TAB_SIZE;
            }
            if para.tab_stops.is_some() {
                mask |= para_mask::TAB_STOPS;
            }
            if para.font_alignment.is_some() {
                mask |= para_mask::FONT_ALIGNMENT;
            }
            if para.character_wrap.is_some() {
                mask |= para_mask::CHARACTER_WRAP;
            }
            if para.word_wrap.is_some() {
                mask |= para_mask::WORD_WRAP;
            }
            if para.overflow.is_some() {
                mask |= para_mask::OVERFLOW;
            }
            if para.text_direction.is_some() {
                mask |= para_mask::TEXT_DIRECTION;
            }

            data.extend_from_slice(&mask.to_le_bytes());

            // Write properties according to mask
            if mask
                & (para_mask::HAS_BULLET
                    | para_mask::BULLET_HAS_FONT
                    | para_mask::BULLET_HAS_COLOR
                    | para_mask::BULLET_HAS_SIZE)
                != 0
            {
                let mut flags = 0u16;
                if para.bullet_enabled.unwrap_or(para.bullet_char.is_some()) {
                    flags |= 0x0001;
                }
                if para
                    .bullet_font_enabled
                    .unwrap_or(para.bullet_font_index.is_some())
                {
                    flags |= 0x0002;
                }
                if para
                    .bullet_color_enabled
                    .unwrap_or(para.bullet_color.is_some())
                {
                    flags |= 0x0004;
                }
                if para
                    .bullet_size_enabled
                    .unwrap_or(para.bullet_size.is_some())
                {
                    flags |= 0x0008;
                }
                data.extend_from_slice(&flags.to_le_bytes());
            }
            if mask & para_mask::BULLET_CHAR != 0 {
                let bullet = para.bullet_char.unwrap_or('•');
                let ch = u16::try_from(bullet as u32).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT bullet characters must fit in one UTF-16 code unit",
                    )
                })?;
                data.extend_from_slice(&ch.to_le_bytes());
            }
            if mask & para_mask::BULLET_FONT != 0 {
                data.extend_from_slice(&para.bullet_font_index.unwrap_or(0).to_le_bytes());
            }
            if mask & para_mask::BULLET_SIZE != 0 {
                let size = para.bullet_size.unwrap_or(100);
                validation::validate_bullet_size(size)?;
                data.extend_from_slice(&size.to_le_bytes());
            }
            if mask & para_mask::BULLET_COLOR != 0 {
                let color = para.bullet_color.unwrap_or(TextColor::BLACK);
                validation::validate_bullet_color(color)?;
                data.extend_from_slice(&color.to_ppt_color().to_le_bytes());
            }
            if mask & para_mask::ALIGNMENT != 0 {
                data.extend_from_slice(&(para.alignment as u16).to_le_bytes());
            }
            if mask & para_mask::LINE_SPACING != 0 {
                data.extend_from_slice(&para.line_spacing.to_le_bytes());
            }
            if mask & para_mask::SPACE_BEFORE != 0 {
                data.extend_from_slice(&para.space_before.to_le_bytes());
            }
            if mask & para_mask::SPACE_AFTER != 0 {
                data.extend_from_slice(&para.space_after.to_le_bytes());
            }
            if mask & para_mask::LEFT_MARGIN != 0 {
                data.extend_from_slice(&para.left_margin.to_le_bytes());
            }
            if mask & para_mask::INDENT != 0 {
                data.extend_from_slice(&para.indent.to_le_bytes());
            }
            if mask & para_mask::DEFAULT_TAB_SIZE != 0 {
                data.extend_from_slice(&para.default_tab_size.unwrap_or(0).to_le_bytes());
            }
            if mask & para_mask::TAB_STOPS != 0 {
                let tab_stops = para.tab_stops.as_deref().unwrap_or_default();
                let count = u16::try_from(tab_stops.len()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT paragraph has more than 65535 tab stops",
                    )
                })?;
                data.extend_from_slice(&count.to_le_bytes());
                for tab_stop in tab_stops {
                    data.extend_from_slice(&tab_stop.position.to_le_bytes());
                    data.extend_from_slice(&(tab_stop.alignment as u16).to_le_bytes());
                }
            }
            if mask & para_mask::FONT_ALIGNMENT != 0 {
                data.extend_from_slice(
                    &(para.font_alignment.unwrap_or(TextFontAlign::Roman) as u16).to_le_bytes(),
                );
            }
            if mask & (para_mask::CHARACTER_WRAP | para_mask::WORD_WRAP | para_mask::OVERFLOW) != 0
            {
                let mut flags = 0u16;
                if para.character_wrap.unwrap_or(false) {
                    flags |= 0x0001;
                }
                if para.word_wrap.unwrap_or(false) {
                    flags |= 0x0002;
                }
                if para.overflow.unwrap_or(false) {
                    flags |= 0x0004;
                }
                data.extend_from_slice(&flags.to_le_bytes());
            }
            if mask & para_mask::TEXT_DIRECTION != 0 {
                data.extend_from_slice(
                    &(para.text_direction.unwrap_or(TextDirection::LeftToRight) as u16)
                        .to_le_bytes(),
                );
            }
        }

        // Character properties (TextCFRun entries)
        // Write one entry per run. The last run in each paragraph gets +1 for CR/terminator.
        for para in &self.paragraphs {
            let num_runs = para.runs.len();

            if num_runs == 0 {
                // Cover the paragraph separator or the implicit final paragraph break.
                data.extend_from_slice(&1u32.to_le_bytes());
                data.extend_from_slice(&0u32.to_le_bytes());
                continue;
            }

            for (run_idx, run) in para.runs.iter().enumerate() {
                validation::validate_font_size(run.font_size)?;
                validation::validate_run_color(run.color)?;
                if let Some(position) = run.baseline_position {
                    validation::validate_baseline_position(position)?;
                }
                if let Some(id) = run.style.pp9_run_id {
                    validation::validate_pp9_run_id(id)?;
                }
                validation::validate_style_mask(run.style.specified_mask)?;
                let is_last_run = run_idx == num_runs - 1;

                // Character count for this run
                // Last run of last paragraph gets +1 for terminator
                // Last run of non-last paragraph gets +1 for CR separator
                let run_units = u32::try_from(run.text.encode_utf16().count()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint text run exceeds the PPT size limit",
                    )
                })?;
                let char_count = if is_last_run {
                    run_units.checked_add(1).ok_or_else(|| {
                        std::io::Error::new(
                            std::io::ErrorKind::InvalidInput,
                            "PowerPoint text run exceeds the PPT size limit",
                        )
                    })?
                } else {
                    run_units
                };
                data.extend_from_slice(&char_count.to_le_bytes());

                // Build mask
                let mut mask = run.style.to_mask();
                mask |= char_mask::FONT_SIZE; // Always include font size
                mask |= char_mask::FONT_COLOR; // Always include color
                mask |= char_mask::FONT_REF; // Always include font reference
                if run.asian_font_index.is_some() {
                    mask |= char_mask::ASIAN_FONT_REF;
                }
                if run.ansi_font_index.is_some() {
                    mask |= char_mask::ANSI_FONT_REF;
                }
                if run.symbol_font_index.is_some() {
                    mask |= char_mask::SYMBOL_FONT_REF;
                }
                if run.baseline_position.is_some() {
                    mask |= char_mask::POSITION;
                }

                data.extend_from_slice(&mask.to_le_bytes());

                // Font style flags (only if any flags are set)
                if mask & 0xFFFF != 0 {
                    let flags = run.style.to_flags();
                    data.extend_from_slice(&flags.to_le_bytes());
                }

                // Font reference (if font_ref bit is set)
                data.extend_from_slice(&run.font_index.to_le_bytes());

                if let Some(index) = run.asian_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }
                if let Some(index) = run.ansi_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }
                if let Some(index) = run.symbol_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }

                // Font size is stored directly in points.
                data.extend_from_slice(&run.font_size.to_le_bytes());

                // Color (POI format: R | G<<8 | B<<16 | 0xFE<<24)
                let color = run.color.to_ppt_color();
                data.extend_from_slice(&color.to_le_bytes());

                if let Some(position) = run.baseline_position {
                    data.extend_from_slice(&position.to_le_bytes());
                }
            }
        }

        Ok(data)
    }

    /// Get total character count
    pub fn total_chars(&self) -> u32 {
        self.paragraphs.iter().map(|p| p.char_count()).sum()
    }
}

impl Default for TextPropsBuilder {
    fn default() -> Self {
        Self::new()
    }
}

use super::{
    ControlWord, Cow, Destination, Parser, RtfError, RtfResult, SmallVec, StyleBlock, Token,
    append_transport_bytes,
};

/// `\uN` code units with the token each was read from.
#[derive(Default)]
pub(super) struct UnicodeRun {
    pub(super) units: SmallVec<[u16; 4]>,
    tokens: SmallVec<[usize; 4]>,
}

impl UnicodeRun {
    pub(super) fn len(&self) -> usize {
        self.units.len()
    }

    /// Append one code unit read from `token`.
    pub(super) fn push(&mut self, unit: u16, token: usize) {
        self.units.push(unit);
        self.tokens.push(token);
    }
}

impl Parser<'_> {
    /// Byte offset of a token in the parser input, for repair records.
    pub(super) fn token_offset(&self, token: usize) -> usize {
        self.token_spans
            .and_then(|spans| spans.get(token))
            .map_or(0, |span| span.start)
    }

    /// Substitute U+FFFD for the unpaired surrogate read from `token`.
    pub(super) fn substitute_unpaired_surrogate(&mut self, output: &mut String, token: usize) {
        output.push(char::REPLACEMENT_CHARACTER);
        let offset = self.token_offset(token);
        self.tolerance
            .record(crate::leniency::UnicodeDefect::UnpairedSurrogate, offset);
    }

    /// Push one non-high-surrogate `\uN` unit read from the cursor token,
    /// substituting U+FFFD for an unpaired low surrogate.
    pub(super) fn push_single_unicode_unit(&mut self, output: &mut String, unit: u16) {
        if let Some(character) = char::from_u32(u32::from(unit)) {
            output.push(character);
        } else {
            self.substitute_unpaired_surrogate(output, self.pos);
        }
    }

    /// Collect a run of `\uN` code units, the first (`first_code`) read from
    /// the token at the cursor.
    ///
    /// Consecutive `\uN` controls join the run. Word also writes each half
    /// of a surrogate pair as its own `\uN` followed by its own fallback
    /// (`\u-10179?\u-8694?`); when the fallback characters still owed after
    /// a high surrogate are immediately followed by a low surrogate, they are
    /// consumed and the low surrogate joins the run, so the pair decodes as
    /// one scalar the way Word and LibreOffice show it.
    ///
    /// Returns how many collected units still owe their `skip` fallback
    /// characters, which follow the cursor. With `bounded`, the run is held
    /// to the shared intermediate text budget.
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
    )]
    pub(super) fn collect_unicode_run(
        &mut self,
        first_code: i32,
        skip: usize,
        run: &mut UnicodeRun,
        bounded: bool,
    ) -> RtfResult<usize> {
        let mut awaiting_fallback = 0usize;
        let mut code = first_code;
        loop {
            if bounded {
                Self::push_bounded_unicode_code(&mut run.units, code)?;
                run.tokens.push(self.pos);
            } else {
                run.push(code as u16, self.pos);
            }
            self.pos += 1;
            awaiting_fallback = awaiting_fallback.saturating_add(1);
            if let Some(Token::Control(ControlWord::Unicode(next))) = self.tokens.get(self.pos) {
                code = *next;
                continue;
            }
            let high = (0xD800..=0xDBFF).contains(&(code as u16));
            let owed = skip.saturating_mul(awaiting_fallback);
            match self.low_surrogate_after_fallback(owed) {
                Some((low_token, low)) if high => {
                    self.pos = low_token;
                    awaiting_fallback = 0;
                    code = low;
                },
                _ => return Ok(awaiting_fallback),
            }
        }
    }

    /// The `\uN` low surrogate (and its token) that directly follows the
    /// cursor once `owed` fallback characters are skipped, if any.
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "RTF \\uN parameters are signed 16-bit; the u16 wrap implements the specified negative-value conversion"
    )]
    fn low_surrogate_after_fallback(&self, mut owed: usize) -> Option<(usize, i32)> {
        let mut token = self.pos;
        while owed > 0 {
            match self.tokens.get(token)? {
                Token::Text(text) => {
                    owed = owed.checked_sub(text.chars().count())?;
                },
                Token::Control(control) if super::control_symbol_text(control).is_some() => {
                    owed -= 1;
                },
                Token::OpenBrace | Token::CloseBrace | Token::Control(_) | Token::Binary(_) => {
                    return None;
                },
            }
            token += 1;
        }
        match self.tokens.get(token)? {
            Token::Control(ControlWord::Unicode(code))
                if (0xDC00..=0xDFFF).contains(&(*code as u16)) =>
            {
                Some((token, *code))
            },
            _ => None,
        }
    }

    /// Decode a collected run into `output`, substituting U+FFFD for each
    /// unpaired UTF-16 surrogate the way Word and LibreOffice do, and
    /// recording every substitution as [`UnicodeDefect::UnpairedSurrogate`]
    /// at the offset of the `\uN` that carried the unit.
    ///
    /// [`UnicodeDefect::UnpairedSurrogate`]: crate::leniency::UnicodeDefect::UnpairedSurrogate
    pub(super) fn decode_unicode_run(&mut self, run: &UnicodeRun, output: &mut String) {
        let mut unit = 0usize;
        for decoded in std::char::decode_utf16(run.units.iter().copied()) {
            if let Ok(character) = decoded {
                output.push(character);
                unit = unit.saturating_add(character.len_utf16());
            } else {
                let token = run.tokens.get(unit).copied().unwrap_or_default();
                self.substitute_unpaired_surrogate(output, token);
                unit = unit.saturating_add(1);
            }
        }
    }

    /// Parse Unicode character sequence with fallback handling.
    ///
    /// RTF Unicode format: `\uN` where N is a signed 16-bit decimal value
    /// Followed by `\ucN` fallback characters (usually ANSI representation)
    ///
    /// Handles compound Unicode characters (surrogate pairs for emoji, etc.)
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "RTF \\uN parameters are signed 16-bit; the wrapping conversions implement the specified negative-value and fallback-skip behavior"
    )]
    pub(super) fn parse_unicode_sequence(&mut self, first_code: i32) -> RtfResult<()> {
        let skip_count = self.current_state()?.unicode_skip as usize;

        // Collect consecutive Unicode values, joining surrogate pairs.
        let mut unicode_values = UnicodeRun::default();
        let awaiting_fallback =
            self.collect_unicode_run(first_code, skip_count, &mut unicode_values, false)?;

        // Skip fallback characters based on unicode_skip count
        // Fallback chars are for non-Unicode readers (usually hex escapes or plain ASCII)
        let mut fallback_skip = skip_count * awaiting_fallback;
        let mut fallback_remainder = None;

        // Handle fallback: skip the next N characters/tokens
        while fallback_skip > 0 {
            let Some(token) = self.tokens.get(self.pos) else {
                break;
            };
            match token {
                Token::Text(text) => {
                    let character_count = text.chars().count();
                    if character_count <= fallback_skip {
                        fallback_skip -= character_count;
                        self.pos += 1;
                    } else {
                        fallback_remainder =
                            Some(text.chars().skip(fallback_skip).collect::<String>());
                        fallback_skip = 0;
                        self.pos += 1;
                    }
                },
                Token::Control(ControlWord::Unicode(_)) => {
                    // Next unicode, don't skip
                    break;
                },
                Token::OpenBrace | Token::CloseBrace | Token::Control(_) | Token::Binary(_) => {
                    // Treat other tokens as single character
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
            }
        }

        // Convert Unicode values to UTF-8 string
        let mut unicode_str = String::with_capacity(unicode_values.len().saturating_mul(3));
        self.decode_unicode_run(&unicode_values, &mut unicode_str);

        let (destination, in_table, table_nesting_level) = {
            let state = self.current_state()?;
            (state.destination, state.in_table, state.table_nesting_level)
        };
        if destination == Destination::DocumentBody && (in_table || table_nesting_level >= 2) {
            self.append_table_text(unicode_str.as_bytes(), table_nesting_level)?;
            if let Some(remainder) = fallback_remainder {
                self.append_table_text(remainder.as_bytes(), table_nesting_level)?;
            }
        } else if destination == Destination::DocumentBody {
            let state = self.current_state()?.clone();
            // Add the Unicode sequence to the document as its own formatted block.
            let allocated = self.arena.alloc_str(&unicode_str);
            let start = self.body_text_len;
            if state.revision_type == Some(super::super::super::annotation::RevisionType::Deletion)
            {
                self.append_revision_text(&state, allocated, start, start)?;
            } else {
                let block =
                    StyleBlock::new(Cow::Borrowed(allocated), state.formatting, state.paragraph);
                self.body_text_len =
                    self.body_text_len
                        .checked_add(allocated.len())
                        .ok_or_else(|| {
                            RtfError::MalformedDocument("RTF body text length overflow".into())
                        })?;
                self.prepare_body_block_push()?;
                self.blocks.push(block);
                self.append_revision_text(&state, allocated, start, self.body_text_len)?;
            }

            // A fallback and subsequent text often share one lexer token. Preserve
            // the portion after the configured fallback character count.
            if let Some(remainder) = fallback_remainder {
                let mut buffer = SmallVec::<[u8; 256]>::new();
                append_transport_bytes(&mut buffer, &remainder)?;
                self.flush_text_buffer(&mut buffer)?;
            }
        }

        Ok(())
    }
}

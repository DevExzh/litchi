use super::*;

impl<'a> Parser<'a> {
    /// Parse Unicode character sequence with fallback handling.
    ///
    /// RTF Unicode format: `\uN` where N is a signed 16-bit decimal value
    /// Followed by `\ucN` fallback characters (usually ANSI representation)
    ///
    /// Handles compound Unicode characters (surrogate pairs for emoji, etc.)
    pub(super) fn parse_unicode_sequence(&mut self, first_code: i32) -> RtfResult<()> {
        let skip_count = self.current_state()?.unicode_skip as usize;

        // Collect all consecutive unicode values (for surrogate pairs)
        let mut unicode_values = SmallVec::<[u16; 4]>::new();

        // Convert signed 16-bit value to unsigned
        unicode_values.push(first_code as u16);
        self.pos += 1;

        // Look ahead for additional Unicode characters (compound characters)
        while let Some(token) = self.tokens.get(self.pos) {
            if let Token::Control(ControlWord::Unicode(code)) = token {
                unicode_values.push(*code as u16);
                self.pos += 1;
            } else {
                break;
            }
        }

        // Skip fallback characters based on unicode_skip count
        // Fallback chars are for non-Unicode readers (usually hex escapes or plain ASCII)
        let mut fallback_skip = skip_count * unicode_values.len();
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
                _ => {
                    // Treat other tokens as single character
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
            }
        }

        // Convert Unicode values to UTF-8 string
        let unicode_str = String::from_utf16(&unicode_values)
            .map_err(|e| RtfError::InvalidUnicode(format!("Invalid Unicode sequence: {}", e)))?;

        let state = self.current_state()?.clone();
        if state.destination == Destination::DocumentBody
            && (state.in_table || state.table_nesting_level >= 2)
        {
            self.append_table_text(unicode_str.as_bytes(), state.table_nesting_level)?;
            if let Some(remainder) = fallback_remainder {
                self.append_table_text(remainder.as_bytes(), state.table_nesting_level)?;
            }
        } else if state.destination == Destination::DocumentBody {
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

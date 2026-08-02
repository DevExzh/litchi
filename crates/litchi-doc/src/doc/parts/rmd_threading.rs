//! Word e-mail review threading data (`RmdThreading`, MS-DOC 2.9.230).
//!
//! The structure describes the e-mail messages a document was routed through
//! for review and the per-author display data. It is parsed as inert
//! metadata: message identifiers are stored verbatim and no message is ever
//! contacted, opened, or rendered.

use super::super::CommentDateTime;
use super::super::package::{DocError, Result};
use super::super::revision::decode_dttm;
use super::fib::FileInformationBlock;

/// Table-pointer index of `fcRmdThreading`/`lcbRmdThreading`.
const RMD_THREADING: usize = 94;
/// `fExtend` marker shared by all six `RmdThreading` STTBs.
const STTB_F_EXTEND: u16 = 0xFFFF;
/// `cbExtra` of `SttbMessage`: one `MDP` per entry (MS-DOC 2.9.155).
const MDP_SIZE: u16 = 8;
/// `cbExtra` of the author/message attribute STTBs: one author index.
const ATTRIB_EXTRA_SIZE: u16 = 2;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// Display properties of one review e-mail message (MS-DOC 2.9.155 `MDP`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MessageDisplayProperties {
    /// Creation time of the message, when the packed DTTM is not an ignored
    /// zero date.
    created: Option<CommentDateTime>,
    /// Index into the document's `SttbfRMark` author table.
    author_index: i16,
}

impl MessageDisplayProperties {
    /// Creation time of the e-mail message.
    pub fn created(&self) -> Option<CommentDateTime> {
        self.created
    }

    /// Index into the `SttbfRMark` table of the message's author.
    pub fn author_index(&self) -> i16 {
        self.author_index
    }
}

/// One review e-mail message parallel to an author in `SttbfRMark`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThreadingMessage {
    /// Message identifier, stored verbatim and never contacted or opened.
    message_id: String,
    /// Display properties, absent when the message identifier is empty and
    /// the extra data is ignored per MS-DOC 2.9.230.
    display: Option<MessageDisplayProperties>,
}

impl ThreadingMessage {
    /// The message identifier. Empty when the corresponding author did not
    /// author an e-mail message.
    pub fn message_id(&self) -> &str {
        &self.message_id
    }

    /// Display properties of the message.
    pub fn display(&self) -> Option<MessageDisplayProperties> {
        self.display
    }
}

/// A document's e-mail review threading data (MS-DOC 2.9.230).
///
/// The author and message attribute/value STTBs carry no defined semantics
/// and are ignored per MS-DOC 2.9.230; they are validated structurally but
/// not exposed.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentRmdThreading {
    messages: Vec<ThreadingMessage>,
    personal_styles: Vec<String>,
}

impl DocumentRmdThreading {
    /// The review messages parallel to the document's `SttbfRMark` authors.
    pub fn messages(&self) -> &[ThreadingMessage] {
        &self.messages
    }

    /// The personal styles parallel to the document's `SttbfRMark` authors.
    pub fn personal_styles(&self) -> &[String] {
        &self.personal_styles
    }

    /// Parse the `RmdThreading` from the table stream, or `None` when the
    /// document carries no review threading data.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentRmdThreading>> {
        let Some((offset, length)) = fib.get_table_pointer(RMD_THREADING) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("RmdThreading offset does not fit in memory"))?;
        let end = start
            .checked_add(
                usize::try_from(length)
                    .map_err(|_| corrupted("RmdThreading length does not fit in memory"))?,
            )
            .ok_or_else(|| corrupted("RmdThreading range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("RmdThreading extends past the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `RmdThreading` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<DocumentRmdThreading> {
        let mut offset = 0usize;
        let (message_ids, mdps) = parse_sttb(data, &mut offset, MDP_SIZE, "SttbMessage")?;
        let (personal_styles, _) = parse_sttb(data, &mut offset, 0, "SttbStyle")?;
        // The attribute/value STTBs are ignored per MS-DOC 2.9.230 but are
        // still part of the structure and must be well-formed.
        let _ = parse_sttb(data, &mut offset, ATTRIB_EXTRA_SIZE, "SttbAuthorAttrib")?;
        let _ = parse_sttb(data, &mut offset, 0, "SttbAuthorValue")?;
        let _ = parse_sttb(data, &mut offset, ATTRIB_EXTRA_SIZE, "SttbMessageAttrib")?;
        let _ = parse_sttb(data, &mut offset, 0, "SttbMessageValue")?;
        if offset != data.len() {
            return Err(corrupted("RmdThreading contains trailing bytes"));
        }

        let messages = message_ids
            .into_iter()
            .zip(mdps)
            .map(|(message_id, extra)| {
                let display = if message_id.is_empty() {
                    None
                } else {
                    let dttm = litchi_core::binary::read_u32_le(extra, 0).map_err(|error| {
                        corrupted(format!("invalid SttbMessage MDP dttm: {error}"))
                    })?;
                    let created = decode_dttm(dttm)
                        .map_err(|_| corrupted("SttbMessage MDP contains an invalid DTTM"))?;
                    let author_index = read_u16(extra, 6, "SttbMessage MDP ibstAuthor")? as i16;
                    Some(MessageDisplayProperties {
                        created,
                        author_index,
                    })
                };
                Ok(ThreadingMessage {
                    message_id,
                    display,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(DocumentRmdThreading {
            messages,
            personal_styles,
        })
    }
}

/// Parse one extended STTB beginning at `*offset`, returning its strings, the
/// per-string extra data, and advancing `*offset` past the table.
fn parse_sttb<'a>(
    data: &'a [u8],
    offset: &mut usize,
    expected_extra: u16,
    name: &str,
) -> Result<(Vec<String>, Vec<&'a [u8]>)> {
    if read_u16(data, *offset, &format!("{name} fExtend"))? != STTB_F_EXTEND {
        return Err(corrupted(format!("{name} fExtend is not 0xFFFF")));
    }
    let count = usize::from(read_u16(data, *offset + 2, &format!("{name} cData"))?);
    let extra_size = read_u16(data, *offset + 4, &format!("{name} cbExtra"))?;
    if extra_size != expected_extra {
        return Err(corrupted(format!("{name} cbExtra is not {expected_extra}")));
    }
    *offset = offset
        .checked_add(6)
        .ok_or_else(|| corrupted(format!("{name} offset overflows")))?;

    let mut strings = Vec::with_capacity(count);
    let mut extras = Vec::with_capacity(count);
    for index in 0..count {
        let field = |what: &str| format!("{name} string {index} {what}");
        let unit_count = usize::from(read_u16(data, *offset, &field("length"))?);
        *offset = offset
            .checked_add(2)
            .ok_or_else(|| corrupted(field("offset overflows")))?;
        let byte_count = unit_count
            .checked_mul(2)
            .ok_or_else(|| corrupted(field("length overflows")))?;
        let end = offset
            .checked_add(byte_count)
            .ok_or_else(|| corrupted(field("range overflows")))?;
        let units = data
            .get(*offset..end)
            .ok_or_else(|| corrupted(field("is truncated")))?;
        let units: Vec<u16> = units
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        strings.push(
            String::from_utf16(&units).map_err(|_| corrupted(field("contains invalid UTF-16")))?,
        );
        *offset = end;
        let extra_end = end
            .checked_add(usize::from(extra_size))
            .ok_or_else(|| corrupted(field("extra range overflows")))?;
        extras.push(
            data.get(end..extra_end)
                .ok_or_else(|| corrupted(field("extra data is truncated")))?,
        );
        *offset = extra_end;
    }
    Ok((strings, extras))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sttb(cb_extra: u16, entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&cb_extra.to_le_bytes());
        for (text, extra) in entries {
            let units: Vec<u16> = text.encode_utf16().collect();
            data.extend_from_slice(&(units.len() as u16).to_le_bytes());
            for unit in units {
                data.extend_from_slice(&unit.to_le_bytes());
            }
            assert_eq!(extra.len(), usize::from(cb_extra));
            data.extend_from_slice(extra);
        }
        data
    }

    fn mdp(dttm: u32, author_index: i16) -> [u8; 8] {
        let mut extra = [0u8; 8];
        extra[..4].copy_from_slice(&dttm.to_le_bytes());
        extra[6..].copy_from_slice(&author_index.to_le_bytes());
        extra
    }

    fn rmd_threading(messages: &[(&str, &[u8])], styles: &[&str]) -> Vec<u8> {
        let mut data = sttb(MDP_SIZE, messages);
        data.extend_from_slice(&sttb(
            0,
            &styles.iter().map(|s| (*s, &[][..])).collect::<Vec<_>>(),
        ));
        data.extend_from_slice(&sttb(ATTRIB_EXTRA_SIZE, &[]));
        data.extend_from_slice(&sttb(0, &[]));
        data.extend_from_slice(&sttb(ATTRIB_EXTRA_SIZE, &[]));
        data.extend_from_slice(&sttb(0, &[]));
        data
    }

    #[test]
    fn parses_messages_and_styles() {
        // 2026-07-15 10:30, Wednesday: minute=30, hour=10, day=15, month=7,
        // year=126, weekday=3.
        let dttm = 30 | (10 << 6) | (15 << 11) | (7 << 16) | (126 << 20) | (3 << 29);
        let data = rmd_threading(
            &[
                ("<message-one@example.com>", &mdp(dttm, 0)),
                ("", &mdp(0, 0)),
            ],
            &["Reviewer A", ""],
        );
        let parsed = DocumentRmdThreading::parse_bytes(&data).unwrap();
        assert_eq!(parsed.messages().len(), 2);
        let first = &parsed.messages()[0];
        assert_eq!(first.message_id(), "<message-one@example.com>");
        let display = first.display().expect("non-empty message keeps its MDP");
        let created = display.created().expect("valid DTTM decodes");
        assert_eq!(
            (
                created.year,
                created.month,
                created.day,
                created.hour,
                created.minute
            ),
            (2026, 7, 15, 10, 30)
        );
        assert_eq!(display.author_index(), 0);
        // An empty message identifier voids the extra data.
        assert!(parsed.messages()[1].display().is_none());
        assert_eq!(parsed.personal_styles(), &["Reviewer A", ""]);
    }

    #[test]
    fn rejects_malformed_tables() {
        // Truncated before all six STTBs.
        assert!(DocumentRmdThreading::parse_bytes(&rmd_threading(&[], &[])[..20]).is_err());
        // Wrong cbExtra on SttbMessage.
        let mut wrong_extra = rmd_threading(&[], &[]);
        wrong_extra[4] = 4;
        assert!(DocumentRmdThreading::parse_bytes(&wrong_extra).is_err());
        // Missing fExtend marker on SttbStyle.
        let mut missing_marker = rmd_threading(&[], &[]);
        missing_marker[6] = 0;
        assert!(DocumentRmdThreading::parse_bytes(&missing_marker).is_err());
        // Trailing bytes after the sixth STTB.
        let mut trailing = rmd_threading(&[], &[]);
        trailing.push(0);
        assert!(DocumentRmdThreading::parse_bytes(&trailing).is_err());
        // Invalid DTTM on a non-empty message.
        let bad_dttm = mdp(60, 0); // minute 60 is out of range
        assert!(
            DocumentRmdThreading::parse_bytes(&rmd_threading(&[("<id>", &bad_dttm)], &[])).is_err()
        );
    }
}

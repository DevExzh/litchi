//! Final DOC stream assembly and OLE2 package writing.

use crate::encryption::encrypt_document_streams_for_write;
use crate::sprm_operations::*;
use crate::writer::fib::FibBuilder;
use crate::writer::font_table::FontTableBuilder;
use crate::writer::piece_table::{Piece, PieceTableBuilder};
use litchi_cfb::writer::OleWriter;

use super::super::{codec, model::*};

/// Fully assembled DOC streams before compound-file packaging.
pub(in crate::writer::core) struct DocOutputStreams {
    word_document: Vec<u8>,
    table: Vec<u8>,
    data: Vec<u8>,
}

impl Writer {
    pub(in crate::writer::core) fn validate_as_attached_glossary(&self) -> Result<(), WriteError> {
        if self.glossary_metadata.is_none() {
            return Err(WriteError::InvalidData(
                "attached DOC glossary requires glossary metadata".to_string(),
            ));
        }
        if self.attached_glossary.is_some() {
            return Err(WriteError::InvalidData(
                "attached DOC glossaries cannot contain another attached glossary".to_string(),
            ));
        }
        if self.encryption.is_some() {
            return Err(WriteError::InvalidData(
                "an attached DOC glossary cannot have independent encryption".to_string(),
            ));
        }
        if self.vba_project.is_some() {
            return Err(WriteError::InvalidData(
                "an attached DOC glossary cannot contain an independent VBA project".to_string(),
            ));
        }
        Ok(())
    }

    fn encryption_table_header_len(&self) -> Result<usize, WriteError> {
        self.encryption
            .as_ref()
            .map(|value| value.profile.table_header_len())
            .transpose()
            .map(|value| value.unwrap_or(0))
            .map_err(WriteError::InvalidData)
    }

    fn encrypt_output_streams(
        &self,
        word_document: &mut [u8],
        table_stream: &mut [u8],
        data_stream: &mut [u8],
    ) -> Result<(), WriteError> {
        let Some(encryption) = &self.encryption else {
            return Ok(());
        };
        encrypt_document_streams_for_write(
            encryption.profile,
            encryption.password.as_str(),
            word_document,
            table_stream,
            data_stream,
        )
        .map_err(WriteError::InvalidData)
    }

    fn populate_compound_document(
        &self,
        ole_writer: &mut OleWriter,
        word_document_stream: &[u8],
        table_stream: &[u8],
        data_stream: &[u8],
    ) -> Result<(), WriteError> {
        ole_writer.set_root_clsid(WORD_DOCUMENT_CLSID);

        // Preserve the conventional stream order so WordDocument occupies the
        // first regular FAT sector, followed by the table and Data streams.
        ole_writer.create_stream(&["WordDocument"], word_document_stream)?;
        ole_writer.create_stream(&["1Table"], table_stream)?;
        ole_writer.create_stream(&["Data"], data_stream)?;

        let compobj_data = crate::writer::ole_metadata::generate_compobj_stream();
        let ole_data = crate::writer::ole_metadata::generate_ole_stream();
        ole_writer.create_stream(&["\x01CompObj"], &compobj_data)?;
        ole_writer.create_stream(&["\x01Ole"], &ole_data)?;

        if let Some(project) = &self.vba_project {
            ole_writer.create_storage(&[VBA_PROJECT_STORAGE_NAME])?;
            project.write_into(ole_writer, &[VBA_PROJECT_STORAGE_NAME])?;
        }
        Ok(())
    }
}
impl Writer {
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), WriteError> {
        self.build_ole_writer()?.save(path)?;
        Ok(())
    }

    /// Build and validate the three core DOC streams.
    fn build_output_streams(&mut self) -> Result<DocOutputStreams, WriteError> {
        self.build_output_streams_with_data_prefix(Vec::new())
    }

    /// Build the DOC streams while retaining an existing shared Data prefix.
    fn build_output_streams_with_data_prefix(
        &mut self,
        data_prefix: Vec<u8>,
    ) -> Result<DocOutputStreams, WriteError> {
        if self.attached_glossary.is_some() && self.glossary_metadata.is_some() {
            return Err(WriteError::InvalidData(
                "a DOC template cannot be both glossary-only and contain an attached glossary"
                    .to_string(),
            ));
        }
        self.validate_style_references()?;
        let table_header_len = self.encryption_table_header_len()?;

        // Based on Apache POI's HWPFDocument.write() implementation

        let mut word_document_stream = Vec::new();
        let mut table_stream = vec![0u8; table_header_len];

        // Reserve space for FIB (Word 2007+ format = 1248 bytes, includes cswNew)
        let fib_placeholder = vec![0u8; 1248];
        word_document_stream.extend_from_slice(&fib_placeholder);

        // fcMin will be set to padded start of text (after 512 alignment below)

        // Build text stream and piece table
        let mut text_stream = Vec::new();
        let mut data_stream = data_prefix;
        let mut floating_anchors: Vec<(u32, FloatingAnchorKind)> = Vec::new();
        let mut current_cp = 0u32;
        let mut pieces = Vec::new();
        let mut chpx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut papx_entries: Vec<(u32, u32, Vec<u8>)> = Vec::new();
        let mut font_builder = FontTableBuilder::new();
        let revision_data = self.build_revision_writer_data()?;

        // Pad to 512-byte boundary before text
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        let text_fc_start = word_document_stream.len() as u32;
        let fc_min: u32 = text_fc_start;

        // Build one sorted list for all main-story reference characters.
        let mut main_refs: Vec<(u32, MainReferenceKind, usize)> = Vec::new();
        for (idx, entry) in self.footnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Footnote, idx));
        }
        for (idx, entry) in self.endnotes.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Endnote, idx));
        }
        for (idx, entry) in self.comments.iter().enumerate() {
            main_refs.push((entry.ref_position, MainReferenceKind::Comment, idx));
        }
        main_refs.sort_by_key(|reference| reference.0);

        let mut field_char_cps: Vec<(u32, u16)> = Vec::new();
        let mut footnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut endnote_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut comment_actual_cps: Vec<(usize, u32)> = Vec::new();
        let mut reference_inject_idx: usize = 0;

        for paragraph in &self.paragraphs {
            let fc_para_start = text_fc_start + text_stream.len() as u32;
            let mut para_chars: u32 = 0;
            let mut last_run_index_for_para: Option<usize> = None;
            for run in &paragraph.runs {
                let run_fc_start = text_fc_start + text_stream.len() as u32;
                let run_text = &run.text;
                let run_len_chars = utf16_code_unit_len(run_text)?;
                let grpprl = codec::build_revision_chpx_grpprl(
                    &run.formatting,
                    &mut font_builder,
                    revision_data.as_ref(),
                )?;
                // Pictures: append the OfficeArtWordDrawing block to the
                // Data stream and point sprmCPicLocation at it. Floating
                // pictures and shapes also record their anchor CP for the
                // PlcfSpa.
                let grpprl = if let Some(picture_index) = run.picture_index {
                    let entry = self.pictures.get(picture_index as usize).ok_or_else(|| {
                        WriteError::InvalidData(format!(
                            "DOC picture index {picture_index} is out of range"
                        ))
                    })?;
                    let pic_offset = u32::try_from(data_stream.len()).map_err(|_| {
                        WriteError::InvalidData(
                            "DOC Data stream exceeds 32-bit FC space".to_string(),
                        )
                    })?;
                    crate::writer::images::write_picture_block(
                        &entry.picture,
                        entry.shape_id,
                        &mut data_stream,
                    )?;
                    if entry.floating.is_some() {
                        floating_anchors.push((
                            current_cp + para_chars,
                            FloatingAnchorKind::Picture(picture_index),
                        ));
                    }
                    let mut grpprl = grpprl;
                    grpprl.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
                    grpprl.extend_from_slice(&pic_offset.to_le_bytes());
                    grpprl
                } else {
                    grpprl
                };
                if let Some(shape_index) = run.shape_index {
                    floating_anchors.push((
                        current_cp + para_chars,
                        FloatingAnchorKind::Shape(shape_index),
                    ));
                }

                let mut utf16_offset = 0u32;
                for ch in run_text.chars() {
                    let cp = current_cp + para_chars + utf16_offset;
                    match ch as u32 {
                        0x0013 => field_char_cps.push((cp, 0x13)),
                        0x0014 => field_char_cps.push((cp, 0x14)),
                        0x0015 => field_char_cps.push((cp, 0x15)),
                        _ => {},
                    }
                    utf16_offset += ch.len_utf16() as u32;
                }
                debug_assert_eq!(utf16_offset, run_len_chars);

                for u in run_text.encode_utf16() {
                    text_stream.extend_from_slice(&u.to_le_bytes());
                }
                let run_fc_end = run_fc_start + run_len_chars * 2;
                chpx_entries.push((run_fc_start, run_fc_end, grpprl));
                para_chars += run_len_chars;
                last_run_index_for_para = Some(chpx_entries.len() - 1);
            }

            while reference_inject_idx < main_refs.len() {
                let (ref_cp, kind, entry_idx) = main_refs[reference_inject_idx];
                if ref_cp <= current_cp + para_chars {
                    let actual_cp = current_cp + para_chars;
                    let fc_ref = text_fc_start + text_stream.len() as u32;
                    let marker = match kind {
                        MainReferenceKind::Footnote | MainReferenceKind::Endnote => 0x0002u16,
                        MainReferenceKind::Comment => 0x0005u16,
                    };
                    text_stream.extend_from_slice(&marker.to_le_bytes());
                    let fc_ref_end = fc_ref + 2;
                    let ref_grpprl = codec::build_chpx_grpprl(
                        &CharacterFormatting {
                            special: Some(true),
                            ..Default::default()
                        },
                        &mut font_builder,
                    );
                    chpx_entries.push((fc_ref, fc_ref_end, ref_grpprl));
                    para_chars += 1;
                    last_run_index_for_para = Some(chpx_entries.len() - 1);
                    match kind {
                        MainReferenceKind::Footnote => {
                            footnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Endnote => {
                            endnote_actual_cps.push((entry_idx, actual_cp));
                        },
                        MainReferenceKind::Comment => {
                            comment_actual_cps.push((entry_idx, actual_cp));
                        },
                    }
                    reference_inject_idx += 1;
                } else {
                    break;
                }
            }

            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            if let Some(last_idx) = last_run_index_for_para {
                chpx_entries[last_idx].1 += 2;
            }
            let fc_para_end = text_fc_start + text_stream.len() as u32;
            let pap_grpprl =
                codec::build_revision_papx_grpprl(&paragraph.formatting, revision_data.as_ref())?;
            papx_entries.push((fc_para_start, fc_para_end, pap_grpprl));

            let fc_offset = fc_para_start;
            pieces.push(Piece::new(
                current_cp,
                current_cp + para_chars + 1,
                fc_offset,
                true,
            ));
            current_cp += para_chars + 1;
        }

        self.append_tables_to_main_story(
            text_fc_start,
            &mut text_stream,
            &mut current_cp,
            &mut pieces,
            &mut chpx_entries,
            &mut papx_entries,
            &mut field_char_cps,
            &mut font_builder,
            revision_data.as_ref(),
        )?;
        if current_cp == 0 {
            Self::append_empty_main_paragraph(
                text_fc_start,
                &mut text_stream,
                &mut current_cp,
                &mut pieces,
                &mut chpx_entries,
                &mut papx_entries,
            )?;
        }

        let text_length = current_cp;

        footnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        endnote_actual_cps.sort_by_key(|&(idx, _)| idx);
        comment_actual_cps.sort_by_key(|&(idx, _)| idx);
        let ftn_ref_cps: Vec<u32> = footnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let edn_ref_cps: Vec<u32> = endnote_actual_cps.iter().map(|&(_, cp)| cp).collect();
        let comment_ref_cps: Vec<u32> = comment_actual_cps.iter().map(|&(_, cp)| cp).collect();

        let footnote_plcfs = Self::build_note_story(
            &self.footnotes,
            &ftn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )?;

        // Header pictures: append their OfficeArtWordDrawing blocks to the
        // Data stream so the header story can point sprmCPicLocation at them.
        let mut header_pic_offsets: Vec<u32> = Vec::with_capacity(self.header_pictures.len());
        for entry in &self.header_pictures {
            let pic_offset = u32::try_from(data_stream.len()).map_err(|_| {
                WriteError::InvalidData("DOC Data stream exceeds 32-bit FC space".to_string())
            })?;
            crate::writer::images::write_picture_block(
                &entry.picture,
                entry.shape_id,
                &mut data_stream,
            )?;
            header_pic_offsets.push(pic_offset);
        }

        let header_plcfhdd = self.build_header_story(
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
            &header_pic_offsets,
        )?;

        let comment_story = Self::build_comment_story(
            &self.comments,
            &comment_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )?;

        let endnote_plcfs = Self::build_note_story(
            &self.endnotes,
            &edn_ref_cps,
            text_length,
            text_fc_start,
            &mut text_stream,
            &mut chpx_entries,
            &mut papx_entries,
            &mut pieces,
            &mut current_cp,
            &mut font_builder,
        )?;
        // Build textbox story (appends textbox text after the endnote story).
        // Entry order follows the anchor CPs so the FTXBXS indices match the
        // ClientTextbox TXIDs emitted into the drawing group below.
        floating_anchors.sort_by_key(|&(anchor_cp, _)| anchor_cp);
        let textbox_shapes: Vec<&WriterShape> = floating_anchors
            .iter()
            .filter_map(|&(_, kind)| match kind {
                FloatingAnchorKind::Shape(index) => {
                    let entry = &self.shapes[index as usize];
                    entry.text.as_ref().map(|_| entry)
                },
                FloatingAnchorKind::Picture(_) => None,
            })
            .collect();
        let mut txbx_start_cps: Vec<u32> = Vec::new();
        let mut ccp_txbx = 0u32;
        if !textbox_shapes.is_empty() {
            let txbx_story_start_cp = current_cp;
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            for entry in &textbox_shapes {
                let text = entry.text.as_deref().expect("filtered on text presence");
                txbx_start_cps.push(current_cp - txbx_story_start_cp);
                // '\n' (and '\r' / "\r\n") separate plain-text paragraphs.
                for paragraph in text.replace("\r\n", "\n").replace('\r', "\n").split('\n') {
                    let para_len = utf16_code_unit_len(paragraph)?;
                    for unit in paragraph.encode_utf16() {
                        text_stream.extend_from_slice(&unit.to_le_bytes());
                    }
                    text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                    current_cp += para_len + 1;
                }
                // Trailing CR of this text box's text, as Word writes.
                text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
                current_cp += 1;
            }
            // Story-final CR, included in ccpTxbx.
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            current_cp += 1;
            ccp_txbx = current_cp - txbx_story_start_cp;
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((fc_story_start, fc_story_end, Vec::new()));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                txbx_story_start_cp,
                current_cp,
                fc_story_start,
                true,
            ));
        }

        // Build header textbox story (after the main textbox story). Entry
        // order follows the header-story anchors so the FTXBXS indices match
        // the ClientTextbox TXIDs emitted into the header drawing below.
        let header_textbox_ids: Vec<u32> = header_plcfhdd
            .as_ref()
            .map(|header| {
                header
                    .shape_anchor_cps
                    .iter()
                    .filter_map(|&(_, kind)| match kind {
                        FloatingAnchorKind::Shape(index) => {
                            Some(self.header_shapes[index as usize].shape_id)
                        },
                        FloatingAnchorKind::Picture(_) => None,
                    })
                    .collect()
            })
            .unwrap_or_default();
        let header_texts: Vec<&str> = header_plcfhdd
            .as_ref()
            .map(|header| {
                header
                    .shape_anchor_cps
                    .iter()
                    .filter_map(|&(_, kind)| match kind {
                        FloatingAnchorKind::Shape(index) => {
                            self.header_shapes[index as usize].text.as_deref()
                        },
                        FloatingAnchorKind::Picture(_) => None,
                    })
                    .collect()
            })
            .unwrap_or_default();
        let mut hdr_txbx_start_cps: Vec<u32> = Vec::new();
        let mut ccp_hdr_txbx = 0u32;
        if !header_texts.is_empty() {
            let hdr_story_start_cp = current_cp;
            let fc_story_start = text_fc_start + text_stream.len() as u32;
            let (start_cps, ccp) =
                codec::write_textbox_story_text(&header_texts, &mut text_stream, &mut current_cp)?;
            hdr_txbx_start_cps = start_cps;
            ccp_hdr_txbx = ccp;
            let fc_story_end = text_fc_start + text_stream.len() as u32;
            chpx_entries.push((fc_story_start, fc_story_end, Vec::new()));
            papx_entries.push((
                fc_story_start,
                fc_story_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(
                hdr_story_start_cp,
                current_cp,
                fc_story_start,
                true,
            ));
        }

        let bookmark_tables = Self::build_bookmark_tables(&self.bookmarks, current_cp)?;
        let smart_tag_tables = crate::writer::smart_tags::build_tables(
            &self.smart_tags,
            &self.smart_tag_recognizer_ranges,
            current_cp,
        )?;

        // Mandatory trailing paragraph mark when ANY subdocument exists (same as save()).
        let has_subdocs = footnote_plcfs.is_some()
            || header_plcfhdd.is_some()
            || comment_story.is_some()
            || endnote_plcfs.is_some()
            || ccp_txbx > 0;
        if has_subdocs {
            let fc_trailing = text_fc_start + text_stream.len() as u32;
            text_stream.extend_from_slice(&0x000Du16.to_le_bytes());
            let fc_trailing_end = fc_trailing + 2;
            chpx_entries.push((fc_trailing, fc_trailing_end, Vec::new()));
            papx_entries.push((
                fc_trailing,
                fc_trailing_end,
                codec::build_papx_grpprl(&ParagraphFormatting::default()),
            ));
            pieces.push(Piece::new(current_cp, current_cp + 1, fc_trailing, true));
            current_cp += 1;
        }
        let proofing_maximum_cp = current_cp
            .checked_add(if has_subdocs { 1 } else { 2 })
            .ok_or_else(|| {
                WriteError::InvalidData("document-parts proofing CP ceiling overflows".into())
            })?;

        let mut fib = FibBuilder::new();
        fib.set_main_text(0, text_length);
        if let Some((_, _, ftn_cp)) = &footnote_plcfs {
            fib.set_ccp_ftn(*ftn_cp);
        }
        if let Some(header) = &header_plcfhdd {
            fib.set_ccp_hdd(header.char_count);
        }
        if let Some(comment) = &comment_story {
            fib.set_ccp_atn(comment.char_count);
        }
        if let Some((_, _, edn_cp)) = &endnote_plcfs {
            fib.set_ccp_edn(*edn_cp);
        }
        if ccp_txbx > 0 {
            fib.set_ccp_txbx(ccp_txbx);
        }
        if ccp_hdr_txbx > 0 {
            fib.set_ccp_hdr_txbx(ccp_hdr_txbx);
        }

        let mut table_offset = table_stream.len() as u32;

        let stylesheet_data = crate::writer::stylesheet::generate_stylesheet(
            &self.styles,
            revision_data.as_ref().map(|data| &data.indexes),
        )
        .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        fib.set_stshf(table_offset, stylesheet_data.len() as u32);
        table_stream.extend_from_slice(&stylesheet_data);
        table_offset = table_stream.len() as u32;

        let mut piece_table = PieceTableBuilder::new();
        for piece in pieces {
            piece_table.add_piece(piece);
        }
        let clx_data = piece_table.generate()?;
        fib.set_clx(table_offset, clx_data.len() as u32);
        table_stream.extend_from_slice(&clx_data);
        table_offset = table_stream.len() as u32;

        // DocumentProperties
        let mut doc_grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            doc_grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            doc_grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            doc_grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            doc_grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            doc_grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            doc_grpf_ihdt |= 0x20;
        }
        let facing_pages = self.header_even.is_some() || self.footer_even.is_some();
        let dop_data = crate::writer::dop::generate_dop(
            facing_pages,
            doc_grpf_ihdt,
            !smart_tag_tables.is_empty(),
        );
        fib.set_dop(table_offset, dop_data.len() as u32);
        table_stream.extend_from_slice(&dop_data);
        table_offset = table_stream.len() as u32;
        table_offset = crate::writer::auxiliary_strings::append_auxiliary_string_tables(
            &mut fib,
            &mut table_stream,
            &self.associated_strings,
            self.saved_by_table.as_ref(),
            table_offset,
        )?;
        table_offset = crate::writer::glossary::append_glossary_tables(
            &mut fib,
            &mut table_stream,
            self.glossary_metadata.as_ref(),
            table_offset,
            text_length,
            &text_stream,
        )?;

        // Write PlcfHdd if present
        if let Some(header) = &header_plcfhdd {
            fib.set_plcfhdd(table_offset, header.plcfhdd.len() as u32);
            table_stream.extend_from_slice(&header.plcfhdd);
            table_offset = table_stream.len() as u32;
            if !header.fields.is_empty() {
                fib.set_plcffld_hdr(table_offset, header.fields.len() as u32);
                table_stream.extend_from_slice(&header.fields);
                table_offset = table_stream.len() as u32;
            }
        }

        // Write footnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &footnote_plcfs {
            fib.set_plcffnd_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcffnd_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        // Write endnote PLCFs if present
        if let Some((ref_bytes, txt_bytes, _)) = &endnote_plcfs {
            fib.set_plcfend_ref(table_offset, ref_bytes.len() as u32);
            table_stream.extend_from_slice(ref_bytes);
            table_offset = table_stream.len() as u32;

            fib.set_plcfend_txt(table_offset, txt_bytes.len() as u32);
            table_stream.extend_from_slice(txt_bytes);
            table_offset = table_stream.len() as u32;
        }

        if let Some(comment) = &comment_story {
            Self::append_comment_tables(&mut fib, &mut table_stream, comment);
            table_offset = table_stream.len() as u32;
        }
        if let Some(bookmarks) = &bookmark_tables {
            Self::append_bookmark_tables(&mut fib, &mut table_stream, bookmarks);
            table_offset = table_stream.len() as u32;
        }
        if !smart_tag_tables.is_empty() {
            Self::append_smart_tag_tables(&mut fib, &mut table_stream, &smart_tag_tables);
            table_offset = table_stream.len() as u32;
        }
        if let Some(revisions) = &revision_data {
            Self::append_revision_author_table(&mut fib, &mut table_stream, revisions);
            table_offset = table_stream.len() as u32;
        }
        table_offset = crate::writer::proofing::append_proofing_tables(
            &mut fib,
            &mut table_stream,
            &self.proofing_tables,
            table_offset,
            proofing_maximum_cp,
        )?;

        // Write PlcfFldMom if there are field characters
        if !field_char_cps.is_empty() {
            let main_text_bytes = usize::try_from(text_length)
                .ok()
                .and_then(|value| value.checked_mul(2))
                .and_then(|length| text_stream.get(..length))
                .ok_or_else(|| {
                    WriteError::InvalidData(
                        "DOC main field story exceeds the text stream".to_string(),
                    )
                })?;
            let plcffld = crate::writer::fields::build_plcffld(
                &field_char_cps,
                text_length,
                main_text_bytes,
            )?;
            if !plcffld.is_empty() {
                fib.set_plcffld_mom(table_offset, plcffld.len() as u32);
                table_stream.extend_from_slice(&plcffld);
                table_offset = table_stream.len() as u32;
            }
        }

        // Write numbering tables if present
        if !self.numbering.is_empty() {
            let (plflst_header, lvl_data) = self.numbering.build_plflst()?;
            fib.set_plflst(table_offset, plflst_header.len() as u32);
            table_stream.extend_from_slice(&plflst_header);
            table_stream.extend_from_slice(&lvl_data);
            table_offset = table_stream.len() as u32;

            let plflfo = self.numbering.build_plflfo();
            fib.set_plflfo(table_offset, plflfo.len() as u32);
            table_stream.extend_from_slice(&plflfo);
            table_offset = table_stream.len() as u32;

            if let Some(list_names) = self.numbering.build_sttb_list_names()? {
                fib.set_sttb_list_names(table_offset, list_names.len() as u32);
                table_stream.extend_from_slice(&list_names);
                table_offset = table_stream.len() as u32;
            }
            if let Some(list_templates) = self.numbering.build_sttb_rgtplc()? {
                fib.set_sttb_rgtplc(table_offset, list_templates.len() as u32);
                table_stream.extend_from_slice(&list_templates);
                table_offset = table_stream.len() as u32;
            }
        }

        // 6-8. Bin tables and section table written AFTER FKPs (need page numbers).

        let font_table = font_builder.generate();
        fib.set_sttbfffn(table_offset, font_table.len() as u32);
        table_stream.extend_from_slice(&font_table);

        // Append text and write FKPs
        word_document_stream.extend_from_slice(&text_stream);

        // Capture fcMac AFTER text, BEFORE FKPs (POI line 703)
        let fc_mac_value = word_document_stream.len() as u32;

        // Write FKPs to WordDocument stream at 512-byte aligned offsets
        let current_size = word_document_stream.len();
        let padding_needed = (512 - (current_size % 512)) % 512;
        word_document_stream.resize(current_size + padding_needed, 0);

        // ── CHPX FKPs (multi-page) ──
        let chpx_first_page = (word_document_stream.len() / 512) as u32;
        let mut chpx_builder = crate::writer::fkp::ChpxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &chpx_entries {
            chpx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let chpx_pages = chpx_builder.generate_pages()?;
        for page in &chpx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── PAPX FKPs (multi-page) ──
        let papx_first_page = (word_document_stream.len() / 512) as u32;
        let mut papx_builder = crate::writer::fkp::PapxFkpBuilder::new();
        for (fc_s, fc_e, grpprl) in &papx_entries {
            papx_builder.add_entry(*fc_s, *fc_e, grpprl.clone());
        }
        let papx_pages = papx_builder.generate_pages()?;
        for page in &papx_pages.pages {
            word_document_stream.extend_from_slice(page);
        }

        // ── Write bin tables to table stream ──
        let chpx_bin_table = crate::writer::bin_table::generate_bin_table_from_pages(
            &chpx_pages.ranges,
            chpx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_chpx(table_offset, chpx_bin_table.len() as u32);
        table_stream.extend_from_slice(&chpx_bin_table);

        let papx_bin_table = crate::writer::bin_table::generate_bin_table_from_pages(
            &papx_pages.ranges,
            papx_first_page,
        );
        table_offset = table_stream.len() as u32;
        fib.set_plcfbte_papx(table_offset, papx_bin_table.len() as u32);
        table_stream.extend_from_slice(&papx_bin_table);

        // Write SEPX to WordDocument stream (after text and FKPs)
        let sepx_offset = word_document_stream.len() as u32;
        let mut grpf_ihdt: u8 = 0;
        if self.header_even.is_some() {
            grpf_ihdt |= 0x01;
        }
        if self.header_odd.is_some() {
            grpf_ihdt |= 0x02;
        }
        if self.footer_even.is_some() {
            grpf_ihdt |= 0x04;
        }
        if self.footer_odd.is_some() {
            grpf_ihdt |= 0x08;
        }
        if self.header_first.is_some() {
            grpf_ihdt |= 0x10;
        }
        if self.footer_first.is_some() {
            grpf_ihdt |= 0x20;
        }
        let first_page = self.header_first.is_some() || self.footer_first.is_some();
        let section_revision = self
            .section_formatting_revision
            .as_ref()
            .map(|revision| {
                Ok::<_, WriteError>((
                    revision_data
                        .as_ref()
                        .expect("section revisions initialize revision writer data")
                        .indexes[&revision.author],
                    pack_dttm(revision.timestamp)?,
                ))
            })
            .transpose()?;
        let sepx_data = crate::writer::section::generate_sepx_with_properties(
            first_page,
            grpf_ihdt,
            section_revision,
            self.section_columns.as_ref(),
            self.section_right_to_left,
            self.section_text_flow,
            self.section_page_borders.as_ref(),
        )
        .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        word_document_stream.extend_from_slice(&sepx_data);

        // Write section table to table stream
        let total_cp = current_cp;
        let section_table = crate::writer::section::generate_section_table(total_cp, sepx_offset);
        table_offset = table_stream.len() as u32;
        fib.set_plcfsed(table_offset, section_table.len() as u32);
        table_stream.extend_from_slice(&section_table);

        // Floating pictures and shapes: shape position tables (PlcfSpaMom /
        // PlcfSpaHdr), the textbox story PLCs, and the drawing group
        // (fcDggInfo OfficeArtContent) that anchors the shapes to the
        // document's drawing layer.
        let header_anchor_cps: &[(u32, FloatingAnchorKind)] = header_plcfhdd
            .as_ref()
            .map(|header| header.shape_anchor_cps.as_slice())
            .unwrap_or(&[]);
        if !floating_anchors.is_empty() || !header_anchor_cps.is_empty() {
            table_offset = table_stream.len() as u32;
            let floating_shapes: Vec<crate::writer::images::FloatingShapeInfo<'_>> =
                floating_anchors
                    .iter()
                    .map(|&(anchor_cp, kind)| match kind {
                        FloatingAnchorKind::Picture(picture_index) => {
                            let entry = &self.pictures[picture_index as usize];
                            crate::writer::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: crate::writer::images::FloatingShapeContent::Picture(
                                    &entry.picture,
                                ),
                                width_twips: entry.picture.width_twips(),
                                height_twips: entry.picture.height_twips(),
                                position: entry.floating.as_ref().expect(
                                    "floating anchors are only recorded for floating pictures",
                                ),
                                text: None,
                            }
                        },
                        FloatingAnchorKind::Shape(shape_index) => {
                            let entry = &self.shapes[shape_index as usize];
                            crate::writer::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: crate::writer::images::FloatingShapeContent::Primitive(
                                    &entry.shape,
                                ),
                                width_twips: entry.shape.width_twips(),
                                height_twips: entry.shape.height_twips(),
                                position: &entry.position,
                                text: entry.text.as_deref(),
                            }
                        },
                    })
                    .collect();
            let header_floating_shapes: Vec<crate::writer::images::FloatingShapeInfo<'_>> =
                header_anchor_cps
                    .iter()
                    .map(|&(anchor_cp, kind)| match kind {
                        FloatingAnchorKind::Shape(shape_index) => {
                            let entry = &self.header_shapes[shape_index as usize];
                            crate::writer::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: crate::writer::images::FloatingShapeContent::Primitive(
                                    &entry.shape,
                                ),
                                width_twips: entry.shape.width_twips(),
                                height_twips: entry.shape.height_twips(),
                                position: &entry.position,
                                text: entry.text.as_deref(),
                            }
                        },
                        FloatingAnchorKind::Picture(picture_index) => {
                            let entry = &self.header_pictures[picture_index as usize];
                            crate::writer::images::FloatingShapeInfo {
                                anchor_cp,
                                shape_id: entry.shape_id,
                                content: crate::writer::images::FloatingShapeContent::Picture(
                                    &entry.picture,
                                ),
                                width_twips: entry.picture.width_twips(),
                                height_twips: entry.picture.height_twips(),
                                position: entry
                                    .floating
                                    .as_ref()
                                    .expect("header pictures always have a floating position"),
                                text: None,
                            }
                        },
                    })
                    .collect();
            if !txbx_start_cps.is_empty() {
                let txbx_shape_ids: Vec<u32> =
                    textbox_shapes.iter().map(|entry| entry.shape_id).collect();
                let plcf_txbx = crate::writer::shapes::build_plcf_txbx_txt(
                    &txbx_shape_ids,
                    &txbx_start_cps,
                    ccp_txbx,
                );
                fib.set_plcftxbx_txt(table_offset, plcf_txbx.len() as u32);
                table_stream.extend_from_slice(&plcf_txbx);
                table_offset = table_stream.len() as u32;
            }
            if !hdr_txbx_start_cps.is_empty() {
                let plcf_hdr_txbx = crate::writer::shapes::build_plcf_txbx_txt(
                    &header_textbox_ids,
                    &hdr_txbx_start_cps,
                    ccp_hdr_txbx,
                );
                fib.set_plcf_hdr_txbx_txt(table_offset, plcf_hdr_txbx.len() as u32);
                table_stream.extend_from_slice(&plcf_hdr_txbx);
                table_offset = table_stream.len() as u32;
            }
            if !floating_shapes.is_empty() {
                let plcf_spa = crate::writer::images::build_plcf_spa(&floating_shapes, text_length);
                fib.set_plc_spa_mom(table_offset, plcf_spa.len() as u32);
                table_stream.extend_from_slice(&plcf_spa);
                table_offset = table_stream.len() as u32;
            }
            if !header_floating_shapes.is_empty() {
                let header_char_count = header_plcfhdd
                    .as_ref()
                    .map(|header| header.char_count)
                    .unwrap_or(0);
                let plcf_spa_hdr = crate::writer::images::build_plcf_spa(
                    &header_floating_shapes,
                    header_char_count,
                );
                fib.set_plc_spa_hdr(table_offset, plcf_spa_hdr.len() as u32);
                table_stream.extend_from_slice(&plcf_spa_hdr);
                table_offset = table_stream.len() as u32;
            }

            let total_shapes = (self.pictures.len() + self.shapes.len()) as u32;
            let dgg_info = crate::writer::images::build_dgg_info(
                &floating_shapes,
                &header_floating_shapes,
                total_shapes,
            )?;
            fib.set_dgg_info(table_offset, dgg_info.len() as u32);
            table_stream.extend_from_slice(&dgg_info);
        }

        // Set FibBase fields
        let cb_mac = word_document_stream.len() as u32;
        fib.set_base_fields(fc_min, fc_mac_value, cb_mac);
        let fib_data = fib.generate()?;
        word_document_stream[0..fib_data.len()].copy_from_slice(&fib_data);

        // Ensure both streams are large (>= 4096) so WordDocument is allocated in regular FAT
        fn pad_to_4096(stream: &mut Vec<u8>) {
            let remainder = stream.len() % 4096;
            if remainder != 0 {
                let padding = 4096 - remainder;
                stream.resize(stream.len() + padding, 0);
            }
        }
        pad_to_4096(&mut word_document_stream);
        pad_to_4096(&mut table_stream);

        // POI writes a zero-filled Data stream when the document has no pictures.
        let mut data_stream = if data_stream.is_empty() {
            vec![0u8; 4096]
        } else {
            pad_to_4096(&mut data_stream);
            data_stream
        };
        if let Some(glossary) = self.attached_glossary.as_mut() {
            glossary.validate_as_attached_glossary()?;
            let data_prefix = std::mem::take(&mut data_stream);
            let mut glossary_streams =
                glossary.build_output_streams_with_data_prefix(data_prefix)?;
            crate::writer::attached_glossary::merge_attached_glossary(
                &mut word_document_stream,
                &mut table_stream,
                &mut glossary_streams.word_document,
                &mut glossary_streams.table,
            )?;
            data_stream = glossary_streams.data;
        }
        self.encrypt_output_streams(
            &mut word_document_stream,
            &mut table_stream,
            &mut data_stream,
        )?;

        Ok(DocOutputStreams {
            word_document: word_document_stream,
            table: table_stream,
            data: data_stream,
        })
    }

    /// Build the complete compound document after validating every staged structure.
    fn build_ole_writer(&mut self) -> Result<OleWriter, WriteError> {
        let streams = self.build_output_streams()?;
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(
            &mut ole_writer,
            &streams.word_document,
            &streams.table,
            &streams.data,
        )?;
        Ok(ole_writer)
    }

    /// Write the document to a seekable output.
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), WriteError> {
        self.build_ole_writer()?.write_to(writer)?;
        Ok(())
    }

    // Helper methods for DOC writer:
    // The following are implemented via the modular components:
    // - Generating FIB structure (File Information Block)
    // - Building piece table for text storage
    // - Generating SPRM sequences for character formatting (CHP)
    // - Generating SPRM sequences for paragraph formatting (PAP)
    // - Building FKP (Formatted Disk Page) structures
    // - Generating table properties (TAP)
    // - Encoding text to Word's internal format
    // - Managing style definitions
    // - Font table generation
}

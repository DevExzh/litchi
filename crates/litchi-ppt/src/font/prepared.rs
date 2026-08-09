//! Optional checked bridge from prepared OpenType programs to inert PPT EOT facets.

use super::{EmbeddedFont, Facet, FontCollection, FontCollections, Scope};
use litchi_fonts::{FontError, Style};

pub(crate) const fn facet_for_style(style: Style) -> Facet {
    match style {
        Style::Regular => Facet::Plain,
        Style::Bold => Facet::Bold,
        Style::Italic => Facet::Italic,
        Style::BoldItalic => Facet::BoldItalic,
    }
}

#[allow(
    clippy::expect_used,
    reason = "callers preflight the prepared-font ordinal before staging, so the lookup cannot fail"
)]
pub(crate) fn stage_facet(
    collection: &mut FontCollection,
    index: u16,
    facet: Facet,
    data: Vec<u8>,
) -> Option<EmbeddedFont> {
    let font = collection
        .get_mut(index)
        .expect("prepared-font ordinal is preflighted before encoding");
    let replacement = EmbeddedFont::from_preserved(facet, data);
    if let Some(position) = font
        .embedded_fonts
        .iter()
        .position(|value| value.style == facet as u8)
    {
        return Some(std::mem::replace(
            &mut font.embedded_fonts[position],
            replacement,
        ));
    }
    let position = font
        .embedded_fonts
        .partition_point(|value| value.style < facet as u8);
    font.embedded_fonts.insert(position, replacement);
    None
}

pub(crate) fn restore_encoded(
    font: &mut litchi_fonts::Prepared,
    encoded: Vec<u8>,
    limits: litchi_fonts::embedding::powerpoint::Limits,
) -> Result<(), FontError> {
    litchi_fonts::embedding::powerpoint::restore_with(font, encoded, limits)
}

#[allow(
    clippy::expect_used,
    reason = "the facet was staged into this collection by `stage_facet` immediately before rollback, so the owner and facet lookups cannot fail"
)]
pub(crate) fn restore_staged(
    font: &mut litchi_fonts::Prepared,
    candidate: &mut FontCollections,
    scope: Scope,
    index: u16,
    facet: Facet,
    limits: litchi_fonts::embedding::powerpoint::Limits,
) -> Result<(), FontError> {
    let owner = candidate
        .collection_mut(scope)
        .and_then(|collection| collection.get_mut(index))
        .expect("staged prepared-font owner remains available for rollback");
    let position = owner
        .embedded_fonts
        .iter()
        .position(|value| value.style == facet as u8)
        .expect("staged prepared facet remains available for rollback");
    let encoded = owner.embedded_fonts.remove(position).data;
    match encoded.try_unwrap_vec() {
        Ok(payload) => restore_encoded(font, payload, limits),
        Err(shared) => {
            restore_encoded(font, shared.as_slice().to_vec(), limits)?;
            Err(FontError::EmbeddingFailed(
                "staged PowerPoint font unexpectedly acquired a shared owner".into(),
            ))
        },
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::super::{EotIntent, EotLimits, Facet, PreparedFont, Scope, Snapshot};
    use crate::Writer;
    use litchi_fonts::{
        Charset, Family, FontProperties, License, Panose, Pitch, Signature, Style,
        embedding::powerpoint::View,
    };
    use std::io::Cursor;

    #[test]
    fn writer_encodes_style_and_publishes_an_inert_facet() {
        let mut writer = Writer::new();
        let mut prepared = prepared_font(0, Style::Bold, false);
        let replaced = writer
            .set_prepared_font(
                Scope::Base,
                0,
                &mut prepared,
                EotIntent::PreviewPrint,
                EotLimits::default(),
            )
            .unwrap();

        assert!(replaced.is_none());
        assert!(prepared.data.is_empty());
        let font = writer.font_collections().get_base(0).unwrap();
        assert_eq!(font.font_flags & 1, 0);
        let facet = font.facet(Facet::Bold).unwrap();
        assert_eq!(
            View::parse(facet.bytes()).unwrap().font_data(),
            test_sfnt(0)
        );
    }

    #[test]
    fn license_and_intent_failure_leave_writer_and_transaction_unchanged() {
        let mut writer = Writer::new();
        let before = writer.font_collections().clone();
        let mut denied = prepared_font(0x0004, Style::Regular, false);
        let source_program = denied.data.clone();
        assert!(
            writer
                .set_prepared_font(
                    Scope::Base,
                    0,
                    &mut denied,
                    EotIntent::Editable,
                    EotLimits::default(),
                )
                .is_err()
        );
        assert_eq!(writer.font_collections(), &before);
        assert_eq!(denied.data, source_program);

        let source = snapshot();
        let mut transaction = source.edit().unwrap();
        let fonts_before = transaction.fonts().clone();
        let mut denied_facet = prepared_font(0x0004, Style::Regular, false);
        let denied_source_program = denied_facet.data.clone();
        assert!(
            transaction
                .set_prepared_facet(
                    Scope::Base,
                    0,
                    &mut denied_facet,
                    EotIntent::Editable,
                    EotLimits::default(),
                )
                .is_err()
        );
        assert_eq!(transaction.fonts(), &fonts_before);
        assert!(transaction.changes().is_empty());
        assert_eq!(denied_facet.data, denied_source_program);
    }

    #[test]
    fn transaction_publishes_prepared_font_and_synchronizes_subset_state() {
        let source = snapshot();
        let mut transaction = source.edit().unwrap();
        let mut prepared = prepared_font(0, Style::Italic, true);
        transaction
            .set_prepared_facet(
                Scope::Base,
                0,
                &mut prepared,
                EotIntent::PreviewPrint,
                EotLimits::default(),
            )
            .unwrap();
        let commit = transaction.commit().unwrap();
        let font = commit.fonts().get_base(0).unwrap();
        assert!(font.embedded_subset);
        assert_eq!(font.font_flags & 1, 1);
        let view = View::parse(font.facet(Facet::Italic).unwrap().bytes()).unwrap();
        assert!(view.subsetted());
    }

    #[test]
    fn serialization_failure_restores_writer_catalog_and_prepared_program() {
        let mut writer = Writer::new();
        let before = writer.font_collections().clone();
        let mut prepared = prepared_font(0, Style::Regular, false);
        prepared.data.reserve(1024);
        let allocation = prepared.data.as_ptr();
        let source_program = prepared.data.clone();
        let mut ppt_limits = super::super::Limits::default();
        ppt_limits.records.max_record_bytes = 100;

        assert!(
            writer
                .set_prepared_font_with_limits(
                    Scope::Base,
                    0,
                    &mut prepared,
                    EotIntent::PreviewPrint,
                    EotLimits::default(),
                    ppt_limits,
                )
                .is_err()
        );
        assert_eq!(writer.font_collections(), &before);
        assert_eq!(prepared.data, source_program);
        assert_eq!(prepared.data.as_ptr(), allocation);
    }

    #[test]
    fn transaction_candidate_limit_failure_restores_all_state() {
        let mut source = snapshot();
        source.limits.fonts.max_embedded_bytes = 1;
        let mut transaction = source.edit().unwrap();
        let before = transaction.fonts().clone();
        let mut prepared = prepared_font(0, Style::Regular, false);
        prepared.data.reserve(1024);
        let allocation = prepared.data.as_ptr();
        let source_program = prepared.data.clone();

        assert!(
            transaction
                .set_prepared_facet(
                    Scope::Base,
                    0,
                    &mut prepared,
                    EotIntent::PreviewPrint,
                    EotLimits::default(),
                )
                .is_err()
        );
        assert_eq!(transaction.fonts(), &before);
        assert!(transaction.changes().is_empty());
        assert_eq!(prepared.data, source_program);
        assert_eq!(prepared.data.as_ptr(), allocation);
    }

    fn snapshot() -> Snapshot {
        let mut writer = Writer::new();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        Snapshot::from_bytes(output.into_inner()).unwrap()
    }

    fn prepared_font(fs_type: u16, style: Style, subsetted: bool) -> PreparedFont {
        PreparedFont {
            name: "Litchi Test".into(),
            style,
            data: test_sfnt(fs_type),
            properties: FontProperties::new(
                License::new(fs_type).unwrap(),
                Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
                Some(Charset::ANSI),
                Family::Roman,
                Pitch::Variable,
                Signature::new([1, 2, 3, 4], [5, 6]),
            ),
            subsetted,
        }
    }

    fn test_sfnt(fs_type: u16) -> Vec<u8> {
        let mut os2 = vec![0; 96];
        set_u16(&mut os2, 0, 2);
        set_u16(&mut os2, 4, 400);
        set_u16(&mut os2, 6, 5);
        set_u16(&mut os2, 8, fs_type);
        os2[32..42].copy_from_slice(&[2, 11, 6, 4, 2, 2, 2, 2, 2, 4]);
        set_u32(&mut os2, 42, 1);
        set_u32(&mut os2, 78, 1);

        let mut head = vec![0; 54];
        set_u32(&mut head, 0, 0x0001_0000);
        set_u32(&mut head, 8, 0x1234_5678);
        set_u32(&mut head, 12, 0x5f0f_3cf5);
        set_u16(&mut head, 18, 1000);

        let name = name_table(&[
            (1, "Litchi Test"),
            (2, "Regular"),
            (4, "Litchi Test Regular"),
            (5, "Version 1.0"),
        ]);
        sfnt(&[(b"OS/2", os2), (b"head", head), (b"name", name)])
    }

    fn name_table(values: &[(u16, &str)]) -> Vec<u8> {
        let string_offset = 6 + values.len() * 12;
        let mut strings = Vec::new();
        let mut records = Vec::new();
        for (id, value) in values {
            let offset = strings.len();
            for unit in value.encode_utf16() {
                strings.extend_from_slice(&unit.to_be_bytes());
            }
            records.push((*id, offset, strings.len() - offset));
        }
        let mut output = vec![0; string_offset];
        set_u16(&mut output, 2, u16::try_from(values.len()).unwrap());
        set_u16(&mut output, 4, u16::try_from(string_offset).unwrap());
        for (index, (id, offset, length)) in records.into_iter().enumerate() {
            let start = 6 + index * 12;
            set_u16(&mut output, start, 3);
            set_u16(&mut output, start + 2, 1);
            set_u16(&mut output, start + 4, 0x0409);
            set_u16(&mut output, start + 6, id);
            set_u16(&mut output, start + 8, u16::try_from(length).unwrap());
            set_u16(&mut output, start + 10, u16::try_from(offset).unwrap());
        }
        output.extend_from_slice(&strings);
        output
    }

    fn sfnt(tables: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
        let directory = 12 + tables.len() * 16;
        let mut offsets = Vec::new();
        let mut length = directory;
        for (_, table) in tables {
            length = (length + 3) & !3;
            offsets.push(length);
            length += table.len();
        }
        let mut output = vec![0; length];
        set_u32(&mut output, 0, 0x0001_0000);
        set_u16(&mut output, 4, u16::try_from(tables.len()).unwrap());
        for (index, ((tag, table), offset)) in tables.iter().zip(offsets).enumerate() {
            let record = 12 + index * 16;
            output[record..record + 4].copy_from_slice(*tag);
            set_u32(&mut output, record + 8, u32::try_from(offset).unwrap());
            set_u32(
                &mut output,
                record + 12,
                u32::try_from(table.len()).unwrap(),
            );
            output[offset..offset + table.len()].copy_from_slice(table);
        }
        output
    }

    fn set_u16(bytes: &mut [u8], offset: usize, value: u16) {
        bytes[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
    }

    fn set_u32(bytes: &mut [u8], offset: usize, value: u32) {
        bytes[offset..offset + 4].copy_from_slice(&value.to_be_bytes());
    }
}

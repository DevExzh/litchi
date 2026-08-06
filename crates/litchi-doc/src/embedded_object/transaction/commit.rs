//! Persistence and CHPX/ObjectPool rewrite helpers.

use super::super::Limits;
use super::super::codec::{
    CLX, PLCFBTE_CHPX, PLCFFLD_MOM, align512, append_table_block, corrupted, object_preview_sprms,
    object_separator_sprms, parse_bte, serialize_clx, serialize_fields,
};
use super::super::model::Editor;
use super::super::storage::{OBJECT_POOL, add_wrapped_storage, discover_targets};
use crate::package::{Error as PackageError, Result};
use crate::writer::ChpxFkpBuilder;
use litchi_ole_common::object::{Editor as ObjectEditor, Target};

impl Editor {
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    pub(in crate::embedded_object) fn add_object_storage(
        &mut self,
        target: Target,
        compound_file: Vec<u8>,
        limits: Limits,
    ) -> Result<()> {
        if self.object_pool_exists {
            self.package
                .add_storage(target, compound_file)
                .map_err(PackageError::from)?;
            return Ok(());
        }

        let object_pool_target =
            Target::new(OBJECT_POOL, [OBJECT_POOL]).map_err(PackageError::from)?;
        let wrapped = add_wrapped_storage(target.key(), compound_file, limits)?;
        self.package
            .add_storage(object_pool_target, wrapped)
            .map_err(PackageError::from)?;
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        let (targets, _) = discover_targets(&bytes, limits)?;
        self.package = ObjectEditor::open(bytes, targets, limits).map_err(PackageError::from)?;
        self.object_pool_exists = true;
        Ok(())
    }

    pub(in crate::embedded_object) fn append_object_chpx(
        &mut self,
        text_fc: u32,
        text: &[u16],
        separator: u32,
        result: u32,
        storage_id: u32,
        data_offset: u32,
    ) -> Result<()> {
        let byte_end = text_fc
            .checked_add(
                u32::try_from(text.len() * 2).map_err(|_| corrupted("text bytes overflow"))?,
            )
            .ok_or_else(|| corrupted("text FC overflow"))?;
        let sep_fc = text_fc + separator * 2;
        let result_fc = text_fc + result * 2;
        let (old_fc, old_pages) = parse_bte(&self.word, &self.table, PLCFBTE_CHPX)?;
        let fkp_start = old_fc.last().copied().unwrap_or(text_fc);
        if fkp_start > text_fc {
            return Err(corrupted(
                "new object FC overlaps the existing CHPX bin table",
            ));
        }
        let mut builder = ChpxFkpBuilder::new();
        if fkp_start < sep_fc {
            builder.add_entry(fkp_start, sep_fc, Vec::new());
        }
        builder.add_entry(sep_fc, sep_fc + 2, object_separator_sprms(storage_id));
        builder.add_entry(result_fc, result_fc + 2, object_preview_sprms(data_offset));
        if result_fc + 2 < byte_end {
            builder.add_entry(result_fc + 2, byte_end, Vec::new());
        }
        let pages = builder.generate_pages().map_err(PackageError::from)?;
        if pages.pages.len() != 1 {
            return Err(corrupted("object CHPX unexpectedly spans multiple FKPs"));
        }
        let page_offset = align512(self.word.len())?;
        self.word.resize(page_offset, 0);
        self.word.extend_from_slice(&pages.pages[0]);
        let page_number =
            u32::try_from(page_offset / 512).map_err(|_| corrupted("FKP page exceeds u32"))?;
        let mut fc = old_fc;
        let mut page_numbers = old_pages;
        if fc.is_empty() {
            fc.push(fkp_start);
        }
        fc.push(byte_end);
        page_numbers.push(page_number);
        let mut plc = Vec::new();
        for value in fc {
            plc.extend_from_slice(&value.to_le_bytes());
        }
        for value in page_numbers {
            plc.extend_from_slice(&value.to_le_bytes());
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_CHPX, &plc)?;
        Ok(())
    }

    pub(in crate::embedded_object) fn append_table_replacements(&mut self) -> Result<()> {
        let clx = serialize_clx(&self.pieces)?;
        append_table_block(&mut self.word, &mut self.table, CLX, &clx)?;
        let fields = serialize_fields(&self.fields, self.main_ccp)?;
        append_table_block(&mut self.word, &mut self.table, PLCFFLD_MOM, &fields)?;
        Ok(())
    }
}

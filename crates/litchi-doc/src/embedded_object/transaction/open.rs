//! Opening and indexing the DOC/ObjectPool snapshot.

use super::super::Limits;
use super::super::codec::{
    CLX, FIB_CCP_TEXT, FIB_FC_LCB, corrupted, discover_targets, parse_clx, parse_fields, u16_at,
    u32_at, validate_existing_fields,
};
use super::super::model::Editor;
use crate::package::{Error as PackageError, Result};
use litchi_ole_common::object::Editor as ObjectEditor;

impl Editor {
    pub fn open(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let (targets, object_pool_exists) = discover_targets(&bytes, limits)?;
        let package = ObjectEditor::open(bytes, targets, limits).map_err(PackageError::from)?;
        let word_path = vec!["WordDocument".to_string()];
        let word = package
            .stream(&word_path)
            .ok_or_else(|| corrupted("WordDocument stream is missing"))?
            .to_vec();
        if word.len() < FIB_FC_LCB + (CLX + 1) * 8 || u16_at(&word, 0)? != 0xA5EC {
            return Err(corrupted("unsupported pre-Word-97 or truncated FIB"));
        }
        let flags = u16_at(&word, 10)?;
        if flags & 0x0100 != 0 || u32_at(&word, 14)? != 0 {
            return Err(corrupted("encrypted DOC cannot be edited"));
        }
        let table_path = vec![
            if flags & 0x0200 != 0 {
                "1Table"
            } else {
                "0Table"
            }
            .to_string(),
        ];
        let table = package
            .stream(&table_path)
            .ok_or_else(|| corrupted("selected Table stream is missing"))?
            .to_vec();
        let data_path = vec!["Data".to_string()];
        let data = package.stream(&data_path).unwrap_or(&[]).to_vec();
        let main_ccp = u32_at(&word, FIB_CCP_TEXT)?;
        let pieces = parse_clx(&word, &table)?;
        if pieces.last().is_none_or(|piece| piece.end < main_ccp) {
            return Err(corrupted("piece table does not cover the main story"));
        }
        let fields = parse_fields(&word, &table, main_ccp)?;
        validate_existing_fields(&fields, main_ccp)?;
        Ok(Self {
            package,
            object_pool_exists,
            limits,
            word_path,
            table_path,
            data_path,
            word,
            table,
            data,
            pieces,
            fields,
            main_ccp,
            changed: false,
        })
    }
}

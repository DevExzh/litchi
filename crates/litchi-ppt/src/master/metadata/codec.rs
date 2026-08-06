//! Record framing for the contextual `SlideNameAtom` owner.

use super::model::Name;
use super::validation;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) fn read(context: Context, root: &Record) -> Result<Option<Name>> {
    validation::validate(context, root)?;
    root.children
        .iter()
        .find(|record| validation::is_name_record(record))
        .map(|record| Name::from_wire(&record.data))
        .transpose()
}

pub(super) fn encode(name: &Name) -> Result<Record> {
    let data = name.wire()?;
    let data_length = u32::try_from(data.len())
        .map_err(|_| Error::InvalidFormat("SlideNameAtom payload exceeds u32".into()))?;
    Ok(Record {
        record_type: RecordType::CString,
        record_type_raw: RecordType::CString.as_u16(),
        version: 0,
        instance: 3,
        data_length,
        data,
        children: Vec::new(),
    })
}

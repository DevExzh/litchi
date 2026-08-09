//! Record framing for the contextual `TemplateNameAtom` owner.

use super::model::Name;
use super::validation;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn read(context: Context, root: &Record) -> Result<Option<Name>> {
    let index = validation::template_index(context, root)?;
    index
        .map(|child_index| Name::from_wire(&root.children[child_index].data))
        .transpose()
}

pub(crate) fn encode(name: &Name) -> Result<Record> {
    let data = name.wire()?;
    let data_length = u32::try_from(data.len())
        .map_err(|_err| Error::InvalidFormat("TemplateNameAtom payload exceeds u32".into()))?;
    Ok(Record {
        record_type: RecordType::CString,
        record_type_raw: RecordType::CString.as_u16(),
        version: 0,
        instance: 2,
        data_length,
        data,
        children: Vec::new(),
    })
}

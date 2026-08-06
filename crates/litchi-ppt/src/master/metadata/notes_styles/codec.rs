//! Record framing for `RoundTripNotesMasterTextStyles12Atom`.

use super::model::Styles;
use super::validation;
use crate::consts::RecordType;
use crate::master_layout::Context;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn read(context: Context, root: &Record) -> Result<Option<Styles>> {
    let index = validation::styles_index(context, root)?;
    index
        .map(|index| Styles::from_package(root.children[index].data.clone()))
        .transpose()
}

pub(crate) fn encode(styles: &Styles) -> Result<Record> {
    let data = styles.bytes().to_vec();
    let data_length = u32::try_from(data.len())
        .map_err(|_| Error::InvalidFormat("notes-master text styles package exceeds u32".into()))?;
    Ok(Record {
        record_type: RecordType::RoundTripNotesMasterTextStyles12Atom,
        record_type_raw: RecordType::RoundTripNotesMasterTextStyles12Atom.as_u16(),
        version: 0,
        instance: 0,
        data_length,
        data,
        children: Vec::new(),
    })
}

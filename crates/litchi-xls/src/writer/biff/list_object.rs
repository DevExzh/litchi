use crate::sort_data::{Config, Parent};
use crate::{ListObject, Result};
use std::io::Write;

pub(crate) fn write_list_objects<W: Write>(
    writer: &mut W,
    tables: &[ListObject],
    sort_data: Option<&Config>,
) -> Result<()> {
    if tables.is_empty() {
        return Ok(());
    }
    writer.write_all(&crate::list_object::feature_header_record(tables)?)?;
    for table in tables {
        for record in table.to_feature_record_bytes()? {
            writer.write_all(&record)?;
        }
        for record in table.to_following_record_bytes()? {
            writer.write_all(&record)?;
        }
        if let Some(sort) = sort_data
            && matches!(sort.parent(), Parent::Table { id } if id == table.id().value())
        {
            sort.write_biff_records(writer)?;
        }
    }
    Ok(())
}

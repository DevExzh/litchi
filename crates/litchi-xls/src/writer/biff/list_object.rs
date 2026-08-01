use crate::{XlsListObject, XlsResult, XlsSortData, XlsSortParent};
use std::io::Write;

pub(crate) fn write_list_objects<W: Write>(
    writer: &mut W,
    tables: &[XlsListObject],
    sort_data: Option<&XlsSortData>,
) -> XlsResult<()> {
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
        if sort_data.is_some_and(
            |sort| matches!(sort.parent(), XlsSortParent::Table { id } if id == table.id().value()),
        ) {
            sort_data.unwrap().write_biff_records(writer)?;
        }
    }
    Ok(())
}

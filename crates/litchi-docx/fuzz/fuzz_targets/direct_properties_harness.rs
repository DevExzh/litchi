//! Shared bounded direct-property parser exercise.
pub fn exercise(data: &[u8]) -> bool {
    let Some((&selector, xml)) = data.split_first() else {
        return false;
    };
    if xml.len() > 64 * 1024 {
        return false;
    }
    match selector % 5 {
        0 => litchi_docx::Run::new(xml.to_vec())
            .direct_properties()
            .is_ok(),
        1 => {
            let paragraph = litchi_docx::Paragraph::new(xml.to_vec());
            let _ = paragraph.addressable_runs();
            paragraph.direct_properties().is_ok()
        },
        2 => litchi_docx::Table::new(xml.to_vec()).properties().is_ok(),
        3 => litchi_docx::table::Row::new(xml.to_vec())
            .properties()
            .is_ok(),
        _ => litchi_docx::Cell::new(xml.to_vec()).properties().is_ok(),
    }
}

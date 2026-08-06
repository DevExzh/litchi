use crate::writer::core::Writer;
use crate::writer::fib::FibBuilder;
use crate::writer::smart_tags::SmartTagTableData;
impl Writer {
    pub(in crate::writer::core::package) fn append_smart_tag_tables(
        fib: &mut FibBuilder,
        table_stream: &mut Vec<u8>,
        smart_tags: &SmartTagTableData,
    ) {
        if let Some(data) = &smart_tags.infos {
            let offset = table_stream.len() as u32;
            fib.set_sttbf_bkmk_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.starts {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.ends {
            let offset = table_stream.len() as u32;
            fib.set_plcf_bkl_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.factoid_data {
            let offset = table_stream.len() as u32;
            fib.set_factoid_data(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
        if let Some(data) = &smart_tags.recognizer_ranges {
            let offset = table_stream.len() as u32;
            fib.set_plcf_factoid(offset, data.len() as u32);
            table_stream.extend_from_slice(data);
        }
    }
}

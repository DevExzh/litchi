fn main() {
    println!(
        "{{\"metadata_bytes\":{},\"cell_record_bytes\":{},\"cell_bytes\":{}}}",
        std::mem::size_of::<litchi_xls::FormulaMetadata>(),
        std::mem::size_of::<litchi_xls::records::CellRecord>(),
        std::mem::size_of::<litchi_xls::Cell>(),
    );
}

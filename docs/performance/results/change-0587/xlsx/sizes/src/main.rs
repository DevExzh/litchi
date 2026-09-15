use std::mem::size_of;
use litchi_xlsx::cell::{Cell, Content, Date, ErrorValue, Number, Text, Unknown, Value};
use litchi_xlsx::{Address, ColumnIndex, Formula, Rect, RowIndex};
fn main() {
    macro_rules! p { ($($t:ty),*) => { $( println!("{:<40} {:>4}", stringify!($t), size_of::<$t>()); )* } }
    p!(Cell, Option<Cell>, Value, Option<Value>, Content, Formula, Option<Formula>, Unknown, Text, Number, Date, ErrorValue,
       Address, Option<Address>, Rect, Option<Rect>, RowIndex, ColumnIndex, Option<u32>, Option<usize>, (Address, Cell),
       litchi_xlsx::workbook::SourceCell, litchi_xlsx::workbook::SourceCellView, litchi_xlsx::cell::View<'static>,
       litchi_xlsx::row::Props, litchi_xlsx::Style, litchi_xlsx::streaming::StreamingCell<'static>);
}

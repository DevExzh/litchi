use std::borrow::Cow;
use litchi_rtf::{RtfDocument,RtfWriter};
fn main() {
 let mut d=RtfDocument::parse(r"{\rtf1\trowd\cellx1000\intbl\absw720 A\line B\cell\row}").unwrap();
 let cell=&mut d.tables_mut()[0].rows_mut()[0].cells_mut()[0];
 let before=cell.paragraphs().to_vec(); let text=cell.text().to_owned();
 println!("before text={text:?} paragraphs={before:?}");
 cell.set_text(Cow::Owned(text)).unwrap();
 println!("after paragraphs={:?}",cell.paragraphs());
 let mut bytes=Vec::new(); RtfWriter::new(&mut bytes).write_document(&d).unwrap();
 let reopened=RtfDocument::parse_bytes(&bytes).unwrap();
 println!("reopened paragraphs={:?}",reopened.tables()[0].rows()[0].cells()[0].paragraphs());
 assert_eq!(before,reopened.tables()[0].rows()[0].cells()[0].paragraphs(),"identical cell text must retain soft-line vs paragraph distinction");
}

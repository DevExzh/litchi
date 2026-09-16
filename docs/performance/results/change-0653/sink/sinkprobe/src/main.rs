//! Change 0653 witness for the DOCX text sink's cumulative namespace-binding
//! limit (`MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS`, 4096).
//!
//! Change 0664 recorded that `Document::write_text_to` refuses its
//! marker-bearing DOCX corpus with `semantic DOCX XML exceeds 4096 namespace
//! bindings` while the byte-identical marker-free control is admitted: the
//! limit counts `xmlns:` attributes cumulatively over the whole parse, and the
//! markup-compatibility writer re-declared every in-scope binding on every
//! emitted start tag, so a 33-declaration root exhausted the budget after about
//! 124 elements. Change 0653 does not touch the limit; it changes how many
//! declarations the writer emits.
//!
//! usage: sinkprobe <docx> [<docx> ...]

use std::io::Cursor;

use litchi_core::TextOutputOptions;

fn main() {
    for path in std::env::args().skip(1) {
        let mut sink = Cursor::new(Vec::new());
        let outcome = litchi_docx::Package::open(&path).and_then(|package| {
            let document = package.document()?;
            match document.write_text_to(&mut sink, TextOutputOptions::default()) {
                Ok(report) => Ok(Ok(report)),
                Err(error) => Ok(Err(error.to_string())),
            }
        });
        match outcome {
            Ok(Ok(report)) => println!("{path}\tOK\t{} bytes\t{report:?}", sink.get_ref().len()),
            Ok(Err(message)) => println!("{path}\tERR\t{message}"),
            Err(error) => println!("{path}\tOPEN-ERR\t{error}"),
        }
    }
}

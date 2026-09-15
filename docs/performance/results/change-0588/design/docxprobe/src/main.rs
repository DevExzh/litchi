//! Scratch probe for change 0588: does `Paragraph::extensions()` already depend
//! on the MCE codec's rewrite path having run? Opens one .docx and reports the
//! typed result for paragraph 0.
fn main() {
    let path = std::env::args().nth(1).expect("path");
    let package = litchi_docx::Package::open(&path).expect("open");
    let document = package.document().expect("document");
    match document.paragraph(0) {
        Ok(Some(paragraph)) => match paragraph.extensions() {
            Ok(extensions) => println!("OK\t{:?}", extensions.ids().para_id()),
            Err(error) => println!("ERR\t{error}"),
        },
        Ok(None) => println!("NO-PARAGRAPH"),
        Err(error) => println!("PARA-ERR\t{error}"),
    }
}

//! Small package fixtures shared by ODT integration tests.

use litchi_odt::core::PackageWriter;

pub(crate) fn package(mimetype: &str, files: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    for (path, bytes) in files {
        writer.add_file(path, bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

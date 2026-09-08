fn main() {
    println!("{}", std::mem::size_of::<soapberry_zip::ZipArchiveWriter<std::io::Sink>>());
}

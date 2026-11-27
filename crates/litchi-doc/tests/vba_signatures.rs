use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::raw::Reader;
use litchi_doc::{Builder, SignatureName};
use litchi_ole_common::object::{
    Edit as PackageEditor, Limits, OpenOptions as ObjectOpenOptions, Package as ObjectPackage,
    Targets,
};
use std::io::Cursor;

fn open_object(
    bytes: Vec<u8>,
    targets: Targets,
    limits: Limits,
) -> Result<PackageEditor, litchi_ole_common::Error> {
    ObjectPackage::open_with_options(
        bytes,
        ObjectOpenOptions::new().targets(targets).limits(limits),
    )?
    .edit()
}

fn object_bytes(editor: &PackageEditor) -> Vec<u8> {
    let package = editor.clone().commit().unwrap().into_package();
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap();
    bytes
}

fn word_blob() -> Vec<u8> {
    let mut bytes = vec![0; 56];
    bytes[0..2].copy_from_slice(&27u16.to_le_bytes());
    bytes[2..6].copy_from_slice(&45u32.to_le_bytes());
    bytes[6..10].copy_from_slice(&8u32.to_le_bytes());
    bytes[10..14].copy_from_slice(&2u32.to_le_bytes());
    bytes[14..18].copy_from_slice(&44u32.to_le_bytes());
    bytes[18..22].copy_from_slice(&3u32.to_le_bytes());
    bytes[22..26].copy_from_slice(&46u32.to_le_bytes());
    bytes[30..34].copy_from_slice(&49u32.to_le_bytes());
    bytes[42..46].copy_from_slice(&51u32.to_le_bytes());
    bytes[46..48].copy_from_slice(&[9, 8]);
    bytes[48..51].copy_from_slice(&[7, 6, 5]);
    bytes[55] = 0xEE;
    bytes
}

fn xst(value: &str) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut bytes = Vec::with_capacity(2 + units.len() * 2);
    bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
    bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
    bytes
}

fn stw_user() -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0xFFFFu16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&4u16.to_le_bytes());
    for name in ["Sign", "Other"] {
        bytes.extend(xst(name));
        bytes.extend_from_slice(&[0xA5, 0x5A, 0xC3, 0x3C]);
    }
    bytes.extend(word_blob());
    bytes.extend(xst("unrecognized"));
    bytes
}

#[test]
fn malformed_stw_user_error_is_deferred_and_cached() {
    let mut writer = Builder::new();
    writer.add_paragraph("Body").expect("fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .unwrap()
        .write_to(&mut output)
        .expect("fixture DOC");

    let mut package = open_object(output.into_inner(), Targets::default(), Limits::default())
        .expect("fixture package");
    let word_path = ["WordDocument".to_string()];
    let mut word = package
        .stream(&word_path)
        .expect("WordDocument stream")
        .to_vec();
    let pair = 154 + litchi_doc::parts::vba_signature::FIB_INDEX_STW_USER * 8;
    word[pair..pair + 4].copy_from_slice(&0u32.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&1u32.to_le_bytes());
    package
        .put_stream(&word_path, word)
        .expect("malformed StwUser pointer");
    let bytes = object_bytes(&package);

    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let first = document
        .vba_signatures()
        .expect_err("malformed StwUser metadata");
    assert_eq!(
        document
            .vba_signatures()
            .expect_err("cached malformed StwUser metadata")
            .to_string(),
        first.to_string()
    );
}

#[test]
fn document_facade_defers_and_reads_stw_user_signature_bytes() {
    let mut writer = Builder::new();
    writer.add_paragraph("Body").expect("fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .unwrap()
        .write_to(&mut output)
        .expect("fixture DOC");

    let mut package = open_object(output.into_inner(), Targets::default(), Limits::default())
        .expect("fixture package");
    let word_path = ["WordDocument".to_string()];
    let mut word = package
        .stream(&word_path)
        .expect("WordDocument stream")
        .to_vec();
    let fib = FileInformationBlock::parse(&word).expect("fixture FIB");
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    let source = stw_user();
    let offset = u32::try_from(table.len()).expect("table offset");
    let length = u32::try_from(source.len()).expect("StwUser length");
    table.extend_from_slice(&source);
    package
        .put_stream(&table_path, table)
        .expect("StwUser table");
    let pair = 154 + litchi_doc::parts::vba_signature::FIB_INDEX_STW_USER * 8;
    word[pair..pair + 4].copy_from_slice(&offset.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&length.to_le_bytes());
    package
        .put_stream(&word_path, word)
        .expect("StwUser pointer");
    let bytes = object_bytes(&package);

    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let signatures = document
        .vba_signatures()
        .expect("valid StwUser")
        .expect("recognized signature");
    let signature = signatures.get(SignatureName::Sign).expect("Sign variable");
    assert_eq!(signature.bytes(), word_blob().as_slice());
    assert_eq!(signature.snapshot().info().signature(), [9, 8]);
}

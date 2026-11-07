//! Inert DDE/OLE metadata on native XLS parents; no link activation or target resolution.
use litchi_xls as link;
use litchi_xls::Package;
#[path = "../../../tools/test_support/native_xls_workbook.rs"]
mod native_xls_workbook;
use native_xls_workbook::{frame, native, publish_range, records, workbook};
const PARENTS: [&str; 3] = [
    "excel-metadata.xls",
    "excel-recorded-macro.xlt",
    "excel-links-objects.xls",
];
const TARGET: &str = "Native.Server\u{3}Topic 東京";
fn unicode(value: &str, short: bool) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut out = if short {
        vec![u8::try_from(units.len()).unwrap()]
    } else {
        u16::try_from(units.len()).unwrap().to_le_bytes().to_vec()
    };
    let narrow = units.iter().all(|u| *u < 256);
    out.push(u8::from(!narrow));
    if narrow {
        out.extend(units.iter().map(|u| u8::try_from(*u).unwrap()));
    } else {
        out.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
    out
}
fn supbook(target: &str) -> Vec<u8> {
    let mut b = vec![0, 0];
    b.extend(unicode(target, false));
    frame(0x1ae, &b)
}
fn operand(value: &link::CachedValue) -> Vec<u8> {
    use link::CachedValue::*;
    match value {
        Blank => vec![0, 0xa5, 0xa5, 0xa5, 0xa5, 0xa5, 0xa5, 0xa5, 0xa5],
        Number(n) => [vec![1], n.to_le_bytes().to_vec()].concat(),
        Text(s) => [vec![2], unicode(s, false)].concat(),
        Boolean(v) => vec![4, u8::from(*v), 0, 0x91, 0x91, 0x91, 0x91, 0x91, 0x91],
        Error(link::ErrorValue::DivisionByZero) => {
            vec![0x10, 7, 0, 0x82, 0x82, 0x82, 0x82, 0x82, 0x82]
        },
        Error(_) => panic!("unexpected test error variant"),
    }
}
fn mixed() -> Vec<link::CachedValue> {
    use link::CachedValue::*;
    vec![
        Blank,
        Number(42.5),
        Text("東京🍎".into()),
        Boolean(true),
        Error(link::ErrorValue::DivisionByZero),
        Number(0.0),
    ]
}
fn name(flags: u16, storage: u32, name: &str, matrix: Option<&link::ValueMatrix>) -> Vec<u8> {
    let mut body = flags.to_le_bytes().to_vec();
    body.extend(storage.to_le_bytes());
    body.extend(unicode(name, true));
    let mut output = Vec::new();
    let mut kind = 0x23;
    if let Some(m) = matrix {
        m.validate().unwrap();
        body.push(m.last_column);
        body.extend(m.last_row.to_le_bytes());
        for value in &m.values {
            let v = operand(value);
            if body.len() + v.len() > 8224 {
                output.extend(frame(kind, &body));
                kind = 0x3c;
                body.clear();
            }
            body.extend(v);
        }
    }
    output.extend(frame(kind, &body));
    output
}
fn matrix() -> link::ValueMatrix {
    link::ValueMatrix {
        last_column: 2,
        last_row: 399,
        values: std::iter::repeat_with(mixed).take(200).flatten().collect(),
    }
}
fn addition(matrix: &link::ValueMatrix) -> Vec<u8> {
    let mut bytes = supbook(TARGET);
    bytes.extend(name(8, 0, "StdDocumentName", None));
    bytes.extend(name(2 | (44 << 5), 0, "Value", Some(matrix)));
    bytes.extend(supbook("Example.Document\u{3}source.ole"));
    bytes.extend(name(
        0x8000 | 0x10 | 4 | (2 << 5),
        0x42,
        "Object",
        Some(&link::ValueMatrix {
            last_column: 2,
            last_row: 1,
            values: mixed(),
        }),
    ));
    bytes
}
fn insertion(book: &[u8]) -> usize {
    let frames = records(book);
    if let Some(r) = frames.iter().find(|r| r.1 == 0x17) {
        return r.0;
    }
    let country = frames.iter().position(|r| r.1 == 0x8c).unwrap();
    frames[country + 1].0
}
fn owner_end(book: &[u8]) -> usize {
    let frames = records(book);
    let start = frames.iter().position(|r| r.1 == 0x1ae).unwrap();
    frames[start..]
        .iter()
        .find(|r| !matches!(r.1, 0x1ae | 0x23 | 0x59 | 0x5a | 0x17 | 0x3c))
        .unwrap()
        .0
}
fn seed(file: &str, cache: &link::ValueMatrix) -> (Package, usize, usize) {
    let bytes = native(file);
    let p = Package::from_bytes(&bytes).unwrap();
    let old = p.reader_ref().unwrap().external_links().clone();
    let book = workbook(&p);
    let at = insertion(&book);
    let q = publish_range(&p, at, at, &addition(cache));
    let read = q.reader_ref().unwrap().external_links();
    assert_eq!(
        &read.supporting_books()[..old.supporting_books().len()],
        old.supporting_books()
    );
    assert_eq!(
        &read.external_names()[..old.external_names().len()],
        old.external_names()
    );
    assert_eq!(read.sheet_references(), old.sheet_references());
    assert_eq!(native(file), bytes);
    (q, old.supporting_books().len(), old.external_names().len())
}
fn assert_matrix(name: &link::Name, expected: &link::ValueMatrix) {
    let link::NameBody::DdeOrOle {
        matrix: Some(m), ..
    } = name.body()
    else {
        panic!("cached item")
    };
    assert_eq!(m, expected);
}
#[test]
fn native_xls_dde_ole_metadata_reads_exact_flags_names_and_continued_values() {
    let m = matrix();
    for file in PARENTS {
        let (p, b, n) = seed(file, &m);
        let l = p.reader_ref().unwrap().external_links();
        assert_eq!(
            &l.supporting_books()[b..],
            [
                link::SupportingBook::DdeOrOle {
                    encoded_virtual_path: TARGET.into()
                },
                link::SupportingBook::DdeOrOle {
                    encoded_virtual_path: "Example.Document\u{3}source.ole".into()
                }
            ]
        );
        let standard = &l.external_names()[n];
        assert!(standard.standard_document_name());
        assert_eq!(
            standard.body(),
            &link::NameBody::DdeStandardDocumentName {
                name: "StdDocumentName".into()
            }
        );
        let dde = &l.external_names()[n + 1];
        assert_eq!(dde.supporting_book_index() as usize, b);
        assert!(dde.automatic() && !dde.ole_link() && !dde.picture() && !dde.displayed_as_icon());
        assert_eq!(dde.clipboard_format(), link::ClipboardFormat::UnicodeText);
        assert_matrix(dde, &m);
        let ole = &l.external_names()[n + 2];
        assert_eq!(ole.supporting_book_index() as usize, b + 1);
        assert!(ole.ole_link() && ole.picture() && ole.displayed_as_icon());
        assert_eq!(
            ole.clipboard_format(),
            link::ClipboardFormat::EnhancedMetafile
        );
        assert_eq!(
            ole.body(),
            &link::NameBody::DdeOrOle {
                storage_id: 0x42,
                name: "Object".into(),
                matrix: Some(link::ValueMatrix {
                    last_column: 2,
                    last_row: 1,
                    values: mixed()
                })
            }
        );
        assert!(records(&workbook(&p)).iter().any(|r| r.1 == 0x3c));
        let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        p.save(output.path()).unwrap();
        let q = Package::open(output.path()).unwrap();
        assert_eq!(q.to_bytes().unwrap(), p.to_bytes().unwrap());
        assert_eq!(q.reader_ref().unwrap().external_links(), l);
    }
}

fn name_span(book: &[u8], index: usize) -> (usize, usize) {
    let frames = records(book);
    let start = frames
        .iter()
        .enumerate()
        .filter(|(_, r)| r.1 == 0x23)
        .nth(index)
        .unwrap()
        .0;
    let mut end = start + 1;
    while frames[end].1 == 0x3c {
        end += 1;
    }
    (frames[start].0, frames[end].0)
}
fn authored_name(value: &str, matrix: &link::ValueMatrix) -> Vec<u8> {
    use litchi_xls::writer::{Builder, DdeOrOleItemOptions, DdeOrOleLinkOptions};
    let mut builder = Builder::new();
    builder.add_worksheet("Carrier").unwrap();
    builder
        .add_dde_or_ole_link(DdeOrOleLinkOptions {
            encoded_virtual_path: TARGET.into(),
            items: vec![DdeOrOleItemOptions {
                name: value.into(),
                automatic: true,
                picture: false,
                standard_document_name: false,
                ole_link: false,
                clipboard_format: link::ClipboardFormat::UnicodeText,
                displayed_as_icon: false,
                storage_id: 0,
                matrix: Some(matrix.clone()),
            }],
        })
        .unwrap();
    let mut bytes = Vec::new();
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    let p = Package::from_bytes(&bytes).unwrap();
    assert_matrix(
        &p.reader_ref().unwrap().external_links().external_names()[0],
        matrix,
    );
    let book = workbook(&p);
    let (a, b) = name_span(&book, 0);
    book[a..b].to_vec()
}

#[test]
fn native_xls_dde_cache_builder_replacement_rechunks_whole_values_and_restores_source() {
    let m = link::ValueMatrix {
        last_column: 255,
        last_row: 3,
        values: (0..1024)
            .map(|i| link::CachedValue::Number(f64::from(i)))
            .collect(),
    };
    let renamed = "Renamed cache 日本語🍎";
    let replacement = authored_name(renamed, &m);
    for file in PARENTS {
        let (p, _, n) = seed(file, &m);
        let before = workbook(&p);
        let (start, end) = name_span(&before, n + 1);
        let source_frames = records(&before[start..end]);
        assert_eq!(source_frames[0].2.len(), 8224);
        let frames = records(&replacement);
        assert!(frames.len() > 1);
        assert!(frames.iter().all(|r| r.2.len() <= 8224));
        let q = publish_range(&p, start, end, &replacement);
        let items = q.reader_ref().unwrap().external_links().external_names();
        assert_matrix(&items[n + 1], &m);
        let link::NameBody::DdeOrOle { name, .. } = items[n + 1].body() else {
            panic!("cache")
        };
        assert_eq!(name, renamed);
        let original = p.reader_ref().unwrap().external_links();
        assert_eq!(
            q.reader_ref().unwrap().external_links().supporting_books(),
            original.supporting_books()
        );
        assert_eq!(
            q.reader_ref().unwrap().external_links().sheet_references(),
            original.sheet_references()
        );
        for (i, item) in items.iter().enumerate() {
            if i != n + 1 {
                assert_eq!(item, &original.external_names()[i]);
            }
        }
        let (new_start, new_end) = name_span(&workbook(&q), n + 1);
        let restored = publish_range(&q, new_start, new_end, &before[start..end]);
        assert_eq!(workbook(&restored), before);
    }
}

fn link_bytes(p: &Package) -> Vec<u8> {
    let book = workbook(p);
    let start = records(&book).iter().find(|r| r.1 == 0x1ae).unwrap().0;
    book[start..owner_end(&book)].to_vec()
}
#[test]
fn native_xls_dde_ole_continued_caches_survive_ordinary_cell_publication() {
    use native_xls_workbook::{cfb_metadata, streams};
    let m = matrix();
    for file in PARENTS {
        let (p, _, _) = seed(file, &m);
        let before = p.to_bytes().unwrap();
        let owned = link_bytes(&p);
        let mut edit = p.edit().unwrap();
        let cell = litchi_xls::cell_values::Reference::new(120, 25).unwrap();
        edit.insert_cell("Sheet1", cell, litchi_xls::cell_values::Value::Number(88.0))
            .unwrap();
        let c = edit.commit().unwrap();
        let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        c.package().save(f.path()).unwrap();
        let q = Package::open(f.path()).unwrap();
        let after = q.to_bytes().unwrap();
        assert_eq!(link_bytes(&q), owned);
        assert_eq!(
            q.reader_ref().unwrap().external_links(),
            p.reader_ref().unwrap().external_links()
        );
        assert_eq!(
            q.workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(cell)
                .unwrap(),
            Some(litchi_core::sheet::CellValue::Float(88.0))
        );
        let a = streams(&before);
        let b = streams(&after);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (name, data) in a {
            if name != ["Workbook"] {
                assert_eq!(b[&name], data);
            }
        }
        assert_eq!(
            cfb_metadata::directory(&before),
            cfb_metadata::directory(&after)
        );
        assert_eq!(
            p.apply(c.patch()).unwrap().package().to_bytes().unwrap(),
            after
        );
        assert_eq!(
            q.apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            before
        );
    }
}

fn raw_workbook(p: &Package, bytes: Vec<u8>) -> Vec<u8> {
    let owner = litchi_ole_common::object::Package::open(p.to_bytes().unwrap()).unwrap();
    let mut edit = owner.edit().unwrap();
    edit.put_stream(&["Workbook".into()], bytes).unwrap();
    edit.commit().unwrap().patch().after().to_vec()
}
#[test]
fn native_xls_dde_ole_invalid_items_and_split_values_have_exact_reader_errors() {
    let bytes = native(PARENTS[0]);
    let (p, _, n) = seed(PARENTS[0], &matrix());
    let book = workbook(&p);
    let frames = records(&book);
    let primary = frames
        .iter()
        .enumerate()
        .filter(|(_, r)| r.1 == 0x23)
        .nth(n + 1)
        .unwrap()
        .0;
    let standard = frames.iter().filter(|r| r.1 == 0x23).nth(n).unwrap().0;
    let mut cases = Vec::new();
    let mut invalid = book.clone();
    invalid[standard + 4 + 8] = b'X';
    cases.push((invalid, "ExternDdeLinkNoOper body is invalid"));
    let mut invalid = book.clone();
    let moper = frames[primary].0 + 4 + 6 + unicode("Value", true).len();
    invalid[moper + 1..moper + 3].copy_from_slice(&400u16.to_le_bytes());
    cases.push((
        invalid,
        "DDE/OLE MOper matrix ended before all declared values were read",
    ));
    let mut first = frames[primary].2.clone();
    let mut second = frames[primary + 1].2.clone();
    assert_eq!(frames[primary + 1].1, 0x3c);
    let tail = first.pop().unwrap();
    second.insert(0, tail);
    let mut replacement = frame(0x23, &first);
    replacement.extend(frame(0x3c, &second));
    let a = frames[primary].0;
    let b = frames[primary + 2].0;
    assert_eq!(replacement.len(), b - a);
    let mut invalid = book.clone();
    invalid[a..b].copy_from_slice(&replacement);
    cases.push((invalid, "SerAr value is truncated or split across records"));
    for (invalid, message) in cases {
        assert!(
            matches!(Package::from_bytes(&raw_workbook(&p,invalid)),Err(litchi_xls::package::Error::Reader(link::Error::InvalidRecord{record_type:0x23,message:m})) if m==message)
        );
    }
    assert_eq!(native(PARENTS[0]), bytes);
}

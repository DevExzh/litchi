//! Compare semantic CFB directory metadata independently of sector allocation.
use std::{collections::BTreeMap, io::Cursor};
type Metadata = (u8, String, u32, u64, u64);
pub(crate) fn directory(bytes: &[u8]) -> BTreeMap<Vec<String>, Metadata> {
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap();
    fn value(e: &litchi_cfb::DirectoryEntry) -> Metadata {
        (
            e.entry_type,
            e.clsid.clone(),
            e.state_bits,
            e.creation_time,
            e.modified_time,
        )
    }
    fn walk(
        ole: &litchi_cfb::OleFile<Cursor<&[u8]>>,
        path: Vec<String>,
        out: &mut BTreeMap<Vec<String>, Metadata>,
    ) {
        for e in ole
            .list_directory_entries(&path.iter().map(String::as_str).collect::<Vec<_>>())
            .unwrap()
        {
            let mut p = path.clone();
            p.push(e.name.clone());
            out.insert(p.clone(), value(e));
            if e.entry_type == 1 {
                walk(ole, p, out);
            }
        }
    }
    let mut out = BTreeMap::new();
    out.insert(Vec::new(), value(ole.root_entry().unwrap()));
    walk(&ole, Vec::new(), &mut out);
    out
}

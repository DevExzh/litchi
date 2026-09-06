#!/usr/bin/env python3
"""Apply the reviewed long-name refinement only after all owned CPU jobs stop."""
import hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3]
p=REPO/'crates/litchi-opc/src/source_backed.rs';s=p.read_text()
assert hashlib.sha256(p.read_bytes()).hexdigest()=='5f4e4d0fd21acfd1581772b7dc078db87e175fc65607338cdac4ad0b55ee1168', 'historical draft only; current source differs'
(ROOT/'draft-history/pre-metadata-bound-source_backed.rs.txt').write_bytes(p.read_bytes())
needle='impl AuthorizedPrecompressedPart {\n'
s=s.replace(needle,needle+'''    fn fixed_memory_bytes(partname: &PackURI, content_type: &str) -> Result<u64> {
        let metadata = std::mem::size_of::<Self>()
            .checked_add(partname.as_str().len())
            .and_then(|bytes| bytes.checked_add(content_type.len()))
            .and_then(|bytes| bytes.checked_add(256))
            .ok_or_else(|| overlay_unavailable("precompressed token metadata size overflows"))?;
        let metadata = u64::try_from(metadata)
            .map_err(|_| overlay_unavailable("precompressed token metadata size exceeds u64"))?;
        Ok(metadata.max(PRECOMPRESSED_WRITER_FIXED_OVERHEAD))
    }

''',1)
s=s.replace('let fixed_memory_reservation = reserve(PRECOMPRESSED_WRITER_FIXED_OVERHEAD)?;', 'let fixed_memory_reservation = reserve(AuthorizedPrecompressedPart::fixed_memory_bytes(\n            &payload.source_partname, payload.source_content_type.as_str())?)?;',1)
a=s.index('    fn authorize_precompressed_inner')
z=s.index('        metadata\n            .compressed_size()\n            .checked_mul(2)',a)
s=s[:z]+'''        let fixed_bytes = AuthorizedPrecompressedPart::fixed_memory_bytes(
            &part.partname, &part.content_type)?;
'''+s[z:]
a=s.index('    fn authorize_precompressed_inner');head=s[:a];tail=s[a:]
tail=tail.replace('.checked_add(PRECOMPRESSED_WRITER_FIXED_OVERHEAD)', '.checked_add(fixed_bytes)',1).replace('self.reserve_topology_memory(PRECOMPRESSED_WRITER_FIXED_OVERHEAD)?', 'self.reserve_topology_memory(fixed_bytes)?',1)
s=head+tail
s=s.replace('    /// for the retained token metadata, including zero-length captures. New', '    /// for the retained token metadata, including zero-length captures and\n    /// long Part names/content types. New')
s=s.replace('        let reserve = |bytes| {', '        let reserve = |bytes| {\n            if bytes == 0 { return Ok(None); }', 1)
p.write_text(s)
p=REPO/'crates/litchi-opc/tests/source_part_transfer.rs';(ROOT/'draft-history/pre-metadata-bound-source_part_transfer.rs.txt').write_bytes(p.read_bytes())
p.write_text(p.read_text()+'''

#[test]
fn retained_empty_capture_charges_long_source_metadata() {
    let name = format!("custom/{}.bin", "x".repeat(16384));
    let uri = pack(&format!("/{name}"));
    let types = content_types(r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#);
    let root = root_relationships();
    let bytes = stored_archive(&[
        Entry { name: b"[Content_Types].xml", data: &types },
        Entry { name: b"_rels/.rels", data: &root },
        Entry { name: b"word/document.xml", data: b"<document/>" },
        Entry { name: name.as_bytes(), data: b"" },
    ]);
    let source = CaptureObservedSource::new(TestSource::new(bytes), 65536);
    let (budget, _cancel, context) = managed_context(256 * 1024);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(), ReadLimits::default(), context).unwrap();
    let (data, token) = package.part(&uri).unwrap().data_and_authorize_precompressed().unwrap();
    let capture = token.into_retained();
    drop(data);
    drop(package);
    let unit = budget.used(Resource::Memory);
    assert!(unit >= uri.as_str().len() as u64);
    let before = source.counts();
    let mut captures = vec![capture];
    for _ in 0..64 {
        match captures.last().unwrap().authorize_for_publication() {
            Ok(token) => captures.push(token.into_retained()),
            Err(OpcError::Execution(ExecutionError::ResourceLimit(limit))) if limit.resource == Resource::Memory => break,
            Err(error) => panic!("unexpected retained metadata refusal: {error}"),
        }
        assert_eq!(budget.used(Resource::Memory), unit * captures.len() as u64);
    }
    assert!(captures.len() < 64);
    assert_eq!(source.counts(), before);
    drop(captures);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}
''')

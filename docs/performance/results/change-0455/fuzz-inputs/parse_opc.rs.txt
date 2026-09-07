#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_opc::{
    AuthoredXmlFragment, OpcPackage, PackURI, ReadLimits, SourceBackedPackage, SourceTopologyPlan,
};

fn source_xml_insertion(bytes: &[u8]) -> Option<usize> {
    bytes.windows(2).rposition(|window| window == b"</")
}

fuzz_target!(|data: &[u8]| {
    // Bound both eager parsing and verified compressed/decoded materialization.
    if data.len() > 1024 * 1024 {
        return;
    }
    let limits = ReadLimits::builder()
        .max_input_bytes(1024 * 1024)
        .unwrap()
        .max_archive_members(128)
        .unwrap()
        .max_parts(128)
        .unwrap()
        .max_relationship_parts(128)
        .unwrap()
        .max_archive_compressed_bytes(1024 * 1024)
        .unwrap()
        .max_archive_entry_bytes(1024 * 1024)
        .unwrap()
        .max_archive_total_bytes(4 * 1024 * 1024)
        .unwrap()
        .max_part_bytes(1024 * 1024)
        .unwrap()
        .build()
        .unwrap();
    if let Ok(package) = SourceBackedPackage::from_vec_with_limits(data.to_vec(), limits) {
        let mut source_xml_mutated = false;
        let mut plan = SourceTopologyPlan::new();
        for part in package.iter_parts() {
            // Capture first so malformed captures are not filtered out by a
            // successful ordinary decode. Repeat through the warm cache path.
            if let Ok((cold, token)) = part.data_and_authorize_precompressed() {
                if let Ok((warm, warm_token)) = part.data_and_authorize_precompressed() {
                    assert_eq!(cold.as_bytes(), warm.as_bytes());
                    drop(warm_token);
                }
                let retained = token.into_retained();
                let copy = retained.clone();
                let first = retained.authorize_for_publication();
                let second = copy.authorize_for_publication();
                drop(first);
                drop(second);
            }

            // Exercise the opaque source-XML proof, one bounded insertion, and
            // topology replacement/addition. The insertion point is chosen
            // immediately before the final closing tag when the source has
            // one; malformed/self-closing XML simply takes the next part.
            if !source_xml_mutated
                && let Ok(source_xml) = part.source_xml()
                && let Some(insertion) = source_xml_insertion(source_xml.bytes())
                && let Ok(proof) = source_xml.checked_range(insertion..insertion, &[])
                && let Ok(fragment) = AuthoredXmlFragment::markup(b"<litchi-fuzz/>".to_vec())
                && let Ok(mut publication) = source_xml.into_publication()
                && publication.replace(proof, fragment).is_ok()
                && let Ok(edited) = publication.finish()
            {
                source_xml_mutated = true;
                let target = part.partname().clone();
                if plan
                    .try_replace_source_xml_part(target, edited.clone())
                    .is_ok()
                {
                    // A fixed, package-local URI keeps the edit bounded. A
                    // collision is harmless and still exercises the guarded
                    // addition path.
                    if let Ok(copy_name) = PackURI::new("/litchi-fuzz-source-copy.xml") {
                        let _ = plan.try_add_source_xml_part(copy_name, edited);
                    }
                }
            }
        }
        // Exercise lossless relationship appends on mutated source XML. This
        // external reference is inert; no resource is fetched. The input,
        // member, decoded-byte and relationship limits above remain in force,
        // and the discard sink does not retain an archive-sized output buffer.
        if plan
            .try_add_external_relationship(
                PackURI::new("/").unwrap(),
                "rIdLitchiFuzzAppend",
                "urn:litchi:fuzz:external",
                "https://example.invalid/fuzz",
            )
            .is_ok()
        {
            let _ = package.write_topology_to_stream(std::io::sink(), plan);
        }
    }
    // Primary entrypoint: parse a ZIP-backed OPC package from a byte slice.
    if let Ok(pkg) = OpcPackage::from_bytes_with_limits(data, limits) {
        // Exercise package-level relationships.
        let rels = pkg.rels();
        for rel in rels.iter() {
            let _ = rel.r_id();
            let _ = rel.reltype();
            let _ = rel.target_ref();
            let _ = rel.is_external();
            // Resolving target partname can fail on malformed inputs; ignore.
            let _ = rel.target_partname();
        }

        // Iterate all parts and inspect surface API.
        for part in pkg.iter_parts() {
            let _ = part.partname();
            let _ = part.content_type();
            let _ = part.blob().len();

            // Walk part-level relationships too.
            let prels = part.rels();
            for rel in prels.iter() {
                let _ = rel.r_id();
                let _ = rel.reltype();
                let _ = rel.target_ref();
                let _ = rel.is_external();
                let _ = rel.target_partname();
            }
        }

        // Try to follow the main document relationship if present.
        let _ = pkg.main_document_part();
        let _ = pkg.part_count();
    }
});

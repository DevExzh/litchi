//! Bounded source-scope hashing without copying retained payloads.

use litchi_core::patch::PatchError;
use sha2::{Digest as _, Sha256};

use super::{MAX_CLOSURE_RECORDS, PartState, Patch};

type Result<T> = std::result::Result<T, PatchError>;

pub(super) fn scope_hash_from_patch_scope(
    patch: &Patch,
    before: bool,
    maximum: usize,
) -> Result<String> {
    let lexical = patch.lexical.as_ref().ok_or_else(invalid_scope)?;
    let count = lexical
        .relationships
        .len()
        .checked_add(patch.parts.len())
        .and_then(|count| count.checked_add(1))
        .ok_or_else(invalid_scope)?;
    if count > MAX_CLOSURE_RECORDS || lexical.parts.len() > MAX_CLOSURE_RECORDS {
        return Err(invalid_scope());
    }
    // Every record needs at least a kind, a name length and a presence byte;
    // XML tokens also carry a payload length. Refuse impossible byte budgets
    // before reserving the sorting indices.
    let minimum = count
        .checked_mul(10)
        .and_then(|size| {
            lexical
                .parts
                .len()
                .checked_mul(17)
                .and_then(|xml| size.checked_add(xml))
        })
        .and_then(|size| size.checked_add(20))
        .ok_or_else(invalid_scope)?;
    charge(&mut 0, minimum, maximum)?;
    let relationships = sorted_indices(&lexical.relationships, |value| value.owner.as_str())?;
    let parts = sorted_indices(&patch.parts, |value| value.name.as_str())?;
    let xml = sorted_indices(&lexical.parts, |value| value.name.as_str())?;
    let relationship_lookup = folded_indices(&lexical.relationships, |value| value.owner.as_str())?;
    let part_lookup = folded_indices(&patch.parts, |value| value.name.as_str())?;
    // Lexical XML tokens cannot substitute for a complete scoped part state:
    // binary data and content types must also be retained in the patch.
    let protection = patch.protection.as_ref().ok_or_else(invalid_scope)?;
    for (scope, source) in [(&protection.source, true), (&protection.destination, false)] {
        for name in &scope.owned_parts {
            let index = part_lookup
                .binary_search_by(|index| {
                    folded_cmp(patch.parts[*index].name.as_str(), name.as_str())
                })
                .map_err(|_| invalid_scope())?;
            let part = &patch.parts[part_lookup[index]];
            if (if source { &part.before } else { &part.after }).is_none() {
                return Err(invalid_scope());
            }
        }
    }
    for token in &lexical.parts {
        let index = part_lookup
            .binary_search_by(|index| {
                folded_cmp(patch.parts[*index].name.as_str(), token.name.as_str())
            })
            .map_err(|_| invalid_scope())?;
        let part = &patch.parts[part_lookup[index]];
        if token.before_exists != part.before.is_some()
            || token.after_exists != part.after.is_some()
        {
            return Err(invalid_scope());
        }
    }
    // XML and relationship identity are case insensitive even though the
    // deterministic stream retains the original spelling and exact sort order.
    drop(folded_indices(&lexical.parts, |value| value.name.as_str())?);
    let mut sink = HashSink::new(maximum);
    sink.append(b"WCS2")?;
    sink.integer(count)?;
    sink.append(&[0])?;
    sink.bytes(b"[Content_Types].xml")?;
    sink.boolean(true)?;
    sink.bytes(if before {
        lexical.content_types.before.bytes()
    } else {
        lexical.content_types.after.bytes()
    })?;
    for &index in &relationships {
        let change = &lexical.relationships[index];
        sink.append(&[1])?;
        sink.bytes(change.owner.as_str().as_bytes())?;
        let token = if before {
            change.before.as_ref()
        } else {
            change.after.as_ref()
        };
        sink.boolean(token.is_some())?;
        if let Some(token) = token {
            sink.bytes(token.bytes())?;
            sink.boolean(token.member_present())?;
            sink.bytes(token.bytes())?;
        }
    }
    for index in parts {
        let change = &patch.parts[index];
        sink.append(&[2])?;
        sink.bytes(change.name.as_str().as_bytes())?;
        let state = if before {
            change.before.as_ref()
        } else {
            change.after.as_ref()
        };
        sink.boolean(state.is_some())?;
        if let Some(state) = state {
            sink.bytes(state.content_type.as_bytes())?;
            sink.bytes(&state.data)?;
            let token = relationship_lookup
                .binary_search_by(|index| {
                    folded_cmp(
                        lexical.relationships[*index].owner.as_str(),
                        change.name.as_str(),
                    )
                })
                .ok()
                .map(|index| &lexical.relationships[relationship_lookup[index]])
                .and_then(|value| {
                    if before {
                        value.before.as_ref()
                    } else {
                        value.after.as_ref()
                    }
                });
            if let Some(token) = token {
                sink.boolean(token.member_present())?;
                sink.bytes(token.bytes())?;
            } else {
                sink.boolean(!state.relationships.is_empty())?;
                // Match the ordinary closure's canonical fallback without an XML buffer.
                let mut length = 0usize;
                canonical_relationships(state, |bytes| charge(&mut length, bytes.len(), maximum))?;
                sink.integer(length)?;
                canonical_relationships(state, |bytes| sink.append(bytes))?;
            }
        }
    }
    sink.integer(xml.len())?;
    for index in xml {
        let change = &lexical.parts[index];
        sink.bytes(change.name.as_str().as_bytes())?;
        sink.boolean(if before {
            change.before_exists
        } else {
            change.after_exists
        })?;
        let token = if before {
            change.before.as_ref()
        } else {
            change.after.as_ref()
        };
        sink.bytes(token.map_or(&[], |token| token.bytes()))?;
    }
    sink.finish()
}

fn sorted_indices<T>(values: &[T], key: impl Fn(&T) -> &str) -> Result<Vec<usize>> {
    let mut indices = Vec::new();
    indices
        .try_reserve_exact(values.len())
        .map_err(|_| PatchError::Allocation)?;
    indices.extend(0..values.len());
    indices.sort_unstable_by(|left, right| {
        key(&values[*left])
            .as_bytes()
            .cmp(key(&values[*right]).as_bytes())
    });
    Ok(indices)
}

fn folded_cmp(left: &str, right: &str) -> std::cmp::Ordering {
    left.bytes()
        .map(|byte| byte.to_ascii_lowercase())
        .cmp(right.bytes().map(|byte| byte.to_ascii_lowercase()))
}

fn folded_indices<T>(values: &[T], key: impl Fn(&T) -> &str) -> Result<Vec<usize>> {
    let mut indices = Vec::new();
    indices
        .try_reserve_exact(values.len())
        .map_err(|_| PatchError::Allocation)?;
    indices.extend(0..values.len());
    indices.sort_unstable_by(|left, right| folded_cmp(key(&values[*left]), key(&values[*right])));
    if indices
        .windows(2)
        .any(|pair| folded_cmp(key(&values[pair[0]]), key(&values[pair[1]])).is_eq())
    {
        return Err(invalid_scope());
    }
    Ok(indices)
}

fn invalid_scope() -> PatchError {
    PatchError::InvalidText {
        field: "Web Extensions source scope",
    }
}

fn charge(total: &mut usize, amount: usize, maximum: usize) -> Result<()> {
    let next = total.checked_add(amount).ok_or_else(invalid_scope)?;
    if next > maximum {
        return Err(PatchError::InvalidText {
            field: "Web Extensions durable scope byte limit",
        });
    }
    *total = next;
    Ok(())
}

struct HashSink {
    digest: Sha256,
    length: usize,
    maximum: usize,
}

impl HashSink {
    fn new(maximum: usize) -> Self {
        Self {
            digest: Sha256::new(),
            length: 0,
            maximum,
        }
    }

    fn append(&mut self, bytes: &[u8]) -> Result<()> {
        charge(&mut self.length, bytes.len(), self.maximum)?;
        self.digest.update(bytes);
        Ok(())
    }

    fn integer(&mut self, value: usize) -> Result<()> {
        self.append(
            &u64::try_from(value)
                .map_err(|_| invalid_scope())?
                .to_le_bytes(),
        )
    }

    fn boolean(&mut self, value: bool) -> Result<()> {
        self.append(&[u8::from(value)])
    }

    fn bytes(&mut self, bytes: &[u8]) -> Result<()> {
        self.integer(bytes.len())?;
        self.append(bytes)
    }

    fn finish(self) -> Result<String> {
        const HEX: &[u8; 16] = b"0123456789abcdef";
        let mut output = String::new();
        output
            .try_reserve_exact(64)
            .map_err(|_| PatchError::Allocation)?;
        for byte in self.digest.finalize() {
            output.push(char::from(HEX[usize::from(byte >> 4)]));
            output.push(char::from(HEX[usize::from(byte & 15)]));
        }
        Ok(output)
    }
}

fn canonical_relationships(
    state: &PartState,
    mut emit: impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    emit(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)?;
    for relation in &state.relationships {
        emit(b"<Relationship Id=\"")?;
        escaped_attribute(&relation.id, &mut emit)?;
        emit(b"\" Type=\"")?;
        escaped_attribute(&relation.relationship_type, &mut emit)?;
        emit(b"\" Target=\"")?;
        escaped_attribute(&relation.target, &mut emit)?;
        emit(b"\"")?;
        if relation.external {
            emit(b" TargetMode=\"External\"")?;
        }
        emit(b" />")?;
    }
    emit(b"</Relationships>")
}

fn escaped_attribute(value: &str, emit: &mut impl FnMut(&[u8]) -> Result<()>) -> Result<()> {
    let bytes = value.as_bytes();
    let mut start = 0;
    for (index, byte) in bytes.iter().enumerate() {
        let replacement: &[u8] = match byte {
            b'&' => b"&amp;",
            b'\"' => b"&quot;",
            b'<' => b"&lt;",
            b'>' => b"&gt;",
            _ => continue,
        };
        emit(&bytes[start..index])?;
        emit(replacement)?;
        start = index + 1;
    }
    emit(&bytes[start..])
}

#[cfg(test)]
#[allow(clippy::unwrap_used, reason = "bounded hashing fixtures must succeed")]
mod tests {
    use super::*;

    fn creation_patch() -> Patch {
        use crate::web::{AddIn, Conformance, Pane, Panes, Reference, Store};
        let mut panes = Panes::new();
        panes
            .push(Pane::new(
                AddIn::new(
                    "scope-instance",
                    Reference::new("scope-addin", "1.0", Store::Omex).unwrap(),
                )
                .unwrap(),
            ))
            .unwrap();
        crate::web::plan_put(
            &litchi_opc::OpcPackage::new(),
            panes,
            Conformance::Transitional,
        )
        .unwrap()
    }

    // Retain the prior allocated WCS2 encoder as a test oracle for the
    // streaming replacement. This path is never used in production.
    fn reference_hash(patch: &Patch, before: bool) -> String {
        use super::super::{Encoder, collect_patch_records, encode_member};
        let records = collect_patch_records(patch, true, usize::MAX).unwrap();
        let lexical = patch.lexical.as_ref().unwrap();
        let mut encoder = Encoder::new(b"WCS2", usize::MAX).unwrap();
        encoder.u64(records.len()).unwrap();
        for record in records {
            encoder.u8(record.kind).unwrap();
            encoder.text(&record.name).unwrap();
            encode_member(
                &mut encoder,
                record.kind,
                if before {
                    record.before.as_ref()
                } else {
                    record.after.as_ref()
                },
            )
            .unwrap();
        }
        let mut xml: Vec<_> = lexical.parts.iter().collect();
        xml.sort_unstable_by(|left, right| {
            left.name
                .as_str()
                .as_bytes()
                .cmp(right.name.as_str().as_bytes())
        });
        encoder.u64(xml.len()).unwrap();
        for change in xml {
            encoder.text(change.name.as_str()).unwrap();
            encoder
                .bool(if before {
                    change.before_exists
                } else {
                    change.after_exists
                })
                .unwrap();
            let token = if before {
                change.before.as_ref()
            } else {
                change.after.as_ref()
            };
            encoder
                .bytes(token.map_or(&[], |token| token.bytes()))
                .unwrap();
        }
        let mut sink = HashSink::new(usize::MAX);
        sink.append(&encoder.finish()).unwrap();
        sink.finish().unwrap()
    }

    #[test]
    fn streamed_scope_matches_retained_wcs2_reference_for_both_directions() {
        let patch = creation_patch();
        for patch in [&patch, &patch.inverse()] {
            for before in [true, false] {
                assert_eq!(
                    scope_hash_from_patch_scope(patch, before, usize::MAX).unwrap(),
                    reference_hash(patch, before)
                );
            }
        }
    }

    #[test]
    fn incomplete_and_duplicate_scopes_are_refused() {
        let mut patch = creation_patch();
        patch.parts = Box::new([]);
        assert!(scope_hash_from_patch_scope(&patch, true, usize::MAX).is_err());
        assert!(folded_indices(&["/Mixed.xml", "/mixed.xml"], |name| name).is_err());
        assert!(folded_indices(&["/same.xml", "/same.xml"], |name| name).is_err());
        let names = ["/Z.xml", "/a.xml"];
        let indices = folded_indices(&names, |name| name).unwrap();
        let index = indices
            .binary_search_by(|index| folded_cmp(names[*index], "/A.XML"))
            .unwrap();
        assert_eq!(indices[index], 1);
    }

    #[test]
    fn streaming_hash_has_exact_byte_boundary() {
        let mut sink = HashSink::new(3);
        sink.append(b"a").unwrap();
        sink.append(b"bc").unwrap();
        assert!(sink.append(b"d").is_err());
        assert_eq!(
            sink.finish().unwrap(),
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
        );
        assert!(HashSink::new(0).append(b"WCS2").is_err());
    }

    #[test]
    fn streamed_canonical_relationships_match_closure_fallback() {
        let state = PartState {
            content_type: String::new(),
            data: std::sync::Arc::new(Vec::new()),
            relationships: vec![super::super::super::RelationshipState {
                id: "id&\"".into(),
                relationship_type: "urn:<test>".into(),
                target: "https://example.test/é?a=1&b=2".into(),
                external: true,
            }]
            .into_boxed_slice(),
        };
        let mut output = Vec::new();
        canonical_relationships(&state, |bytes| {
            output.extend_from_slice(bytes);
            Ok(())
        })
        .unwrap();
        assert_eq!(output, super::super::canonical_relationships(&state));
    }
}

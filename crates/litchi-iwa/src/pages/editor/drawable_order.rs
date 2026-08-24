//! Native back-to-front ordering for Pages document drawables.

use litchi_iwa_common::comment::DrawableId;
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::pages_drawable_order_codec::{
    self as drawable_order_codec, DecodeError, DecodeOptions, DrawableOrderWrite, WireResourceLimit,
};

use super::*;
use crate::drawable_order::{DrawableLayerMove, move_drawable_layer, validate_unique_drawables};
use crate::package::PackageLimits;

const PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE: u32 = 10_015;

impl PagesEditor {
    /// List document drawable identifiers from the back-most to the front-most layer.
    pub fn body_drawable_order(&self) -> Result<Vec<DrawableId>> {
        pages_drawable_order(self.package())
    }

    /// Set the exact back-to-front order of every identifier in the native
    /// drawable-order list.
    ///
    /// The supplied slice must be a permutation of [`Self::body_drawable_order`].
    /// Reference payloads and unrelated native wire fields are retained verbatim.
    pub fn set_body_drawable_order(&mut self, ordered_drawable_ids: &[DrawableId]) -> Result<()> {
        let current = self.body_drawable_order()?;
        if current == ordered_drawable_ids {
            return Ok(());
        }
        let mut staged = self.package().clone();
        rewrite_pages_drawable_order(&mut staged, ordered_drawable_ids)?;
        let verified = Self::from_package(staged)?;
        if verified.body_drawable_order()? != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Pages drawable-order update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Move one drawable using the native Arrange layer semantics.
    ///
    /// Returns `false` when the requested move is already at the requested
    /// boundary; otherwise the changed order is committed transactionally.
    pub fn move_body_drawable(
        &mut self,
        drawable_object_id: DrawableId,
        movement: DrawableLayerMove,
    ) -> Result<bool> {
        let current = self.body_drawable_order()?;
        let Some(ordered) = move_drawable_layer(&current, drawable_object_id, movement)? else {
            return Ok(false);
        };
        self.set_body_drawable_order(&ordered)?;
        Ok(true)
    }
}

fn pages_drawable_order(package: &IWorkPackage) -> Result<Vec<DrawableId>> {
    let document = root_document(package)?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let archive_name = find_object_archive(package, z_order_id.identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(z_order_id.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages drawable z-order object {} is missing",
            z_order_id.identifier
        ))
    })?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE);
    let message = messages.next().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages drawable z-order object {} must have exactly one payload",
            z_order_id.identifier
        ))
    })?;
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Pages drawable z-order object {} must have exactly one payload",
            z_order_id.identifier
        )));
    }
    let options = drawable_order_decode_options(package.limits(), &message.data, false);
    let (snapshot, _report) =
        drawable_order_codec::decode_drawable_order_with_report(&message.data, options)
            .map_err(map_drawable_order_error)?;
    let ordered = drawable_ids_from_snapshot(snapshot)?;
    for identifier in &ordered {
        find_object_archive(package, identifier.get())?;
    }
    validate_unique_drawables(&ordered, "Pages drawable order")?;
    Ok(ordered)
}

fn rewrite_pages_drawable_order(
    package: &mut IWorkPackage,
    ordered_drawable_ids: &[DrawableId],
) -> Result<()> {
    let current = pages_drawable_order(package)?;
    let document = root_document(package)?;
    let z_order_id = document.drawables_zorder.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no drawable z-order object".to_owned())
    })?;
    let archive_name = find_object_archive(package, z_order_id.identifier)?;
    let limits = package.limits();
    let requested_raw = raw_drawable_ids(ordered_drawable_ids)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(z_order_id.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable z-order object {} is missing",
                z_order_id.identifier
            ))
        })?;
        let mut indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE)
            .map(|(index, _)| index);
        let message_index = indexes.next().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages drawable z-order object {} must have exactly one payload",
                z_order_id.identifier
            ))
        })?;
        if indexes.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages drawable z-order object {} must have exactly one payload",
                z_order_id.identifier
            )));
        }
        let original = object.messages[message_index].data.as_slice();
        let options = drawable_order_decode_options(limits, original, true);
        let (previous, _report) =
            drawable_order_codec::decode_drawable_order_with_report(original, options)
                .map_err(map_drawable_order_error)?;
        let mut previous_ids = Vec::new();
        previous_ids
            .try_reserve_exact(previous.len())
            .map_err(|_| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "Pages drawable-order identifier staging",
                    amount: previous.len(),
                })
            })?;
        previous_ids.extend(previous.identifiers());
        let expected_ids = current.iter().map(|identifier| identifier.get());
        if previous_ids.iter().copied().ne(expected_ids) {
            return Err(Error::InvalidFormat(
                "Pages drawable-order source changed during mutation".to_owned(),
            ));
        }
        let (data, _report) = drawable_order_codec::rewrite_drawable_order_with_report(
            original,
            DrawableOrderWrite::new(&requested_raw),
            options,
        )
        .map_err(map_drawable_order_error)?;
        let (verified, _report) = drawable_order_codec::decode_drawable_order_with_report(
            data.as_slice(),
            drawable_order_decode_options(limits, &data, false),
        )
        .map_err(map_drawable_order_error)?;
        let verified = drawable_ids_from_snapshot(verified)?;
        if verified != ordered_drawable_ids {
            return Err(Error::InvalidFormat(
                "Pages drawable-order wire patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn drawable_ids_from_snapshot(
    snapshot: litchi_iwa_protos::pages_drawable_order_codec::DrawableOrderSnapshot<'_>,
) -> Result<Vec<DrawableId>> {
    let mut ordered = Vec::new();
    ordered.try_reserve_exact(snapshot.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages drawable-order identifiers",
            amount: snapshot.len(),
        })
    })?;
    for identifier in snapshot.identifiers() {
        ordered.push(DrawableId::from_raw(identifier).map_err(crate::Error::from)?);
    }
    Ok(ordered)
}

fn raw_drawable_ids(identifiers: &[DrawableId]) -> Result<Vec<u64>> {
    let mut raw = Vec::new();
    raw.try_reserve_exact(identifiers.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages drawable-order rewrite identifiers",
            amount: identifiers.len(),
        })
    })?;
    raw.extend(identifiers.iter().map(|identifier| identifier.get()));
    Ok(raw)
}

const DRAWABLE_ORDER_RECURSION_LIMIT: u32 = 8;
const DRAWABLE_ORDER_FIELD_MULTIPLIER: usize = 8;
const DRAWABLE_ORDER_WORK_MULTIPLIER: usize = 64;

fn drawable_order_decode_options(
    limits: PackageLimits,
    source: &[u8],
    rewrite: bool,
) -> DecodeOptions {
    let stream_limit = limits
        .max_iwa_stream_bytes()
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let archive_limit = limits
        .archive_limits()
        .max_archive_bytes()
        .min(stream_limit)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let max_input_bytes = source.len().max(1).min(archive_limit);
    let max_output_bytes = if rewrite {
        archive_limit.min(WireLimits::MAX_OUTPUT_BYTES)
    } else {
        source.len().max(1).min(archive_limit)
    };
    let max_fields = source
        .len()
        .saturating_mul(DRAWABLE_ORDER_FIELD_MULTIPLIER)
        .clamp(1, WireLimits::MAX_FIELDS);
    let max_work_bytes = source
        .len()
        .saturating_mul(DRAWABLE_ORDER_WORK_MULTIPLIER)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let max_references = source.len().clamp(1, WireLimits::MAX_FIELDS);
    DecodeOptions::new(
        max_input_bytes,
        max_output_bytes,
        max_fields,
        max_work_bytes,
        DRAWABLE_ORDER_RECURSION_LIMIT,
        max_references,
    )
}

fn map_drawable_order_error(error: DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages drawable-order output",
            amount,
        });
    }
    match error.resource_limit() {
        Some(WireResourceLimit::InputBytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit: maximum,
            })
        },
        Some(WireResourceLimit::OutputBytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                observed,
                limit: maximum,
            })
        },
        Some(WireResourceLimit::Fields { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: maximum,
            })
        },
        Some(WireResourceLimit::WorkBytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed,
                limit: maximum,
            })
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: usize::try_from(observed).unwrap_or(usize::MAX),
                limit: usize::try_from(maximum).unwrap_or(usize::MAX),
            })
        },
        Some(WireResourceLimit::References { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: maximum,
            })
        },
        Some(_) => Error::InvalidFormat(format!(
            "Pages drawable-order payload failed strict resource validation: {error}"
        )),
        None => Error::InvalidFormat(format!(
            "Pages drawable-order payload failed strict validation: {error}"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::DrawableLayerMove;
    use crate::shapes::{DrawablePoint, DrawableSize};

    const OVERLAP_POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
    const OVERLAP_SIZE: DrawableSize = DrawableSize {
        width: 240.0,
        height: 180.0,
    };

    #[test]
    fn scratch_document_supports_drawable_order_crud() {
        let mut editor = PagesEditor::create_with_text("Layered").unwrap();
        let first = append_rectangle(&mut editor, "First");
        let second = append_rectangle(&mut editor, "Second");
        let third = append_rectangle(&mut editor, "Third");
        let original = editor.body_drawable_order().unwrap();
        assert!(original.contains(&first));
        assert!(original.contains(&second));
        assert!(original.contains(&third));

        assert!(
            editor
                .move_body_drawable(first, DrawableLayerMove::ToFront)
                .unwrap()
        );
        assert_eq!(editor.body_drawable_order().unwrap().last(), Some(&first));
        assert!(
            editor
                .move_body_drawable(first, DrawableLayerMove::ToBack)
                .unwrap()
        );
        assert_eq!(editor.body_drawable_order().unwrap().first(), Some(&first));

        let mut reversed = original.clone();
        reversed.reverse();
        editor.set_body_drawable_order(&reversed).unwrap();
        assert_eq!(editor.body_drawable_order().unwrap(), reversed);
        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_drawable_order().unwrap(), reversed);

        let bytes = editor.to_bytes().unwrap();
        let mut duplicate = reversed.clone();
        duplicate[1] = duplicate[0];
        assert!(editor.set_body_drawable_order(&duplicate).is_err());
        assert_eq!(editor.to_bytes().unwrap(), bytes);

        editor.set_body_drawable_order(&original).unwrap();
        assert!(
            !editor
                .move_body_drawable(original[0], DrawableLayerMove::ToBack)
                .unwrap()
        );
    }

    #[test]
    fn wave70_reorder_keeps_each_reference_payload_with_its_identifier_and_is_reversible() {
        let mut editor = layered_editor();
        let original = editor.body_drawable_order().unwrap();
        assert!(original.len() >= 3);

        let enriched = z_order_payload(&editor);
        let references = crate::wire::repeated_length_delimited_payloads(&enriched, 1).unwrap();
        assert_eq!(references.len(), original.len());
        let mut expected = Vec::with_capacity(references.len());
        let mut source = Vec::new();
        for (index, reference) in references.iter().enumerate() {
            let mut enriched_reference = reference.to_vec();
            crate::wire::append_varint_field(
                &mut enriched_reference,
                2,
                u64::from(100 + index as u32),
            )
            .unwrap();
            crate::wire::append_varint_field(&mut enriched_reference, 3, (index % 2) as u64)
                .unwrap();
            crate::wire::append_varint_field(
                &mut enriched_reference,
                90,
                0xfeed_0000_u64 + index as u64,
            )
            .unwrap();
            crate::wire::append_length_delimited_field(&mut source, 1, &enriched_reference)
                .unwrap();
            expected.push((reference_identifier(reference), enriched_reference));
        }
        crate::wire::append_varint_field(&mut source, 99, 0xfeed_face).unwrap();
        install_z_order_payload(&mut editor, source);
        let before = editor.to_bytes().unwrap();

        let mut requested = original.clone();
        requested.swap(0, 2);
        editor.set_body_drawable_order(&requested).unwrap();
        let candidate = z_order_payload(&editor);
        let candidate_references =
            crate::wire::repeated_length_delimited_payloads(&candidate, 1).unwrap();
        assert_eq!(
            candidate_references
                .iter()
                .map(|reference| reference_identifier(reference))
                .collect::<Vec<_>>(),
            requested
                .iter()
                .map(|identifier| identifier.get())
                .collect::<Vec<_>>()
        );
        for reference in candidate_references {
            let identifier = reference_identifier(reference);
            let (_, expected_reference) = expected
                .iter()
                .find(|(expected_identifier, _)| *expected_identifier == identifier)
                .expect("candidate identifier was in source order");
            assert_eq!(reference, expected_reference.as_slice());
        }
        assert!(
            candidate
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0xce])
        );

        editor.set_body_drawable_order(&original).unwrap();
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn wave70_malformed_nested_reference_fields_fail_without_publication() {
        let malformed_references = [
            vec![0x0a, 0x01, 0x01],       // required identifier has the wrong wire type
            vec![0x08, 0x01, 0x08, 0x02], // duplicate required identifier
            vec![0x08, 0x81, 0x00],       // non-canonical identifier varint
            vec![0x08, 0x01, 0x10, 0x02, 0x10, 0x03], // duplicate deprecated type
        ];
        for malformed in malformed_references {
            let package = malformed_z_order_package(&malformed);
            let before = package.to_bytes().unwrap();
            let parsed = PagesEditor::from_package(package);
            let Ok(mut editor) = parsed else {
                continue;
            };
            assert!(editor.body_drawable_order().is_err());
            assert_eq!(editor.to_bytes().unwrap(), before);
            let current = [DrawableId::from_raw(1).unwrap()];
            assert!(editor.set_body_drawable_order(&current).is_err());
            assert_eq!(editor.to_bytes().unwrap(), before);
        }
    }

    #[test]
    fn wave70_duplicate_source_identifiers_fail_without_publication() {
        let editor = layered_editor();
        let original = editor.body_drawable_order().unwrap();
        let payload = z_order_payload(&editor);
        let references = crate::wire::repeated_length_delimited_payloads(&payload, 1).unwrap();
        let mut duplicate_payload = Vec::new();
        for reference in references.iter().take(references.len().saturating_sub(1)) {
            crate::wire::append_length_delimited_field(&mut duplicate_payload, 1, reference)
                .unwrap();
        }
        crate::wire::append_length_delimited_field(&mut duplicate_payload, 1, references[0])
            .unwrap();
        let package = package_with_z_order_payload(&editor, duplicate_payload);
        let before = package.to_bytes().unwrap();
        let parsed = PagesEditor::from_package(package);
        let Ok(mut editor) = parsed else {
            return;
        };
        assert!(editor.body_drawable_order().is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
        let requested = original.iter().rev().copied().collect::<Vec<_>>();
        assert!(editor.set_body_drawable_order(&requested).is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn wave70_non_permutations_fail_atomically() {
        let mut editor = layered_editor();
        let original = editor.body_drawable_order().unwrap();
        let before = editor.to_bytes().unwrap();
        let foreign = DrawableId::from_raw(0x1_0000).unwrap();
        let invalid = [
            original[..original.len() - 1].to_vec(),
            vec![original[0], original[0], original[1]],
            vec![original[0], original[1], foreign],
        ];
        for requested in invalid {
            assert!(editor.set_body_drawable_order(&requested).is_err());
            assert_eq!(editor.to_bytes().unwrap(), before);
        }
    }

    fn append_rectangle(editor: &mut PagesEditor, text: &str) -> DrawableId {
        let anchor = editor.body_text().unwrap().encode_utf16().count();
        DrawableId::from_raw(
            editor
                .add_body_rectangle(anchor, text, OVERLAP_POSITION, OVERLAP_SIZE)
                .unwrap()
                .drawable_object_id,
        )
        .unwrap()
    }

    fn layered_editor() -> PagesEditor {
        let mut editor = PagesEditor::create_with_text("Layered").unwrap();
        append_rectangle(&mut editor, "First");
        append_rectangle(&mut editor, "Second");
        append_rectangle(&mut editor, "Third");
        editor
    }

    fn z_order_payload(editor: &PagesEditor) -> Vec<u8> {
        let document = root_document(editor.package()).unwrap();
        let z_order_id = document.drawables_zorder.unwrap().identifier;
        let archive_name = find_object_archive(editor.package(), z_order_id).unwrap();
        let archive = editor.package().archive(&archive_name).unwrap();
        archive
            .object(z_order_id)
            .unwrap()
            .messages
            .iter()
            .find(|message| message.type_ == PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE)
            .unwrap()
            .data
            .clone()
    }

    fn install_z_order_payload(editor: &mut PagesEditor, payload: Vec<u8>) {
        let package = package_with_z_order_payload(editor, payload);
        *editor = PagesEditor::from_package(package).unwrap();
    }

    fn package_with_z_order_payload(editor: &PagesEditor, payload: Vec<u8>) -> IWorkPackage {
        let document = root_document(editor.package()).unwrap();
        let z_order_id = document.drawables_zorder.unwrap().identifier;
        let archive_name = find_object_archive(editor.package(), z_order_id).unwrap();
        let mut package = editor.package().clone();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(z_order_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE)
                    .unwrap();
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: PAGES_DRAWABLE_ORDER_PAYLOAD_TYPE,
                        data: payload,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        package
    }

    fn malformed_z_order_package(malformed_reference: &[u8]) -> IWorkPackage {
        let editor = layered_editor();
        let source = z_order_payload(&editor);
        let references = crate::wire::repeated_length_delimited_payloads(&source, 1).unwrap();
        let mut payload = Vec::new();
        crate::wire::append_length_delimited_field(&mut payload, 1, malformed_reference).unwrap();
        for reference in references.iter().skip(1) {
            crate::wire::append_length_delimited_field(&mut payload, 1, reference).unwrap();
        }
        package_with_z_order_payload(&editor, payload)
    }

    fn reference_identifier(reference: &[u8]) -> u64 {
        let (key, key_length) =
            litchi_iwa_common::varint::decode_varint_from_bytes(reference).unwrap();
        assert_eq!(key, 0x08, "reference identifier field");
        litchi_iwa_common::varint::decode_varint_from_bytes(&reference[key_length..])
            .unwrap()
            .0
    }
}

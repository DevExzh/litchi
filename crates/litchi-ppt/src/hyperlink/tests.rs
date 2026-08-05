use super::*;
use crate::consts::RecordType;
use crate::records::Record;

#[cfg(test)]
mod interaction_protocol_tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: RecordType, data: &[u8]) -> Vec<u8> {
        encode_record(version, instance, kind.as_u16(), data).unwrap()
    }

    fn interaction(instance: u16, atom: &[u8], macro_data: Option<&[u8]>) -> Vec<u8> {
        let mut children = record(0, 0, RecordType::InteractiveInfoAtom, atom);
        if let Some(data) = macro_data {
            children.extend(record(0, 2, RecordType::CString, data));
        }
        record(0x0F, instance, RecordType::InteractiveInfo, &children)
    }

    fn atom() -> [u8; 16] {
        InteractiveInfoAtom {
            sound_id: 7,
            hyperlink_id: 11,
            action: InteractionAction::Macro,
            ole_verb: 19,
            jump: InteractionJump::LastSlideViewed,
            animated: true,
            stop_sound: true,
            custom_show_return: true,
            visited: true,
            link_target: InteractionLinkTarget::OtherFile,
            unused: [0xAA, 0xBB, 0xCC],
        }
        .to_payload()
    }

    #[test]
    fn exact_round_trip_preserves_trigger_macro_bytes_and_undefined_data() {
        let macro_data = [b'R', 0, b'u', 0, b'n', 0, 0, 0, 1, 0];
        let bytes = interaction(1, &atom(), Some(&macro_data));
        let parsed = Interaction::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.trigger, InteractionTrigger::MouseOver);
        assert_eq!(parsed.macro_name.as_deref(), Some("Run"));
        assert_eq!(
            parsed.macro_name_atom().unwrap().unwrap().raw_utf16(),
            macro_data
        );
        assert_eq!(parsed.unused, [0xAA, 0xBB, 0xCC]);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn canonical_constructor_and_record_accessor_round_trip() {
        let mut value = Interaction::new(
            InteractionTrigger::Click,
            InteractionAction::CustomShow,
            InteractionLinkTarget::CustomShow,
        )
        .with_macro_name("Quarterly show")
        .unwrap();
        value.hyperlink_id = 42;
        value.custom_show_return = true;
        let bytes = value.to_bytes().unwrap();
        assert_eq!(
            Interaction::parse(&value.to_record().unwrap()).unwrap(),
            value
        );
        assert_eq!(Interaction::parse_bytes(&bytes).unwrap(), value);
    }

    #[test]
    fn preserves_ignored_macro_name_without_activating_it() {
        let bytes = interaction(
            0,
            &InteractiveInfoAtom {
                action: InteractionAction::Hyperlink,
                ..Interaction::new(
                    InteractionTrigger::Click,
                    InteractionAction::Hyperlink,
                    InteractionLinkTarget::Url,
                )
                .atom()
            }
            .to_payload(),
            Some(&[b'X', 0]),
        );
        let parsed = Interaction::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.macro_name.as_deref(), Some("X"));
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_bad_container_atom_and_child_order() {
        let valid_atom = atom();
        for instance in [2u16, 0x0FFF] {
            assert!(Interaction::parse_bytes(&interaction(instance, &valid_atom, None)).is_err());
        }
        assert!(Interaction::parse_bytes(&record(0, 0, RecordType::InteractiveInfo, &[])).is_err());
        assert!(Interaction::parse_bytes(&interaction(0, &valid_atom[..15], None)).is_err());

        let name = record(0, 2, RecordType::CString, &[b'A', 0]);
        let atom_record = record(0, 0, RecordType::InteractiveInfoAtom, &valid_atom);
        assert!(
            Interaction::parse_bytes(&record(
                0x0F,
                0,
                RecordType::InteractiveInfo,
                &[name, atom_record].concat()
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_enum_reserved_and_printable_string_violations() {
        for (offset, value) in [(8usize, 8u8), (10, 7), (12, 4)] {
            let mut bad = atom();
            bad[offset] = value;
            assert!(Interaction::parse_bytes(&interaction(0, &bad, None)).is_err());
        }
        let mut reserved = atom();
        reserved[11] |= 0x10;
        assert!(Interaction::parse_bytes(&interaction(0, &reserved, None)).is_err());
        assert!(Interaction::parse_bytes(&interaction(0, &atom(), Some(&[1]))).is_err());
        assert!(Interaction::parse_bytes(&interaction(0, &atom(), Some(&[1, 0]))).is_err());
        assert!(Interaction::parse_bytes(&interaction(0, &atom(), Some(&[0, 0xD8]))).is_err());
    }

    #[test]
    fn enforces_record_and_macro_limits_and_exact_consumption() {
        let bytes = interaction(0, &atom(), Some(&[b'A', 0, b'B', 0]));
        assert!(
            Interaction::parse_bytes_with_limits(
                &bytes,
                InteractionLimits {
                    max_record_bytes: bytes.len() - 1,
                    ..Default::default()
                }
            )
            .is_err()
        );
        assert!(
            Interaction::parse_bytes_with_limits(
                &bytes,
                InteractionLimits {
                    max_macro_name_bytes: 2,
                    ..Default::default()
                }
            )
            .is_err()
        );
        let mut trailing = bytes;
        trailing.push(0);
        assert!(Interaction::parse_bytes(&trailing).is_err());

        let value = Interaction::new(
            InteractionTrigger::Click,
            InteractionAction::NoAction,
            InteractionLinkTarget::Nil,
        );
        let limits = InteractionLimits {
            max_record_bytes: 31,
            ..Default::default()
        };
        assert!(value.validate_with_limits(limits).is_err());
        assert!(value.to_bytes_with_limits(limits).is_err());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn unicode(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn hyperlink(id: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode("Example")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("https://example.test")));
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("section")));
        record_bytes(0x0f, 0, 4055, &payload)
    }

    fn external_object_list(seed: i32, hyperlinks: &[Vec<u8>]) -> Record {
        let mut payload = record_bytes(0, 0, 1034, &seed.to_le_bytes());
        for hyperlink in hyperlinks {
            payload.extend_from_slice(hyperlink);
        }
        Record {
            record_type: RecordType::ExObjList,
            record_type_raw: 1033,
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
    }

    fn hyperlink9(id: u32, screen_tip: Option<&str>, flags: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        if let Some(screen_tip) = screen_tip {
            payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode(screen_tip)));
        }
        payload.extend_from_slice(&record_bytes(0, 0, 4120, &flags.to_le_bytes()));
        record_bytes(0x0f, 0, 4068, &payload)
    }

    fn prog_tags_record(blob_payload: &[u8]) -> Record {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        Record {
            record_type: RecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    fn interaction_record(trigger: u16, flags: u8, target: u8) -> Record {
        let mut atom = [0u8; 16];
        atom[4..8].copy_from_slice(&3u32.to_le_bytes());
        atom[8] = 4;
        atom[10] = 0;
        atom[11] = flags;
        atom[12] = target;
        let payload = record_bytes(0, 0, 4083, &atom);
        let bytes = record_bytes(0x0f, trigger, 4082, &payload);
        Record::parse(&bytes, 0).unwrap().0
    }

    fn root(list: Option<Record>, extensions: &[Vec<u8>]) -> Record {
        let mut children = Vec::new();
        if let Some(list) = list {
            children.push(list);
        }
        if !extensions.is_empty() {
            let blob: Vec<u8> = extensions.iter().flatten().copied().collect();
            children.push(prog_tags_record(&blob));
        }
        Record {
            record_type: RecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_and_merges_powerpoint9_hyperlinks() {
        let root = root(
            Some(external_object_list(7, &[hyperlink(3)])),
            &[hyperlink9(3, Some("Open example"), 7)],
        );
        let hyperlinks = Hyperlinks::parse(&root).unwrap();
        assert_eq!(hyperlinks.id_seed, Some(7));
        let hyperlink = hyperlinks.get(3).unwrap();
        assert_eq!(hyperlink.friendly_name.as_deref(), Some("Example"));
        assert_eq!(hyperlink.target.as_deref(), Some("https://example.test"));
        assert_eq!(hyperlink.location.as_deref(), Some("section"));
        let extension = hyperlink.extension.as_ref().unwrap();
        assert_eq!(extension.screen_tip.as_deref(), Some("Open example"));
        assert!(extension.inserted_with_dialog);
        assert!(extension.location_is_named_show);
        assert!(extension.named_show_returns_to_slide);
    }

    #[test]
    fn parses_and_resolves_interactive_hyperlinks() {
        let interaction = Interaction::parse(&interaction_record(0, 0x09, 8)).unwrap();
        assert_eq!(interaction.trigger, InteractionTrigger::Click);
        assert_eq!(interaction.action, InteractionAction::Hyperlink);
        assert_eq!(interaction.link_target, InteractionLinkTarget::Url);
        assert!(interaction.animated);
        assert!(interaction.visited);

        let hyperlinks =
            Hyperlinks::parse(&root(Some(external_object_list(3, &[hyperlink(3)])), &[])).unwrap();
        assert_eq!(
            interaction
                .hyperlink(&hyperlinks)
                .unwrap()
                .target
                .as_deref(),
            Some("https://example.test")
        );

        assert!(Interaction::parse(&interaction_record(2, 0, 8)).is_err());
        assert!(Interaction::parse(&interaction_record(0, 0x10, 8)).is_err());
        assert!(Interaction::parse(&interaction_record(0, 0, 4)).is_err());
    }

    #[test]
    fn accepts_optional_base_strings_and_absent_extensions() {
        let atom_only = record_bytes(
            0x0f,
            0,
            4055,
            &record_bytes(0, 0, 4051, &1u32.to_le_bytes()),
        );
        let hyperlinks =
            Hyperlinks::parse(&root(Some(external_object_list(1, &[atom_only])), &[])).unwrap();
        assert_eq!(hyperlinks.get(1).unwrap().target, None);
    }

    #[test]
    fn rejects_invalid_hyperlink_ids_and_extensions() {
        assert!(
            Hyperlinks::parse(&root(Some(external_object_list(2, &[hyperlink(3)])), &[],)).is_err()
        );
        assert!(
            Hyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3), hyperlink(3)])),
                &[],
            ))
            .is_err()
        );
        assert!(
            Hyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(4, None, 0)],
            ))
            .is_err()
        );
        assert!(
            Hyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 8)],
            ))
            .is_err()
        );
        assert!(
            Hyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 0), hyperlink9(3, None, 0)],
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_malformed_hyperlink_strings_and_child_order() {
        let mut invalid_utf16 = hyperlink(1);
        invalid_utf16[28..30].copy_from_slice(&0xd800u16.to_le_bytes());
        assert!(
            Hyperlinks::parse(&root(Some(external_object_list(1, &[invalid_utf16])), &[],))
                .is_err()
        );

        let mut payload = record_bytes(0, 0, 4051, &1u32.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("late")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("early")));
        let out_of_order = record_bytes(0x0f, 0, 4055, &payload);
        assert!(
            Hyperlinks::parse(&root(Some(external_object_list(1, &[out_of_order])), &[],)).is_err()
        );
    }
}

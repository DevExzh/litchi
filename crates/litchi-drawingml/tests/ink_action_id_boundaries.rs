//! ID spelling must not change the semantics of a source-backed removal.
use litchi_drawingml::ink::{
    ACTION_NAMESPACE,
    actions::{
        self, ActionDataDraft, ActionSelector, ActionType, Edit, Limits, OpaquePayload, RootChild,
    },
};

#[test]
fn short_ids_and_non_reference_text_do_not_block_removal() {
    for (removed, retained, comment) in [
        ("a", "b", ""),
        (
            "stroke",
            "stroke2",
            "<!-- stroke is an ordinary comment -->",
        ),
    ] {
        let xml = format!(
            r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms">{comment}<a:action xml:id="{removed}" type="add" startTime="0"/><a:action xml:id="{retained}" type="add" startTime="1"/></a:actions>"#
        );
        let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
        edit.remove_action(ActionSelector::ordinal(0)).unwrap();
        let commit = edit
            .finish()
            .expect("unreferenced action can be removed regardless of ID spelling");
        let ids: Vec<_> = commit
            .profile()
            .children()
            .iter()
            .map(|child| match child {
                RootChild::Action(action) => action.xml_id().unwrap(),
                _ => panic!("unexpected group"),
            })
            .collect();
        assert_eq!(ids, [retained]);
        assert_eq!(
            commit.patch().inverse().apply(commit.as_bytes()).unwrap(),
            xml.as_bytes()
        );
        assert!(
            std::str::from_utf8(commit.as_bytes())
                .unwrap()
                .contains(comment)
        );
    }
}

#[test]
fn encoded_local_fragment_cannot_bypass_reference_closure() {
    for reference in ["#a", "  #%61  ", "&#x23;%61"] {
        let xml = format!(
            r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action xml:id="a" type="add" startTime="0"/><a:action xml:id="b" type="add" startTime="1"><a:actionData ref="{reference}"/></a:action></a:actions>"#
        );
        let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
        edit.remove_action(ActionSelector::ordinal(0)).unwrap();
        assert!(
            edit.finish().is_err(),
            "local reference {reference:?} must retain the same target after XML/URI decoding"
        );
    }
}

#[test]
fn removing_action_cannot_orphan_id_owned_by_opaque_trace() {
    let xml = format!(
        r##"<a:actions xmlns:a="{ACTION_NAMESPACE}" xmlns:i="http://www.w3.org/2003/InkML" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"><a:actionData><i:trace xml:id="stroke">0 0, 1 1</i:trace></a:actionData></a:action><a:action type="add" startTime="1"><a:actionData ref="#stroke"/></a:action></a:actions>"##
    );
    let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
    edit.remove_action(ActionSelector::ordinal(0)).unwrap();
    assert!(
        edit.finish().is_err(),
        "removal must account for IDs declared inside a removed opaque payload"
    );
}

#[test]
fn trace_draft_cannot_silently_disappear_as_comment_only_payload() {
    for payload in ["<!-- preserved comment -->", " \n\t "] {
        let xml = format!(
            r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
        );
        let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
        let result = (|| {
            let data = ActionDataDraft::new().trace(OpaquePayload::new(payload)?)?;
            edit.add_data(ActionSelector::ordinal(0), data)?;
            edit.finish()
        })();
        assert!(
            result.is_err(),
            "a trace draft must contain a trace element, not only {payload:?}"
        );
    }
}

#[test]
fn empty_custom_type_remains_an_exact_source_noop() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type='' startTime="0"/></a:actions>"#
    );
    let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
    edit.set_action_type(ActionSelector::ordinal(0), ActionType::custom("").unwrap())
        .unwrap();
    assert_eq!(edit.finish().unwrap().as_bytes(), xml.as_bytes());
}

#[test]
fn trace_draft_cannot_be_read_back_as_a_transform() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
    );
    let payload = format!(r#"<a:transform xmlns:a="{ACTION_NAMESPACE}"/>"#);
    let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
    let result = (|| {
        let data = ActionDataDraft::new().trace(OpaquePayload::new(payload)?)?;
        edit.add_data(ActionSelector::ordinal(0), data)?;
        edit.finish()
    })();
    assert!(
        result.is_err(),
        "typed trace intent cannot silently become a transform"
    );
}

#[test]
fn opaque_descendants_count_toward_the_authored_depth_limit() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
    );
    let profile = actions::read_profile(xml.as_bytes()).unwrap();
    let data = ActionDataDraft::new().trace(OpaquePayload::new(
        r#"<i:trace xmlns:i="http://www.w3.org/2003/InkML" xmlns:x="urn:future"><x:future>0 0</x:future></i:trace>"#,
    ).unwrap()).unwrap();
    let mut generous = Edit::new(profile.clone()).unwrap();
    generous
        .add_data(ActionSelector::ordinal(0), data.clone())
        .unwrap();
    generous
        .finish()
        .expect("opaque payload is valid under the normal depth limit");
    let mut bounded = Edit::with_limits(
        profile,
        Limits {
            max_depth: 4,
            ..Limits::default()
        },
    )
    .unwrap();
    let result = bounded
        .add_data(ActionSelector::ordinal(0), data)
        .and_then(|()| bounded.finish());
    assert!(
        result.is_err(),
        "root/action/data/trace/descendant requires depth five"
    );
}

#[test]
fn malformed_opaque_trace_xml_is_refused_before_publication() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
    );
    for body in [
        "&unknown;",
        "&#x1;",
        "bad]]>text",
        "\u{1}",
        "<!-- bad--comment -->",
        "<x:future xmlns:x='urn:future' value='bad<value'/>",
        "<1bad/>",
    ] {
        let payload =
            format!(r#"<i:trace xmlns:i="http://www.w3.org/2003/InkML">{body}</i:trace>"#);
        let result = (|| {
            let data = ActionDataDraft::new().trace(OpaquePayload::new(payload)?)?;
            let mut edit = Edit::new(actions::read_profile(xml.as_bytes())?)?;
            edit.add_data(ActionSelector::ordinal(0), data)?;
            edit.finish()
        })();
        assert!(result.is_err(), "invalid opaque XML was accepted: {body:?}");
    }
}

#[test]
fn legal_opaque_comment_and_cdata_survive_authored_trace() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
    );
    let payload = r#"<i:trace xmlns:i="http://www.w3.org/2003/InkML"><!-- <? retained --> <![CDATA[<?x & raw <]]></i:trace>"#;
    let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
    let data = ActionDataDraft::new()
        .trace(OpaquePayload::new(payload).unwrap())
        .unwrap();
    edit.add_data(ActionSelector::ordinal(0), data).unwrap();
    let commit = edit.finish().unwrap();
    assert!(
        commit
            .as_bytes()
            .windows(payload.len())
            .any(|bytes| bytes == payload.as_bytes())
    );
    assert_eq!(
        commit.patch().inverse().apply(commit.as_bytes()).unwrap(),
        xml.as_bytes()
    );
}

#[test]
fn clearing_action_cannot_orphan_a_retained_data_reference() {
    let xml = format!(
        r##"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"><a:actionData xml:id="d1"/></a:action><a:action type="add" startTime="1"><a:actionData ref="#d1"/></a:action></a:actions>"##
    );
    let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
    let result = edit
        .clear(ActionSelector::ordinal(0))
        .and_then(|()| edit.finish());
    assert!(
        result.is_err(),
        "clear must preserve retained reference closure"
    );
}

#[test]
fn source_edit_refuses_invalid_payload_namespace_context_before_finish() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"/></a:actions>"#
    );
    for payload in [
        "<inkml:trace/>",
        "<inkml:trace xmlns:inkml=''/>",
        "<i:trace xmlns:i='http://www.w3.org/2003/InkML'><x:future/></i:trace>",
        "<i:trace xmlns:i='http://www.w3.org/2003/InkML' xmlns:x='urn:future' xmlns:y='urn:future' x:value='1' y:value='2'/>",
    ] {
        let mut edit = Edit::new(actions::read_profile(xml.as_bytes()).unwrap()).unwrap();
        let result = ActionDataDraft::new()
            .trace(OpaquePayload::new(payload).unwrap())
            .and_then(|data| edit.add_data(ActionSelector::ordinal(0), data));
        assert!(
            result.is_err(),
            "payload must be refused during staging: {payload}"
        );
        assert_eq!(edit.finish().unwrap().as_bytes(), xml.as_bytes());
    }
}

#[test]
fn cleared_children_release_the_final_node_budget() {
    let xml = format!(
        r#"<a:actions xmlns:a="{ACTION_NAMESPACE}" lengthUnit="cm" timeUnit="ms"><a:action type="add" startTime="0"><a:actionData/></a:action><a:action type="add" startTime="1"><a:actionData/></a:action></a:actions>"#
    );
    let profile = actions::read_profile(xml.as_bytes()).unwrap();
    let mut edit = Edit::with_limits(
        profile,
        Limits {
            max_nodes: 5,
            ..Limits::default()
        },
    )
    .unwrap();
    edit.clear(ActionSelector::ordinal(0)).unwrap();
    edit.add_data(ActionSelector::ordinal(1), ActionDataDraft::new())
        .unwrap();
    let commit = edit.finish().expect("five final nodes fit the node budget");
    assert_eq!(
        commit.patch().inverse().apply(commit.as_bytes()).unwrap(),
        xml.as_bytes()
    );
}

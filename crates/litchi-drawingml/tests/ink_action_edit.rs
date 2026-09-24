//! Independent tests for detached and source-backed ink-action editing.
//!
//! These tests cover the public action facade only.  InkML descendants are
//! deliberately used as opaque payloads; this suite does not assert a full
//! InkML model or action execution semantics.

use litchi_drawingml::ink::{
    ACTION_NAMESPACE, INKML_NAMESPACE,
    actions::{
        self, ActionChild, ActionDataDraft, ActionDraft, ActionGroupDraft, ActionParent,
        ActionSelector, ActionType, ChildSelector, DataChild, DataGroupDraft, DataSelector, Draft,
        Edit, LengthUnit, Limits, OpaquePayload, PropertyDraft, RootChild, TimeUnit,
    },
};
use litchi_drawingml::{Error, Result};

fn expect_invalid<T>(result: Result<T>, fragment: &str) {
    match result {
        Err(Error::Invalid(message)) => assert!(
            message.contains(fragment),
            "invalid error {message:?} does not contain {fragment:?}"
        ),
        Err(other) => panic!("expected Invalid containing {fragment:?}, got {other:?}"),
        Ok(_) => panic!("expected Invalid containing {fragment:?}, got success"),
    }
}

fn expect_limit<T>(result: Result<T>, resource: &'static str, limit: usize) {
    match result {
        Err(Error::Limit {
            resource: actual_resource,
            limit: actual_limit,
        }) => {
            assert_eq!(actual_resource, resource);
            assert_eq!(actual_limit, limit);
        },
        Err(other) => panic!("expected {resource:?} limit {limit}, got {other:?}"),
        Ok(_) => panic!("expected {resource:?} limit {limit}, got success"),
    }
}

fn source() -> Vec<u8> {
    format!(
        r##"<?xml version="1.0" encoding="UTF-8"?><!--leading-->
<a:actions xmlns:a="{ACTION_NAMESPACE}" xmlns:i="{INKML_NAMESPACE}" xmlns:x="urn:future" lengthUnit="cm" timeUnit="ms" xml:id="root">
  <!--root comment-->
  <i:definitions><!--definition comment--><x:future x:flag="v"/></i:definitions>
  <a:action xml:id="a0" type="add" startTime="0">
    <!--property comment-->
    <a:property name="kind" value="old"/>
    <a:actionData xml:id="d0" name="stroke" ref="#d1">
      <!--data comment-->
      <i:trace><x:future x:payload="yes"/></i:trace>
    </a:actionData>
    <a:actionDataGroup xml:id="dg0" name="group">
      <a:actionData xml:id="d1" name="other"><i:traceView/></a:actionData>
      <a:actionData xml:id="d2" name="third"><i:trace/></a:actionData>
    </a:actionDataGroup>
  </a:action>
  <a:actionGroup xml:id="ag0" type="transform" startTime="1.0">
    <!--group comment-->
    <a:action xml:id="a1" type="remove" startTime="2"><a:actionData/></a:action>
  </a:actionGroup>
</a:actions><!--trailing-->"##
    )
    .into_bytes()
}

fn profile() -> actions::Profile {
    actions::read_profile(&source()).expect("valid source-backed action profile")
}

fn bare_action(id: &str, action_type: ActionType, start_time: &str) -> ActionDraft {
    ActionDraft::new(action_type, start_time)
        .expect("action draft")
        .with_xml_id(id)
        .expect("action xml:id")
}

fn default_limits() -> Limits {
    Limits::default()
}

fn ordered_source() -> Vec<u8> {
    format!(
        r#"<p:actions xmlns:p="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"><p:action xml:id="a0" type="add" startTime="0"><p:property name="p0" value="zero"/><p:property name="p1" value="one"/><p:property name="p2" value="two"/></p:action><p:action xml:id="a1" type="remove" startTime="1"/><p:action xml:id="a2" type="transform" startTime="2"/></p:actions>"#
    )
    .into_bytes()
}

fn alias_only_source() -> Vec<u8> {
    format!(
        r#"<p:actions xmlns:p="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"><p:action xml:id="a0" type="add" startTime="0"/></p:actions>"#
    )
    .into_bytes()
}

fn self_closing_source() -> Vec<u8> {
    format!(r#"<p:actions xmlns:p="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"/>"#).into_bytes()
}

fn whitespace_close_source() -> Vec<u8> {
    format!(
        r#"<p:actions xmlns:p="{ACTION_NAMESPACE}" lengthUnit="m" timeUnit="s"><p:action xml:id="a0" type="add" startTime="0"/></p:actions   >"#
    )
    .into_bytes()
}

fn action_with_data_group_id(action_id: &str, data_group_id: &str) -> ActionDraft {
    let data_group = DataGroupDraft::new(ActionDataDraft::new())
        .expect("data group")
        .with_xml_id(data_group_id)
        .expect("data-group id");
    bare_action(action_id, ActionType::Add, "0")
        .data_group(data_group)
        .expect("data group child")
}

#[test]
fn detached_builder_round_trips_ordered_children_and_opaque_payloads() {
    let definitions = OpaquePayload::new(format!(
        r#"<inkml:definitions xmlns:inkml="{INKML_NAMESPACE}" xmlns:v="urn:vendor"><v:future/></inkml:definitions>"#
    ))
    .expect("definitions payload");
    let transform = OpaquePayload::new(format!(
        r#"<iact:transform xmlns:iact="{ACTION_NAMESPACE}" matrix="1,0,0,1"/>"#
    ))
    .expect("transform payload");
    let trace = OpaquePayload::new(format!(
        r#"<i:trace xmlns:i="{INKML_NAMESPACE}"><v:future xmlns:v="urn:vendor">opaque</v:future></i:trace>"#
    ))
    .expect("trace payload");
    let trace_view = OpaquePayload::new(format!(r#"<i:traceView xmlns:i="{INKML_NAMESPACE}"/>"#))
        .expect("trace-view payload");

    let property = PropertyDraft::new("name&<", "value \" ' & < >").expect("property");
    let data = ActionDataDraft::new()
        .with_xml_id("d-new")
        .expect("data id")
        .with_name("stroke & <future>")
        .expect("data name")
        .with_reference("#d-new&ref")
        .expect("data ref")
        .transform(transform.clone())
        .expect("transform first")
        .trace(trace.clone())
        .expect("trace")
        .trace_view(trace_view.clone())
        .expect("trace view");
    let grouped_data = ActionDataDraft::new()
        .with_name("group data")
        .expect("group data name");
    let data_group = DataGroupDraft::new(grouped_data)
        .expect("data-group")
        .with_xml_id("dg-new")
        .expect("data-group id")
        .with_name("group & data")
        .expect("data-group name");
    let action = bare_action(
        "a-new",
        ActionType::Custom("future&<\"'".into()),
        "  +0.25\n",
    )
    .property(property)
    .expect("property before data")
    .data(data)
    .expect("direct data")
    .data_group(data_group)
    .expect("data group");
    let group_action = bare_action("a-group", ActionType::Remove, "1");
    let group = ActionGroupDraft::with_metadata(ActionType::Transform, " 2.0 ", group_action)
        .expect("action group")
        .with_xml_id("ag-new")
        .expect("action-group id");

    let prepared = Draft::with_units(LengthUnit::Centimeter, TimeUnit::Millisecond)
        .expect("draft")
        .with_xml_id("root-new")
        .expect("root id")
        .definitions(definitions)
        .expect("definitions first")
        .action(action)
        .expect("direct action")
        .action_group(group)
        .expect("action group")
        .finish()
        .expect("detached authoring");
    assert_eq!(prepared.to_vec().expect("sink copy"), prepared.as_bytes());
    let xml = std::str::from_utf8(prepared.as_bytes()).expect("UTF-8 XML");
    assert!(xml.starts_with("<iact:actions xmlns:iact=\""));
    assert!(xml.contains("name&amp;&lt;"));
    assert!(xml.contains("value &quot; &apos; &amp; &lt; >"));
    assert!(xml.contains("type=\"future&amp;&lt;&quot;&apos;\""));
    assert!(xml.contains("<v:future xmlns:v=\"urn:vendor\">opaque</v:future>"));
    assert!(xml.contains("<iact:transform xmlns:iact=\""));

    let profile = prepared.profile();
    assert_eq!(profile.xml_id(), Some("root-new"));
    assert_eq!(profile.length_unit(), LengthUnit::Centimeter);
    assert_eq!(profile.time_unit(), TimeUnit::Millisecond);
    assert_eq!(profile.children().len(), 2);
    let RootChild::Action(action) = &profile.children()[0] else {
        panic!("first root child should be action")
    };
    assert_eq!(action.xml_id(), Some("a-new"));
    assert_eq!(action.action_type().as_str(), "future&<\"'");
    assert_eq!(action.start_time(), "+0.25");
    assert_eq!(action.properties().len(), 1);
    assert_eq!(action.properties()[0].name(), "name&<");
    assert_eq!(action.properties()[0].value(), "value \" ' & < >");
    assert_eq!(action.children().len(), 3);
    let ActionChild::Data(data) = &action.children()[1] else {
        panic!("second action child should be actionData")
    };
    assert_eq!(data.xml_id(), Some("d-new"));
    assert_eq!(data.name(), "stroke & <future>");
    assert_eq!(data.reference(), Some("#d-new&ref"));
    assert!(matches!(data.children()[0], DataChild::Transform(_)));
    assert_eq!(data.children().len(), 3);
    let RootChild::ActionGroup(group) = &profile.children()[1] else {
        panic!("second root child should be actionGroup")
    };
    assert_eq!(group.xml_id(), Some("ag-new"));
    assert_eq!(group.actions().len(), 1);
    assert_eq!(group.start_time(), "2.0");
}

#[test]
fn scalar_edits_preserve_source_spans_comments_and_opaque_descendants() {
    let original = source();
    let profile = actions::read_profile(&original).expect("profile");
    let mut edit = Edit::from_profile(&profile).expect("edit");
    let action = ActionSelector::ordinal(0);
    let data = DataSelector::Action { action, index: 0 };
    let property = ChildSelector::Property { action, index: 0 };
    edit.set_action_type(action, ActionType::Custom("future&<\"'".into()))
        .expect("action type");
    edit.set_start_time(action, " \n -2.50\t")
        .expect("start time");
    edit.set_property_name(property, "name&<\"'")
        .expect("property name");
    edit.set_property_value(property, "value&<\"'")
        .expect("property value");
    edit.set_data_name(data, "data&<\"'").expect("data name");
    edit.set_data_reference(data, Some("#d1&<\"'"))
        .expect("data reference");
    let commit = edit.finish().expect("scalar edit");
    let edited = commit.as_bytes();
    assert_ne!(edited, original.as_slice());
    assert_eq!(
        edited.as_ptr(),
        commit.prepared().profile().source().as_ptr(),
        "changed typed readback should reuse the emitted source allocation"
    );
    for preserved in [
        b"<!--leading-->".as_slice(),
        b"<!--root comment-->".as_slice(),
        b"<!--property comment-->".as_slice(),
        b"<!--data comment-->".as_slice(),
        b"<!--definition comment-->".as_slice(),
        b"<x:future x:payload=\"yes\"/>".as_slice(),
        b"<!--trailing-->".as_slice(),
    ] {
        assert!(
            edited
                .windows(preserved.len())
                .any(|window| window == preserved)
        );
    }
    let text = std::str::from_utf8(edited).expect("edited XML");
    assert!(text.contains("type=\"future&amp;&lt;&quot;&apos;\""));
    assert!(text.contains("name=\"name&amp;&lt;&quot;&apos;\""));
    assert!(text.contains("value=\"value&amp;&lt;&quot;&apos;\""));
    assert!(text.contains("ref=\"#d1&amp;&lt;&quot;&apos;\""));
    let updated = commit.prepared().profile();
    let RootChild::Action(action_value) = &updated.children()[0] else {
        panic!("root action expected")
    };
    assert_eq!(action_value.action_type().as_str(), "future&<\"'");
    assert_eq!(action_value.start_time(), "-2.50");
    assert_eq!(action_value.properties()[0].name(), "name&<\"'");
    assert_eq!(action_value.properties()[0].value(), "value&<\"'");
    let ActionChild::Data(data_value) = &action_value.children()[1] else {
        panic!("direct action data expected")
    };
    assert_eq!(data_value.name(), "data&<\"'");
    assert_eq!(data_value.reference(), Some("#d1&<\"'"));
}

#[test]
fn no_op_identity_edits_and_patches_are_exact_and_source_checked() {
    let original = source();
    let profile = actions::read_profile(&original).expect("profile");
    let mut identity = Edit::from_profile(&profile).expect("edit");
    identity
        .set_action_id(ActionSelector::ordinal(0), "a0")
        .expect("same action id is a no-op");
    identity
        .set_data_id(
            DataSelector::Action {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "d0",
        )
        .expect("same data id is a no-op");
    let no_op = identity.finish().expect("identity no-op");
    assert_eq!(no_op.as_bytes(), original.as_slice());
    assert_eq!(no_op.patch().before(), original.as_slice());
    assert_eq!(no_op.patch().after(), original.as_slice());
    assert_eq!(no_op.patch().apply(&original).expect("patch"), original);
    assert_eq!(
        no_op
            .patch()
            .inverse()
            .apply(&original)
            .expect("inverse patch"),
        original
    );
    let mut stale = original.clone();
    let last = stale.len() - 1;
    stale[last] = b'!';
    expect_invalid(no_op.patch().apply(&stale), "patch source is stale");

    let mut changed_id = Edit::from_profile(&profile).expect("edit");
    expect_invalid(
        changed_id.set_action_id(ActionSelector::ordinal(0), "renamed"),
        "modeled reference closure",
    );
    assert_eq!(changed_id.profile().source(), original.as_slice());
    let mut changed_data_id = Edit::from_profile(&profile).expect("edit");
    expect_invalid(
        changed_data_id.set_data_id(
            DataSelector::Action {
                action: ActionSelector::ordinal(0),
                index: 0,
            },
            "renamed-data",
        ),
        "modeled reference closure",
    );
    assert_eq!(changed_data_id.profile().source(), original.as_slice());
}

#[test]
fn add_move_and_remove_operations_preserve_order_and_reference_closure() {
    let profile = profile();
    let mut add = Edit::from_profile(&profile).expect("edit");
    add.add_action(ActionParent::Root, bare_action("a2", ActionType::Add, "3"))
        .expect("root action");
    add.add_action(
        ActionParent::Group(0),
        bare_action("a2g", ActionType::Remove, "4"),
    )
    .expect("group action");
    let added = add.finish().expect("add actions");
    assert_eq!(added.prepared().profile().actions().count(), 2);
    assert_eq!(
        added
            .prepared()
            .profile()
            .action_groups()
            .next()
            .expect("group")
            .actions()
            .len(),
        2
    );
    let mut reorder = Edit::from_profile(added.prepared().profile()).expect("reorder edit");
    reorder
        .move_before(ActionSelector::direct(1), ActionSelector::direct(0))
        .expect("same root sequence move");
    let reordered = reorder.finish().expect("reorder");
    let direct_actions: Vec<_> = reordered.prepared().profile().actions().collect();
    assert_eq!(direct_actions[0].xml_id(), Some("a2"));
    assert_eq!(direct_actions[1].xml_id(), Some("a0"));

    let mut duplicate = Edit::from_profile(&profile).expect("duplicate edit");
    duplicate
        .add_action(
            ActionParent::Root,
            bare_action("d0", ActionType::Custom("duplicate".into()), "5"),
        )
        .expect("queue duplicate");
    expect_invalid(duplicate.finish(), "duplicate xml:id values");

    let mut remove_referenced = Edit::from_profile(&profile).expect("remove edit");
    remove_referenced
        .remove_group_data(DataSelector::Group {
            action: ActionSelector::ordinal(0),
            group: 0,
            index: 0,
        })
        .expect("queue referenced removal");
    expect_invalid(remove_referenced.finish(), "break an xml:id reference");

    let mut remove = Edit::from_profile(&profile).expect("remove edit");
    let action = ActionSelector::ordinal(0);
    let data = DataSelector::Action { action, index: 0 };
    remove
        .remove_property(ChildSelector::Property { action, index: 0 })
        .expect("remove property");
    remove
        .remove_group_data(DataSelector::Group {
            action,
            group: 0,
            index: 0,
        })
        .expect("remove one group data");
    remove
        .set_data_reference(data, None::<&str>)
        .expect("clear reference before removing data");
    remove.remove_data(data).expect("remove direct data");
    let removed = remove.finish().expect("removals");
    let RootChild::Action(action_value) = &removed.prepared().profile().children()[0] else {
        panic!("root action expected")
    };
    assert!(action_value.properties().is_empty());
    assert!(matches!(
        action_value.children()[0],
        ActionChild::DataGroup(_)
    ));
    let ActionChild::DataGroup(group) = &action_value.children()[0] else {
        panic!("data group expected")
    };
    assert_eq!(group.data().len(), 1);
    assert_eq!(group.data()[0].xml_id(), Some("d2"));

    let mut remove_action = Edit::from_profile(&profile).expect("action removal");
    remove_action
        .set_data_reference(data, None::<&str>)
        .expect("clear reference");
    remove_action
        .remove_action(action)
        .expect("remove root action");
    let removed_action = remove_action.finish().expect("root action removed");
    assert_eq!(removed_action.prepared().profile().actions().count(), 0);
    assert_eq!(
        removed_action.prepared().profile().action_groups().count(),
        1
    );
}

#[test]
fn forward_and_inverse_moves_reorder_source_sequence() {
    let initial = actions::read_profile(&ordered_source()).expect("ordered profile");
    let mut forward = Edit::from_profile(&initial).expect("forward move edit");
    forward
        .move_before(ActionSelector::direct(0), ActionSelector::direct(2))
        .expect("move first action before third");
    let moved = forward.finish().expect("forward move");
    let moved_ids: Vec<_> = moved
        .prepared()
        .profile()
        .actions()
        .map(|action| action.xml_id())
        .collect();
    assert_eq!(moved_ids, [Some("a1"), Some("a0"), Some("a2")]);

    let mut inverse = Edit::from_profile(moved.prepared().profile()).expect("inverse edit");
    inverse
        .move_before(ActionSelector::direct(1), ActionSelector::direct(0))
        .expect("move first action back before second");
    let restored = inverse.finish().expect("inverse move");
    let restored_ids: Vec<_> = restored
        .prepared()
        .profile()
        .actions()
        .map(|action| action.xml_id())
        .collect();
    assert_eq!(restored_ids, [Some("a0"), Some("a1"), Some("a2")]);
}

#[test]
fn failed_edits_are_atomic_and_invalid_selectors_do_not_mutate_source_snapshot() {
    let profile = profile();
    let original = profile.source().to_vec();
    let mut bad_selector = Edit::from_profile(&profile).expect("edit");
    bad_selector
        .remove_data(DataSelector::Action {
            action: ActionSelector::ordinal(99),
            index: 0,
        })
        .expect("queue selector check until finish");
    assert_eq!(bad_selector.profile().source(), original.as_slice());
    expect_invalid(
        bad_selector.finish(),
        "action ordinal selector is out of range",
    );

    let mut bad_group = Edit::from_profile(&profile).expect("edit");
    bad_group
        .add_action(
            ActionParent::Group(99),
            bare_action("new", ActionType::Add, "1"),
        )
        .expect("queue group check until finish");
    assert_eq!(bad_group.profile().source(), original.as_slice());
    expect_invalid(bad_group.finish(), "action group selector is out of range");

    let mut wrong_remove = Edit::from_profile(&profile).expect("edit");
    expect_invalid(
        wrong_remove.remove_data_group(ChildSelector::Data {
            action: ActionSelector::ordinal(0),
            index: 0,
        }),
        "not a data group",
    );
    assert_eq!(wrong_remove.profile().source(), original.as_slice());
}

#[test]
fn detached_limits_reserve_output_payload_nodes_depth_actions_groups_and_scalars() {
    let mut invalid_limits = default_limits();
    invalid_limits.max_output_bytes = 0;
    expect_limit(
        Draft::new(invalid_limits, LengthUnit::Meter, TimeUnit::Second),
        "ink action output bytes",
        litchi_drawingml::ink::MAX_SOURCE_BYTES,
    );

    let action = || bare_action("a", ActionType::Add, "0");
    let mut output_limit = default_limits();
    output_limit.max_output_bytes = 64;
    expect_limit(
        Draft::new(output_limit, LengthUnit::Meter, TimeUnit::Second)
            .expect("output-limited draft")
            .action(action())
            .expect("action")
            .finish(),
        "ink action output bytes",
        64,
    );

    let mut payload_limit = default_limits();
    payload_limit.max_payload_bytes = 4;
    let payload = OpaquePayload::new(b"<i:trace xmlns:i=\"urn:test\"/>").expect("payload");
    let with_payload = bare_action("a", ActionType::Add, "0")
        .data(ActionDataDraft::new().trace(payload).expect("trace child"))
        .expect("data child");
    expect_limit(
        Draft::new(payload_limit, LengthUnit::Meter, TimeUnit::Second)
            .expect("payload-limited draft")
            .action(with_payload)
            .expect("action")
            .finish(),
        "ink action payload bytes",
        4,
    );

    let mut nodes_limit = default_limits();
    nodes_limit.max_nodes = 1;
    expect_limit(
        Draft::new(nodes_limit, LengthUnit::Meter, TimeUnit::Second)
            .expect("node-limited draft")
            .action(action())
            .expect("action")
            .finish(),
        "ink action XML nodes",
        1,
    );

    let mut depth_limit = default_limits();
    depth_limit.max_depth = 2;
    let nested = bare_action("a", ActionType::Add, "0")
        .data(ActionDataDraft::new())
        .expect("data child");
    expect_limit(
        Draft::new(depth_limit, LengthUnit::Meter, TimeUnit::Second)
            .expect("depth-limited draft")
            .action(nested)
            .expect("action")
            .finish(),
        "ink action XML depth",
        2,
    );

    let mut action_limit = default_limits();
    action_limit.max_actions = 1;
    let two_actions = Draft::new(action_limit, LengthUnit::Meter, TimeUnit::Second)
        .expect("action-limited draft")
        .action(action())
        .expect("first action")
        .action(bare_action("b", ActionType::Remove, "1"))
        .expect("second action");
    expect_limit(two_actions.finish(), "ink action records", 1);

    let mut group_limit = default_limits();
    group_limit.max_action_groups = 1;
    let first_group = ActionGroupDraft::new(action()).expect("first group");
    let second_group =
        ActionGroupDraft::new(bare_action("b", ActionType::Remove, "1")).expect("second group");
    let two_groups = Draft::new(group_limit, LengthUnit::Meter, TimeUnit::Second)
        .expect("group-limited draft")
        .action_group(first_group)
        .expect("first group")
        .action_group(second_group)
        .expect("second group");
    expect_limit(two_groups.finish(), "ink action groups", 1);

    let mut scalar_limit = default_limits();
    scalar_limit.max_scalar_bytes = 4;
    let long_type = bare_action("a", ActionType::Custom("five!".into()), "0");
    expect_limit(
        Draft::new(scalar_limit, LengthUnit::Meter, TimeUnit::Second)
            .expect("scalar-limited draft")
            .action(long_type)
            .expect("action")
            .finish(),
        "action type",
        4,
    );
}

#[test]
fn edit_output_limit_is_checked_before_replacement_allocation() {
    let profile = profile();
    let mut limits = default_limits();
    limits.max_output_bytes = profile.source().len();
    let mut edit = Edit::with_limits(profile.clone(), limits).expect("source fits bound");
    edit.set_property_value(
        ChildSelector::Property {
            action: ActionSelector::ordinal(0),
            index: 0,
        },
        "expanded property value that exceeds the exact source budget",
    )
    .expect("queue expansion");
    expect_limit(
        edit.finish(),
        "ink action output bytes",
        limits.max_output_bytes,
    );
    assert_eq!(profile.source(), source().as_slice());
}

#[test]
fn scalar_control_references_have_exact_lengths_in_detached_and_source_edits() {
    let controls = "\t\n\r";
    let property = PropertyDraft::new(format!("property{controls}"), format!("value{controls}"))
        .expect("control property");
    let data = ActionDataDraft::new()
        .with_name(format!("data{controls}"))
        .expect("control data name")
        .with_reference(format!("reference{controls}"))
        .expect("control data reference");
    let action = ActionDraft::new(
        ActionType::custom(format!("custom{controls}")).expect("control action type"),
        "0",
    )
    .expect("control action")
    .property(property)
    .expect("control property order")
    .data(data)
    .expect("control data");
    let prepared = Draft::with_units(LengthUnit::Meter, TimeUnit::Second)
        .expect("control draft")
        .action(action)
        .expect("control root action")
        .finish()
        .expect("control detached output");
    let detached_xml = std::str::from_utf8(prepared.as_bytes()).expect("detached XML");
    assert!(detached_xml.contains("&#x9;&#xA;&#xD;"));
    let RootChild::Action(detached_action) = &prepared.profile().children()[0] else {
        panic!("detached action expected")
    };
    assert_eq!(
        detached_action.action_type().as_str(),
        format!("custom{controls}")
    );
    assert_eq!(
        detached_action.properties()[0].name(),
        format!("property{controls}")
    );
    assert_eq!(
        detached_action.properties()[0].value(),
        format!("value{controls}")
    );
    let ActionChild::Data(detached_data) = &detached_action.children()[1] else {
        panic!("detached data expected")
    };
    assert_eq!(detached_data.name(), format!("data{controls}"));
    assert_eq!(
        detached_data.reference(),
        Some(format!("reference{controls}").as_str())
    );

    let profile = profile();
    let mut edit = Edit::from_profile(&profile).expect("source edit");
    let action = ActionSelector::ordinal(0);
    let property = ChildSelector::Property { action, index: 0 };
    let data = DataSelector::Action { action, index: 0 };
    edit.set_action_type(action, ActionType::custom(controls).expect("control type"))
        .expect("source action type");
    edit.set_property_value(property, controls)
        .expect("source property value");
    edit.set_data_name(data, controls)
        .expect("source data name");
    edit.set_data_reference(data, Some(controls))
        .expect("source data reference");
    let commit = edit.finish().expect("source control output");
    let source_xml = std::str::from_utf8(commit.as_bytes()).expect("source XML");
    assert!(source_xml.contains("type=\"&#x9;&#xA;&#xD;\""));
    assert!(source_xml.contains("value=\"&#x9;&#xA;&#xD;\""));
    assert!(source_xml.contains("name=\"&#x9;&#xA;&#xD;\""));
    assert!(source_xml.contains("ref=\"&#x9;&#xA;&#xD;\""));
}

#[test]
fn source_selectors_remain_stable_across_removal_and_move() {
    let ordered = ordered_source();
    let profile = actions::read_profile(&ordered).expect("ordered profile");
    let mut remove_then_edit = Edit::from_profile(&profile).expect("remove edit");
    let first = ActionSelector::ordinal(0);
    remove_then_edit
        .remove_property(ChildSelector::Property {
            action: first,
            index: 0,
        })
        .expect("remove first property");
    remove_then_edit
        .set_property_value(
            ChildSelector::Property {
                action: first,
                index: 1,
            },
            "edited-after-removal",
        )
        .expect("edit property after removal");
    let removed = remove_then_edit.finish().expect("remove then edit");
    let RootChild::Action(removed_action) = &removed.prepared().profile().children()[0] else {
        panic!("removed action expected")
    };
    assert_eq!(removed_action.properties().len(), 2);
    assert_eq!(removed_action.properties()[0].name(), "p1");
    assert_eq!(
        removed_action.properties()[0].value(),
        "edited-after-removal"
    );
    assert_eq!(removed_action.properties()[1].name(), "p2");
    assert_eq!(removed_action.properties()[1].value(), "two");

    let mut move_then_edit = Edit::from_profile(&profile).expect("move edit");
    move_then_edit
        .move_before(ActionSelector::direct(2), ActionSelector::direct(0))
        .expect("move action");
    move_then_edit
        .set_action_type(ActionSelector::direct(1), ActionType::Transform)
        .expect("edit moved sequence");
    let moved = move_then_edit.finish().expect("move then edit");
    let moved_actions: Vec<_> = moved.prepared().profile().actions().collect();
    assert_eq!(moved_actions[0].xml_id(), Some("a2"));
    assert_eq!(moved_actions[0].action_type().as_str(), "transform");
    assert_eq!(moved_actions[1].xml_id(), Some("a0"));
    assert_eq!(moved_actions[1].action_type().as_str(), "add");
    assert_eq!(moved_actions[2].xml_id(), Some("a1"));
    assert_eq!(moved_actions[2].action_type().as_str(), "transform");
}

#[test]
fn equal_scalar_edit_reuses_the_retained_source_when_observable() {
    let profile = profile();
    let source_pointer = profile.source().as_ptr();
    let mut edit = Edit::from_profile(&profile).expect("equal edit");
    edit.set_property_value(
        ChildSelector::Property {
            action: ActionSelector::ordinal(0),
            index: 0,
        },
        "old",
    )
    .expect("equal property value");
    let commit = edit.finish().expect("equal edit result");
    assert_eq!(commit.as_bytes(), profile.source());
    assert_eq!(commit.as_bytes().as_ptr(), source_pointer);
}

#[test]
fn identity_reservations_cover_root_action_group_and_data_group_ids() {
    for duplicate_id in ["root", "ag0", "dg0"] {
        let mut edit = Edit::from_profile(&profile()).expect("duplicate edit");
        edit.add_action(
            ActionParent::Root,
            bare_action(duplicate_id, ActionType::Add, "3"),
        )
        .expect("queue duplicate action");
        expect_invalid(edit.finish(), "duplicate xml:id values");
    }

    let mut data_group_duplicate = Edit::from_profile(&profile()).expect("data-group edit");
    data_group_duplicate
        .add_action(
            ActionParent::Root,
            action_with_data_group_id("a-new", "dg0"),
        )
        .expect("queue duplicate data group");
    expect_invalid(data_group_duplicate.finish(), "duplicate xml:id values");

    let root_and_group_duplicate = Draft::with_units(LengthUnit::Meter, TimeUnit::Second)
        .expect("root/group draft")
        .with_xml_id("same-id")
        .expect("root id")
        .action_group(
            ActionGroupDraft::new(bare_action("group-action", ActionType::Add, "0"))
                .expect("group")
                .with_xml_id("same-id")
                .expect("group id"),
        )
        .expect("root/group child");
    expect_invalid(root_and_group_duplicate.finish(), "duplicate xml:id");

    let first_group = DataGroupDraft::new(ActionDataDraft::new())
        .expect("first data group")
        .with_xml_id("same-data-group")
        .expect("first data-group id");
    let second_group = DataGroupDraft::new(ActionDataDraft::new())
        .expect("second data group")
        .with_xml_id("same-data-group")
        .expect("second data-group id");
    let duplicate_data_groups = Draft::with_units(LengthUnit::Meter, TimeUnit::Second)
        .expect("data-group draft")
        .action(
            bare_action("data-group-action", ActionType::Add, "0")
                .data_group(first_group)
                .expect("first data group child")
                .data_group(second_group)
                .expect("second data group child"),
        )
        .expect("data-group action");
    expect_invalid(duplicate_data_groups.finish(), "duplicate xml:id");
}

#[test]
fn removing_an_action_referenced_by_data_is_refused() {
    let referenced_source = String::from_utf8(source())
        .expect("source UTF-8")
        .replace("ref=\"#d1\"", "ref=\"#a1\"")
        .into_bytes();
    let profile = actions::read_profile(&referenced_source).expect("referenced profile");
    let mut edit = Edit::from_profile(&profile).expect("referenced edit");
    edit.remove_action(ActionSelector::ordinal(1))
        .expect("queue referenced action removal");
    expect_invalid(edit.finish(), "break an xml:id reference");
}

#[test]
fn restored_reference_before_target_removal_is_still_refused() {
    let profile = profile();
    let d0 = DataSelector::Action {
        action: ActionSelector::ordinal(0),
        index: 0,
    };
    let d1 = DataSelector::Group {
        action: ActionSelector::ordinal(0),
        group: 0,
        index: 0,
    };
    let mut edit = Edit::from_profile(&profile).expect("reference edit");
    edit.set_data_reference(d0, None::<&str>)
        .expect("clear source reference");
    edit.set_data_reference(d0, Some("#d1"))
        .expect("restore source reference");
    edit.remove_group_data(d1)
        .expect("queue referenced target removal");
    expect_invalid(edit.finish(), "break an xml:id reference");
}

#[test]
fn removing_a_referenced_data_group_id_is_refused() {
    let referenced_source = String::from_utf8(source())
        .expect("source UTF-8")
        .replace("ref=\"#d1\"", "ref=\"#dg0\"")
        .into_bytes();
    let profile = actions::read_profile(&referenced_source).expect("referenced group profile");
    let mut edit = Edit::from_profile(&profile).expect("group removal edit");
    edit.remove_data_group(ChildSelector::DataGroup {
        action: ActionSelector::ordinal(0),
        index: 0,
    })
    .expect("queue referenced data-group removal");
    expect_invalid(edit.finish(), "break an xml:id reference");
}

#[test]
fn changed_edits_refuse_duplicate_ids_but_exact_no_op_replays() {
    let duplicate_source = String::from_utf8(source())
        .expect("source UTF-8")
        .replace("xml:id=\"d2\"", "xml:id=\"d1\"")
        .into_bytes();
    let profile = actions::read_profile(&duplicate_source).expect("duplicate-id profile");
    let property = ChildSelector::Property {
        action: ActionSelector::ordinal(0),
        index: 0,
    };

    let mut no_op = Edit::from_profile(&profile).expect("duplicate no-op edit");
    no_op
        .set_property_value(property, "old")
        .expect("queue exact no-op");
    let retained = no_op.finish().expect("exact duplicate no-op");
    assert_eq!(retained.as_bytes(), duplicate_source.as_slice());

    let mut changed = Edit::from_profile(&profile).expect("duplicate changed edit");
    changed
        .set_property_value(property, "changed")
        .expect("queue duplicate-source edit");
    expect_invalid(changed.finish(), "duplicate xml:id");
}

fn expect_aggregate_limit(
    profile: &actions::Profile,
    limits: Limits,
    action: ActionDraft,
    resource: &'static str,
    value: usize,
) {
    match Edit::with_limits(profile.clone(), limits) {
        Err(error) => expect_limit::<()>(Err(error), resource, value),
        Ok(mut edit) => {
            edit.add_action(ActionParent::Root, action)
                .expect("queue aggregate action");
            expect_limit(edit.finish(), resource, value);
        },
    }
}

#[test]
fn aggregate_existing_and_added_records_are_bounded() {
    let profile = profile();
    let mut action_limits = default_limits();
    action_limits.max_actions = 2;
    expect_aggregate_limit(
        &profile,
        action_limits,
        bare_action("aggregate-action", ActionType::Add, "3"),
        "ink action records",
        2,
    );

    let mut node_limits = default_limits();
    node_limits.max_nodes = 2;
    expect_aggregate_limit(
        &profile,
        node_limits,
        bare_action("aggregate-node", ActionType::Add, "3"),
        "ink action XML nodes",
        2,
    );
}

#[test]
fn alias_only_self_closing_and_whitespace_close_roots_are_handled_without_rewriting() {
    let alias_source = alias_only_source();
    let alias_profile = actions::read_profile(&alias_source).expect("alias-only profile");
    let mut alias_edit = Edit::from_profile(&alias_profile).expect("alias-only edit");
    alias_edit
        .add_action(
            ActionParent::Root,
            bare_action("alias-added", ActionType::Add, "1"),
        )
        .expect("alias-only add");
    let alias_commit = alias_edit.finish().expect("alias-only result");
    assert!(
        std::str::from_utf8(alias_commit.as_bytes())
            .expect("alias-only XML")
            .contains("<p:action xml:id=\"alias-added\"")
    );
    assert_eq!(alias_commit.prepared().profile().actions().count(), 2);

    let self_closing_profile =
        actions::read_profile(&self_closing_source()).expect("self-closing profile");
    let mut first_action_edit =
        Edit::from_profile(&self_closing_profile).expect("self-closing action edit");
    first_action_edit
        .add_action(
            ActionParent::Root,
            bare_action("self-first", ActionType::Add, "0"),
        )
        .expect("queue first self-closing action");
    let first_action = first_action_edit
        .finish()
        .expect("first self-closing action");
    let first_action_xml = std::str::from_utf8(first_action.as_bytes()).expect("action XML");
    assert!(first_action_xml.contains("<p:actions xmlns:p="));
    assert!(first_action_xml.contains("</p:actions>"));
    assert_eq!(first_action.prepared().profile().actions().count(), 1);

    let mut first_group_edit =
        Edit::from_profile(&self_closing_profile).expect("self-closing group edit");
    first_group_edit
        .add_action_group(
            ActionGroupDraft::new(bare_action("self-group-action", ActionType::Remove, "1"))
                .expect("self-closing group"),
        )
        .expect("queue first self-closing group");
    let first_group = first_group_edit.finish().expect("first self-closing group");
    let first_group_xml = std::str::from_utf8(first_group.as_bytes()).expect("group XML");
    assert!(first_group_xml.contains("<p:actionGroup"));
    assert!(first_group_xml.contains("</p:actions>"));
    assert_eq!(first_group.prepared().profile().action_groups().count(), 1);
    let RootChild::ActionGroup(group) = &first_group.prepared().profile().children()[0] else {
        panic!("root action group expected")
    };
    assert_eq!(group.actions().len(), 1);
    assert_eq!(group.actions()[0].xml_id(), Some("self-group-action"));

    let mut multi_add = Edit::from_profile(&self_closing_profile).expect("multi-add edit");
    multi_add
        .add_action(
            ActionParent::Root,
            bare_action("self-first", ActionType::Add, "0"),
        )
        .expect("queue first multi-add action");
    multi_add
        .add_action(
            ActionParent::Root,
            bare_action("self-second", ActionType::Transform, "1"),
        )
        .expect("queue second multi-add action");
    let multi_commit = multi_add.finish().expect("multi-add self-closing root");
    let multi_xml = std::str::from_utf8(multi_commit.as_bytes()).expect("multi-add XML");
    assert_eq!(multi_xml.matches("<p:actions").count(), 1);
    assert_eq!(multi_xml.matches("</p:actions>").count(), 1);
    let multi_ids: Vec<_> = multi_commit
        .prepared()
        .profile()
        .actions()
        .map(|action| action.xml_id())
        .collect();
    assert_eq!(multi_ids, [Some("self-first"), Some("self-second")]);
    assert_eq!(
        multi_commit
            .patch()
            .inverse()
            .apply(multi_commit.as_bytes())
            .expect("multi-add inverse"),
        self_closing_source()
    );

    let whitespace_source = whitespace_close_source();
    let whitespace_profile =
        actions::read_profile(&whitespace_source).expect("whitespace-close profile");
    let mut whitespace_edit = Edit::from_profile(&whitespace_profile).expect("whitespace edit");
    whitespace_edit
        .add_action(
            ActionParent::Root,
            bare_action("whitespace-added", ActionType::Add, "1"),
        )
        .expect("whitespace-close add");
    let whitespace_commit = whitespace_edit.finish().expect("whitespace-close result");
    assert_eq!(whitespace_commit.prepared().profile().actions().count(), 2);
    assert!(
        std::str::from_utf8(whitespace_commit.as_bytes())
            .expect("whitespace-close XML")
            .contains("</p:actions   >")
    );
}

#[test]
fn absent_data_reference_none_is_an_exact_no_op() {
    let profile = profile();
    let data = DataSelector::Group {
        action: ActionSelector::ordinal(0),
        group: 0,
        index: 0,
    };
    let source_pointer = profile.source().as_ptr();
    let mut edit = Edit::from_profile(&profile).expect("absent-reference edit");
    edit.set_data_reference(data, None::<&str>)
        .expect("queue absent-reference no-op");
    let commit = edit.finish().expect("absent-reference no-op");
    assert_eq!(commit.as_bytes(), profile.source());
    assert_eq!(commit.as_bytes().as_ptr(), source_pointer);
}

#[test]
fn group_and_data_group_metadata_edits_preserve_source_and_exact_no_ops() {
    let profile = profile();
    let group_data = ChildSelector::DataGroup {
        action: ActionSelector::ordinal(0),
        index: 0,
    };
    let mut edit = Edit::from_profile(&profile).expect("metadata edit");
    edit.set_action_group_type(0, ActionType::Remove)
        .expect("group type");
    edit.set_group_start_time(0, "2.5")
        .expect("group start time alias");
    edit.set_data_group_name(group_data, "renamed-group")
        .expect("data group name");
    let commit = edit.finish().expect("metadata result");
    let xml = std::str::from_utf8(commit.as_bytes()).expect("metadata XML");
    assert!(xml.contains("type=\"remove\" startTime=\"2.5\""));
    assert!(xml.contains("<a:actionDataGroup xml:id=\"dg0\" name=\"renamed-group\""));
    assert!(xml.contains("<!--group comment-->"));
    let RootChild::ActionGroup(group) = &commit.prepared().profile().children()[1] else {
        panic!("group child expected")
    };
    assert_eq!(group.action_type().as_str(), "remove");
    assert_eq!(group.start_time(), "2.5");
    let RootChild::Action(action) = &commit.prepared().profile().children()[0] else {
        panic!("action child expected")
    };
    let Some(ActionChild::DataGroup(data_group)) = action
        .children()
        .iter()
        .find(|child| matches!(child, ActionChild::DataGroup(_)))
    else {
        panic!("data group child expected")
    };
    assert_eq!(data_group.name(), "renamed-group");

    let source_pointer = profile.source().as_ptr();
    let mut no_op = Edit::from_profile(&profile).expect("metadata no-op");
    no_op
        .set_group_type(0, ActionType::Transform)
        .expect("same group type alias");
    no_op
        .set_action_group_start_time(0, "1.0")
        .expect("same group start time");
    no_op
        .set_data_group_name(group_data, "group")
        .expect("same data group name");
    let retained = no_op.finish().expect("metadata no-op result");
    assert_eq!(retained.as_bytes(), profile.source());
    assert_eq!(retained.as_bytes().as_ptr(), source_pointer);
}

#[test]
fn adding_data_then_property_to_empty_action_preserves_schema_order_and_inverse() {
    let source = alias_only_source();
    let profile = actions::read_profile(&source).expect("empty action profile");
    let action = ActionSelector::ordinal(0);
    let mut edit = Edit::from_profile(&profile).expect("empty action edit");
    edit.add_data(action, ActionDataDraft::new())
        .expect("queue data insertion");
    edit.add_property(
        action,
        PropertyDraft::new("added-property", "added-value").expect("property"),
    )
    .expect("queue property insertion");
    let commit = edit.finish().expect("data then property edit");
    let xml = std::str::from_utf8(commit.as_bytes()).expect("edited empty action XML");
    assert!(
        xml.find("<p:property").expect("property XML")
            < xml.find("<p:actionData").expect("data XML")
    );
    let action_value = &commit.prepared().profile().children()[0];
    let RootChild::Action(action_value) = action_value else {
        panic!("root action expected")
    };
    assert_eq!(action_value.properties().len(), 1);
    assert!(matches!(
        action_value.children()[0],
        ActionChild::Property(_)
    ));
    assert!(matches!(action_value.children()[1], ActionChild::Data(_)));
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.as_bytes())
            .expect("empty action inverse"),
        source
    );
}

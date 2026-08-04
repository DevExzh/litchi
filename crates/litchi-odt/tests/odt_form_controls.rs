//! Classic ODF form controls (`office:forms`) in text documents.
//!
//! The reader parses form containers and the common control kinds into
//! strictly typed inert data — control ids, names, labels, current values,
//! checked/selected state, options, and event-listener metadata that is
//! retained but never executed — while `DocumentBuilder` and
//! `MutableDocument` author, edit, and remove controls with packaged
//! round trips.

use litchi_odt::{
    ButtonControl, CheckboxControl, CheckboxState, ComboItem, ComboboxControl, ControlForm,
    Document, DocumentBuilder, FixedTextControl, FlatOpenDocument, FormControlKind, FormNode,
    GenericForm, HiddenControl, InteractiveForm, ListOption, ListboxControl, MutableDocument,
    RadioControl, SelectionForm, TextControl, TypedValueControl, TypedValueControlKind,
    TypedValueForm, VisualForm,
};

/// Flat text document holding one form with every common control kind.
const FIXTURE: &str = "../../test-data/odf/odt/form-controls.fodt";

fn text_form() -> ControlForm {
    let mut form = ControlForm::new("texts");
    let mut name = TextControl::text("Name", "name_field");
    name.current_value = Some("Ada".to_string());
    name.title = Some("Your name".to_string());
    name.max_length = Some(40);
    form.add_control(name).unwrap();
    let mut bio = TextControl::textarea("Bio", "bio_field");
    bio.readonly = Some(true);
    bio.paragraphs = vec!["first line".to_string(), "second line".to_string()];
    form.add_control(bio).unwrap();
    form
}

fn interactive_form() -> InteractiveForm {
    let mut form = InteractiveForm::new("buttons");
    let mut subscribe = CheckboxControl::new("Subscribe", "subscribe_box");
    subscribe.label = Some("Subscribe to newsletter".to_string());
    subscribe.current_state = Some(CheckboxState::Checked);
    form.add_control(subscribe).unwrap();
    let mut go = ButtonControl::new("Go", "go_button");
    go.label = Some("Go".to_string());
    form.add_control(go).unwrap();
    form
}

fn selection_form() -> SelectionForm {
    let mut form = SelectionForm::new("choices");
    let mut city = ComboboxControl::new("City", "city_combo");
    city.current_value = Some("Paris".to_string());
    city.dropdown = Some(true);
    city.items = vec![
        ComboItem {
            label: Some("Paris".to_string()),
            text: String::new(),
        },
        ComboItem {
            label: Some("London".to_string()),
            text: String::new(),
        },
    ];
    form.add_control(city).unwrap();
    let mut color = ListboxControl::new("Color", "color_list");
    color.dropdown = Some(true);
    color.options = vec![
        ListOption {
            label: Some("Red".to_string()),
            value: Some("r".to_string()),
            selected: None,
            current_selected: Some(true),
            text: String::new(),
        },
        ListOption {
            label: Some("Blue".to_string()),
            value: Some("b".to_string()),
            selected: None,
            current_selected: None,
            text: String::new(),
        },
    ];
    form.add_control(color).unwrap();
    form
}

fn typed_value_form() -> TypedValueForm {
    let mut form = TypedValueForm::new("values");
    let mut age = TypedValueControl::new(TypedValueControlKind::Number, "Age", "age_number");
    age.current_value = Some("36".to_string());
    form.add_control(age).unwrap();
    let mut when = TypedValueControl::new(TypedValueControlKind::Date, "When", "when_date");
    when.current_value = Some("2024-01-31".to_string());
    form.add_control(when).unwrap();
    let mut at = TypedValueControl::new(TypedValueControlKind::Time, "At", "at_time");
    at.current_value = Some("12:30:00".to_string());
    form.add_control(at).unwrap();
    form
}

fn generic_form() -> GenericForm {
    let mut form = GenericForm::new("labels");
    let mut caption = FixedTextControl::new("Caption", "caption_label");
    caption.label = Some("Static caption".to_string());
    caption.form_for = Some("name_field".to_string());
    form.add_control(caption).unwrap();
    let mut token = HiddenControl::new("Token", "hidden_token");
    token.value = Some("inert-token".to_string());
    form.add_control(token).unwrap();
    form
}

fn visual_form() -> VisualForm {
    let mut form = VisualForm::new("radios");
    let mut choice = RadioControl::new("Choice", "choice_a");
    choice.label = Some("Option A".to_string());
    choice.value = Some("a".to_string());
    choice.current_selected = Some(true);
    form.add_control(choice).unwrap();
    form
}

#[test]
fn reads_fixture_form_controls_as_typed_inert_data() {
    let document = FlatOpenDocument::open(FIXTURE).unwrap();
    let forms = document.forms().unwrap();
    assert_eq!(forms.groups.len(), 1);
    assert!(forms.groups[0].apply_design_mode == Some(true));
    assert_eq!(forms.groups[0].forms.len(), 1);
    let form = &forms.groups[0].forms[0];
    assert_eq!(form.name.as_deref(), Some("registration"));

    let controls: Vec<&litchi_odt::FormControl> = form
        .children
        .iter()
        .map(|node| match node {
            FormNode::Control(control) => control,
            FormNode::Form(_) => panic!("no nested forms in fixture"),
        })
        .collect();
    let kinds: Vec<&FormControlKind> = controls.iter().map(|control| &control.kind).collect();
    assert_eq!(
        kinds,
        [
            &FormControlKind::Text,
            &FormControlKind::TextArea,
            &FormControlKind::CheckBox,
            &FormControlKind::Button,
            &FormControlKind::ComboBox,
            &FormControlKind::ListBox,
            &FormControlKind::Radio,
            &FormControlKind::FixedText,
            &FormControlKind::Hidden,
            &FormControlKind::Number,
            &FormControlKind::Date,
            &FormControlKind::Time,
        ]
    );

    // Ids, names, labels, current values, and checked state are typed data.
    assert_eq!(controls[0].xml_id.as_deref(), Some("name_field"));
    assert_eq!(controls[0].name.as_deref(), Some("Name"));
    assert_eq!(controls[0].current_value.as_deref(), Some("Ada"));
    assert_eq!(controls[2].current_state.as_deref(), Some("checked"));
    assert_eq!(
        controls[2].label.as_deref(),
        Some("Subscribe to newsletter")
    );
    assert_eq!(controls[4].current_value.as_deref(), Some("Paris"));
    assert_eq!(controls[6].current_selected, Some(true));
    assert_eq!(controls[8].value.as_deref(), Some("inert-token"));
    assert_eq!(controls[9].current_value.as_deref(), Some("36"));

    // Combobox items and listbox options are nested typed controls.
    let FormNode::Control(item) = &controls[4].children[0] else {
        panic!("combobox items are nested controls")
    };
    assert_eq!(item.kind, FormControlKind::Item);
    assert_eq!(item.label.as_deref(), Some("Paris"));
    let FormNode::Control(option) = &controls[5].children[0] else {
        panic!("listbox options are nested controls")
    };
    assert_eq!(option.kind, FormControlKind::Option);
    assert_eq!(option.current_selected, Some(true));
    assert_eq!(option.value.as_deref(), Some("r"));

    // The event declaration is retained as inert metadata, never executed.
    assert!(forms.has_event_listeners);
    let listener = forms
        .event_listeners
        .iter()
        .find(|listener| listener.macro_name.is_some())
        .expect("checkbox event listener is preserved");
    assert_eq!(listener.macro_name.as_deref(), Some("macro:///never.open"));
    assert_eq!(listener.event_name.as_deref(), Some("form:performaction"));

    // The paragraph anchor resolves to the declared control without fetching.
    assert!(forms.control_shapes[0].resolved_control.is_some());

    // The dedicated flat facade retains this inventory without exposing a
    // mutable conversion; authored mutation is covered by the packaged tests
    // below and never executes the retained event metadata.
    assert!(document.forms().is_ok());
}

#[test]
fn builder_authored_forms_round_trip_the_package() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Registration form").unwrap();
    builder.add_control_form(&text_form()).unwrap();
    builder.add_interactive_form(&interactive_form()).unwrap();
    builder.add_selection_form(&selection_form()).unwrap();
    builder.add_typed_value_form(&typed_value_form()).unwrap();
    builder.add_generic_form(&generic_form()).unwrap();
    builder.add_visual_form(&visual_form()).unwrap();
    // Duplicate form names are rejected atomically.
    assert!(builder.add_control_form(&text_form()).is_err());

    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    let forms = document.forms().unwrap();
    assert_eq!(forms.groups.len(), 1);
    let names: Vec<Option<String>> = forms.groups[0]
        .forms
        .iter()
        .map(|form| form.name.clone())
        .collect();
    // The builder emits one `<form:form>` per control family in family order.
    assert_eq!(
        names,
        ["texts", "buttons", "choices", "radios", "labels", "values"]
            .map(|name| Some(name.to_string()))
    );

    let mutable = MutableDocument::from_document(document).unwrap();
    assert_eq!(mutable.text_controls().unwrap(), text_form().controls);
    assert_eq!(
        mutable.interactive_controls().unwrap(),
        interactive_form().controls
    );
    assert_eq!(
        mutable.selection_controls().unwrap(),
        selection_form().controls
    );
    assert_eq!(
        mutable.typed_value_controls().unwrap(),
        typed_value_form().controls
    );
    assert_eq!(
        mutable.generic_form_controls().unwrap(),
        generic_form().controls
    );
    assert_eq!(mutable.visual_controls().unwrap(), visual_form().controls);
}

#[test]
fn mutable_inserts_replaces_and_removes_controls() {
    // The typed mutation APIs only operate on macro-free documents, so the
    // round trip starts from a builder-authored package.
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Registration form").unwrap();
    builder.add_control_form(&text_form()).unwrap();
    builder.add_interactive_form(&interactive_form()).unwrap();
    builder.add_selection_form(&selection_form()).unwrap();
    builder.add_typed_value_form(&typed_value_form()).unwrap();
    builder.add_generic_form(&generic_form()).unwrap();
    builder.add_visual_form(&visual_form()).unwrap();
    let document = Document::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutableDocument::from_document(document).unwrap();

    // Insert a new textarea into the first form (document order).
    let extra = TextControl::textarea("Notes", "notes_area");
    mutable.insert_text_control(0, &extra).unwrap();
    // Replace the name field's current value.
    let mut renamed = TextControl::text("Name", "name_field");
    renamed.current_value = Some("Grace".to_string());
    renamed.title = Some("Your name".to_string());
    renamed.max_length = Some(40);
    let old = mutable.replace_text_control(0, &renamed).unwrap();
    assert_eq!(old.current_value.as_deref(), Some("Ada"));
    // Remove the bio textarea and the hidden token.
    let removed = mutable.remove_text_control(1).unwrap();
    assert_eq!(removed.xml_id, "bio_field");
    let removed_hidden = mutable.remove_generic_form_control(1).unwrap();
    assert_eq!(removed_hidden.xml_id(), "hidden_token");
    // Insert a second radio option into the radio form (document order index 3)
    // and retitle the date control.
    let mut choice_b = RadioControl::new("ChoiceB", "choice_b");
    choice_b.label = Some("Option B".to_string());
    choice_b.value = Some("b".to_string());
    mutable.insert_visual_control(3, &choice_b.into()).unwrap();
    let mut when = TypedValueControl::new(TypedValueControlKind::Date, "When", "when_date");
    when.current_value = Some("2024-02-29".to_string());
    when.title = Some("Pick a day".to_string());
    mutable.replace_typed_value_control(1, &when).unwrap();
    // Out-of-bounds edits fail without changing the document.
    assert!(mutable.remove_text_control(42).is_err());
    assert!(mutable.insert_text_control(42, &extra).is_err());

    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let reopened = MutableDocument::from_document(reopened).unwrap();

    let texts = reopened.text_controls().unwrap();
    assert_eq!(texts.len(), 2);
    assert_eq!(texts[0].current_value.as_deref(), Some("Grace"));
    assert_eq!(texts[1].xml_id, "notes_area");

    assert_eq!(reopened.generic_form_controls().unwrap().len(), 1);
    let radios = reopened.visual_controls().unwrap();
    assert_eq!(radios.len(), 2);
    let values = reopened.typed_value_controls().unwrap();
    assert_eq!(values[1].current_value.as_deref(), Some("2024-02-29"));
    assert_eq!(values[1].title.as_deref(), Some("Pick a day"));

    // The untouched controls survive the round trip unchanged.
    assert_eq!(
        reopened.interactive_controls().unwrap(),
        interactive_form().controls
    );
    assert_eq!(
        reopened.selection_controls().unwrap(),
        selection_form().controls
    );
}

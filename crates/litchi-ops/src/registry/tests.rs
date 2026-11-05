use super::*;
use crate::node::spec::Subject;
use crate::value::{LengthSpec, ValueType};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Setter {
    Bold,
    Size,
    Name,
}

const GAP: GapSpec = GapSpec {
    subject: Subject::Property {
        element: "run",
        key: "fill",
    },
    code: Code::UnsupportedCapability,
    follow_up: "M2",
    reason: "the leaf setter is pending",
};
const TABLE: Table<Setter, 5, 2> = Table::new(
    [
        Entry {
            spec: PropSpec::new("fontSize", ValueType::Length(LengthSpec::FONT_SIZE)),
            access: Access::SetAndAdd(Setter::Size),
            constraint: Constraint::Length {
                min: Length::from_points(1),
                max: Length::from_points(4000),
            },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("bold", ValueType::Bool),
            access: Access::Set(Setter::Bold),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("name", ValueType::Text),
            access: Access::Add(Setter::Name),
            constraint: Constraint::Chars { max: 3 },
            clear: false,
        },
        Entry::read(PropSpec::new("id", ValueType::Text)),
        Entry {
            spec: PropSpec::new("fill", ValueType::Text),
            access: Access::Gap(&GAP),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [
        Alias {
            name: "size",
            entry: 0,
        },
        Alias {
            name: "font.bold",
            entry: 1,
        },
    ],
);

#[test]
fn lookup_is_folded_and_preserves_declaration_indices() {
    assert_eq!(TABLE.lookup("FONTSIZE").unwrap().0, 0);
    assert_eq!(TABLE.lookup("BoLd").unwrap().0, 1);
    assert_eq!(TABLE.lookup("SIZE").unwrap().0, 0);
    assert_eq!(TABLE.lookup("FONT.BOLD").unwrap().0, 1);
    assert!(TABLE.lookup(" fontSize").is_none());
    assert!(TABLE.lookup("böld").is_none());
    let props = TABLE.readable();
    assert_eq!(props[0].name, "fontSize");
    assert!(props[0].access.set && props[0].access.add);
    assert!(props[1].access.set && !props[1].access.add);
    assert!(!props[2].access.set && props[2].access.add);
    assert!(!props[3].access.set && !props[4].access.add);
}

#[test]
fn access_coercion_and_constraints_run_before_staging() {
    let (_, size) = TABLE.resolve("size").unwrap();
    let result = size
        .prepare(Operation::Set, &PropValue::from(12), |setter, value| {
            assert_eq!(setter, Setter::Size);
            assert!(
                matches!(value, Some(Scalar::Length(value)) if value == Length::from_points(12))
            );
            Ok(42)
        })
        .unwrap();
    assert_eq!(result, 42);
    for value in [
        PropValue::from("-1pt"),
        PropValue::from("4001pt"),
        PropValue::from("secret-invalid-value"),
    ] {
        let error = size
            .prepare(Operation::Set, &value, |_, _| -> Result<()> {
                panic!("invalid values must not stage")
            })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some("fontSize"));
        assert!(!format!("{error:?}").contains("secret-invalid-value"));
    }
    size.prepare(Operation::Set, &PropValue::Null, |_, value| {
        assert!(value.is_none());
        Ok(())
    })
    .unwrap();
    let (_, bold) = TABLE.resolve("bold").unwrap();
    assert_eq!(
        bold.prepare(Operation::Add, &true.into(), |_, _| Ok(()))
            .unwrap_err()
            .code(),
        Code::UnsupportedProperty
    );
    assert_eq!(
        bold.prepare(Operation::Set, &PropValue::Null, |_, _| Ok(()))
            .unwrap_err()
            .code(),
        Code::InvalidValue
    );
    let (_, name) = TABLE.resolve("name").unwrap();
    name.prepare(Operation::Add, &"你好🙂".into(), |_, _| Ok(()))
        .unwrap();
    assert_eq!(
        name.prepare(Operation::Add, &"你好🙂!".into(), |_, _| Ok(()))
            .unwrap_err()
            .code(),
        Code::InvalidValue
    );
}

#[test]
fn errors_distinguish_unknown_read_only_and_leaf_gaps() {
    let unknown = TABLE.resolve("bolld").err().unwrap();
    assert_eq!(unknown.code(), Code::UnsupportedProperty);
    assert_eq!(unknown.suggestion(), Some("did you mean 'bold'?"));
    let (_, read_only) = TABLE.resolve("id").unwrap();
    assert_eq!(
        read_only
            .prepare(Operation::Set, &"private".into(), |_, _| Ok(()))
            .unwrap_err()
            .property(),
        Some("id")
    );
    let (_, gap) = TABLE.resolve("fill").unwrap();
    let error = gap
        .prepare(Operation::Set, &"private".into(), |_, _| Ok(()))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.suggestion(), Some("M2"));
}

#[test]
fn malformed_static_declarations_are_rejected() {
    let read = |name| Entry::<Setter>::read(PropSpec::new(name, ValueType::Text));
    assert!(std::panic::catch_unwind(|| Table::new([read("name"), read("NAME")], [])).is_err());
    assert!(
        std::panic::catch_unwind(|| Table::new(
            [read("name")],
            [Alias {
                name: "NAME",
                entry: 0
            }]
        ))
        .is_err()
    );
    assert!(
        std::panic::catch_unwind(|| Table::new(
            [read("name")],
            [Alias {
                name: "label",
                entry: 1
            }]
        ))
        .is_err()
    );
}

#[test]
fn all_constraint_domains_refuse_instead_of_clamping() {
    let invalid = Entry {
        spec: PropSpec::new("count", ValueType::Bool),
        access: Access::Set(Setter::Bold),
        constraint: Constraint::Int { min: 1, max: 3 },
        clear: false,
    };
    assert!(std::panic::catch_unwind(|| Table::new([invalid], [])).is_err());

    assert!(
        Constraint::Int { min: 1, max: 3 }
            .check(Scalar::Int(3))
            .is_ok()
    );
    assert!(
        Constraint::Int { min: 1, max: 3 }
            .check(Scalar::Int(4))
            .is_err()
    );
    assert!(
        Constraint::Number { min: 0.0, max: 1.0 }
            .check(Scalar::Number(f64::NAN))
            .is_err()
    );
    assert!(
        Constraint::Number { min: 0.0, max: 1.0 }
            .check(Scalar::Number(0.5))
            .is_ok()
    );
    assert!(Constraint::NotEmpty.check(Scalar::Text("")).is_err());
    assert!(Constraint::NotEmpty.check(Scalar::Text(" ")).is_ok());
    assert!(Constraint::NotEmpty.check(Scalar::Int(1)).is_err());
}

crate::registry::registry! {
    READ_PROPS, READ_TABLE: Setter {
        Entry::read(PropSpec::new("second", ValueType::Bool)),
        Entry::read(PropSpec::new("first", ValueType::Text)),
    }
}

#[test]
fn registry_macro_keeps_read_projection_and_lookup_together() {
    assert_eq!(READ_PROPS[0].name, "second");
    assert_eq!(READ_TABLE.lookup("FIRST").unwrap().0, 1);
    assert!(!READ_PROPS[0].access.set);
}

//! OpenDocument database front ends (`.odb`): tables, their columns and
//! queries, read from the `litchi_odb::Database` catalog and written through
//! `litchi_odb::Edit`.

use litchi_core::FileFormat;
use litchi_core::selector::Path;

use super::tree::{count, text};
use super::write::{
    Commit, Patch, added_path, bound, element_refusal, finish, key_refusal, refusal, required,
    resolve, staging_error,
};
use super::{Flavor, OdfTree, Owner, TreeBuilder, addable, element, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::ValueType;

kinds! {
    /// The element kinds of an ODB database.
    Kind: &crate::path::vocabulary::DATABASE, root Database {
        Database => "database",
        Table => "table",
        Column => "column",
        Query => "query",
    }
}

const DATABASE_PROPS: &[PropSpec] = &[
    PropSpec::new("tableCount", ValueType::Int),
    PropSpec::new("queryCount", ValueType::Int),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
];

const TABLE_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    PropSpec::new("columnCount", ValueType::Int),
];

const COLUMN_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    PropSpec::new("typeName", ValueType::Text),
];

const QUERY_PROPS: &[PropSpec] = &[
    addable("name", ValueType::Text),
    addable("command", ValueType::Text),
    PropSpec::new("columnCount", ValueType::Int),
];

static OPERATION_GAPS: [GapSpec; 6] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "table",
            context: "a table definition (columns, keys and types)",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODB typed table creation",
        reason: "creating a table needs a typed column schema the unified registry does not have yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "column",
            context: "a column definition",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODB typed column creation",
        reason: "creating a column needs a typed data type the unified registry does not have yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "query",
            context: "the `name` property (renaming a query)",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODB query rename",
        reason: "the leaf has no query rename; remove and add the query instead",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "column",
            context: "a column into another table",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0292 ODB cross-table column transfer",
        reason: "a column belongs to its table's schema; columns move and swap within their table",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "table",
            context: "a database that mixes table representations and definitions",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0292 ODB mixed table collections",
        reason: "representations and definitions order in separate collections, so one unified table order cannot be published",
    },
    super::METADATA_PATH_GAP,
];

/// The ODB spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Database,
    elements: &[
        element(
            Kind::Database,
            "database",
            &["database"],
            &[],
            &[Kind::Table, Kind::Query],
            &["/"],
            DATABASE_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table"],
            &[Kind::Database],
            &[Kind::Column],
            &["/table[N]", "/table[NAME]"],
            TABLE_PROPS,
        ),
        element(
            Kind::Column,
            "column",
            &["column"],
            &[Kind::Table],
            &[],
            &["/table[N]/column[N]", "/table[N]/column[NAME]"],
            COLUMN_PROPS,
        ),
        element(
            Kind::Query,
            "query",
            &["query"],
            &[Kind::Database],
            &[],
            &["/query[N]", "/query[NAME]"],
            QUERY_PROPS,
        ),
    ],
    gaps: &[],
};

crate::node::spec::check_format_spec!(SPEC);

/// The ODB flavor.
pub(crate) struct Odb;

impl Flavor for Odb {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Odb;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Database, Kind::Table];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["table", "column", "query"],
    addable: &["query"],
    removable: &["table", "column", "query"],
    gaps: &OPERATION_GAPS,
};

/// Build the tree of `package`.
pub(crate) fn build(package: &litchi_odb::Package) -> Result<super::Tree> {
    let database = package.database();
    let catalog = database
        .catalog()
        .map_err(|error| super::odf_error(error, FileFormat::Odb))?;
    let mut tree = TreeBuilder::new(SPEC.elements, &[]);
    let metadata = database.metadata();
    tree.root()
        .prop("tableCount", count(catalog.tables().len()))
        .prop("queryCount", count(catalog.queries().len()))
        .prop(
            "title",
            text(metadata.and_then(|metadata| metadata.title.as_deref())),
        )
        .prop(
            "author",
            text(metadata.and_then(|metadata| metadata.author.as_deref())),
        );
    for (table_index, table) in catalog.tables().iter().enumerate() {
        let parent = tree
            .push(super::tree::ROOT, Kind::Table)
            .prop("name", text(Some(table.name())))
            .prop("columnCount", count(table.columns().len()))
            .label(Some(table.name()))
            .leaf(table_index, 0)
            .done();
        for (column_index, column) in table.columns().iter().enumerate() {
            tree.push(parent, Kind::Column)
                .prop("name", text(Some(column.name())))
                .prop("typeName", text(column.type_name()))
                .label(Some(column.name()))
                .leaf(table_index, column_index)
                .done();
        }
    }
    for (index, query) in catalog.queries().iter().enumerate() {
        tree.push(super::tree::ROOT, Kind::Query)
            .prop("name", text(Some(query.name())))
            .prop("command", text(Some(query.command())))
            .prop("columnCount", count(query.columns().len()))
            .text(query.command())
            .label(Some(query.name()))
            .leaf(index, 0)
            .done();
    }
    Ok(tree.build())
}

/// The name a table, column or query node carries.
fn own_name(node: &super::Node<'_>) -> String {
    node.item().label().unwrap_or_default().to_owned()
}

/// The table name of a column node: its parent's label.
fn table_of(file: &OdfTree, node: &super::Node<'_>) -> String {
    let leaf = node.item().leaf()[0];
    file.tree()
        .items()
        .iter()
        .find(|item| item.type_name() == "table" && item.leaf()[0] == leaf)
        .and_then(|item| item.label())
        .unwrap_or_default()
        .to_owned()
}

/// Set table, column or query properties.
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odb::Package,
    groups: &[super::write::Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut edit = package.edit();
    let mut operations = 0usize;
    let staging = |error| staging_error(error, FileFormat::Odb);
    for (node, props) in groups
        .iter()
        .flat_map(|(nodes, props)| nodes.iter().map(move |node| (node, *props)))
    {
        let kind = node.item().type_name();
        let own = own_name(node);
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            match (kind, key) {
                ("table", "name") => {
                    let name = required(super::write::text(value, "name")?, "name")?;
                    if name != own {
                        edit.rename_table(&own, &name).map_err(staging)?;
                    }
                },
                ("column", "name") => {
                    let name = required(super::write::text(value, "name")?, "name")?;
                    if name != own {
                        let table = table_of(file, node);
                        edit.rename_column(&table, &own, &name).map_err(staging)?;
                    }
                },
                ("query", "command") => {
                    let command = required(super::write::text(value, "command")?, "command")?;
                    let current = node.item().text().unwrap_or_default();
                    if command != current {
                        edit.set_query_command(&own, command).map_err(staging)?;
                    }
                },
                ("query", "name") => {
                    return Err(refusal(
                        Code::UnsupportedCapability,
                        "the ODB leaf has no query rename; remove and add the query instead",
                    )
                    .with_property("name".to_owned()));
                },
                ("table" | "column" | "query", other) => return Err(key_refusal(node, other)),
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    commit(file, edit, operations, limits)
}

fn commit(
    file: &OdfTree,
    edit: litchi_odb::Edit<'_>,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let commit = edit
        .commit()
        .map_err(|error| super::write::commit_error(error, FileFormat::Odb))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    finish(
        file,
        changed,
        Owner::Odb(package),
        Patch::Odb(patch),
        operations,
        limits,
    )
}

/// Add a query under the database root.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odb::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let parent_node = super::write::parent::<Odb>(file, parent)?;
    if !node.kind().eq_ignore_ascii_case("query") {
        return Err(refusal(
            Code::UnsupportedCapability,
            format!(
                "adding `{}` elements to an ODB database is not supported",
                node.kind()
            ),
        )
        .with_suggestion(
            "add a query: NewNode::new(\"query\").prop(\"name\", ..).prop(\"command\", ..)",
        ));
    }
    if parent_node.item().type_name() != "database" {
        return Err(refusal(
            Code::InvalidPath,
            "queries are added under the database root `/`",
        ));
    }
    if node
        .placement()
        .is_some_and(|placement| !matches!(placement, crate::Placement::End))
    {
        return Err(refusal(
            Code::UnsupportedCapability,
            "ODB queries are appended; a placement is not supported",
        ));
    }
    let mut name = None;
    let mut command = None;
    for (key, value) in node.props().iter() {
        match key.as_str() {
            "name" => name = super::write::text(value, "name")?,
            "command" => command = super::write::text(value, "command")?,
            other => {
                return Err(refusal(
                    Code::UnsupportedProperty,
                    format!("`{other}` is not a property `add` accepts for an ODB query"),
                )
                .with_property(other.to_owned()));
            },
        }
    }
    let name = required(name, "name")?;
    let command = required(command, "command")?;
    let mut edit = package.edit();
    edit.add_query(litchi_odb::Query::new(name.as_str(), command))
        .map_err(|error| staging_error(error, FileFormat::Odb))?;
    let commit = commit(file, edit, 1, limits)?;
    let path = added_path(&commit.after, &[], "query", None)?;
    Ok((commit, path))
}

/// Remove tables, columns or queries.
pub(crate) fn remove(
    file: &OdfTree,
    package: &litchi_odb::Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Odb>(file, target, limits.max_expanded_targets())?;
    let mut edit = package.edit();
    let staging = |error| staging_error(error, FileFormat::Odb);
    for node in &nodes {
        let own = own_name(node);
        match node.item().type_name() {
            "table" => edit.remove_table(&own).map_err(staging)?,
            "column" => {
                let table = table_of(file, node);
                edit.remove_column(&table, &own).map_err(staging)?;
            },
            "query" => edit.remove_query(&own).map_err(staging)?,
            _ => return Err(element_refusal("remove", node)),
        }
    }
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    commit(file, edit, nodes.len(), limits)
}

/// The leaf position space of a table move: refused when the catalog mixes
/// table representations and definitions, whose collections order apart.
fn single_table_kind(package: &litchi_odb::Package) -> Result<()> {
    let catalog = package
        .database()
        .catalog()
        .map_err(|error| super::odf_error(error, FileFormat::Odb))?;
    let mut kinds = catalog.tables().iter().map(litchi_odb::Table::kind);
    let first = kinds.next();
    if kinds.any(|kind| Some(kind) != first) {
        return Err(refusal(
            Code::UnsupportedCapability,
            "the database mixes table representations and definitions, which order in separate collections",
        )
        .with_suggestion("follow-up: 0292 ODB mixed table collections"));
    }
    Ok(())
}

/// Stage moving `node` (a table, column or query) to final position `to`
/// among its siblings.
fn stage_move(
    file: &OdfTree,
    edit: &mut litchi_odb::Edit<'_>,
    node: &super::Node<'_>,
    to: usize,
) -> Result<()> {
    let own = own_name(node);
    let staging = |error| staging_error(error, FileFormat::Odb);
    match node.item().type_name() {
        "table" => edit.move_table(&own, to).map_err(staging),
        "column" => edit
            .move_column(&table_of(file, node), &own, to)
            .map_err(staging),
        "query" => edit.move_query(&own, to).map_err(staging),
        _ => Err(element_refusal("move", node)),
    }
}

/// Move a table, a column within its table, or a query (record 0292).
pub(crate) fn move_to(
    file: &OdfTree,
    package: &litchi_odb::Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = super::write::parent::<Odb>(file, path)?;
    if node.item().type_name() == "database" {
        return Err(element_refusal("move", &node));
    }
    let siblings = super::write::siblings(file, &node)?;
    let super::write::Destination::Within(to) =
        super::write::destination::<Odb>(file, &node, &siblings, placement)?
    else {
        return Err(refusal(
            Code::UnsupportedCapability,
            "an ODB column moves within its own table",
        )
        .with_suggestion("follow-up: 0292 ODB cross-table column transfer"));
    };
    if node.item().type_name() == "table" {
        single_table_kind(package)?;
    }
    let mut edit = package.edit();
    stage_move(file, &mut edit, &node, to)?;
    commit(file, edit, 1, limits)
}

/// Exchange two tables, two columns of one table, or two queries.
pub(crate) fn swap(
    file: &OdfTree,
    package: &litchi_odb::Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let (left, right) = super::write::swap_pair::<Odb>(file, left, right)?;
    if left.item().type_name() == "database" {
        return Err(element_refusal("swap", &left));
    }
    let left_siblings = super::write::siblings(file, &left)?;
    let right_siblings = super::write::siblings(file, &right)?;
    if left_siblings.parent != right_siblings.parent {
        return Err(refusal(
            Code::UnsupportedCapability,
            "ODB columns swap within one table",
        )
        .with_suggestion("follow-up: 0292 ODB cross-table column transfer"));
    }
    if left.item().type_name() == "table" {
        single_table_kind(package)?;
    }
    let (first, second, i, j) = if left_siblings.from <= right_siblings.from {
        (left, right, left_siblings.from, right_siblings.from)
    } else {
        (right, left, right_siblings.from, left_siblings.from)
    };
    let mut edit = package.edit();
    if i != j {
        stage_move(file, &mut edit, &first, j)?;
        stage_move(file, &mut edit, &second, i)?;
    }
    commit(file, edit, 2, limits)
}

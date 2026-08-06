use super::super::codec::invalid;
use super::super::package::*;
use super::super::*;
use super::*;
pub(in crate::web) fn canonicalize_pane_snapshot_resources(
    pane: &mut Pane,
    existing_panes: &[Pane],
    skip_index: Option<usize>,
) -> Result<()> {
    for index in 0..pane.snapshot_resources.len() {
        let (previous, incoming) = pane.snapshot_resources.split_at_mut(index);
        let incoming = &mut incoming[0];
        for existing in previous {
            reconcile_snapshot_resource(incoming, existing, "another resource in the same pane")?;
        }
    }

    for incoming in &mut pane.snapshot_resources {
        for (pane_index, existing_pane) in existing_panes.iter().enumerate() {
            if skip_index == Some(pane_index) {
                continue;
            }
            for existing in &existing_pane.snapshot_resources {
                reconcile_snapshot_resource(incoming, existing, "an existing pane resource")?;
            }
        }
    }
    Ok(())
}

pub(in crate::web) fn reconcile_snapshot_resource(
    incoming: &mut SnapshotResource,
    existing: &SnapshotResource,
    conflict_scope: &str,
) -> Result<()> {
    let (
        SnapshotTarget::Internal {
            part_name,
            content_type,
            data,
        },
        SnapshotTarget::Internal {
            part_name: existing_name,
            content_type: existing_content_type,
            data: existing_data,
        },
    ) = (&mut incoming.target, &existing.target)
    else {
        return Ok(());
    };
    if !part_names_conflict(part_name, existing_name) {
        return Ok(());
    }
    if part_name.is_equivalent_to(existing_name)
        && content_type == existing_content_type
        && data.as_slice() == existing_data.as_slice()
    {
        *part_name = existing_name.clone();
        return Ok(());
    }
    invalid(format!(
        "snapshot part '{}' conflicts with {conflict_scope}",
        part_name.as_str()
    ))
}

//! The removal plan the ported CalculationEngine phase ([`formula`](crate::structure::formula))
//! reads and extends, in the shapes of the Pages body-table deletion owner.

use std::collections::HashMap;
use std::num::NonZeroU64;

use super::{Fault, FaultResult, allocation};
use crate::TableSource;

/// Where one source object lives.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DeletionObjectLocation {
    pub component_index: usize,
    pub object_index: usize,
}

/// The source-wide object index.
#[derive(Debug, Default)]
pub struct DeletionObjectIndex {
    by_identifier: HashMap<NonZeroU64, DeletionObjectLocation>,
}

impl DeletionObjectIndex {
    /// Index every object of `package`; an identifier stored twice is
    /// refused.
    pub fn new(package: &dyn TableSource) -> FaultResult<Self> {
        let catalog = package.catalog();
        let mut by_identifier = HashMap::new();
        for (component_index, component) in catalog.iter().enumerate() {
            by_identifier
                .try_reserve(component.archive().objects.len())
                .map_err(|_error| allocation(component.archive().objects.len()))?;
            for (object_index, object) in component.archive().objects.iter().enumerate() {
                let identifier = object
                    .archive_info
                    .identifier
                    .and_then(NonZeroU64::new)
                    .ok_or(Fault::Invalid)?;
                if by_identifier
                    .insert(
                        identifier,
                        DeletionObjectLocation {
                            component_index,
                            object_index,
                        },
                    )
                    .is_some()
                {
                    return Err(Fault::Unsupported);
                }
            }
        }
        Ok(Self { by_identifier })
    }

    pub fn get(&self, identifier: NonZeroU64) -> Option<&DeletionObjectLocation> {
        self.by_identifier.get(&identifier)
    }
}

/// One payload replacement and the header references it prunes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MessageEdit {
    pub component_index: usize,
    pub object_index: usize,
    pub message_index: usize,
    pub object_identifier: NonZeroU64,
    pub message_type: u32,
    pub data: Vec<u8>,
    pub remove_object_references: Vec<NonZeroU64>,
    pub remove_data_references: Vec<NonZeroU64>,
}

/// One object selected for removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ObjectRemoval {
    pub component_index: usize,
    pub object_index: usize,
    pub identifier: NonZeroU64,
}

/// The message edits and object removals one phase proved.
#[derive(Debug, Clone, Default)]
pub struct RemovalPlan {
    pub message_edits: Vec<MessageEdit>,
    pub object_removals: Vec<ObjectRemoval>,
    pub formula_contexts: Vec<NonZeroU64>,
}

/// The graph phase's output the formula phase reads.
#[derive(Debug)]
pub struct GraphPlan {
    pub index: DeletionObjectIndex,
    pub removals: RemovalPlan,
}

/// The formula phase's output.
#[derive(Debug)]
pub struct FormulaPlan {
    pub removals: RemovalPlan,
}

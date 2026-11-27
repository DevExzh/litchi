//! The package metadata member (`Index/Metadata.iwa`) route.

use crate::TableSource;

/// The package metadata member every iWork package stores.
pub const ENTRY_NAME: &str = "Index/Metadata.iwa";
pub(crate) const MESSAGE_TYPE: u32 = 11_006;

/// The native route of the unique package metadata message.
#[derive(Debug, Clone, Copy)]
pub struct MessageRoute {
    /// Component index in the catalog.
    pub component_index: usize,
    /// Object index inside the component.
    pub object_index: usize,
    /// Message index inside the object.
    pub message_index: usize,
}

pub(crate) fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

pub(crate) fn unique_message_route(source: &dyn TableSource) -> Option<MessageRoute> {
    let mut route = None;
    for (component_index, component) in source.catalog().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != MESSAGE_TYPE {
                    continue;
                }
                if route.is_some() || component.name() != ENTRY_NAME {
                    return None;
                }
                route = Some(MessageRoute {
                    component_index,
                    object_index,
                    message_index,
                });
            }
        }
    }
    route
}

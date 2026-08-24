use super::Package;

pub(super) const ENTRY_NAME: &str = "Index/Metadata.iwa";
pub(super) const MESSAGE_TYPE: u32 = 11_006;

#[derive(Debug, Clone, Copy)]
pub(super) struct MessageRoute {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
}

pub(super) fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

pub(super) fn unique_message_route(source: &Package) -> Option<MessageRoute> {
    let mut route = None;
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
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

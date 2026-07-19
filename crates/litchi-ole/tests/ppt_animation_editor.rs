use litchi_ole::ppt::animation::{ExtendedTimeNode, TimeNodeAtom, TimeNodeKind};
use litchi_ole::ppt::writer::{PersistPtrBuilder, UserEditAtom};
use litchi_ole::ppt::{PowerPointAnimationEditor, PowerPointAnimationEditorLimits, PowerPointAnimationScope};
use litchi_ole::OleWriter;
use std::io::Cursor;

fn record(version:u16,kind:u16,payload:&[u8])->Vec<u8>{let mut out=version.to_le_bytes().to_vec();out.extend_from_slice(&kind.to_le_bytes());out.extend_from_slice(&(payload.len()as u32).to_le_bytes());out.extend_from_slice(payload);out}

fn generated_ppt()->Vec<u8>{
    let document=record(0x0f,1000,&[]);
    let slide=record(0x0f,1006,&[]);
    let master=record(0x0f,1016,&[]);
    let mut stream=document.clone();let slide_offset=stream.len()as u32;stream.extend(slide);
    let master_offset=stream.len()as u32;stream.extend(master);
    let mut ptr=PersistPtrBuilder::new();ptr.set_offset(1,0);ptr.set_offset(2,slide_offset);ptr.set_offset(3,master_offset);
    let dir_offset=stream.len()as u32;stream.extend(ptr.generate_full_record());
    let edit_offset=stream.len()as u32;stream.extend(UserEditAtom::new_minimal(dir_offset,1,3,0).generate_record());
    let mut current=vec![0;28];current[12..16].copy_from_slice(&0xE391_C05Fu32.to_le_bytes());current[16..20].copy_from_slice(&edit_offset.to_le_bytes());
    let mut writer=OleWriter::new();writer.create_stream(&["PowerPoint Document"],&stream).unwrap();writer.create_stream(&["Current User"],&current).unwrap();
    let mut output=Cursor::new(Vec::new());writer.write_to(&mut output).unwrap();output.into_inner()
}

#[test]
fn generated_timeline_add_update_reorder_remove_and_reopen() {
    let mut editor=PowerPointAnimationEditor::open(generated_ppt(),PowerPointAnimationEditorLimits::default()).unwrap();
    assert_eq!(editor.find(2).unwrap().scope,PowerPointAnimationScope::Slide);
    assert_eq!(editor.find(3).unwrap().scope,PowerPointAnimationScope::MainMaster);
    let first=ExtendedTimeNode{atom:TimeNodeAtom{node_type:Some(TimeNodeKind::Parallel),duration_ms:Some(100),..Default::default()},..Default::default()};
    let second=ExtendedTimeNode{atom:TimeNodeAtom{node_type:Some(TimeNodeKind::Sequential),duration_ms:Some(200),..Default::default()},..Default::default()};
    editor.add(2,0,first).unwrap();editor.add(2,1,second.clone()).unwrap();editor.reorder(2,&[1,0]).unwrap();
    editor.add(3,0,ExtendedTimeNode::default()).unwrap();
    editor.update(2,0,second).unwrap();editor.remove(2,1).unwrap();
    let bytes=editor.finish().unwrap();let reopened=PowerPointAnimationEditor::open(bytes,PowerPointAnimationEditorLimits::default()).unwrap();
    assert_eq!(reopened.find(2).unwrap().extension.time_node.unwrap().children.len(),1);
    assert_eq!(reopened.find(3).unwrap().extension.time_node.unwrap().children.len(),1);
}

#[test]
fn invalid_indexes_and_limits_are_atomic() {
    let mut editor=PowerPointAnimationEditor::open(generated_ppt(),PowerPointAnimationEditorLimits::default()).unwrap();
    let before=editor.timelines();assert!(editor.add(2,1,ExtendedTimeNode::default()).is_err());assert!(editor.reorder(2,&[0]).is_err());assert_eq!(editor.timelines(),before);
}

#[test]
fn poi_and_libreoffice_animation_fixtures_are_strictly_gated() {
    let root=std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in [root.join("3rdparty/poi/test-data/slideshow/sound.ppt"),root.join("3rdparty/poi/test-data/slideshow/datetime.ppt")]{let original=std::fs::read(&path).unwrap();match PowerPointAnimationEditor::open(original.clone(),PowerPointAnimationEditorLimits::default()){Ok(editor)=>{let _=editor.timelines();assert_eq!(std::fs::read(&path).unwrap(),original)},Err(_)=>assert_eq!(std::fs::read(&path).unwrap(),original)}}
}

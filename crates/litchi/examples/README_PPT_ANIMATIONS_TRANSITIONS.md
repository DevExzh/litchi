# PPT Animation and Transition Examples

These examples demonstrate the newly implemented animation and transition features for PowerPoint (.ppt) files.

## Examples Overview

### 1. `ppt_animations_example.rs`
Showcases various animation effects organized by category:

**Slide 1: Entrance Animations**
- Fade In
- Fly In (from left)
- Zoom
- Wipe (from bottom)

**Slide 2: Emphasis Animations**
- Pulse
- Spin
- Teeter
- Wave

**Slide 3: Exit Animations**
- Dissolve
- Swivel
- Bounce
- Split (from left)

**Slide 4: Speed Variations**
- Demonstrates all five speed levels: Very Slow (5s), Slow (3s), Medium (2s), Fast (1s), Very Fast (0.5s)

**Output:** `ppt_animations_showcase.ppt`

### 2. `ppt_transitions_example.rs`
Demonstrates slide transition effects:

**Transitions Showcased:**
- Classic: Fade, Dissolve, Wipe
- Directional: Push (left), Wipe (bottom), Split (horizontal)
- Pattern: Blinds (vertical), Checkerboard, Box (in)
- Modern: Zoom, Random, Morph

**Advance Modes:**
- On Click (manual advance)
- Automatic (with timing: 3s, 5s)
- Both (click or wait)

**Output:** `ppt_transitions_showcase.ppt`

### 3. `ppt_animations_transitions_combined.rs`
A comprehensive example combining animations and transitions:

**Slide 1:** Title with Zoom animation + Fade transition
**Slide 2:** Sequential build effects with FlyIn animations + Push transition
**Slide 3:** Emphasis effects (Pulse, Spin, GrowAndTurn, Bounce) + Wipe transition
**Slide 4:** Exit animations + Split transition
**Slide 5:** Auto-advance demonstration (4 seconds) + Dissolve transition
**Slide 6:** Conclusion with Zoom animation + Random transition

**Output:** `ppt_complete_showcase.ppt`

## Running the Examples

```bash
# Animation showcase
cargo run --example ppt_animations_example

# Transition showcase
cargo run --example ppt_transitions_example

# Combined demonstration
cargo run --example ppt_animations_transitions_combined
```

## Output Files

Each example generates:
1. **Main PPT file**: The complete presentation to open in PowerPoint
2. **Binary samples**: Raw animation/transition data for debugging (`.bin` files)

## Verification in PowerPoint

To verify the features are working correctly:

1. Open the generated `.ppt` file in **Microsoft PowerPoint**
2. Enter **Slide Show mode** (F5 or Slide Show → From Beginning)
3. Click through slides to see animations trigger
4. Observe slide transitions between slides
5. For the auto-advance demo, watch the slide automatically progress

## Animation Features Implemented

### Animation Effects (18 types)
- **Entrance:** Appear, FadeIn, FlyIn, Zoom, Wipe, Split, Dissolve, Box, Checkerboard, Blinds, RandomBars, GrowAndTurn, Swivel
- **Emphasis:** Pulse, Spin, Teeter, Wave, Bounce
- **Exit:** All entrance effects can be used as exits

### Animation Properties
- **Build Types:** Entrance, Emphasis, Exit, MotionPath
- **Speed:** Very Slow (5s), Slow (3s), Medium (2s), Fast (1s), Very Fast (0.5s)
- **Direction:** None, Horizontal, Vertical, FromTop, FromBottom, FromLeft, FromRight, FromTopLeft, FromTopRight, FromBottomLeft, FromBottomRight, In, Out
- **Trigger:** OnClick, WithPrevious, AfterPrevious

## Transition Features Implemented

### Transition Types (46 types)
Classic, modern, and directional transitions including:
- Fade, Dissolve, Wipe, Push, Split, Cover, Uncover
- Blinds, Checkerboard, Box, Comb, Wheel, Wedge
- Zoom, Random, Newsflash, Vortex, Shred, Switch, Flip
- Gallery, Cube, Doors, Window, Ferris, Conveyor, Rotate
- Pan, Glitter, Honeycomb, Flash, Ripple, Fracture, Crush
- Peel, PageCurl, Airplane, Origami, Morph

### Transition Properties
- **Speed:** Slow (2s), Medium (1s), Fast (0.5s)
- **Direction:** Type-specific (horizontal, vertical, directional)
- **Advance Mode:**
  - OnClick: Manual advance
  - Automatic: Timed advance (specify milliseconds)
  - Both: Click or wait for timer
- **Sound:** Built-in and external sound support (20 built-in sounds)

## Technical Details

### Shape ID Assignment
Shape IDs in PPT are assigned sequentially starting from 1024. In these examples, we manually track shape IDs to associate animations with specific shapes.

### Binary Format
The examples also save raw binary data (`.bin` files) showing the actual PPT record structure for animations and transitions, useful for:
- Debugging the binary format
- Understanding the record structure
- Comparing with Microsoft PowerPoint's output

### Animation Records
- `AnimationInfo` container (type 4116)
- `BuildList` container (type 2000)
- `BuildAtom` records (type 2001)
- `TimeNode` containers (type 0xF127)

### Transition Records
- `SSSlideInfoAtom` (type 1017)
- Contains transition type, speed, direction, and advance settings

## API Usage

### Creating Animations

```rust
use litchi::ole::ppt::animation::{
    AnimationInfo, BuildInfo, BuildLevel, BuildType,
    AnimationEffect, EffectSpeed, EffectDirection, AnimationTrigger,
};

let mut animation = AnimationInfo::new();
let mut build_list = BuildInfo::new();

build_list.add_build(BuildLevel {
    build_type: BuildType::Entrance,
    shape_id: 1024,  // Target shape ID
    build_order: 0,
    effect: AnimationEffect::FadeIn,
    speed: EffectSpeed::Medium,
    direction: EffectDirection::None,
    trigger: AnimationTrigger::OnClick,
});

animation.build_list = Some(build_list);
```

### Creating Transitions

```rust
use litchi::ole::ppt::transition::{
    TransitionInfo, TransitionType, TransitionSpeed,
    TransitionDirection, AdvanceMode,
};

let transition = TransitionInfo::with_type(TransitionType::Fade)
    .with_speed(TransitionSpeed::Medium)
    .with_direction(TransitionDirection::None)
    .with_advance_mode(AdvanceMode::OnClick);

// For auto-advance
let auto_transition = TransitionInfo::with_type(TransitionType::Dissolve)
    .with_speed(TransitionSpeed::Slow)
    .with_advance_mode(AdvanceMode::Automatic)
    .with_advance_time(3000);  // 3 seconds
```

## Known Limitations

1. Shape IDs must be manually tracked (PptWriter API doesn't return shape IDs)
2. Animation and transition data is created but not yet fully integrated into PptWriter's save method
3. Some advanced animation features (motion paths, time nodes) are defined but not fully implemented in the writer
4. The examples demonstrate the data structures and binary format; full integration with PptWriter is in progress

## Next Steps

To fully integrate animations and transitions into the PptWriter:

1. Add methods to PptWriter API:
   - `set_slide_animation(slide_index, animation_info)`
   - `set_slide_transition(slide_index, transition_info)`
2. Integrate animation/transition records into the save pipeline
3. Update persist pointer mappings to include animation records
4. Add shape ID tracking to PptWriter's shape creation methods

## References

- MS-PPT Specification: Animation and Transition Records
- Apache POI HSLF Implementation
- PowerPoint Animation Effects Documentation

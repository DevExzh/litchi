#!/bin/bash
# Mermaid diagrams — every property + both render modes, enumerating the mermaid
# type gallery.
#
#   render=native  → editable PowerPoint shapes + connectors (no browser).
#                    Supported types: flowchart / graph, sequenceDiagram.
#   render=image   → full-fidelity PNG via real mermaid.js (headless browser).
#                    Covers EVERY mermaid type; source stamped into alt-text.
#   render=auto    → (default) image when a browser is present, else native.
#
# A diagram is an ADD-ONLY synthesizer (like 'equation'): no persistent 'diagram'
# node — the whole picture is ONE object and Add returns its path. Native mode →
# a group of editable shapes+connectors (/slide[N]/group[K]); image mode → a
# single PNG picture (/slide[N]/picture[K]). Both are addressable/movable as a unit.
#
# Source props are interchangeable — mermaid= (canonical), text=, dsl=, or src=
# (a .mmd file). Placement props x/y/width/height define a box the diagram is
# fitted into (aspect ratio preserved, centred in the box). poster= is deck-wide
# (see the note at the end) so it is documented, not baked into this multi-slide file.
#
# NOTE: intentionally NO `set -e` — render=image needs a headless browser
# (Chrome / Chromium / Edge); without one those slides are skipped with a clear
# message while the native slides still build.
#
# Auto-resident disabled: render=image launches a browser and can take a few
# seconds, so each add runs as its own process rather than contending for the
# resident's single command pipe. (The Python SDK twin keeps a resident — its
# client retries a busy connect, so it is safe there.)
export OFFICECLI_NO_AUTO_RESIDENT=1

DIR="$(dirname "$0")"
PPTX="$DIR/diagram.pptx"
MMD="$DIR/pie.mmd"          # loaded via src= on the pie gallery slide

BOX=(--prop x=1in --prop y=1.2in --prop width=11.3in --prop height=5.8in)

# The one flowchart the native + image comparison both render. Exercises the
# node-shape vocabulary the native synthesizer maps: ([stadium]) {diamond} [rect]
# [(database)] [[subroutine]] {{hexagon}} [/parallelogram/] ((circle)); and edge
# forms -->|label|  -.->  ==>  --x .
FLOW="flowchart TD
  A([Start]) --> B{Decision}
  B -->|yes| C[Process]
  B -->|no| D[(Database)]
  C --> E[[Subroutine]]
  D -.-> F{{Prepare}}
  E ==> G((Done))
  F --> G
  A --> H[/Input/]
  H --x B"

rm -f "$PPTX"
officecli create "$PPTX"

title() { # $1=slide  $2=text
    officecli add "$PPTX" "/slide[$1]" --type textbox \
        --prop text="$2" --prop size=24 --prop bold=true \
        --prop x=0.5in --prop y=0.3in --prop width=12.3in --prop height=0.6in
}

# ─────────────────────────────────────────────────────────────────────────────
# Slide 1 — Title
# ─────────────────────────────────────────────────────────────────────────────
officecli add "$PPTX" / --type slide
officecli add "$PPTX" '/slide[1]' --type textbox \
    --prop text="Mermaid Diagrams" --prop size=44 --prop bold=true \
    --prop x=1in --prop y=2.5in --prop width=11.3in --prop height=1in --prop align=center
officecli add "$PPTX" '/slide[1]' --type textbox \
    --prop text="native editable shapes  ·  full-fidelity PNG for every mermaid type" \
    --prop size=20 --prop color=595959 \
    --prop x=1in --prop y=3.6in --prop width=11.3in --prop height=0.6in --prop align=center

# ─────────────────────────────────────────────────────────────────────────────
# Slide 2 — NATIVE flowchart (mermaid= source). Editable shapes + connectors,
#           fitted and CENTRED in the placement box.
# ─────────────────────────────────────────────────────────────────────────────
officecli add "$PPTX" / --type slide
title 2 "render=native — flowchart (editable shapes + connectors)"
officecli add "$PPTX" '/slide[2]' --type diagram --prop render=native --prop mermaid="$FLOW" "${BOX[@]}"
# The whole native diagram is ONE group at the returned path — read its box back.
# `set /slide[2]/group[1] --prop width=…` resizes it as a unit (fonts re-bake);
# `remove /slide[2]/group[1]` deletes the group and every child.
officecli get "$PPTX" '/slide[2]/group[1]'

# ─────────────────────────────────────────────────────────────────────────────
# Slide 3 — IMAGE of the SAME flowchart (dsl= alias). Real mermaid.js → PNG.
#           Apples-to-apples comparison with slide 2.
# ─────────────────────────────────────────────────────────────────────────────
officecli add "$PPTX" / --type slide
title 3 "render=image — the same flowchart as a full-fidelity PNG"
officecli add "$PPTX" '/slide[3]' --type diagram --prop render=image --prop dsl="$FLOW" "${BOX[@]}"

# ─────────────────────────────────────────────────────────────────────────────
# Slide 4 — NATIVE sequenceDiagram (text= alias). The 2nd native-supported type.
# ─────────────────────────────────────────────────────────────────────────────
officecli add "$PPTX" / --type slide
title 4 "render=native — sequenceDiagram"
officecli add "$PPTX" '/slide[4]' --type diagram --prop render=native --prop text="sequenceDiagram
  participant U as User
  participant S as Server
  participant D as Database
  U->>S: Login request
  S->>D: Validate credentials
  D-->>S: OK
  S-->>U: Session token" "${BOX[@]}"

# ─────────────────────────────────────────────────────────────────────────────
# Slides 5+ — IMAGE gallery: every other mermaid type, proving render=image
#             covers the full mermaid surface (these have no native synthesizer).
# ─────────────────────────────────────────────────────────────────────────────
# The pie source is written to a .mmd file and loaded with src= (alias path=).
cat > "$MMD" << 'EOF'
pie showData title Traffic Sources
    "Organic Search" : 45
    "Direct" : 30
    "Referral" : 15
    "Social" : 10
EOF

# One helper per gallery slide (macOS ships bash 3.2 — NO associative arrays, so
# we use plain functions, not `declare -A`). `n` tracks the current slide number.
n=4
img() {     # $1=title suffix   $2=mermaid source (inline)
    n=$((n + 1))
    officecli add "$PPTX" / --type slide
    title "$n" "render=image — $1"
    officecli add "$PPTX" "/slide[$n]" --type diagram --prop render=image --prop text="$2" "${BOX[@]}"
}
img_src() { # $1=title suffix   $2=.mmd file path (loaded with src=, alias path=)
    n=$((n + 1))
    officecli add "$PPTX" / --type slide
    title "$n" "render=image — $1"
    officecli add "$PPTX" "/slide[$n]" --type diagram --prop render=image --prop src="$2" "${BOX[@]}"
}

img_src "pie — proportions" "$MMD"

img "classDiagram — UML classes" "classDiagram
  class Animal { +int age +run() }
  class Dog { +bark() }
  Animal <|-- Dog"

img "stateDiagram-v2 — states" "stateDiagram-v2
  [*] --> Idle
  Idle --> Running: start
  Running --> Idle: pause
  Running --> [*]: stop"

img "erDiagram — entities & relations" "erDiagram
  CUSTOMER ||--o{ ORDER : places
  ORDER ||--|{ LINE_ITEM : contains"

img "gantt — project schedule" "gantt
  title Project Plan
  dateFormat YYYY-MM-DD
  axisFormat %b %d
  tickInterval 1week
  weekday monday
  todayMarker off
  section Design
  Research :a1, 2024-01-01, 5d
  Draft :after a1, 4d
  section Build
  Code :2024-01-12, 6d"

img "journey — user journey" "journey
  title My working day
  section Morning
  Standup: 3: Me, Team
  Code: 5: Me
  section Afternoon
  Review: 2: Me"

img "gitGraph — branch/merge" "gitGraph
  commit
  branch develop
  commit
  checkout main
  merge develop
  commit"

img "mindmap — hierarchy" "mindmap
  root((mermaid))
    Origins
      History
    Uses
      Docs
      Diagrams"

img "timeline — events" "timeline
  title Release History
  2019 : v1
  2021 : v2 : v2.1
  2023 : v3"

img "quadrantChart — 2x2 matrix" "quadrantChart
  title Reach vs Engagement
  x-axis Low Reach --> High Reach
  y-axis Low Engagement --> High Engagement
  quadrant-1 Expand
  quadrant-2 Promote
  quadrant-3 Re-evaluate
  quadrant-4 Improve
  Campaign A: [0.3, 0.6]
  Campaign B: [0.45, 0.23]"

img "requirementDiagram — requirements" "requirementDiagram
  requirement test_req {
    id: 1
    text: the test text
    risk: high
    verifymethod: test
  }
  element test_entity {
    type: simulation
  }
  test_entity - satisfies -> test_req"

img "C4Context — system context" "C4Context
  title System Context
  Person(customer, \"Customer\")
  System(banking, \"Internet Banking\")
  Rel(customer, banking, \"Uses\")"

img "sankey-beta — flow volumes" "sankey-beta
Agricultural,Bio-conversion,124
Bio-conversion,Losses,26
Bio-conversion,Solid,280
Bio-conversion,Gas,81"

img "xychart-beta — bar/line" "xychart-beta
  title \"Monthly Revenue\"
  x-axis [jan, feb, mar, apr]
  y-axis \"Revenue (k\$)\" 0 --> 100
  bar [30, 50, 65, 80]
  line [30, 50, 65, 80]"

img "block-beta — block layout" "block-beta
  columns 3
  a[\"Ingest\"] b[\"Process\"] c[\"Store\"]
  d[\"Log\"]"

img "packet-beta — byte layout" "packet-beta
  0-15: \"Source Port\"
  16-31: \"Destination Port\"
  32-63: \"Sequence Number\""

img "kanban — board" "kanban
  Todo
    t1[Design]
  In Progress
    t2[Build]
  Done
    t3[Ship]"

img "architecture-beta — cloud services" "architecture-beta
  group api(cloud)[API]
  service db(database)[Database] in api
  service server(server)[Server] in api
  db:L -- R:server"

img "radar-beta — multi-axis scores" "radar-beta
  title Skill Assessment
  axis a[\"Coding\"], b[\"Design\"], c[\"Testing\"], d[\"Docs\"]
  curve x{80, 60, 70, 50}"

# poster= is deck-wide: pptx has a single presentation slide size, so poster=true
# grows the WHOLE deck to the diagram's natural size (export-a-diagram-as-a-slide).
# It is mutually exclusive with a multi-slide showcase — use it in a single-diagram
# file:  officecli add poster.pptx '/slide[1]' --type diagram --prop poster=true \
#            --prop mermaid="flowchart LR; A --> B --> C"

officecli validate "$PPTX"
echo "Created: $PPTX ($n slides)"

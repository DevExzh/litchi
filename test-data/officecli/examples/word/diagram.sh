#!/bin/bash
# Mermaid diagrams in Word — native (editable shapes) + image (full-fidelity PNG),
# enumerating the mermaid type gallery.
#
#   render=native  → editable Word drawing shapes + connectors (no browser).
#                    Supported types: flowchart / graph, sequenceDiagram. Placed as
#                    FLOATING shapes anchored to the page margin, so give each native
#                    diagram its own page (page break) to avoid overlapping text.
#   render=image   → full-fidelity PNG via real mermaid.js (headless browser). Covers
#                    EVERY mermaid type and is INLINE, so it flows with the text like
#                    any picture — the natural choice for a flowing document.
#   render=auto    → (default) image when a browser is present, else native.
#
# A diagram is an ADD-ONLY synthesizer (like 'equation'): no persistent 'diagram'
# node — the whole picture is ONE object and Add returns its path. Native mode →
# a group of editable shapes (/body/group[N]); image mode → an inline picture in a
# paragraph. Source props are interchangeable: mermaid= (canonical), text=, dsl=, or
# src= (a .mmd file). width/height fit the diagram (aspect preserved). Word has no
# slide, so there is no x/y and no poster (those are pptx-only).
#
# NOTE: intentionally NO `set -e` — render=image needs a headless browser
# (Chrome / Chromium / Edge); without one those diagrams are skipped with a clear
# message while the native ones still build. Auto-resident is disabled so each add
# is its own process (render=image launches a browser and can take a few seconds).
export OFFICECLI_NO_AUTO_RESIDENT=1

DIR="$(dirname "$0")"
DOCX="$DIR/diagram.docx"
MMD="$DIR/pie.mmd"

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

rm -f "$DOCX"
officecli create "$DOCX"

head() { officecli add "$DOCX" /body --type paragraph --prop text="$1" --prop bold=true --prop size=16; }
para() { officecli add "$DOCX" /body --type paragraph --prop text="$1"; }
pbreak() { officecli add "$DOCX" /body --type paragraph --prop text="$1" --prop bold=true --prop size=16 --prop pageBreakBefore=true; }

# Title
officecli add "$DOCX" /body --type paragraph --prop text="Mermaid Diagrams in Word" --prop bold=true --prop size=24
para "The same --type diagram element as pptx, in a .docx body. Native builds editable Word shapes; image embeds a full-fidelity PNG of any mermaid type."

# ─────────────────────────────────────────────────────────────────────────────
# render=native — editable flowchart (its own page: native diagrams float)
# ─────────────────────────────────────────────────────────────────────────────
pbreak "render=native — flowchart (editable shapes + connectors)"
officecli add "$DOCX" /body --type diagram --prop render=native --prop mermaid="$FLOW" --prop width=12cm
# The whole native diagram is ONE group — read its box back. set /body/group[1]
# --prop width=… resizes it (fonts re-bake); remove /body/group[1] deletes it.
officecli get "$DOCX" '/body/group[1]'

# render=native — sequenceDiagram (text= alias), on its own page
pbreak "render=native — sequenceDiagram"
officecli add "$DOCX" /body --type diagram --prop render=native --prop text="sequenceDiagram
  participant U as User
  participant S as Server
  participant D as Database
  U->>S: Login request
  S->>D: Validate credentials
  D-->>S: OK
  S-->>U: Session token" --prop width=13cm

# ─────────────────────────────────────────────────────────────────────────────
# render=image — inline PNGs that flow with the text (every mermaid type)
# ─────────────────────────────────────────────────────────────────────────────
pbreak "render=image — full-fidelity PNG (real mermaid.js)"
para "Image diagrams are inline pictures, so they flow with the text. The same flowchart as above, now a PNG (dsl= is an alias of mermaid=):"
officecli add "$DOCX" /body --type diagram --prop render=image --prop dsl="$FLOW" --prop width=14cm

# The pie source is loaded from a .mmd file via src= (alias path=).
cat > "$MMD" << 'EOF'
pie showData title Traffic Sources
    "Organic Search" : 45
    "Direct" : 30
    "Referral" : 15
    "Social" : 10
EOF

imgdoc() {     # $1=heading  $2=mermaid source (inline, text=)
    head "render=image — $1"
    officecli add "$DOCX" /body --type diagram --prop render=image --prop text="$2" --prop width=14cm
}
imgdoc_src() { # $1=heading  $2=.mmd file (src=)
    head "render=image — $1"
    officecli add "$DOCX" /body --type diagram --prop render=image --prop src="$2" --prop width=10cm
}

imgdoc_src "pie — proportions" "$MMD"

imgdoc "classDiagram — UML classes" "classDiagram
  class Animal { +int age +run() }
  class Dog { +bark() }
  Animal <|-- Dog"

imgdoc "stateDiagram-v2 — states" "stateDiagram-v2
  [*] --> Idle
  Idle --> Running: start
  Running --> Idle: pause
  Running --> [*]: stop"

imgdoc "erDiagram — entities & relations" "erDiagram
  CUSTOMER ||--o{ ORDER : places
  ORDER ||--|{ LINE_ITEM : contains"

imgdoc "gantt — project schedule" "gantt
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

imgdoc "journey — user journey" "journey
  title My working day
  section Morning
  Standup: 3: Me, Team
  Code: 5: Me
  section Afternoon
  Review: 2: Me"

imgdoc "gitGraph — branch/merge" "gitGraph
  commit
  branch develop
  commit
  checkout main
  merge develop
  commit"

imgdoc "mindmap — hierarchy" "mindmap
  root((mermaid))
    Origins
      History
    Uses
      Docs
      Diagrams"

imgdoc "timeline — events" "timeline
  title Release History
  2019 : v1
  2021 : v2 : v2.1
  2023 : v3"

imgdoc "quadrantChart — 2x2 matrix" "quadrantChart
  title Reach vs Engagement
  x-axis Low Reach --> High Reach
  y-axis Low Engagement --> High Engagement
  quadrant-1 Expand
  quadrant-2 Promote
  quadrant-3 Re-evaluate
  quadrant-4 Improve
  Campaign A: [0.3, 0.6]
  Campaign B: [0.45, 0.23]"

imgdoc "requirementDiagram — requirements" "requirementDiagram
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

imgdoc "C4Context — system context" "C4Context
  title System Context
  Person(customer, \"Customer\")
  System(banking, \"Internet Banking\")
  Rel(customer, banking, \"Uses\")"

imgdoc "sankey-beta — flow volumes" "sankey-beta
Agricultural,Bio-conversion,124
Bio-conversion,Losses,26
Bio-conversion,Solid,280
Bio-conversion,Gas,81"

imgdoc "xychart-beta — bar/line" "xychart-beta
  title \"Monthly Revenue\"
  x-axis [jan, feb, mar, apr]
  y-axis \"Revenue (k\$)\" 0 --> 100
  bar [30, 50, 65, 80]
  line [30, 50, 65, 80]"

imgdoc "block-beta — block layout" "block-beta
  columns 3
  a[\"Ingest\"] b[\"Process\"] c[\"Store\"]
  d[\"Log\"]"

imgdoc "packet-beta — byte layout" "packet-beta
  0-15: \"Source Port\"
  16-31: \"Destination Port\"
  32-63: \"Sequence Number\""

imgdoc "kanban — board" "kanban
  Todo
    t1[Design]
  In Progress
    t2[Build]
  Done
    t3[Ship]"

imgdoc "architecture-beta — cloud services" "architecture-beta
  group api(cloud)[API]
  service db(database)[Database] in api
  service server(server)[Server] in api
  db:L -- R:server"

imgdoc "radar-beta — multi-axis scores" "radar-beta
  title Skill Assessment
  axis a[\"Coding\"], b[\"Design\"], c[\"Testing\"], d[\"Docs\"]
  curve x{80, 60, 70, 50}"

officecli validate "$DOCX"
echo "Created: $DOCX"

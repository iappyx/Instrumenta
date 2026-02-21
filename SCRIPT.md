# Instrumenta Script Language

The Instrumenta Script Language is a simple scripting language built into Instrumenta that lets you automate and batch-manipulate shapes in PowerPoint. Instead of clicking through menus repeatedly, you write a short script and run it with a single click.

Scripts work on the **current slide** only.

---

## Basic concepts

### Working set
The working set is the collection of shapes that `SET`, `ROTATE`, and `GROUP` commands operate on. You define it using `SELECT`, `USE SELECTION`, or `INSERT`. It persists until you define a new one.

```
SELECT ALL
SET font.size = 12      # applies to all shapes

SELECT WHERE type = TEXTBOX
SET font.bold = TRUE    # applies only to textboxes
```

### Comments and blank lines
Lines starting with `#` are comments and are ignored. Blank lines are also ignored.

```
# This is a comment
SELECT ALL

# Another comment
SET font.size = 12
```

### Case insensitivity
Commands and keywords are case-insensitive. `SELECT ALL`, `select all` and `Select All` all work the same.

---

## Commands

### SELECT
Defines the working set by filtering shapes on the current slide. Also syncs the PowerPoint selection visually.

```
SELECT ALL
SELECT WHERE name = "myshape"
SELECT WHERE name CONTAINS "box"
SELECT WHERE name STARTSWITH "slide_"
SELECT WHERE type = <type>
```

#### Available types for SELECT WHERE type =

**Basic shapes**
`RECTANGLE`, `ROUNDEDRECTANGLE`, `OVAL`, `TRIANGLE`, `RIGHTTRIANGLE`, `DIAMOND`, `PARALLELOGRAM`, `TRAPEZOID`, `HEXAGON`, `PENTAGON`, `OCTAGON`

**Arrows**
`ARROWRIGHT`, `ARROWLEFT`, `ARROWUP`, `ARROWDOWN`, `ARROWLEFTRIGHT`, `CHEVRON`, `PENTAGON_ARROW`, `CIRCULARRIGHTARROW`

**Flowchart**
`FLOWCHART_PROCESS`, `FLOWCHART_DECISION`, `FLOWCHART_TERMINATOR`, `FLOWCHART_DATA`, `FLOWCHART_DOCUMENT`, `FLOWCHART_CONNECTOR`

**Callouts**
`CALLOUT_RECT`, `CALLOUT_OVAL`, `CALLOUT_CLOUD`

**Other**
`TEXTBOX`, `PICTURE`, `TABLE`, `LINE`

Name and type filters support loop variable expressions:

```
REPEAT 5 AS i
    SELECT WHERE name = "box_"+i
    SET fill.color = #003366
END REPEAT
```

---

### USE SELECTION
Sets the working set to whatever is currently selected in PowerPoint. Useful for running a script against a manual selection.

```
USE SELECTION
SET font.size = 14
SET font.bold = TRUE
```

> After a `CALL` command, use `USE SELECTION` to re-sync the working set if the called sub changed the PowerPoint selection.

---

### INSERT
Creates a new shape on the current slide and makes it the working set.

```
INSERT <type> AT x, y WIDTH w HEIGHT h [NAME "myname"] [TEXT "hello"]
```

- All numeric arguments (`x`, `y`, `w`, `h`) support expressions including loop variables
- `NAME` and `TEXT` support string expressions like `"box_"+i`
- If `NAME` is omitted, a name is auto-generated (e.g. `script_rect_1`)
- If a shape with the given name already exists, `INSERT` fails with an error
- `TEXT` is supported on all shapes that have a text frame

#### Available shape types for INSERT

**Basic shapes**
`RECTANGLE`, `ROUNDEDRECTANGLE`, `OVAL`, `TRIANGLE`, `RIGHTTRIANGLE`, `DIAMOND`, `PARALLELOGRAM`, `TRAPEZOID`, `HEXAGON`, `PENTAGON`, `OCTAGON`

**Arrows**
`ARROWRIGHT`, `ARROWLEFT`, `ARROWUP`, `ARROWDOWN`, `ARROWLEFTRIGHT`, `CHEVRON`, `PENTAGON_ARROW`, `CIRCULARRIGHTARROW`

**Flowchart**
`FLOWCHART_PROCESS`, `FLOWCHART_DECISION`, `FLOWCHART_TERMINATOR`, `FLOWCHART_DATA`, `FLOWCHART_DOCUMENT`, `FLOWCHART_CONNECTOR`

**Callouts**
`CALLOUT_RECT`, `CALLOUT_OVAL`, `CALLOUT_CLOUD`

**Text**
`TEXTBOX`

```
REPEAT 4 AS col
    INSERT CHEVRON AT 50+(col*110), 100 WIDTH 100 HEIGHT 60 NAME "step_"+col TEXT "Step "+col
    SET fill.color = #003366
    SET font.color = #FFFFFF
END REPEAT
```

#### INSERT LINE
Lines use a different syntax because they are defined by two points rather than a position and size.

```
INSERT LINE FROM x1, y1 TO x2, y2 [NAME "myname"]
```

- All coordinates support expressions and variables
- `NAME` is optional — auto-generates `script_line_1` if omitted
- After insert, the line becomes the working set so you can immediately apply `SET border.color`, `SET border.width`, `SET border.style`, and `SET opacity`
- `TEXT` is not supported on lines

```
# Simple horizontal divider
INSERT LINE FROM 40, 200 TO 680, 200 NAME "divider"
SET border.color = #CCCCCC
SET border.width = 1

# Diagonal line using variables
SET VAR x1 = 50
SET VAR y1 = 50
SET VAR x2 = 350
SET VAR y2 = 300
INSERT LINE FROM x1, y1 TO x2, y2 NAME "diagonal"
SET border.color = #003366
SET border.width = 2
SET border.style = DASH

# Grid lines using a loop
SET VAR gridX = 100
SET VAR gridY = 60
SET VAR gridH = 200
SET VAR colW = 80
REPEAT 4 AS i
    INSERT LINE FROM gridX+(i*colW), gridY TO gridX+(i*colW), gridY+gridH NAME "vline_"+i
    SET border.color = #DDDDDD
    SET border.width = 1
END REPEAT
```

---

### DELETE
Deletes shapes from the current slide. Clears the working set.

```
DELETE WHERE name = "myname"
DELETE WHERE name CONTAINS "mytext"
DELETE WHERE name STARTSWITH "script_"
DELETE WHERE type = RECTANGLE
DELETE SELECTION
```

> **Tip:** Start scripts with `DELETE WHERE name STARTSWITH "..."` to make them safely re-runnable.

---

### SET
Applies a property to all shapes in the current working set.

#### Font properties
| Property | Value | Example |
|---|---|---|
| `font.size` | expression | `SET font.size = 10+i*2` |
| `font.bold` | TRUE / FALSE | `SET font.bold = TRUE` |
| `font.italic` | TRUE / FALSE | `SET font.italic = FALSE` |
| `font.underline` | TRUE / FALSE | `SET font.underline = TRUE` |
| `font.color` | #RRGGBB | `SET font.color = #FFFFFF` |
| `font.name` | font name | `SET font.name = "Calibri"` |

#### Fill properties
| Property | Value | Example |
|---|---|---|
| `fill.color` | #RRGGBB | `SET fill.color = #003366` |
| `fill.transparent` | TRUE / FALSE | `SET fill.transparent = TRUE` |
| `fill.gradient` | `"#color1,#color2"` | `SET fill.gradient = "#003366,#56A8E0"` |
| `fill.gradient.direction` | HORIZONTAL / VERTICAL / DIAGONAL / DIAGONAL_DOWN | `SET fill.gradient.direction = VERTICAL` |

> `fill.gradient` applies a two-stop gradient (horizontal by default). Use `fill.gradient.direction` after `fill.gradient` to change the angle — it reads the existing colors and re-applies them in the new direction.

#### Text alignment
| Property | Value | Example |
|---|---|---|
| `text` | string | `SET text = "new label"` |
| `text.align` | LEFT / CENTER / RIGHT / JUSTIFY | `SET text.align = CENTER` |
| `text.valign` | TOP / MIDDLE / BOTTOM | `SET text.valign = MIDDLE` |

#### Shadow
| Property | Value | Example |
|---|---|---|
| `shadow` | TRUE / FALSE | `SET shadow = TRUE` |
| `shadow.color` | #RRGGBB | `SET shadow.color = #AAAAAA` |
| `shadow.offset.x` | expression (pt) | `SET shadow.offset.x = 4` |
| `shadow.offset.y` | expression (pt) | `SET shadow.offset.y = 4` |

> Setting `shadow.color` or `shadow.offset.*` automatically enables the shadow.

#### Connector
| Property | Value | Example |
|---|---|---|
| `connector.style` | STRAIGHT / ELBOW / CURVED | `SET connector.style = ELBOW` |

> `connector.style` only applies to shapes created with `INSERT CONNECTOR`. Logs a warning if the shape is not a connector.

#### Other
| Property | Value | Example |
|---|---|---|
| `opacity` | 0–100 | `SET opacity = 100-i*10` |
| `name` | string | `SET name = "new_name"` |
| `z.order` | FRONT / BACK / FORWARD / BACKWARD | `SET z.order = FRONT` |
| Property | Value | Example |
|---|---|---|
| `border.color` | #RRGGBB | `SET border.color = #000000` |
| `border.width` | expression (pt) | `SET border.width = 2` |
| `border.visible` | TRUE / FALSE | `SET border.visible = FALSE` |
| `border.style` | SOLID / DASH / DOT / DASHDOT | `SET border.style = DASH` |
| `border.radius` | 0–100 | `SET border.radius = 50` |

> Setting `border.color` or `border.width` automatically makes the border visible. `border.radius` only applies to `ROUNDEDRECTANGLE` shapes — 0 = square corners, 100 = fully rounded.

#### Size and position
| Property | Value | Example |
|---|---|---|
| `width` | expression | `SET width = 100+i*10` |
| `height` | expression | `SET height = 80` |
| `position.x` | expression | `SET position.x = 50+col*120` |
| `position.y` | expression | `SET position.y = 50+row*80` |

```
# Gradient fill with direction
INSERT RECTANGLE AT 50, 50 WIDTH 200 HEIGHT 60 NAME "grad_box"
SET fill.gradient = "#003366,#56A8E0"
SET fill.gradient.direction = VERTICAL
SET border.visible = FALSE

# Rounded rectangle with custom corner radius
INSERT ROUNDEDRECTANGLE AT 50, 130 WIDTH 200 HEIGHT 60 NAME "rr_box"
SET fill.color = #003366
SET border.radius = 30
SET border.visible = FALSE
SET font.color = #FFFFFF
SET text.align = CENTER
SET text.valign = MIDDLE

# Shadow
INSERT RECTANGLE AT 50, 210 WIDTH 200 HEIGHT 60 NAME "shadow_box"
SET fill.color = #FFFFFF
SET border.color = #CCCCCC
SET border.width = 1
SET shadow = TRUE
SET shadow.color = #AAAAAA
SET shadow.offset.x = 4
SET shadow.offset.y = 4

# Connector style
INSERT CONNECTOR FROM "grad_box" TO "rr_box" NAME "conn_1"
SET connector.style = ELBOW
SET border.color = #003366
SET border.width = 2

# Z-order and text
SELECT WHERE name = "shadow_box"
SET z.order = FRONT
SET text = "Updated"
```

---

### SET VAR
Declares a named variable for use throughout the script. Variables can hold numeric values or strings, and can be used in any expression where a number or string is expected.

```
SET VAR name = <numeric expression>
SET VAR name = "string value"
```

- Variable names are case-insensitive and may contain letters, digits, and underscores
- Numeric variables work in any numeric expression: `AT`, `WIDTH`, `HEIGHT`, `ROTATE`, `SET` values
- String variables work in `NAME`, `TEXT`, color values, font names, and any string expression
- String values must be quoted — including colors (e.g. `"#003366"`)
- Variables can reference previously declared variables in their definition
- Declaring a variable that already exists overwrites it

```
# Declare constants at the top
SET VAR originX = 40
SET VAR originY = 60
SET VAR cellW = 140
SET VAR cellH = 90
SET VAR prefix = "grid_"
SET VAR primary = "#003366"
SET VAR fontName = "Calibri"

# Use them throughout the script
REPEAT 3 AS row
    REPEAT 4 AS col
        INSERT RECTANGLE AT originX+(col*cellW), originY+(row*cellH) WIDTH cellW HEIGHT cellH NAME prefix+row+"_"+col
        SET fill.color = primary
        SET font.name = fontName
    END REPEAT
END REPEAT
```

#### Numeric variables
Numeric variables are substituted into expressions before evaluation, so they work identically to loop variables.

```
SET VAR gap = 10
SET VAR boxW = 120

INSERT RECTANGLE AT 50, 100 WIDTH boxW HEIGHT 60 NAME "box_a"
INSERT RECTANGLE AT 50+(boxW+gap), 100 WIDTH boxW HEIGHT 60 NAME "box_b"
INSERT RECTANGLE AT 50+(boxW+gap)*2, 100 WIDTH boxW HEIGHT 60 NAME "box_c"
```

#### String variables
String variables are substituted into string expressions before evaluation. They must be declared with quoted values.

```
SET VAR prefix = "chart_"
SET VAR title = "Revenue"
SET VAR accent = "#E8A000"

INSERT RECTANGLE AT 50, 50 WIDTH 200 HEIGHT 40 NAME prefix+"header" TEXT title
SET fill.color = accent
SET font.color = "#FFFFFF"
```

> Color values assigned to string variables must be quoted: `SET VAR primary = "#003366"`, not `SET VAR primary = #003366`.

#### Built-in variables
A set of read-only variables is automatically populated at script start from the active presentation. You can use them in any expression but cannot overwrite them — a `SET VAR` targeting these names will be ignored with a warning.

| Variable | Value |
|---|---|
| `slideWidth` | Width of the current slide in points |
| `slideHeight` | Height of the current slide in points |
| `slideCenterX` | `slideWidth / 2` |
| `slideCenterY` | `slideHeight / 2` |

A standard 16:9 slide is 720 × 405 pt, so `slideWidth = 720`, `slideHeight = 405`, `slideCenterX = 360`, `slideCenterY = 202.5`.

```
# Center a box on any slide size
SET VAR boxW = 200
SET VAR boxH = 80
INSERT RECTANGLE AT slideCenterX-(boxW/2), slideCenterY-(boxH/2) WIDTH boxW HEIGHT boxH NAME "center_box"

# Right-align a column to the slide edge
SET VAR margin = 40
SET VAR colW = 160
INSERT RECTANGLE AT slideWidth-margin-colW, 60 WIDTH colW HEIGHT 280 NAME "right_col"
```

#### Reading shape properties — GET
You can read a property from an existing named shape into a variable using `GET`:

```
SET VAR name = GET property FROM "shapename"
```

Supported properties: `position.x`, `position.y`, `width`, `height`, `opacity`, `rotation`, `font.size`, `name`, `text`

```
# Position a new shape directly below an existing one
SET VAR srcY = GET position.y FROM "header_box"
SET VAR srcH = GET height FROM "header_box"
INSERT RECTANGLE AT 50, srcY+srcH+10 WIDTH 200 HEIGHT 60 NAME "body_box"

# Copy width from one shape to another
SET VAR refW = GET width FROM "col_a"
SELECT WHERE name = "col_b"
SET width = refW
```

#### Interactive input — INPUT
Pauses the script and shows a text box prompting the user for a value. If the entered value is numeric it is stored as a number, otherwise as a string. Cancelling or leaving blank skips the assignment with a warning.

```
SET VAR name = INPUT "prompt text"
```

```
SET VAR slideTitle = INPUT "Enter slide title:"
SET VAR colCount = INPUT "How many columns?"

INSERT TEXTBOX AT 40, 20 WIDTH 400 HEIGHT 40 NAME "title_box" TEXT slideTitle
SET font.size = 24
SET font.bold = TRUE

REPEAT colCount AS i
    INSERT RECTANGLE AT 40+(i*160), 80 WIDTH 150 HEIGHT 200 NAME "col_"+i
    SET fill.color = #003366
    SET font.color = #FFFFFF
END REPEAT
```

---

### ROTATE
Rotates all shapes in the working set. Supports absolute and relative rotation.

```
ROTATE <angle>       # set absolute rotation in degrees
ROTATE BY <angle>    # rotate relative to current angle
```

Both forms support expressions and loop variables.

```
ROTATE 45             # set rotation to exactly 45 degrees
ROTATE BY 90          # add 90 degrees to current rotation
ROTATE BY i*15        # rotate progressively in a loop
```

```
REPEAT 6 AS i
    INSERT DIAMOND AT 50+(i*90), 100 WIDTH 70 HEIGHT 70 NAME "diamond_"+i
    ROTATE BY i*10
END REPEAT
```

---

### GROUP
Groups all shapes in the working set into a single group, which then becomes the new working set. Optionally assigns a name to the group.

```
GROUP
GROUP NAME "mygroup"
```

Requires at least 2 shapes in the working set.

```
SELECT WHERE name STARTSWITH "step_"
GROUP NAME "process_flow"
SET opacity = 80
```

After grouping, you can `SET` properties on the group as a whole, or `SELECT` it by name later.

---

### UNGROUP
Ungroups all groups in the current working set. The ungrouped child shapes become the new working set. Non-group shapes in the working set are passed through unchanged.

```
UNGROUP
```

```
SELECT WHERE name = "process_flow"
UNGROUP
SET font.size = 10      # applies to all ungrouped children
```

---

### DUPLICATE
Duplicates all shapes in the current working set. The duplicates become the new working set.

```
DUPLICATE
DUPLICATE OFFSET dx, dy
```

- Without `OFFSET`, duplicates are placed 10pt right and 10pt down from the originals
- `OFFSET` accepts expressions and variables
- Each duplicate is named `originalname_dup1`, `_dup2`, etc.

```
INSERT RECTANGLE AT 50, 50 WIDTH 120 HEIGHT 60 NAME "card"
SET fill.color = #003366
SET border.visible = FALSE

DUPLICATE OFFSET 130, 0
SET fill.color = #0055AA

DUPLICATE OFFSET 260, 0
SET fill.color = #56A8E0
```

---

### INSERT CONNECTOR
Creates a straight connector line between two named shapes. The connector stays attached when shapes are moved.

```
INSERT CONNECTOR FROM "shapename1" TO "shapename2" [NAME "myname"]
```

- Both shape names support string expressions and variables
- `NAME` is optional — auto-generates `script_conn_1` if omitted
- The connector becomes the working set after insert, so you can immediately apply `SET border.color`, `SET border.width`, `SET border.style`
- Connection points are auto-selected by PowerPoint (connection point 0 = closest)

```
INSERT RECTANGLE AT 50, 100 WIDTH 120 HEIGHT 60 NAME "box_a"
SET fill.color = #003366
SET border.visible = FALSE

INSERT RECTANGLE AT 300, 100 WIDTH 120 HEIGHT 60 NAME "box_b"
SET fill.color = #003366
SET border.visible = FALSE

INSERT CONNECTOR FROM "box_a" TO "box_b" NAME "conn_ab"
SET border.color = #003366
SET border.width = 2

# With loop variables in names
REPEAT 3 AS i
    INSERT CONNECTOR FROM "step_"+i TO "step_"+(i+1) NAME "conn_"+i
    SET border.color = #56A8E0
    SET border.width = 1
END REPEAT
```

---

### CALL
Calls any public VBA sub by name — including all Instrumenta functions.

```
SELECT WHERE name STARTSWITH "chart_"
CALL ObjectsAlignTops

# Re-sync working set after CALL if needed
USE SELECTION
SET font.size = 10
```

> `CALL` does **not** update the working set. Use `USE SELECTION` afterwards if the called sub changed the PowerPoint selection.

---

### REPEAT
Loops a block of commands a fixed number of times, with a loop variable.

```
REPEAT count AS variable [FROM start] [STEP step]
    ...
END REPEAT
```

- `count` — number of iterations
- `AS variable` — loop variable name (usable in expressions)
- `FROM start` — starting value (default: 0)
- `STEP step` — increment per iteration (default: 1)

```
REPEAT 5 AS i               # i = 0, 1, 2, 3, 4
REPEAT 5 AS i FROM 10       # i = 10, 11, 12, 13, 14
REPEAT 5 AS i STEP 2        # i = 0, 2, 4, 6, 8
REPEAT 5 AS i FROM 1 STEP 2 # i = 1, 3, 5, 7, 9
```

Loops can be nested:

```
REPEAT 3 AS row
    REPEAT 4 AS col
        INSERT RECTANGLE AT 50+(col*120), 50+(row*90) WIDTH 110 HEIGHT 80 NAME "cell_"+row+"_"+col
    END REPEAT
END REPEAT
```

### BREAK
Exits the innermost loop immediately.

```
REPEAT 100 AS i
    SELECT WHERE name = "box_"+i
    IF i > 5
        BREAK
    END IF
    SET fill.color = #003366
END REPEAT
```

In nested loops, `BREAK` exits only the innermost loop.

---

### IF / ELSE IF / ELSE / END IF

```
IF condition
    ...
ELSE IF condition
    ...
ELSE
    ...
END IF
```

#### Comparison operators
| Operator | Meaning |
|---|---|
| `=` | equal |
| `<>` | not equal |
| `>` | greater than |
| `<` | less than |
| `>=` | greater or equal |
| `<=` | less or equal |

#### Logical operators
| Operator | Meaning |
|---|---|
| `AND` | both conditions true |
| `OR` | either condition true |
| `NOT` | negation |

#### Precedence (highest to lowest)
`( )` → `NOT` → `AND` → `OR`

```
REPEAT 5 AS i
    INSERT RECTANGLE AT 50+(i*110), 100 WIDTH 100 HEIGHT 80 NAME "box_"+i

    IF i = 0
        SET fill.color = #CC0000
        SET font.color = #FFFFFF
    ELSE IF i > 0 AND i < 4
        SET fill.color = #003366
        SET font.color = #FFFFFF
    ELSE
        SET fill.color = #E8F0F8
        SET font.color = #000000
    END IF
END REPEAT
```

---

## Expressions

### Numeric expressions
Used anywhere a number is expected: `AT`, `WIDTH`, `HEIGHT`, `ROTATE`, and `SET` values.

```
AT 50+i*120, 100
WIDTH (col+1)*100
SET font.size = 10+i*2
SET position.x = 50+(col*(width+gap))
SET opacity = 100-(i*10)
ROTATE BY i*15
```

**Operators:** `+` `-` `*` `/` with parentheses. Standard precedence: `( )` → `* /` → `+ -`.

### String expressions
Used in `NAME`, `TEXT`, and string-based SET values.

```
NAME "box_"+i
NAME "row_"+row+"_col_"+col
TEXT "Step "+i
```

String parts in quotes are literals. Numbers and variables are automatically converted to integers when they are whole numbers.

---

## Full example

```
# Clean up from previous run
DELETE WHERE name STARTSWITH "grid_"
DELETE WHERE name STARTSWITH "step_"

# Build a 3x4 grid with styled cells and borders
REPEAT 3 AS row
    REPEAT 4 AS col
        INSERT RECTANGLE AT 50+(col*150), 60+(row*100) WIDTH 140 HEIGHT 90 NAME "grid_"+row+"_"+col TEXT "R"+row+" C"+col

        IF row = 0 AND col = 0
            SET fill.color = #CC0000
            SET font.color = #FFFFFF
            SET font.bold = TRUE
        ELSE IF row = 0
            SET fill.color = #003366
            SET font.color = #FFFFFF
            SET font.bold = TRUE
        ELSE IF col = 0
            SET fill.color = #0055AA
            SET font.color = #FFFFFF
        ELSE IF (row+col) = 4
            SET fill.color = #E8A000
            SET font.color = #FFFFFF
        ELSE
            SET fill.color = #E8F0F8
            SET font.color = #000000
        END IF

        SET font.size = 11
        SET font.name = "Calibri"
        SET border.color = #CCCCCC
        SET border.width = 1
    END REPEAT
END REPEAT

# Align and group the grid
SELECT WHERE name STARTSWITH "grid_"
CALL ObjectsAlignTops
GROUP NAME "grid_group"

# Build a chevron process flow below the grid
REPEAT 5 AS i
    INSERT CHEVRON AT 50+(i*130), 360 WIDTH 120 HEIGHT 50 NAME "step_"+i TEXT "Step "+(i+1)
    SET fill.color = #003366
    SET font.color = #FFFFFF
    SET font.size = 11
    SET font.bold = TRUE
END REPEAT
```

---

## Tips

**Make scripts re-runnable.** Start with `DELETE WHERE name STARTSWITH "..."` to clean up before inserting. Use a consistent prefix for all shapes a script creates.

**Use naming conventions.** Prefixing shape names (e.g. `chart_`, `layout_`, `grid_`) makes `STARTSWITH` very powerful for selecting groups of related shapes, including inside loops.

**Points reference.** PowerPoint uses points as its unit. A standard 16:9 slide is 720 × 405 pt. 1 cm ≈ 28.35 pt. Font size 12 = 12 pt.

**CALL any public sub.** `CALL` works with any public sub in any loaded VBA module, not just Instrumenta functions.

**No parentheses needed for simple conditions.** `IF i > 2 AND i < 7` works fine without parentheses. Use them only when you need to override the default `AND` before `OR` precedence.

**ROTATE BY vs ROTATE.** Use `ROTATE` to set an exact angle regardless of current rotation. Use `ROTATE BY` to add to the existing angle — useful inside loops to create fan or spiral effects.

**GROUP then SET.** After `GROUP`, the group itself becomes the working set, so you can immediately apply `SET` or `ROTATE` to the whole group as one object.

**Use SET VAR for constants.** Declare layout values like `originX`, `cellW`, and colors at the top of a script using `SET VAR`. Change them once and the whole layout updates.

**GET for relative positioning.** `SET VAR y = GET position.y FROM "header"` then `SET VAR h = GET height FROM "header"` lets you place the next shape exactly below an existing one without hardcoding coordinates. Makes scripts resilient to layout changes.

**INPUT for reusable scripts.** Scripts with `INPUT` prompts can be shared and run by anyone without editing the script. Use it for column counts, title text, colors, or any value that changes between uses.

**border.radius on ROUNDEDRECTANGLE.** The default radius is whatever PowerPoint picks. Use `SET border.radius = 20` for subtle rounding or `SET border.radius = 100` for pill shapes.

**fill.gradient.direction must come after fill.gradient.** The direction command reads the existing gradient stop colors and re-applies them, so `fill.gradient` must be set first on the same shape.

**connector.style after INSERT CONNECTOR.** Switch a connector from STRAIGHT to ELBOW or CURVED immediately after inserting it, before moving to the next shape.

**UNGROUP then SET.** `UNGROUP` puts all children back in the working set, so you can immediately `SET` a property across all of them in one step — useful when you want to restyle a group's contents without ungrouping manually in PowerPoint.
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

#### Border properties
| Property | Value | Example |
|---|---|---|
| `border.color` | #RRGGBB | `SET border.color = #000000` |
| `border.width` | expression (pt) | `SET border.width = 2` |
| `border.visible` | TRUE / FALSE | `SET border.visible = FALSE` |
| `border.style` | SOLID / DASH / DOT / DASHDOT | `SET border.style = DASH` |

> Setting `border.color` or `border.width` automatically makes the border visible.

#### Size and position
| Property | Value | Example |
|---|---|---|
| `width` | expression | `SET width = 100+i*10` |
| `height` | expression | `SET height = 80` |
| `position.x` | expression | `SET position.x = 50+col*120` |
| `position.y` | expression | `SET position.y = 50+row*80` |

#### Other
| Property | Value | Example |
|---|---|---|
| `opacity` | 0–100 | `SET opacity = 100-i*10` |
| `name` | string | `SET name = "new_name"` |

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


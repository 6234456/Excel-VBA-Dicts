# CLAUDE.md — Excel-VBA-Dicts

## Project Overview

Excel-VBA-Dicts is a VBA library that brings **functional programming** idioms (map, filter, reduce, groupBy, leftJoin, etc.) to Excel/VBA. It implements native data structures (Dicts, Lists, HashSets, TreeSets) entirely in VBA, with full macOS-Excel compatibility, JSON support, and spreadsheet I/O.

A companion Python module (`xlDataStructure.py`) provides an equivalent `xlDict` class using `win32com` for automation on Windows.

---

## Repository Structure

```
Excel-VBA-Dicts/
├── Dicts.bas          # Core dictionary class (TreeMap-like, sorted keys)
├── Lists.bas          # Dynamic array list with functional ops
├── HashSets.bas       # Hash set backed by Lists
├── TreeSets.bas       # Sorted set (BST) backed by Nodes
├── Nodes.bas          # BST node used by TreeSets
├── util.bas           # Utility: filesystem workbook iteration helper
├── Test.bas           # Unit tests using Debug.Assert
├── xlDataStructure.py # Python xlDict class (win32com automation)
├── README.md
└── LICENSE            # MIT
```

All `.bas` files are VBA class/module files meant to be imported into an Excel workbook via the VBA IDE (Alt+F11 → File → Import File).

---

## Module Descriptions

### `Dicts.bas` — Core Dictionary

A sorted dictionary (TreeMap-like) using `HashSets` for keys and a `Collection` for values.

**Dependencies:** `Lists`, `HashSets`, `Nodes`
**Author:** Xiou Yang
**Last updated:** 20.04.2023

**Key enums:**
```vba
Enum ProcessWith
    key = 0
    value = 1
    RangedValue = 2
End Enum

Enum AggregateMethod
    AggMap = 0
    AggReduce = 1
    Aggfilter = 2
End Enum
```

**Core API:**

| Method/Property | Description |
|---|---|
| `add(k, v)` | Add key-value pair; returns `Me` (chainable) |
| `Item(key)` | Get/set by key (default property) |
| `exists(key)` | Check if key exists |
| `Remove(e)` / `RemoveAll()` / `clear()` | Remove entries |
| `Count` | Number of entries (optional recursive) |
| `keysArr` / `valsArr` | Keys/values as arrays |
| `label` / `setLabel(rng)` | Column labels (for DataFrame-like usage) |
| `getByLabel(k, label)` | Retrieve element by key + column label |

**Functional operations (inline expression syntax):**

| Method | Description |
|---|---|
| `map(operation)` | Transform values using string expression |
| `mapKey(operation)` | Transform keys |
| `filter(operation)` | Filter by value |
| `filterKey(operation)` | Filter by key |
| `reduce(operation, initialVal)` | Reduce values |
| `reduceKey(operation, initialVal)` | Reduce keys |
| `ranged(operation, aggregate)` | Windowed/rolling aggregation |
| `groupBy(attr, valCol)` | Group by attribute |
| `leftJoin(dict2)` | Left join two dicts |
| `diff(dict2)` | Keys in self not in dict2 |
| `union(dict2)` | Merge (keep original values by default) |
| `intersect(dict2)` | Keys in both |
| `mergeMap(op, other)` | Element-wise binary operation across two dicts |
| `sort(isAscending)` | Sort by key |

**Callback variants (for complex logic):**

| Method | Description |
|---|---|
| `mapX(callback)` | map with external VBA function |
| `filterX(callback)` | filter with external VBA function |
| `reduceX(callback, initVal)` | reduce with external VBA function |
| `productX(other, callback)` | cartesian product with callback |

**Spreadsheet I/O:**

| Method | Description |
|---|---|
| `load(sht, keyCol, valCol, ...)` | Load from worksheet (vertical) |
| `loadH(sht, keyRow, valRow, ...)` | Load from worksheet (horizontal) |
| `loadStruct(sht, keyCol1, keyCol2, valCol, ...)` | Load 2-level keyed structure |
| `unload(sht, keyPos, ...)` | Write values back to sheet |
| `dump(sht, ...)` | Write keys + values to sheet |
| `fromRng(rng)` | Load from a Range object |
| `fromMatrix(l)` | Build from 2D Lists |
| `fromArray(arr)` | Build from array/Lists/Collection |

**Utilities:**

| Method | Description |
|---|---|
| `p()` | Debug.Print the dict |
| `pk()` | Debug.Print keys only |
| `toString()` | String representation |
| `toJSON(exportTo)` | Serialize to JSON |
| `fromString(s)` | Deserialize from JSON-like string |
| `rng(start, ending, steps)` | Generate a range array |
| `x(sht, row)` / `y(sht, col)` | Get last used column/row on sheet |
| `reg(pattern, flag)` | Create a VBScript RegExp object |
| `frequencyCount(rng)` | Count occurrences |
| `feed(d, isIncremental)` | Merge another dict's values |
| `reset(v)` | Reset all values to a constant |
| `nulls(toVal)` | Handle null values |
| `isDict(obj)` / `isInstanceOf(obj, type)` | Type checking |

---

### `Lists.bas` — Dynamic Array List

A dynamically resizing array with functional programming support.

**Last updated:** 13.08.2024

**Core API:**

| Method/Property | Description |
|---|---|
| `add(ele)` | Append element; returns `Me` |
| `addAll(arr)` | Append all from array/collection |
| `addAt(ele, index)` | Insert at index |
| `remove(ele)` / `removeAt(index)` | Remove element |
| `getVal(index, index2)` | Get element (supports 2D indexing) |
| `setVal(index, ele)` | Set element at index |
| `length` | Current number of elements |
| `contains(ele)` | Membership test |
| `indexOf(ele)` | First index of element |
| `last()` | Last element |

**Functional operations:**

| Method | Description |
|---|---|
| `map(operation)` | Transform each element |
| `filter(judgement)` | Filter elements |
| `reduce(operation, initialVal)` | Reduce to single value |
| `reduceRight(operation, initialVal)` | Reduce from right |
| `every(judgement)` | True if all match |
| `some(judgement)` | True if any match |
| `mapX(callback)` | map with external VBA function |
| `filterX(callback)` | filter with external VBA function |
| `reduceX(callback, initVal)` | reduce with external VBA function |
| `sortX(callback)` | sort with external comparator |
| `product(op, list2)` | Element-wise binary operation |
| `mapList(op, reduceOp)` | Map each sub-list then reduce |

**Slice/Reshape:**

| Method | Description |
|---|---|
| `slice(from, to, step)` | Python-like slicing (negative indices supported) |
| `sliceBy(arr)` | Slice using index array |
| `subList(from, to)` | Sublist by range |
| `subgroupBy(l, offset)` | Group consecutive elements |
| `take(n)` / `drop(n)` / `dropLast(n)` | Head/tail operations |
| `flatten()` | Flatten nested lists |
| `zip(...)` | Zip multiple arrays |
| `zipMe()` | Transpose a list of lists |
| `permutation()` | All permutations |
| `reverse()` | In-place reverse |
| `sort(isAscending)` | Sort |
| `unique()` | Remove duplicates |

**Conversion:**

| Method | Description |
|---|---|
| `toArray()` | Convert to Variant array |
| `toDict()` / `toDicts()` | Convert to Dicts (index → value) |
| `toMap(asArr)` | Convert to Dicts |
| `fromSerial(start, ending, steps)` | Create sequential list |
| `fromArray(arr)` | Load from array/Collection |
| `fromDict(d)` | Load from Dicts |
| `fromRng(rng, orientation)` | Load from Range |
| `load(rng)` / `loadSht(sht)` | Load from sheet |
| `toRng(rng)` | Write to range |
| `join(delimiter)` | Join as string |

**Stats:**

| Method | Description |
|---|---|
| `min_()` / `max_()` | Min/max |
| `avg()` | Average |

---

### `HashSets.bas` — Hash Set

A set backed by `Lists`. Used internally by `Dicts` for key management.

**Dependencies:** `Lists`

| Method | Description |
|---|---|
| `add(e, update)` | Add element |
| `contains(e)` | Membership test |
| `remove(e)` | Remove element |
| `ceiling(e, asNode)` | Smallest element >= e (delegates to TreeSets) |
| `clear()` | Clear all |
| `toString()` / `p()` | Print |
| `toArray()` | Export to array |

---

### `TreeSets.bas` — Sorted Set (BST)

A sorted set implemented as a Binary Search Tree. Supports ordered operations.

**Dependencies:** `Lists`, `Nodes`

| Method | Description |
|---|---|
| `add(e, update)` | Insert element |
| `remove(e)` | Delete element |
| `contains(e)` | Membership test |
| `ceiling(e, asNode)` | Smallest >= e |
| `floor(e, asNode)` | Largest <= e |
| `min_()` / `max_()` | Min/max elements |
| `toArray()` | In-order traversal as array |
| `toString()` / `p()` | Print |
| `clear()` | Clear tree |

---

### `Nodes.bas` — BST Node

Internal node class for `TreeSets`.

| Property | Description |
|---|---|
| `value` | Stored value |
| `leftNode` / `RightNode` | Child nodes |
| `index` | Position index |
| `init(l, r, i, val)` | Initialize node |

---

### `util.bas` — Workbook Iterator

A utility module (not a class) providing a template for batch-processing workbooks in a directory.

```vba
Public Sub processWorkbooksInthePath(Optional path As String = "src", Optional readOnly As Boolean = True)
```

Iterates all `.xls`, `.xlsm`, `.xlsx` files in a subfolder and calls the user-defined interface:
```vba
Sub interface_processWorkbook(byref wb as workbook, byref this as workbook)
```

---

### `xlDataStructure.py` — Python Companion

A Python `xlDict` class using `win32com` for COM automation of Excel (Windows only).

**Requires:** `pywin32` (`pip install pywin32`)

**Key methods:**

| Method | Description |
|---|---|
| `__init__(sht, keyCol, valCol, ...)` | Load from Excel sheet |
| `unload(sht, ...)` | Write values back to sheet |
| `dump(sht, ...)` | Write keys + values |
| `map(keyFun, valFun)` | Transform key/value |
| `toJSON()` | Serialize to JSON |
| `titledDict(title)` | Map values to named dict |
| `append(...)` | Merge another sheet's data |
| `reduced()` | Sum multi-column values |
| `simplify()` | Collapse single-key nested dicts |
| `__and__` / `__or__` / `__sub__` / `__xor__` | Set operations (intersection/union/diff/sym-diff) |

---

## Expression Syntax for Functional Operations

The inline functional operations (`map`, `filter`, `reduce`) use a **string-based expression mini-language**:

- `_` — placeholder for the current element/value
- `?` — placeholder for accumulator (reduce only)
- `{i}` — placeholder for current index
- `{1}` / `{2}` — left/right operands in `mergeMap`

Expressions are VBA `Evaluate`-compatible formulas. Examples:

```vba
' Double all values
dict.map("_*2")

' Keep values > 4
dict.filter("_>4")

' Sum values starting from 0
dict.reduce("_+?", 0)

' Keep keys > 8
dict.filterKey("_>8")

' NPV via rolling reduce, then filter results > 5
dict.ranged("?+_/1.1^({i}+1)", AggregateMethod.AggReduce).filter("_>5")
```

---

## External Callback Syntax

For complex operations not expressible as simple strings, use the `*X` variants with an external VBA function reference (as a string `"Module.FunctionName"`):

```vba
' Call Test.callback for each element
l.mapX("Test.callback")

' Call Test.f as filter predicate
l.filterX("Test.f")

' Call Test.r as reducer
l.reduceX("Test.r", New Dicts)
```

The callback function receives the current element (and accumulator for reduce) via the `callback` property of the calling object.

---

## Testing

Tests live in `Test.bas` and use VBA's built-in `Debug.Assert`. Run the `Test()` subroutine from the VBA IDE (F5 or Run → Run Sub).

**There is no automated test runner.** Tests must be executed manually within Excel.

Key test patterns:
```vba
Debug.Assert d.Count = 0           ' count check
Debug.Assert .reduce("_+?", -1) = 54   ' arithmetic reduce
Debug.Assert .filterKey("_>8").Count = 2  ' filtered count
Debug.Assert l.fromSerial(10, 15).mapX("Test.callback").slice(-1).getVal(0) = "225_"
```

---

## Development Conventions

### VBA Coding Style
- `Option Explicit` is required in all modules.
- All classes use `Class_Initialize` / `Class_Terminate` for setup/teardown.
- Private fields use `p` prefix (e.g., `pKeys`, `pLen`, `pArr`).
- Methods return `Me` (i.e., `Set method = Me`) for **method chaining**.
- `Nothing` checks use `isInstanceOf(obj, "Nothing")` or `isNothing(obj)`.
- Decimal point replacement: numeric string expressions use `.` → `,` conversion for locale compatibility (`replaceDecimalPoint` parameter).

### Method Chaining Pattern
```vba
Dim d As New Dicts
d.load("Sheet1", 1, 2).filter("_>100").map("_*1.1").p
```

### Null / Empty Handling
- Missing values return `Null` (not `Empty` or `Nothing`).
- Use `.nulls(defaultVal)` to replace nulls in a dict.
- Use `nullVal(setValTo)` in Lists.

### Cross-Platform (macOS)
- No `Scripting.Dictionary` (Windows-only ActiveX) is used — all data structures are pure VBA.
- `util.bas` uses `scripting.filesystemobject` which is Windows-only; avoid in macOS-targeted code.

### File Naming
- `.bas` extension = VBA module/class file.
- Class modules that represent objects (Dicts, Lists, etc.) use PascalCase naming.
- The `Test.bas` module name prefix is used in `mapX`/`filterX` callback strings (e.g., `"Test.callback"`).

---

## Import Order

When importing into a new workbook, respect dependency order:

1. `Nodes.bas`
2. `Lists.bas`
3. `HashSets.bas`
4. `TreeSets.bas`
5. `Dicts.bas`
6. `util.bas` (optional)
7. `Test.bas` (optional, development only)

---

## Common Patterns

### Load spreadsheet data and filter
```vba
Dim d As New Dicts
With d.load("Sheet1", 1, 2)          ' keyCol=1, valCol=2
    .filter("_>100").p               ' print entries where value > 100
End With
```

### Build a dict from an array
```vba
Dim d As New Dicts
Set d = d.fromArray(Array(10, 20, 30))
' keys: 1,2,3 → values: 10,20,30
```

### Functional pipeline
```vba
Dim l As New Lists
l.fromSerial(1, 10).map("_*2").filter("_>10").reduce("_+?", 0)
' = sum of doubled values > 10 from 1..10
```

### GroupBy with aggregation
```vba
' Group by column 1, aggregate column 2 with xlSum
d.groupBy(1, 2, xlSum).p
```

### JSON round-trip
```vba
Dim json As String
json = d.toJSON()

Dim d2 As New Dicts
Set d2 = d2.fromString(json)
```

### Labeled DataFrame access
```vba
d.label = Array("id", "name", "score")
Debug.Print d.getByLabel(3, "score")  ' score for key 3
```

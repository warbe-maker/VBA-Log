## VBA-Service-Log
### Summary
Provides comprehensive methods/services and properties for writing a log file either with already formatted entries or column oriented entries.

### Usage
### Simple
Writing log entries is as simple as possible thanks to sensible defaults
```vb
    Dim Log As New clsLog
    Log.Entry "This is a log entry line"
```
Writes a single log entry to the default log file (see the properties ***FileFullName***, ***FileName***, and ***Path***).
```vb
    Dim Log As New clsLog
    Log.ColsDelimiter = " " ' defaults to |
    Log.Entry "xxxxxxxxxx", "yyyyyyyyyyyyyyyyyyyy", "zzzzzzzz"
    Log.Entry "xxxxxxxx", "yyyyyyyy", "zzzzzzzzz"
```
Writes two log entries **left aligned in columns**:<br>
```
Item-1     Item-2               Item-3
xxxxxxxxxx yyyyyyyyyyyyyyyyyyyy zzzzzzzz
xxx        yyyyyyy              zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
```
>When no ***Widths*** are explicitly specified the columns width is determined by the width of the first row's items, the width of the rightmost column is unlimited by default. 

### Methods
| Method Name         | Function |
|---------------------|----------|
|***Dsply***              | Displays the log-file by means of the application associated with the file's extension, which defaults to .log|
|***Entry***              | Specifies either a single string or a number of items written aligned in columns. For the latter see the [implicit column width ++and++ alignment specification](#implicit-column-width-and-alignment-specification).|
|***EntryItemsAlignment*** | ParamArray if strings, Explicitly specifies the alignment for each column's item, defaults to **L**eft adjusted when not provided.. <br>Example:<br>**"L","C","R","L"** col1=Left, col2=centered, col3=rigth. |
|***Headers***             | ParamArray of strings, specifying any number number of strings written as a single line ***column*** header, written automatically with the first call of the ***Entry*** method provided it specifies column aligned items. The alignment of the header string defaults to centered. An implicit specification of the alignment is possible by means of vertical bars [\|).<br>Examples:<br>- **"\| xx\|yy \|zz\| aa \|"** xx=right adjusted, yy=left adjusted, zz, aa=centered.<br>- **"xxx", "yyyy", "zzz"** are written centered provided not specified explicitly by means of the ***HeaderAlignment*** method. |
|***HeadersAlignment***    | Defaults to ***C***entered for each of the specified headers |
|***Title***               | ParamArray of strings, Specifies the - optionally multi-line - title of a new series of log entries. Triggers the writing of the column headers provided specified.<br>Examples:<br>- **"Any title"** will be centered,<br> - **"\| &nbsp;&nbsp;&nbsp;Any title"** will be left adjusted including all leading spaces.|
|***Widths***              | ParamArray of integer values, explicitly specifies [the columns (minimum) width](#column-width-specification).|

### Properties
| Name          | Description |
|---------------|-------------|
|***FileFullName*** | ReadWrite, string expression, specifies the full name of the log-file defaults to a file named like the `ActiveWorkbook` [^1] with an ".log" extension |
|***FileName***     | Specifies the log-files name, defaults to the  `ActiveWorkbook's` [^1] BaseName with an `.log` file extension. |
|***KeepDays***     | Specifies the number of days a new log-file is kept before it is deleted and re-created.|
|***LogFile***      | Expression representing a file object. |
|***Path***         | String expression, defaults to the `ActiveWorkbook's` [^1] parent folder. |
|***WithTimeStamp***| Boolean expression, defaults to true, when true each log line is prefixed with a time stamp in the format `yy-mm-dd-hh:mm:ss` |

### Installation
Download and open the dedicated development Workbook [VBLogService.xlsb][1] and in the VB-Editor copy (drag and drop) the clsLog Class-Module into your VB-Project. Alternatively copy the below to the clipboard and into a new Class-Module (throughout this README named ******clsLog******).
```
Still waiting for the final code version!
```
## Column alignment details
### Margin and delimiter
The columns margin and the columns delimiter are determined by the specification of column ***Headers***. When specified the columns delimiter defaults `|` (vertical bar) and the margin defaults to a  `" "` (single space). When no ***Headers*** were specified the columns delimiter defaults to a `" "` (single space and the margin defaults to a vbNullString.

### Column width specification
#### Explicit columns width specification
An explicit column ***Widths*** specification is regarded the **minimum** width. It may be expanded by the width of the corresponding column's ***Headers*** (when specified) and the width of the very first ***Entry*** line's items width (whereby the width of the rightmost column is unlimited.<br>Examples:
  - `|20|30|25|` specifies the minimum width for 3 columns
  - `20,30,25` same as above.
  - `,,,30` specifies the minimum width only for the 3rd column. For all the other columns the width is determined by the ***Headers*** (when specified) and the width of the very first ***Entry*** line's items width 
#### Implicit columns width specification
See [Implicit column width and alignment specification](#implicit-column-width-and-alignment-specification)

### Column alignment specification
#### Explicit column alignment specification
- For ***Headers***: ***HeadersAlignment*** method
- For ***Entry*** items: ***EntryItemsAlignment*** method

For each column the alignment is not explicitly specified by means of the corresponding method, the alignment follows the [Implicit column width and alignment specification](#implicit-column-width-and-alignment-specification). [^2]

Examples:
  - `|C|L|R|` specifies the alignments centered, left adjusted and right adjusted for the first 3 columns
  - `"C","L","R"` same as above.
  - >Note: Any string not beginning with C, L, or R defaults to **C**entered

#### Implicit column width and alignment specification 
For all columns the width and/or the alignment has not explicitly specified, both is derived from an implicit specification as follows. The specification may be a single vertical bar (|) delimited string or an array of string expressions.

| &nbsp;&nbsp;&nbsp;&nbsp;Example&nbsp;&nbsp;&nbsp;&nbsp; | Alignment | Width | Alignment Rule |
|---------------------|:----------:|:-----:|-------------|
| `|xxx.|`<br>`|.xxxx..|`<br>`"xxx."`<br>`".xxx.."` | left<br>left<br>left<br>left       | 5<br>7<br>5<br>7 | A number of trailing spaces greater than the number of leading spaces indicates **L**eft adjusted.|
| `|xxx|`<br>`|.xxx.|`| centered   | 4     | None or an equal number of leading and trailing spaces indicates **C**entered. |
| `|.xxx|`<br>`|....xxx..|`            | right<br>right      | 5<br>9     | A number of leading spaces less than the number of trailing spaces indicates **L**eft adjusted. |

>The calculated final width encloses the string in at least one leading a trailing [***ColsMargin***](#class-module-properties), which default to a single space and is supposed for the above examples. 

[^1]: When the `ActiveWorkbook` is used as the default for the log-file's location the log-file is located in the serviced Workbook's parent folder. When the service writing the log is for the Workbook itself `ThisWorkbook` and `ActiveWorkbook` are the same, when the service is provided by another Workbook for the  servicing Workbook will be `ThisWorkbook` and the serviced Workbook will be the `ActiveWorkbook`. In both cases the log-file written into the **serviced Workbook's** parent folder.
 

[1]: https://github.com/warbe-maker/VBA-Log-Service/blob/main/VBALogService.xlsb?raw=true
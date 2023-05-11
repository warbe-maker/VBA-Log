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

### Class Module Methods
| Method Name         | Function |
|---------------------|----------|
|***Dsply***              | Displays the log-file by means of the application associated with the file's extension, which defaults to .log|
|***Entry***              | ParamArray of strings. When only one string (without any vertical bars [\|) is specified, the string is written as specified, when multiple strings are specified the strings are written left aligned in columns, when a single string with vertical bars (\|) is specified the elements are written aligned in columns whereby the kind of alignment is implicitly specified by leading and trailing spaces.<br>Example: **"\| xx\|yy \|zz\| aa \|"** xx is right ajusted, yy is left adjusted, zz and aa are aligned centered.<br>Note: The alignment of the items is independent of the alignment of the column headers.|
|***EntryItemsAlignment*** | ParamArray if strings, Explicitly specifies the alignment for each column's item, defaults to **L**eft adjusted when not provided.. <br>Example:<br>**"L","C","R","L"** col1=Left, col2=centered, col3=rigth. |
|***Headers***             | ParamArray of strings, specifying any number number of strings written as a single line ***column*** header, written automatically with the first call of the ***Entry*** method provided it specifies column aligned items. The alignment of the header string defaults to centered. An implicit specification of the alignment is possible by means of vertical bars [\|).<br>Examples:<br>- **"\| xx\|yy \|zz\| aa \|"** xx=right adjusted, yy=left adjusted, zz, aa=centered.<br>- **"xxx", "yyyy", "zzz"** are written centered provided not specified explicitly by means of the ***HeaderAlignment*** method. |
|***HeadersAlignment***    | Defaults to ***C***entered for each of the specified headers |
|***Title***               | ParamArray of strings, Specifies the - optionally multi-line - title of a new series of log entries. Triggers the writing of the column headers provided specified.<br>Examples:<br>- **"Any title"** will be centered,<br> - **"\| &nbsp;&nbsp;&nbsp;Any title"** will be left adjusted including all leading spaces.|
|***Widths***              | ParamArray of integer values, explicitly specifies [the columns (minimum) width](#column-width-specification).|


### Class Module Properties
| Name          | Description |
|---------------|-------------|
|***ColsDelimiter***| String expression. defaults to a vertical bar (`|`), when set to a single space the ***ColsMargin*** is set to a `VBNullString`|
|***ColsMargin***   | String expression, defaults to a single space (` `), when set to a `VBNustring` the ***ColsDelimiter*** is set to a vertical bar (`|`) |
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
## Specifics
### Column width specification
- **Explicit specification**: When none is specified by means of the ***Widths*** method the width defaults to the maximum of the width of the corresponding column's header and the width of the very first entry's items width, whereby the rightmost columns width is unlimited by default. When no ***Headers*** had been specified the column width defaults to the first ***Entry***'s written item of the corresponding column. Examples:
  - `|20|30|25|` specifies the minimum width for 3 columns
  - `20,30,25` same as above.
- **Implicit specification**: See [Implicit column width and alignment specification](#implicit-column-width-and-alignment-specification)

### Column alignment specification
- **Explicit header alignment**: When none is specified by means of the ***HeadersAlignment*** method and none is specified implicitly by means of the ***Headers*** method, the alignment defaults to **C**entered. Examples:
  - `|C|L|R|` specifies the alignments centered, left adjusted and right adjusted for the first 3 columns
  - `"C","L","R"` same as above.
  - >Note: Any string not beginning with C, L, or R defaults to **C**entered
- **Explicit ***Entry*** items alignment**: When none is specified by means of the ***EntryItemsAlignment*** method and none is specified implicitly by means of the ***Entry*** method, the alignment defaults to **L**eft adjusted. Examples:
  - `|C|L|R|` specifies the alignments centered, left adjusted and right adjusted for the first 3 columns
  - `"C","L","R"` same as above.
  - >Note: Any string not beginning with C, L, or R defaults to **L**eft adjusted
- **Implicit ***Headers*** alignment**: See [Implicit column width and alignment specification](#implicit-column-width-and-alignment-specification)
- **Implicit ***Entry*** item alignment**: See [Implicit column width and alignment specification](#implicit-column-width-and-alignment-specification)


#### Implicit column width and alignment specification 
In general the implicit specification is done by means of vertical bars (|) indicating columns. Note: Spaces are indicated by dots (.).

| &nbsp;&nbsp;&nbsp;&nbsp;Example&nbsp;&nbsp;&nbsp;&nbsp; | Adjustment | Width | Common rule |
|---------------------|------------|-------|-------------|
| `|xxx.|`<br>`|.xxxx..|` | left<br>left       | 5<br>7     | **Alignment**: A number of trailing spaces greater than the number of leading spaces indicates **L**eft adjusted.|
| `|xxx|`<br>`|.xxx.|`| centered   | 4     | None or an equal number of leading and trailing spaces indicates **C**entered. |
| `|.xxx|`<br>`|....xxx..|`            | right<br>right      | 5<br>9     | A number of leading spaces less than the number of trailing spaces indicates **L**eft adjusted. |

>The calculated final width encloses the string in at least one leading a trailing [***ColsMargin***](#class-module-properties), which default to a single space and is supposed for the above examples. 

[^1]: When the `ActiveWorkbook` is used as the default for the log-file's location the log-file is located in the serviced Workbook's parent folder. When the service writing the log is for the Workbook itself `ThisWorkbook` and `ActiveWorkbook` are the same, when the service is provided by another Workbook for the  servicing Workbook will be `ThisWorkbook` and the serviced Workbook will be the `ActiveWorkbook`. In both cases the log-file written into the **serviced Workbook's** parent folder.

[1]: https://github.com/warbe-maker/VBA-Log-Service/blob/main/VBALogService.xlsb?raw=true
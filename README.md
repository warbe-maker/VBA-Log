## VBA-Service-Log
### Usage examples
Writing log entries is as simple as possible thanks to sensible defaults:
#### Example 1
```vb
    Dim Log As New clsLog
    Log.Entry "This is a log entry line"
```
Writes a single log entry to the default log file (see the [properties](#properties) ***FileFullName***, ***FileName***, and ***Path***).
#### Example 2
```vb
    Dim Log As New clsLog
    Log.Entry "xxxxxxxxxx ", "yyyyyyyyyyyyyyyyyyyy ", "zzzzzzzz " 
    Log.Entry "xxx", "yyyyyyy", "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz"
```
Writes two log entries aligned in columns, with the alignment and the column width [implicitly](#implicit-column-alignment-specification) specified:
```
=====================================================================
xxxxxxxxxx yyyyyyyyyyyyyyyyyyyy zzzzzzzz
xxx        yyyyyyy              zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
```
Note:
- A delimiter line is automatically inserted when the new series of log entries is not the first in the log-file
- When no [***MaxItemLengths***](#methods) is specified the columns width is determined by the width of the first row's items whereby the width of the rightmost column is unlimited by default. 

### Methods
| Method Name         | Function |
|---------------------|----------|
|***AlignmentHeaders***    | ParamArray of string expressions, specifies a header line with column headers, may be repeated for multiple column headers, optionally implicitly specifies the [alignment](#explicit-column-alignment-specification) for each column's item.<br>Example:<br>**"L","C","R","L"** col1=Left, col2=centered, col3=rigth. |
|***AlignmentItems*** | ParamArray of string expressions, explicitly specifies the [alignment](#explicit-column-alignment-specification) for each column's item. <br>Example:<br>**"L","C","R","L"** col1=Left, col2=centered, col3=rigth. |
|***Dsply***              | Displays the log-file by means of the application associated with the file's extension, which defaults to .log|
|***Entry***              | Specifies either a single string or a number of items written aligned in columns. For the latter see the [implicit column width ++and++ alignment specification](#implicit-column-alignment-specification).|
|***Headers***             | ParamArray of string expressions, specifies a header line with column headers, may be repeated for multiple column headers, optionally implicitly specifies the column headers' alignment |
|***MaxItemLengths***        | ParamArray of integer values, specifies the [maximum lenght of items aligned in columns](#column-width).|
|***NewLog***              | Explicit indication that the next ***Entry*** writes the first of a new series of log entries. I.e. with the next ***Entry*** a delimiter line (======) is written - provided its not a new log file. In most cases this method is unnecessary because the begin of a new series of log entries is implicitly considered with the ***Title*** method, the ***Headers****method and in case the ***Entry*** method changes from a single string to column aligned items or vice versa.|
|***Title***               | ParamArray of strings, Specifies the - optionally multi-line - title of a new series of log entries. Triggers the writing of the column headers provided specified.<br>Examples:<br>- **"Any title"** will be centered,<br> - **"\| &nbsp;&nbsp;&nbsp;Any title"** will be left adjusted including all leading spaces.|

### Properties
| Name          | Description |
|---------------|-------------|
|***ColsDelimiter***| Defaults to a vertical bar (\|) when ***Headers*** are specified and defaults to a single space when no ***Headers*** are specified. |
|***FileFullName*** | ReadWrite, string expression, specifies the full name of the log-file defaults to a file named like the `ActiveWorkbook`[^1] with an ".log" extension |
|***FileName***     | Specifies the log-files name, defaults to the  `ActiveWorkbook's` [^1] BaseName with an `.log` file extension. |
|***KeepDays***     | Specifies the number of days a new log-file is kept before it is deleted and re-created.|
|***LogFile***      | Expression representing a file object. |
|***Path***         | String expression, defaults to the `ActiveWorkbook's` [^1] parent folder. |
|***WithTimeStamp***| Boolean expression, defaults to `True`. When `True` each log line is prefixed with a time stamp in the format `yy-mm-dd-hh:mm:ss` |

### Installation
Download and open the dedicated development Workbook [VBLogService.xlsb][1] and in the VB-Editor copy (drag and drop) the clsLog Class-Module into your VB-Project. [^2]

## Column alignment details
### The Columns Delimiter
When ***Headers*** are specified the columns delimiter defaults to a  `|` (vertical bar), else to a single space.
### The Columns Margin
When the [columns delimiter](#the-columns-delimiter) is a `|` (vertical bar) the margin defaults to a single space, when it is a single space the margin is a `vbNullString`.

### Column Width
The column width is the space between two [column delimiters] which may be a `|` (vertical bar) or a single space. The final width of a column considers:
 - the ***MaxItemLengths*** (when specified for the column)
 - a leading and trailing column margin which depends on the ***ColsDelimiter***
 - the width of the columns ***Headers*** (when specified)
 - the length of the first ***Entry*** item.
 
 Examples with a ***MaxItemLengths*** = 30 explicitly specified for a column's items:

| Example | Conditions | Final Cols Width |
| --------|------------|------------------|
|<nobr>`| <--max-item-length-column-n--> |`| The cols delimiter is a `|` (vertical bar), the case by default when ***Headers*** are specified of when ***ColsDelimiter*** explicitly specifies it. In both cases the columns margin is a single leading and trailing space.| 32 |
|<nobr>` <--max-item-length-column-n--> `| The columns delimiter is a single space, either by default when no ***Headers*** are specified or when explicitly specified by the ***ColsDelimiter***| 30 |
|<nobr>` <--max-item-length-column-n--> .: `| - The columns delimiter is a "&nbsp;" (single space), either by default when no ***Headers*** are specified or when explicitly specified by the ***ColsDelimiter***<br>- The ***AlignmentIems*** specified `"L."` for the column indicating fill with `.`(dots) terminated by a `:`(colon) | 33 |

#### The maximum items' length specification
The explicit specification of the maximum items' length ensures enough column space in case neither the header nor the first ***Entry***'s items specifies the length/width implicitly. Since this usually is unlikely the case an explicit maximum length specification is rather common.
The specified [***MaxItemLengths***](#methods) becomes the minimum width of the corresponding column. This width may still be expanded by the width of the corresponding column's ***Headers*** (when specified and greater).<br>Examples:
  - `|20|30|25|` specifies the maximum item length for column 1, 2, and 3
  - `20,30,25` same as above.
  - `,,,30` specifies the maximum item length only for the 3rd column. For all the other columns the width is determined by the ***Headers*** (when specified) and the width of the very first ***Entry*** line's items width

#### Implicit columns width specification

See [Implicit column width and alignment specification specification](#implicit-column-alignment-specification)

### Alignment specification
Columns alignment may be specified
- **explicitly**
  - for [headers](#explicit-headers-alignment-specification): by the ***AlignmentHeaders*** method
  - for ***Entry*** [items](#explicit-items-alignment-specification) by the ***AlignmentItems*** method
- **implicitly**
  - for [headers](#implicit-headers-alignment-specification) by the ***Headers*** method
  - for items by the ***Entry*** method
  - for log title lines by the ***Title*** method.

#### Explicit headers alignment specification
```vbs
    .AlignmentHeaders "|C|L|R|"     ' centered, left adjusted and right adjusted
    .AlignmentHeaders "C","L","R"   ' same as above
    .AlignmentHeaders "|x|x | x|"   ' same as above, follows the same rules as the implicit alignment soec 
```
For each header column the alignment is not explicitly specified by means of the corresponding method, the alignment follows the [Implicit column width and alignment specification](#implicit-column-alignment-specification). [^2]

##### Explicit items alignment specification
```vb
    ' 1: centered, 2: left adjusted, 3: right adjusted
    .AlignmentItems "|C|L|R|"     
    ' 1: centered, 2: left adjusted filled with . (dots), 3: right adjusted
    .AlignmentItems "|C|L.|R|"
    ' 1: centered, 2: left adjusted filled with . (dots) teminated by a : (colon), 3: right adjusted
    .AlignmentItems "|C|L.:|R|"   
    
    ' same as above but with individual (ParamArray) strings
    .AlignmentItems "C","L","R"
    .AlignmentItems "C","L.","R"
    .AlignmentItems "C","L.:","R"
    
    ' same as above, follows the same rules as the implicit alignment spec 
    .AlignmentItems "|x|x | x|"   
    .AlignmentItems "|x|x.| x|"     ' xxxxx .........  
    .AlignmentItems "|x|x.:| x|"    ' xxxxx ........:    
```

#### Implicit headers alignment specification
For headers the implicit alignment specification will be the common method since the header is a fixed string. Example of a 3 line header: Note that the alignment is specified by the first line only and subsequent lines are aligned accordingly.
```vb
    .Headers "|  Column  | Column  |  Column |"
    .Headers "|     1    |   2     |    3    |"
    .Headers "|(centered)| (left)  | (right) |"
```

#### Implicit alignment specification rules
| Alignment      | Rule |
|----------------|------|
| Left adjusted  | 1. The number of leading spaces is less than the number of trailing spaces.<br>2. Leading  spaces are preserved. |
| Centered       | 1. The number of leading and trailing spaces is equal (may be 0)<br>2. Leading and trailing spaces are dropped. |
| Right adjusted | 1. The number of trailing spaces is less than the number of leading spaces.<br>2. Trailing spaces are preserved.

### The columns margin (depending on the columns delimiter)

| Columns Delimiter        | Columns Margin              | Comment |
|--------------------------|-----------------------------|---------|
|<nobr>`"|"` (vertical bar)|<nobr>" " (single space)     | Default when ***Headers*** are specified. The final column width will thus add two spaces.  |
| " " (single space)       |<nobr> "" (`vbNullString`)   | Default when no ***Headers*** are specified. The final width will be the maximum of the minimum width specifed expanded in case the ***Headers*** or the first ***Entry*** items occupy more space. |

[^1]: When the `ActiveWorkbook` is used as the default for the log-file's location the log-file is located in the serviced Workbook's parent folder. When the service writing the log is for the Workbook itself `ThisWorkbook` and `ActiveWorkbook` are the same, when the service is provided by another Workbook for the  servicing Workbook will be `ThisWorkbook` and the serviced Workbook will be the `ActiveWorkbook`. In both cases the log-file written into the **serviced Workbook's** parent folder.
 
[^2]: The Workbook (its dedicated parent folder respectively) is dedicated to the Class-Module's development and test and provides a full regression test which compares the result of a series of test with a file containing the expected results.

[1]: https://github.com/warbe-maker/VBA-Log-Service/blob/main/VBALogService.xlsb?raw=true
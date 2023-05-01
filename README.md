## VBA-Service-Log
### Summary
Provides comprehensive methods/services and properties for writing a log file either with already formatted entries or column oriented entries.
### Installation

### Usage

### Methods
| Method Name | Function |
|-------------|----------|
| Entry       | String expression, written to the log-file |
| Items       | Writes any number of strings to the log-file aligned in columns |
| ColsHeader  | Specifies any number of strings written as column header, either automatically with the first call of the method _Items_ of explicitly with the mehtod _WriteHeader_.  |
| ColsAlignmentHeader | Defaults to ***C***entered for each of the specified headers |
| ColsAlignmentItems | Defaults to ***L***eft adjusted for each of the items provided with the _Items_ method. |
| | |
| WriteHeader | Writes the specified header to the log-file |
| ColsWidth | ParamArray of column width, defaults to the width of the corresponding column header, when no column headers are provided, the width is determined by the width of the items written in the first log line. |
| Dsply       | Displays the log-file |


### Properties
| Name          | Description |
|---------------|-------------|
| FileFullName  | ReadWrite, string expression, specifies the full name of the log-file defaults to a file named like the `ActiveWorkbook` with an ".log" extension |
| FileName      | |
| KeepDays      | |
| LogFile       | Expression representing a file object. |
| Path          | String expression, defaults to the `ActiveWorkbook's` parent folder. |
| Title         | String expression, specifying the title printed above the log-lines. |
| WithTimeStamp | Boolean expression, defaults to true, when true each log line is prefixed with a time stamp in the format `yy-mm-dd-hh:mm:ss` |
| ColsMargin    | Default to a single space, printed left and right of the vertical bar \| used to spearate the columns _ColDelimiter_ |
| ColsDelimiter | Defaults to a vertical bar (\|), may be replaced e.b. by a space. |

 

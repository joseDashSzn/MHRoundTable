# MHRoundTable
Walkthrough GAS

# Menu Health Google Apps Script
## Send Items Custom Function
### Function Summary
Will be used to send info on restricted items to the appropriate sheet for specialists to handle.


## Setting Global Variables
```Javascript
let ss = SpreadsheetApp.getActive();
let sheet1 = ss.getSheetByName('Sheet1');
let sheet2 = ss.getSheetByName('Sheet2');
```
* We declare the `ss`, `sheet1`, `sheet2` variables outside of our function so that they can be used by other functions in the future. This is called Scope.
    * More info on [Javascript Scope](https://www.w3schools.com/js/js_scope.asp)
    * `Let` and `Var` are also tied to Scope. You can read more about that in the link above or [practice here](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/explore-differences-between-the-var-and-let-keywords).
    * [Global Scope](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/global-scope-and-functions),[Local Scope](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/local-scope-and-functions),[Global vs Local](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/global-vs--local-scope-in-functions)

* `ss` is the Alias we have given our currently active Spreadsheet 
  * Whenever we reference `ss` just think Spreadsheet
  * To look at the different `Properties` and `Methods` (commands) available, be sure to look at the [Developers Docs here](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#methods)!
  * Yes, there technically is a difference between a Spreadsheet and a Sheet. 

## What is a Property and what is a Method? 
```javascript
let stringFirstName = "Dumbledore" // we've just declared a variable (stringFirstName) and set it equal to the string "Dumbledore"
let arrayOfNames = ["Jack","Jill", "Snow White"]
```
* Property - Returns metadata about the object
  * `stringFirstName.length` will return 10
    * [String Length Property](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/length)  
  * `arrayOfNames.length` will return 3
    * [Array Length Property](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/length)
  * Notice how the `.length` property will return the length of the object. It returns info on the object
* Method - Functions (can interact with the data)
  * `stringFirstName.toUpperCase()` shall return a *new* string with all uppercase letters (`DUMBLEDORE`)
    * [String to Upper Case Method](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/toUpperCase)
  * `arrayOfNames.pop()` shall return a new array with the last Element removed
    * [Array Pop Method](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/pop)
## Creating Our Function
```Javascript
function onEdit(e) {
  
}
```
* [onEdit(e)](https://developers.google.com/apps-script/guides/triggers#onedite) is a built in Google Apps Script trigger that is triggered only when a *user* changes the value of a cell in the spreadsheet. Scripts CANNOT trigger this event, must be a user. 
* We could technically use the event object (`e`) to reference the range and cells but can circle back to that at a later date

## Working Down from Spreadsheet Level
```Javascript
function onEdit(e) {
let activeCell = ss.getActiveCell();
}
```
* After looking over the [Methods available for the Spreadsheet Object](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#methods), I've decided to use the [.getCurrentCell()](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getcurrentcell) method to return the currently highlighted cell and save it under as the variable `activeCell`. After some reading, it seems this method will `return` a `range object`
  * This means we won't be able to use any Spreadsheet Methods on our `activeCell` variable, we'll have to look at the [Methods available for the Range Objects](https://developers.google.com/apps-script/reference/spreadsheet/range). 
  * Note: If you run `console.log(activeCell)`, you will see an object filled with all the different methods and properties we can use with `activeCell`
  ```javascript
  console.log(activeCell)
  /* returns
  { 
  toString: [Function],
  setNumberFormat: [Function],
  getNumberFormat: [Function],
  setComment: [Function],
  getComment: [Function],
  getRow: [Function],
  getColumn: [Function],
  clear: [Function],
  getValue: [Function]
  } */
  ```
```javascript
function onEdit(e) {
let activeCell = ss.getActiveCell();
let activeRow = activeCell.getRow();
let currentColumn = activeCell.getColumn();
let tabValidation = 'Sheet1'; //This will be the name of the only sheet we want the script to work on.

//keyword that is used in SendItems is trimmed and set to lower case
let keyword = activeCell.getValue().trim().toLowerCase()

}
```
* I'm using the `getRow` and `getColumn` methods as they return the row number and row column that we are on. 
* Column Position will be useful later to make sure we're working on the correct columns
* We'll also use the `getValue` method to pull the data that is `inside` the cell instead of a reference to the cell object. (Yes, Cell Objects are a thing.)
  * We then chain the string methods `.trim()` and `.toLowerCase()` just in case, looking to remove unintentional spaces or caps lock
  * Makes it easier to compare because in JavaScript `"cbd"` and `"CBD"` are not equal

```javascript
//entire data range of the Sheet 1. (A1:AQ27)
let sheet1Range = sheet1.getDataRange();
  
// Object with Numbers pertaining to sheet headers on Spreadsheet. 1 = Col A, 2 = Col B, 4 = Col F etc.
const sheetHeaders = {
  // Numbered Column Positions
  // Business Name, Store Name, Store ID, Order Protocol, Is Partner, AO Emails,DM Contact,Error Category,Menu ID
  cbd: [1,2,4,6,7,10,11,12,16]
}
  // getSheet Headers depending on Keyword
  let arr = sheetHeaders[keyword];
```
* `.getDataRange` is a [Sheet Method](https://developers.google.com/apps-script/reference/spreadsheet/sheet) (not spreadsheet), that returns a range. We will save this range object as `sheet1Range`. It will only return the range with data *present* and for our data that range is `(A1:AQ27)`. 
  * By default a spreadsheet has 1000 rows and 26 columns. If we only fill in 10 rows and 3 Columns, `.getDataRange` will only specify the range with data. i.e (Range will be A1:C10)
  * `.getDataRange` is incredibly helpful for dynamic data. The alternative `.getRange()` requires the range to be specified.
  * If we wanted to specify the same range `A1:C10` with .getRange() it would look like so `sheet1.getRange(1,1,10,3)` (start from the first row, first column, grab 10 rows, grab 3 columns)

* `sheetHeaders` is an Object we've declared. 
  * cbd is a property, and I've set it equal to an array of numbers
    * The numbers in the array point to the columns we are grabbing
    * sheetHeaders.cbd or sheetHeaders[cbd] or sheetHeaders["cbd"] are all valid ways of accessing the property
  * To learn more about [Javascript Objects](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Working_with_Objects)
  * To [Practice with more Javascript Objects](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/build-javascript-objects)

* `arr` is then set equal to one of the array of numbers we've created in sheetHeaders. 
  * Depends entirely on what keyword is passed on. If the keyword does not exist in the Object then nothing will happen

``` javascript
// empty array to push data into
  let rowData = [];

  // pushing data to empty the array (rowData)
  arr.forEach(el => 
  rowData.push(sheet1Range.getCell(activeRow,el).getValue())
  )
```
* `.forEach` is an [Array Method](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach). It will loop over every single element in an array and perform an action or a function on the elements. 
* We are running `.forEach` on the `arr` variable (`arr.forEach()`). The array within this variable contains the Column Positions we numbered earlier 
* `rowData.push(sheet1Range.getCell(activeRow,el).getValue())`
  * `rowData` is an ***EMPTY array*** we declared earlier and will be the destination for the data we grab
  * For more [Practice with Javascript Arrays](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/store-multiple-values-in-one-variable-using-javascript-arrays)
  * `.push()` is an [Array Method](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/push), that will push new elements into the array it is called on (we will push new elements into `rowData`)
    * `sheet1Range` was the range we declared earlier. Normally dynamic, but for our test data the range it grabs is `(A1:AQ27)`
    * We then use our current `activeRow` and the `Column Positions` from earlier to locate the cells using `.getCell()` 
      * `sheet1Range.getCell(2,2)` would grab the cell in the second row, second column (row 2, store_name)
      * [.getCell](https://developers.google.com/apps-script/reference/spreadsheet/range#getcellrow,-column) returns a range, so we'll need to use `.getValue()` to grab the actual data inside the cell. We then push this data to our empty array `rowData`.
        * `rowData` will look like this if, on the second row
          * `[Cole Street Market, Cole Street Market, 1159291, Non POS, TRUE,	, colestmarket325@gmail.com, Other, 1697471]` 

```javascript
let rowDataValues = new Array(rowData)
```
* Now that we've built an array with the necessary data (`rowData`), we will need to adjust its format slightly so that we can paste it to a different range. 
* Currently, `rowData` is a single array or 1-Dimensional. 
  * `[Cole Street Market, Cole Street Market, 1159291, Non POS, TRUE,	, colestmarket325@gmail.com, Other, 1697471]` 
* In order for us to paste the data to a different range, we'll have to transform our data into a 2D or Nested Array. An array within an array. 
    * [Practice with 2d Arrays](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/nest-one-array-within-another-array)
    * [More 2d Array Practice](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/access-multi-dimensional-arrays-with-indexes)
    * A 2d Array is needed to paste data into multiple cells, in case multiple rows are required. 
    * If pasting data to a single cell, 2d array not needed 
    ```javascript
    // 2d array example
    [
        ['Cole Street Market', 'Cole Street Market', 1159291, "Non POS", TRUE,	, "generic@gmail.com", "Other", 1697471]
    ]
    ``` 

* We create our 2d array, by passing our current array through an array constructor function.
    * `let rowDataValues = new Array(rowData)`
    * [New Array Constructor Function](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/Array)

```javascript
let sheet2LastRow = sheet2.getDataRange().getLastRow();
let sheet2Range = sheet2.getRange(sheet2LastRow+1,1,1,arr.length) 
//{starting positions}: one past the last row, starting at the first column, only one row, same amount of columns as length of the Headers Array
```

* `sheet2LastRow` is a Range Method. It returns a number, that number is the last row with data present. 
    * If row 1 has data, but row 2 is empty, it will return 1. 
* We combine the last row from Sheet2 with .getRange() in order to specify the empty row using `sheet2LastRow + 1`, as our first parameter for `.getRange()`
* the second parameter specifies we start at the first column, 
* The third parameter indicate we will only be focusing on 1 row, 
* And the fourth paratmeter indicates how many rows we will be focusing on. In this case the number of columns we are focusing on is set to the length of our `arr` variable we created earlier. 
    * We need to make sure the Number of Rows and Number of Columns matches if we want to paste our data. Its easiest to do that if we use the `.length` properties to our advantage. That way the columns we focus on will always match the number of columns we specified.

```javascript
if(keyword == "cbd" && currentColumn == 26 && activeRow > 1 && ss.getSheetName() == tabValidation){
    sheet2Range.setValues(rowDataValues); 
  };
```
* Finally, if the keyword matches `cbd`, and the `currentColumn` is Col Z (26), and the Active Row is greater than 1, and the name of our sheet matches the tabValidation variable we created earlier, then it will paste the ddata into Sheet2. 
* [Practice with Conditional Logic](https://www.freecodecamp.org/learn/javascript-algorithms-and-data-structures/basic-javascript/use-conditional-logic-with-if-statements)

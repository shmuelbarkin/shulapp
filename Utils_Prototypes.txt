function isNumeric_(n) 
{
  return !isNaN(parseFloat(n)) && isFinite(n);
} // isNumeric_()



// http://stackoverflow.com/a/29329794
function isDate_(myDate) 
{
  return myDate.constructor.toString().indexOf("Date") > -1;
} // isDate_()



// Add the specified number of days to the given date
Date.prototype.addDays = function(days) {
  var dat = new Date(this.valueOf())
  dat.setDate(dat.getDate() + days);
  return dat;
} // addDays()



// Returns true if string starts with suffix, and false otherwise
String.prototype.endsWith = function(suffix)
{
  return (this.match(suffix + '$') == suffix);
};



// Returns true if string contains sub_string, and false otherwise
String.prototype.contains = function(sub_string)
{
  return (this.indexOf(sub_string) > -1);
};



Array.prototype.unique = function(field)
{
  var
    o = {},
    r = [];
  
  // Check if field has been filled up
  if (field)
  {
    // Treat the array as 2D
    for (var i = 0; i < this.length; i += 1) o[this[i][field]] = this[i][field];
  }
  else
  {
    // Treat the array as 1D
    for (var i = 0; i < this.length; i += 1) o[this[i]] = this[i];
  } // if field
  
  for (var j in o) r.push(o[j]);
  return r;
}; // unique()



// Searches within arrayToSearch for the first match of valueToSearch and returns that row
// Assumes that arrayToSearch is either:
  // an array of objects (when propertyToSearch is a string)
  // a rectangular 2D array (when propertyToSearch is a number)
function findFirstMatch_(arrayToSearch, propertyToSearch, valueToSearch)
{
  // Check if required parameters are valid
  if (
    // Array
    (arrayToSearch) && (arrayToSearch.length)
    // Property
    && (
      // String and not empty
      ((typeof propertyToSearch === 'string') && (propertyToSearch.length > 0))
      // Number and non-negative
      || ((typeof propertyToSearch === 'number') && (propertyToSearch >= 0))
    )
  )
  {
    // Go through all elements in the array
    for (var i = 0, iLen = arrayToSearch.length; i < iLen; i++)
    {
      // Check if match
      if (arrayToSearch[i][propertyToSearch] == valueToSearch)
      {
        return arrayToSearch[i];
      } // if match
    } // for all array elements i    
  } // if valid parameters
  
  return {};  
} // findFirstMatch_()



// Searches within arrayToSearch for the first match of valueToSearch and returns the index of that row
function findFirstIndex_(arrayToSearch, propertyToSearch, valueToSearch)
{
  // Check if required parameters are valid
  if (
    // Array
    (arrayToSearch) && (arrayToSearch.length)
    // Property
    && (
      // String and not empty
      ((typeof propertyToSearch === 'string') && (propertyToSearch.length > 0))
      // Number and non-negative
      || ((typeof propertyToSearch === 'number') && (propertyToSearch >= 0))
    )
  )
  {
    // Go through all elements in the array
    for (var i = 0, iLen = arrayToSearch.length; i < iLen; i++)
    {
      // Check if match
      if (arrayToSearch[i][propertyToSearch] === valueToSearch)
      {
        return i;
      } // if match
    } // for all array elements i    
  } // if valid parameters
  
  return -1;  
} // findFirstIndex_()
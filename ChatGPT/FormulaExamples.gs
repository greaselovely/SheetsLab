// FormulaExamples.gs
// SheetsLab - Advanced formula demonstrations

/**
 * Computes the moving average for a range of numbers.
 *
 * @param {range} dataRange - A range of numerical data.
 * @param {number} period - The number of data points to average.
 * @return {Array} A column of moving averages.
 * @customfunction
 */
function MOVINGAVERAGE(dataRange, period) {
    if (period <= 0) return [];
    var data = dataRange.flat();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      if (i < period - 1) {
        result.push(''); // Not enough data for the period
      } else {
        var sum = 0;
        for (var j = i - period + 1; j <= i; j++) {
          sum += data[j];
        }
        result.push(sum / period);
      }
    }
    return result;
  }
  
  /**
   * Performs a weighted sum on two ranges: values and corresponding weights.
   *
   * @param {range} valuesRange - A range of numbers representing values.
   * @param {range} weightsRange - A range of numbers representing weights.
   * @return {number} The weighted sum.
   * @customfunction
   */
  function WEIGHTEDSUM(valuesRange, weightsRange) {
    var values = valuesRange.flat();
    var weights = weightsRange.flat();
    if (values.length !== weights.length) {
      throw new Error("Values and weights ranges must be of the same length.");
    }
    var total = 0;
    for (var i = 0; i < values.length; i++) {
      total += values[i] * weights[i];
    }
    return total;
  }
  
  /**
   * Generates a dynamic two-dimensional array filled with incremental numbers.
   *
   * @param {number} rows - Number of rows for the output array.
   * @param {number} cols - Number of columns for the output array.
   * @return {Array} A two-dimensional array with incremental numbers.
   * @customfunction
   */
  function GENERATEARRAY(rows, cols) {
    var result = [];
    var counter = 1;
    for (var i = 0; i < rows; i++) {
      var row = [];
      for (var j = 0; j < cols; j++) {
        row.push(counter++);
      }
      result.push(row);
    }
    return result;
  }
  
  /**
   * Sums two ranges element-wise, demonstrating array formula capabilities.
   *
   * @param {range} range1 - The first range of numbers.
   * @param {range} range2 - The second range of numbers.
   * @return {Array} An array containing the element-wise sums.
   * @customfunction
   */
  function ELEMENTWISEADD(range1, range2) {
    var arr1 = range1.flat();
    var arr2 = range2.flat();
    if (arr1.length !== arr2.length) {
      throw new Error("Input ranges must have the same number of elements.");
    }
    var result = [];
    for (var i = 0; i < arr1.length; i++) {
      result.push(arr1[i] + arr2[i]);
    }
    return result;
  }
  
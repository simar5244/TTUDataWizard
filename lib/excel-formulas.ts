// Comprehensive list of Excel formulas supported
export interface ExcelFormula {
  name: string;
  syntax: string;
  description: string;
  category: 'Math' | 'Text' | 'Logic' | 'Date' | 'Lookup' | 'Statistical' | 'Financial';
  examples: string[];
}

export const EXCEL_FORMULAS: ExcelFormula[] = [
  // Math
  { name: 'SUM', syntax: 'SUM(A,B,...)', description: 'Add all numbers', category: 'Math', examples: ['SUM(A,B)', 'SUM(A,B,C)'] },
  { name: 'AVERAGE', syntax: 'AVERAGE(A,B,...)', description: 'Average of numbers', category: 'Math', examples: ['AVERAGE(A,B)'] },
  { name: 'MAX', syntax: 'MAX(A,B,...)', description: 'Largest value', category: 'Math', examples: ['MAX(A,B,C)'] },
  { name: 'MIN', syntax: 'MIN(A,B,...)', description: 'Smallest value', category: 'Math', examples: ['MIN(A,B,C)'] },
  { name: 'ABS', syntax: 'ABS(A)', description: 'Absolute value', category: 'Math', examples: ['ABS(A)'] },
  { name: 'ROUND', syntax: 'ROUND(A,decimals)', description: 'Round to decimals', category: 'Math', examples: ['ROUND(A,2)', 'ROUND(A,0)'] },
  { name: 'ROUNDUP', syntax: 'ROUNDUP(A,decimals)', description: 'Round up', category: 'Math', examples: ['ROUNDUP(A,0)'] },
  { name: 'ROUNDDOWN', syntax: 'ROUNDDOWN(A,decimals)', description: 'Round down', category: 'Math', examples: ['ROUNDDOWN(A,0)'] },
  { name: 'CEILING', syntax: 'CEILING(A,significance)', description: 'Round up to multiple', category: 'Math', examples: ['CEILING(A,10)'] },
  { name: 'FLOOR', syntax: 'FLOOR(A,significance)', description: 'Round down to multiple', category: 'Math', examples: ['FLOOR(A,10)'] },
  { name: 'MOD', syntax: 'MOD(A,B)', description: 'Remainder after division', category: 'Math', examples: ['MOD(A,B)'] },
  { name: 'POWER', syntax: 'POWER(A,B)', description: 'A raised to B', category: 'Math', examples: ['POWER(A,2)', 'POWER(A,3)'] },
  { name: 'SQRT', syntax: 'SQRT(A)', description: 'Square root', category: 'Math', examples: ['SQRT(A)'] },
  { name: 'EXP', syntax: 'EXP(A)', description: 'e raised to A', category: 'Math', examples: ['EXP(A)'] },
  { name: 'LN', syntax: 'LN(A)', description: 'Natural log', category: 'Math', examples: ['LN(A)'] },
  { name: 'LOG', syntax: 'LOG(A,base)', description: 'Log with base', category: 'Math', examples: ['LOG(A,10)'] },
  { name: 'LOG10', syntax: 'LOG10(A)', description: 'Base-10 log', category: 'Math', examples: ['LOG10(A)'] },
  { name: 'PI', syntax: 'PI()', description: 'Pi constant', category: 'Math', examples: ['PI()'] },
  { name: 'RAND', syntax: 'RAND()', description: 'Random 0-1', category: 'Math', examples: ['RAND()'] },
  { name: 'SIGN', syntax: 'SIGN(A)', description: 'Sign of number (-1,0,1)', category: 'Math', examples: ['SIGN(A)'] },
  { name: 'TRUNC', syntax: 'TRUNC(A)', description: 'Truncate decimal', category: 'Math', examples: ['TRUNC(A)'] },
  
  // Text
  { name: 'TRIM', syntax: 'TRIM(A)', description: 'Remove extra spaces', category: 'Text', examples: ['TRIM(A)'] },
  { name: 'CLEAN', syntax: 'CLEAN(A)', description: 'Remove non-printable', category: 'Text', examples: ['CLEAN(A)'] },
  { name: 'UPPER', syntax: 'UPPER(A)', description: 'Convert to uppercase', category: 'Text', examples: ['UPPER(A)'] },
  { name: 'LOWER', syntax: 'LOWER(A)', description: 'Convert to lowercase', category: 'Text', examples: ['LOWER(A)'] },
  { name: 'PROPER', syntax: 'PROPER(A)', description: 'Capitalize words', category: 'Text', examples: ['PROPER(A)'] },
  { name: 'LEFT', syntax: 'LEFT(A,n)', description: 'First n characters', category: 'Text', examples: ['LEFT(A,5)'] },
  { name: 'RIGHT', syntax: 'RIGHT(A,n)', description: 'Last n characters', category: 'Text', examples: ['RIGHT(A,5)'] },
  { name: 'MID', syntax: 'MID(A,start,n)', description: 'n chars from start', category: 'Text', examples: ['MID(A,2,5)'] },
  { name: 'LEN', syntax: 'LEN(A)', description: 'Character count', category: 'Text', examples: ['LEN(A)'] },
  { name: 'FIND', syntax: 'FIND(text,A)', description: 'Position of text in A', category: 'Text', examples: ['FIND("x",A)'] },
  { name: 'SUBSTITUTE', syntax: 'SUBSTITUTE(A,old,new)', description: 'Replace text', category: 'Text', examples: ['SUBSTITUTE(A,"old","new")'] },
  { name: 'REPLACE', syntax: 'REPLACE(A,start,n,new)', description: 'Replace chars', category: 'Text', examples: ['REPLACE(A,2,3,"xyz")'] },
  { name: 'CONCATENATE', syntax: 'CONCATENATE(A,B,...)', description: 'Join text', category: 'Text', examples: ['CONCATENATE(A,B)', 'CONCATENATE(A," ",B)'] },
  { name: 'TEXT', syntax: 'TEXT(A,format)', description: 'Format as text', category: 'Text', examples: ['TEXT(A,"0.00")'] },
  { name: 'VALUE', syntax: 'VALUE(A)', description: 'Text to number', category: 'Text', examples: ['VALUE(A)'] },
  { name: 'REPT', syntax: 'REPT(A,n)', description: 'Repeat text n times', category: 'Text', examples: ['REPT("-",10)'] },
  { name: 'SUBSTITUTE', syntax: 'SUBSTITUTE(A,old,new)', description: 'Replace occurrences', category: 'Text', examples: ['SUBSTITUTE(A," ","")'] },
  { name: 'CHAR', syntax: 'CHAR(code)', description: 'Character from code', category: 'Text', examples: ['CHAR(65)'] },
  { name: 'CODE', syntax: 'CODE(A)', description: 'ASCII code of first char', category: 'Text', examples: ['CODE(A)'] },
  { name: 'DOLLAR', syntax: 'DOLLAR(A,decimals)', description: 'Currency format', category: 'Text', examples: ['DOLLAR(A,2)'] },
  { name: 'FIXED', syntax: 'FIXED(A,decimals)', description: 'Number as text', category: 'Text', examples: ['FIXED(A,2)'] },
  { name: 'T', syntax: 'T(A)', description: 'Value if text', category: 'Text', examples: ['T(A)'] },
  
  // Logic
  { name: 'IF', syntax: 'IF(condition, true, false)', description: 'Conditional', category: 'Logic', examples: ['IF(A>B, A, B)', 'IF(A>0, "Yes", "No")'] },
  { name: 'AND', syntax: 'AND(A,B,...)', description: 'All true?', category: 'Logic', examples: ['AND(A>0, B>0)'] },
  { name: 'OR', syntax: 'OR(A,B,...)', description: 'Any true?', category: 'Logic', examples: ['OR(A>0, B>0)'] },
  { name: 'NOT', syntax: 'NOT(A)', description: 'Logical NOT', category: 'Logic', examples: ['NOT(A=B)'] },
  { name: 'TRUE', syntax: 'TRUE()', description: 'True value', category: 'Logic', examples: ['TRUE()'] },
  { name: 'FALSE', syntax: 'FALSE()', description: 'False value', category: 'Logic', examples: ['FALSE()'] },
  { name: 'IFERROR', syntax: 'IFERROR(A, value_if_error)', description: 'Catch errors', category: 'Logic', examples: ['IFERROR(A/B, 0)'] },
  { name: 'IFNA', syntax: 'IFNA(A, value_if_na)', description: 'Catch #N/A', category: 'Logic', examples: ['IFNA(A, "Not found")'] },
  { name: 'XOR', syntax: 'XOR(A,B)', description: 'Exclusive OR', category: 'Logic', examples: ['XOR(A,B)'] },
  
  // Date
  { name: 'TODAY', syntax: 'TODAY()', description: 'Current date', category: 'Date', examples: ['TODAY()'] },
  { name: 'NOW', syntax: 'NOW()', description: 'Current datetime', category: 'Date', examples: ['NOW()'] },
  { name: 'DATE', syntax: 'DATE(year,month,day)', description: 'Create date', category: 'Date', examples: ['DATE(2024,1,15)'] },
  { name: 'TIME', syntax: 'TIME(hour,minute,second)', description: 'Create time', category: 'Date', examples: ['TIME(12,0,0)'] },
  { name: 'YEAR', syntax: 'YEAR(A)', description: 'Year from date', category: 'Date', examples: ['YEAR(A)'] },
  { name: 'MONTH', syntax: 'MONTH(A)', description: 'Month from date', category: 'Date', examples: ['MONTH(A)'] },
  { name: 'DAY', syntax: 'DAY(A)', description: 'Day from date', category: 'Date', examples: ['DAY(A)'] },
  { name: 'HOUR', syntax: 'HOUR(A)', description: 'Hour from time', category: 'Date', examples: ['HOUR(A)'] },
  { name: 'MINUTE', syntax: 'MINUTE(A)', description: 'Minute from time', category: 'Date', examples: ['MINUTE(A)'] },
  { name: 'SECOND', syntax: 'SECOND(A)', description: 'Second from time', category: 'Date', examples: ['SECOND(A)'] },
  { name: 'WEEKDAY', syntax: 'WEEKDAY(A)', description: 'Day of week', category: 'Date', examples: ['WEEKDAY(A)'] },
  { name: 'WEEKNUM', syntax: 'WEEKNUM(A)', description: 'Week number', category: 'Date', examples: ['WEEKNUM(A)'] },
  { name: 'DATEDIF', syntax: 'DATEDIF(start,end,unit)', description: 'Difference', category: 'Date', examples: ['DATEDIF(A,B,"D")'] },
  { name: 'DAYS', syntax: 'DAYS(end,start)', description: 'Days between', category: 'Date', examples: ['DAYS(B,A)'] },
  { name: 'EDATE', syntax: 'EDATE(A,months)', description: 'Add months', category: 'Date', examples: ['EDATE(A,3)'] },
  { name: 'EOMONTH', syntax: 'EOMONTH(A,months)', description: 'End of month', category: 'Date', examples: ['EOMONTH(A,0)'] },
  { name: 'ISOWEEKNUM', syntax: 'ISOWEEKNUM(A)', description: 'ISO week number', category: 'Date', examples: ['ISOWEEKNUM(A)'] },
  { name: 'NETWORKDAYS', syntax: 'NETWORKDAYS(start,end)', description: 'Workdays', category: 'Date', examples: ['NETWORKDAYS(A,B)'] },
  { name: 'WORKDAY', syntax: 'WORKDAY(A,days)', description: 'Future workday', category: 'Date', examples: ['WORKDAY(A,5)'] },
  
  // Statistical
  { name: 'COUNT', syntax: 'COUNT(A,B,...)', description: 'Count numbers', category: 'Statistical', examples: ['COUNT(A,B,C)'] },
  { name: 'COUNTA', syntax: 'COUNTA(A,B,...)', description: 'Count non-empty', category: 'Statistical', examples: ['COUNTA(A,B,C)'] },
  { name: 'COUNTBLANK', syntax: 'COUNTBLANK(A)', description: 'Count empty', category: 'Statistical', examples: ['COUNTBLANK(A)'] },
  { name: 'COUNTIF', syntax: 'COUNTIF(A,condition)', description: 'Conditional count', category: 'Statistical', examples: ['COUNTIF(A,">0")'] },
  { name: 'SUMIF', syntax: 'SUMIF(A,condition,sum_range)', description: 'Conditional sum', category: 'Statistical', examples: ['SUMIF(A,">0",B)'] },
  { name: 'AVERAGEIF', syntax: 'AVERAGEIF(A,condition,avg_range)', description: 'Conditional avg', category: 'Statistical', examples: ['AVERAGEIF(A,">0",B)'] },
  { name: 'MEDIAN', syntax: 'MEDIAN(A,B,...)', description: 'Median value', category: 'Statistical', examples: ['MEDIAN(A,B,C)'] },
  { name: 'MODE', syntax: 'MODE(A,B,...)', description: 'Most common', category: 'Statistical', examples: ['MODE(A,B,C)'] },
  { name: 'STDEV', syntax: 'STDEV(A,B,...)', description: 'Std deviation', category: 'Statistical', examples: ['STDEV(A,B,C)'] },
  { name: 'VAR', syntax: 'VAR(A,B,...)', description: 'Variance', category: 'Statistical', examples: ['VAR(A,B,C)'] },
  { name: 'PERCENTILE', syntax: 'PERCENTILE(A,k)', description: 'kth percentile', category: 'Statistical', examples: ['PERCENTILE(A,0.5)'] },
  { name: 'QUARTILE', syntax: 'QUARTILE(A,quart)', description: 'Quartile', category: 'Statistical', examples: ['QUARTILE(A,1)'] },
  { name: 'RANK', syntax: 'RANK(A,range)', description: 'Rank of value', category: 'Statistical', examples: ['RANK(A,B)'] },
  { name: 'LARGE', syntax: 'LARGE(A,k)', description: 'kth largest', category: 'Statistical', examples: ['LARGE(A,2)'] },
  { name: 'SMALL', syntax: 'SMALL(A,k)', description: 'kth smallest', category: 'Statistical', examples: ['SMALL(A,2)'] },
  
  // Financial
  { name: 'PV', syntax: 'PV(rate,nper,pmt)', description: 'Present value', category: 'Financial', examples: ['PV(0.05,10,100)'] },
  { name: 'FV', syntax: 'FV(rate,nper,pmt)', description: 'Future value', category: 'Financial', examples: ['FV(0.05,10,100)'] },
  { name: 'PMT', syntax: 'PMT(rate,nper,pv)', description: 'Payment', category: 'Financial', examples: ['PMT(0.05,360,200000)'] },
  { name: 'IPMT', syntax: 'IPMT(rate,per,nper,pv)', description: 'Interest payment', category: 'Financial', examples: ['IPMT(0.05,1,10,1000)'] },
  { name: 'PPMT', syntax: 'PPMT(rate,per,nper,pv)', description: 'Principal payment', category: 'Financial', examples: ['PPMT(0.05,1,10,1000)'] },
  { name: 'NPER', syntax: 'NPER(rate,pmt,pv)', description: 'Periods', category: 'Financial', examples: ['NPER(0.05,100,1000)'] },
  { name: 'RATE', syntax: 'RATE(nper,pmt,pv)', description: 'Interest rate', category: 'Financial', examples: ['RATE(10,100,1000)'] },
  { name: 'NPV', syntax: 'NPV(rate,values)', description: 'Net present value', category: 'Financial', examples: ['NPV(0.1,A,B,C)'] },
  { name: 'IRR', syntax: 'IRR(values)', description: 'Internal rate return', category: 'Financial', examples: ['IRR(A,B,C)'] },
  { name: 'DB', syntax: 'DB(cost,salvage,life,period)', description: 'Declining balance', category: 'Financial', examples: ['DB(10000,1000,5,1)'] },
  { name: 'DDB', syntax: 'DDB(cost,salvage,life,period)', description: 'Double declining', category: 'Financial', examples: ['DDB(10000,1000,5,1)'] },
  { name: 'SLN', syntax: 'SLN(cost,salvage,life)', description: 'Straight line depreciation', category: 'Financial', examples: ['SLN(10000,1000,5)'] },
  { name: 'SYD', syntax: 'SYD(cost,salvage,life,period)', description: 'Sum of years digits', category: 'Financial', examples: ['SYD(10000,1000,5,1)'] },
];

export const FORMULA_CATEGORIES = ['All', 'Math', 'Text', 'Logic', 'Date', 'Statistical', 'Financial'] as const;

export function getFormulasByCategory(category: typeof FORMULA_CATEGORIES[number]): ExcelFormula[] {
  if (category === 'All') return EXCEL_FORMULAS;
  return EXCEL_FORMULAS.filter(f => f.category === category);
}

export function searchFormulas(query: string): ExcelFormula[] {
  const q = query.toLowerCase();
  return EXCEL_FORMULAS.filter(f => 
    f.name.toLowerCase().includes(q) || 
    f.description.toLowerCase().includes(q)
  );
}

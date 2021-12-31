# Table Parser
### Ruby Metaprogramming

Project for Scripting Languages course at Faculty of Computing.

The goal of the project is to implement a table parser, which supports both xlsx and xls formats, using metaprogramming concepts in Ruby programming language.


List of commands:
* returns row at given index:
```ruby
table.row(i)
```
* table.row(i)\[j\] - returns element at j from row at i
* table\["column_name"\] - returns the column with given header name
* table\["column_name"\]\[i\] - returns element at given index from column with given header name
* t.column_name - accessing column with given header name
* t.column_name.sum - returns the sum of all elements in given column
* t.column_name.cell_value - returns the whole row which has given value
* t1+t2 - returns sum of two tables (values will be written in t1)
* t1-t2 - returns subtraction of two tables (values will be written in t1)
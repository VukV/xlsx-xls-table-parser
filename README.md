# Table Parser
### Ruby Metaprogramming

Project for Scripting Languages course at Faculty of Computing.

The goal of the project is to implement a table parser, which supports both xlsx and xls formats, using metaprogramming concepts in Ruby programming language.

### List of commands:
Return row at given index:
```ruby
table.row(i)
```
Return element at j from row at i:
```ruby
table.row(i)[j]
```
Return the column with given header name:
```ruby
table["column_name"]
```
Return element at given index from column with given header name:
```ruby
table["column_name"][i]
```
Access column with given header name:
```ruby
t.column_name
```
Return the sum of all elements in given column:
```ruby
t.column_name.sum
```
Return the whole row with given cell value:
```ruby
t.column_name.cell_value
```
Return the sum of two tables (values will be written in t1):
```ruby
t1 + t2
```
Return the subtraction of two tables (values will be written in t1):
```ruby
t1 - t2
```
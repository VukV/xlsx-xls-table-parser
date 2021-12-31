require_relative 'parser'

table_xlsx = Table.new('../tables/table1.xlsx', 'Sheet1')
puts "XLSX:"

puts "cela tabela:"
table_xlsx.print_table
print "\n"

print "drugi red: ", table_xlsx.row(1), "\n"
print "drugi red, treci element: ", table_xlsx.row(1)[2], "\n"
print "peti red, drugi element: ", table_xlsx.row(4)[1], "\n"

#todo test cases

table_xls = Table.new('../tables/table2.xls', 'Sheet1')
puts "XLS"

puts "cela tabela:"
table_xls.print_table
print "\n"

print "drugi red: ", table_xls.row(1), "\n"
print "drugi red, treci element: ", table_xls.row(1)[2], "\n"
print "peti red, drugi element: ", table_xls.row(4)[1], "\n"
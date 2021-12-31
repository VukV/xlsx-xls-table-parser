require_relative 'parser'

print "--------------- XLSX ---------------", "\n"

table_xlsx = Table.new('../tables/table1.xlsx', 'Sheet1')

puts "cela tabela:"
table_xlsx.print_table
print "\n"

print "drugi red: ", table_xlsx.row(1), "\n"
print "drugi red, treci element: ", table_xlsx.row(1)[2], "\n"
print "peti red, drugi element: ", table_xlsx.row(4)[1], "\n"

print "druga kolona: ", table_xlsx["Second"], "\n"
print "druga kolona, cetvrti element: ", table_xlsx["Second"][3], "\n"

third_column_xlsx = table_xlsx.Second
print "druga kolona - ime: ", third_column_xlsx.name, "\n"
print "druga kolona - celije: ", third_column_xlsx.fields, "\n"
print "druga kolona - zbir: ", third_column_xlsx.sum, "\n"
#todo col.col_name.cell

table_xlsx_second = Table.new('../tables/table1.xlsx', 'Sheet1')
puts "\n", "sabiranje:"
table_xlsx+table_xlsx_second
table_xlsx.print_table

puts "\n", "oduzimanje:"
table_xlsx-table_xlsx_second
table_xlsx.print_table


print "\n", "--------------- XLS ---------------", "\n"

table_xls = Table.new('../tables/table2.xls', 'Sheet1')

puts "cela tabela:"
table_xls.print_table
print "\n"

print "drugi red: ", table_xls.row(1), "\n"
print "drugi red, treci element: ", table_xls.row(1)[2], "\n"
print "peti red, drugi element: ", table_xls.row(4)[1], "\n"

print "druga kolona: ", table_xls["Second"], "\n"
print "druga kolona, cetvrti element: ", table_xls["Second"][3], "\n"

third_column_xls = table_xls.Third
print "treca kolona - ime: ", third_column_xls.name, "\n"
print "treca kolona - celije: ", third_column_xls.fields, "\n"
print "treca kolona - zbir: ", third_column_xls.sum, "\n"
#todo col.col_name.cell

table_xls_second = Table.new('../tables/table2.xls', 'Sheet1')
puts "\n", "sabiranje:"
table_xls+table_xls_second
table_xls.print_table

puts "\n", "oduzimanje:"
table_xls-table_xls_second
table_xls.print_table
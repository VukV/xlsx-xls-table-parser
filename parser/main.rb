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
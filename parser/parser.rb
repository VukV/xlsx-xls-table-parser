require 'roo'
require 'spreadsheet'

class Column
    attr_accessor :name, :fields, :table

    def initialize(name, fields, table)
        @name = name
        @fields = fields
        @table = table
    end

    def add_data(data)
        @fields << data
    end

    def sum
        fields_sum = 0
        @fields.each do |field|
            fields_sum += field
        end

        return fields_sum
    end

    def make_cell_methods
        @fields.each do |cell|
            self.define_singleton_method("find_#{cell}") do
                index = -1
                @table.table.each_with_index do |row, i|
                    if i == 0
                        row.each_with_index do |row_cell, j|
                            if row_cell == @name
                                index = j
                            end
                        end
                    else
                        row.each_with_index do |row_cell, j|
                            if row_cell == cell and j == index
                                return row
                            end
                        end
                    end
                end
            end
        end
    end
end

class Table
    attr_accessor :table, :columns, :table_file, :sheet_name
    include Enumerable

    def initialize(path, sheet_name)
        @table = Array.new
        @columns = Array.new
        @sheet_name = sheet_name

        if path.end_with?(".xlsx")
            @table_file = Roo::Spreadsheet.open(path, {:expand_merged_ranges => true})
            
            counter = 0;
            rows_to_remove = Array.new
            
            @table_file.sheet(sheet_name).each_row_streaming do |row|
                row_formulas = row.map {|cell| cell}

                if row_formulas.to_s.include? "@formula=\"SUBTOTAL" or row_formulas.to_s.include? "@formula=\"TOTAL"
                    rows_to_remove << counter
                end

                counter += 1                
            end

            @table_file.sheet(sheet_name).each_with_index do |row, i|
                if rows_to_remove.include? i
                    next
                end
                
                self.make_row(row, 'XLSX')
            end
           
            self.init_columns

        elsif path.end_with?(".xls")
            @table_file = Spreadsheet.open(path)

            for row in @table_file.worksheet(sheet_name)
                self.make_row(row, 'XLS')
            end
            
            self.fix_columns
            self.init_columns
        else
            raise "Wrong file format!"
        end
    end

    def make_row(row, format)
        if row.all? { |x| x.nil? }
            return
        end

        skip_row = false
        row.each_with_index do |row_data, i|
            if format == 'XLS'
                if row_data.to_s.include? "Formula"
                    skip_row = true
                end
            end

            if row_data == nil
                row[i] = 0
            end
        end

        if not skip_row
            @table << row
        end
    end

    def fix_columns
        table_tmp = @table.transpose

        table_tmp.delete_if do |col|
            if col.all? { |x| x == 0 }
                true
            end
        end
        
        @table = table_tmp.transpose

        @table.each_with_index do |row, i|
            row.each_with_index do |cell, j|                
                if cell !~ /\D/
                    @table[i][j] = cell.round
                end
            end
        end
    end

    def init_columns
        table_columns = @table.transpose
        
        table_columns.each do |col|
            current_column = Column.new(col[0], col[1..-1], self)
            @columns << current_column
        end

        @columns.each do |col|
            col_name = col.name

            self.define_singleton_method("#{col_name}") do
                @columns.each do |column|
                    if column.name == col_name
                        return col
                    end
                end
            end

            col.make_cell_methods
        end

    end

    def [](name)
        @columns.each do |col|
            if col.name == name
                return col.fields
            end
        end
    end

    def +(second_table)
        if @table[0] != second_table.table[0]
            raise "Headers are not the same!"
        end

        second_table = second_table.table
        second_table.each_with_index do |row, i|
            if i == 0
                next
            end
            @table << row
        end

        @columns = @columns.clear
        self.init_columns
    end

    def -(second_table)
        if @table[0] != second_table.table[0]
            raise "Headers are not the same!"
        end

        second_table = second_table.table
        second_table.each_with_index do |row, i|
            if i == 0
                next
            end
            
            @table.delete_if do |table_row|
                if table_row == row
                    true
                end
            end
        end

        @columns = @columns.clear
        self.init_columns
    end

    def each(&block)
        @table.each do |table_data|
            block.call(table_data)
        end
    end

    def row(index)
        return @table[index]
    end

    def print_table
        for row in @table
            print row, "\n"
        end
    end
end
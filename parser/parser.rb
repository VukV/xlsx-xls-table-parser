require 'roo'
require 'spreadsheet'

class Column
    attr_accessor :name, :fields

    def initialize(name, fields)
        @name = name
        @fields = fields
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

            for row in @table_file.sheet(sheet_name)
                self.make_row(row, 'XLSX')
            end

            #todo dodaj kolone

        elsif path.end_with?(".xls")
            @table_file = Spreadsheet.open(path)

            for row in @table_file.worksheet(sheet_name)
                self.make_row(row, 'XLS')
            end

            fix_columns()
            #todo dodaj kolone
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
            
            #TODO SREDI BUG
            if format == 'XLSX'
                current_sheet = @table_file.sheet(@sheet_name)
                column_index = @table_file.sheet(@sheet_name).first_column - 1
                print i, column_index, current_sheet.column(column_index + i), current_sheet.formula(row, current_sheet.column(column_index + i)).to_s, "\n"
                if current_sheet.formula(row, current_sheet.column(column_index + i)).to_s.include? "TOTAL" #automatically checks for subtotal
                    puts 'usao if'
                    skip_row = true
                end

            elsif format == 'XLS'
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
        
        table_tmp.each do |col|
            if col.all? { |x| x == 0 }
                table_tmp.delete(col)
            end
        end
        
        @table = table_tmp.transpose
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
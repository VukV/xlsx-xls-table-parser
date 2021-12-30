require 'roo'

class Column
    attr_accessor :name, :fields

    def initialize(name)
        @name = name
        @fields = Array.new
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
    attr_accessor :table, :columns
    include Enumerable

    def initialize(path)
        @table = Array.new
        @columns = Array.new
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
            puts row
        end
    end
end
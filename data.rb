module D
  require 'rubyXL'
  require 'json'
  require 'byebug'
  require 'ostruct'

  Column = OpenStruct
  Columns = [
    Column.new(name: "name", position_in_file: 0),
    Column.new(name: "author", position_in_file: 1),
    Column.new(name: "district", position_in_file: 2),
    Column.new(name: "lat", position_in_file: 4),
    Column.new(name: "lng", position_in_file: 5),
    Column.new(name: "description", position_in_file: 7),
    Column.new(name: "popularity", position_in_file: 8),
    Column.new(name: "parking", position_in_file: 9),
    Column.new(name: "neighborhood", position_in_file: 10),
    Column.new(name: "permanent", position_in_file: 11),
    Column.new(name: "tips", position_in_file: 13),
    Column.new(name: "hash_tags", position_in_file: 14),
    Column.new(name: "images", position_in_file: 15),
  ]
  # Row = OpenStruct.new(*Columns.map(&:name).map(&:to_sym))
  Row = OpenStruct.new
  Columns.each do |c|
    Row[c.name] = c.position_in_file
  end

  class XlsxColumn
    include Comparable
    attr :position

    def initialize(position, values: Columns)
      raise ArgumentError, "There are only #{values.length} columns. #{position} is too big" if position >= values.length
      @position = position
    end

    def succ
      XlsxColumn.new(@position + 1)
    end

    def <=>(other)
      @position <=> other.position
    end

    def to_s
      Columns[@position].name
    end

    def to_int
      Columns[@position].position_in_file
    end
  end

  module Importer
    CantParseThatXlsXFile = Class.new(RuntimeError)

    def import(from, to)
      workbook = begin
                   RubyXL::Parser.parse(from)
                 rescue => e
                   raise CantParseThatXlsXFile, "Cannot parse file: #{e.message}"
                 end
      worksheet = workbook[0]

      # Transform file
      column_range = XlsxColumn.new(0)..XlsxColumn.new(Columns.length - 1) # Use custom range
      rows = worksheet.map.with_index do |row, i|
        next if i == 0
        values = row.cells.values_at(*column_range).map { |v| v ? v.value : nil }
        values = values.
          map { |v| v.to_s&.strip }. # Drop extra spaces
          map { |v| v&.gsub('"', '\"') } # Escape scopes

        r = OpenStruct.new
        r.id = i
        r.name = values[0]
        r.author = "@aesthie.fun"
        r.district = values[2]
        r.coordinates = {
          lat: values[3].to_f,
          lng: values[4].to_f,
        }
        r.description = values[5]
        r.info = [
          { popularity: values[6] },
          { parking: values[7] },
          { permanent: values[9] },
        ]
        r.hash_tags = values[11]&.split(',')&.map(&:strip) || ''
        r.tips = values[10]&.split('.')&.map(&:strip) || ''
        r.images = values[12]&.split(',')&.map(&:strip) || ''
        r
      end[1..-1].reject { |row| row.nil? }
      # Write down file
      File.open(to, "w+") do |file|
        # Prelude
        file.puts "{ \"data\" : ["

        # Content
        rows.each.with_index do |row, index|
          file.puts "," if index > 0
          file.write(row.to_h.to_json)
        end

        # Epilogue
        file.puts "] }"
      end
    end
    module_function :import
  end
end

D::Importer.import('./locations.xlsx', './locations.json')
puts 'done'

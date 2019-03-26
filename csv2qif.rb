#!/usr/bin/env ruby

require 'csvreader'
require 'qif'

class VoBaLe_CSV
  def self.read(path, opts = {sep: ';'})
    @meta = {}; csv = ""; eof = File.foreach(path).count # wie viele zeilen hat die datei?

    File.readlines(path).each_with_index do |line, row|
      raw_line = line
      line = line.force_encoding('iso-8859-1').encode('utf-8')
      
      # in den ersten 12 zeilen sind header-informationen enthalten...
      case row
        when 5-1 then
          if match = line.match(/;"(\d*).*;"([\d\.]*)/)
            @meta[:blz], @meta[:datum] = match.captures; end
        when 6-1 then
          if match = line.match(/;"(\d*).*;"([\d\:]*)/)
            @meta[:konto], @meta[:uhrzeit] = match.captures; end
        when 7-1 then
          if match = line.match(/;"(.*)";;".*";"(.*)"/)
            @meta[:abfrager], @meta[:kontoinhaber] = match.captures; end
        when 9-1 then
          if match = line.match(/;"([\d\.]*)".*;"([\d\.]*)/)
            @meta[:zeitraum_von], @meta[:zeitraum_bis] = match.captures; end
        when 13-1 then
          headers = CsvReader.parse(line, opts).first.to_a.reject {|i| i.strip.empty?}
          @meta[:headers] = []
          headers.each do |header| 
            @meta[:headers] << header
              .downcase
              .gsub(/\./, '')
              .gsub(/[[:punct:]]/, '_')
              .gsub(/[äöüß]/) do |match|
              case match
                when 'ä' then 'ae'
                when 'ö' then 'oe'
                when 'ü' then 'ue'
                when 'ß' then 'ss'
              end
            end
          end
        when eof-2 then
          if match = line.match(/"([\d\.]*).*;"([\d\.,]*)"/)
            @meta[:datum_anfangssaldo], @meta[:anfangssaldo] = match.captures; end
        when eof-1 then
          if match = line.match(/"([\d\.]*).*;"([\d\.,]*)"/)
            @meta[:datum_endsaldo], @meta[:endsaldo] = match.captures; end
      end
      
      # vor 13 ist noch header, nach eof-3 ist footer
      csv << line unless row < 13 || row > eof-3
    end

    @csv = CsvReader.parse(csv, opts)
    @csv = self.rows_to_hashes
    return @meta, @csv
  end

  def self.rows_to_hashes
    hashes = []
    @csv.each do |row|
      hash = {}
      @meta[:headers].each_with_index do |header, i|
        hash[header] = row[i] unless row[i].empty?
      end
      hashes << hash
    end
    hashes
  end
end

Dir['rein/*.csv'].each do |file|
  meta, csv = VoBaLe_CSV.read(file)
  pp csv
  pp meta
end

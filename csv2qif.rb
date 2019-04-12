#!/usr/bin/env ruby

require 'csvreader'
require 'qif'
require 'date'

class VoBaLe_CSV
  def self.read(path, opts = {sep: ';'})
    puts "Reading from #{path}"

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

  def self.german_number_to_float(value, prefix_marker)
    # string als fließkommazahl parsen
    value.gsub! /\./, ''
    value.gsub! /,/, '.'
    value = value.to_f
    # "soll" nach "minus" konvertieren
    # falls nötig, sonst positiv bleiben
    (prefix_marker == 'S') ? -value : value
  end

  def self.rows_to_hashes
    hashes = []
    @csv.each do |row|
      hash = {}
      @meta[:headers].each_with_index do |header, i|
        hash[header.to_sym] = row[i] unless row[i].empty?
      end; if hash[:umsatz]
        hash[:umsatz] = self.german_number_to_float(hash[:umsatz], row.last)
      end
      hashes << hash
    end
    hashes
  end
end

class QIF_File
  def self.normalize_date(value)
    value.gsub! /\./, '/' # datumsformat der lib einhalten!
    if match = value.match(/(\d+)\/(\d+)\/(\d+)/)
      d, m, y = match.captures
      last_day_that_month = Date.civil(y.to_i, m.to_i, -1).day
      # hier müssen wir evtl. die dämliche VoBaLe korrigieren...
      # bei denen gibts nämlich auch mal einen 30. Februar und
      # so Schmarrn...
      d = last_day_that_month if d.to_i > last_day_that_month
      value = [d, m, y].join('/')
    end
    value
  end

  def self.write(path, csv_data, csv_meta)
    puts "Writing to #{path}"
    n_records = 0

    Qif::Writer.open(path, type = 'Bank', format = 'dd/mm/yyyy') do |qif_file|
      csv_data.each do |row|
        next if row[:umsatz].nil? # zeilen ohne umsatz sind keine überweisungen
        puts "\n--- Transaction ---"

        row.each do |key, value|
          case key
            when :buchungstag, :valuta then row[key] = self.normalize_date(value)
          end
        end

        transaction = Qif::Transaction.new(
          #
          # Genutzte QIF-Felder:
          # --------------------
          # date, amount, status (cleared yes/no?), memo,
          # payee, address (up to 5 lines)
          #
          # Ungenutzte QIF-Felder:
          # ----------------------
          # number (id; da nicht im CSV enthalten),
          # split_category, split_memo, split_amount (ggf. von der lib befüllt?),
          # end (von der lib befüllt?),
          # reference, name, description (alle von der lib deprecated),
          # category (nicht im CSV enthalten?)
          #
          date: (row[:valuta] || row[:buchungstag]), # wenns tag der wertstellung nicht gibt ist buchungstag besser als nx
          amount: row[:umsatz], # bereits negative (S) oder positive (H) zahl - siehe VoBaLe_CSV#rows_to_hashes...
          status: (row[:valuta] ? 'cleared' : 'uncleared'), # wenn es ein valuta-datum gibt, sollte der umsatz bereits gebucht sein...
          memo: row[:vorgang_verwendungszweck], # einfach so wie es ist...
          # wenn negativ: z.B. Jan hat Geld von der eG bekommen
          # wenn positiv: z.B. die eG hat Geld vom Finanzamt zurückbekommen
          payee: (row[:umsatz] < 0 ? row[:empfaenger_zahlungspflichtiger] : row[:auftraggeber_zahlungsempfaenger]),
          # wenn negativ: z.B. die eG hat Geld an Jan gezahlt
          # wenn positiv: z.B. die eG hat Geld ans Finanzamt zahlen müssen
          address: (row[:umsatz] < 0 ? row[:auftraggeber_zahlungsempfaenger] : row[:empfaenger_zahlungspflichtiger])
          #
        ); qif_file << transaction; n_records += 1

        pp row
      end
    end
    n_records
  end
end

begin
  n_records = 0
  Dir['rein/*.csv'].each do |file|
    meta, csv = VoBaLe_CSV.read(file)
    n_records += QIF_File.write(file.gsub(/rein/, 'raus').gsub(/\.csv/, '.qif'), csv, meta)
  end
  puts "\nKeine Fehler.\n#{n_records} Datensätze in QIF-Datei geschrieben."
rescue Exception => e
  puts "\nEs gab einen Fehler (\"#{e}\").\nBitte schreib Jonathan <jrs+konstruktiv@weitnahbei.de> eine E-Mail!"
end

print "Drück die ENTER-Taste, um das Programm zu beenden."; gets

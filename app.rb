=begin
  Problem
    Retrieve emails and check for attachments. For a compatible attachment,
    parse the attached file, and convert to Prophet 21 tab-delimited import files.
    Log any errors and email to support.

  Examples
    Tomlinsons .xlsx orders
    Pet Supplies Plus .csv orders
    Promotional orders submitted by a fillable PDF Form with a Submit Button
      *** PDF data structure yet to be determined.
  
  Data Structures
    Input - Email Objects containing file object attachments
    Output - Tab delimited order import files
    YAML file for configuration
    Log files for errors and previous emails/attachments processed
      * SQL table???

  Algorithm
    get_email method

    process_tomlinsons method

    process_psp method

    process_promo method

    P21Order class

    submit_order method

=end

require 'pry'
require 'pry-byebug'
require 'mail'
require 'creek'
require 'tiny_tds'

SQL_DEFAULTS = ['SET ANSI_DEFAULTS ON',
                'SET QUOTED_IDENTIFIER ON',
                'SET CURSOR_CLOSE_ON_COMMIT OFF',
                'SET IMPLICIT_TRANSACTIONS OFF',
                'SET TEXTSIZE 2147483647',
                'SET CONCAT_NULL_YIELDS_NULL ON',
                'USE [P21]']

# Connect to SQL Server
def sql_default_settings(client)
  puts 'Setting SQL defaults...'

  SQL_DEFAULTS.each do |sql|
    result = client.execute(sql)
    result.do
    result.cancel
  end

  puts 'Done'
end

def log_to_sql(log_type, log_description)
  client = TinyTds::Client.new username: 'admin',
                           password: 'Nel$0n!($&',
                           host: 'SQL1', port: 1433
  puts 'Connecting to SQL Server...'
  puts 'Done' if client.active?
  sql_default_settings(client)
  
  sql = "INSERT INTO [nws_log_emailed_orders] ([date_created], [log_type], "\
        "[log_description]) VALUES (CAST('#{DateTime.now}' AS DATETIME2), "\
        "'#{client.escape(log_type)}', '#{client.escape(log_description)}')"
  puts sql
  result = client.execute(sql)
  p result.insert
  puts "#{result.affected_rows} row(s) affected" # if result.affected_rows > 0
  client.close
end


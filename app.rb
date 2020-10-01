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

# Connect to SQL Server
@client = TinyTds::Client.new username: 'admin', password: 'Nel$0n!($&', host: 'SQL1', port: 1433
puts 'Connecting to SQL Server'

puts 'Done' if @client.active?

@client.close
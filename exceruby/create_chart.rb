require "./my_excel.rb"
require "./csv_to_excel_cpu.rb"
require "./csv_to_excel_jvmstat.rb"
require "./csv_to_excel_memory.rb"
require "./png_to_excel.rb"

#------------------------------------------------------------------------------
#                  Get time range
#------------------------------------------------------------------------------

# Get start time
time_from = ""
while 1  do
  p "> input START_TIME with HH:MM:SS format"
  time_from = STDIN.gets.chomp

  if time_from =~ /^([0-1][0-9]|[2][0-3]):[0-5][0-9]:[0-5][0-9]$/ then
    p time_from
     break
  else
    p "ERROR: please re-type START_TIME."
  end

end

# Get end time
time_to = ""
while 1  do
  p "> input END_TIME with HH:MM:SS format"
  time_to = STDIN.gets.chomp

  if time_to =~ /^([0-1][0-9]|[2][0-3]):[0-5][0-9]:[0-5][0-9]$/ && time_from < time_to then
    p time_to
     break
  else
    p "ERROR: please re-type END_TIME."
  end

end

p "Pleae select data directory"


#------------------------------------------------------------------------------
#                  select directory
#------------------------------------------------------------------------------
shell =  WIN32OLE.new('Shell.Application')
path_obj = shell.BrowseForFolder(0, 'Please select data directory', 0)
root_path = ( path_obj.Items.Item.path ).gsub(/\\/,"/")

#------------------------------------------------------------------------------
#                  create chart
#------------------------------------------------------------------------------

Dir::glob( root_path + "/**/*.csv" ).each do |full_path|

    filename = full_path.gsub(/^.*\//,"")

    case filename

    when /^cpu\.csv$/ #cpu.csv
        p full_path
        
        csv_to_excel_cpu( full_path, time_from, time_to )

    when /^jvmstat_.*csv$/ #jvmstat.csv
        p full_path
        csv_to_excel_jvmstat( full_path, time_from, time_to)


    when /^memory\.csv$/ #memory.csv
        p full_path
        csv_to_excel_memory( full_path, time_from, time_to)
    end

end

#------------------------------------------------------------------------------
#                   create excel file
#------------------------------------------------------------------------------

col_num = 1
row_num = 1

if File.exist?(root_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx") then
    File.unlink root_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx"
end

excel_obj = excel_new

workbook_obj = excel_create_workbook(excel_obj)
worksheet_obj = workbook_get_worksheet(workbook_obj,1)
worksheet_set_name(worksheet_obj,"FROM_" + time_from.gsub(/:/,"") + "_TO_" + time_to.gsub(/:/,""))

Dir::glob( root_path + "/**/*.png" ).each do |full_path|
    p full_path
    png_to_excel( worksheet_obj, full_path, col_num, row_num, time_from, time_to )
    row_num = row_num + 18

end

excel_obj.displayAlerts  = false # force save without alert
workbook_save( workbook_obj, root_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx")

excel_quit(excel_obj)

p "FINISHED save as " + root_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx";
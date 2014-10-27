def png_to_excel( file_full_path, save_path, col_num, row_num ,time_from, time_to)

    excel_obj = excel_new

    if File.exist?(save_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx") then
        workbook_obj = excel_obj.Workbooks.Open( save_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx")
        worksheet_obj = workbook_get_worksheet(workbook_obj,1)
    else
	    workbook_obj = excel_create_workbook(excel_obj)
	    worksheet_obj = workbook_get_worksheet(workbook_obj,1)
	    worksheet_set_name(worksheet_obj,"FROM_" + time_from.gsub(/:/,"") + "_TO_" + time_to.gsub(/:/,""))
    end

    excel_screen_up(excel_obj,false)

    filename = file_full_path.gsub(/^.*\//,"")

    case filename
    when /^cpu\.png$/ #cpu.png file 
    	worksheet_paste_text(worksheet_obj,file_full_path.gsub(/\/cpu.png/,"") ,col_num, row_num )
    	row_num = row_num + 1
    end

    worksheet_paste_png(worksheet_obj, file_full_path ,col_num, row_num )

    excel_screen_up(excel_obj,true)

    excel_obj.displayAlerts  = false # force save without alert
    workbook_save( workbook_obj, save_path + "/" + time_from.gsub(/:/,"") + "_" + time_to.gsub(/:/,"") + ".xlsx")

    excel_quit(excel_obj)
    sleep(0.01)
end
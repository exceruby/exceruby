def png_to_excel( worksheet_obj, file_full_path,  col_num, row_num ,time_from, time_to)

    filename = file_full_path.gsub(/^.*\//,"")

    case filename
    when /^cpu\.png$/ #cpu.png file 
    	worksheet_paste_text(worksheet_obj,file_full_path.gsub(/\/cpu.png/,"") ,col_num, row_num )
    	row_num = row_num + 1
    end

    worksheet_paste_png(worksheet_obj, file_full_path ,col_num, row_num )

end
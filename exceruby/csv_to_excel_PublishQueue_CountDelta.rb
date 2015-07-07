def csv_to_excel_PublishQueue_CountDelta( file_full_path, time_from, time_to )
    excel_obj = excel_new # create excel object
    workbook_obj = excel_create_workbook(excel_obj)
    worksheet_obj = workbook_get_worksheet(workbook_obj,1)
    worksheet_set_name(worksheet_obj,"_PublishQueue__CountDelta")

    excel_screen_up(excel_obj,false)
    pre_line_cells = nil
    col_num, row_num = worksheet_load_csv(worksheet_obj, file_full_path, time_from, time_to ) do | line, preline, file_i, excel_i |
            line_cells = line.split(",").map {|cell| cell.strip }
            time = line_cells[0]
            [ time, line_cells ]

    end
    
    #グラフ作成(Heap)
    chart_obj = worksheet_create_chart(worksheet_obj, point(2,2), WIN32OLE::XlLine)
    chart_resize(chart_obj, 100,250,700,220 )


# left, top, right, bottom (1～)
#def range(left, top, right, bottom )
#    point(left, top) + ":" + point(right,bottom)
#end

    chart_set_source_data(chart_obj,worksheet_obj, range(1,1,1,col_num), range(2,1,2,col_num))
    chart_set_title(chart_obj,"_PublishQueue__CountDelta")
    chart_set_plotArea(chart_obj,20,50,600,200)

    chart_series(chart_obj, 1, rgb(   0, 255,   0 ), 2.5) # all


    chart_export(chart_obj,file_full_path.gsub(/\.csv$/,"_PublishQueue__CountDelta.png")) # save chart as png file


    excel_obj.displayAlerts  = false # force save file without alert
    workbook_save( workbook_obj,file_full_path.gsub(/\.csv$/,".xlsx") )
    excel_screen_up(excel_obj,true)
    excel_quit(excel_obj)
end
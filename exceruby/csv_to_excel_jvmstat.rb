
def csv_to_excel_jvmstat( file_full_path, time_from, time_to )
    excel_obj = excel_new # create excel object
    workbook_obj = excel_create_workbook(excel_obj)
    worksheet_obj = workbook_get_worksheet(workbook_obj,1)
    worksheet_set_name(worksheet_obj,"Heap_GC")

    excel_screen_up(excel_obj,false)
    pre_line_cells = nil
    col_num, row_num = worksheet_load_csv(worksheet_obj, file_full_path, time_from, time_to ) do | line, preline, file_i, excel_i |
            line_cells = line.split(",").map {|cell| cell.strip }
            preline_cells = preline.split(",").map {|cell| cell.strip } if preline!=nil
            if file_i == 0 then
                line_cells.push "dYGCT"
                line_cells.push "dFGCT"
            elsif preline!=nil and excel_i > 1
                line_cells.push (line_cells[13].to_f - preline_cells[13].to_f ).to_s
                line_cells.push (line_cells[15].to_f - preline_cells[15].to_f ).to_s       
            end
            time = line_cells[0]
            pre_line_cells = line_cells
            [ time, line_cells ]
    end
    
    #グラフ作成(Heap)
    chart_obj = worksheet_create_chart(worksheet_obj, point(2,2), WIN32OLE::XlLine)
    chart_resize(chart_obj, 100,250,700,220 )
    chart_set_source_data(chart_obj,worksheet_obj, range(1,1,1,col_num), range(3,1,10,col_num))
    chart_set_title(chart_obj,"Heap")
    chart_set_plotArea(chart_obj,20,50,600,200)

    chart_series(chart_obj, 1, rgb(   0,   0, 225 ), 2.5) # S0C
    chart_series(chart_obj, 2, rgb(   0, 255,   0 ), 2.5) # S1C
    chart_series(chart_obj, 3, rgb( 150, 150, 225 ), 2.5) # S0U
    chart_series(chart_obj, 4, rgb( 150, 255, 150 ), 2.5) # S1U
    chart_series(chart_obj, 5, rgb( 230, 180,  86 ), 2.5) # EC
    chart_series(chart_obj, 6, rgb( 255, 255, 150 ), 2.5) # EU
    chart_series(chart_obj, 7, rgb( 156,   0,   0 ), 2.5) # OC
    chart_series(chart_obj, 8, rgb( 255,   0,   0 ), 2.5) # OU    

    chart_export(chart_obj,file_full_path.gsub(/\.csv$/,"_heap.png")) # save chart as png file


    #グラフ作成(GC)
    chart_obj = worksheet_create_chart(worksheet_obj, point(2,2), WIN32OLE::XlLine)
    chart_resize(chart_obj, 350,250,700,220 )

    chart_set_source_data(chart_obj,worksheet_obj, range(1,1,1,col_num), range(18,1,19,col_num))
    chart_set_title(chart_obj,"GC")
    chart_set_plotArea(chart_obj,20,50,600,200)

    chart_export(chart_obj,file_full_path.gsub(/\.csv$/,"_gc.png")) # save chart as png file

    excel_obj.displayAlerts  = false # force save file without alert
    workbook_save( workbook_obj,file_full_path.gsub(/\.csv$/,".xlsx") )
    excel_screen_up(excel_obj,true)
    excel_quit(excel_obj)
end
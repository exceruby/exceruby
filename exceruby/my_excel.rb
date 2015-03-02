require 'win32ole'
WIN32OLE.const_load(WIN32OLE.new('Excel.Application'))

def num2alpha( num )
    alpha = 'A'
    (num-1).times { alpha.succ! }
    alpha
end

def point( col_num, row_num )
    num2alpha(col_num) + row_num.to_s
end

# left, top, right, bottom (1～)
def range(left, top, right, bottom )
    point(left, top) + ":" + point(right,bottom)
end

def rgb(red,green,blue)
    red + green * 256 + blue * 256 * 256
end


def excel_new
    excel_obj = WIN32OLE.new('Excel.Application') 
    excel_obj.visible = true                      
    excel_obj.displayAlerts = false               
    excel_obj
end

def excel_create_workbook( excel_obj, sheets_in_new_workbook: 1)
    excel_obj.SheetsInNewWorkbook = sheets_in_new_workbook
    workbook_obj = excel_obj.workbooks.add                
    workbook_obj
end

def worksheet_set_name(worksheet_obj,name)
    worksheet_obj.Name = name
end

def worksheet_paste_text(worksheet_obj,text ,col_num, row_num)
    worksheet_obj.Cells.Item(row_num,col_num).Value = text
end


def worksheet_paste_png(worksheet_obj,file_full_path ,col_num, row_num)
    pic =  worksheet_obj.Pictures.Insert(file_full_path.gsub(/\//,"\\")).Cut
    pasteCell = worksheet_obj.Cells.Item(row_num,col_num)
    worksheet_obj.paste(pasteCell, pic)
end

def excel_quit( excel_obj )
    excel_obj.quit 
end

def excel_get_workbook( excel_obj, name_or_index )
    workbook_obj = excel_obj.Workbooks name_or_index
    workbook_obj
end

def excel_screen_up( excel_obj, ture_or_false )
    excel_obj.ScreenUpdating = ture_or_false
end

def workbook_save( workbook_obj, filename)
    workbook_obj.saveAs filename.gsub(/\//,"\\")
end


def workbook_get_worksheet( workbook_obj, name_or_index )
    worksheet_obj = workbook_obj.Worksheets( name_or_index )
    worksheet_obj
end

def worksheet_write( worksheet_obj, range, value )
    worksheet_obj.Range( range ).value = value
end

def worksheet_create_chart( worksheet_obj, select_cell, chart_type)
    worksheet_obj.Range( select_cell ).Select
    chart_obj = worksheet_obj.Shapes.AddChart
    chart_obj.Chart.ChartType = chart_type
    chart_obj
end

def chart_resize(chart_obj, top, left, width, height )
    chart_obj.Top    = top
    chart_obj.Left   = left 
    chart_obj.Width  = width
    chart_obj.Height = height
end

def worksheet_load_csv(worksheet_obj,filename, time_from, time_to ) 
    col_num = 0   
    row_num = 0   
    File.open(filename) do |file|
        preline = nil
        file.each_with_index do |line,i|
            time,line_cells = yield(line,preline,i,col_num)

            if col_num == 0 or ( ( time_from <= time ) and ( time <= time_to ) )
                col_num += 1
                row_num = line_cells.size if row_num < line_cells.size
                preline = line
                worksheet_obj.Range(range( 1, 1, 1, col_num )).NumberFormatLocal = "h:mm:ss;@"

                worksheet_obj.Range(range( 1, col_num, line_cells.size, col_num )).value = line_cells
         end
        end

    end
    return col_num, row_num
end


    
def worksheet_create_chart(worksheet_obj, point, type)
    worksheet_obj.Range(point).Select
    chart_obj = worksheet_obj.Shapes.AddChart
    chart_obj.Chart.ChartType = type
    chart_obj
    
end

def chart_resize(chart_obj, top, left, width, height )
    chart_obj.Top    = top
    chart_obj.Left   = left 
    chart_obj.Width  = width
    chart_obj.Height = height
end

def chart_set_title(chart_obj, title)
    chart_obj.Chart.HasTitle = true
    chart_obj.Chart.ChartTitle.Characters.Text = title #Title
    chart_obj.Chart.ChartTitle.Font.Size  = 16         #FontSize
    chart_obj.Chart.ChartTitle.Font.Color = rgb(0,0,0) #palette(0～56) or RGB(57～)
    #chart.ChartTitle.Font.Bold = true       #FontBold
end

def chart_set_axes_max(chart_obj, max)
    chart_obj.Chart.Axes(WIN32OLE::XlValue).MaximumScale = max
end

def chart_set_axes_min(chart_obj, min)
    chart_obj.Chart.Axes(WIN32OLE::XlValue).MinimumScale = min
end

def chart_set_source_data(chart_obj,worksheet_obj, axis_range, data_range)
    chart_obj.Chart.SetSourceData worksheet_obj.Range( axis_range + "," + data_range )
end

def chart_set_plotArea(chart_obj, top, left, width, height )
    chart_obj.Chart.PlotArea.Select
    chart_obj.Chart.PlotArea.InsideTop    = top
    chart_obj.Chart.PlotArea.InsideLeft   = left
    chart_obj.Chart.PlotArea.InsideWidth  = width
    chart_obj.Chart.PlotArea.InsideHeight = height
end

def chart_export(chart_obj,filename)
    chart_obj.Chart.Export(filename.gsub(/\//,"\\"))
end

def chart_series(chart_obj, series_no, rgb, weight)
    begin
        series = chart_obj.Chart.SeriesCollection(series_no)
        series.Select
        series.Format.Line.Weight = weight
        series.Format.Line.Visible = false
        series.Format.Line.Visible = true
        series.Format.Line.ForeColor.RGB = rgb
    rescue => evar
        p "ERROR! Series:" + series_no.to_s
    end

end



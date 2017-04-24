Dim kow(100),a
Set Exclcre=createobject("excel.application")
Set Exclopen=Exclcre.Workbooks.Open("D:\Wikipedia Automation")
'Exclcre.Application.Visible = true

Set tc = Exclcre.Sheets("Test Case")
Row_Count_tc= tc.usedrange.rows.count
Column_Count_tc =tc.UsedRange.Columns.Count

Set ts = Exclcre.sheets("Test Step")
Row_Count_ts= ts.usedrange.rows.count
 Column_Count_ts=ts.UsedRange.Columns.Count

For i=2 To (Row_Count_tc) Step 1
	a=Exclopen.sheets(1).cells(i,4).value
    
    						If a=1 Then
    						id=Exclopen.sheets(1).cells(i,1).value
    						'tsid=Exclopen.sheets(2).cells(i,1).value
    						For j=2  To Row_Count_ts Step 1
                                   tsid=Exclopen.sheets(2).cells(j,1).value						
          						if(id=tsid) then
    							keyword=Exclopen.sheets(2).cells(j,5).value
    							Call executeFlow(keyword)
    							
    							End if
    							
    						Next                                              
						End  if 
next



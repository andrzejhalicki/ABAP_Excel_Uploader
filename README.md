This class lets you quickly upload excel file into internal table.

 In **constructor** pass following parameters:
 * -> iv_filename - path to Excel file
 * -> iv_last_column_number - number of columns in Excel
 * -> iv_number_of_rows - number of ALL rows (including header)
 * -> iv_first_data_row - number of first row with actual data
 * -> iv_names_row - row number where fieldnames are stored
 * -> iv_dataelements_row - row number where dataelement names are stored

 Next you will get access to public table mt_data.
 Example implementation:
 ```
 DATA(lo_excel_uploader) = NEW zcl_excel_uploader( iv_filename           = p_file
                                                   iv_last_column_number = p_colnum
                                                   iv_number_of_rows     = p_rownum
*                                                 iv_first_data_row     = 4
*                                                 iv_names_row          = 1
*                                                 iv_dataelements_row   = 2 ).

 FIELD-SYMBOLS: <lt_data> TYPE ANY TABLE,
                <ls_data> TYPE any.
                
 ASSIGN lo_excel_uplaoder->mt_data->* TO <lt_data>.
 
 DATA(lv_field_key) = 'ID'.
 READ TABLE <lt_data> ASSIGNING <ls_data> WITH KEY (lv_field_key) = '1'.

 LOOP AT <lt_data> ASSIGNING FIELD-SYMBOL(<ls_data>).
   ASSIGN COMPONENT 'ID' OF STRUCTURE <ls_data> TO FIELD-SYMBOL(<lv_id>).
 ...
 ENDLOOP.
 ```
 You can also quickly display the table with method **display( )**.
```
lo_excel_uploader->display( ).
```

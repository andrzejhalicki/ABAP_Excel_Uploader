REPORT z_excelupload.

DATA gt_filetable TYPE filetable.

FIELD-SYMBOLS <gt_table> TYPE ANY TABLE.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-100.
PARAMETERS: p_file LIKE rlgrap-filename MEMORY ID ad_local_path OBLIGATORY MODIF ID fil.
PARAMETERS: p_colnum TYPE i OBLIGATORY DEFAULT 4.
PARAMETERS: p_rownum TYPE i OBLIGATORY DEFAULT 5.
SELECTION-SCREEN END OF BLOCK b1.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  DATA: lv_return_code TYPE i.
  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      window_title            = 'Upload Excel'
      default_filename        = ''
      initial_directory       = CONV #( p_file )
      multiselection          = ''
    CHANGING
      file_table              = gt_filetable
      rc                      = lv_return_code
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5
  ).
  IF sy-subrc <> 0.
    EXIT.
  ENDIF.

  p_file = gt_filetable[ 1 ].

START-OF-SELECTION.

NEW zcl_excel_uploader( iv_filename           = p_file
                        iv_last_column_number = p_colnum
                        iv_number_of_rows     = p_rownum
*                       iv_first_data_row     = 4
*                       iv_names_row          = 1
*                       iv_dataelements_row   = 2
                                                        )->display( ).

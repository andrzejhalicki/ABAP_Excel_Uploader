* ---------------------------------------------------------------------------------------+
* | Excel Uploader
* +-------------------------------------------------------------------------------------------------+
* | Author: Andrzej Halicki
* | Date: 27.07.2019
* | SAP community: https://people.sap.com/andrzejhalicki
* | LinkedIn: /in/halickiandrzej
* +--------------------------------------------------------------------------------------
* | This class lets you quickly upload excel file into internal table.
* |
* | In constructor pass following parameters:
* |  -> iv_filename - path to Excel file
* |  -> iv_last_column_number - number of columns in Excel
* |  -> iv_number_of_rows - number of ALL rows (including header)
* |  -> iv_first_data_row - number of first row with actual data
* |  -> iv_names_row - row number where fieldnames are stored
* |  -> iv_dataelements_row - row number where dataelement names are stored
* |
* | Next you will get access to public table mt_data.
* | Example implementation:
* |
* | FIELD-SYMBOLS: <lt_data> TYPE ANY TABLE,
* |                <ls_data> TYPE any.
* | ASSIGN mt_data->* TO <lt_data>.
* | DATA(lv_field_key) = 'ID'.
* | READ TABLE <lt_data> ASSIGNING <ls_data> WITH KEY (lv_field_key) = '1'.
* |
* | LOOP AT <lt_data> ASSIGNING FIELD-SYMBOL(<ls_data>).
* |   ASSIGN COMPONENT 'ID' OF STRUCTURE <ls_data> TO FIELD-SYMBOL(<lv_id>).
* | ...
* | ENDLOOP.
* |
* | You can also quickly display the table with method display( ).
* +--------------------------------------------------------------------------------------

CLASS zcl_excel_uploader DEFINITION PUBLIC FINAL.
  PUBLIC SECTION.
    DATA:
      mt_data       TYPE REF TO data.
    METHODS:
      constructor IMPORTING iv_filename           TYPE rlgrap-filename
                            iv_last_column_number TYPE i
                            iv_number_of_rows     TYPE i
                            iv_first_data_row     TYPE i DEFAULT 4
                            iv_names_row          TYPE i DEFAULT 1
                            iv_dataelements_row   TYPE i DEFAULT 2,
      display.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_tab,
             name    TYPE char30,
             dt_name TYPE char30,
           END OF ty_tab,
           tt_tab TYPE TABLE OF ty_tab WITH EMPTY KEY.
    DATA:
      mt_excel_data         TYPE TABLE OF alsmex_tabline,
      mt_name_dataelement   TYPE tt_tab,
      mt_components         TYPE cl_abap_structdescr=>component_table,
      ms_data               TYPE REF TO data,
      mv_filename           TYPE rlgrap-filename,
      mv_last_column_number TYPE i,
      mv_number_of_rows     TYPE i,
      mv_first_data_row     TYPE i,
      mv_names_row          TYPE i,
      mv_dataelements_row   TYPE i,

      mo_salv               TYPE REF TO cl_salv_table.

    METHODS:
      upload_excel,
      build_table_from_excel,
      build_components_table RETURNING VALUE(rt_components) TYPE cl_abap_structdescr=>component_table,
      fill_data_table,
      get_column_names.

ENDCLASS.



CLASS ZCL_EXCEL_UPLOADER IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_UPLOADER->BUILD_COMPONENTS_TABLE
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RT_COMPONENTS                  TYPE        CL_ABAP_STRUCTDESCR=>COMPONENT_TABLE
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD build_components_table.

    DATA: lr_description TYPE REF TO cl_abap_typedescr.

    mt_name_dataelement = VALUE #( FOR i IN mt_excel_data
                                   WHERE ( row = mv_dataelements_row )
                                    (
                                      name    = mt_excel_data[ col = i-col row = mv_names_row ]-value
                                      dt_name = i-value
                                     )
                                  ).

    LOOP AT mt_name_dataelement ASSIGNING FIELD-SYMBOL(<ls_name_dataelement>).
      cl_abap_typedescr=>describe_by_name(
        EXPORTING
            p_name = CONV #( <ls_name_dataelement>-dt_name )
        RECEIVING
            p_descr_ref = lr_description
        EXCEPTIONS
            type_not_found = 1
            OTHERS         = 2 ).
      IF sy-subrc <> 0.

      ELSE.
        APPEND VALUE #( name = <ls_name_dataelement>-name
                        type = cl_abap_elemdescr=>get_by_kind( p_type_kind = lr_description->type_kind
                                                               p_length    = lr_description->length
                                                               p_decimals  = lr_description->decimals ) ) TO rt_components.
      ENDIF.
    ENDLOOP.

    CHECK mt_components IS INITIAL.
    mt_components = rt_components.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_UPLOADER->BUILD_TABLE_FROM_EXCEL
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD build_table_from_excel.

    TRY.
        DATA(lr_structure) = cl_abap_structdescr=>create( me->build_components_table( ) ).
        DATA(lr_table) = cl_abap_tabledescr=>create( p_line_type = CAST #( lr_structure ) ).
      CATCH cx_sy_struct_creation.

      CATCH cx_sy_table_creation.

    ENDTRY.

    CREATE DATA ms_data TYPE HANDLE lr_structure.

    CREATE DATA mt_data TYPE HANDLE lr_table.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_UPLOADER->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_FILENAME                    TYPE        RLGRAP-FILENAME
* | [--->] IV_LAST_COLUMN_NUMBER          TYPE        I
* | [--->] IV_NUMBER_OF_ROWS              TYPE        I
* | [--->] IV_FIRST_DATA_ROW              TYPE        I (default =4)
* | [--->] IV_NAMES_ROW                   TYPE        I (default =1)
* | [--->] IV_DATAELEMENTS_ROW            TYPE        I (default =2)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD constructor.

    mv_filename = iv_filename.
    mv_last_column_number = iv_last_column_number.
    mv_number_of_rows = iv_number_of_rows.
    mv_first_data_row = iv_first_data_row.
    mv_names_row       = iv_names_row.
    mv_dataelements_row  = iv_dataelements_row.

    me->upload_excel( ).

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_UPLOADER->DISPLAY
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD display.
    FIELD-SYMBOLS <lt_table> TYPE ANY TABLE.

    ASSIGN mt_data->* TO <lt_table>.

    TRY.
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table = mo_salv
          CHANGING
            t_table      = <lt_table> ).
      CATCH cx_salv_msg.

    ENDTRY.

    me->get_column_names( ).
    mo_salv->display( ).

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_UPLOADER->FILL_DATA_TABLE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD fill_data_table.

    FIELD-SYMBOLS: <data>      TYPE ANY TABLE,
                   <data_line> TYPE any.

    ASSIGN mt_data->* TO <data>.
    ASSIGN ms_data->* TO <data_line>.

    LOOP AT mt_excel_data ASSIGNING FIELD-SYMBOL(<excel_data_line>) WHERE row >= mv_first_data_row.
      ASSIGN COMPONENT mt_excel_data[ row = mv_names_row col = <excel_data_line>-col ]-value OF STRUCTURE <data_line> TO FIELD-SYMBOL(<cell>).
      IF sy-subrc = 0.
        <cell> = <excel_data_line>-value.
      ENDIF.
      AT END OF row.
        INSERT <data_line> INTO TABLE <data>.
        CLEAR <data_line>.
      ENDAT.
    ENDLOOP.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_UPLOADER->GET_COLUMN_NAMES
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_column_names.

    DATA lt_dd04 TYPE TABLE OF dd04t.
    DATA(lo_columns) = mo_salv->get_columns( ).

    LOOP AT mt_name_dataelement ASSIGNING FIELD-SYMBOL(<ls_name_dataelement>).
      DATA(lo_column) = lo_columns->get_column( columnname = <ls_name_dataelement>-name ).
      DATA(lv_roll_name) = CONV dd04l-rollname( <ls_name_dataelement>-dt_name ).
      CALL FUNCTION 'DD_DTEL_GET'
        EXPORTING
          langu         = sy-langu
          roll_name     = lv_roll_name
        TABLES
          dd04t_tab_n   = lt_dd04
        EXCEPTIONS
          illegal_value = 1
          OTHERS        = 2.
      IF sy-subrc = 0.
        lo_column->set_long_text( value = lt_dd04[ 1 ]-scrtext_l ).
        lo_column->set_medium_text( value = lt_dd04[ 1 ]-scrtext_m ).
        lo_column->set_short_text( value = lt_dd04[ 1 ]-scrtext_s ).
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_UPLOADER->UPLOAD_EXCEL
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD upload_excel.

    cl_progress_indicator=>progress_indicate(
      EXPORTING
        i_text               = |Uploading excel file...|
        i_processed          = 0
        i_total              = 1
        i_output_immediately = abap_false
    ).

    CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
      EXPORTING
        filename                = mv_filename
        i_begin_col             = 1
        i_begin_row             = 1
        i_end_col               = mv_last_column_number
        i_end_row               = mv_number_of_rows
      TABLES
        intern                  = mt_excel_data
      EXCEPTIONS
        inconsistent_parameters = 1
        upload_ole              = 2
        OTHERS                  = 3.
    IF sy-subrc <> 0.

    ENDIF.

    me->build_table_from_excel( ).
    me->fill_data_table( ).

  ENDMETHOD.
ENDCLASS.

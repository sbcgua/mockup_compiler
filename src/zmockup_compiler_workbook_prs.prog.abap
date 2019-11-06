**********************************************************************
* WORKBOOK PARSER
**********************************************************************

class ltcl_workbook_parser_test definition deferred.

class lcx_excel definition inheriting from lcx_error.
  public section.
    class-methods excel_error
      importing
        msg type string
      raising
        lcx_excel.
endclass.

class lcx_excel implementation.
  method excel_error.
    raise exception type lcx_excel
      exporting
        msg = msg.
  endmethod.
endclass.

**********************************************************************

interface lif_excel.

  types:
    begin of ty_sheet_content,
      cell_row    type int4,
      cell_column type int4,
      cell_value  type string,
      cell_coords type string,
      cell_style  type uuid,
      data_type   type string,
    end of ty_sheet_content,
    tt_sheet_content type standard table of ty_sheet_content with default key.

  types:
    begin of ty_style,
      id     type uuid,
      format type string,
    end of ty_style,
    tt_styles type standard table of ty_style with key id.

  methods get_sheet_names
    returning
      value(rt_sheet_names) type string_table.

  methods get_sheet_content
    importing
      iv_sheet_name type string
    returning
      value(rt_content) type tt_sheet_content
    raising
      lcx_excel.

  methods get_styles
    returning
      value(rt_styles) type tt_styles
    raising
      lcx_excel.

*  method get_styles

endinterface.

class lcl_excel_abap2xlsx definition final.
  public section.
    interfaces lif_excel.

    types:
      begin of ty_ws_item,
        title     type string,
        worksheet type ref to zcl_excel_worksheet,
      end of ty_ws_item,
      tt_worksheets type standard table of ty_ws_item with key title.

    class-methods load
      importing
        iv_xdata type xstring
      returning
        value(ro_excel) type ref to lcl_excel_abap2xlsx
      raising
        lcx_excel.
  private section.
    data mo_excel type ref to zcl_excel.
    data mt_worksheets type tt_worksheets.
endclass.

class lcl_excel_abap2xlsx implementation.

  method load.

    data lx_xls type ref to zcx_excel.
    data lo_reader type ref to zif_excel_reader.

    create object ro_excel.

    try.
      create object lo_reader type zcl_excel_reader_2007.
      ro_excel->mo_excel = lo_reader->load( iv_xdata ).
    catch zcx_excel into lx_xls.
      lcx_excel=>excel_error( 'zcl_excel error: ' && lx_xls->get_text( ) ). "#EC NOTEXT
    endtry.

  endmethod.

  method lif_excel~get_sheet_names.

    field-symbols <w> like line of mt_worksheets.
    if mt_worksheets is not initial.
      loop at mt_worksheets assigning <w>.
        append <w>-title to rt_sheet_names.
      endloop.
      return.
    endif.

    data lo_iter type ref to cl_object_collection_iterator.

    lo_iter = mo_excel->get_worksheets_iterator( ).
    while lo_iter->has_next( ) is not initial.
      append initial line to mt_worksheets assigning <w>.
      <w>-worksheet ?= lo_iter->get_next( ).
      <w>-title      = <w>-worksheet->get_title( ).
    endwhile.

  endmethod.

  method lif_excel~get_sheet_content.

    field-symbols <w> like line of mt_worksheets.

    read table mt_worksheets assigning <w> with key title = iv_sheet_name.
    if sy-subrc <> 0.
      lcx_excel=>excel_error( msg = |Workbook does not contain [{ iv_sheet_name }] sheet| ). "#EC NOTEXT
    endif.

    field-symbols <src> like line of <w>-worksheet->sheet_content.
    field-symbols <dst> like line of rt_content.
    loop at <w>-worksheet->sheet_content assigning <src>.
      append initial line to rt_content assigning <dst>.
      move-corresponding <src> to <dst>.
    endloop.

  endmethod.

  method lif_excel~get_styles.

    data:
      lo_style type ref to zcl_excel_style,
      lo_iter  type ref to cl_object_collection_iterator.
    field-symbols <s> like line of rt_styles.

    lo_iter = mo_excel->get_styles_iterator( ).
    while lo_iter->has_next( ) is not initial.
      lo_style ?= lo_iter->get_next( ).
      append initial line to rt_styles assigning <s>.
      <s>-id = lo_style->get_guid( ).
      <s>-format = lo_style->number_format->format_code.
    endwhile.

  endmethod.

endclass.

**********************************************************************

class lcl_workbook_parser definition final
  friends ltcl_workbook_parser_test.
  public section.
    types:
      begin of ty_mock,
        name type string,
        data type string,
      end of ty_mock,
      tt_mocks type standard table of ty_mock with key name,

      begin of ty_sheet_cell,
        row type i,
        col type i,
        value type string,
        type type char1,
      end of ty_sheet_cell,
      tt_sheet_data type standard table of ty_sheet_cell with key row col,

      begin of ty_ws_item,
        title     type string,
        worksheet type ref to zcl_excel_worksheet,
      end of ty_ws_item,
      tt_worksheets type standard table of ty_ws_item with key title,

      begin of ty_range,
        row_min type i,
        row_max type i,
        col_min type i,
        col_max type i,
      end of ty_range,
      tt_uuid type standard table of uuid with default key.

    constants content_sheet_name type string value '_contents'.
    constants exclude_sheet_name type string value '_exclude'.

    class-methods parse
      importing
        iv_data type xstring
      returning value(rt_mocks) type tt_mocks
      raising lcx_error.

  private section.
    class-methods read_contents
      importing
        it_content type zexcel_t_cell_data
      returning value(rt_sheets_to_save) type string_table
      raising lcx_error zcx_excel.

    class-methods read_exclude
      importing
        it_exclude type zexcel_t_cell_data
      returning value(rt_sheets_to_exclude) type string_table
      raising lcx_error zcx_excel.

    class-methods get_sheets_list
      importing
        io_excel type ref to zcl_excel
      returning value(rt_worksheets) type tt_worksheets.

    class-methods find_date_styles
      importing
        io_excel type ref to zcl_excel
      returning value(rt_style_uuids) type tt_uuid.

    class-methods clip_header
      importing
        it_head type string_table
      changing
        cs_range type ty_range.

    class-methods clip_rows
      importing
        it_content   type zexcel_t_cell_data
        iv_start_row type i default 1
      changing
        cs_range type ty_range.

    class-methods clip_range
      importing
        it_content type zexcel_t_cell_data
      returning value(rs_range) type ty_range
      raising lcx_error zcx_excel.

    class-methods convert_to_date
      importing iv_value type string
      returning value(rv_date) type string
      raising zcx_excel.

    class-methods read_row
      importing
        it_content type zexcel_t_cell_data
        i_row type i
        i_colmin type i default 1
        i_colmax type i default 9999
        it_date_styles type tt_uuid optional
      returning value(rt_values) type string_table
      raising zcx_excel.

    class-methods is_row_empty
      importing
        it_content type zexcel_t_cell_data
        i_row type i
        i_colmin type i default 1
        i_colmax type i
      returning value(rv_yes) type abap_bool.

    class-methods convert_sheet
      importing
        it_content type zexcel_t_cell_data
        it_date_styles type tt_uuid optional
      returning value(rv_data) type string
      raising lcx_error zcx_excel.

endclass.


class lcl_workbook_parser implementation.

  method parse.
    data:
      lx_xls    type ref to zcx_excel,
      lo_excel  type ref to zcl_excel,
      lo_reader type ref to zif_excel_reader.

    try.
      create object lo_reader type zcl_excel_reader_2007.
      lo_excel = lo_reader->load( iv_data ).
    catch zcx_excel into lx_xls.
      lcx_error=>raise( 'zcl_excel error: ' && lx_xls->get_text( ) ). "#EC NOTEXT
    endtry.

    " Get list of work sheets
    data lt_worksheets type tt_worksheets.
    data lt_sheets_to_save type string_table.
    lt_worksheets = get_sheets_list( lo_excel ).

    " Check and read content
    field-symbols <ws> like line of lt_worksheets.
    read table lt_worksheets assigning <ws> with key title = content_sheet_name.
    if sy-subrc is not initial.
      loop at lt_worksheets assigning <ws>.
        append <ws>-title to lt_sheets_to_save. " Just parse all
      endloop.
    else.
      try.
        lt_sheets_to_save = read_contents( <ws>-worksheet->sheet_content ).
      catch zcx_excel into lx_xls.
        lcx_error=>raise( 'Excel error: ' && lx_xls->get_text( ) ). "#EC NOTEXT
      endtry.

      " Check all sheets exist
      field-symbols <sheet_name> like line of lt_sheets_to_save.
      loop at lt_sheets_to_save assigning <sheet_name>.
        read table lt_worksheets with key title = <sheet_name> transporting no fields.
        if sy-subrc is not initial.
          lcx_error=>raise( msg = |Workbook does not contain [{ <sheet_name> }] sheet| ). "#EC NOTEXT
        endif.
      endloop.
    endif.

    " Read excludes
    data lt_excludes type sorted table of string with non-unique key table_line.
    read table lt_worksheets assigning <ws> with key title = exclude_sheet_name.
    if sy-subrc = 0.
      try.
        lt_excludes = read_exclude( <ws>-worksheet->sheet_content ).
      catch zcx_excel into lx_xls.
        lcx_error=>raise( 'Excel error: ' && lx_xls->get_text( ) ). "#EC NOTEXT
      endtry.

      " exclude sheets
      data lv_index type sy-tabix.
      loop at lt_sheets_to_save assigning <sheet_name>.
        lv_index = sy-tabix.
        read table lt_excludes with key table_line = <sheet_name> transporting no fields.
        if sy-subrc = 0 or <sheet_name> = exclude_sheet_name.
          delete lt_sheets_to_save index lv_index.
        endif.
      endloop.
    endif.

    data lt_date_styles type tt_uuid.
    lt_date_styles = find_date_styles( lo_excel ).

    " convert sheets
    try.
      field-symbols <mock> like line of rt_mocks.
      loop at lt_sheets_to_save assigning <sheet_name>.
        read table lt_worksheets with key title = <sheet_name> assigning <ws>.
        append initial line to rt_mocks assigning <mock>.
        <mock>-name = <sheet_name>.
        <mock>-data = convert_sheet( it_content = <ws>-worksheet->sheet_content it_date_styles = lt_date_styles ).
      endloop.
    catch zcx_excel into lx_xls.
      lcx_error=>raise( 'Excel error: ' && lx_xls->get_text( ) ). "#EC NOTEXT
    endtry.

  endmethod.  " parse.

  method convert_sheet.

    data ls_range type ty_range.
    ls_range = clip_range( it_content ).

    data lt_data type string_table.
    data lt_values type string_table.
    data lv_temp_line type string.
    do ls_range-row_max - ls_range-row_min + 1 times. " starts from 1 always
      lt_values = read_row(
        it_content     = it_content
        it_date_styles = it_date_styles
        i_row          = sy-index + ls_range-row_min - 1
        i_colmin       = ls_range-col_min
        i_colmax       = ls_range-col_max ).
      lv_temp_line = concat_lines_of( table = lt_values sep = cl_abap_char_utilities=>horizontal_tab ).
      append lv_temp_line to lt_data.
    enddo.

    rv_data = concat_lines_of( table = lt_data sep = cl_abap_char_utilities=>newline ).

  endmethod.  " convert_sheet.

  method get_sheets_list.
    data lo_iter type ref to cl_object_collection_iterator.

    field-symbols <ws> like line of rt_worksheets.

    lo_iter = io_excel->get_worksheets_iterator( ).
    while lo_iter->has_next( ) is not initial.
      append initial line to rt_worksheets assigning <ws>.
      <ws>-worksheet ?= lo_iter->get_next( ).
      <ws>-title      = <ws>-worksheet->get_title( ).
    endwhile.

  endmethod.  " get_sheets_list.

  method find_date_styles.
    data:
          fmt type string,
          uuid type uuid,
          lo_style type ref to zcl_excel_style,
          lo_iter  type ref to cl_object_collection_iterator.

    lo_iter = io_excel->get_styles_iterator( ).
    while lo_iter->has_next( ) is not initial.
      lo_style ?= lo_iter->get_next( ).
      fmt = lo_style->number_format->format_code.
      if fmt ca 'd' and fmt ca 'm' and fmt ca 'y'. " Guess it is date ...
        uuid = lo_style->get_guid( ).
        append uuid to rt_style_uuids.
      endif.
    endwhile.
  endmethod.  " find_date_styles.

  method convert_to_date.
    if iv_value is initial.
      return. " rv_date also initial
    endif.
    rv_date = zcl_excel_common=>excel_string_to_date( iv_value ).
    rv_date = rv_date+6(2) && '.' && rv_date+4(2) && '.' && rv_date+0(4). " TODO reafactor
  endmethod.  " convert_to_date.

  method read_row.
    assert i_row > 0.
    assert i_colmin > 0.
    assert i_colmin <= i_colmax.

    field-symbols <c> like line of it_content.
    data last_col type i.
    data tmp type string.
    last_col = i_colmin - 1.
    loop at it_content assigning <c> where cell_row = i_row and cell_column between i_colmin and i_colmax.
      do <c>-cell_column - last_col - 1 times. " fill blanks
        append initial line to rt_values.
      enddo.
      last_col = <c>-cell_column.
      tmp      = <c>-cell_value.
      if it_date_styles is supplied and <c>-cell_style is not initial.
        read table it_date_styles with key table_line = <c>-cell_style transporting no fields.
        if sy-subrc is initial.
          tmp = convert_to_date( tmp ).
        endif.
      endif.
      append tmp to rt_values.
    endloop.
    if last_col < i_colmax and i_colmax is supplied. " fill blanks
      do i_colmax - last_col times.
        append initial line to rt_values.
      enddo.
    endif.

  endmethod.  " read_row.

  method is_row_empty.

    assert i_row > 0.
    assert i_colmin > 0.
    assert i_colmin <= i_colmax.

    field-symbols <c> like line of it_content.
    rv_yes = abap_true.

    loop at it_content assigning <c> where cell_row = i_row and cell_column between i_colmin and i_colmax.
      if <c>-cell_value is not initial.
        rv_yes = abap_false.
        exit.
      endif.
    endloop.

  endmethod. "is_row_empty.

  method clip_header.
    field-symbols <str> like line of it_head.
    cs_range-col_min = 1.
    cs_range-col_max = lines( it_head ).

    " clip on first empty field
    read table it_head with key table_line = '' transporting no fields.
    if sy-subrc is initial.
      cs_range-col_max = sy-tabix - 1.
    endif.

    " clip technical fields from the end
    while cs_range-col_max > 0.
      read table it_head index cs_range-col_max assigning <str>.
      if <str>+0(1) = '_'.
        cs_range-col_max = cs_range-col_max - 1.
      else.
        exit.
      endif.
    endwhile.

    " clip technical fields from the start
    while cs_range-col_min < cs_range-col_max.
      read table it_head index cs_range-col_min assigning <str>.
      if <str>+0(1) = '_'.
        cs_range-col_min = cs_range-col_min + 1.
      else.
        exit.
      endif.
    endwhile.

  endmethod.  " clip_header.

  method clip_rows.
    data lv_is_empty type abap_bool.

    cs_range-row_min = iv_start_row.
    cs_range-row_max = iv_start_row.

    do.
      lv_is_empty = is_row_empty(
        it_content = it_content
        i_row = cs_range-row_max + 1
        i_colmin = cs_range-col_min
        i_colmax = cs_range-col_max ).
      if lv_is_empty = abap_true.
        exit.
      else.
        cs_range-row_max = cs_range-row_max + 1.
      endif.
    enddo.

  endmethod.  " clip_rows.

  method clip_range.
    field-symbols <c> like line of it_content.
    " Assuming content is sorted
    read table it_content assigning <c> index 1.
    if sy-subrc is not initial or <c>-cell_row <> 1 or <c>-cell_column <> 1 or <c>-cell_value is initial.
      lcx_error=>raise( msg = 'Sheet data must start at A1 cell' ). "#EC NOTEXT
    endif.

    data lt_head type string_table.
    data lv_start_row type i value 1.
    field-symbols <cell> type string.

    lt_head = read_row(
      it_content = it_content
      i_row      = 1 ).

    read table lt_head index 1 assigning <cell>.
    if <cell>+0(1) = '#'. " Skip first comment row
      lv_start_row = 2.
      lt_head = read_row(
        it_content = it_content
        i_row      = 2 ).
    endif.

    clip_header(
      exporting
        it_head = lt_head
      changing
        cs_range = rs_range ).
    if rs_range-col_max < rs_range-col_min.
      lcx_error=>raise( msg = 'Only contains technical field (_...) in the header' ). "#EC NOTEXT
    endif.

    clip_rows(
      exporting
        it_content   = it_content
        iv_start_row = lv_start_row
      changing
        cs_range = rs_range ).

  endmethod.  " clip_range.

  method read_contents.

    data ls_range type ty_range.
    ls_range = clip_range( it_content ).

    if ls_range-col_max - ls_range-col_min + 1 < 2.
      lcx_error=>raise( msg = '_contents sheet must have at least 2 columns' ). "#EC NOTEXT
    endif.

    if ls_range-row_max < 2.
      lcx_error=>raise( msg = '_contents sheet must have at least 1 sheet config' ). "#EC NOTEXT
    endif.

    ls_range-col_max = ls_range-col_min + 1. " Consider only 1st 2 columns

    data lt_values type string_table.
    field-symbols <str> type string.
    do ls_range-row_max - ls_range-row_min times.
      lt_values = read_row(
        it_content = it_content
        i_row      = sy-index + 1
        i_colmin   = ls_range-col_min
        i_colmax   = ls_range-col_max ).
      read table lt_values index 2 assigning <str>.
      if <str> is not initial.
        read table lt_values index 1 assigning <str>.
        append <str> to rt_sheets_to_save.
      endif.
    enddo.

  endmethod.  " read_contents.

  method read_exclude.

    data ls_range type ty_range.
    ls_range = clip_range( it_exclude ).

    if ls_range-col_max - ls_range-col_min + 1 < 1.
      lcx_error=>raise( msg = '_exclude sheet must have at least 1 column' ). "#EC NOTEXT
    endif.

    if ls_range-row_max < 2.
      lcx_error=>raise( msg = '_exclude sheet must have at least 1 sheet config' ). "#EC NOTEXT
    endif.

    ls_range-col_max = ls_range-col_min. " Consider only 1st column

    data lt_values type string_table.
    field-symbols <str> type string.
    do ls_range-row_max - ls_range-row_min times.
      lt_values = read_row(
        it_content = it_exclude
        i_row      = sy-index + 1
        i_colmin   = ls_range-col_min
        i_colmax   = ls_range-col_max ).
      read table lt_values index 1 assigning <str>.
      append <str> to rt_sheets_to_exclude.
    enddo.

  endmethod.  " read_exclude.

endclass.

**********************************************************************
* UNIT TESTS
**********************************************************************

class ltcl_workbook_parser_test definition final for testing
  duration short
  risk level harmless.

  private section.
    data mt_dummy_sheet type zexcel_t_cell_data.
    data mt_dummy_contents type zexcel_t_cell_data.
    data mt_dummy_exclude type zexcel_t_cell_data.

    methods setup.
    methods read_row for testing.
    methods is_row_empty for testing.
    methods clip_header for testing.
    methods clip_rows for testing.
    methods read_contents for testing.
    methods convert_sheet for testing.
    methods read_exclude for testing.

endclass.

class ltcl_workbook_parser_test implementation.

  method setup.

    data cell like line of mt_dummy_sheet.

    define _add_cell.
      cell-cell_row = &2.
      cell-cell_column = &3.
      cell-cell_value = &4.
      insert cell into table &1.
    end-of-definition.

    _add_cell mt_dummy_sheet 1 1 '_idx'.
    _add_cell mt_dummy_sheet 1 2 'col1'.
    _add_cell mt_dummy_sheet 1 3 'col2'.
    _add_cell mt_dummy_sheet 1 4 '_comment'.
    _add_cell mt_dummy_sheet 1 6 'aux_col'.

    _add_cell mt_dummy_sheet 2 1 '1'.
    _add_cell mt_dummy_sheet 2 2 'A'.
    _add_cell mt_dummy_sheet 2 3 '10'.
    _add_cell mt_dummy_sheet 2 4 'some comment'.

    _add_cell mt_dummy_sheet 3 1 '2'.
    _add_cell mt_dummy_sheet 3 2 'B'.
    _add_cell mt_dummy_sheet 3 3 '20'.

    _add_cell mt_dummy_sheet 5 2 'some excluded data'.

    " _contents

    _add_cell mt_dummy_contents 1 1 'SheetName'.
    _add_cell mt_dummy_contents 1 2 'ToSave'.
    _add_cell mt_dummy_contents 1 3 'comment'.
    _add_cell mt_dummy_contents 2 1 'Sheet1'.
    _add_cell mt_dummy_contents 2 2 'X'.
    _add_cell mt_dummy_contents 3 1 'Sheet2'.
    _add_cell mt_dummy_contents 3 2 ''.
    _add_cell mt_dummy_contents 4 1 'Sheet3'.
    _add_cell mt_dummy_contents 4 2 'X'.

    " _exclude
    _add_cell mt_dummy_exclude 1 1 'SheetName'.
    _add_cell mt_dummy_exclude 1 2 'Some comment'.
    _add_cell mt_dummy_exclude 2 1 'SheetX'.
    _add_cell mt_dummy_exclude 2 2 'comment 1'.
    _add_cell mt_dummy_exclude 3 1 'SheetY'.

  endmethod.   " setup.

  method read_row.
    data lx type ref to cx_root.
    data lt_act type string_table.
    data lt_exp type string_table.

    clear lt_exp.
    append '_idx' to lt_exp.
    append 'col1' to lt_exp.
    append 'col2' to lt_exp.
    append '_comment' to lt_exp.
    append '' to lt_exp.
    append 'aux_col' to lt_exp.

    try.

      lt_act = lcl_workbook_parser=>read_row(
        it_content = mt_dummy_sheet
        i_row      = 1
        i_colmax   = 6 ).
      cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).

      lt_act = lcl_workbook_parser=>read_row(
        it_content = mt_dummy_sheet
        i_row      = 1 ).
      cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).

      " 2
      lt_act = lcl_workbook_parser=>read_row(
        it_content = mt_dummy_sheet
        i_row      = 5
        i_colmax   = 3 ).

      clear lt_exp.
      append '' to lt_exp.
      append 'some excluded data' to lt_exp.
      append '' to lt_exp.

      cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).

      " 3
      lt_act = lcl_workbook_parser=>read_row(
        it_content = mt_dummy_sheet
        i_row      = 5
        i_colmin   = 2
        i_colmax   = 4 ).

      clear lt_exp.
      append 'some excluded data' to lt_exp.
      append '' to lt_exp.
      append '' to lt_exp.

      cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).
    catch cx_root into lx.
      cl_abap_unit_assert=>fail( 'Unexpected error' ).
    endtry.
  endmethod.  " read_row.

  method is_row_empty.

    data lv_act type abap_bool.

    lv_act = lcl_workbook_parser=>is_row_empty(
      it_content = mt_dummy_sheet
      i_row      = 1
      i_colmax   = 4 ).
    cl_abap_unit_assert=>assert_false( lv_act ).

    lv_act = lcl_workbook_parser=>is_row_empty(
      it_content = mt_dummy_sheet
      i_row      = 5
      i_colmax   = 4 ).
    cl_abap_unit_assert=>assert_false( lv_act ).

    lv_act = lcl_workbook_parser=>is_row_empty(
      it_content = mt_dummy_sheet
      i_row      = 4
      i_colmax   = 4 ).
    cl_abap_unit_assert=>assert_true( lv_act ).

    lv_act = lcl_workbook_parser=>is_row_empty(
      it_content = mt_dummy_sheet
      i_row      = 5
      i_colmin   = 3
      i_colmax   = 4 ).
    cl_abap_unit_assert=>assert_true( lv_act ).

  endmethod.  " is_row_empty.

  method clip_header.
    data lt_head type string_table.
    data ls_range type lcl_workbook_parser=>ty_range.

    append '_idx' to lt_head.
    append 'col1' to lt_head.
    append 'col2' to lt_head.
    append '_comment' to lt_head.
    append '' to lt_head.
    append 'aux_col' to lt_head.

    lcl_workbook_parser=>clip_header(
      exporting it_head = lt_head
      changing  cs_range = ls_range ).

    cl_abap_unit_assert=>assert_equals( act = ls_range-col_min exp = 2 ).
    cl_abap_unit_assert=>assert_equals( act = ls_range-col_max exp = 3 ).

  endmethod.  " clip_header.

  method clip_rows.
    data ls_range type lcl_workbook_parser=>ty_range.

    ls_range-col_min = 2.
    ls_range-col_max = 3.

    lcl_workbook_parser=>clip_rows(
      exporting it_content = mt_dummy_sheet
      changing  cs_range = ls_range ).

    cl_abap_unit_assert=>assert_equals( act = ls_range-row_min exp = 1 ).
    cl_abap_unit_assert=>assert_equals( act = ls_range-row_max exp = 3 ).

  endmethod.  " clip_rows.

  method read_contents.
    data lt_act type string_table.
    data lt_exp type string_table.

    try .
      lt_act = lcl_workbook_parser=>read_contents( mt_dummy_contents ).
    catch lcx_error zcx_excel.
      cl_abap_unit_assert=>fail( 'Unexpected error' ).
    endtry.

    append 'Sheet1' to lt_exp.
    append 'Sheet3' to lt_exp.

    cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).

  endmethod.  " read_contents.

  method convert_sheet.
    data lv_act type string.
    data lf  like cl_abap_char_utilities=>newline value cl_abap_char_utilities=>newline.
    data tab like cl_abap_char_utilities=>horizontal_tab value cl_abap_char_utilities=>horizontal_tab.

    try.
      lv_act = lcl_workbook_parser=>convert_sheet( mt_dummy_sheet ).
    catch lcx_error zcx_excel.
      cl_abap_unit_assert=>fail( 'Unexpected error' ).
    endtry.

    cl_abap_unit_assert=>assert_equals(
      act = lv_act
      exp = |col1{ tab }col2{ lf }A{ tab }10{ lf }B{ tab }20| ).

  endmethod.  " convert_sheet.

  method read_exclude.
    data lt_act type string_table.
    data lt_exp type string_table.

    try .
      lt_act = lcl_workbook_parser=>read_exclude( mt_dummy_exclude ).
    catch lcx_error zcx_excel.
      cl_abap_unit_assert=>fail( 'Unexpected error' ).
    endtry.

    append 'SheetX' to lt_exp.
    append 'SheetY' to lt_exp.

    cl_abap_unit_assert=>assert_equals( act = lt_act exp = lt_exp ).

  endmethod.  " read_exclude.

endclass.

report zmockup_compiler.

**********************************************************************
* EXCEPTIONS
**********************************************************************

class lcx_error definition inheriting from cx_static_check.
  public section.
    interfaces if_t100_message.
    data msg type string read-only.

    methods constructor
      importing msg type string.

    class-methods raise
      importing msg  type string
      raising lcx_error.
endclass.

class lcx_error implementation.

  method constructor.
    super->constructor( ).
    me->msg = msg.
    me->if_t100_message~t100key-msgid = 'SY'. " & & & &
    me->if_t100_message~t100key-msgno = '499'.
    me->if_t100_message~t100key-attr1 = 'MSG'.
  endmethod.

  method raise.
    raise exception type lcx_error
      exporting
        msg = msg.
  endmethod.

endclass.

**********************************************************************
* WORKBOOK PARSER
**********************************************************************

class lcl_unit_tests definition deferred.

class lcl_workbook_parser definition final
  friends lcl_unit_tests.
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
        it_content type zexcel_t_cell_data
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
    lt_worksheets = get_sheets_list( lo_excel ).

    " Check and read content
    field-symbols <ws> like line of lt_worksheets.
    read table lt_worksheets assigning <ws> with key title = content_sheet_name.
    if sy-subrc is not initial.
      lcx_error=>raise( msg = 'Workbook does not contain _contents sheet' ). "#EC NOTEXT
    endif.

    try.
      data lt_sheets_to_save type string_table.
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
    do ls_range-row_max times. " starts from 1 always
      lt_values = read_row(
        it_content     = it_content
        it_date_styles = it_date_styles
        i_row          = sy-index
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

    cs_range-row_min = 1.
    cs_range-row_max = 1.

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

    lt_head = read_row(
      it_content = it_content
      i_row      = 1 ).
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
        it_content = it_content
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

endclass.

**********************************************************************
* APP
**********************************************************************

class lcl_app definition final.
  public section.

    constants c_file_mask type string value '*.xlsx'.

    methods constructor
      importing
        iv_dir      type string
        iv_include  type string optional
        iv_mime_key type string
        iv_do_watch type abap_bool default abap_false
      raising lcx_error zcx_w3mime_error.

    methods run
      raising lcx_error zcx_w3mime_error.

    methods process_xls
      importing
        iv_path type string
      raising lcx_error zcx_w3mime_error.

    methods update_mime
      raising lcx_error zcx_w3mime_error.

    methods handle_changed for event changed of zcl_w3mime_poller importing changed_list.
    methods handle_error   for event error of zcl_w3mime_poller importing error_text.

  private section.
    data:
          mo_poller   type ref to zcl_w3mime_poller,
          mo_zip      type ref to zcl_w3mime_zip_writer,
          mv_dir      type string,
          mv_include  type string,
          mv_mime_key type wwwdata-objid.

    class-methods fmt_dt
      importing iv_ts         type zcl_w3mime_poller=>ty_file_state-timestamp
      returning value(rv_str) type string.

    methods process_includes
      importing
        iv_subdir type string default ''
      returning value(rv_num) type i
      raising lcx_error zcx_w3mime_error.

endclass.

class lcl_app implementation.
  method constructor.
    data lv_sep type c.

    if iv_mime_key is initial.
      lcx_error=>raise( 'iv_mime_key must be specified' ). "#EC NOTEXT
    endif.

    if iv_dir is initial.
      lcx_error=>raise( 'iv_dir must be specified' ). "#EC NOTEXT
    endif.

    create object mo_zip.
    mv_mime_key = iv_mime_key.
    mv_dir      = iv_dir.
    mv_include  = iv_include.

    cl_gui_frontend_services=>get_file_separator( changing file_separator = lv_sep ).
    if substring( val = mv_dir off = strlen( mv_dir ) - 1 len = 1 ) <> lv_sep.
      mv_dir = mv_dir && lv_sep.
    endif.

    if mv_include is not initial and substring( val = mv_include off = strlen( mv_include ) - 1 len = 1 ) <> lv_sep.
      mv_include = mv_include && lv_sep.
    endif.

    data lt_targets type zcl_w3mime_poller=>tt_target.
    field-symbols: <ft> like line of lt_targets.

    if iv_do_watch = abap_true.
      append initial line to lt_targets assigning <ft>.
      <ft>-directory = mv_dir.
      <ft>-filter    = c_file_mask.

      create object mo_poller
        exporting
          it_targets  = lt_targets
          iv_interval = 1.
      set handler me->handle_changed for mo_poller.
      set handler me->handle_error for mo_poller.
    endif.

  endmethod.

  method run.
    data lt_files type zcl_w3mime_fs=>tt_files.
    field-symbols <f> like line of lt_files.

    write: / 'Processing directory:', mv_dir. "#EC NOTEXT
    lt_files = zcl_w3mime_fs=>read_dir(
      iv_dir    = mv_dir
      iv_filter = c_file_mask ).

    data lv_num_files type i.
    lv_num_files = lines( lt_files ).

    loop at lt_files assigning <f>.
      write: /3 <f>-filename.

      cl_progress_indicator=>progress_indicate(
        i_text               = | Processing XL { sy-tabix } / { lv_num_files }: { <f>-filename }| "#EC NOTEXT
        i_processed          = sy-tabix
        i_total              = lv_num_files
        i_output_immediately = abap_true ).

      process_xls( mv_dir && <f>-filename ).
    endloop.

    if mv_include is not initial.
      write: / 'Processing includes...'. "#EC NOTEXT
      lv_num_files = process_includes( ).
      write: lv_num_files, 'found'.
    endif.

    update_mime( ).
    write: / 'MIME object updated:', mv_mime_key. "#EC NOTEXT

    if mo_poller is bound.
      uline.
      write: / 'Start polling ...'. "#EC NOTEXT
      mo_poller->start( ).
    endif.

  endmethod.

  method process_xls.
    data lv_blob type xstring.
    lv_blob = zcl_w3mime_fs=>read_file_x( iv_path ).

    data lt_mocks type lcl_workbook_parser=>tt_mocks.
    field-symbols <mock> like line of lt_mocks.
    lt_mocks = lcl_workbook_parser=>parse( lv_blob ).

    data lv_folder_name type string.
    zcl_w3mime_fs=>parse_path(
      exporting
        iv_path = iv_path
      importing
        ev_filename = lv_folder_name ).
    lv_folder_name = to_upper( lv_folder_name ).

    loop at lt_mocks assigning <mock>.
      <mock>-name = lv_folder_name && '/' && to_lower( <mock>-name ) && '.txt'.
      mo_zip->add(
        iv_filename = <mock>-name
        iv_data     = <mock>-data ).
    endloop.

  endmethod.  " process_xls.

  method process_includes.
    data lt_files type zcl_w3mime_fs=>tt_files.
    field-symbols <f> like line of lt_files.

    data:
          lv_root          type string,
          lv_relative_path type string,
          lv_zip_path      type string,
          lv_file_path     type string.

    lv_root  = zcl_w3mime_fs=>join_path( iv_p1 = mv_include iv_p2 = iv_subdir ).
    lt_files = zcl_w3mime_fs=>read_dir( lv_root ).

    cl_progress_indicator=>progress_indicate( " Hmmm, refactor ?
      i_text               = | Processing includes: { iv_subdir }| "#EC NOTEXT
      i_processed          = 1
      i_total              = 1
      i_output_immediately = abap_true ).

    loop at lt_files assigning <f>.
      lv_relative_path = zcl_w3mime_fs=>join_path( iv_p1 = iv_subdir iv_p2 = |{ <f>-filename }| ).

      if <f>-isdir is initial.
        lv_zip_path  = replace(
          val = lv_relative_path
          sub = zcl_w3mime_fs=>c_sep
          with = '/'
          occ = 0 ).
        lv_file_path = zcl_w3mime_fs=>join_path(
          iv_p1 = lv_root
          iv_p2 = |{ <f>-filename }| ).

        mo_zip->addx(
          iv_filename = lv_zip_path
          iv_xdata    = zcl_w3mime_fs=>read_file_x( lv_file_path )  ).

        rv_num = rv_num + 1.
      else.
        rv_num = rv_num + process_includes( lv_relative_path ).
      endif.
    endloop.

  endmethod.  " process_includes.

  method update_mime.
    data lv_blob type xstring.
    lv_blob = mo_zip->get_blob( ).
    zcl_w3mime_storage=>update_object_x(
      iv_key  = mv_mime_key
      iv_data = lv_blob ).
  endmethod.  " update_mime.

  method fmt_dt.
    data ts type char14.
    ts = iv_ts.
    rv_str = |{ ts+0(4) }-{ ts+4(2) }-{ ts+6(2) } |
          && |{ ts+8(2) }:{ ts+10(2) }:{ ts+12(2) }|.
  endmethod.  " format_dt.

  method handle_changed.
    data lt_list like changed_list.
    data lt_filenames type string_table.
    data lstr type string.
    data lext type string.
    field-symbols <i> like line of lt_list.

    loop at changed_list assigning <i>.
      zcl_w3mime_fs=>parse_path( exporting iv_path = <i>-path importing ev_filename = lstr ev_extension = lext ).
      check not ( strlen( lstr ) >= 2 and substring( val = lstr len = 2 ) = '~$' ).
      append <i> to lt_list.
      lstr = lstr && lext.
      append lstr to lt_filenames.
    endloop.

    if lines( lt_list ) = 0.
      return.
    endif.

    lstr = |{ fmt_dt( <i>-timestamp ) }: { concat_lines_of( table = lt_filenames sep = ', ' ) }|.
    write / lstr.

    data lx type ref to cx_static_check.
    try.
      loop at lt_list assigning <i>.
        process_xls( <i>-path ).
      endloop.
      update_mime( ).
    catch lcx_error zcx_w3mime_error into lx.
      lstr = lx->get_text( ).
      message lstr type 'E'.
    endtry.

  endmethod.  "handle_changed

  method handle_error.
    message error_text type 'E'.
  endmethod.  " handle_error.

endclass.

**********************************************************************
* UNIT TESTS
**********************************************************************

class lcl_unit_tests definition final for testing
  duration short
  risk level harmless.

  private section.
    data mt_dummy_sheet type zexcel_t_cell_data.
    data mt_dummy_contents type zexcel_t_cell_data.

    methods setup.
    methods read_row for testing.
    methods is_row_empty for testing.
    methods clip_header for testing.
    methods clip_rows for testing.
    methods read_contents for testing.
    methods convert_sheet for testing.

endclass.

class lcl_unit_tests implementation.

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

endclass.

**********************************************************************
* ENTRY POINT
**********************************************************************
constants:
  GC_DIR_PARAM_NAME TYPE CHAR20 VALUE 'ZMOCKUP_COMPILER_DIR',
  GC_OBJ_PARAM_NAME TYPE CHAR20 VALUE 'ZMOCKUP_COMPILER_OBJ',
  GC_INC_PARAM_NAME TYPE CHAR20 VALUE 'ZMOCKUP_COMPILER_INC'.

form main using pv_srcdir pv_incdir pv_mimename pv_watch.
  data lo_app type ref to lcl_app.
  data lx type ref to cx_static_check.
  data l_str type string.

  if pv_srcdir is initial.
    message 'Source directory parameter is mandatory' type 'S' display like 'E'. "#EC NOTEXT
  endif.

  if pv_mimename is initial.
    message 'MIME object name is mandatory' type 'S' display like 'E'. "#EC NOTEXT
  endif.

  try.
    create object lo_app
      exporting
        iv_dir      = |{ pv_srcdir }|
        iv_include  = |{ pv_incdir }|
        iv_mime_key = |{ pv_mimename }|
        iv_do_watch = pv_watch.
    lo_app->run( ).

    " For debug
*    zcl_w3mime_utils=>download(
*        iv_filename = 'c:\sap\test.zip'
*        iv_key = 'ZMLB_TEST' ).

  catch lcx_error zcx_w3mime_error into lx.
    l_str = lx->get_text( ).
    message l_str type 'E'.
  endtry.

endform.

form f4_include_path changing c_path type char255.
  c_path = zcl_w3mime_fs=>choose_dir_dialog( ).
  if c_path is not initial.
    set parameter id GC_INC_PARAM_NAME field c_path.
  endif.
endform.

form f4_srcdir_path changing c_path type char255.
  c_path = zcl_w3mime_fs=>choose_dir_dialog( ).
  if c_path is not initial.
    set parameter id GC_DIR_PARAM_NAME field c_path.
  endif.
endform.

form f4_mime_path changing c_path.
  c_path = zcl_w3mime_storage=>choose_mime_dialog( ).
  if c_path is not initial.
    set parameter id GC_OBJ_PARAM_NAME field c_path.
  endif.
endform.

**********************************************************************
* SELECTION SCREEN
**********************************************************************

selection-screen begin of block b1 with frame title txt_b1.

selection-screen begin of line.
selection-screen comment (24) t_mime for field p_mime.
parameters p_mime type w3objid visible length 40.
selection-screen comment (24) t_mime2.
selection-screen end of line.

selection-screen begin of line.
selection-screen comment (24) t_dir for field p_dir.
parameters p_dir type char255 visible length 40.
selection-screen end of line.

selection-screen begin of line.
selection-screen comment (24) t_inc for field p_inc.
parameters p_inc type char255 visible length 40.
selection-screen comment (24) t_inc2.
selection-screen end of line.

selection-screen begin of line.
selection-screen comment (24) t_watch for field p_watch.
parameters p_watch type abap_bool as checkbox.
selection-screen end of line.

selection-screen end of block b1.

initialization.
  txt_b1   = 'Compile parameters'.        "#EC NOTEXT
  t_dir    = 'Source directory'.          "#EC NOTEXT
  t_inc    = 'Includes directory'.        "#EC NOTEXT
  t_inc2   = '  (optional)'.              "#EC NOTEXT
  t_mime   = 'Target W3MI object'.        "#EC NOTEXT
  t_mime2  = '  (must be existing one)'.  "#EC NOTEXT
  t_watch  = 'Keep watching source'.      "#EC NOTEXT

  get parameter id GC_DIR_PARAM_NAME field p_dir.
  get parameter id GC_INC_PARAM_NAME field p_inc.
  get parameter id GC_OBJ_PARAM_NAME field p_mime.

at selection-screen on value-request for p_dir.
  perform f4_srcdir_path changing p_dir.

at selection-screen on value-request for p_mime.
  perform f4_mime_path changing p_mime.

at selection-screen on value-request for p_inc.
  perform f4_include_path changing p_inc.

at selection-screen on p_dir.
  if p_dir is not initial.
    set parameter id GC_DIR_PARAM_NAME field p_dir.
  endif.

at selection-screen on p_inc.
  if p_inc is not initial.
    set parameter id GC_INC_PARAM_NAME field p_inc.
  endif.

at selection-screen on p_mime.
  if p_mime is not initial.
    set parameter id GC_OBJ_PARAM_NAME field p_mime.
  endif.

start-of-selection.
  perform main using p_dir p_inc p_mime p_watch.

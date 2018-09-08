**********************************************************************
* APP
**********************************************************************

class lcl_app definition final.
  public section.

    constants c_xlfile_mask type string value '*.xlsx'.

    methods constructor
      importing
        iv_dir      type string
        iv_include  type string optional
        iv_mime_key type string
        iv_rebuild  type abap_bool default abap_false
        iv_do_watch type abap_bool default abap_false
      raising lcx_error zcx_w3mime_error.

    methods run
      raising lcx_error zcx_w3mime_error.

    methods handle_changed for event changed of zcl_w3mime_poller importing changed_list.
    methods handle_error   for event error of zcl_w3mime_poller importing error_text.

  private section.

    types:
      begin of ty_src_timestamp,
        type      type char1,
        src_file  type string,
        timestamp type char14,
      end of ty_src_timestamp,
      tt_src_timestamp type standard table of ty_src_timestamp with key type src_file.

    data:
          mv_do_watch    type abap_bool,
          mt_src_ts      type tt_src_timestamp,
          mo_poller      type ref to zcl_w3mime_poller,
          mo_zip         type ref to zcl_w3mime_zip_writer,
          mv_dir         type string,
          mv_include_dir type string,
          mv_mime_key    type wwwdata-objid,
          mt_inc_dirs    type string_table.

    methods process_excel
      importing
        iv_path type string
      raising lcx_error zcx_w3mime_error.

    methods process_include
      importing
        iv_path type string
      raising lcx_error zcx_w3mime_error.

    methods update_mime_object
      raising lcx_error zcx_w3mime_error.

    class-methods fmt_dt
      importing iv_ts         type zcl_w3mime_poller=>ty_file_state-timestamp
      returning value(rv_str) type string.

    class-methods is_tempfile
      importing iv_filename type string
      returning value(rv_yes) type abap_bool.

    methods update_src_timestamp
      importing
        iv_type      type char1
        iv_filename  type string
        iv_timestamp type zcl_w3mime_poller=>ty_file_state-timestamp
      returning value(rv_updated) type abap_bool.

    methods process_excel_dir
      raising lcx_error zcx_w3mime_error.

    methods process_includes_dir
      importing
        iv_inc_dir type string
      returning value(rv_num) type i
      raising lcx_error zcx_w3mime_error.

    methods write_meta.
    methods read_meta.
    methods start_watcher
      raising lcx_error zcx_w3mime_error.

endclass.

class lcl_app implementation.

  method constructor.
    if iv_mime_key is initial.
      lcx_error=>raise( 'iv_mime_key must be specified' ). "#EC NOTEXT
    endif.

    if iv_dir is initial.
      lcx_error=>raise( 'iv_dir must be specified' ). "#EC NOTEXT
    endif.

    if iv_include is not initial and zcl_w3mime_fs=>path_is_relative( iv_to = iv_dir iv_from = iv_include ) = abap_true.
      lcx_error=>raise( 'iv_dir cannot be relevant to iv_include' ). "#EC NOTEXT
    endif.

    data lo_zip type ref to cl_abap_zip.
    if iv_rebuild = abap_false.
      try.
        create object lo_zip.
        lo_zip->load(
          exporting
            zip = zcl_w3mime_storage=>read_object_x( |{ iv_mime_key }| )
          exceptions others = 4 ).
      catch zcx_w3mime_error.
        sy-subrc = 4. " Ignore errors
      endtry.
    endif.

    create object mo_zip exporting io_zip = lo_zip.
    if lo_zip is bound.
      read_meta( ).
    endif.

    mv_mime_key    = iv_mime_key.
    mv_dir         = zcl_w3mime_fs=>path_ensure_dir_tail( iv_dir ).
    mv_include_dir = zcl_w3mime_fs=>path_ensure_dir_tail( iv_include ).
    mv_do_watch    = iv_do_watch.

  endmethod.

  method run.
    write: / 'Processing directory:', mv_dir. "#EC NOTEXT
    process_excel_dir( ).

    data lv_num_files type i.
    if mv_include_dir is not initial.
      write: /.
      write: / 'Processing includes...'. "#EC NOTEXT
      lv_num_files = process_includes_dir( mv_include_dir ).
      write: /3 'Found (changed):', lv_num_files.
    endif.

    write: /.
    if mo_zip->is_dirty( ) = abap_true.
      write_meta( ).
      update_mime_object( ).
      write: / 'MIME object updated:', mv_mime_key. "#EC NOTEXT
    else.
      write: / 'Changes not detected, update skipped'. "#EC NOTEXT
    endif.

    if mv_do_watch = abap_true.
      uline.
      write: / 'Start polling ...'. "#EC NOTEXT
      start_watcher( ).
    endif.

  endmethod.

  method process_excel_dir.
    data lt_files type zcl_w3mime_fs=>tt_files.
    field-symbols <f> like line of lt_files.

    lt_files = zcl_w3mime_fs=>read_dir(
      iv_dir    = mv_dir
      iv_filter = c_xlfile_mask ).

    data lv_num_files type i.
    lv_num_files = lines( lt_files ).

    loop at lt_files assigning <f>.
      if is_tempfile( |{ <f>-filename }| ) = abap_true.
        lv_num_files = lv_num_files - 1.
        continue.
      endif.

      cl_progress_indicator=>progress_indicate(
        i_text               = | Processing XL { sy-tabix } / { lv_num_files }: { <f>-filename }| "#EC NOTEXT
        i_processed          = sy-tabix
        i_total              = lv_num_files
        i_output_immediately = abap_true ).

      data lv_updated   type abap_bool.
      lv_updated = update_src_timestamp(
        iv_type      = 'X'
        iv_filename  = |{ <f>-filename }|
        iv_timestamp = <f>-writedate && <f>-writetime ).

      write: /3 <f>-filename.
      if lv_updated = abap_false.
        write at 50 '<skipped, meta timestamp unchanged>'. "#EC NOTEXT
        continue.
      endif.

      process_excel( mv_dir && <f>-filename ).
    endloop.
  endmethod.

  method process_excel.
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

  endmethod.  " process_excel.

  method process_include.
    if zcl_w3mime_fs=>path_is_relative( iv_to = iv_path iv_from = mv_include_dir ) <> abap_true.
      lcx_error=>raise( 'Unexpected include path' ).
    endif.

    data lv_relative_path type string.
    lv_relative_path = zcl_w3mime_fs=>path_relative(
      iv_from = mv_include_dir
      iv_to   = iv_path ).

    lv_relative_path = replace(
      val = lv_relative_path
      sub = zcl_w3mime_fs=>c_sep
      with = '/'
      occ = 0 ).

    mo_zip->addx(
      iv_filename = lv_relative_path
      iv_xdata    = zcl_w3mime_fs=>read_file_x( iv_path )  ).

  endmethod.

  method process_includes_dir.
    data lt_files type zcl_w3mime_fs=>tt_files.
    data lv_subdir type string.
    field-symbols <f> like line of lt_files.

    append iv_inc_dir to mt_inc_dirs. " remember for future watch targets

    lv_subdir  = zcl_w3mime_fs=>path_relative( iv_from = mv_include_dir iv_to = iv_inc_dir ).
    if lv_subdir is not initial.
      write /3 lv_subdir.
    endif.
    cl_progress_indicator=>progress_indicate(
      i_text               = | Processing includes: { lv_subdir }| "#EC NOTEXT
      i_processed          = 1
      i_total              = 1
      i_output_immediately = abap_true ).

    lt_files = zcl_w3mime_fs=>read_dir( iv_inc_dir ).
    loop at lt_files assigning <f>.
      data lv_full_path type string.
      lv_full_path = zcl_w3mime_fs=>path_join(
        iv_p1 = iv_inc_dir
        iv_p2 = |{ <f>-filename }| ).

      if <f>-isdir is initial.
        data lv_updated type abap_bool.
        lv_updated = update_src_timestamp(
          iv_type      = 'I'
          iv_filename  = zcl_w3mime_fs=>path_relative( iv_to = lv_full_path iv_from = mv_include_dir )
          iv_timestamp = <f>-writedate && <f>-writetime ).
        check lv_updated = abap_true.

        " TODO count skipped
        process_include( lv_full_path ).
        rv_num = rv_num + 1.
      else.
        rv_num = rv_num + process_includes_dir( lv_full_path ).
      endif.
    endloop.

  endmethod.  " process_includes_dir.

  method update_mime_object.
    data lv_blob type xstring.
    lv_blob = mo_zip->get_blob( ).
    zcl_w3mime_storage=>update_object_x(
      iv_key  = mv_mime_key
      iv_data = lv_blob ).
  endmethod.  " update_mime_object.

  method fmt_dt.
    data ts type char14.
    ts = iv_ts.
    rv_str = |{ ts+0(4) }-{ ts+4(2) }-{ ts+6(2) } |
          && |{ ts+8(2) }:{ ts+10(2) }:{ ts+12(2) }|.
  endmethod.  " format_dt.

  method handle_changed.
    data lx           type ref to cx_static_check.
    data l_msg        type string.
    data lt_filenames type string_table.
    field-symbols <i> like line of changed_list.

    try.
      loop at changed_list assigning <i>.
        data l_fname type string.
        data l_ext type string.
        zcl_w3mime_fs=>parse_path(
          exporting
            iv_path = <i>-path
          importing
            ev_filename  = l_fname
            ev_extension = l_ext ).
        l_fname = l_fname && l_ext.

        check is_tempfile( l_fname ) <> abap_true.
        append l_fname to lt_filenames.

        if mv_include_dir is not initial
          and zcl_w3mime_fs=>path_is_relative( iv_to = <i>-path iv_from = mv_include_dir ) = abap_true.
          process_include( <i>-path ).
          update_src_timestamp(
            iv_type      = 'I'
            iv_filename  = zcl_w3mime_fs=>path_relative( iv_to = <i>-path iv_from = mv_include_dir )
            iv_timestamp = <i>-timestamp ).
        else.
          process_excel( <i>-path ).
          update_src_timestamp(
            iv_type      = 'X'
            iv_filename  = l_fname
            iv_timestamp = <i>-timestamp ).
        endif.
      endloop.

      if lines( lt_filenames ) = 0.
        return.
      endif.

      write_meta( ).
      update_mime_object( ).
    catch lcx_error zcx_w3mime_error into lx.
      l_msg = lx->get_text( ).
      message l_msg type 'E'.
    endtry.

    " Report result
    l_msg = |{ fmt_dt( <i>-timestamp ) }: { concat_lines_of( table = lt_filenames sep = ', ' ) }|.
    write / l_msg.

  endmethod.  "handle_changed

  method handle_error.
    message error_text type 'E'.
  endmethod.  " handle_error.

  method is_tempfile.
    rv_yes = boolc( strlen( iv_filename ) >= 2 and substring( val = iv_filename len = 2 ) = '~$' ).
  endmethod.

  method update_src_timestamp.
    field-symbols <m> like line of mt_src_ts.
    read table mt_src_ts with key type = iv_type src_file = iv_filename assigning <m>.
    if sy-subrc is not initial. " new file
      append initial line to mt_src_ts assigning <m>.
      <m>-type     = iv_type.
      <m>-src_file = iv_filename.
    endif.

    if <m>-timestamp <> iv_timestamp.
      <m>-timestamp = iv_timestamp.
      rv_updated = abap_true.
    endif.
  endmethod.

  method write_meta.
    constants:
      lc_tab like cl_abap_char_utilities=>horizontal_tab value cl_abap_char_utilities=>horizontal_tab,
      lc_lf like cl_abap_char_utilities=>newline value cl_abap_char_utilities=>newline.

    data lt_lines type string_table.
    data ltmp type string.

    " TODO replace with TEXT2TAB
    field-symbols <m> like line of mt_src_ts.
    loop at mt_src_ts assigning <m>.
      ltmp = <m>-type && lc_tab && <m>-src_file && lc_tab && <m>-timestamp.
      append ltmp to lt_lines.
    endloop.

    mo_zip->add(
      iv_filename = '.meta/src_files'
      iv_data     = concat_lines_of( table = lt_lines sep = lc_lf ) ).
  endmethod.

  method read_meta.
    constants:
      lc_tab like cl_abap_char_utilities=>horizontal_tab value cl_abap_char_utilities=>horizontal_tab,
      lc_lf like cl_abap_char_utilities=>newline value cl_abap_char_utilities=>newline.
    data lv_src_files_blob type string.

    try.
      lv_src_files_blob = mo_zip->read( '.meta/src_files' ).
    catch zcx_w3mime_error.
      return. " Ignore errors
    endtry.

    " TODO replace with TEXT2TAB
    data lt_lines type string_table.
    split lv_src_files_blob at lc_lf into table lt_lines.

    data ls_src_ts like line of mt_src_ts.
    field-symbols <i> type string.
    loop at lt_lines assigning <i>.
      split <i> at lc_tab into ls_src_ts-type ls_src_ts-src_file ls_src_ts-timestamp.
      append ls_src_ts to mt_src_ts.
    endloop.

  endmethod.

  method start_watcher.

    data lt_targets type zcl_w3mime_poller=>tt_target.
    field-symbols: <ft> like line of lt_targets.

    append initial line to lt_targets assigning <ft>.
    <ft>-directory = mv_dir.
    <ft>-filter    = c_xlfile_mask.

    field-symbols <i> like line of mt_inc_dirs.
    loop at mt_inc_dirs assigning <i>.
      append initial line to lt_targets assigning <ft>.
      <ft>-directory = <i>.
      <ft>-filter    = '*.*'.
    endloop.

    create object mo_poller
      exporting
        it_targets  = lt_targets
        iv_interval = 1.
    set handler me->handle_changed for mo_poller.
    set handler me->handle_error for mo_poller.

    mo_poller->start( ).

  endmethod.

endclass.

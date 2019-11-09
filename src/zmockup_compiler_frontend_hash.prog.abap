class lcl_frontend_hash definition final.
  public section.
    constants c_tmp_filename type string value 'mockup_compiler_frontend_hash.tmp'.

    types:
      begin of ty_hash,
        path type string,
        hash type string,
      end of ty_hash.

    methods constructor.
    methods is_working
      returning
        value(rv_yes) type abap_bool.
    methods get_hash
      importing
        iv_file type string
      returning
        value(rv_hash) type string.

  private section.
    data mv_workdir type string.
    data mv_tmp_path type string.
    data mv_is_working type abap_bool.
    data mv_is_checked type abap_bool.

    methods run_get_filehash
      importing
        iv_file type string optional
      returning
        value(rv_stdout) type string.

    class-methods parse_hash_stdout
      importing
        iv_stdout type string
      returning
        value(rs_hash) type ty_hash.

    methods check_is_working.

endclass.

class lcl_frontend_hash implementation.
  method constructor.
    cl_gui_frontend_services=>get_sapgui_workdir( changing sapworkdir = mv_workdir ).
    cl_gui_cfw=>flush( ).
    mv_tmp_path = mv_workdir && '\' && c_tmp_filename.
    check_is_working( ).
  endmethod.

  method get_hash.

    if mv_is_working = abap_false.
      return.
    endif.

    assert abap_true = cl_gui_frontend_services=>file_exist( iv_file ).

    data lv_stdout type string.
    data ls_hash type ty_hash.
    lv_stdout = run_get_filehash( iv_file ).
    ls_hash   = parse_hash_stdout( lv_stdout ).

    if ls_hash-path <> iv_file.
      return. " Hmmm
    endif.

    rv_hash = to_lower( ls_hash-hash ).

  endmethod.

  method run_get_filehash.

    data rc type i.
    cl_gui_frontend_services=>file_delete(
      exporting
        filename = mv_tmp_path
      changing
        rc = rc
      exceptions
        others = 1 ).

    if iv_file is initial. " Check ?
      cl_gui_frontend_services=>execute(
        exporting
          application = 'powershell'
          parameter = |-command "get-filehash -? > '{ mv_tmp_path }'"|
          minimized = 'X'
          synchronous = 'X'
        exceptions
          others = 1 ).
    else.
      cl_gui_frontend_services=>execute(
        exporting
          application = 'powershell'
          parameter = |-command "get-filehash -Algorithm SHA1 '{ iv_file }' \| Format-List  > '{ mv_tmp_path }'"|
          minimized = 'X'
          synchronous = 'X'
        exceptions
          others = 1 ).
    endif.
    if sy-subrc <> 0.
      return.
    endif.

    data lv_size type i.
    data lt_data type table of raw255.
    cl_gui_frontend_services=>gui_upload(
      exporting
        filename   = mv_tmp_path
        filetype   = 'BIN'
      importing
        filelength = lv_size
      changing
        data_tab   = lt_data
      exceptions
        others = 1 ).
    if sy-subrc <> 0.
      return.
    endif.

    data outputx type xstring.
    call function 'SCMS_BINARY_TO_XSTRING'
      exporting
        input_length = lv_size
      importing
        buffer       = outputx
      tables
        binary_tab   = lt_data.

    data lo_conv type ref to cl_abap_conv_in_ce.
    try.
      lv_size = xstrlen( outputx ).
      lo_conv = cl_abap_conv_in_ce=>create(
          input    = outputx
          encoding = |{ cl_lxe_constants=>c_sap_codepage_utf16 }| ).
      lo_conv->read(
        exporting
          n = lv_size
        importing
          data = rv_stdout ).
    catch cx_root.
      return.
    endtry.

  endmethod.

  method parse_hash_stdout.

    data lt_lines type string_table.
    field-symbols <i> like line of lt_lines.

    split iv_stdout at cl_abap_char_utilities=>cr_lf into table lt_lines.

    loop at lt_lines assigning <i>.
      if <i> cp 'Path *'.
        rs_hash-path = substring_after( val = <i> sub = ':' ).
        condense rs_hash-path.
      endif.
      if <i> cp 'Hash *'.
        rs_hash-hash = substring_after( val = <i> sub = ':' ).
        condense rs_hash-hash.
      endif.
    endloop.
  endmethod.

  method is_working.
    if mv_is_checked = abap_false.
      check_is_working( ).
    endif.
    rv_yes = mv_is_working.
  endmethod.

  method check_is_working.

    data lv_stdout type string.

    mv_is_working = abap_false.
    lv_stdout = run_get_filehash( ). " output help
    mv_is_checked = abap_true.

    if lv_stdout is initial.
      return. " false
    endif.

    if 0 < find( val = lv_stdout sub = 'Get-FileHash -LiteralPath' ). " Part of help detected
      mv_is_working = abap_true.
    endif.

  endmethod.

endclass.

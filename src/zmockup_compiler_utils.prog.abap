class lcl_utils definition final.
  public section.

  class-methods fmt_dt
    importing
      iv_ts         type zcl_w3mime_poller=>ty_file_state-timestamp
    returning
      value(rv_str) type string.

  class-methods is_tempfile
    importing
      iv_filename type string
    returning
      value(rv_yes) type abap_bool.

  class-methods sha1
    importing
      iv_data type xstring
    returning
      value(rv_hash) type hash160
    raising
      lcx_error.
endclass.

class lcl_utils implementation.

  method fmt_dt.
    data ts type char14.
    ts = iv_ts.
    rv_str = |{ ts+0(4) }-{ ts+4(2) }-{ ts+6(2) } |
          && |{ ts+8(2) }:{ ts+10(2) }:{ ts+12(2) }|.
  endmethod.

  method is_tempfile.
    rv_yes = boolc( strlen( iv_filename ) >= 2 and substring( val = iv_filename len = 2 ) = '~$' ).
  endmethod.

  method sha1.

    data lv_hash type hash160.

    call function 'CALCULATE_HASH_FOR_RAW'
      exporting
        data           = iv_data
      importing
        hash           = lv_hash
      exceptions
        unknown_alg    = 1
        param_error    = 2
        internal_error = 3
        others         = 4.
    if sy-subrc <> 0.
      lcx_error=>raise( 'sha1 calculation error' ).
    endif.

    rv_hash = to_lower( lv_hash ).

  endmethod.

endclass.

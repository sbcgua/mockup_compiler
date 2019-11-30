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

    data lo_iter type ref to cl_object_collection_iterator.
    field-symbols <w> like line of mt_worksheets.

    if mt_worksheets is initial.
      lo_iter = mo_excel->get_worksheets_iterator( ).
      while lo_iter->has_next( ) is not initial.
        append initial line to mt_worksheets assigning <w>.
        <w>-worksheet ?= lo_iter->get_next( ).
        <w>-title      = <w>-worksheet->get_title( ).
      endwhile.
    endif.

    loop at mt_worksheets assigning <w>.
      append <w>-title to rt_sheet_names.
    endloop.

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
* UNIT TEST
**********************************************************************

class ltcl_excel_abap2xlsx definition final
  for testing
  risk level harmless
  duration short.
  private section.

    types:
      begin of ty_content_sample,
        contents type lif_excel=>tt_sheet_content,
        excludes type lif_excel=>tt_sheet_content,
        simple type lif_excel=>tt_sheet_content,
        with_dates type lif_excel=>tt_sheet_content,
      end of ty_content_sample.

    methods load for testing raising zcx_w3mime_error lcx_excel.
    methods assert_sheet_names importing ii_excel type ref to lif_excel.
    methods assert_styles importing ii_excel type ref to lif_excel raising lcx_excel.
    methods assert_content importing ii_excel type ref to lif_excel raising lcx_excel.
    methods get_expected_content returning value(rs_content_samples) type ty_content_sample.

    methods clear_styles changing ct_content type lif_excel=>tt_sheet_content.
endclass.

class ltcl_excel_abap2xlsx implementation.

  method load.

    data xstr type xstring.
    xstr = zcl_w3mime_storage=>read_object_x( 'ZMOCKUP_COMPILER_UNIT_TEST' ).

    data li_excel type ref to lif_excel.
    li_excel = lcl_excel_abap2xlsx=>load( xstr ).

    assert_sheet_names( li_excel ).
    assert_styles( li_excel ).
    assert_content( li_excel ).

  endmethod.

  method assert_sheet_names.

    data lt_exp type string_table.

    append '_contents' to lt_exp.
    append '_exclude' to lt_exp.
    append 'Sheet1' to lt_exp.
    append 'Sheet2' to lt_exp.
    append 'Sheet3' to lt_exp.
    append 'Sheet4' to lt_exp.

    cl_abap_unit_assert=>assert_equals(
      act = ii_excel->get_sheet_names( )
      exp = lt_exp ).

  endmethod.

  method assert_styles.

    data lt_act type lif_excel=>tt_styles.
    data style like line of lt_act.

    lt_act = ii_excel->get_styles( ).
    read table lt_act into style index 7.

    cl_abap_unit_assert=>assert_equals(
      act = lines( lt_act )
      exp = 8 ).
    cl_abap_unit_assert=>assert_equals(
      act = style-format
      exp = 'mm-dd-yy' ).

  endmethod.

  method clear_styles.
    field-symbols <c> like line of ct_content.
    loop at ct_content assigning <c>.
      clear <c>-cell_style.
    endloop.
  endmethod.

  method assert_content.

    data ls_samples type ty_content_sample.
    data lt_act type lif_excel=>tt_sheet_content.
    ls_samples = get_expected_content( ).

    clear_styles( changing ct_content = ls_samples-contents ).
    clear_styles( changing ct_content = ls_samples-excludes ).
    clear_styles( changing ct_content = ls_samples-simple ).
    clear_styles( changing ct_content = ls_samples-with_dates ).

    lt_act = ii_excel->get_sheet_content( '_contents' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-contents ).

    lt_act = ii_excel->get_sheet_content( '_exclude' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-excludes ).

    lt_act = ii_excel->get_sheet_content( 'Sheet1' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-simple ).

    lt_act = ii_excel->get_sheet_content( 'Sheet2' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-with_dates ).

    lt_act = ii_excel->get_sheet_content( 'Sheet3' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-with_dates ).

    lt_act = ii_excel->get_sheet_content( 'Sheet4' ).
    clear_styles( changing ct_content = lt_act ).
    cl_abap_unit_assert=>assert_equals( act = lt_act exp = ls_samples-with_dates ).

  endmethod.

  method get_expected_content.
    data lt_content type lif_excel=>tt_sheet_content.
    field-symbols <i> like line of lt_content.

    define _add_cell.
      append initial line to lt_content assigning <i>.
      <i>-cell_row     = &1.
      <i>-cell_column  = &2.
      <i>-cell_value   = &3.
      <i>-cell_coords  = &4.
      <i>-cell_style   = &5.
      <i>-data_type    = &6.
    end-of-definition.

    clear lt_content.
    _add_cell 1 1 'Content'    'A1'  '0230522560691EEA84F5D3267CA7FD65'  's'.
    _add_cell 1 2 'SaveToText' 'B1'  '0230522560691EEA84F5D3267CA7FD65'  's'.
    _add_cell 2 1 'Sheet1'     'A2'  '00000000000000000000000000000000'  's'.
    _add_cell 2 2 'X'          'B2'  '00000000000000000000000000000000'  's'.
    _add_cell 3 1 'Sheet2'     'A3'  '00000000000000000000000000000000'  's'.
    _add_cell 4 1 'Sheet3'     'A4'  '00000000000000000000000000000000'  's'.
    _add_cell 4 2 'X'          'B4'  '00000000000000000000000000000000'  's'.
    _add_cell 5 1 'Sheet4'     'A5'  '00000000000000000000000000000000'  's'.
    _add_cell 5 2 'X'          'B5'  '00000000000000000000000000000000'  's'.
    rs_content_samples-contents = lt_content.

    clear lt_content.
    _add_cell 1 1 'to_exclude' 'A1' '0230522560691EEA84F5D3267CA87D65'  's'.
    _add_cell 2 1 'Sheet3'     'A2' '00000000000000000000000000000000'  's'.
    rs_content_samples-excludes = lt_content.

    clear lt_content.
    _add_cell 1 1 'Column1'   'A1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 1 2 'Column2'   'B1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 2 1 'A'         'A2' '00000000000000000000000000000000' 's'.
    _add_cell 2 2 '1'         'B2' '0230522560691EEA84F5D3267CA81D65' ''.
    _add_cell 3 1 'B'         'A3' '00000000000000000000000000000000' 's'.
    _add_cell 3 2 '2'         'B3' '00000000000000000000000000000000' ''.
    _add_cell 4 1 'C'         'A4' '00000000000000000000000000000000' 's'.
    _add_cell 4 2 '3'         'B4' '00000000000000000000000000000000' ''.
    _add_cell 6 1 'More_data' 'A6' '00000000000000000000000000000000' 's'.
    _add_cell 6 2 'to_skip'   'B6' '00000000000000000000000000000000' 's'.
    rs_content_samples-simple = lt_content.

    clear lt_content.
    _add_cell 1 1 'A'     'A1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 1 2 'B'     'B1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 1 3 'C'     'C1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 1 4 'D'     'D1' '0230522560691EEA84F5D3267CA7FD65' 's'.
    _add_cell 2 1 'Vasya' 'A2' '00000000000000000000000000000000' 's'.
    _add_cell 2 2 '43344' 'B2' '0230522560691EEA84F5D3267CA85D65' ''.
    _add_cell 2 3 '15'    'C2' '00000000000000000000000000000000' ''.
    _add_cell 2 4 '1'     'D2' '00000000000000000000000000000000' 'b'.
    _add_cell 3 1 'Petya' 'A3' '00000000000000000000000000000000' 's'.
    _add_cell 3 2 '43345' 'B3' '0230522560691EEA84F5D3267CA85D65' ''.
    _add_cell 3 3 '16.37' 'C3' '0230522560691EEA84F5D3267CA83D65' ''.
    _add_cell 3 4 '0'     'D3' '00000000000000000000000000000000' 'b'.
    rs_content_samples-with_dates = lt_content.

  endmethod.

endclass.

**********************************************************************
* UNIT TESTS
**********************************************************************

class ltcl_workbook_parser_test definition final for testing
  duration short
  risk level harmless.

  private section.
    data mt_dummy_sheet    type lif_excel=>tt_sheet_content.
    data mt_dummy_contents type lif_excel=>tt_sheet_content.
    data mt_dummy_exclude  type lif_excel=>tt_sheet_content.

    methods setup.
    methods read_row for testing.
    methods is_row_empty for testing.
    methods clip_header for testing.
    methods clip_rows for testing.
    methods read_contents for testing.
    methods convert_sheet for testing.
    methods read_exclude for testing.
    " TODO integration test

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

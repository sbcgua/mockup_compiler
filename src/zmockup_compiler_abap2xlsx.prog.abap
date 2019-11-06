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

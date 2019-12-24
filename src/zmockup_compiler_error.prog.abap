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

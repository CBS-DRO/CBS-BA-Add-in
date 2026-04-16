option explicit

public sub showerror(modulename as string, procedurename as string, _
      errornumber as long, errordescription as string)
   '* purpose  :  provide a central error handling mechanism.
   on error goto proc_err
   dim message as string
   dim title as string

   '* build the error message.
   message = "error number: " & errornumber & vbcrlf & _
                "description: " & errordescription & vbcrlf & vbcrlf & _
                "module: " & modulename & vbcrlf & _
                "procedure: " & procedurename
   
    '* build the title for the message box.
    title = "error in excel " & application.version & ", " _
        & application.operatingsystem & ", solvertable version december 2013"

    msgbox message, vbcritical, title

proc_exit:
   exit sub
   
proc_err:
   resume next
   
end sub




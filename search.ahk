;Start
FileDelete *.txt

;Open PS xlsx------------------------------------------------------------------
#ps_file = %A_ScriptDir%\ps.xlsx

FileSelectFile, ps_file,,c:\users\jthobbs\downloads\,,*.xlsx

xl         := ComObjCreate("Excel.Application")
xl.Visible := False
book       := xl.Workbooks.Open(ps_file)

;Save as plain-text
save_path = %A_ScriptDir%\_ps.txt
book.SaveAs(save_path, -4158) ;Tab-delmited file

;Close XLS file
book.Close(savechanges := False)
xl.Quit

;Search database
FileRead, data, _ps.txt

Search:
    Inputbox, q, InvoiceID, %display_row%
    If q =
    {
        exit
    }
    q := RegExReplace(q, "Unique.*", "")
    q := RegExReplace(q, "Invoice.*", "")
    
    display_row = Not found. ;Default value
    
    if q contains amount=
    {
        search_type = amount
    }
    else
    {
        search_type = keyword
    }
    
    ;Query database
    Loop, parse, data, `n
    {
        row := A_LoopField
        Stringsplit, r, row, `t
        invoiceID := r9
        vendor := r8
        documentID := r7
        desc := r13
        amount := r14
        
        if search_type = keyword
        {
            If row contains %q%
            {
                display_row = Vendor: %vendor%`nInvoiceID: %invoiceID%`nDocumentID: %documentID%`nDescr: %desc%`nAmount: %amount%
                clipboard := documentID
                continue
            }
        }
        
        if search_type = amount
        {
            StringReplace, amount, amount, amount=,,All
            If amount = %q%
            {
                display_row = Vendor: %vendor%`nInvoiceID: %invoiceID%`nDocumentID: %documentID%`nDescr: %desc%`nAmount: %amount%
                clipboard := documentID
                continue
            }
        }        
    }
    
    goto Search
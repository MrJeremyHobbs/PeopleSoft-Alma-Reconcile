# PeopleSoft-Alma-Reconcile
A simple search utility written in AutoHotkey to search a Transactions - Actuals for Download.xlsx from PeopleSoft dashboards and paste the DocumentID into a paid invoice in Alma.

Enter the InvoiceID or other string to search the transactions file.

If found, the DocumentID will be copied to the clipboard.

In Alma, from the Waiting for Payment queue, you can change the status of the invoice as "Paid" and paste the DocumentID into the "Payment Identifier" field.
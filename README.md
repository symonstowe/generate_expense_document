# generate_expense_document
Given a gmail account get the attachements from all identified emails and put them together in a single file.

1. Emails with images of receipts are sent to an address
2. The company that will be billed for the expense should be included in the subject line
3. The text content of the email should be a description of the expense
4. The attachments will be sorted in the document with a heading for each date that an email was received

Required inputs:
  company - the company that will be billed for the expense
  curr - the currency that is used for the document
  email - gmail address
Optional inputs:
  doc_title - desired output title of the document 
  start_date - start date of the expense form
  end_date - end date of the expense report 
  unread - load data only from undread emails
  
  Outputs:
   doc_title.tex - latex document with header including info and personal company info
                 - Info on date payable and payment info 

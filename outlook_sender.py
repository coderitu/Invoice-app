import win32com.client
from app_logger import log_info, log_error

def send_test_email(
    to_email,
    member_name,
    invoice_no,
    shared_mailbox,
    pdf_path,
    email_template,
    test_mode=True
):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to_email
        mail.CC = "membership.accounts@wdclubsafrica.com"
        mail.Subject = f"Invoice {invoice_no}"

        mail.Body = email_template.format(
            name=member_name,
            invoice_no=invoice_no
        )

        mail.SentOnBehalfOfName = shared_mailbox
        mail.Attachments.Add(pdf_path)

        if test_mode:
            mail.Display()
        else:
            mail.Send()

        log_info(f"Email {'prepared' if test_mode else 'sent'} | Invoice: {invoice_no} | To: {to_email}")
        return True, "Email successful"

    except Exception as e:
        log_error(f"Email failed | Invoice: {invoice_no} | To: {to_email} | Error: {e}")
        return False, str(e)

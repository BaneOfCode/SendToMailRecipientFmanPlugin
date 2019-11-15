from fman import DirectoryPaneCommand, show_alert
import subprocess
from fman.url import as_human_readable, basename
import sys
import win32com.client as win32

class SendToMailRecipient(DirectoryPaneCommand):
    def __call__(self):
        chosen_file = self.pane.get_file_under_cursor()
        file_to_send = as_human_readable(chosen_file)
        file_name = basename(chosen_file)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = "Attachment  file: " + file_name
        mail.Attachments.Add(Source=file_to_send)
        mail.Display(True)

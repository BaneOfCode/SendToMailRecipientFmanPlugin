from fman import DirectoryPaneCommand, show_alert
import subprocess
from fman.url import as_human_readable, basename
import sys
import win32com.client as win32

class SendToMailRecipient(DirectoryPaneCommand):
    def __call__(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        chosen_files = self.pane.get_selected_files()
        if not chosen_files:
            show_alert('No file selected.')
            return
        for chosen_file in chosen_files:
            file_to_send = as_human_readable(chosen_file)
            mail.Attachments.Add(Source=file_to_send)

        mail.Display(True)

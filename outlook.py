import win32com.client
import os
from datetime import datetime, timedelta


# WORKS! IT MANAGES TO FILTER INBOX ITEMS BY DAYS BACK!!

def download_pdf_from_outlook(save_folder_path, subject_filter=None, sender_filter=None, days=7):
    save_folder_path = os.path.abspath(save_folder_path)
    os.makedirs(save_folder_path, exist_ok=True)
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    
    try:
        # Format date properly for Outlook
        cutoff_date = datetime.now() - timedelta(days=days)
        received_str = cutoff_date.strftime('%m/%d/%Y')
        
        # Test different filter formats
        # Try 1: Basic date filter
        filter_str = "@SQL=urn:schemas:httpmail:datereceived >= '" + received_str + "'"
        
        # Try 2: If you want to add subject filter
        if subject_filter:
            filter_str += " AND urn:schemas:httpmail:subject LIKE '%" + subject_filter + "%'"
            
        print(f"Using filter: {filter_str}")  # Debug print
        
        messages = inbox.Items.Restrict(filter_str)
        messages.Sort("[ReceivedTime]", True)
                
    except Exception as e:
        print(f"Error with filter: {str(e)}")
        # Fallback to original method if filtering fails
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
    
    print(f"\nApplying filters:")
    print(f"- Emails since: {received_str}")
    print(f"- Subject containing: {subject_filter}")
    
    downloaded_count = 0
    email_count = 0
    
    for message in messages:
        try:
            email_count += 1
            print(f"\nProcessing email #{email_count}:")
            print(f"Subject: {message.Subject}")
            print(f"From: {message.SenderEmailAddress}")
            
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    if attachment.FileName.lower().endswith('.pdf'):
                        save_path = os.path.join(save_folder_path, attachment.FileName)
                        
                        if os.path.exists(save_path):
                            print(f"File already exists: {attachment.FileName}")
                            continue
                            
                        try:
                            attachment.SaveAsFile(save_path)
                            downloaded_count += 1
                            print(f"Saved PDF: {attachment.FileName}")
                        except Exception as save_error:
                            print(f"Error saving attachment: {save_error}")
                            
        except Exception as e:
            print(f"Error processing email: {str(e)}")
            continue
    
    print(f"\nSummary:")
    print(f"Emails processed: {email_count}")
    print(f"PDFs downloaded: {downloaded_count}")
    return downloaded_count

def main():
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    save_folder_path = os.path.join(current_dir, "PO_not_in_excel")
    
    subject_filter = "Hotel booking"  
    days = 7    
    
    try:
        count = download_pdf_from_outlook(
            save_folder_path=save_folder_path,
            subject_filter=subject_filter,
            sender_filter=None,
            days=days
        )    
        print(f"\nScript completed successfully!")
    except Exception as e:
        print(f"Error in main: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()